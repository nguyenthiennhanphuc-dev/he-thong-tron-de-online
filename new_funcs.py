def parse_and_build_pools(files):
    merged_stream = merge_and_track_files(files)
    doc = Document(merged_stream)
    sections, outro = extract_sections_xml(doc)
    
    global_pools = {} 
    structure = []
    
    title_to_part = {}
    part_idx = 0
    
    for sec in sections:
        title = sec['title']
        if title != 'CHUNG' and title not in title_to_part:
            match = re.search(r'PHẦN\s+([IVX]+)', title.upper())
            if match:
                roman = match.group(1)
                if roman == 'I': title_to_part[title] = 'part_1'
                elif roman == 'II': title_to_part[title] = 'part_2'
                elif roman == 'III': title_to_part[title] = 'part_3'
                else: 
                    part_idx += 1
                    title_to_part[title] = f'part_{part_idx}'
            else:
                part_idx += 1
                title_to_part[title] = f'part_{part_idx}'
            
        if title not in global_pools: global_pools[title] = []
        
        start_q_idx = sum(1 for item in global_pools[title] if item['type'] == 'q')
        start_q_idx += sum(len(item['data']['questions']) for item in global_pools[title] if item['type'] == 'g')
        
        current_q_idx = start_q_idx + 1
        
        for item in sec['items']:
            if item['type'] == 'q':
                item['original_index'] = current_q_idx
                current_q_idx += 1
            elif item['type'] == 'g':
                item['original_indices'] = list(range(current_q_idx, current_q_idx + len(item['data']['questions'])))
                current_q_idx += len(item['data']['questions'])
                
        global_pools[title].extend(sec['items'])
        if len(sec['items']) > 0 and sec['items'][0]['is_temp']:
            if not any(s['title'] == title for s in structure):
                structure.append({'title': title, 'intro': sec['intro'], 'needed': len(sec['items'])})
            else:
                for s in structure:
                    if s['title'] == title: s['needed'] += len(sec['items'])

    return global_pools, structure, title_to_part, outro, merged_stream.getvalue()

def generate_preview(global_pools, structure, title_to_part, manual_answers=None):
    preview_data = []
    
    q_count_mock = 1
    for sec_template in structure:
        title = sec_template['title']
        pool = global_pools.get(title, [])
        part_key = title_to_part.get(title)
        part_manual_ans = manual_answers.get(part_key, {}) if manual_answers else {}
        
        part_preview = {'title': title, 'questions': []}
        
        for item in pool:
            if item['type'] == 'q':
                manual_ans_for_item = part_manual_ans.get(item.get('original_index'))
                correct_label, els = process_question_block(item['data'], q_count_mock, shuffle_a=False, keep_color=True, manual_ans=manual_ans_for_item)
                
                full_text = '\n'.join(p.text for p in els if p.text.strip())
                part_preview['questions'].append({
                    'text': full_text, 
                    'answer': correct_label,
                    'orig_idx': item.get('original_index')
                })
                q_count_mock += 1
            elif item['type'] == 'g':
                g_data = copy.deepcopy(item['data'])
                intro_text = '\n'.join(p.text for p in g_data['intro'] if p.text.strip())
                
                group_qs = g_data['questions']
                sub_q_indices = item.get('original_indices', [])
                zipped_qs = list(zip(group_qs, sub_q_indices))
                
                for q_blocks, orig_sub_idx in zipped_qs:
                    manual_ans_for_sub = part_manual_ans.get(orig_sub_idx)
                    correct_label, els = process_question_block(q_blocks, q_count_mock, shuffle_a=False, keep_color=True, manual_ans=manual_ans_for_sub)
                    full_text = '\n'.join(p.text for p in els if p.text.strip())
                    
                    part_preview['questions'].append({
                        'text': f'[Nhóm dùng chung]\n{intro_text}\n{full_text}', 
                        'answer': correct_label,
                        'orig_idx': orig_sub_idx
                    })
                    q_count_mock += 1
                    
        preview_data.append(part_preview)
    return preview_data

def generate_mixed_exams(global_pools, structure, title_to_part, outro, merged_bytes, num_versions=4, q_sort_mode='random', shuffle_a=True, even_split=True, keep_color=False, manual_answers=None):
    zip_buffer = io.BytesIO()
    master_key_doc = Document()
    master_key_doc.add_heading('BẢNG ĐÁP ÁN TỔNG HỢP & TRUY VẾT NGUỒN', 0)

    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zip_file:
        for v in range(num_versions):
            ma_de = str(random.randint(100, 999))
            
            merged_stream = io.BytesIO(merged_bytes)
            new_doc = Document(merged_stream)
            body = new_doc._body._body
            
            sectPr = body.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
            if sectPr is not None: body.remove(sectPr)
            for child in list(body):
                if child.tag.endswith(('}p', '}tbl')): body.remove(child)
            
            ans_key = {}
            source_info = []
            q_count = 1
            elements_to_keep = []
            
            for sec_template in structure:
                for p in sec_template['intro']: elements_to_keep.append(copy.deepcopy(p._element))
                title = sec_template['title']
                pool = global_pools.get(title, [])
                needed = sec_template['needed']
                
                actual_needed = min(needed, len(pool))
                randomize = (q_sort_mode == 'random')
                if even_split:
                    selected_items = allocate_questions_fairly(pool, actual_needed, randomize=randomize)
                else:
                    if randomize:
                        selected_items = random.sample(pool, actual_needed)
                    else:
                        selected_items = pool[:actual_needed]
                
                if q_sort_mode == 'random':
                    random.shuffle(selected_items)
                elif q_sort_mode == 'reverse':
                    selected_items.reverse()
                
                part_key = title_to_part.get(title)
                part_manual_ans = manual_answers.get(part_key, {}) if manual_answers else {}
                
                for item in selected_items:
                    src_name = item['src']
                    if item['type'] == 'q':
                        manual_ans_for_item = part_manual_ans.get(item.get('original_index'))
                        correct_label, els = process_question_block(item['data'], q_count, shuffle_a, keep_color, manual_ans=manual_ans_for_item)
                        elements_to_keep.extend(els)
                        ans_key[q_count] = correct_label
                        source_info.append(f'Câu {q_count}: Nguồn từ file [{src_name}]')
                        q_count += 1
                    elif item['type'] == 'g':
                        g_data = copy.deepcopy(item['data'])
                        for p in g_data['intro']: 
                            p.text = re.sub(r'^(Câu|Bài)\s+\d+', f'Câu {q_count}', p.text, flags=re.IGNORECASE)
                            elements_to_keep.append(copy.deepcopy(p._element))
                        
                        group_qs = g_data['questions']
                        sub_q_indices = item.get('original_indices', [])
                        
                        zipped_qs = list(zip(group_qs, sub_q_indices))
                        
                        if q_sort_mode == 'random':
                            random.shuffle(zipped_qs)
                        elif q_sort_mode == 'reverse':
                            zipped_qs.reverse()
                            
                        for q_blocks, orig_sub_idx in zipped_qs:
                            manual_ans_for_sub = part_manual_ans.get(orig_sub_idx)
                            correct_label, els = process_question_block(q_blocks, q_count, shuffle_a, keep_color, manual_ans=manual_ans_for_sub)
                            elements_to_keep.extend(els)
                            ans_key[q_count] = correct_label
                            source_info.append(f'Câu {q_count}: Nguồn từ file [{src_name}]')
                            q_count += 1
            
            for p in outro: elements_to_keep.append(copy.deepcopy(p._element))
            for el in elements_to_keep: body.append(el)
                
            try:
                new_doc.add_page_break()
                new_doc.add_paragraph(f'BẢNG ĐÁP ÁN - MÃ ĐỀ {ma_de}').bold = True
                table = new_doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Câu'; hdr_cells[1].text = 'Đáp án'; hdr_cells[2].text = 'Câu'; hdr_cells[3].text = 'Đáp án'
                total_qs = q_count - 1
                for i in range(1, total_qs + 1, 2):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(i); row_cells[1].text = ans_key.get(i, '')
                    if i + 1 <= total_qs:
                        row_cells[2].text = str(i + 1); row_cells[3].text = ans_key.get(i + 1, '')
            except Exception:
                new_doc.add_paragraph('--- ĐÁP ÁN ---')
                for i in range(1, q_count): new_doc.add_paragraph(f'Câu {i}: {ans_key.get(i, "")}')
                    
            new_doc.add_paragraph('')
            new_doc.add_paragraph('📍 TRUY VẾT NGUỒN GỐC CÂU HỎI').bold = True
            for info in source_info: new_doc.add_paragraph(info)
                
            if sectPr is not None: body.append(sectPr)
            buf = io.BytesIO(); new_doc.save(buf)
            zip_file.writestr(f'De_{ma_de}.docx', buf.getvalue())
            
            master_key_doc.add_heading(f'MÃ ĐỀ: {ma_de}', level=1)
            try:
                table_mk = master_key_doc.add_table(rows=1, cols=4); table_mk.style = 'Table Grid'
                h_mk = table_mk.rows[0].cells
                h_mk[0].text = 'Câu'; h_mk[1].text = 'Đáp án'; h_mk[2].text = 'Câu'; h_mk[3].text = 'Đáp án'
                for i in range(1, q_count, 2):
                    r_mk = table_mk.add_row().cells
                    r_mk[0].text = str(i); r_mk[1].text = ans_key.get(i, '')
                    if i + 1 < q_count: r_mk[2].text = str(i + 1); r_mk[3].text = ans_key.get(i + 1, '')
            except Exception: pass
            
            master_key_doc.add_paragraph('📍 Nguồn gốc câu hỏi:').bold = True
            for info in source_info: master_key_doc.add_paragraph(info, style='List Bullet')
            master_key_doc.add_page_break()

    zip_buffer.seek(0)
    master_buf = io.BytesIO(); master_key_doc.save(master_buf); master_buf.seek(0)
    return zip_buffer, master_buf
