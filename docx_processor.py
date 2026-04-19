import io
import zipfile
import random
import re
import copy
from docx import Document
from docxcompose.composer import Composer
from docx.shared import RGBColor

def fix_docx_floats(file_stream):
    """Gọt số thập phân chống crash"""
    file_stream.seek(0)
    zip_in = zipfile.ZipFile(file_stream)
    zip_buffer = io.BytesIO()
    zip_out = zipfile.ZipFile(zip_buffer, 'w')
    for item in zip_in.infolist():
        content = zip_in.read(item.filename)
        if item.filename.endswith('.xml'):
            content_str = content.decode('utf-8')
            header_end = content_str.find('?>') + 2
            if header_end > 1:
                header = content_str[:header_end]
                body = content_str[header_end:]
                body = re.sub(r'="(\-?[0-9]+)\.[0-9]+"', r'="\1"', body)
                content_str = header + body
            else:
                content_str = re.sub(r'="(\-?[0-9]+)\.[0-9]+"', r'="\1"', content_str)
            content = content_str.encode('utf-8')
        zip_out.writestr(item, content)
    zip_out.close()
    zip_buffer.seek(0)
    return zip_buffer

def merge_and_track_files(files):
    """Hợp thể các file để không làm đứt liên kết kho ảnh"""
    master_stream = fix_docx_floats(files[0])
    master = Document(master_stream)
    p_src = master.add_paragraph(f"@SOURCE_FILE: {files[0].name}")
    master._body._body.insert(0, p_src._element) 
    composer = Composer(master)
    for f in files[1:]:
        f_stream = fix_docx_floats(f)
        doc_append = Document(f_stream)
        p_src_app = doc_append.add_paragraph(f"@SOURCE_FILE: {f.name}")
        doc_append._body._body.insert(0, p_src_app._element)
        composer.append(doc_append)
    merged_io = io.BytesIO()
    composer.doc.save(merged_io)
    merged_io.seek(0)
    return merged_io

def clean_source_tags(paragraph):
    """Tự động dọn dẹp các mã nguồn như [1], """
    for run in paragraph.runs:
        if run.text: run.text = re.sub(r'\[.*?\]\s*', '', run.text)

def enhance_visuals(paragraph):
    """Phóng to chữ cho học sinh dễ đọc và căn giữa hình ảnh"""
    for run in paragraph.runs:
        if run.font.size is None: run.font.size = 177800
    if 'Graphic' in paragraph._p.xml or 'pic:pic' in paragraph._p.xml:
        paragraph.alignment = 1

def extract_sections_xml(doc):
    """Máy quét phân khu (Phần I, II, III), Nhóm chung và Đọc chữ HẾT"""
    sections = []
    current_section = {"title": "CHUNG", "intro": [], "items": []} 
    current_q = []
    current_g = None
    outro = []
    is_outro = False
    current_src = "Chưa rõ nguồn"
    is_template = True
    file_count = 0
    
    for p in doc.paragraphs:
        clean_source_tags(p)
        enhance_visuals(p)
        text = p.text.strip()
        
        if text.startswith("@SOURCE_FILE:"):
            if current_q:
                if current_g: current_g["questions"].append(current_q)
                else: current_section["items"].append({"type": "q", "data": current_q, "src": current_src, "is_temp": is_template})
                current_q = []
            if current_g:
                current_section["items"].append({"type": "g", "data": current_g, "src": current_src, "is_temp": is_template})
                current_g = None
            if current_section["intro"] or current_section["items"]:
                sections.append(current_section)
            current_section = {"title": "CHUNG", "intro": [], "items": []}
            is_outro = False
            current_src = text.replace("@SOURCE_FILE:", "").strip()
            file_count += 1
            if file_count > 1: is_template = False
            p.text = "" 
            continue
            
        if is_outro:
            if is_template: outro.append(p)
            continue
            
        is_het = re.match(r'^[\-\—\_\s\*]*HẾT[\-\—\_\s\*]*$', text.upper())
        is_phan = re.match(r'^PHẦN\s+[IVX\d]+', text.upper())
        is_start_group = "@BẮT ĐẦU" in text.upper()
        is_end_group = "@KẾT THÚC" in text.upper()
        is_cau = re.match(r'^Câu\s+\d+[\.:]', text, re.IGNORECASE)
        
        if is_het:
            is_outro = True
            if current_q:
                if current_g: current_g["questions"].append(current_q)
                else: current_section["items"].append({"type": "q", "data": current_q, "src": current_src, "is_temp": is_template})
                current_q = []
            if current_g:
                current_section["items"].append({"type": "g", "data": current_g, "src": current_src, "is_temp": is_template})
                current_g = None
            if current_section["intro"] or current_section["items"]:
                sections.append(current_section)
                current_section = {"title": "CHUNG", "intro": [], "items": []}
            if is_template: outro.append(p)
            continue
        
        if is_phan:
            if current_q:
                if current_g: current_g["questions"].append(current_q)
                else: current_section["items"].append({"type": "q", "data": current_q, "src": current_src, "is_temp": is_template})
                current_q = []
            if current_g:
                current_section["items"].append({"type": "g", "data": current_g, "src": current_src, "is_temp": is_template})
                current_g = None
            if current_section["intro"] or current_section["items"]:
                sections.append(current_section)
            current_section = {"title": text, "intro": [p], "items": []}
        elif is_start_group:
            if current_q:
                if current_g: current_g["questions"].append(current_q)
                else: current_section["items"].append({"type": "q", "data": current_q, "src": current_src, "is_temp": is_template})
                current_q = []
            if current_g: current_section["items"].append({"type": "g", "data": current_g, "src": current_src, "is_temp": is_template})
            current_g = {"intro": [p], "questions": []}
        elif is_end_group:
            if current_g:
                current_g["intro"].append(p)
                if current_q:
                    current_g["questions"].append(current_q)
                    current_q = []
                current_section["items"].append({"type": "g", "data": current_g, "src": current_src, "is_temp": is_template})
                current_g = None
            else:
                if current_q: current_q.append(p)
                else: current_section["intro"].append(p)
        elif is_cau:
            if current_q:
                if current_g: current_g["questions"].append(current_q)
                else: current_section["items"].append({"type": "q", "data": current_q, "src": current_src, "is_temp": is_template})
            current_q = [p]
        else:
            if current_q: current_q.append(p)
            elif current_g: current_g["intro"].append(p)
            else: current_section["intro"].append(p)
            
    if current_q:
        if current_g: current_g["questions"].append(current_q)
        else: current_section["items"].append({"type": "q", "data": current_q, "src": current_src, "is_temp": is_template})
    if current_g:
        current_section["items"].append({"type": "g", "data": current_g, "src": current_src, "is_temp": is_template})
    if current_section["intro"] or current_section["items"]:
        sections.append(current_section)
    return sections, outro

def fix_choice_label(paragraph, new_label):
    try:
        for run in paragraph.runs:
            text = run.text.strip()
            if re.match(r'^[A-D][\.\)]', text):
                run.text = re.sub(r'^[A-D][\.\)]', new_label, run.text, count=1)
                break
            elif text in ['A', 'B', 'C', 'D']:
                run.text = run.text.replace(text, new_label.replace('.', ''))
                break
    except Exception: pass

def is_correct_choice(paragraph):
    """
    Sniper Model: Chỉ quét Run đầu tiên chứa chữ để miễn nhiễm 100% với lỗi từ MathType.
    """
    try:
        for run in paragraph.runs:
            if run.text.strip(): 
                xml_str = run._element.xml
                if 'w:val="FF0000"' in xml_str or 'w:val="red"' in xml_str.lower() or '<w:u ' in xml_str or '<w:u/>' in xml_str:
                    return True
                if run.font.color and str(run.font.color.rgb) == 'FF0000':
                    return True
                if run.font.underline:
                    return True
                return False 
    except Exception: pass
    
    try:
        xml_str = paragraph._p.xml
        if 'w:val="FF0000"' in xml_str or 'w:val="red"' in xml_str.lower() or '<w:u ' in xml_str or '<w:u/>' in xml_str:
            return True
    except Exception: pass
    return False

def extract_answer_text(paragraph):
    """Hút phần nội dung chữ màu đỏ/gạch chân (Dự phòng)"""
    ans_parts = []
    for run in paragraph.runs:
        try:
            xml_str = run._element.xml
            if 'w:val="FF0000"' in xml_str or 'w:val="red"' in xml_str.lower() or '<w:u ' in xml_str or '<w:u/>' in xml_str:
                if run.text: ans_parts.append(run.text)
        except Exception: pass
    return "".join(ans_parts).strip()

def process_question_block(q_blocks, q_num, shuffle_a):
    """Trái tim xử lý: Nhận diện, Đổi màu tàng hình, Rút trích Đ/S và Lời giải"""
    q_blocks = copy.deepcopy(q_blocks)
    first_p = q_blocks[0]
    first_p.text = re.sub(r'^Câu\s+\d+', f'Câu {q_num}', first_p.text, flags=re.IGNORECASE)
    
    correct_label = ""
    mcq_indices = []
    tf_indices = []
    
    for i, p in enumerate(q_blocks):
        text = p.text.strip()
        if re.match(r'^[A-D][\.\)]', text): mcq_indices.append(i)
        elif re.match(r'^[a-d][\.\)]', text): tf_indices.append(i)
        
    elements_to_keep = []
    
    if len(mcq_indices) == 4:
        choices = [q_blocks[i] for i in mcq_indices]
        if shuffle_a: random.shuffle(choices)
        labels = ['A', 'B', 'C', 'D']; labels_text = ['A.', 'B.', 'C.', 'D.']
        for i, orig_idx in enumerate(mcq_indices):
            q_blocks[orig_idx] = choices[i]
            fix_choice_label(q_blocks[orig_idx], labels_text[i])
            if is_correct_choice(choices[i]): correct_label = labels[i]
        if not correct_label and not shuffle_a:
            c_idx = 0
            for p in q_blocks:
                if re.match(r'^[A-D][\.\)]', p.text.strip()):
                    if is_correct_choice(p) and c_idx < 4: correct_label = labels[c_idx]
                    c_idx += 1
                    
        for p in q_blocks:
            for run in p.runs:
                try:
                    xml_str = run._element.xml
                    if 'w:val="FF0000"' in xml_str or 'w:val="red"' in xml_str.lower() or '<w:u ' in xml_str or '<w:u/>' in xml_str:
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.underline = False
                except Exception: pass
            elements_to_keep.append(copy.deepcopy(p._element))
            
    elif len(tf_indices) == 4:
        tf_answers = []
        for i in tf_indices:
            if is_correct_choice(q_blocks[i]): tf_answers.append("Đ")
            else: tf_answers.append("S")
            
        correct_label = ", ".join([f"{['a','b','c','d'][idx]}-{ans}" for idx, ans in enumerate(tf_answers)])
        
        for p in q_blocks:
            for run in p.runs:
                try:
                    xml_str = run._element.xml
                    if 'w:val="FF0000"' in xml_str or 'w:val="red"' in xml_str.lower() or '<w:u ' in xml_str or '<w:u/>' in xml_str:
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        run.font.underline = False
                except Exception: pass
            elements_to_keep.append(copy.deepcopy(p._element))
            
    else:
        ans_texts = []
        for p in q_blocks:
            text_strip = p.text.strip()
            
            match_prefix = re.match(r'^(Đáp án|ĐS|Kết quả|Lời giải|Đáp số)[\s\:\.]*(.*)', text_strip, re.IGNORECASE)
            if match_prefix:
                ans_val = match_prefix.group(2).strip()
                if ans_val: ans_texts.append(ans_val)
                continue 
                
            txt_red = extract_answer_text(p)
            if txt_red:
                clean_txt = re.sub(r'^(Đáp án|ĐS|Kết quả|Lời giải|Đáp số)[\s\:]*', '', txt_red, flags=re.IGNORECASE).strip()
                if clean_txt: ans_texts.append(clean_txt)
                for run in p.runs:
                    try:
                        xml_str = run._element.xml
                        if 'w:val="FF0000"' in xml_str or 'w:val="red"' in xml_str.lower() or '<w:u ' in xml_str or '<w:u/>' in xml_str:
                            run.text = "" 
                    except Exception: pass
                    
            elements_to_keep.append(copy.deepcopy(p._element))
            
        if ans_texts: correct_label = " ; ".join(ans_texts)
            
    return correct_label, elements_to_keep

def allocate_questions_fairly(pool, needed):
    """
    Thuật toán chia đều số lượng câu hỏi từ nhiều nguồn (files) khác nhau.
    Có khả năng bù trừ (nếu file A thiếu câu thì bốc bù từ file B).
    """
    pool_by_src = {}
    for item in pool:
        src = item["src"]
        if src not in pool_by_src:
            pool_by_src[src] = []
        pool_by_src[src].append(item)
        
    sources = list(pool_by_src.keys())
    if not sources or needed <= 0:
        return []
        
    quota = {src: 0 for src in sources}
    base = needed // len(sources)
    rem = needed % len(sources)
    
    # Bước 1: Chia đều cơ bản
    for src in sources:
        quota[src] = base
    # Phân phát số dư (nếu chia không hết)
    for src in random.sample(sources, rem):
        quota[src] += 1
        
    # Bước 2: Bù trừ (Nếu có file không đủ số lượng quota yêu cầu)
    while True:
        shortfall = 0
        active_sources = []
        for src in sources:
            avail = len(pool_by_src[src])
            if quota[src] > avail:
                shortfall += (quota[src] - avail)
                quota[src] = avail # Chốt lấy tối đa số câu file này có
            elif quota[src] < avail:
                active_sources.append(src) # Các file còn dư dả để bốc bù
        
        if shortfall == 0 or not active_sources:
            break # Hoàn hảo hoặc hết file để bù
            
        add_base = shortfall // len(active_sources)
        add_rem = shortfall % len(active_sources)
        for src in active_sources:
            quota[src] += add_base
        for src in random.sample(active_sources, add_rem):
            quota[src] += 1

    # Bước 3: Bốc ngẫu nhiên dựa trên Quota đã chốt
    selected_items = []
    for src in sources:
        if quota[src] > 0:
            selected_items.extend(random.sample(pool_by_src[src], min(quota[src], len(pool_by_src[src]))))
            
    return selected_items

def process_and_shuffle_multi(files, num_versions=4, shuffle_q=True, shuffle_a=True, even_split=True):
    merged_stream = merge_and_track_files(files)
    doc = Document(merged_stream)
    sections, outro = extract_sections_xml(doc)
    
    global_pools = {} 
    structure = []
    
    for sec in sections:
        title = sec["title"]
        if title not in global_pools: global_pools[title] = []
        global_pools[title].extend(sec["items"])
        if len(sec["items"]) > 0 and sec["items"][0]["is_temp"]:
            if not any(s["title"] == title for s in structure):
                structure.append({"title": title, "intro": sec["intro"], "needed": len(sec["items"])})
            else:
                for s in structure:
                    if s["title"] == title: s["needed"] += len(sec["items"])

    zip_buffer = io.BytesIO()
    master_key_doc = Document()
    master_key_doc.add_heading('BẢNG ĐÁP ÁN TỔNG HỢP & TRUY VẾT NGUỒN', 0)

    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
        for v in range(num_versions):
            ma_de = str(random.randint(100, 999))
            merged_stream.seek(0)
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
                for p in sec_template["intro"]: elements_to_keep.append(copy.deepcopy(p._element))
                title = sec_template["title"]
                pool = global_pools.get(title, [])
                needed = sec_template["needed"]
                
                # ------ GỌI THUẬT TOÁN CHIA ĐỀU TẠI ĐÂY ------
                actual_needed = min(needed, len(pool))
                if even_split:
                    selected_items = allocate_questions_fairly(pool, actual_needed)
                else:
                    selected_items = random.sample(pool, actual_needed)
                
                if shuffle_q: random.shuffle(selected_items)
                # ----------------------------------------------
                
                for item in selected_items:
                    src_name = item["src"]
                    if item["type"] == "q":
                        correct_label, els = process_question_block(item["data"], q_count, shuffle_a)
                        elements_to_keep.extend(els)
                        ans_key[q_count] = correct_label
                        source_info.append(f"Câu {q_count}: Nguồn từ file [{src_name}]")
                        q_count += 1
                    elif item["type"] == "g":
                        for p in item["data"]["intro"]: elements_to_keep.append(copy.deepcopy(p._element))
                        group_qs = item["data"]["questions"]
                        if shuffle_q:
                            group_qs = copy.deepcopy(group_qs)
                            random.shuffle(group_qs)
                        for q_blocks in group_qs:
                            correct_label, els = process_question_block(q_blocks, q_count, shuffle_a)
                            elements_to_keep.extend(els)
                            ans_key[q_count] = correct_label
                            source_info.append(f"Câu {q_count}: Nguồn từ file [{src_name}]")
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
                    row_cells[0].text = str(i); row_cells[1].text = ans_key.get(i, "")
                    if i + 1 <= total_qs:
                        row_cells[2].text = str(i + 1); row_cells[3].text = ans_key.get(i + 1, "")
            except Exception:
                new_doc.add_paragraph("--- ĐÁP ÁN ---")
                for i in range(1, q_count): new_doc.add_paragraph(f"Câu {i}: {ans_key.get(i, '')}")
                    
            new_doc.add_paragraph("")
            new_doc.add_paragraph("📍 TRUY VẾT NGUỒN GỐC CÂU HỎI").bold = True
            for info in source_info: new_doc.add_paragraph(info)
                
            if sectPr is not None: body.append(sectPr)
            buf = io.BytesIO(); new_doc.save(buf)
            zip_file.writestr(f"De_{ma_de}.docx", buf.getvalue())
            
            master_key_doc.add_heading(f"MÃ ĐỀ: {ma_de}", level=1)
            try:
                table_mk = master_key_doc.add_table(rows=1, cols=4); table_mk.style = 'Table Grid'
                h_mk = table_mk.rows[0].cells
                h_mk[0].text = 'Câu'; h_mk[1].text = 'Đáp án'; h_mk[2].text = 'Câu'; h_mk[3].text = 'Đáp án'
                for i in range(1, q_count, 2):
                    r_mk = table_mk.add_row().cells
                    r_mk[0].text = str(i); r_mk[1].text = ans_key.get(i, "")
                    if i + 1 < q_count: r_mk[2].text = str(i + 1); r_mk[3].text = ans_key.get(i + 1, "")
            except Exception: pass
            
            master_key_doc.add_paragraph("📍 Nguồn gốc câu hỏi:").bold = True
            for info in source_info: master_key_doc.add_paragraph(info, style='List Bullet')
            master_key_doc.add_page_break()

    zip_buffer.seek(0)
    master_buf = io.BytesIO(); master_key_doc.save(master_buf); master_buf.seek(0)
    return zip_buffer, master_buf