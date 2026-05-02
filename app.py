import streamlit as st
import io
import os
import re
from docx_processor import parse_and_build_pools, generate_preview, generate_mixed_exams

if 'preview_data' not in st.session_state: st.session_state.preview_data = None
if 'parsed_pools' not in st.session_state: st.session_state.parsed_pools = None
if 'parsed_structure' not in st.session_state: st.session_state.parsed_structure = None
if 'parsed_title_to_part' not in st.session_state: st.session_state.parsed_title_to_part = None
if 'parsed_outro' not in st.session_state: st.session_state.parsed_outro = None
if 'merged_bytes' not in st.session_state: st.session_state.merged_bytes = None

def clear_preview():
    st.session_state.preview_data = None
    st.session_state.parsed_pools = None

def parse_manual_answers(text, type_):
    ans_dict = {}
    if not text.strip(): return ans_dict
    
    if type_ == "mcq":
        matches = re.finditer(r'(\d+)[\.\:\s]*([A-D])', text, re.IGNORECASE)
        for m in matches:
            ans_dict[int(m.group(1))] = m.group(2).upper()
            
    elif type_ == "tf":
        matches = re.finditer(r'(\d+)[\.\:\s\(\[\{]*([ĐS])[\,\.\s]*([ĐS])[\,\.\s]*([ĐS])[\,\.\s]*([ĐS])', text, re.IGNORECASE)
        for m in matches:
            ans_dict[int(m.group(1))] = [m.group(2).upper(), m.group(3).upper(), m.group(4).upper(), m.group(5).upper()]
            
    elif type_ == "sa":
        matches = re.finditer(r'(\d+)[\.\:]\s*([^\|]+?)(?=\s*\|\s*\d+[\.\:]|\s*,\s*\d+[\.\:]|$)', text)
        for m in matches:
            ans_val = re.sub(r',$', '', m.group(2)).strip()
            ans_dict[int(m.group(1))] = ans_val
            
    return ans_dict

st.set_page_config(page_title="Hệ Thống Ngân Hàng Đề", page_icon="📚", layout="wide")

if "zip_tests" not in st.session_state: st.session_state.zip_tests = None
if "master_key" not in st.session_state: st.session_state.master_key = None

with st.sidebar:
    st.header("⚙️ Cấu Hình Hệ Thống")
    
    st.subheader("1. Thông tin Tiêu đề (Header)")
    use_header = st.checkbox("Thêm bảng tiêu đề", value=True)
    if use_header:
        # Đã cập nhật thông tin mặc định theo yêu cầu của thầy Nguyễn Thiện
        ubnd = st.text_input("UBND/SỞ GD&ĐT", value="UBND XÃ VĨNH LẠI")
        truong = st.text_input("Trường", value="TRƯỜNG THCS ỨNG HÒE")
        ky_thi = st.text_input("Kỳ Thi", value="ĐỀ KIỂM TRA GIỮA KỲ I")
        mon_thi = st.text_input("Môn Thi", value="KHTN 7")
        thoi_gian = st.text_input("Thời gian", value="90 phút")
        nam_hoc = st.text_input("Năm học", value="2025 - 2026")

    st.subheader("2. Tùy chọn Trộn & Lấy Ngẫu Nhiên")
    num_versions = st.number_input("Số lượng đề mới muốn tạo:", min_value=1, max_value=24, value=4, step=1)
    q_mode_label = st.radio(
        "Chế độ sắp xếp Câu hỏi:",
        options=["🔀 Trộn ngẫu nhiên", "🔄 Đảo ngược thứ tự (VD: 12 → 1)", "📌 Giữ nguyên thứ tự gốc"],
        index=0
    )
    q_sort_mode_map = {
        "🔀 Trộn ngẫu nhiên": "random",
        "🔄 Đảo ngược thứ tự (VD: 12 → 1)": "reverse",
        "📌 Giữ nguyên thứ tự gốc": "keep"
    }
    q_sort_mode = q_sort_mode_map[q_mode_label]
    shuffle_a = st.checkbox("Trộn vị trí Đáp án (A,B,C,D)", value=True)
    
    # --- TÍNH NĂNG CHIA ĐỀU CÂU HỎI ---
    even_split = st.checkbox("⚖️ Tỉ lệ chia đều câu hỏi từ các file nguồn", value=True, help="Nếu bật, hệ thống sẽ lấy số lượng câu hỏi bằng nhau từ các đề bạn tải lên. Nếu 1 file thiếu, sẽ tự động bốc bù từ file khác.")
    keep_color = st.checkbox("🖍️ Giữ nguyên màu đỏ/gạch chân đáp án trong đề (Bản Giáo viên)", value=False, help="Bật tùy chọn này nếu bạn muốn xuất ra đề thi dành cho giáo viên (vẫn hiển thị đáp án đúng màu đỏ).")
    
    st.info("💡 Lưu ý: Hệ thống sẽ lấy cấu trúc (số câu mỗi Phần) theo File đầu tiên bạn tải lên làm chuẩn.")
    
    # --- THÔNG TIN TÁC GIẢ ---
    st.markdown("---")
    st.success("👨‍💻 **Tác giả:** Nguyễn Thiện\n\n📞 **SĐT/Zalo:** 0988250112")

st.markdown("<h1 style='text-align: center; color: #1f4e79;'>📚 HỆ THỐNG TỔNG HỢP & TRỘN ĐỀ ĐA NGUỒN</h1>", unsafe_allow_html=True)
st.markdown("---")

# --- KHU VỰC HƯỚNG DẪN VÀ TẢI FILE MẪU ---
with st.expander("📖 HƯỚNG DẪN SỬ DỤNG & TẢI FILE MẪU", expanded=True):
    col_text, col_btn = st.columns([3, 1])
    with col_text:
        st.markdown("""
        **Quy tắc CẬP NHẬT cực kỳ nhàn rỗi và CHỐNG LỖI 100%:**
        - **Phần I & Phần II:** **CHỈ CẦN** Gạch chân hoặc Tô đỏ ký hiệu chữ cái đầu tiên (VD: <u>A.</u>, <u>a)</u>) của phương án ĐÚNG. **Tuyệt đối không tô đỏ công thức Toán (MathType)** để tránh lỗi.
        - **Phần III (Trả lời ngắn):** Xuống dòng gõ `Đáp án: [Kết quả]` hoặc `Đáp số: [Kết quả]`. KHÔNG CẦN tô màu đỏ. Hệ thống sẽ tự bốc đáp án và tự động xóa sạch dòng này trên đề in của học sinh.
        - **Nhóm dùng chung:** Kẹp các câu hỏi chung 1 hình ảnh/đoạn văn giữa thẻ **@BẮT ĐẦU DÙNG CHUNG@** và **@KẾT THÚC DÙNG CHUNG@**.
        """, unsafe_allow_html=True)
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        sample_file_path = "ĐỀ MẪU TRỘN ĐỀ ĐA NĂNG.docx"
        if os.path.exists(sample_file_path):
            with open(sample_file_path, "rb") as f:
                st.download_button("📥 TẢI ĐỀ MẪU CHUẨN", f.read(), "De_Mau_Chuan.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
        else:
            st.warning(f"💡 Chưa tìm thấy file mẫu '{sample_file_path}' trong thư mục.")

st.markdown("---")

# --- KHU VỰC TẢI LÊN NHIỀU FILE ---
uploaded_files = st.file_uploader("📁 Tải lên các file Word đề gốc (Bạn có thể bôi đen chọn nhiều file)", type=['docx'], accept_multiple_files=True, on_change=clear_preview)

if not uploaded_files:
    st.info("👉 Vui lòng tải lên ít nhất 1 file đề gốc (.docx) để bắt đầu.")
else:
    st.success(f"✅ Đã tải lên {len(uploaded_files)} file nguồn. Bấm nút dưới đây để trộn!")
    
    with st.expander("✍️ NHẬP ĐÁP ÁN THỦ CÔNG (Tùy chọn - Dùng thay thế cho việc tô đỏ trong Word)", expanded=False):
        st.markdown("**Hướng dẫn:** Nếu bạn không muốn mở file Word để tô đỏ đáp án, bạn có thể nhập trực tiếp chuỗi đáp án vào đây. Hệ thống sẽ ưu tiên dùng đáp án này.")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.text_area("Phần I: Trắc nghiệm (VD: 1A 2B 3C)", key="manual_p1", height=150)
        with col_m2:
            st.text_area("Phần II: Đúng/Sai (VD: 1: Đ S Đ S)", key="manual_p2", height=150)
        with col_m3:
            st.text_area("Phần III: Trả lời ngắn (VD: 1: 15 | 2: 20)", key="manual_p3", height=150)
            
    if st.button("🔍 Phân tích & Chuẩn hóa đề", type="primary"):
        with st.spinner('Đang phân tích cấu trúc đề và bóc tách đáp án...'):
            try:
                manual_answers = {
                    "part_1": parse_manual_answers(st.session_state.get("manual_p1", ""), "mcq"),
                    "part_2": parse_manual_answers(st.session_state.get("manual_p2", ""), "tf"),
                    "part_3": parse_manual_answers(st.session_state.get("manual_p3", ""), "sa")
                }
                
                pools, structure, title_to_part, outro, merged_bytes = parse_and_build_pools(uploaded_files)
                preview_data = generate_preview(pools, structure, title_to_part, manual_answers)
                
                st.session_state.parsed_pools = pools
                st.session_state.parsed_structure = structure
                st.session_state.parsed_title_to_part = title_to_part
                st.session_state.parsed_outro = outro
                st.session_state.merged_bytes = merged_bytes
                st.session_state.preview_data = preview_data
                
                st.success("✅ Phân tích xong! Vui lòng kiểm tra lại đáp án bên dưới.")
                st.info(f"DEBUG: Số phần trong đề: {len(structure)} | Preview size: {len(preview_data) if preview_data else 0}")
            except Exception as e:
                import traceback
                traceback.print_exc()
                st.error(f"Có lỗi xảy ra: {e}")

    if st.session_state.preview_data:
        st.markdown("### 📋 XEM TRƯỚC (PREVIEW) NỘI DUNG")
        st.info("💡 Lưu ý: Đây chỉ là giao diện hiển thị chữ thô để bạn kiểm tra nhanh việc nhận diện Đáp Án. File Word gốc chứa công thức/hình ảnh vẫn được bảo toàn 100% khi trộn đề.")
        
        for part in st.session_state.preview_data:
            with st.expander(f"Phần: {part['title']} ({len(part['questions'])} câu hỏi/nhóm)", expanded=True):
                for q in part['questions']:
                    st.markdown(f"**Câu hỏi (Gốc ID: {q.get('orig_idx', '?')}):**")
                    st.text(q['text'])
                    st.markdown(f"👉 **Đáp án hệ thống nhận diện:** `{q['answer']}`")
                    st.markdown("---")
                    
        if st.button("🔀 Bắt Đầu Tổng Hợp & Trộn Đề", type="primary"):
            with st.spinner('Đang trộn đề và xuất file...'):
                try:
                    manual_answers = {
                        "part_1": parse_manual_answers(st.session_state.get("manual_p1", ""), "mcq"),
                        "part_2": parse_manual_answers(st.session_state.get("manual_p2", ""), "tf"),
                        "part_3": parse_manual_answers(st.session_state.get("manual_p3", ""), "sa")
                    }
                    
                    zip_result, master_key_result = generate_mixed_exams(
                        global_pools=st.session_state.parsed_pools,
                        structure=st.session_state.parsed_structure,
                        title_to_part=st.session_state.parsed_title_to_part,
                        outro=st.session_state.parsed_outro,
                        merged_bytes=st.session_state.merged_bytes,
                        num_versions=num_versions,
                        q_sort_mode=q_sort_mode,
                        shuffle_a=shuffle_a,
                        even_split=even_split,
                        keep_color=keep_color,
                        manual_answers=manual_answers
                    )
                    st.session_state.zip_tests = zip_result
                    st.session_state.master_key = master_key_result
                    st.success("🎉 Đã tạo xong! Mời bạn tải file ở các nút bên dưới.")
                except Exception as e:
                    st.error(f"Có lỗi xảy ra trong quá trình trộn đề: {e}")

# --- KHU VỰC TẢI KẾT QUẢ ---
if st.session_state.zip_tests and st.session_state.master_key:
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label=f"📦 1. TẢI BỘ {num_versions} ĐỀ ĐÃ TRỘN (.ZIP)",
            data=st.session_state.zip_tests,
            file_name="Bo_De_Thi_Tong_Hop.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )
    with col2:
        st.download_button(
            label="📑 2. TẢI BẢNG ĐÁP ÁN & TRUY VẾT NGUỒN (.DOCX)",
            data=st.session_state.master_key,
            file_name="Dap_An_Va_Truy_Vet_Nguon.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )