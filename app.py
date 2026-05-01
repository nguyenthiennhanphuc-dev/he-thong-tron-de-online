import streamlit as st
import io
import os
from docx_processor import process_and_shuffle_multi

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
uploaded_files = st.file_uploader("📁 Tải lên các file Word đề gốc (Bạn có thể bôi đen chọn nhiều file)", type=['docx'], accept_multiple_files=True)

if not uploaded_files:
    st.info("👉 Vui lòng tải lên ít nhất 1 file đề gốc (.docx) để bắt đầu.")
else:
    st.success(f"✅ Đã tải lên {len(uploaded_files)} file nguồn. Bấm nút dưới đây để trộn!")
    
    if st.button("Bắt đầu Tổng hợp & Trộn đề", type="primary"):
        with st.spinner('Đang phân tích XML, chia đều tỷ lệ câu hỏi và truy vết nguồn gốc...'):
            try:
                zip_result, master_key_result = process_and_shuffle_multi(
                    files=uploaded_files,
                    num_versions=num_versions,
                    q_sort_mode=q_sort_mode,
                    shuffle_a=shuffle_a,
                    even_split=even_split
                )
                st.session_state.zip_tests = zip_result
                st.session_state.master_key = master_key_result
                st.success("🎉 Đã tạo xong! Mời bạn tải file ở các nút bên dưới.")
            except Exception as e:
                st.error(f"Có lỗi xảy ra trong quá trình xử lý: {e}")

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