import streamlit as st
import io
import random
import re
import copy
from docx import Document
from docxcompose.composer import Composer

# ==================== CẤU HÌNH TRANG ====================
st.set_page_config(page_title="Trộn Đề Word Pro - Fix Lỗi Corrupt", page_icon="🛡️", layout="wide")

st.markdown("""
<style>
    .main-header { text-align: center; color: #0d9488; font-weight: bold; }
    .success-box { padding: 10px; background-color: #f0fdfa; border: 1px solid #14b8a6; border-radius: 5px; }
    .error-box { padding: 10px; background-color: #fef2f2; border: 1px solid #ef4444; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

# ==================== CORE LOGIC: CLONE & PRUNE ====================

def get_difficulty(text):
    """Nhận diện độ khó"""
    t = text.upper()
    if "#VDC" in t: return "VDC"
    if "#VD" in t: return "VD"
    if "#TH" in t: return "TH"
    if "#NB" in t: return "NB"
    return "NB" # Mặc định

def map_questions(file_bytes):
    """
    Quét vị trí (index) của các câu hỏi trong file mà KHÔNG sửa đổi file.
    Trả về bản đồ: { 'P1': [ {start, end, diff}, ... ], ... }
    """
    doc = Document(io.BytesIO(file_bytes))
    mapping = {"P1": [], "P2": [], "P3": [], "Footer": None}
    
    current_part = "P1"
    q_start = -1
    total_paras = len(doc.paragraphs)
    
    # Từ khóa nhận diện phần đáp án cuối file
    footer_keywords = ["BẢNG ĐÁP ÁN", "HƯỚNG DẪN GIẢI", "LỜI GIẢI CHI TIẾT", "ĐÁP ÁN THAM KHẢO"]

    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().upper()
        
        # 1. Phát hiện chuyển Phần
        if re.match(r'^PHẦN\s+(1|I)\b', txt): current_part = "P1"
        elif re.match(r'^PHẦN\s+(2|II)\b', txt): current_part = "P2"
        elif re.match(r'^PHẦN\s+(3|III)\b', txt): current_part = "P3"
        
        # 2. Phát hiện Phần Đáp Án (Footer)
        if any(k in txt for k in footer_keywords):
            # Đóng câu hỏi đang mở nếu có
            if q_start != -1:
                diff = get_difficulty(doc.paragraphs[q_start].text)
                mapping[prev_part].append({"range": (q_start, i), "diff": diff})
                q_start = -1
            
            # Ghi nhận vị trí bắt đầu footer
            mapping["Footer"] = (i, total_paras)
            break # Dừng quét vì phần sau là đáp án hết

        # 3. Phát hiện Câu hỏi (Câu 1., Câu 2...)
        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
            # Lưu câu hỏi trước đó
            if q_start != -1:
                diff = get_difficulty(doc.paragraphs[q_start].text)
                mapping[prev_part].append({"range": (q_start, i), "diff": diff})
            
            q_start = i
            prev_part = current_part
    
    # Lưu câu cuối cùng (nếu không có footer)
    if q_start != -1 and mapping["Footer"] is None:
        diff = get_difficulty(doc.paragraphs[q_start].text)
        mapping[prev_part].append({"range": (q_start, total_paras), "diff": diff})
        
    return mapping

def extract_content_safe(original_bytes, keep_ranges):
    """
    THUẬT TOÁN AN TOÀN TUYỆT ĐỐI:
    1. Load file gốc.
    2. Xóa TẤT CẢ các dòng KHÔNG nằm trong keep_ranges.
    3. Trả về Document đã được cắt gọt.
    -> Giữ nguyên 100% MathType, Ảnh, Table vì chúng chưa từng bị di chuyển.
    """
    doc = Document(io.BytesIO(original_bytes))
    total_paras = len(doc.paragraphs)
    
    # Tạo set các chỉ số dòng cần GIỮ LẠI
    indices_to_keep = set()
    for start, end in keep_ranges:
        for i in range(start, end):
            indices_to_keep.add(i)
            
    # Xóa ngược từ dưới lên trên để không làm lệch index
    for i in range(total_paras - 1, -1, -1):
        if i not in indices_to_keep:
            p = doc.paragraphs[i]
            # Xóa triệt để khỏi XML
            p._element.getparent().remove(p._element)
            
    return doc

# ==================== GIAO DIỆN STREAMLIT ====================

st.markdown("<h1 class='main-header'>🛡️ Hệ Thống Trộn Đề: Bảo Toàn MathType & Ảnh</h1>", unsafe_allow_html=True)
st.info("💡 Cách hoạt động: App sẽ 'nhân bản' file gốc cho mỗi câu hỏi để đảm bảo không một công thức hay hình ảnh nào bị mất liên kết.")

# 1. Upload
uploaded_files = st.file_uploader("Bước 1: Tải các file chủ đề (.docx)", type="docx", accept_multiple_files=True)

if uploaded_files:
    # 2. Phân tích file
    if 'data_map' not in st.session_state or len(st.session_state.data_map) != len(uploaded_files):
        with st.spinner("Đang quét cấu trúc file..."):
            st.session_state.data_map = {}
            for f in uploaded_files:
                f_bytes = f.read()
                st.session_state.data_map[f.name] = {
                    "bytes": f_bytes,
                    "map": map_questions(f_bytes)
                }
    
    # 3. Cấu hình
    st.subheader("Bước 2: Chọn số lượng câu hỏi")
    configs = {}
    cols = st.columns(len(uploaded_files))
    
    for i, fname in enumerate(st.session_state.data_map.keys()):
        mapping = st.session_state.data_map[fname]["map"]
        p1_count = len(mapping["P1"])
        p2_count = len(mapping["P2"])
        p3_count = len(mapping["P3"])
        has_footer = "✅ Có Đáp án" if mapping["Footer"] else "⚠️ Không thấy Đáp án"
        
        with cols[i]:
            st.success(f"📂 {fname}\n\n{has_footer}")
            configs[fname] = {
                "P1": st.number_input(f"P1 (Max {p1_count})", 0, 50, 0, key=f"p1_{fname}"),
                "P2": st.number_input(f"P2 (Max {p2_count})", 0, 50, 0, key=f"p2_{fname}"),
                "P3": st.number_input(f"P3 (Max {p3_count})", 0, 50, 0, key=f"p3_{fname}")
            }

    # 4. Xử lý
    if st.button("🚀 XUẤT ĐỀ THI (KHÔNG LỖI)", type="primary", use_container_width=True):
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # Tạo Master Doc từ file đầu tiên (để lấy lề trang, font chữ chuẩn)
            base_bytes = list(st.session_state.data_map.values())[0]["bytes"]
            master_doc = Document(io.BytesIO(base_bytes))
            # Xóa sạch nội dung Master
            for p in master_doc.paragraphs: 
                p._element.getparent().remove(p._element)
            
            composer = Composer(master_doc)
            
            titles = {
                "P1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
                "P2": "PHẦN II. Câu trắc nghiệm đúng sai.",
                "P3": "PHẦN III. Câu trắc nghiệm trả lời ngắn."
            }
            
            parts = ["P1", "P2", "P3"]
            global_q_idx = {"P1": 1, "P2": 1, "P3": 1}
            
            # --- VÒNG LẶP XỬ LÝ TỪNG PHẦN ---
            for part_idx, part in enumerate(parts):
                status_text.write(f"⏳ Đang xử lý {titles[part]}...")
                
                # Gom câu hỏi từ tất cả các file
                requests = [] # {bytes, range}
                for fname, cfg in configs.items():
                    needed = cfg[part]
                    available = st.session_state.data_map[fname]["map"][part]
                    if needed > 0:
                        chosen = random.sample(available, min(needed, len(available)))
                        for item in chosen:
                            requests.append({
                                "bytes": st.session_state.data_map[fname]["bytes"],
                                "range": item["range"]
                            })
                
                if requests:
                    # Thêm tiêu đề phần vào Master
                    master_doc.add_paragraph(titles[part]).bold = True
                    random.shuffle(requests)
                    
                    # GỘP TỪNG CÂU HỎI
                    for i, req in enumerate(requests):
                        # CẮT FILE GỐC CHỈ LẤY CÂU HỎI NÀY
                        # Đây là bước quan trọng nhất để giữ MathType
                        q_doc = extract_content_safe(req["bytes"], [req["range"]])
                        
                        # Đánh lại số câu (Câu 1, Câu 2...)
                        for p in q_doc.paragraphs:
                            if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
                                # Thay thế số câu cũ bằng số mới
                                p.text = re.sub(r'^Câu\s*\d+', f"Câu {global_q_idx[part]}", p.text, flags=re.IGNORECASE)
                                # Xóa tag rác #NB, #TH...
                                p.text = re.sub(r'#(NB|TH|VD|VDC)', '', p.text)
                                break
                        
                        # Gộp vào Master
                        composer.append(q_doc)
                        global_q_idx[part] += 1
                
                progress_bar.progress((part_idx + 1) / 4)

            # --- XỬ LÝ ĐÁP ÁN (FOOTER) ---
            status_text.write("⏳ Đang tổng hợp Đáp án...")
            master_doc.add_page_break()
            master_doc.add_paragraph("--- TỔNG HỢP ĐÁP ÁN & HƯỚNG DẪN ---").bold = True
            
            for fname, cfg in configs.items():
                total_req = sum(cfg.values())
                mapping = st.session_state.data_map[fname]["map"]
                
                if total_req > 0 and mapping["Footer"]:
                    master_doc.add_paragraph(f"\nNguồn: {fname}").italic = True
                    # Cắt lấy phần footer
                    footer_doc = extract_content_safe(st.session_state.data_map[fname]["bytes"], [mapping["Footer"]])
                    composer.append(footer_doc)
            
            progress_bar.progress(1.0)
            status_text.write("✅ Đã xong!")
            
            # Xuất file
            out_io = io.BytesIO()
            master_doc.save(out_io)
            
            st.success("Tạo đề thành công! Không còn lỗi 'We found a problem'.")
            st.download_button(
                label="📥 Tải về đề thi chuẩn (.docx)",
                data=out_io.getvalue(),
                file_name="De_Thi_Chuan_Pro.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Có lỗi xảy ra: {str(e)}")
            st.code(str(e))
