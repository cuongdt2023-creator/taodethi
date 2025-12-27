import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# ==================== CẤU HÌNH GIAO DIỆN ====================
st.set_page_config(page_title="AIOMT Premium - Gộp Đề Chuẩn", page_icon="🎯", layout="wide")

st.markdown("""
<style>
    .main-header { text-align: center; color: #0d9488; }
    .file-box { border: 1px solid #e2e8f0; padding: 10px; border-radius: 8px; background: #f8fafc; }
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC XỬ LÝ WORD ====================

def get_difficulty(para_text):
    """Nhận diện độ khó dựa trên thẻ #NB, #TH..."""
    t = para_text.upper()
    for tag in ["#VDC", "#VD", "#TH", "#NB"]:
        if tag in t: return tag[1:]
    return "NB"

def split_docx_to_bank(file_bytes):
    """
    Tách file word thành các câu hỏi riêng biệt.
    Mỗi câu hỏi được lưu dưới dạng một Document tạm để giữ nguyên hình ảnh/công thức.
    """
    source_doc = Document(io.BytesIO(file_bytes))
    bank = {"P1": [], "P2": [], "P3": []}
    
    current_part = "P1"
    questions = []
    temp_q_elements = []

    for p in source_doc.paragraphs:
        txt = p.text.upper()
        # Nhận diện chuyển phần
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"

        # Nhận diện bắt đầu câu hỏi mới
        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
            if temp_q_elements:
                # Lưu câu hỏi trước đó
                q_doc = Document() # Tạo doc mới cho câu hỏi này
                # Copy các đoạn văn vào doc tạm
                for elem in temp_q_elements:
                    new_p = q_doc.add_paragraph(elem.text, elem.style)
                
                diff = get_difficulty(temp_q_elements[0].text)
                bank[current_part_at_start].append({"doc": q_doc, "diff": diff})
            
            temp_q_elements = [p]
            current_part_at_start = current_part
        elif temp_q_elements:
            temp_q_elements.append(p)

    # Lưu câu cuối cùng
    if temp_q_elements:
        q_doc = Document()
        for elem in temp_q_elements: q_doc.add_paragraph(elem.text, elem.style)
        bank[current_part_at_start].append({"doc": q_doc, "diff": get_difficulty(temp_q_elements[0].text)})
    
    return bank

# ==================== GIAO DIỆN STREAMLIT ====================

st.markdown("<h1 class='main-header'>🎯 Hệ Thống Tạo Đề Tổng Hợp Master</h1>", unsafe_allow_html=True)

files = st.file_uploader("Bước 1: Tải các file chủ đề (Ngân hàng câu hỏi)", type="docx", accept_multiple_files=True)

if files:
    # Khởi tạo ngân hàng dữ liệu
    if 'bank' not in st.session_state or len(st.session_state.bank) != len(files):
        with st.spinner("Đang phân tích dữ liệu các file..."):
            st.session_state.bank = {f.name: split_docx_to_bank(f.read()) for f in files}

    st.subheader("Bước 2: Chọn số lượng câu hỏi từ mỗi chủ đề")
    configs = {}
    cols = st.columns(len(files))
    
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]:
            st.markdown(f"<div class='file-box'><b>📂 {fname[:20]}</b></div>", unsafe_allow_html=True)
            p1 = st.number_input(f"P1 (Câu)", 0, 50, 0, key=f"p1_{fname}")
            p2 = st.number_input(f"P2 (Câu)", 0, 50, 0, key=f"p2_{fname}")
            p3 = st.number_input(f"P3 (Câu)", 0, 50, 0, key=f"p3_{fname}")
            configs[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🚀 XUẤT ĐỀ THI TỔNG HỢP CHUẨN CẤU TRÚC", type="primary", use_container_width=True):
        try:
            # 1. Khởi tạo tài liệu đích từ file đầu tiên (để lấy định dạng trang/font)
            result_doc = Document(io.BytesIO(files[0].getvalue()))
            # Xóa hết nội dung cũ nhưng giữ lại Section (Lề, khổ giấy)
            for p in result_doc.paragraphs:
                p._element.getparent().remove(p._element)
            
            composer = Composer(result_doc)
            
            titles = {
                "P1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 12.",
                "P2": "PHẦN II. Câu trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu 4.",
                "P3": "PHẦN III. Câu trắc nghiệm trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu 6."
            }

            for p_key in ["P1", "P2", "P3"]:
                # Gom câu hỏi từ các file
                selected_pool = []
                for fname, cfg in configs.items():
                    pool = st.session_state.bank[fname][p_key]
                    num_to_take = min(cfg[p_key], len(pool))
                    if num_to_take > 0:
                        selected_pool.extend(random.sample(pool, num_to_take))
                
                if selected_pool:
                    # Chèn tiêu đề Phần
                    title_p = result_doc.add_paragraph()
                    run = title_p.add_run(titles[p_key])
                    run.bold = True
                    run.font.size = 14 * 12700 # Quy đổi sang DXA tương ứng size 14

                    random.shuffle(selected_pool)
                    
                    for idx, q_data in enumerate(selected_pool):
                        q_doc = q_data["doc"]
                        # Đánh lại số câu ở đoạn văn đầu tiên của mỗi câu hỏi
                        first_para = q_doc.paragraphs[0]
                        first_para.text = re.sub(r'^Câu\s*\d+', f"Câu {idx+1}", first_para.text, flags=re.IGNORECASE)
                        first_para.text = re.sub(r'#(NB|TH|VD|VDC)', '', first_para.text)
                        
                        # Dùng composer để gộp giữ nguyên hình ảnh/công thức
                        composer.append(q_doc)

            # Xu
