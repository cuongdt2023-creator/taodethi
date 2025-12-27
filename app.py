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
    .file-box { border: 1px solid #e2e8f0; padding: 10px; border-radius: 8px; background: #f8fafc; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# ==================== LOGIC XỬ LÝ CHUYÊN SÂU ====================

def split_docx_into_questions(file_bytes):
    """
    Tách file gốc thành từng file Document nhỏ cho mỗi câu hỏi.
    Cách này giúp Composer giữ được toàn bộ Media (hình, công thức).
    """
    source_stream = io.BytesIO(file_bytes)
    source_doc = Document(source_stream)
    bank = {"P1": [], "P2": [], "P3": []}
    
    current_part = "P1"
    # Dùng để chứa các câu hỏi tạm thời dưới dạng Document riêng biệt
    temp_elements = []
    
    for p in source_doc.paragraphs:
        txt = p.text.upper()
        # Chuyển phần khi gặp từ khóa
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"

        # Nếu gặp chữ "Câu", bắt đầu một file Document mới cho câu đó
        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
            if temp_elements:
                # Gói các đoạn văn trước đó thành 1 file Word ảo
                q_doc = Document(io.BytesIO(file_bytes)) # Copy toàn bộ template gốc
                for para in q_doc.paragraphs: # Xóa sạch chỉ để lại khung
                    para._element.getparent().remove(para._element)
                
                # Thêm nội dung câu hỏi vào file ảo này
                for elem in temp_elements:
                    new_p = q_doc.add_paragraph()
                    new_p._element.getparent().replace(new_p._element, elem._element)
                
                bank[current_part_at_start].append(q_doc)
            
            temp_elements = [p]
            current_part_at_start = current_part
        elif temp_elements:
            temp_elements.append(p)

    return bank

# ==================== GIAO DIỆN CHÍNH ====================

st.markdown("<h1 class='main-header'>🎯 Hệ Thống Gộp Đề Bảo Toàn Hình Ảnh & Công Thức</h1>", unsafe_allow_html=True)

files = st.file_uploader("1. Tải các file đề chủ đề (.docx)", type="docx", accept_multiple_files=True)

if files:
    if 'bank' not in st.session_state or len(st.session_state.bank) != len(files):
        with st.spinner("Đang trích xuất dữ liệu thông minh..."):
            st.session_state.bank = {f.name: split_docx_into_questions(f.read()) for f in files}

    configs = {}
    st.subheader("2. Thiết lập số câu")
    cols = st.columns(len(files))
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]:
            st.markdown(f"<div class='file-box'>📂 <b>{fname[:15]}...</b></div>", unsafe_allow_html=True)
            configs[fname] = {
                "P1": st.number_input(f"P1 (Câu)", 0, 50, 0, key=f"p1_{fname}"),
                "P2": st.number_input(f"P2 (Câu)", 0, 50, 0, key=f"p2_{fname}"),
                "P3": st.number_input(f"P3 (Câu)", 0, 50, 0, key=f"p3_{fname}")
            }

    if st.button("🚀 XUẤT ĐỀ THI TỔNG HỢP CHUẨN", type="primary", use_container_width=True):
        # Tạo file đích giữ nguyên Section Properties (lề trang) của file đầu tiên
        final_doc = Document(io.BytesIO(files[0].getvalue()))
        for p in final_doc.paragraphs:
            final_doc._element.body.remove(p._element)
        
        composer = Composer(final_doc)
        
        titles = {
            "P1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
            "P2": "PHẦN II. Câu trắc nghiệm đúng sai.",
            "P3": "PHẦN III. Câu trắc nghiệm trả lời ngắn."
        }

        for p_key in ["P1", "P2", "P3"]:
            selected_docs = []
            for fname, cfg in configs.items():
                pool = st.session_state.bank[fname][p_key]
                num = min(cfg[p_key], len(pool))
                if num > 0:
                    selected_docs.extend(random.sample(pool, num))
            
            if selected_docs:
                # Thêm tiêu đề Phần
                t_para = final_doc.add_paragraph()
                t_para.add_run(titles[p_key]).bold = True
                
                random.shuffle(selected_docs)
                for idx, q_doc in enumerate(selected_docs):
                    # Đánh lại số câu trực tiếp trong Document tạm
                    for p in q_doc.paragraphs:
                        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
                            p.text = re.sub(r'^Câu\s*\d+', f"Câu {idx+1}", p.text, flags=re.IGNORECASE)
                            break
                    
                    # Gộp file
                    composer.append(q_doc)

        out_io = io.BytesIO()
        final_doc.save(out_io)
        st.success("✅ Đề thi đã sẵn sàng!")
        st.download_button("📥 Tải đề ngay", out_io.getvalue(), "De_Tong_Hop_Chuan.docx")

st.info("💡 Lưu ý: Hãy đảm bảo bạn đã thêm 'docxcompose' và 'python-docx' vào file requirements.txt.")
