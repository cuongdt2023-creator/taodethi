import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Cấu hình giao diện
st.set_page_config(page_title="AIOMT - Bảo Toàn Công Thức & Hình Ảnh", layout="wide")

def get_difficulty(para_text):
    t = para_text.upper()
    for tag in ["#VDC", "#VD", "#TH", "#NB"]:
        if tag in t: return tag[1:]
    return "NB"

def split_docx_to_questions(file_bytes):
    """
    Tách file gốc thành từng câu hỏi. 
    Mỗi câu hỏi được lưu tạm dưới dạng một đối tượng Document riêng để bảo toàn Media.
    """
    source_doc = Document(io.BytesIO(file_bytes))
    bank = {"P1": [], "P2": [], "P3": []}
    
    current_part = "P1"
    questions_data = []
    temp_paras = []

    for p in source_doc.paragraphs:
        txt = p.text.upper()
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"

        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
            if temp_paras:
                # Tạo một document nhỏ chứa duy nhất câu hỏi này để giữ nguyên hình/ảnh
                q_doc = Document(io.BytesIO(file_bytes)) 
                # Xóa sạch mọi thứ trong q_doc, chỉ để lại các đoạn văn của câu hỏi này
                target_body = q_doc._element.body
                for child in list(target_body):
                    if child.tag.endswith('sectPr'): continue
                    target_body.remove(child)
                
                for para in temp_paras:
                    target_body.append(para._element)
                
                diff = get_difficulty(temp_paras[0].text)
                bank[start_part].append({"doc": q_doc, "diff": diff})
            
            temp_paras = [p]
            start_part = current_part
        elif temp_paras:
            temp_paras.append(p)

    # Lưu câu cuối
    if temp_paras:
        q_doc = Document(io.BytesIO(file_bytes))
        target_body = q_doc._element.body
        for child in list(target_body):
            if child.tag.endswith('sectPr'): continue
            target_body.remove(child)
        for para in temp_paras: target_body.append(para._element)
        bank[start_part].append({"doc": q_doc, "diff": get_difficulty(temp_paras[0].text)})

    return bank

st.title("🎯 Tạo Đề Tổng Hợp: Giữ Nguyên Hình Ảnh & Công Thức")

uploaded_files = st.file_uploader("Tải các file chủ đề (.docx)", type="docx", accept_multiple_files=True)

if uploaded_files:
    if 'bank' not in st.session_state:
        st.session_state.bank = {f.name: split_docx_to_questions(f.read()) for f in uploaded_files}

    st.subheader("Chọn số lượng câu hỏi từ mỗi file")
    configs = {}
    cols = st.columns(len(uploaded_files))
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]:
            st.info(f"📁 {fname}")
            p1 = st.number_input(f"P1", 0, 50, 0, key=f"p1_{fname}")
            p2 = st.number_input(f"P2", 0, 50, 0, key=f"p2_{fname}")
            p3 = st.number_input(f"P3", 0, 50, 0, key=f"p3_{fname}")
            configs[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🚀 XUẤT ĐỀ THI TỔNG HỢP", type="primary", use_container_width=True):
        # 1. Lấy file đầu tiên làm mẫu (Template) để giữ Margin, Font, Header/Footer
        main_doc = Document(io.BytesIO(uploaded_files[0].getvalue()))
        # Xóa sạch nội dung cũ trong body nhưng giữ lại sectPr (định dạng trang)
        body = main_doc._element.body
        for child in list(body):
            if not child.tag.endswith('sectPr'):
                body.remove(child)
        
        composer = Composer(main_doc)
        
        titles = {
            "P1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
            "P2": "PHẦN II. Câu trắc nghiệm đúng sai.",
            "P3": "PHẦN III. Câu trắc nghiệm trả lời ngắn."
        }

        for p_key in ["P1", "P2", "P3"]:
            selected = []
            for fname, cfg in configs.items():
                pool = st.session_state.bank[fname][p_key]
                num = min(cfg[p_key], len(pool))
                if num > 0: selected.extend(random.sample(pool, num))
            
            if selected:
                # Thêm tiêu đề phần vào main_doc
                title_para = main_doc.add_paragraph()
                run = title_para.add_run(titles[p_key])
                run.bold = True
                run.font.size = 14 * 12700 # Size 14

                random.shuffle(selected)
                for idx, q_data in enumerate(selected):
                    q_doc = q_data["doc"]
                    # Đánh lại số câu trong Document tạm
                    for p in q_doc.paragraphs:
                        if "Câu" in p.text:
                            p.text = re.sub(r'^Câu\s*\d+', f"Câu {idx+1}", p.text, flags=re.IGNORECASE)
                            p.text = re.sub(r'#(NB|TH|VD|VDC)', '', p.text)
                            break
                    # Dùng composer để gộp - Đây là bước giữ lại hình ảnh/công thức
                    composer.append(q_doc)

        output = io.BytesIO()
        main_doc.save(output)
        st.success("✅ Đã tạo đề thành công! Hình ảnh và công thức đã được bảo toàn.")
        st.download_button("📥 Tải đề chuẩn (.docx)", output.getvalue(), "De_Tong_Hop_Bao_Toan.docx")
