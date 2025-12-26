import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Cấu hình giao diện
st.set_page_config(page_title="AIOMT - Gộp Đề Chuẩn", layout="wide")

def get_difficulty(para_text):
    t = para_text.upper()
    for tag in ["#VDC", "#VD", "#TH", "#NB"]:
        if tag in t: return tag[1:]
    return "NB"

def parse_docx_to_bank(file_bytes):
    """Phân loại câu hỏi vào 3 kho P1, P2, P3"""
    doc = Document(io.BytesIO(file_bytes))
    bank = {"P1": [], "P2": [], "P3": []}
    curr_p = "P1"
    curr_q = []
    
    for p in doc.paragraphs:
        txt = p.text.upper()
        if "PHẦN 1" in txt: curr_p = "P1"
        elif "PHẦN 2" in txt: curr_p = "P2"
        elif "PHẦN 3" in txt: curr_p = "P3"
        
        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
            if curr_q:
                diff = get_difficulty(curr_q[0].text)
                bank[prev_p].append({"paras": curr_q, "diff": diff})
            curr_q = [p]
            prev_p = curr_p
        elif curr_q:
            curr_q.append(p)
            
    if curr_q:
        bank[prev_p].append({"paras": curr_q, "diff": get_difficulty(curr_q[0].text)})
    return bank

st.title("🧩 Gộp Đề Tổng Hợp - Fix Lỗi Content")

files = st.file_uploader("Tải các file chủ đề", type="docx", accept_multiple_files=True)

if files:
    # Tránh lỗi IndexError bằng cách reset bank khi số lượng file thay đổi
    if 'bank' not in st.session_state or len(st.session_state.bank) != len(files):
        st.session_state.bank = {f.name: parse_docx_to_bank(f.read()) for f in files}

    configs = {}
    cols = st.columns(len(files))
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]: # Fix lỗi IndexError tại dòng 65
            st.info(f"📂 {fname[:15]}...")
            p1 = st.number_input(f"P1", 0, 50, 0, key=f"p1_{fname}")
            p2 = st.number_input(f"P2", 0, 50, 0, key=f"p2_{fname}")
            p3 = st.number_input(f"P3", 0, 50, 0, key=f"p3_{fname}")
            configs[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🚀 TẠO ĐỀ THI TỔNG HỢP", type="primary"):
        # Tạo file đích dựa trên template của file đầu tiên
        template_doc = Document(io.BytesIO(files[0].getvalue()))
        # Xóa hết nội dung cũ trong template
        for p in template_doc.paragraphs:
            p._element.getparent().remove(p._element)
            
        final_composer = Composer(template_doc)
        
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
                if num > 0:
                    selected.extend(random.sample(pool, num))
            
            if selected:
                # Thêm tiêu đề phần
                template_doc.add_paragraph(titles[p_key]).bold = True
                random.shuffle(selected)
                
                for idx, q_data in enumerate(selected):
                    # Tạo một doc tạm cho từng câu để dùng Composer
                    q_doc = Document()
                    for j, p_origin in enumerate(q_data["paras"]):
                        new_p = q_doc.add_paragraph()
                        text = p_origin.text
                        if j == 0: # Đánh lại số câu và xóa tag
                            text = re.sub(r'^Câu\s*\d+', f"Câu {idx+1}", text, flags=re.IGNORECASE)
                            text = re.sub(r'#(NB|TH|VD|VDC)', '', text)
                        new_p.text = text
                    
                    final_composer.append(q_doc)

        output = io.BytesIO()
        template_doc.save(output)
        st.success("🎉 Đã gộp đề thành công!")
        st.download_button("📥 Tải đề chuẩn", output.getvalue(), "De_Tong_Hop.docx")
