import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Hàm cắt lấy 1 đoạn nội dung từ file gốc mà không làm hỏng MathType/Ảnh
def extract_section_safe(source_bytes, start_idx, end_idx):
    doc = Document(io.BytesIO(source_bytes))
    # Xóa tất cả các đoạn văn không nằm trong khoảng cần lấy
    for i in range(len(doc.paragraphs) - 1, -1, -1):
        if not (start_idx <= i < end_idx):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
    return doc

def analyze_file(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    questions = {"P1": [], "P2": [], "P3": []}
    current_part = "P1"
    start_idx = -1
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().upper()
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"
        
        if re.match(r'^Câu\s*\d+', p.text, re.I):
            if start_idx != -1:
                questions[prev_part].append((start_idx, i))
            start_idx = i
            prev_part = current_part
            
    if start_idx != -1:
        questions[prev_part].append((start_idx, len(doc.paragraphs)))
    return questions

# --- GIAO DIỆN ---
st.title("🚀 Tạo Đề Thi Tổng Hợp (Bảo toàn MathType)")

files = st.file_uploader("Tải các file ngân hàng câu hỏi (.docx)", type="docx", accept_multiple_files=True)

if files:
    all_data = {}
    for f in files:
        b = f.read()
        all_data[f.name] = {"bytes": b, "struct": analyze_file(b)}
    
    st.subheader("Cấu hình số câu cần lấy:")
    selected_config = {}
    for fname, data in all_data.items():
        with st.expander(f"File: {fname}"):
            c1, c2, c3 = st.columns(3)
            q1 = c1.number_input(f"Phần 1 (Max {len(data['struct']['P1'])})", 0, len(data['struct']['P1']), 0, key=f"n1_{fname}")
            q2 = c2.number_input(f"Phần 2 (Max {len(data['struct']['P2'])})", 0, len(data['struct']['P2']), 0, key=f"n2_{fname}")
            q3 = c3.number_input(f"Phần 3 (Max {len(data['struct']['P3'])})", 0, len(data['struct']['P3']), 0, key=f"n3_{fname}")
            selected_config[fname] = {"P1": q1, "P2": q2, "P3": q3}

    if st.button("Tạo Đề Mới"):
        # Tạo file tổng (Master)
        master_doc = Document(io.BytesIO(list(all_data.values())[0]["bytes"]))
        for p in master_doc.paragraphs: master_doc._element.body.remove(p._element)
        composer = Composer(master_doc)
        
        current_q_num = 1
        for part in ["P1", "P2", "P3"]:
            # Thêm tiêu đề phần
            master_doc.add_paragraph(f"--- {part} ---").bold = True
            
            for fname, cfg in selected_config.items():
                num_to_take = cfg[part]
                if num_to_take > 0:
                    indices = random.sample(all_data[fname]["struct"][part], num_to_take)
                    for start, end in indices:
                        # Trích xuất "nguyên khối" để giữ MathType/Ảnh
                        q_doc = extract_section_safe(all_data[fname]["bytes"], start, end)
                        
                        # Đánh lại số câu
                        for p in q_doc.paragraphs:
                            if re.match(r'^Câu\s*\d+', p.text, re.I):
                                p.text = re.sub(r'^Câu\s*\d+', f"Câu {current_q_num}", p.text, flags=re.I)
                                current_q_num += 1
                                break
                        composer.append(q_doc)
        
        out = io.BytesIO()
        master_doc.save(out)
        st.success("Đã tạo đề thành công!")
        st.download_button("Tải Đề Tổng Hợp", out.getvalue(), "De_Tong_Hop.docx")
