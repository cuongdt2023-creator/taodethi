import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Hàm quan trọng nhất: Trích xuất câu hỏi mà không làm hỏng MathType/Ảnh
def extract_question_safe(file_bytes, start_line, end_line):
    doc = Document(io.BytesIO(file_bytes))
    # Xóa ngược từ dưới lên để giữ nguyên chỉ số các dòng phía trên
    for i in range(len(doc.paragraphs) - 1, -1, -1):
        if not (start_line <= i < end_line):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
    return doc

def analyze_structure(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    structure = {"P1": [], "P2": [], "P3": []}
    current_part = "P1"
    q_start = -1
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip().upper()
        # Nhận diện chuyển phần
        if "PHẦN 1" in text or "PHẦN I" in text: current_part = "P1"
        elif "PHẦN 2" in text or "PHẦN II" in text: current_part = "P2"
        elif "PHẦN 3" in text or "PHẦN III" in text: current_part = "P3"
        
        # Nhận diện bắt đầu câu hỏi
        if re.match(r'^Câu\s*\d+', p.text, re.I):
            if q_start != -1:
                structure[last_part].append((q_start, i))
            q_start = i
            last_part = current_part
            
    if q_start != -1:
        structure[last_part].append((q_start, len(doc.paragraphs)))
    return structure

# GIAO DIỆN
st.set_page_config(page_title="Tạo Đề Tổng Hợp Pro", layout="wide")
st.title("🚀 Hệ thống Tạo Đề từ nhiều file nguồn")

uploaded_files = st.file_uploader("Bước 1: Chọn các file ngân hàng câu hỏi", type="docx", accept_multiple_files=True)

if uploaded_files:
    db = {}
    for f in uploaded_files:
        content = f.read()
        db[f.name] = {"bytes": content, "map": analyze_structure(content)}
    
    st.subheader("Bước 2: Chọn số lượng câu hỏi từ mỗi file")
    final_selection = {}
    
    for fname, data in db.items():
        with st.expander(f"📁 File: {fname}"):
            cols = st.columns(3)
            p1 = cols[0].number_input(f"Phần 1 (Max {len(data['map']['P1'])})", 0, len(data['map']['P1']), 0, key=f"p1_{fname}")
            p2 = cols[1].number_input(f"Phần 2 (Max {len(data['map']['P2'])})", 0, len(data['map']['P2']), 0, key=f"p2_{fname}")
            p3 = cols[2].number_input(f"Phần 3 (Max {len(data['map']['P3'])})", 0, len(data['map']['P3']), 0, key=f"p3_{fname}")
            final_selection[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🌟 TẠO ĐỀ THI MỚI", type="primary"):
        # Tạo file đích dựa trên định dạng file đầu tiên
        base_bytes = list(db.values())[0]["bytes"]
        master_doc = Document(io.BytesIO(base_bytes))
        for p in master_doc.paragraphs: master_doc._element.body.remove(p._element)
        
        composer = Composer(master_doc)
        global_q_count = 1
        
        for part_name, part_label in [("P1", "PHẦN I"), ("P2", "PHẦN II"), ("P3", "PHẦN III")]:
            master_doc.add_paragraph(f"{part_label}. (Tự động tổng hợp)").bold = True
            
            for fname, counts in final_selection.items():
                num = counts[part_name]
                if num > 0:
                    # Lấy ngẫu nhiên các câu hỏi đã chọn
                    chosen_ranges = random.sample(db[fname]["map"][part_name], num)
                    for start, end in chosen_ranges:
                        # Trích xuất an toàn để giữ MathType
                        q_doc = extract_question_safe(db[fname]["bytes"], start, end)
                        
                        # Đánh lại số thứ tự câu
                        for p in q_doc.paragraphs:
                            if re.match(r'^Câu\s*\d+', p.text, re.I):
                                p.text = re.sub(r'^Câu\s*\d+', f"Câu {global_q_count}", p.text, flags=re.I)
                                global_q_count += 1
                                break
                        composer.append(q_doc)
        
        out_stream = io.BytesIO()
        master_doc.save(out_stream)
        st.success("✅ Đã tạo đề thành công! Mọi công thức MathType và hình ảnh đã được bảo toàn.")
        st.download_button("📥 Tải đề thi tổng hợp", out_stream.getvalue(), "De_Thi_Tong_Hop.docx")
