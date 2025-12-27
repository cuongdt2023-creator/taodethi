import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Hàm này giúp trích xuất nội dung mà không làm mất MathType/Ảnh
def extract_content_safe(source_bytes, start_idx, end_idx):
    # Load lại file gốc để giữ nguyên toàn bộ định nghĩa công thức/ảnh
    doc = Document(io.BytesIO(source_bytes))
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    
    # Xóa ngược từ dưới lên những đoạn không thuộc câu hỏi được chọn
    for i in range(total - 1, -1, -1):
        if not (start_idx <= i < end_idx):
            p = paragraphs[i]._element
            p.getparent().remove(p)
    return doc

def analyze_file(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    questions = {"P1": [], "P2": [], "P3": []}
    current_part = "P1"
    q_start = -1
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().upper()
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"
        
        if re.match(r'^Câu\s*\d+', p.text, re.I):
            if q_start != -1:
                questions[last_part].append((q_start, i))
            q_start = i
            last_part = current_part
            
    if q_start != -1:
        questions[last_part].append((q_start, len(doc.paragraphs)))
    return questions

st.title("🛡️ Tạo Đề Tổng Hợp - Bảo Toàn MathType 100%")

uploaded_files = st.file_uploader("Tải các file đề nguồn (.docx)", type="docx", accept_multiple_files=True)

if uploaded_files:
    file_data = {}
    for f in uploaded_files:
        b = f.read()
        file_data[f.name] = {"bytes": b, "map": analyze_file(b)}
    
    # Giao diện chọn số câu (giữ nguyên logic của bạn)
    # ... (Phần hiển thị number_input cho từng file) ...

    if st.button("🚀 XUẤT ĐỀ THI CHUẨN"):
        # Lấy file đầu tiên làm mẫu định dạng
        master_doc = Document(io.BytesIO(list(file_data.values())[0]["bytes"]))
        for p in master_doc.paragraphs: 
            master_doc._element.body.remove(p._element)
        
        composer = Composer(master_doc)
        count = 1
        
        for part in ["P1", "P2", "P3"]:
            for fname, data in file_data.items():
                # Giả sử bạn đã lưu số câu chọn vào biến 'selected_num'
                # Code này mô phỏng việc lấy câu hỏi
                ranges = data["map"][part]
                for start, end in ranges:
                    # Trích xuất "nguyên khối" để không mất MathType
                    sub_doc = extract_content_safe(data["bytes"], start, end)
                    
                    # Đánh lại số câu mà không làm hỏng công thức đi kèm
                    for p in sub_doc.paragraphs:
                        if re.match(r'^Câu\s*\d+', p.text, re.I):
                            p.text = re.sub(r'^Câu\s*\d+', f"Câu {count}", p.text, flags=re.I)
                            count += 1
                            break
                    composer.append(sub_doc)
        
        out = io.BytesIO()
        master_doc.save(out)
        st.download_button("📥 Tải đề hoàn thiện", out.getvalue(), "De_Thi_Chuan.docx")
