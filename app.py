import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Hàm này giữ nguyên 100% MathType và Ảnh bằng cách xóa phần không dùng
def extract_safe(source_bytes, start_idx, end_idx):
    doc = Document(io.BytesIO(source_bytes))
    total = len(doc.paragraphs)
    # Xóa ngược từ dưới lên để không lệch vị trí
    for i in range(total - 1, -1, -1):
        if not (start_idx <= i < end_idx):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
    return doc

def analyze_questions(file_bytes):
    doc = Document(io.BytesIO(file_bytes))
    mapping = {"P1": [], "P2": [], "P3": []}
    current_part = "P1"
    q_start = -1
    
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip().upper()
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"
        
        if re.match(r'^Câu\s*\d+', p.text, re.I):
            if q_start != -1:
                mapping[last_part].append((q_start, i))
            q_start = i
            last_part = current_part
            
    if q_start != -1:
        mapping[last_part].append((q_start, len(doc.paragraphs)))
    return mapping

st.title("Tạo Đề Tổng Hợp (Bảo Toàn Hệ Phương Trình)")

files = st.file_uploader("Tải các file đề nguồn", type="docx", accept_multiple_files=True)

if files:
    db = {f.name: {"bytes": f.read(), "map": analyze_questions(f.getvalue())} for f in files}
    
    # Giao diện chọn câu (Ví dụ đơn giản)
    selected_counts = {}
    for fname in db:
        st.write(f"--- File: {fname} ---")
        c1, c2, c3 = st.columns(3)
        n1 = c1.number_input(f"P1", 0, 50, 0, key=f"n1_{fname}")
        n2 = c2.number_input(f"P2", 0, 50, 0, key=f"n2_{fname}")
        n3 = c3.number_input(f"P3", 0, 50, 0, key=f"n3_{fname}")
        selected_counts[fname] = {"P1": n1, "P2": n2, "P3": n3}

    if st.button("XUẤT ĐỀ THI"):
        try:
            # Lấy file đầu tiên làm mẫu
            first_file_bytes = list(db.values())[0]["bytes"]
            master_doc = Document(io.BytesIO(first_file_bytes))
            for p in master_doc.paragraphs: master_doc._element.body.remove(p._element)
            
            composer = Composer(master_doc)
            q_num = 1
            
            for part in ["P1", "P2", "P3"]:
                master_doc.add_paragraph(f"PHẦN {part[-1]}.").bold = True
                for fname, counts in selected_counts.items():
                    if counts[part] > 0:
                        chosen = random.sample(db[fname]["map"][part], counts[part])
                        for start, end in chosen:
                            # TRÍCH XUẤT NGUYÊN KHỐI
                            q_doc = extract_safe(db[fname]["bytes"], start, end)
                            # Đổi số câu
                            for p in q_doc.paragraphs:
                                if re.match(r'^Câu\s*\d+', p.text, re.I):
                                    p.text = re.sub(r'^Câu\s*\d+', f"Câu {q_num}", p.text, flags=re.I)
                                    q_num += 1
                                    break
                            composer.append(q_doc)
            
            out = io.BytesIO()
            master_doc.save(out)
            st.download_button("Tải file kết quả", out.getvalue(), "Ket_Qua_Chuan.docx")
        except Exception as e:
            st.error(f"Lỗi: {e}")
