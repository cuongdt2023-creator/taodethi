import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Kỹ thuật "Cắt tỉa": Nhân bản file gốc rồi xóa phần thừa để giữ 100% công thức
def extract_safe(source_bytes, start_idx, end_idx):
    doc = Document(io.BytesIO(source_bytes))
    total = len(doc.paragraphs)
    for i in range(total - 1, -1, -1):
        if not (start_idx <= i < end_idx):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
    return doc

def analyze_structure(file_bytes):
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
            if q_start != -1: mapping[last_part].append((q_start, i))
            q_start, last_part = i, current_part
    if q_start != -1: mapping[last_part].append((q_start, len(doc.paragraphs)))
    return mapping

st.set_page_config(page_title="Tạo Đề Tổng Hợp", layout="wide")
st.title("🚀 Tạo Đề Mới (Bảo Toàn MathType/Hình Ảnh)")

files = st.file_uploader("Tải các file ngân hàng câu hỏi", type="docx", accept_multiple_files=True)

if files:
    db = {f.name: {"bytes": f.read(), "map": analyze_structure(f.getvalue())} for f in files}
    st.info("Chọn số lượng câu hỏi từ mỗi file nguồn:")
    
    selected_config = {}
    for fname in db:
        with st.expander(f"📁 {fname}"):
            c1, c2, c3 = st.columns(3)
            p1 = c1.number_input("Phần I", 0, 50, 0, key=f"p1_{fname}")
            p2 = c2.number_input("Phần II", 0, 50, 0, key=f"p2_{fname}")
            p3 = c3.number_input("Phần III", 0, 50, 0, key=f"p3_{fname}")
            selected_config[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🌟 XUẤT ĐỀ THI TỔNG HỢP", type="primary"):
        # Tạo file Master (lấy định dạng từ file đầu tiên)
        master_doc = Document(io.BytesIO(list(db.values())[0]["bytes"]))
        for p in master_doc.paragraphs: master_doc._element.body.remove(p._element)
        
        composer = Composer(master_doc)
        global_q = 1
        
        for p_key, p_label in [("P1", "PHẦN I"), ("P2", "PHẦN II"), ("P3", "PHẦN III")]:
            master_doc.add_paragraph(f"{p_label}.").bold = True
            for fname, cfg in selected_config.items():
                if cfg[p_key] > 0:
                    chosen = random.sample(db[fname]["map"][p_key], cfg[p_key])
                    for start, end in chosen:
                        q_doc = extract_safe(db[fname]["bytes"], start, end)
                        # Đổi số câu mà không làm hỏng MathType
                        for p in q_doc.paragraphs:
                            if re.match(r'^Câu\s*\d+', p.text, re.I):
                                p.text = re.sub(r'^Câu\s*\d+', f"Câu {global_q}", p.text, flags=re.I)
                                global_q += 1; break
                        composer.append(q_doc)
        
        out = io.BytesIO()
        master_doc.save(out)
        st.success("Đã tạo đề thành công! Hệ phương trình và ảnh được giữ nguyên.")
        st.download_button("📥 Tải đề thi", out.getvalue(), "De_Thi_Tong_Hop.docx")
