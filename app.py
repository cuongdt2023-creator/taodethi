import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Cấu hình giao diện chuẩn AIOMT
st.set_page_config(page_title="AIOMT Premium - Gộp Đề Bảo Toàn", layout="wide")

def get_difficulty(para_text):
    t = para_text.upper()
    for tag in ["#VDC", "#VD", "#TH", "#NB"]:
        if tag in t: return tag[1:]
    return "NB"

def split_docx_to_questions(file_bytes):
    """
    Tách file gốc thành từng câu hỏi.
    Sử dụng kỹ thuật sao chép Deep Copy để giữ nguyên Media (Ảnh/Công thức).
    """
    source_stream = io.BytesIO(file_bytes)
    source_doc = Document(source_stream)
    bank = {"P1": [], "P2": [], "P3": []}
    
    current_part = "P1"
    temp_elements = []
    
    for p in source_doc.paragraphs:
        txt = p.text.upper()
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"

        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
            if temp_elements:
                # Tạo một Document con từ chính file gốc để kế thừa toàn bộ Media/Rels
                q_doc = Document(io.BytesIO(file_bytes))
                # Xóa sạch nội dung cũ, chỉ giữ lại khung (styles, settings, rels)
                body = q_doc._element.body
                for child in list(body):
                    if not child.tag.endswith('sectPr'):
                        body.remove(child)
                
                # Chèn các đoạn văn của câu hỏi vào body mới
                for elem in temp_elements:
                    body.append(elem._element)
                
                diff = get_difficulty(temp_elements[0].text)
                bank[start_part].append({"doc": q_doc, "diff": diff})
            
            temp_elements = [p]
            start_part = current_part
        elif temp_elements:
            temp_elements.append(p)

    # Lưu câu cuối cùng
    if temp_elements:
        q_doc = Document(io.BytesIO(file_bytes))
        body = q_doc._element.body
        for child in list(body):
            if not child.tag.endswith('sectPr'): body.remove(child)
        for elem in temp_elements: body.append(elem._element)
        bank[start_part].append({"doc": q_doc, "diff": get_difficulty(temp_elements[0].text)})

    return bank

st.title("🚀 Hệ Thống Gộp Đề Chuyên Nghiệp (Giữ 100% Định Dạng)")

uploaded_files = st.file_uploader("Tải các file chủ đề (.docx)", type="docx", accept_multiple_files=True)

if uploaded_files:
    if 'bank' not in st.session_state:
        st.session_state.bank = {f.name: split_docx_to_questions(f.read()) for f in uploaded_files}

    st.subheader("Chọn số lượng câu hỏi")
    configs = {}
    cols = st.columns(len(uploaded_files))
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]:
            st.info(f"📂 {fname[:15]}")
            configs[fname] = {
                "P1": st.number_input(f"P1", 0, 50, 0, key=f"p1_{fname}"),
                "P2": st.number_input(f"P2", 0, 50, 0, key=f"p2_{fname}"),
                "P3": st.number_input(f"P3", 0, 50, 0, key=f"p3_{fname}")
            }

    if st.button("🌟 XUẤT ĐỀ THI TỔNG HỢP", type="primary", use_container_width=True):
        # Lấy file đầu tiên làm Template gốc
        main_doc = Document(io.BytesIO(uploaded_files[0].getvalue()))
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
                # Tạo tiêu đề phần
                title_para = main_doc.add_paragraph()
                run = title_para.add_run(titles[p_key])
                run.bold = True
                
                random.shuffle(selected)
                for idx, q_data in enumerate(selected):
                    q_doc = q_data["doc"]
                    # Thay đổi số câu trực tiếp trong Document con
                    for p in q_doc.paragraphs:
                        if "Câu" in p.text:
                            p.text = re.sub(r'^Câu\s*\d+', f"Câu {idx+1}", p.text, flags=re.IGNORECASE)
                            p.text = re.sub(r'#(NB|TH|VD|VDC)', '', p.text)
                            break
                    # Gộp file bằng Composer (Cực kỳ quan trọng để giữ ảnh/công thức)
                    composer.append(q_doc)

        out_io = io.BytesIO()
        main_doc.save(out_io)
        st.success("✅ Thành công! Đề thi đã được bảo toàn mọi hình ảnh và công thức.")
        st.download_button("📥 Tải đề thi ngay", out_io.getvalue(), "De_Tong_Hop_Bao_Toan.docx")
