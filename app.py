import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# ==================== CẤU HÌNH TRANG ====================
st.set_page_config(page_title="AIOMT - Gộp Đề Chuẩn", page_icon="🎯", layout="wide")

def get_question_difficulty(paragraph_text):
    """Xác định độ khó của câu hỏi dựa trên tag #NB, #TH..."""
    t = paragraph_text.upper()
    for tag in ["#VDC", "#VD", "#TH", "#NB"]:
        if tag in t: return tag[1:]
    return "NB"

def split_into_parts(doc):
    """Phân tách tài liệu thành 3 kho lưu trữ Phần 1, 2, 3."""
    sections = {"P1": [], "P2": [], "P3": []}
    current_part = "P1"
    current_question = []
    
    # Duyệt qua các thành phần của tài liệu
    for p in doc.paragraphs:
        txt = p.text.upper()
        if "PHẦN 1" in txt or "PHẦN I" in txt: current_part = "P1"
        elif "PHẦN 2" in txt or "PHẦN II" in txt: current_part = "P2"
        elif "PHẦN 3" in txt or "PHẦN III" in txt: current_part = "P3"
        
        # Nhận diện điểm bắt đầu của một câu hỏi mới
        if re.match(r'^Câu\s*\d+', p.text, re.IGNORECASE):
            if current_question:
                # Lưu câu hỏi cũ vào kho
                diff = get_question_difficulty(current_question[0].text)
                sections[prev_part].append({"content": current_question, "diff": diff})
            current_question = [p]
            prev_part = current_part
        elif current_question:
            current_question.append(p)
            
    # Lưu câu hỏi cuối cùng
    if current_question:
        diff = get_question_difficulty(current_question[0].text)
        sections[prev_part].append({"content": current_question, "diff": diff})
        
    return sections

# ==================== GIAO DIỆN ====================
st.title("🧩 Gộp Đề Tổng Hợp (Fix Lỗi Content)")

uploaded_files = st.file_uploader("Tải các file chủ đề (.docx)", type="docx", accept_multiple_files=True)

if uploaded_files:
    if 'bank' not in st.session_state:
        st.session_state.bank = {}
        for f in uploaded_files:
            doc = Document(io.BytesIO(f.read()))
            st.session_state.bank[f.name] = split_into_parts(doc)

    st.subheader("Cấu hình số câu lấy từ mỗi chủ đề")
    configs = {}
    cols = st.columns(len(uploaded_files))
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]:
            st.info(f"📁 {fname}")
            p1 = st.number_input(f"P1", 0, 50, 0, key=f"p1_{fname}")
            p2 = st.number_input(f"P2", 0, 50, 0, key=f"p2_{fname}")
            p3 = st.number_input(f"P3", 0, 50, 0, key=f"p3_{fname}")
            configs[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🚀 TẠO ĐỀ THI TỔNG HỢP", type="primary", use_container_width=True):
        final_doc = Document() # Tạo file mới
        composer = Composer(final_doc)
        
        titles = {
            "P1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
            "P2": "PHẦN II. Câu trắc nghiệm đúng sai.",
            "P3": "PHẦN III. Câu trắc nghiệm trả lời ngắn."
        }

        for part_key in ["P1", "P2", "P3"]:
            # Lấy danh sách câu hỏi được chọn
            selected_questions = []
            for fname, cfg in configs.items():
                num_needed = cfg[part_key]
                pool = []
                for diff in ["NB", "TH", "VD", "VDC"]:
                    pool.extend(st.session_state.bank[fname][part_key][diff])
                
                if len(pool) >= num_needed:
                    selected_questions.extend(random.sample(pool, num_needed))
                else:
                    selected_questions.extend(pool)

            if selected_questions:
                # 1. Thêm tiêu đề phần
                p_title = final_doc.add_paragraph()
                run = p_title.add_run(titles[part_key])
                run.bold = True
                
                # 2. Trộn thứ tự câu hỏi trong phần
                random.shuffle(selected_questions)
                
                # 3. Chèn nội dung và đánh lại số câu
                for idx, q_data in enumerate(selected_questions):
                    for i, para in enumerate(q_data["content"]):
                        new_p = final_doc.add_paragraph()
                        # Đánh lại số câu tại dòng đầu tiên
                        text = para.text
                        if i == 0:
                            text = re.sub(r'^Câu\s*\d+', f"Câu {idx+1}", text, flags=re.IGNORECASE)
                            text = re.sub(r'#(NB|TH|VD|VDC)', '', text)
                        
                        new_p.text = text
                        # Copy định dạng (đơn giản)
                        new_p.style = para.style

        # Xuất file
        output = io.BytesIO()
        final_doc.save(output)
        st.success("🎉 Đề thi đã được tạo thành công và không còn lỗi cấu trúc!")
        st.download_button("📥 Tải đề thi tổng hợp", output.getvalue(), "De_Tong_Hop_Final.docx")
