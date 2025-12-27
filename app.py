import streamlit as st
import io
import random
import re
from docx import Document
from docxcompose.composer import Composer

# Hàm cốt lõi để bảo toàn MathType và Hình ảnh
def get_cleaned_question_doc(file_bytes, start_idx, end_idx):
    """
    Tạo một bản sao của file gốc và xóa mọi thứ trừ đoạn từ start_idx đến end_idx.
    """
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    
    # Xóa từ dưới lên để không làm thay đổi chỉ số (index) của các đoạn phía trên
    for i in range(total - 1, -1, -1):
        if not (start_idx <= i < end_idx):
            p = paragraphs[i]._element
            p.getparent().remove(p)
            
    return doc

def main():
    st.title("🛡️ Trộn Đề Bảo Toàn Tuyệt Đối MathType & Hình Vẽ")
    st.info("Phương pháp: Cắt tỉa trực tiếp trên file gốc để giữ nguyên 100% định dạng.")

    uploaded_file = st.file_uploader("Tải lên file đề (.docx)", type="docx")

    if uploaded_file:
        file_bytes = uploaded_file.read()
        doc = Document(io.BytesIO(file_bytes))
        
        # Nhận diện vị trí các câu hỏi
        q_map = []
        current_start = -1
        
        for i, p in enumerate(doc.paragraphs):
            # Nhận diện "Câu 1.", "Câu 2."...
            if re.match(r'^Câu\s*\d+', p.text.strip(), re.IGNORECASE):
                if current_start != -1:
                    q_map.append((current_start, i))
                current_start = i
        
        if current_start != -1:
            q_map.append((current_start, len(doc.paragraphs)))

        st.success(f"Tìm thấy {len(q_map)} câu hỏi.")
        
        num_versions = st.number_input("Số lượng mã đề:", 1, 20, 4)

        if st.button("🚀 Bắt đầu trộn đề"):
            # Tạo file Master để gộp (lấy định dạng từ file gốc)
            master_output = Document(io.BytesIO(file_bytes))
            for p in master_output.paragraphs:
                p._element.getparent().remove(p._element)
            
            composer = Composer(master_output)
            
            # Trộn thứ tự
            shuffled_indices = list(range(len(q_map)))
            random.shuffle(shuffled_indices)

            with st.spinner("Đang xử lý bảo toàn dữ liệu..."):
                for new_idx, old_idx in enumerate(shuffled_indices):
                    start, end = q_map[old_idx]
                    
                    # CẮT TỈA: Lấy file chứa duy nhất câu hỏi này từ file gốc
                    temp_doc = get_cleaned_question_doc(file_bytes, start, end)
                    
                    # Đánh lại số câu (vẫn giữ định dạng)
                    for p in temp_doc.paragraphs:
                        if re.match(r'^Câu\s*\d+', p.text.strip(), re.IGNORECASE):
                            p.text = re.sub(r'^Câu\s*\d+', f"Câu {new_idx + 1}", p.text, flags=re.IGNORECASE)
                            break
                    
                    # GỘP AN TOÀN bằng docxcompose
                    composer.append(temp_doc)

            # Xuất file
            out_io = io.BytesIO()
            master_output.save(out_io)
            st.download_button("📥 Tải đề đã trộn", out_io.getvalue(), "De_Thi_Bao_Toan.docx")

if __name__ == "__main__":
    main()
