import streamlit as st
import io
import re
from docx import Document
from docxcompose.composer import Composer

def extract_safe(source_bytes, start_idx, end_idx):
    """
    KỸ THUẬT QUAN TRỌNG NHẤT:
    Mở file gốc, xóa tất cả các dòng KHÔNG nằm trong khoảng (start, end).
    Điều này giữ lại nguyên vẹn các Object MathType/Hệ phương trình 
    vì chúng chưa bao giờ bị copy sang nơi khác.
    """
    doc = Document(io.BytesIO(source_bytes))
    total = len(doc.paragraphs)
    # Xóa ngược từ dưới lên để không làm hỏng index
    for i in range(total - 1, -1, -1):
        if not (start_idx <= i < end_idx):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)
    return doc

# --- PHẦN XỬ LÝ CHÍNH TRÊN STREAMLIT ---
st.title("Hệ Thống Tạo Đề Bảo Toàn Hệ Phương Trình")

uploaded_file = st.file_uploader("Tải file Hệ phương trình.docx", type="docx")

if uploaded_file:
    file_bytes = uploaded_file.read()
    doc = Document(io.BytesIO(file_bytes))
    
    # Quét vị trí các câu hỏi
    q_map = []
    start = -1
    for i, p in enumerate(doc.paragraphs):
        if re.match(r'^Câu\s*\d+', p.text.strip(), re.I):
            if start != -1: q_map.append((start, i))
            start = i
    if start != -1: q_map.append((start, len(doc.paragraphs)))

    st.write(f"Tìm thấy {len(q_map)} câu hỏi có chứa hệ phương trình.")

    if st.button("XUẤT ĐỀ (GIỮ NGUYÊN HỆ PHƯƠNG TRÌNH)"):
        # Tạo file Master từ định dạng file gốc
        master_doc = Document(io.BytesIO(file_bytes))
        for p in master_doc.paragraphs:
            p._element.getparent().remove(p._element)
        
        composer = Composer(master_doc)
        
        # Giả sử lấy 5 câu đầu tiên để test
        for i in range(min(5, len(q_map))):
            s, e = q_map[i]
            # Lấy nguyên khối câu hỏi (bao gồm cả hệ phương trình)
            q_doc = extract_safe(file_bytes, s, e)
            
            # Chỉ sửa số câu trên văn bản, không chạm vào Object công thức
            for p in q_doc.paragraphs:
                if re.match(r'^Câu\s*\d+', p.text, re.I):
                    p.text = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", p.text, flags=re.I)
                    break
            
            # Gộp vào file chính
            composer.append(q_doc)
            
        out = io.BytesIO()
        master_doc.save(out)
        st.download_button("Tải file kết quả chuẩn", out.getvalue(), "Ket_Qua_Chinh_Xac.docx")
