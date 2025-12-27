import io
from docx import Document
from docxcompose.composer import Composer

def combine_questions_to_docx(selected_questions, template_file_bytes):
    """
    selected_questions: Danh sách các đối tượng câu hỏi đã chọn
    template_file_bytes: File mẫu để giữ định dạng trang (lề, khổ giấy)
    """
    # 1. Khởi tạo tài liệu Master từ file mẫu
    master_doc = Document(io.BytesIO(template_file_bytes))
    
    # 2. Xóa sạch nội dung cũ trong file mẫu nhưng giữ lại định dạng trang
    for p in master_doc.paragraphs:
        p._element.getparent().remove(p._element)
    
    # 3. Tạo đối tượng Composer
    composer = Composer(master_doc)
    
    for q_data in selected_questions:
        # q_data["file_bytes"] là nội dung file .docx chứa câu hỏi đó
        temp_doc = Document(io.BytesIO(q_data["file_bytes"]))
        
        # 4. Sử dụng composer.append để gộp câu hỏi vào Master
        # Cách này sẽ tự động mang theo hình ảnh, công thức và bảng biểu
        composer.append(temp_doc)

    # 5. Xuất file cuối cùng
    output = io.BytesIO()
    master_doc.save(output)
    return output.getvalue()
