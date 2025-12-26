import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== CẤU HÌNH ====================
st.set_page_config(page_title="Tạo Đề Từ Nhiều Chủ Đề", page_icon="📚", layout="wide")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# ==================== LOGIC XỬ LÝ XML ====================

def get_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def detect_difficulty(text):
    text_upper = text.upper()
    if "#NB" in text_upper: return "NB"
    if "#TH" in text_upper: return "TH"
    if "#VDC" in text_upper: return "VDC"
    if "#VD" in text_upper: return "VD"
    return "NB"

def clean_tags(block):
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            val = t.firstChild.nodeValue
            new_val = re.sub(r'#(NB|TH|VDC|VD)\b', '', val, flags=re.IGNORECASE)
            t.firstChild.nodeValue = new_val

def parse_docx_to_questions(file_bytes):
    """Phân tích 1 file docx thành danh sách câu hỏi theo từng phần"""
    questions = {"PHAN1": [], "PHAN2": [], "PHAN3": []}
    
    input_buffer = io.BytesIO(file_bytes)
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        doc_xml = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        blocks = [c for c in body.childNodes if c.nodeType == 1 and c.localName in ["p", "tbl"]]
        
        current_part = "PHAN1"
        current_q_blocks = []
        
        for block in blocks:
            txt = get_text(block).upper()
            if "PHẦN 1" in txt: current_part = "PHAN1"
            elif "PHẦN 2" in txt: current_part = "PHAN2"
            elif "PHẦN 3" in txt: current_part = "PHAN3"
            
            if re.match(r'^Câu\s*\d+', get_text(block), re.IGNORECASE):
                if current_q_blocks:
                    questions[prev_part].append({
                        "blocks": current_q_blocks,
                        "diff": detect_difficulty(get_text(current_q_blocks[0]))
                    })
                current_q_blocks = [block]
                prev_part = current_part
            else:
                if current_q_blocks:
                    current_q_blocks.append(block)
        
        if current_q_blocks:
            questions[prev_part].append({
                "blocks": current_q_blocks,
                "diff": detect_difficulty(get_text(current_q_blocks[0]))
            })
            
    return questions

def update_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            txt = t.firstChild.nodeValue
            if "Câu" in txt:
                t.firstChild.nodeValue = re.sub(r'Câu\s*\d+', new_label, txt)
                break

# ==================== GIAO DIỆN STREAMLIT ====================

st.title("🧩 Tạo Đề Tổng Hợp Từ Nhiều Chủ Đề")

uploaded_files = st.file_uploader(
    "Upload các file chủ đề (.docx)", 
    type="docx", 
    accept_multiple_files=True
)

if uploaded_files:
    bank_data = {}
    for f in uploaded_files:
        bank_data[f.name] = parse_docx_to_questions(f.read())
    
    st.divider()
    st.subheader("📊 Thiết lập ma trận câu hỏi cho từng file")
    
    total_config = {}
    
    # Tạo bảng nhập liệu cho mỗi file
    for fname in bank_data.keys():
        with st.expander(f"📁 Chủ đề: {fname}", expanded=True):
            cols = st.columns(4)
            nb = cols[0].number_input(f"NB ({fname})", 0, 50, 0, key=f"{fname}_nb")
            th = cols[1].number_input(f"TH ({fname})", 0, 50, 0, key=f"{fname}_th")
            vd = cols[2].number_input(f"VD ({fname})", 0, 50, 0, key=f"{fname}_vd")
            vdc = cols[3].number_input(f"VDC ({fname})", 0, 50, 0, key=f"{fname}_vdc")
            total_config[fname] = {"NB": nb, "TH": th, "VD": vd, "VDC": vdc}

    if st.button("🚀 Tạo Đề Thi Tổng Hợp", type="primary", use_container_width=True):
        final_selected_blocks = []
        parts = ["PHAN1", "PHAN2", "PHAN3"]
        
        # Để đơn giản, ta sẽ gom tất cả câu hỏi được chọn từ các file theo từng Phần
        all_part_blocks = {"PHAN1": [], "PHAN2": [], "PHAN3": []}
        
        for fname, config in total_config.items():
            file_qs = bank_data[fname]
            for part in parts:
                for diff in ["NB", "TH", "VD", "VDC"]:
                    needed = config[diff]
                    pool = [q for q in file_qs[part] if q['diff'] == diff]
                    if len(pool) < needed:
                        st.warning(f"File {fname} không đủ {needed} câu {diff} ở {part}")
                        selected = pool
                    else:
                        selected = random.sample(pool, needed)
                    
                    for s in selected:
                        all_part_blocks[part].append(s['blocks'])
        
        # Trộn và đánh số lại
        final_doc_blocks = []
        for part in parts:
            random.shuffle(all_part_blocks[part])
            for idx, q_blocks in enumerate(all_part_blocks[part]):
                first_blk = q_blocks[0]
                update_label(first_blk, f"Câu {idx + 1}")
                clean_tags(first_blk)
                final_doc_blocks.extend(q_blocks)

        # Xuất file (Dùng file đầu tiên làm template cho style)
        first_file_bytes = uploaded_files[0].getvalue()
        output_buffer = io.BytesIO()
        
        with zipfile.ZipFile(io.BytesIO(first_file_bytes), 'r') as zin:
            doc_xml = zin.read("word/document.xml").decode('utf-8')
            dom = minidom.parseString(doc_xml)
            body = dom.getElementsByTagNameNS(W_NS, "body")[0]
            
            # Xóa sạch body cũ
            for child in list(body.childNodes):
                if child.nodeType == 1: body.removeChild(child)
            
            # Chèn câu hỏi mới
            for blk in final_doc_blocks:
                body.appendChild(blk)
            
            new_xml = dom.toxml()
            with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, new_xml.encode('utf-8'))
                    else:
                        zout.writestr(item, zin.read(item.filename))

        st.success("✅ Đã tạo xong đề thi tổng hợp!")
        st.download_button("📥 Tải đề thi (.docx)", output_buffer.getvalue(), "De_Tong_Hop.docx")
