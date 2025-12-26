import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== CẤU HÌNH GIAO DIỆN ====================
st.set_page_config(page_title="AIOMT - Fix Lỗi Gộp Đề", page_icon="🛠️", layout="wide")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    """Lấy văn bản thô từ một block XML để nhận diện tiêu đề/câu hỏi."""
    texts = [t.firstChild.nodeValue for t in block.getElementsByTagNameNS(W_NS, "t") if t.firstChild]
    return "".join(texts).strip()

def detect_difficulty(text):
    """Nhận diện độ khó từ các tag #NB, #TH...."""
    t = text.upper()
    for tag in ["#VDC", "#VD", "#TH", "#NB"]:
        if tag in t: return tag[1:]
    return "NB"

def parse_docx(file_bytes):
    """Tách câu hỏi và phân loại theo 3 phần chuẩn."""
    data = {p: {d: [] for d in ["NB", "TH", "VD", "VDC"]} for p in ["P1", "P2", "P3"]}
    with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as zin:
        xml_content = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(xml_content)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        blocks = [c for c in body.childNodes if c.nodeType == 1 and c.localName in ["p", "tbl"]]
        
        curr_p, curr_q = "P1", []
        prev_p = "P1"
        for b in blocks:
            txt = get_text(b).upper()
            if "PHẦN 1" in txt: curr_p = "P1"
            elif "PHẦN 2" in txt: curr_p = "P2"
            elif "PHẦN 3" in txt: curr_p = "P3"
            
            if re.match(r'^Câu\s*\d+', get_text(b), re.IGNORECASE):
                if curr_q:
                    diff = detect_difficulty(get_text(curr_q[0]))
                    data[prev_p][diff].append(curr_q)
                curr_q, prev_p = [b], curr_p
            elif curr_q:
                curr_q.append(b)
        if curr_q:
            data[prev_p][detect_difficulty(get_text(curr_q[0]))].append(curr_q)
    return data

def create_heading_paragraph(text, dom):
    """Tạo Paragraph XML chuẩn cho tiêu đề để tránh lỗi cấu trúc Word."""
    p = dom.createElementNS(W_NS, "w:p")
    pPr = dom.createElementNS(W_NS, "w:pPr")
    # Định dạng in đậm và cỡ chữ cho tiêu đề Phần
    rPr = dom.createElementNS(W_NS, "w:rPr")
    b = dom.createElementNS(W_NS, "w:b")
    rPr.appendChild(b)
    
    r = dom.createElementNS(W_NS, "w:r")
    r.appendChild(rPr)
    t = dom.createElementNS(W_NS, "w:t")
    t.appendChild(dom.createTextNode(text))
    r.appendChild(t)
    p.appendChild(pPr)
    p.appendChild(r)
    return p

# ==================== GIAO DIỆN CHÍNH ====================
st.title("🧩 Trích xuất & Gộp Đề Chuẩn Cấu Trúc")

files = st.file_uploader("Tải các file chủ đề (.docx)", type="docx", accept_multiple_files=True)

if files:
    if 'bank' not in st.session_state:
        st.session_state.bank = {f.name: parse_docx(f.read()) for f in files}

    configs = {}
    cols = st.columns(len(files))
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]:
            st.info(f"📁 {fname}")
            p1 = st.number_input(f"P1 (Câu)", 0, 20, 0, key=f"n1_{fname}")
            p2 = st.number_input(f"P2 (Câu)", 0, 10, 0, key=f"n2_{fname}")
            p3 = st.number_input(f"P3 (Câu)", 0, 10, 0, key=f"n3_{fname}")
            configs[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🚀 XUẤT ĐỀ THI TỔNG HỢP", type="primary", use_container_width=True):
        final_selected = {"P1": [], "P2": [], "P3": []}
        for fname, cfg in configs.items():
            for p in ["P1", "P2", "P3"]:
                pool = []
                for d in ["NB", "TH", "VD", "VDC"]:
                    pool.extend(st.session_state.bank[fname][p][d])
                if len(pool) >= cfg[p] and cfg[p] > 0:
                    final_selected[p].extend(random.sample(pool, cfg[p]))

        output = io.BytesIO()
        # Sử dụng file đầu tiên làm mẫu để lấy các khai báo namespace chuẩn
        with zipfile.ZipFile(io.BytesIO(files[0].getvalue()), 'r') as zin:
            with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        doc_dom = minidom.parseString(zin.read(item.filename).decode('utf-8'))
                        body = doc_dom.getElementsByTagNameNS(W_NS, "body")[0]
                        # Xóa nội dung cũ để xây dựng lại từ đầu
                        for child in list(body.childNodes):
                            if child.nodeType == 1 and child.localName != "sectPr":
                                body.removeChild(child)
                        
                        sectPr = body.getElementsByTagNameNS(W_NS, "sectPr")[-1]
                        
                        titles = {
                            "P1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.",
                            "P2": "PHẦN II. Câu trắc nghiệm đúng sai.",
                            "P3": "PHẦN III. Câu trắc nghiệm trả lời ngắn."
                        }
                        
                        for p in ["P1", "P2", "P3"]:
                            if final_selected[p]:
                                body.insertBefore(create_heading_paragraph(titles[p], doc_dom), sectPr)
                                for idx, q_blocks in enumerate(final_selected[p]):
                                    # Cập nhật số câu và xóa tag
                                    for block in q_blocks:
                                        # Import block từ file gốc vào tài liệu mới để tránh lỗi sở hữu node
                                        imported_block = doc_dom.importNode(block, True)
                                        if block == q_blocks[0]:
                                            t_nodes = imported_block.getElementsByTagNameNS(W_NS, "t")
                                            for t in t_nodes:
                                                if t.firstChild and "Câu" in t.firstChild.nodeValue:
                                                    t.firstChild.nodeValue = re.sub(r'Câu\s*\d+', f"Câu {idx+1}", t.firstChild.nodeValue)
                                                    t.firstChild.nodeValue = re.sub(r'#(NB|TH|VD|VDC)', '', t.firstChild.nodeValue)
                                                    break
                                        body.insertBefore(imported_block, sectPr)
                        
                        zout.writestr(item, doc_dom.toxml().encode('utf-8'))
                    else:
                        zout.writestr(item, zin.read(item.filename))
        
        st.download_button("📥 Tải đề đã sửa lỗi", output.getvalue(), "De_Thi_Chuan.docx")
