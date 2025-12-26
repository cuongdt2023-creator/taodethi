import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== CẤU HÌNH GIAO DIỆN ====================
st.set_page_config(page_title="AIOMT - Tạo Đề Chuẩn 3 Phần", page_icon="📝", layout="wide")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    texts = [t.firstChild.nodeValue for t in block.getElementsByTagNameNS(W_NS, "t") if t.firstChild]
    return "".join(texts).strip()

def detect_difficulty(text):
    t = text.upper()
    for tag in ["#VDC", "#VD", "#TH", "#NB"]:
        if tag in t: return tag[1:]
    return "NB"

def parse_docx(file_bytes):
    """Phân tích file thành: {Phần: {Độ khó: [Danh sách câu]}}"""
    data = {p: {d: [] for d in ["NB", "TH", "VD", "VDC"]} for p in ["P1", "P2", "P3"]}
    with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as zin:
        xml_content = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(xml_content)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        blocks = [c for c in body.childNodes if c.nodeType == 1 and c.localName in ["p", "tbl"]]
        
        curr_p, curr_q = "P1", []
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
            elif curr_q: curr_q.append(b)
        if curr_q: data[prev_p][detect_difficulty(get_text(curr_q[0]))].append(curr_q)
    return data

def create_heading_xml(text, dom):
    """Tạo XML cho tiêu đề Phần"""
    p = dom.createElementNS(W_NS, "w:p")
    pPr = dom.createElementNS(W_NS, "w:pPr")
    jc = dom.createElementNS(W_NS, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", "left")
    pPr.appendChild(jc)
    p.appendChild(pPr)
    
    r = dom.createElementNS(W_NS, "w:r")
    rPr = dom.createElementNS(W_NS, "w:rPr")
    b = dom.createElementNS(W_NS, "w:b")
    rPr.appendChild(b)
    sz = dom.createElementNS(W_NS, "w:sz")
    sz.setAttributeNS(W_NS, "w:val", "28") # Size 14
    rPr.appendChild(sz)
    r.appendChild(rPr)
    
    t = dom.createElementNS(W_NS, "w:t")
    t.appendChild(dom.createTextNode(text))
    r.appendChild(t)
    p.appendChild(r)
    return p

# ==================== GIAO DIỆN CHÍNH ====================
st.title("🧩 Trích xuất & Gộp Đề Theo Cấu Trúc Chuẩn")

files = st.file_uploader("Tải các file chủ đề (.docx)", type="docx", accept_multiple_files=True)

if files:
    if 'bank' not in st.session_state:
        st.session_state.bank = {f.name: parse_docx(f.read()) for f in files}

    st.subheader("1. Chọn số lượng câu hỏi từ mỗi file")
    configs = {}
    cols = st.columns(len(files))
    for i, fname in enumerate(st.session_state.bank.keys()):
        with cols[i]:
            st.info(f"📁 {fname}")
            p1 = st.number_input(f"P1 (Câu)", 0, 20, 0, key=f"n1_{fname}")
            p2 = st.number_input(f"P2 (Câu)", 0, 10, 0, key=f"n2_{fname}")
            p3 = st.number_input(f"P3 (Câu)", 0, 10, 0, key=f"n3_{fname}")
            configs[fname] = {"P1": p1, "P2": p2, "P3": p3}

    if st.button("🚀 XUẤT ĐỀ THI ĐÚNG CẤU TRÚC", type="primary", use_container_width=True):
        final_selected = {"P1": [], "P2": [], "P3": []}
        
        # Bốc câu hỏi
        for fname, cfg in configs.items():
            for p in ["P1", "P2", "P3"]:
                pool = []
                for d in ["NB", "TH", "VD", "VDC"]:
                    pool.extend(st.session_state.bank[fname][p][d])
                if len(pool) < cfg[p]:
                    st.warning(f"File {fname} không đủ câu cho {p}")
                    final_selected[p].extend(pool)
                else:
                    final_selected[p].extend(random.sample(pool, cfg[p]))

        # Tạo file kết quả
        output = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(files[0].getvalue()), 'r') as zin:
            with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        dom = minidom.parseString(zin.read(item.filename).decode('utf-8'))
                        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                        for child in list(body.childNodes):
                            if child.nodeType == 1: body.removeChild(child)
                        
                        # Chèn từng phần
                        titles = {
                            "P1": "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.",
                            "P2": "PHẦN II. Câu trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.",
                            "P3": "PHẦN III. Câu trắc nghiệm trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu 6."
                        }
                        
                        for p in ["P1", "P2", "P3"]:
                            if final_selected[p]:
                                body.appendChild(create_heading_xml(titles[p], dom))
                                random.shuffle(final_selected[p])
                                for idx, q_blocks in enumerate(final_selected[p]):
                                    # Đánh số lại Câu
                                    f_txt = get_text(q_blocks[0])
                                    for t in q_blocks[0].getElementsByTagNameNS(W_NS, "t"):
                                        if t.firstChild and "Câu" in t.firstChild.nodeValue:
                                            t.firstChild.nodeValue = re.sub(r'Câu\s*\d+', f"Câu {idx+1}", t.firstChild.nodeValue)
                                            t.firstChild.nodeValue = re.sub(r'#(NB|TH|VD|VDC)', '', t.firstChild.nodeValue)
                                    for b in q_blocks: body.appendChild(b)
                        
                        zout.writestr(item, dom.toxml().encode('utf-8'))
                    else:
                        zout.writestr(item, zin.read(item.filename))
        
        st.success("✅ Đã gộp đề thành công theo cấu trúc 3 phần!")
        st.download_button("📥 Tải đề chuẩn (.docx)", output.getvalue(), "De_Thi_Chuan_Cau_Truc.docx")
