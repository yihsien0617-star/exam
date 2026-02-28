import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import io

# X 光掃描引擎
def xray_word_doc(file_stream):
    doc = docx.Document(file_stream)
    raw_lines = []
    
    # 逐一掃描文件中的所有元素
    for element in doc.element.body:
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            if para.text.strip():
                raw_lines.append(f"[段落] {para.text.strip()}")
        elif element.tag.endswith('tbl'):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        # 把表格內的換行也攤開
                        for line in cell.text.split('\n'):
                            if line.strip():
                                raw_lines.append(f"[表格] {line.strip()}")
    return raw_lines

# 網頁介面
st.set_page_config(page_title="Word 檔案 X 光機", layout="wide")
st.title("🔦 Word 檔案 X 光機 (除錯專用)")
st.markdown("請上傳一直抓不到解析的 Word 檔，我們來看看底層文字到底長怎樣！")

uploaded_file = st.file_uploader("上傳 Word 檔案", type=['docx'])

if uploaded_file is not None:
    file_stream = io.BytesIO(uploaded_file.read())
    lines = xray_word_doc(file_stream)
    
    st.success(f"讀取完畢！共掃描出 {len(lines)} 行文字。")
    st.write("### 🔍 Python 實際讀取到的內容：")
    
    # 將結果顯示在文字框內方便查看
    st.code("\n".join(lines), language="text")
