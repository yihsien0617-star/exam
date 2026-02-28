import streamlit as st
import docx
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import re
import json
import io

# --- 1. 終極版：遞迴透視掃描工具 (保證一字不漏且順序正確) ---
def iter_block_items(parent):
    """
    遞迴讀取 Word 文件中所有段落和表格內容，保證順序完全一致。
    """
    if isinstance(parent, docx.document.Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            p = Paragraph(child, parent)
            if p.text.strip():
                yield p.text.strip()
        elif isinstance(child, CT_Tbl):
            table = Table(child, parent)
            for row in table.rows:
                for cell in row.cells:
                    # 遞迴進入表格的每一個儲存格
                    for text in iter_block_items(cell):
                        yield text

# --- 2. 核心解析引擎 (增強容錯能力) ---
def parse_exam_docx(file_stream):
    doc = docx.Document(file_stream)
    questions = []
    current_q = None
    
    # 擷取題號、題目與選項
    q_pattern = re.compile(r'^(?:\(([A-E])\)\s*)?(\d+)[\.、]\s*(.*)')
    opt_pattern = re.compile(r'\(([A-E])\)\s*([^()]+?)(?=\([A-E]\)|$)')
    
    # 使用終極掃描工具讀取整份文件
    for text in iter_block_items(doc):
            
        # 1. 抓取題目
        q_match = q_pattern.match(text)
        if q_match:
            if current_q:
                questions.append(current_q)
            ans, num, q_text = q_match.groups()
            current_q = {
                "question_number": int(num),
                "question_text": q_text.strip(),
                "answer": ans if ans else "未提供",
                "options": {},
                "explanation": "",
                "tags": {}
            }
            continue
            
        # 2. 抓取選項 (為了避免誤判，只有在還沒抓到解析時才抓選項)
        opt_matches = opt_pattern.findall(text)
        if opt_matches and current_q and not current_q["explanation"]:
            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
            continue
            
        # 3. 抓取解析與標籤
        if current_q:
            # 【關鍵修復】：使用 re.search 取代 re.match，這樣就算前面有逗號 ",解析:" 也能精準抓到！
            if re.search(r'解\s*析\s*[:：]', text):
                # 把 "解析:" 以及它前面的所有雜訊 (包含逗號) 全部清空
                raw_exp = re.sub(r'^.*?解\s*析\s*[:：]\s*', '', text)
                
                # 切割後面的難度與再現性標籤
                parts = re.split(r'","|",\s*"|"', raw_exp)
                current_q["explanation"] = parts[0].strip()
                
                for part in parts[1:]:
                    if ":" in part or "：" in part:
                        key_val = re.split(r'[:：]', part, 1)
                        if len(key_val) == 2:
                            k = key_val[0].replace(" ", "").strip()
                            v = key_val[1].split('(')[0].strip()
                            current_q["tags"][k] = v
                continue
                
            # 4. 處理跨行文字
            if not current_q["options"] and not current_q["explanation"]:
                # 選項還沒出來前，當作題目的延伸
                current_q["question_text"] += "\n" + text
            elif current_q["explanation"]:
                # 解析已經出來了，後面的文字通通當作解析的延伸
                current_q["explanation"] += "\n" + text

    # 收尾最後一題
    if current_q:
        questions.append(current_q)
        
    return questions

# --- 3. 網頁介面設計 ---
st.set_page_config(page_title="醫檢師國考題庫轉檔工具", page_icon="🧬", layout="wide")

st.title("🧬 醫檢師國考題庫：Word 轉 JSON 工具")
st.markdown("請上傳具有標準格式的國考解析 Word 檔 (`.docx`)，系統將自動擷取題目、選項、答案與解析。")

# 建立兩欄式排版
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("📁 檔案上傳區")
    uploaded_file = st.file_uploader("拖曳或選擇 Word 檔案", type=['docx'])
    
    if uploaded_file is not None:
        with st.spinner('正在解析檔案中...'):
            file_stream = io.BytesIO(uploaded_file.read())
            
            try:
                parsed_data = parse_exam_docx(file_stream)
                st.success(f"✅ 解析成功！共擷取 {len(parsed_data)} 道題目。")
                
                json_str = json.dumps(parsed_data, ensure_ascii=False, indent=4)
                
                st.download_button(
                    label="📥 下載 JSON 題庫檔",
                    data=json_str,
                    file_name=uploaded_file.name.replace(".docx", ".json"),
                    mime="application/json",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ 解析過程中發生錯誤：{e}")

with col2:
    st.subheader("🔍 解析結果即時預覽")
    if uploaded_file is not None and 'parsed_data' in locals():
        tab_preview, tab_json = st.tabs(["畫面預覽", "JSON 原始碼"])
        
        with tab_preview:
            preview_limit = min(10, len(parsed_data))
            st.info(f"僅顯示前 {preview_limit} 題預覽...")
            
            for i in range(preview_limit):
                q = parsed_data[i]
                with st.container(border=True):
                    st.markdown(f"**第 {q['question_number']} 題：{q['question_text']}**")
                    for opt, text in q['options'].items():
                        st.write(f"({opt}) {text}")
                    
                    st.success(f"**標準答案：** {q['answer']}")
                    st.info(f"**解析：** {q['explanation']}")
                    if q['tags']:
                        st.write("**標籤：**", q['tags'])
                        
        with tab_json:
            st.json(parsed_data)
    else:
        st.info("請先於左側上傳檔案，此處將顯示預覽結果。")
