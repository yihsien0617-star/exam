import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import re
import json
import io

# --- 核心解析引擎 (無差別攤平版) ---
def parse_exam_docx(file_stream):
    doc = docx.Document(file_stream)
    
    # 1. 將整份文件攤平成為一行一行的純文字陣列
    lines = []
    seen_tc = set() # 用來避免合併儲存格造成文字重複抓取
    
    for element in doc.element.body:
        # 如果是段落
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            if para.text.strip():
                # 遇到段落內自帶換行，強制切開成獨立行
                for line in para.text.split('\n'):
                    if line.strip(): lines.append(line.strip())
        # 如果是表格
        elif element.tag.endswith('tbl'):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    # 過濾掉合併儲存格產生的重複物件
                    if cell._tc in seen_tc:
                        continue
                    seen_tc.add(cell._tc)
                    if cell.text.strip():
                        # 表格內的文字也強制切成獨立行
                        for line in cell.text.split('\n'):
                            if line.strip(): lines.append(line.strip())

    # 2. 開始逐行精準捕捉
    questions = []
    current_q = None
    
    # 正規表達式 (允許前面有任何空白或雜訊)
    q_pattern = re.compile(r'^\s*(?:\(([A-E])\)\s*)?(\d+)[\.、]\s*(.*)')
    opt_pattern = re.compile(r'\(([A-E])\)\s*([^()]+?)(?=\([A-E]\)|$)')
    
    for text in lines:
        # 建立一個去除所有空白的字串，用來抵抗排版不一的干擾
        clean_text = text.replace(" ", "").replace("　", "")
        
        # [步驟 A] 抓題目
        q_match = q_pattern.match(text)
        if q_match:
            if current_q:
                # 儲存上一題前，清掉解析尾巴多餘的空白
                current_q["explanation"] = current_q["explanation"].strip()
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
            
        if not current_q:
            continue
            
        # [步驟 B] 抓選項
        opt_matches = opt_pattern.findall(text)
        if opt_matches and not current_q["explanation"]:
            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
            continue
            
        # [步驟 C] 抓解析 (只要句子裡含有解析兩字就觸發)
        if "解析:" in clean_text or "解析：" in clean_text or clean_text.startswith("解析"):
            exp_text = re.sub(r'^.*?解\s*析\s*[:：]?\s*', '', text)
            current_q["explanation"] += exp_text + "\n"
            continue
            
        # [步驟 D] 抓標籤 - 難度
        if "難度:" in clean_text or "難度：" in clean_text or clean_text.startswith("難度"):
            diff_text = re.sub(r'^.*?難\s*度\s*[:：]?\s*', '', text)
            # 自動砍掉後面的說明括號，如 "(非常簡單, 簡單...)"
            current_q["tags"]["難度"] = diff_text.split('(')[0].strip()
            continue
            
        # [步驟 E] 抓標籤 - 再現性
        if "再現性:" in clean_text or "再現性：" in clean_text or clean_text.startswith("再現性"):
            rep_text = re.sub(r'^.*?再\s*現\s*性\s*[:：]?\s*', '', text)
            current_q["tags"]["再現性"] = rep_text.split('(')[0].strip()
            continue
            
        # [步驟 F] 處理跨行文字 (題幹太長或解析太長)
        if not current_q["options"] and not current_q["explanation"] and not current_q["tags"]:
            current_q["question_text"] += "\n" + text
        elif current_q["explanation"] and not current_q["tags"]:
            # 如果已經進到解析區，且還沒遇到難度標籤，都算解析的延續
            current_q["explanation"] += text + "\n"

    # 收尾最後一題
    if current_q:
        current_q["explanation"] = current_q["explanation"].strip()
        questions.append(current_q)
        
    return questions

# --- 網頁介面設計 ---
st.set_page_config(page_title="醫檢師國考題庫轉檔工具", page_icon="🧬", layout="wide")

st.title("🧬 醫檢師國考題庫：Word 轉 JSON 工具")
st.markdown("請上傳您的解析 Word 檔 (`.docx`)，系統將自動擷取題目、選項、答案、**解析**與**難度標籤**。")

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("📁 檔案上傳區")
    uploaded_file = st.file_uploader("拖曳或選擇 Word 檔案", type=['docx'])
    
    if uploaded_file is not None:
        with st.spinner('正在使用終極掃描引擎解析中...'):
            file_stream = io.BytesIO(uploaded_file.read())
            
            try:
                parsed_data = parse_exam_docx(file_stream)
                st.success(f"✅ 破解成功！共完美擷取 {len(parsed_data)} 道題目。")
                
                json_str = json.dumps(parsed_data, ensure_ascii=False, indent=4)
                
                st.download_button(
                    label="📥 下載完整 JSON 題庫檔",
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
            st.info(f"僅顯示前 {preview_limit} 題預覽，以確認解析是否出現。")
            
            for i in range(preview_limit):
                q = parsed_data[i]
                with st.container(border=True):
                    st.markdown(f"**第 {q['question_number']} 題：{q['question_text']}**")
                    for opt, text in q['options'].items():
                        st.write(f"({opt}) {text}")
                    
                    st.success(f"**標準答案：** {q['answer']}")
                    # 特別把解析區塊用不同顏色標示
                    if q['explanation']:
                        st.info(f"💡 **老師解析：**\n{q['explanation']}")
                    else:
                        st.error("⚠️ 尚未抓取到解析")
                        
                    if q['tags']:
                        st.write("**📝 標籤數據：**", q['tags'])
                        
        with tab_json:
            st.json(parsed_data)
    else:
        st.info("請先於左側上傳檔案，此處將顯示預覽結果。")
