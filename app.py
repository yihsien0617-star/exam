import streamlit as st
import docx
import re
import json
import io

# --- 核心解析引擎 ---
def parse_exam_docx(file_stream):
    doc = docx.Document(file_stream)
    questions = []
    current_q = None
    
    # 寬容模式的正規表達式
    q_pattern = re.compile(r'^(?:\(([A-E])\)\s*)?(\d+)[\.、]\s*(.*)')
    opt_pattern = re.compile(r'\(([A-E])\)\s*([^()]+?)(?=\([A-E]\)|$)')
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
            
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
            
        opt_matches = opt_pattern.findall(text)
        if opt_matches and current_q:
            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
            continue
            
        if current_q:
            # 處理解析與標籤
            if re.match(r'^解\s*析\s*[:：]', text):
                raw_exp = re.sub(r'^解\s*析\s*[:：]\s*', '', text)
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
                
            # 處理跨行文字
            if not current_q["options"] and not current_q["explanation"]:
                current_q["question_text"] += "\n" + text
            elif current_q["explanation"] and not re.match(r'^解\s*析\s*[:：]', text):
                current_q["explanation"] += "\n" + text

    if current_q:
        questions.append(current_q)
        
    return questions

# --- 網頁介面設計 ---
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
            # 讀取上傳的檔案為記憶體串流
            file_stream = io.BytesIO(uploaded_file.read())
            
            try:
                parsed_data = parse_exam_docx(file_stream)
                st.success(f"✅ 解析成功！共擷取 {len(parsed_data)} 道題目。")
                
                # 將資料轉為 JSON 字串
                json_str = json.dumps(parsed_data, ensure_ascii=False, indent=4)
                
                # 提供下載按鈕
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
        # 使用 tabs 來切換「視覺化預覽」與「原始 JSON」
        tab_preview, tab_json = st.tabs(["畫面預覽", "JSON 原始碼"])
        
        with tab_preview:
            # 預覽前 3 題，避免畫面過長
            preview_limit = min(3, len(parsed_data))
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
