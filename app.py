import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import google.generativeai as genai
import json
import io

# --- 1. 暴力抽取純文字引擎 ---
def extract_raw_text(file_stream):
    doc = docx.Document(file_stream)
    raw_text = []
    
    for element in doc.element.body:
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            if para.text.strip():
                raw_text.append(para.text.strip())
        elif element.tag.endswith('tbl'):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        for line in cell.text.split('\n'):
                            if line.strip():
                                raw_text.append(line.strip())
                                
    return "\n".join(raw_text)

# --- 2. 呼叫 AI 進行語意轉換 (全自動適應版) ---
def parse_with_ai(raw_text, api_key):
    genai.configure(api_key=api_key)
    
    # 動態查詢這把金鑰支援的所有模型
    available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    
    # 自動挑選最好的模型 (優先順序: 1.5-flash-latest > 1.5-flash > 1.5-pro > 1.0-pro)
    target_model = "gemini-pro" # 最基礎的保底模型
    use_json_mode = False
    
    if 'models/gemini-1.5-flash-latest' in available_models:
        target_model = "gemini-1.5-flash-latest"
        use_json_mode = True
    elif 'models/gemini-1.5-flash' in available_models:
        target_model = "gemini-1.5-flash"
        use_json_mode = True
    elif 'models/gemini-1.5-pro-latest' in available_models:
        target_model = "gemini-1.5-pro-latest"
        use_json_mode = True
    elif 'models/gemini-pro' in available_models:
        target_model = "gemini-pro"
        use_json_mode = False

    st.toast(f"🤖 系統自動選擇的模型：{target_model}")
    
    # 根據支援度設定參數
    generation_config = {"temperature": 0.1}
    if use_json_mode:
        generation_config["response_mime_type"] = "application/json"
        
    model = genai.GenerativeModel(
        model_name=target_model,
        generation_config=generation_config
    )
    
    prompt = f"""
    你是一個專業的題庫資料處理專家。
    請將以下的「原始混亂文本」，轉換成結構化的 JSON 陣列 (JSON Array)。
    原始文本包含了題目、選項(A, B, C, D)、答案、解析以及難度等標籤。

    必須輸出的 JSON 格式定義如下：
    [
      {{
        "question_number": 1,
        "question_text": "題目內容(不含題號與答案)",
        "answer": "標準答案(僅限 A/B/C/D，若無請填 '未提供')",
        "options": {{
          "A": "選項A內容",
          "B": "選項B內容",
          "C": "選項C內容",
          "D": "選項D內容"
        }},
        "explanation": "解析內容(請完整保留，若無則留空)",
        "tags": {{
          "難度": "簡單/適中/困難 (若文本有提供)",
          "再現性": "高度/中度/低度 (若文本有提供)"
        }}
      }}
    ]
    
    請注意：只輸出 JSON 格式，不要包含任何 ```json 的 Markdown 標籤。
    
    以下是原始文本：
    ---
    {raw_text}
    """
    
    response = model.generate_content(prompt)
    
    # 清理 AI 可能手癢加上的 Markdown 標籤
    clean_text = response.text.strip()
    if clean_text.startswith("```json"):
        clean_text = clean_text[7:]
    elif clean_text.startswith("```"):
        clean_text = clean_text[3:]
    if clean_text.endswith("```"):
        clean_text = clean_text[:-3]
        
    try:
        json_data = json.loads(clean_text.strip())
        return json_data
    except Exception as e:
        raise ValueError(f"AI 回傳格式錯誤：{e}\n\n回傳內容：{clean_text}")

# --- 3. 網頁介面設計 ---
st.set_page_config(page_title="AI 題庫智慧轉檔", page_icon="🤖", layout="wide")

st.title("🤖 題庫：AI 智慧轉檔工具")
st.markdown("無需強迫修改 Word 格式！直接上傳檔案，讓 AI 幫您讀懂語意並產出資料庫格式。")

with st.sidebar:
    st.header("⚙️ 系統設定")
    api_key = st.text_input("Gemini API Key", type="password")

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("📁 檔案上傳區")
    uploaded_file = st.file_uploader("上傳任一排版的 Word 檔案", type=['docx'])
    
    if uploaded_file is not None:
        if not api_key:
            st.warning("⚠️ 請先在左側欄輸入 API Key 才能進行轉換。")
        else:
            if st.button("🚀 啟動 AI 智慧分析"):
                with st.spinner('AI 正在閱讀並理解題庫中，這可能需要幾十秒...'):
                    try:
                        file_stream = io.BytesIO(uploaded_file.read())
                        raw_text = extract_raw_text(file_stream)
                        parsed_data = parse_with_ai(raw_text, api_key)
                        
                        st.session_state['parsed_data'] = parsed_data
                        st.session_state['file_name'] = uploaded_file.name
                        st.success(f"✅ AI 破解成功！共整理出 {len(parsed_data)} 道題目。")
                        
                    except Exception as e:
                        st.error(f"❌ 發生錯誤：{e}")

    if 'parsed_data' in st.session_state:
        json_str = json.dumps(st.session_state['parsed_data'], ensure_ascii=False, indent=4)
        st.download_button(
            label="📥 下載完整 JSON 題庫檔",
            data=json_str,
            file_name=st.session_state['file_name'].replace(".docx", "_AI.json"),
            mime="application/json",
            use_container_width=True
        )

with col2:
    st.subheader("🔍 AI 解析結果即時預覽")
    if 'parsed_data' in st.session_state:
        tab_preview, tab_json = st.tabs(["畫面預覽", "JSON 原始碼"])
        parsed_data = st.session_state['parsed_data']
        
        with tab_preview:
            preview_limit = min(10, len(parsed_data))
            st.info(f"預覽前 {preview_limit} 題...")
            
            for i in range(preview_limit):
                q = parsed_data[i]
                with st.container(border=True):
                    st.markdown(f"**第 {q.get('question_number', '?')} 題：{q.get('question_text', '')}**")
                    for opt, text in q.get('options', {}).items():
                        st.write(f"({opt}) {text}")
                    st.success(f"**標準答案：** {q.get('answer', '')}")
                    if q.get('explanation'):
                        st.info(f"💡 **解析：**\n{q['explanation']}")
                    if q.get('tags'):
                        st.write("**📝 標籤數據：**", q['tags'])
                        
        with tab_json:
            st.json(parsed_data)
    else:
        st.info("等待 AI 處理完成後，此處將顯示預覽結果。")
