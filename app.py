import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import google.generativeai as genai
import json
import io

# --- 1. 暴力抽取純文字引擎 ---
def extract_raw_text(file_stream):
    """將 Word 檔內所有文字（含表格）暴力抽出，不保留任何排版結構"""
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
                        # 將表格內的換行也攤平
                        for line in cell.text.split('\n'):
                            if line.strip():
                                raw_text.append(line.strip())
                                
    return "\n".join(raw_text)

# --- 2. 呼叫 AI 進行語意轉換 ---
def parse_with_ai(raw_text, api_key):
    """將純文字交給 Gemini 進行語意分析與 JSON 結構化"""
    # 設定 API Key
    genai.configure(api_key=api_key)
    
    # 選擇速度快、適合處理大量文字的 Flash 模型
model = genai.GenerativeModel(
        model_name="gemini-1.5-flash",  # 👈 換回這個最新版的模型
        generation_config={
            "temperature": 0.1,
            "response_mime_type": "application/json",
        }
    )
    
    prompt = f"""
    你是一個專業的醫事檢驗師國考題庫資料處理專家。
    請將以下的「國考題庫原始混亂文本」，轉換成結構化的 JSON 陣列 (JSON Array)。
    原始文本包含了題目、選項(A, B, C, D)、答案、老師解析以及難度等標籤。
    排版可能極度混亂（例如選項沒對齊、解析前面有奇怪符號等），請透過你的語意理解能力來精準擷取。

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
        "explanation": "老師解析內容(請完整保留，若無則留空)",
        "tags": {{
          "難度": "簡單/適中/困難 (若文本有提供)",
          "再現性": "高度/中度/低度 (若文本有提供)"
        }}
      }}
    ]

    請注意：
    1. 確保所有題目都被完整擷取，不可遺漏。
    2. 解析區塊通常緊跟在題目與選項之後，請仔細尋找「解析」字眼。
    
    以下是原始文本：
    ---
    {raw_text}
    """
    
    response = model.generate_content(prompt)
    
    # 將 AI 回傳的字串轉換回 Python 字典格式
    try:
        json_data = json.loads(response.text)
        return json_data
    except Exception as e:
        raise ValueError(f"AI 回傳的格式非有效 JSON，解析失敗：{e}\n\nAI 回傳內容：{response.text}")

# --- 3. 網頁介面設計 ---
st.set_page_config(page_title="AI 國考題庫智慧轉檔", page_icon="🤖", layout="wide")

st.title("🤖 醫檢師國考題庫：AI 智慧轉檔工具")
st.markdown("無需強迫其他老師修改 Word 格式！直接上傳檔案，讓 AI 幫您讀懂語意並產出資料庫格式。")

# 側邊欄：設定 API Key
with st.sidebar:
    st.header("⚙️ 系統設定")
    st.info("請輸入您的 Google Gemini API Key 以啟用 AI 轉換功能。")
    api_key = st.text_input("Gemini API Key", type="password")
    st.markdown("[👉 點此免費獲取 API Key](https://aistudio.google.com/app/apikey)")

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
                        # 步驟 1：暴力抽取純文字
                        file_stream = io.BytesIO(uploaded_file.read())
                        raw_text = extract_raw_text(file_stream)
                        
                        # 步驟 2：呼叫 AI 處理
                        parsed_data = parse_with_ai(raw_text, api_key)
                        
                        st.session_state['parsed_data'] = parsed_data
                        st.session_state['file_name'] = uploaded_file.name
                        st.success(f"✅ AI 破解成功！共整理出 {len(parsed_data)} 道題目。")
                        
                    except Exception as e:
                        st.error(f"❌ 發生錯誤：{e}")

    # 下載按鈕區塊
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
            st.info(f"預覽前 10 題，請檢查 AI 是否正確理解了其他老師的排版邏輯。")
            preview_limit = min(10, len(parsed_data))
            
            for i in range(preview_limit):
                q = parsed_data[i]
                with st.container(border=True):
                    st.markdown(f"**第 {q.get('question_number', '?')} 題：{q.get('question_text', '')}**")
                    
                    options = q.get('options', {})
                    for opt, text in options.items():
                        st.write(f"({opt}) {text}")
                    
                    st.success(f"**標準答案：** {q.get('answer', '')}")
                    
                    if q.get('explanation'):
                        st.info(f"💡 **老師解析：**\n{q['explanation']}")
                    else:
                        st.warning("⚠️ 此題未擷取到解析")
                        
                    if q.get('tags'):
                        st.write("**📝 標籤數據：**", q['tags'])
                        
        with tab_json:
            st.json(parsed_data)
    else:
        st.info("等待 AI 處理完成後，此處將顯示預覽結果。")
