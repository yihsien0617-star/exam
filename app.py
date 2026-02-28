import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import requests
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

# --- 2. 全自動探測與 API 呼叫引擎 ---
def parse_with_ai_rest(raw_text, api_key):
    # 【步驟 A】先向 Google 查詢這把金鑰「真正」能用的模型清單
    list_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    list_resp = requests.get(list_url)
    
    if list_resp.status_code != 200:
        raise ValueError(f"無法驗證 API Key 或取得模型列表：{list_resp.text}")
        
    models_data = list_resp.json().get('models', [])
    # 過濾出支援「生成內容 (generateContent)」的模型
    available_models = [m['name'] for m in models_data if 'generateContent' in m.get('supportedGenerationMethods', [])]
    
    # 【步驟 B】自動挑選最好的模型 (依序向下找，保證一定能中一個)
    target_model = None
    preferences = ['models/gemini-1.5-flash-latest', 'models/gemini-1.5-flash', 'models/gemini-1.5-pro-latest', 'models/gemini-pro']
    
    for pref in preferences:
        if pref in available_models:
            target_model = pref
            break
            
    # 如果真的都沒有上面的首選，就隨便抓一個能用的
    if not target_model:
        if available_models:
            target_model = available_models[0]
        else:
            raise ValueError("您的 API Key 尚未開通任何支援文本生成的模型權限，請重新申請一把金鑰。")

    # 在畫面右下角彈出小提示，告訴我們系統選了哪一個
    st.toast(f"🤖 系統自動為您配對模型：{target_model}")
    
    # 【步驟 C】組合最終的 API 網址並發送請求
    url = f"https://generativelanguage.googleapis.com/v1beta/{target_model}:generateContent?key={api_key}"
    headers = {'Content-Type': 'application/json'}
    
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
    
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1, # 降低隨機性
            # 為了相容較舊的 gemini-pro 模型，這裡不強制鎖死 JSON 模式，改由 prompt 控制
        }
    }
    
    response = requests.post(url, headers=headers, json=payload)
    
    if response.status_code != 200:
        raise ValueError(f"API 連線失敗 (錯誤碼 {response.status_code})：{response.text}")
        
    data = response.json()
    
    try:
        content_text = data['candidates'][0]['content']['parts'][0]['text']
        
        # 清理可能殘留的 Markdown 標籤
        clean_text = content_text.strip()
        if clean_text.startswith("```json"):
            clean_text = clean_text[7:]
        elif clean_text.startswith("```"):
            clean_text = clean_text[3:]
        if clean_text.endswith("```"):
            clean_text = clean_text[:-3]
            
        return json.loads(clean_text.strip())
        
    except Exception as e:
        raise ValueError(f"解析 JSON 失敗：{e}\n回傳內容：{data}")

# --- 3. 網頁介面設計 ---
st.set_page_config(page_title="AI 題庫智慧轉檔", page_icon="🤖", layout="wide")

st.title("🤖 題庫：AI 智慧轉檔工具 (全自動探測版)")
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
                with st.spinner('AI 正在尋找最佳模型並理解題庫中...'):
                    try:
                        file_stream = io.BytesIO(uploaded_file.read())
                        raw_text = extract_raw_text(file_stream)
                        parsed_data = parse_with_ai_rest(raw_text, api_key)
                        
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
