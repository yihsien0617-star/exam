import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import re
import json
import io

# --- 1. 抽取純文字引擎 ---
def extract_raw_text(file_stream):
    """將 Word 檔內所有文字（含表格與段落）按順序抽出"""
    doc = docx.Document(file_stream)
    lines = []
    
    for element in doc.element.body:
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            if para.text.strip():
                lines.append(para.text.strip())
        elif element.tag.endswith('tbl'):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        # 將表格內的換行也攤平
                        for line in cell.text.split('\n'):
                            if line.strip():
                                lines.append(line.strip())
    return lines

# --- 2. 標籤與解析淨化工具 ---
def extract_tags_and_clean(text, current_tags):
    """從解析字串中精準分離出難度、再現性，並去除多餘符號"""
    # 尋找難度
    diff_match = re.search(r'難\s*度[:：]\s*([^\(,"”]+)', text)
    if diff_match:
        current_tags["難度"] = diff_match.group(1).strip()
        
    # 尋找再現性
    rep_match = re.search(r'再\s*現\s*性[:：]\s*([^\(,"”]+)', text)
    if rep_match:
        current_tags["再現性"] = rep_match.group(1).strip()
        
    # 將解析內文中的標籤與後面的選項說明 (極低, 低度...) 徹底刪除
    clean_exp = re.sub(r'[,"]*\s*難\s*度[:：][^"]+"?', '', text)
    clean_exp = re.sub(r'[,"]*\s*再\s*現\s*性[:：][^"]+"?', '', clean_exp)
    clean_exp = re.sub(r'\([^\)]*極低[^\)]*\)', '', clean_exp)
    clean_exp = re.sub(r'\([^\)]*非常簡單[^\)]*\)', '', clean_exp)
    
    # 清除前後殘留的引號與逗號
    return clean_exp.strip('", '), current_tags

# --- 3. 核心精準解析引擎 ---
def parse_unified_format(lines):
    questions = []
    current_q = None
    
    # 擷取題號、答案與題目 (例如: (D) 1. 題目...)
    q_pattern = re.compile(r'^\s*\(([A-E])\)\s*(\d+)[\.、]\s*(.*)')
    # 擷取選項 (例如: (A) 選項內容)
    opt_pattern = re.compile(r'\(([A-E])\)\s*([^()]+?)(?=\([A-E]\)|$)')
    
    for line in lines:
        clean_line = line.strip()
        
        # [步驟 A] 判斷是否為新題目
        q_match = q_pattern.match(clean_line)
        if q_match:
            if current_q:
                current_q["explanation"] = current_q["explanation"].strip()
                questions.append(current_q)
            
            ans, num, q_text = q_match.groups()
            current_q = {
                "question_number": int(num),
                "question_text": q_text.strip(),
                "answer": ans,
                "options": {},
                "explanation": "",
                "tags": {}
            }
            continue
            
        if not current_q:
            continue
            
        # [步驟 B] 判斷是否為選項
        opt_matches = opt_pattern.findall(clean_line)
        if opt_matches and not current_q["explanation"]:
            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
            continue
            
        # [步驟 C] 判斷是否進入解析區塊
        if "解  析:" in clean_line or "解析:" in clean_line or "解析：" in clean_line:
            # 移除開頭的「解析:」字眼
            exp_text = re.sub(r'^.*?(?:解\s*析)[:：]\s*', '', clean_line)
            
            # 呼叫淨化工具，把標籤抽出來，留下乾淨的解析
            clean_exp, updated_tags = extract_tags_and_clean(exp_text, current_q["tags"])
            current_q["tags"] = updated_tags
            current_q["explanation"] += clean_exp + "\n"
            continue
            
        # [步驟 D] 處理跨行文字
        if not current_q["options"] and not current_q["explanation"]:
            # 選項還沒出現，歸類為題幹的延伸
            current_q["question_text"] += "\n" + clean_line
        elif current_q["explanation"]:
            # 解析已經出現，歸類為解析的延伸
            # 預防標籤掉到下一行的情況
            clean_exp, updated_tags = extract_tags_and_clean(clean_line, current_q["tags"])
            current_q["tags"] = updated_tags
            
            # 如果這行只有標籤，淨化後會變成空字串，就不加入解析中
            if clean_exp:
                current_q["explanation"] += clean_exp + "\n"

    # 收尾最後一題
    if current_q:
        current_q["explanation"] = current_q["explanation"].strip()
        questions.append(current_q)
        
    return questions

# --- 4. 網頁介面設計 ---
st.set_page_config(page_title="國考題庫極速轉檔", page_icon="⚡", layout="wide")

st.title("⚡ 國考題庫：極速精準轉檔工具 (統一格式專用)")
st.markdown("此版本專為統一格式之題庫設計，**免連網、免 API Key，100% 本地極速處理**。")

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("📁 檔案上傳區")
    uploaded_file = st.file_uploader("上傳已統一格式的 Word 檔案 (.docx)", type=['docx'])
    
    if uploaded_file is not None:
        with st.spinner('正在極速解析中...'):
            try:
                file_stream = io.BytesIO(uploaded_file.read())
                
                # 執行解析
                lines = extract_raw_text(file_stream)
                parsed_data = parse_unified_format(lines)
                
                st.session_state['parsed_data'] = parsed_data
                st.session_state['file_name'] = uploaded_file.name
                
                st.success(f"✅ 解析完成！共完美擷取 {len(parsed_data)} 道題目。")
                
            except Exception as e:
                st.error(f"❌ 發生錯誤：{e}")

    # 下載按鈕區塊
    if 'parsed_data' in st.session_state:
        json_str = json.dumps(st.session_state['parsed_data'], ensure_ascii=False, indent=4)
        st.download_button(
            label="📥 下載完整 JSON 題庫檔",
            data=json_str,
            file_name=st.session_state['file_name'].replace(".docx", ".json"),
            mime="application/json",
            use_container_width=True
        )

with col2:
    st.subheader("🔍 解析結果即時預覽")
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
        st.info("請上傳檔案以檢視結果。")
