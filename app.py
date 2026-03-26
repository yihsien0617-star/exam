import streamlit as st
import docx
import json
import re
import io

# --- 內部輔助函數 ---
def normalize_text(text):
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    text = re.sub(r'[\s\u3000]+', ' ', text)
    return text.strip()

def _extract_options_v3(q_dict):
    raw = q_dict.pop("_raw_text", "")
    match_A = re.search(r'\(\s*[A]\s*\)', raw)
    if match_A:
        q_dict["question_text"] = raw[:match_A.start()].strip()
        opts_text = raw[match_A.start():]
        opt_pattern = re.compile(r'\(\s*(?P<key>[A-E])\s*\)\s*(?P<val>.*?)(?=(?:\(\s*[A-E]\s*\))|$)')
        options = {}
        for m in opt_pattern.finditer(opts_text):
            options[m.group('key')] = m.group('val').strip()
        q_dict["options"] = options
    else:
        q_dict["question_text"] = raw.strip()
        q_dict["options"] = {}

# --- 網頁介面開始 ---
st.set_page_config(page_title="國考 Word 轉 JSON 神器", page_icon="⚙️")

st.title("⚙️ 國考 Word 題庫轉檔神器 (V3 終極版)")
st.info("請將整理好的 Word 考題上傳，系統會自動將其轉換為國考平台專用的 JSON 格式。")

# 檔案上傳區塊
uploaded_file = st.file_uploader("📂 請選擇您的 Word 檔 (.docx)", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 開始全自動轉換", type="primary", use_container_width=True):
        with st.spinner("正在努力解析 Word 文件中..."):
            try:
                # 讀取上傳的 Word 檔
                doc = docx.Document(uploaded_file)
                
                questions = []
                current_year = "未知年份"
                current_q = None
                skipped_lines = []

                year_pattern = re.compile(r'(\d{2,4})\s*年')
                q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
                exp_pattern = re.compile(r'^[\"\'\,\.\-\s]*解\s*析\s*[:\s](.*)', re.IGNORECASE)

                for para in doc.paragraphs:
                    text = normalize_text(para.text)
                    if not text:
                        continue
                        
                    year_match = year_pattern.search(text)
                    if year_match and not q_start_pattern.search(text): 
                        current_year = text.replace('"', '').replace(',', '').strip()
                        continue
                        
                    q_match = q_start_pattern.match(text)
                    if q_match:
                        if current_q:
                            _extract_options_v3(current_q)
                            questions.append(current_q)
                        
                        ans = q_match.group('ans').strip().upper().replace('，', ',')
                        try:
                            q_num = int(q_match.group('num'))
                        except:
                            q_num = 0
                            
                        q_text = q_match.group('text').strip()
                        
                        current_q = {
                            "question_number": q_num,
                            "answer": ans,
                            "explanation": "",
                            "tags": {
                                "年份": current_year,
                                "主題": "未分類" # 這裡保留了主題欄位讓您後續分類
                            },
                            "_raw_text": q_text
                        }
                        continue
                        
                    exp_match = exp_pattern.match(text)
                    if exp_match and current_q:
                        current_q["explanation"] = exp_match.group(1).strip()
                        continue
                        
                    if current_q:
                        hidden_exp_match = re.search(r'[\"\'\,\.\-\s]*解\s*析\s*[:\s]', text, re.IGNORECASE)
                        if hidden_exp_match:
                            split_idx = hidden_exp_match.start()
                            q_part = text[:split_idx].strip()
                            exp_part = text[hidden_exp_match.end():].strip()
                            if q_part:
                                current_q["_raw_text"] += "\n" + q_part
                            current_q["explanation"] += exp_part
                            continue

                        if current_q["explanation"]:
                            current_q["explanation"] += "\n" + text
                        else:
                            current_q["_raw_text"] += "\n" + text
                    else:
                        if len(text) > 5:
                            skipped_lines.append(f"[{current_year}] {text}")

                if current_q:
                    _extract_options_v3(current_q)
                    questions.append(current_q)

                # --- 轉換完成，準備輸出網頁畫面 ---
                if questions:
                    st.success(f"🎉 轉換大功告成！系統共成功辨識了 **{len(questions)}** 題！")
                    
                    # 將 JSON 轉為字串並提供下載按鈕
                    json_str = json.dumps(questions, ensure_ascii=False, separators=(',', ':'))
                    
                    st.download_button(
                        label="💾 點我下載轉換後的 JSON 題庫檔",
                        data=json_str,
                        file_name=uploaded_file.name.replace(".docx", "_修復版.json"),
                        mime="application/json",
                        type="primary",
                        use_container_width=True
                    )
                    
                    # 顯示漏網之魚報告在網頁上
                    if skipped_lines:
                        st.warning("⚠️ 系統發現以下文字不符合標準格式，已被自動略過。請檢查這之中是否有您漏掉的題目：")
                        with st.expander("👀 點我查看漏題報告細節"):
                            for line in skipped_lines:
                                st.code(line)
                else:
                    st.error("😭 轉換失敗，系統沒有在檔案中找到任何符合標準格式的題目。")

            except Exception as e:
                st.error(f"系統發生錯誤：{e}")
