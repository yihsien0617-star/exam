import streamlit as st
import docx
import json
import re
import io

# --- 內部輔助函數 ---
def normalize_line(text):
    """處理單行文字，解決全半形問題，但絕對不破壞換行"""
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    # 只把連續的「空白鍵」縮減，保留 \n
    text = re.sub(r'[ \t\u3000]+', ' ', text)
    return text.strip()

def _extract_tags_from_explanation(q_dict):
    """🌟 智慧標籤清洗器：自動剔除 (非常簡單, 簡單...) 等雜訊"""
    exp = q_dict.get("explanation", "")
    if not exp: return
    
    # 定義要抓取的標籤與關鍵字
    keywords = [("難度", r'難\s*度'), ("再現性", r'再現性'), ("主題", r'主\s*題'), ("分類", r'分\s*類')]
    
    for tag_name, kw_regex in keywords:
        # 尋找像是 難度: 適中(非常簡單...) 的格式
        pattern = kw_regex + r'\s*[:]\s*(.*?)(?=\"|\n|$)'
        match = re.search(pattern, exp)
        if match:
            val = match.group(1).strip()
            # 神級防呆：把 (非常簡單, 簡單...) 這種操作說明直接刪除
            val = re.sub(r'\(.*?\)', '', val).strip()
            val = val.rstrip(',').strip()
            
            q_dict["tags"][tag_name] = val
            
            # 從解析本體中把這段標籤宣告刪除，保持解析乾淨
            remove_pat = r'[\"\'\,]?\s*' + kw_regex + r'\s*[:]\s*.*?(?=\"|\n|$)[\"\']?\,?'
            exp = re.sub(remove_pat, '', exp, count=1)
            
    # 清理殘留的標點符號
    exp = re.sub(r'^[,\"\'\s]+|[,\"\'\s]+$', '', exp)
    q_dict["explanation"] = exp.strip()

def _extract_options_v5(q_dict):
    """🌟 跨行選項捕捉器"""
    raw = q_dict.pop("_raw_text", "")
    
    match_A = re.search(r'\(\s*[A]\s*\)(?=.*?\(\s*[B]\s*\))', raw, re.DOTALL)
    
    if match_A:
        q_dict["question_text"] = raw[:match_A.start()].strip()
        opts_text = raw[match_A.start():]
        
        opt_pattern = re.compile(r'\(\s*(?P<key>[A-E])\s*\)\s*(?P<val>.*?)(?=(?:\(\s*[A-E]\s*\))|$)', re.DOTALL)
        options = {}
        for m in opt_pattern.finditer(opts_text):
            options[m.group('key')] = m.group('val').replace('\n', ' ').strip()
        q_dict["options"] = options
    else:
        q_dict["question_text"] = raw.strip()
        q_dict["options"] = {}

# --- 網頁介面開始 ---
st.set_page_config(page_title="國考 Word 轉 JSON 神器", page_icon="⚙️")

st.title("⚙️ 國考 Word 題庫轉檔神器 (V5 防彈解析版)")
st.info("請將整理好的 Word 考題上傳。支援 Shift+Enter 密集排版，自動清洗難度與分類標籤。")

uploaded_file = st.file_uploader("📂 請選擇您的 Word 檔 (.docx)", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 開始全自動轉換", type="primary", use_container_width=True):
        with st.spinner("正在強制斷行掃描並萃取分類標籤..."):
            try:
                doc = docx.Document(uploaded_file)
                
                # 🌟 核心防呆：強制將所有段落打散成獨立的行，徹底破解 Shift+Enter 造成的黏連
                all_lines = []
                for para in doc.paragraphs:
                    for line in para.text.split('\n'):
                        clean_line = normalize_line(line)
                        if clean_line:
                            all_lines.append(clean_line)

                questions = []
                current_year = "未知年份"
                current_topic = "未分類"
                current_q = None
                skipped_lines = []

                year_pattern = re.compile(r'(\d{2,4})\s*年')
                q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
                exp_pattern = re.compile(r'^[\"\'\,\.\-\s]*解\s*析\s*[:](.*)', re.IGNORECASE)

                for text in all_lines:
                    # 1. 抓年份
                    year_match = year_pattern.search(text)
                    if year_match and not q_start_pattern.search(text): 
                        current_year = text.replace('"', '').replace(',', '').strip()
                        continue
                        
                    # 2. 抓題目
                    q_match = q_start_pattern.match(text)
                    if q_match:
                        if current_q:
                            _extract_options_v5(current_q)
                            _extract_tags_from_explanation(current_q)
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
                                "主題": current_topic
                            },
                            "_raw_text": q_text
                        }
                        continue
                        
                    # 3. 抓解析
                    exp_match = exp_pattern.match(text)
                    if exp_match and current_q:
                        current_q["explanation"] = exp_match.group(1).strip()
                        continue
                        
                    # 4. 多行文字串接與單元標題偵測
                    if current_q:
                        hidden_exp_match = re.search(r'[\"\'\,\.\-\s]*解\s*析\s*[:]', text, re.IGNORECASE)
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
                        if 2 < len(text) < 30 and not text.startswith("(") and not text.startswith("["):
                            current_topic = text.strip()
                        elif len(text) > 5:
                            skipped_lines.append(f"[{current_year}] {text}")

                # 收尾最後一題
                if current_q:
                    _extract_options_v5(current_q)
                    _extract_tags_from_explanation(current_q)
                    questions.append(current_q)

                if questions:
                    st.success(f"🎉 轉換大功告成！系統共成功辨識了 **{len(questions)}** 題！")
                    
                    json_str = json.dumps(questions, ensure_ascii=False, separators=(',', ':'))
                    
                    st.download_button(
                        label="💾 點我下載完美修復版 JSON",
                        data=json_str,
                        file_name=uploaded_file.name.replace(".docx", "_V5_防彈版.json"),
                        mime="application/json",
                        type="primary",
                        use_container_width=True
                    )
                    
                    if skipped_lines:
                        st.warning("⚠️ 以下文字無法被辨識為題目，若有漏題請檢查 Word 排版：")
                        with st.expander("👀 點我查看忽略文字"):
                            for line in skipped_lines:
                                st.code(line)
                else:
                    st.error("😭 轉換失敗，找不到任何符合格式的題目。")

            except Exception as e:
                st.error(f"系統發生錯誤：{e}")
