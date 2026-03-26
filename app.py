import streamlit as st
import docx
import json
import re

# --- 內部輔助函數 ---
def normalize_line(text):
    """處理單行文字，解決全半形問題"""
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    text = re.sub(r'[ \t\u3000]+', ' ', text)
    return text.strip()

def extract_and_remove_tags(text, q_dict):
    """🌟 全域智慧標籤清洗器：不管標籤藏在哪，通通吸出來"""
    if not text: return text
    
    # 1. 抓取被引號包圍的標籤 (例如: "難  度: 適中(非常簡單...)" 或 "主題: 腫瘤")
    pattern_quoted = r'[\"\'”’]\s*([^\"\'”’]+?)\s*[:]\s*([^\"\'”’]+?)\s*[\"\'”’]'
    for m in re.finditer(pattern_quoted, text):
        raw_key = m.group(1).replace(" ", "")
        raw_val = m.group(2).strip()
        # 清洗掉 (非常簡單...) 等無用註解
        clean_val = re.sub(r'\(.*?\)|（.*?）', '', raw_val).strip()
        q_dict["tags"][raw_key] = clean_val
        
    text = re.sub(pattern_quoted, '', text).strip()
    
    # 2. 抓取無引號的常規標籤
    keywords = ["難度", "再現性", "主題", "分類", "單元", "章節"]
    for kw in keywords:
        kw_regex = kw[0] + r'\s*' + kw[1:] if len(kw) >= 2 else kw
        pattern_plain = r'(?:^|[\n\s,，])(' + kw_regex + r')\s*[:]\s*([^,，\n]+)'
        for m in re.finditer(pattern_plain, text):
            raw_key = kw
            raw_val = m.group(2).strip()
            clean_val = re.sub(r'\(.*?\)|（.*?）', '', raw_val).strip()
            q_dict["tags"][raw_key] = clean_val
            text = text.replace(m.group(0), '')
            
    # 清理殘留的標點符號
    text = re.sub(r'^[,\"\'\s，]+|[,\"\'\s，]+$', '', text)
    return text.strip()

def _extract_tags_from_all(q_dict):
    """掃描解析區與選項區，防止標籤黏在選項 D 後面"""
    # 掃描解析
    q_dict["explanation"] = extract_and_remove_tags(q_dict.get("explanation", ""), q_dict)
    
    # 掃描每一個選項
    for k, v in q_dict.get("options", {}).items():
        q_dict["options"][k] = extract_and_remove_tags(v, q_dict)
        
    # 🌟 統一映射到「主題」：讓前台的下拉選單抓得到
    for k in ["分類", "單元", "章節"]:
        if k in q_dict["tags"] and "主題" not in q_dict["tags"]:
            q_dict["tags"]["主題"] = q_dict["tags"][k]

def _extract_options_v6(q_dict):
    """跨行選項捕捉器"""
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

st.title("⚙️ 國考 Word 題庫轉檔神器 (V6 全域掃描版)")
st.info("請將整理好的 Word 考題上傳。支援全域標籤掃描，無論分類標籤藏在哪裡都能完美萃取。")

uploaded_file = st.file_uploader("📂 請選擇您的 Word 檔 (.docx)", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 開始全自動轉換", type="primary", use_container_width=True):
        with st.spinner("正在執行全域防彈掃描與標籤萃取..."):
            try:
                doc = docx.Document(uploaded_file)
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
                    # 1. 偵測明確的主題標題
                    topic_match = re.match(r'^【?(主題|單元|分類|章節)】?[:\s]*(.+)$', text)
                    if topic_match and not q_start_pattern.search(text):
                        if current_q:
                            _extract_options_v6(current_q)
                            _extract_tags_from_all(current_q)
                            questions.append(current_q)
                            current_q = None # 🌟 關鍵修復：切斷上一題，防止主題被吞掉
                        current_topic = topic_match.group(2).strip()
                        continue

                    # 2. 偵測年份標題
                    year_match = year_pattern.search(text)
                    if year_match and not q_start_pattern.search(text): 
                        if current_q:
                            _extract_options_v6(current_q)
                            _extract_tags_from_all(current_q)
                            questions.append(current_q)
                            current_q = None # 🌟 關鍵修復：切斷上一題
                        current_year = text.replace('"', '').replace(',', '').strip()
                        continue
                        
                    # 3. 抓題目
                    q_match = q_start_pattern.match(text)
                    if q_match:
                        if current_q:
                            _extract_options_v6(current_q)
                            _extract_tags_from_all(current_q)
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
                                "主題": current_topic # 預設套用目前抓到的主題
                            },
                            "_raw_text": q_text
                        }
                        continue
                        
                    # 4. 抓解析
                    exp_match = exp_pattern.match(text)
                    if exp_match and current_q:
                        current_q["explanation"] = exp_match.group(1).strip()
                        continue
                        
                    # 5. 多行文字串接
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
                        # 🌟 智慧判斷無標記的短句作為單元主題
                        if 2 <= len(text) <= 30 and not text.startswith("(") and not text.startswith("["):
                            current_topic = text.strip()
                        elif len(text) > 5:
                            skipped_lines.append(f"[{current_year}] {text}")

                # 收尾最後一題
                if current_q:
                    _extract_options_v6(current_q)
                    _extract_tags_from_all(current_q)
                    questions.append(current_q)

                if questions:
                    st.success(f"🎉 轉換大功告成！系統共成功辨識了 **{len(questions)}** 題！")
                    
                    json_str = json.dumps(questions, ensure_ascii=False, separators=(',', ':'))
                    
                    st.download_button(
                        label="💾 點我下載完美修復版 JSON",
                        data=json_str,
                        file_name=uploaded_file.name.replace(".docx", "_V6_全域掃描版.json"),
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
