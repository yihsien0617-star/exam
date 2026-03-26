import streamlit as st
import docx
import json
import re

# --- 內部輔助函數 ---
def normalize_text(text):
    """標準化全半形與多餘空白"""
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    text = re.sub(r'[\s\u3000]+', ' ', text)
    return text.strip()

def _extract_tags_from_explanation(q_dict):
    """🌟 智慧標籤萃取器：從解析中把難度、再現性、分類抽出來變成 Tags"""
    exp = q_dict.get("explanation", "")
    if not exp: return
    
    # 預期會出現的分類關鍵字
    keywords = ["難度", "難 度", "再現性", "主題", "主 題", "分類", "分 類", "章節"]
    
    for kw in keywords:
        # 尋找像是 "難度: 適中" 或是 難度: 困難 的格式
        pattern_quoted = r'[\"\'\,]?\s*(' + kw + r')\s*[:]\s*([^\"]+?)[\"\']'
        pattern_plain = r'(' + kw + r')\s*[:]\s*([^,\n]+)'
        
        for pat in [pattern_quoted, pattern_plain]:
            match = re.search(pat, exp)
            if match:
                clean_key = match.group(1).replace(" ", "") # 統一去掉空白變成"難度"
                clean_val = match.group(2).strip()
                if clean_val.endswith(','): 
                    clean_val = clean_val[:-1].strip()
                
                # 寫入 Tags 分類中
                q_dict["tags"][clean_key] = clean_val
                
                # 把抽出的字眼從解析中刪除，讓解析保持乾淨
                exp = exp[:match.start()] + exp[match.end():]
                break
                
    # 清理殘留的標點符號
    exp = re.sub(r'^[,\"\'\s]+|[,\"\'\s]+$', '', exp)
    q_dict["explanation"] = exp.strip()

def _extract_options_v4(q_dict):
    """🌟 跨行選項捕捉器：支援跨行抓取選項"""
    raw = q_dict.pop("_raw_text", "")
    
    # 尋找第一個 (A)，並確認後面有 (B)
    match_A = re.search(r'\(\s*[A]\s*\)(?=.*?\(\s*[B]\s*\))', raw, re.DOTALL)
    
    if match_A:
        q_dict["question_text"] = raw[:match_A.start()].strip()
        opts_text = raw[match_A.start():]
        
        # 加入 re.DOTALL，讓選項內容可以跨越換行符號！解決 C、D 選項不見的問題
        opt_pattern = re.compile(r'\(\s*(?P<key>[A-E])\s*\)\s*(?P<val>.*?)(?=(?:\(\s*[A-E]\s*\))|$)', re.DOTALL)
        options = {}
        for m in opt_pattern.finditer(opts_text):
            # 移除選項內部的多餘換行，保持單行整潔
            options[m.group('key')] = m.group('val').replace('\n', ' ').strip()
        q_dict["options"] = options
    else:
        q_dict["question_text"] = raw.strip()
        q_dict["options"] = {}

# --- 網頁介面開始 ---
st.set_page_config(page_title="國考 Word 轉 JSON 神器", page_icon="⚙️")

st.title("⚙️ 國考 Word 題庫轉檔神器 (V4 智慧分類版)")
st.info("請將整理好的 Word 考題上傳，系統會自動轉換、修復漏行選項，並智能萃取難度與分類。")

uploaded_file = st.file_uploader("📂 請選擇您的 Word 檔 (.docx)", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 開始全自動轉換", type="primary", use_container_width=True):
        with st.spinner("正在努力解析文件並萃取分類標籤..."):
            try:
                doc = docx.Document(uploaded_file)
                
                questions = []
                current_year = "未知年份"
                current_topic = "未分類"  # 🌟 新增預設主題
                current_q = None
                skipped_lines = []

                year_pattern = re.compile(r'(\d{2,4})\s*年')
                q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
                exp_pattern = re.compile(r'^[\"\'\,\.\-\s]*解\s*析\s*[:\s](.*)', re.IGNORECASE)

                for para in doc.paragraphs:
                    text = normalize_text(para.text)
                    if not text:
                        continue
                        
                    # 1. 抓年份
                    year_match = year_pattern.search(text)
                    if year_match and not q_start_pattern.search(text): 
                        current_year = text.replace('"', '').replace(',', '').strip()
                        continue
                        
                    # 2. 抓題目
                    q_match = q_start_pattern.match(text)
                    if q_match:
                        if current_q:
                            _extract_options_v4(current_q)
                            _extract_tags_from_explanation(current_q) # 提交前先抽標籤
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
                                "主題": current_topic  # 套用目前最新的主題
                            },
                            "_raw_text": q_text
                        }
                        continue
                        
                    # 3. 抓解析
                    exp_match = exp_pattern.match(text)
                    if exp_match and current_q:
                        current_q["explanation"] = exp_match.group(1).strip()
                        continue
                        
                    # 4. 多行文字串接，或偵測是否為單元標題
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
                        # 🌟 智慧判斷：如果不是題目、不是年份、不是解析，且長度很短，就認定為新的「主題/分類」
                        if 2 < len(text) < 30 and not text.startswith("(") and not text.startswith("["):
                            current_topic = text.strip()
                        elif len(text) > 5:
                            skipped_lines.append(f"[{current_year}] {text}")

                # 收尾最後一題
                if current_q:
                    _extract_options_v4(current_q)
                    _extract_tags_from_explanation(current_q)
                    questions.append(current_q)

                if questions:
                    st.success(f"🎉 轉換大功告成！系統共成功辨識了 **{len(questions)}** 題！")
                    
                    json_str = json.dumps(questions, ensure_ascii=False, separators=(',', ':'))
                    
                    st.download_button(
                        label="💾 點我下載完美修復版 JSON",
                        data=json_str,
                        file_name=uploaded_file.name.replace(".docx", "_V4_完美修復版.json"),
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
