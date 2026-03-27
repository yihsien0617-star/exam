import streamlit as st
import docx
import json
import re

# --- 內部輔助函數 ---
def normalize_line(text):
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    text = re.sub(r'[ \t\u3000]+', ' ', text)
    return text.strip()

def extract_and_remove_tags(text, q_dict):
    if not text: return text
    pattern_quoted = r'[\"\'”’]\s*([^\"\'”’]+?)\s*[:]\s*([^\"\'”’]+?)\s*[\"\'”’]'
    for m in re.finditer(pattern_quoted, text):
        raw_key = m.group(1).replace(" ", "")
        raw_val = m.group(2).strip()
        clean_val = re.sub(r'\(.*?\)|（.*?）', '', raw_val).strip()
        q_dict["tags"][raw_key] = clean_val
    text = re.sub(pattern_quoted, '', text).strip()
    
    keywords = ["難度", "再現性", "主題", "分類", "單元", "章節"]
    for kw in keywords:
        kw_regex = kw[0] + r'\s*' + kw[1:] if len(kw) >= 2 else kw
        pattern_plain = r'(?:^|[\n\s,，])(' + kw_regex + r')\s*[:]\s*([^,，\n]+)'
        for m in re.finditer(pattern_plain, text):
            q_dict["tags"][kw] = re.sub(r'\(.*?\)|（.*?）', '', m.group(2)).strip()
            text = text.replace(m.group(0), '')
            
    return re.sub(r'^[,\"\'\s，]+|[,\"\'\s，]+$', '', text).strip()

def _extract_tags_from_all(q_dict):
    q_dict["explanation"] = extract_and_remove_tags(q_dict.get("explanation", ""), q_dict)
    for k, v in q_dict.get("options", {}).items():
        q_dict["options"][k] = extract_and_remove_tags(v, q_dict)
        
    for k in ["分類", "單元", "章節"]:
        if k in q_dict["tags"] and "主題" not in q_dict["tags"]:
            q_dict["tags"]["主題"] = q_dict["tags"][k]

def _extract_options(q_dict):
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

def auto_categorize(q_dict, mapping):
    """🌟 智慧關鍵字分類引擎：掃描題目與解析，全自動歸類"""
    # 如果已經有明確主題，就不覆蓋
    if q_dict["tags"].get("主題", "未分類") != "未分類":
        return
        
    # 將題目與解析合併作為搜尋範圍
    search_text = q_dict.get("question_text", "") + " " + q_dict.get("explanation", "")
    search_text = search_text.lower()
    
    for topic, keywords in mapping.items():
        for kw in keywords:
            if kw.lower() in search_text:
                q_dict["tags"]["主題"] = topic
                return  # 匹配到第一個就結束

# --- 網頁介面開始 ---
st.set_page_config(page_title="國考題庫全自動轉檔與分類系統", page_icon="⚙️", layout="wide")

st.title("⚙️ 國考題庫轉檔與自動分類系統 (V7 智慧版)")
st.write("支援未來所有專業科目！只要在下方設定關鍵字，系統就會自動幫您閱讀題目並進行分類。")

# --- 智慧分類字典設定區 ---
st.subheader("🧠 第一步：設定自動分類關鍵字字典 (選填)")
st.info("系統會掃描題目，若包含以下關鍵字，將自動套用對應的分類主題。您可以依照不同科目自由修改這份清單！")

default_mapping = {
    "過敏反應": ["IgE", "過敏", "氣喘", "第一型", "第二型", "第三型", "第四型", "hypersensitivity"],
    "腫瘤免疫": ["腫瘤", "癌症", "tumor", "cancer", "TSA", "TAA", "carcinoma"],
    "自體免疫": ["自體免疫", "紅斑性狼瘡", "風濕", "SLE", "RA", "autoimmune"],
    "移植免疫": ["移植", "排斥", "GVHD", "MHC", "HLA", "graft"],
    "先天免疫": ["先天免疫", "巨噬細胞", "補體", "complement", "NK cell", "發炎", "吞噬"],
    "細胞免疫": ["T細胞", "CD4", "CD8", "T cell", "細胞毒殺"],
    "體液免疫": ["B細胞", "B cell", "抗體", "IgG", "IgM", "IgA", "漿細胞"]
}

mapping_str = st.text_area(
    "請以 JSON 格式編輯您的關鍵字字典：", 
    value=json.dumps(default_mapping, ensure_ascii=False, indent=4),
    height=250
)

try:
    topic_mapping = json.loads(mapping_str)
except Exception as e:
    st.error(f"⚠️ 字典格式錯誤，請確認標點符號是否為半形！將暫時使用預設字典。({e})")
    topic_mapping = default_mapping

st.divider()

# --- 檔案上傳區 ---
st.subheader("📂 第二步：上傳 Word 題庫檔 (.docx)")
uploaded_file = st.file_uploader("請選擇準備好的 Word 檔案", type=["docx"])

if uploaded_file is not None:
    if st.button("🚀 開始全自動轉換與分類", type="primary", use_container_width=True):
        with st.spinner("正在執行全域掃描與智慧分類..."):
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
                current_topic = "未分類" # 這裡存放「章節繼承法」的標題
                current_q = None
                skipped_lines = []

                year_pattern = re.compile(r'(\d{2,4})\s*年')
                q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
                exp_pattern = re.compile(r'^[\"\'\,\.\-\s]*解\s*析\s*[:](.*)', re.IGNORECASE)
                # 🌟 嚴謹的章節標題偵測 (支援: 【腫瘤免疫】 或 單元: 腫瘤免疫)
                topic_pattern = re.compile(r'^(?:【([^】]+)】|(?:\w{2}[:：]\s*)(.+))$')

                for text in all_lines:
                    # 1. 偵測章節標題 (章節繼承法)
                    t_match = topic_pattern.match(text)
                    if t_match and not q_start_pattern.search(text):
                        if current_q:
                            _extract_options(current_q)
                            _extract_tags_from_all(current_q)
                            auto_categorize(current_q, topic_mapping) # 提交前進行智慧分類
                            questions.append(current_q)
                            current_q = None
                        current_topic = t_match.group(1) or t_match.group(2)
                        continue

                    # 2. 偵測年份
                    year_match = year_pattern.search(text)
                    if year_match and not q_start_pattern.search(text): 
                        if current_q:
                            _extract_options(current_q)
                            _extract_tags_from_all(current_q)
                            auto_categorize(current_q, topic_mapping)
                            questions.append(current_q)
                            current_q = None
                        current_year = text.replace('"', '').replace(',', '').strip()
                        continue
                        
                    # 3. 抓題目
                    q_match = q_start_pattern.match(text)
                    if q_match:
                        if current_q:
                            _extract_options(current_q)
                            _extract_tags_from_all(current_q)
                            auto_categorize(current_q, topic_mapping)
                            questions.append(current_q)
                        
                        ans = q_match.group('ans').strip().upper().replace('，', ',')
                        try:
                            q_num = int(q_match.group('num'))
                        except:
                            q_num = 0
                            
                        current_q = {
                            "question_number": q_num,
                            "answer": ans,
                            "explanation": "",
                            "tags": {
                                "年份": current_year,
                                "主題": current_topic # 優先套用上方偵測到的章節標題
                            },
                            "_raw_text": q_match.group('text').strip()
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
                            if text[:split_idx].strip():
                                current_q["_raw_text"] += "\n" + text[:split_idx].strip()
                            current_q["explanation"] += text[hidden_exp_match.end():].strip()
                            continue

                        if current_q["explanation"]:
                            current_q["explanation"] += "\n" + text
                        else:
                            current_q["_raw_text"] += "\n" + text
                    else:
                        if len(text) > 5:
                            skipped_lines.append(text)

                # 收尾最後一題
                if current_q:
                    _extract_options(current_q)
                    _extract_tags_from_all(current_q)
                    auto_categorize(current_q, topic_mapping)
                    questions.append(current_q)

                if questions:
                    st.success(f"🎉 轉換大功告成！系統共成功辨識並分類了 **{len(questions)}** 題！")
                    
                    # 統計分類結果，讓老師一眼看出分類狀況
                    st.markdown("#### 📊 本次分類統計結果：")
                    topic_counts = {}
                    for q in questions:
                        t = q["tags"].get("主題", "未分類")
                        topic_counts[t] = topic_counts.get(t, 0) + 1
                    
                    st.write(" | ".join([f"**{k}**: {v}題" for k, v in topic_counts.items()]))
                    
                    json_str = json.dumps(questions, ensure_ascii=False, separators=(',', ':'))
                    
                    st.download_button(
                        label="💾 點我下載分類完畢的 JSON 題庫檔",
                        data=json_str,
                        file_name=uploaded_file.name.replace(".docx", "_V7_智慧分類版.json"),
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
