import streamlit as st
import docx
import json
import re
import pandas as pd
import io

# --- 內部輔助函數 (文字處理與清洗) ---
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
    if q_dict["tags"].get("主題", "未分類") != "未分類": return
    search_text = (q_dict.get("question_text", "") + " " + q_dict.get("explanation", "")).lower()
    for topic, keywords in mapping.items():
        for kw in keywords:
            if kw.lower() in search_text:
                q_dict["tags"]["主題"] = topic
                return 

# --- 網頁介面開始 ---
st.set_page_config(page_title="國考題庫轉檔與協作系統", page_icon="⚙️", layout="wide")

st.title("⚙️ 國考題庫轉檔與協作系統 (V8 雙引擎版)")
st.write("這套系統為降低教師負擔而生。請先將 Word 轉為 Excel 供老師校對，確認無誤後再將 Excel 轉為最終題庫。")

tab1, tab2 = st.tabs(["📝 階段一：Word 轉 Excel (AI 初步掃描)", "💾 階段二：Excel 轉 JSON (最終題庫產出)"])

# ==========================================
# 階段一：Word 轉 Excel
# ==========================================
with tab1:
    st.subheader("🧠 第一步：設定初步分類字典 (選填)")
    st.info("系統會盡力幫您初步分類。沒分好的部分，等一下匯出 Excel 後老師再手動改就好！")
    
    default_mapping = {
        "過敏反應": ["IgE", "過敏", "氣喘", "hypersensitivity"],
        "腫瘤免疫": ["腫瘤", "癌症", "tumor", "cancer", "TSA", "TAA"],
        "自體免疫": ["自體免疫", "紅斑性狼瘡", "風濕", "SLE", "RA"],
        "移植免疫": ["移植", "排斥", "GVHD", "MHC", "HLA"],
        "先天免疫": ["先天免疫", "巨噬細胞", "補體", "complement", "NK cell", "發炎"],
        "細胞免疫": ["T細胞", "CD4", "CD8", "T cell", "細胞毒殺"],
        "體液免疫": ["B細胞", "B cell", "抗體", "IgG", "IgM", "IgA", "漿細胞"]
    }
    
    mapping_str = st.text_area("請定義初步抓取的關鍵字：", value=json.dumps(default_mapping, ensure_ascii=False, indent=4), height=150)
    try:
        topic_mapping = json.loads(mapping_str)
    except:
        topic_mapping = default_mapping

    st.subheader("📂 第二步：上傳原始 Word 題庫檔")
    uploaded_word = st.file_uploader("選擇 Word 檔案 (.docx)", type=["docx"], key="word_uploader")

    if uploaded_word is not None:
        if st.button("🚀 產出 Excel 給老師校對", type="primary", use_container_width=True):
            with st.spinner("正在掃描 Word 並產生 Excel 表格..."):
                try:
                    doc = docx.Document(uploaded_word)
                    all_lines = []
                    for para in doc.paragraphs:
                        for line in para.text.split('\n'):
                            clean_line = normalize_line(line)
                            if clean_line: all_lines.append(clean_line)

                    questions = []
                    current_year = "未知年份"
                    current_topic = "未分類"
                    current_q = None

                    year_pattern = re.compile(r'(\d{2,4})\s*年')
                    q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
                    topic_pattern = re.compile(r'^(?:【([^】]+)】|(?:\w{2}[:：]\s*)(.+))$')

                    for text in all_lines:
                        t_match = topic_pattern.match(text)
                        if t_match and not q_start_pattern.search(text):
                            if current_q:
                                _extract_options(current_q); _extract_tags_from_all(current_q); auto_categorize(current_q, topic_mapping); questions.append(current_q)
                                current_q = None
                            current_topic = t_match.group(1) or t_match.group(2)
                            continue

                        year_match = year_pattern.search(text)
                        if year_match and not q_start_pattern.search(text): 
                            if current_q:
                                _extract_options(current_q); _extract_tags_from_all(current_q); auto_categorize(current_q, topic_mapping); questions.append(current_q)
                                current_q = None
                            current_year = text.replace('"', '').replace(',', '').strip()
                            continue
                            
                        q_match = q_start_pattern.match(text)
                        if q_match:
                            if current_q:
                                _extract_options(current_q); _extract_tags_from_all(current_q); auto_categorize(current_q, topic_mapping); questions.append(current_q)
                            ans = q_match.group('ans').strip().upper().replace('，', ',')
                            q_num = int(q_match.group('num')) if q_match.group('num').isdigit() else 0
                            current_q = {"question_number": q_num, "answer": ans, "explanation": "", "tags": {"年份": current_year, "主題": current_topic}, "_raw_text": q_match.group('text').strip()}
                            continue
                            
                        if current_q:
                            hidden_exp_match = re.search(r'[\"\'\,\.\-\s]*解\s*析\s*[:]', text, re.IGNORECASE)
                            if hidden_exp_match:
                                current_q["_raw_text"] += "\n" + text[:hidden_exp_match.start()].strip()
                                current_q["explanation"] += text[hidden_exp_match.end():].strip()
                                continue

                            if current_q["explanation"]: current_q["explanation"] += "\n" + text
                            else: current_q["_raw_text"] += "\n" + text

                    if current_q:
                        _extract_options(current_q); _extract_tags_from_all(current_q); auto_categorize(current_q, topic_mapping); questions.append(current_q)

                    # --- 轉換為 Excel 格式 ---
                    if questions:
                        excel_rows = []
                        for q in questions:
                            opts = q.get("options", {})
                            excel_rows.append({
                                "年份": q["tags"].get("年份", ""),
                                "題號": q.get("question_number", ""),
                                "主題 (請在此修正)": q["tags"].get("主題", ""),
                                "題目": q.get("question_text", ""),
                                "選項A": opts.get("A", ""),
                                "選項B": opts.get("B", ""),
                                "選項C": opts.get("C", ""),
                                "選項D": opts.get("D", ""),
                                "正確答案": q.get("answer", ""),
                                "解析": q.get("explanation", "")
                            })
                            
                        df = pd.DataFrame(excel_rows)
                        
                        # 匯出至 BytesIO
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df.to_excel(writer, index=False, sheet_name='待校對題庫')
                            
                            # 美化 Excel 寬度
                            worksheet = writer.sheets['待校對題庫']
                            worksheet.set_column('A:B', 8)
                            worksheet.set_column('C:C', 15)
                            worksheet.set_column('D:D', 40)
                            worksheet.set_column('E:H', 20)
                            worksheet.set_column('I:I', 10)
                            worksheet.set_column('J:J', 40)

                        st.success(f"🎉 成功初步解析了 **{len(questions)}** 題！請下載 Excel 檔發派給各科老師校對。")
                        st.download_button(
                            label="📊 下載待校對 Excel 檔",
                            data=output.getvalue(),
                            file_name=uploaded_word.name.replace(".docx", "_待校對.xlsx"),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                    else:
                        st.error("😭 找不到任何題目。")
                except Exception as e:
                    st.error(f"系統發生錯誤：{e}")

# ==========================================
# 階段二：Excel 轉 JSON
# ==========================================
with tab2:
    st.subheader("📥 上傳老師校對完畢的 Excel 檔")
    st.info("當各科老師在 Excel 裡把「主題」和「錯字」都改好後，將檔案上傳至此，系統會直接將其封裝成前台可用的 JSON 系統題庫！")
    
    uploaded_excel = st.file_uploader("選擇已校對的 Excel 檔案 (.xlsx)", type=["xlsx"], key="excel_uploader")

    if uploaded_excel is not None:
        if st.button("💾 封裝為最終 JSON 題庫", type="primary", use_container_width=True):
            with st.spinner("正在讀取 Excel 並封裝為系統題庫..."):
                try:
                    df = pd.read_excel(uploaded_excel)
                    df = df.fillna("") # 將 NaN 轉換為空字串
                    
                    final_questions = []
                    for idx, row in df.iterrows():
                        # 避開完全空白的行
                        if str(row.get("題目", "")).strip() == "": continue
                            
                        # 重組選項
                        opts = {}
                        for k in ['A', 'B', 'C', 'D']:
                            val = str(row.get(f"選項{k}", "")).strip()
                            if val: opts[k] = val
                            
                        q_num = str(row.get("題號", "0")).strip()
                        q_num = int(float(q_num)) if q_num.replace('.', '', 1).isdigit() else 0
                        
                        q = {
                            "question_number": q_num,
                            "answer": str(row.get("正確答案", "")).strip(),
                            "explanation": str(row.get("解析", "")).strip(),
                            "tags": {
                                "年份": str(row.get("年份", "")).strip(),
                                "主題": str(row.get("主題 (請在此修正)", "")).strip()
                            },
                            "question_text": str(row.get("題目", "")).strip(),
                            "options": opts
                        }
                        final_questions.append(q)
                        
                    if final_questions:
                        st.success(f"🎉 封裝成功！共匯入 **{len(final_questions)}** 題完美確認版考題！")
                        
                        json_str = json.dumps(final_questions, ensure_ascii=False, separators=(',', ':'))
                        
                        st.download_button(
                            label="📥 下載最終上線版 JSON 題庫",
                            data=json_str,
                            file_name=uploaded_excel.name.replace(".xlsx", "_最終上線版.json"),
                            mime="application/json",
                            type="primary",
                            use_container_width=True
                        )
                        st.markdown("⚠️ **下一步：** 帶著下載好的 JSON 檔案，前往您的【國考平台前台 -> 管理員登入 -> 科目與題庫管理】進行上傳發布！")
                    else:
                        st.error("😭 Excel 檔案內似乎沒有有效的題目資料。")
                        
                except Exception as e:
                    st.error(f"讀取 Excel 失敗，請確認檔案格式是否正確：{e}")
