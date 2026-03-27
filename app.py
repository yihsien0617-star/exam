import streamlit as st
import docx
import json
import re
import pandas as pd
import io
import uuid
import base64

# ==========================================
# 內部輔助函數 (文字、圖片與標籤清洗)
# ==========================================
def get_para_text_with_images(para, image_db):
    """🌟 完美復刻 V1 的 100% 不漏字抓取，並安全提取圖片"""
    full_text = para.text  # 直接抓取完整段落，無視底層複雜格式，絕不漏字！
    img_placeholders = ""
    for run in para.runs:
        try:
            blips = run._element.xpath('.//*[local-name()="blip"]')
            for blip in blips:
                rId = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if rId:
                    part = para.part.related_parts[rId]
                    b64 = base64.b64encode(part.blob).decode('utf-8')
                    img_id = f"IMG_{uuid.uuid4().hex[:8]}"
                    image_db[img_id] = f"[IMAGE_BASE64:data:{part.content_type};base64,{b64}]"
                    img_placeholders += f"\n[{img_id}]\n"
        except:
            pass
    # 將圖片安全地附掛在該段文字的最後面
    return full_text + img_placeholders

def normalize_line(text):
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    text = re.sub(r'[ \t\u3000]+', ' ', text)
    return text.strip()

def extract_and_remove_tags(text, q_dict):
    """只抓特定標籤，絕不誤吃解析本體"""
    if not text: return text
    keywords = ["難度", "再現性", "主題", "分類", "單元", "章節"]
    for kw in keywords:
        kw_regex = kw[0] + r'\s*' + kw[1:] if len(kw) >= 2 else kw
        pattern = r'[\"\'”’]?\s*(' + kw_regex + r')\s*[:]\s*([^\"\'”’\n]+)[\"\'”’]?'
        for m in re.finditer(pattern, text):
            val = m.group(2).strip()
            val = re.sub(r'\(.*?\)|（.*?）', '', val).strip()
            q_dict["tags"][kw] = val.rstrip(',，').strip()
            text = text.replace(m.group(0), '')
            
    text = re.sub(r'^[,\"\'\s，]+|[,\"\'\s，]+$', '', text)
    return re.sub(r',\s*,', ',', text).strip()

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

def replace_images_in_dict(d, img_db):
    if isinstance(d, dict):
        for k, v in d.items():
            if isinstance(v, str):
                for img_id, img_str in img_db.items():
                    if f"[{img_id}]" in v:
                        d[k] = v.replace(f"[{img_id}]", img_str)
            else:
                replace_images_in_dict(v, img_db)
    elif isinstance(d, list):
        for item in d:
            replace_images_in_dict(item, img_db)

# ==========================================
# 網頁介面開始
# ==========================================
st.set_page_config(page_title="國考題庫轉檔與協作系統", page_icon="⚙️", layout="wide")

st.title("⚙️ 國考題庫轉檔系統 (V11 完美文字回歸版)")
st.write("已完美結合初代 100% 抓字準確率與全自動圖片萃取功能！")

tab1, tab2, tab3 = st.tabs(["🚀 一鍵產出 JSON (推薦)", "📝 階段一：轉為 Excel 供校對", "💾 階段二：Excel 打包 JSON"])

default_mapping = {
    "過敏反應": ["IgE", "過敏", "氣喘", "hypersensitivity"],
    "腫瘤免疫": ["腫瘤", "癌症", "tumor", "cancer", "TSA", "TAA"],
    "自體免疫": ["自體免疫", "紅斑性狼瘡", "風濕", "SLE", "RA"],
    "移植免疫": ["移植", "排斥", "GVHD", "MHC", "HLA"],
    "先天免疫": ["先天免疫", "巨噬細胞", "補體", "complement", "NK cell", "發炎"],
    "細胞免疫": ["T細胞", "CD4", "CD8", "T cell", "細胞毒殺"],
    "體液免疫": ["B細胞", "B cell", "抗體", "IgG", "IgM", "IgA", "漿細胞"]
}

def parse_word_document(uploaded_file, topic_mapping):
    doc = docx.Document(uploaded_file)
    image_db = {}
    all_lines = []
    
    for para in doc.paragraphs:
        raw_text_with_imgs = get_para_text_with_images(para, image_db)
        for line in re.split(r'[\n\v]', raw_text_with_imgs):
            clean_line = normalize_line(line)
            if clean_line: all_lines.append(clean_line)

    questions = []
    current_year = "未知年份"
    current_topic = "未分類"
    current_q = None

    year_pattern = re.compile(r'(\d{2,4})\s*年')
    q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
    topic_pattern = re.compile(r'^(?:【([^】]+)】|(?:\w{2}[:：]\s*)(.+))$')
    
    # 🌟 無敵解析切割法 (相容: 【解析】、解析:、解 析 、[解答] 等各種格式，沒有冒號也抓得到！)
    exp_pattern = re.compile(r'^[\"\'\,\.\-\s【\[<]*解\s*答?\s*析[\s:：\]】>]*(.*)', re.IGNORECASE)

    for text in all_lines:
        t_match = topic_pattern.match(text)
        if t_match and not q_start_pattern.search(text) and not exp_pattern.search(text):
            if current_q:
                _extract_options(current_q); _extract_tags_from_all(current_q); auto_categorize(current_q, topic_mapping); questions.append(current_q)
                current_q = None
            current_topic = t_match.group(1) or t_match.group(2)
            continue

        year_match = year_pattern.search(text)
        if year_match and not q_start_pattern.search(text) and not exp_pattern.search(text): 
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
            
        exp_match = exp_pattern.match(text)
        if exp_match and current_q:
            current_q["explanation"] = exp_match.group(1).strip()
            continue
            
        if current_q:
            # 同行解析暴力擷取
            hidden_exp_match = re.search(r'([\"\'\,\.\-\s【\[<]*)(解\s*答?\s*析[\s:：\]】>]+)(.*)', text, re.IGNORECASE)
            if hidden_exp_match:
                q_part = text[:hidden_exp_match.start(1)].strip()
                if q_part:
                    current_q["_raw_text"] += "\n" + q_part
                current_q["explanation"] += text[hidden_exp_match.end(2):].strip()
                continue

            if current_q["explanation"]: current_q["explanation"] += "\n" + text
            else: current_q["_raw_text"] += "\n" + text

    if current_q:
        _extract_options(current_q); _extract_tags_from_all(current_q); auto_categorize(current_q, topic_mapping); questions.append(current_q)
        
    return questions, image_db

# ==========================================
# Tab 1: 直接產出 JSON 
# ==========================================
with tab1:
    st.info("直接將 Word 轉換為系統可讀的 JSON，完美保留初代文字精準度，解析 100% 完整呈現！")
    mapping_str_1 = st.text_area("關鍵字分類字典：", value=json.dumps(default_mapping, ensure_ascii=False, indent=4), height=150, key="map1")
    try: topic_mapping_1 = json.loads(mapping_str_1)
    except: topic_mapping_1 = default_mapping
    
    uploaded_word_1 = st.file_uploader("上傳 Word 題庫 (.docx)", type=["docx"], key="w1")
    if uploaded_word_1 and st.button("🚀 產出最終 JSON 題庫", type="primary", use_container_width=True):
        with st.spinner("正在萃取圖片與解析..."):
            qs, img_db = parse_word_document(uploaded_word_1, topic_mapping_1)
            if qs:
                replace_images_in_dict(qs, img_db)
                st.success(f"成功解析 {len(qs)} 題！共抽取了 {len(img_db)} 張圖片。")
                json_str = json.dumps(qs, ensure_ascii=False, separators=(',', ':'))
                st.download_button("💾 下載 JSON 上線檔", data=json_str, file_name=uploaded_word_1.name.replace(".docx", "_完美解析版.json"), mime="application/json", type="primary", use_container_width=True)

# ==========================================
# Tab 2: Word 轉 Excel
# ==========================================
with tab2:
    st.info("讓老師用 Excel 校對分類。⚠️ 系統會將圖片暫時替換為 [IMG_xxx] 標記以防 Excel 崩潰，轉回 JSON 時會自動復原！")
    mapping_str_2 = st.text_area("關鍵字分類字典：", value=json.dumps(default_mapping, ensure_ascii=False, indent=4), height=150, key="map2")
    try: topic_mapping_2 = json.loads(mapping_str_2)
    except: topic_mapping_2 = default_mapping
    
    uploaded_word_2 = st.file_uploader("上傳 Word 題庫 (.docx)", type=["docx"], key="w2")
    if uploaded_word_2 and st.button("🚀 產出待校對 Excel 與圖片暫存檔", type="primary", use_container_width=True):
        with st.spinner("正在產生 Excel..."):
            qs, img_db = parse_word_document(uploaded_word_2, topic_mapping_2)
            if qs:
                excel_rows = []
                for q in qs:
                    opts = q.get("options", {})
                    excel_rows.append({
                        "年份": q["tags"].get("年份", ""),
                        "題號": q.get("question_number", ""),
                        "主題 (下拉選單)": q["tags"].get("主題", ""),
                        "題目": q.get("question_text", ""),
                        "選項A": opts.get("A", ""), "選項B": opts.get("B", ""),
                        "選項C": opts.get("C", ""), "選項D": opts.get("D", ""),
                        "正確答案": q.get("answer", ""), "解析": q.get("explanation", "")
                    })
                df = pd.DataFrame(excel_rows)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='待校對題庫')
                    workbook = writer.book
                    worksheet = writer.sheets['待校對題庫']
                    worksheet.set_column('A:B', 8); worksheet.set_column('C:C', 20); worksheet.set_column('D:D', 45)
                    worksheet.set_column('E:H', 22); worksheet.set_column('I:I', 10); worksheet.set_column('J:J', 50)
                    
                    topic_sheet = workbook.add_worksheet('主題清單(可擴充)')
                    topic_sheet.write('A1', '🔽 系統預設主題 (往下新增將自動同步)')
                    topic_sheet.set_column('A:A', 50)
                    for i, t in enumerate(["未分類"] + list(topic_mapping_2.keys())):
                        topic_sheet.write(i + 1, 0, t)
                    worksheet.data_validation('C2:C10000', {'validate': 'list', 'source': "='主題清單(可擴充)'!$A$2:$A$200", 'error_type': 'warning'})

                st.success(f"成功解析 {len(qs)} 題！")
                st.download_button("📊 1. 下載待校對 Excel 檔", data=output.getvalue(), file_name=uploaded_word_2.name.replace(".docx", "_校對用.xlsx"), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                img_json = json.dumps(img_db, ensure_ascii=False)
                st.download_button("🖼️ 2. 下載圖片暫存檔 (image_db.json)", data=img_json, file_name="image_db.json", mime="application/json", use_container_width=True)

# ==========================================
# Tab 3: Excel 轉 JSON
# ==========================================
with tab3:
    st.info("上傳校對好的 Excel，並附上圖片暫存檔，系統會將圖片與解析完美還原至題庫中！")
    uploaded_excel = st.file_uploader("1. 上傳校對完的 Excel (.xlsx)", type=["xlsx"])
    uploaded_img_db = st.file_uploader("2. 上傳圖片暫存檔 (image_db.json) (若本題無圖片可略過)", type=["json"])

    if uploaded_excel and st.button("💾 封裝為最終 JSON 題庫", type="primary", use_container_width=True):
        with st.spinner("正在封裝題庫並還原圖片..."):
            try:
                img_db = json.load(uploaded_img_db) if uploaded_img_db else {}
                df = pd.read_excel(uploaded_excel, sheet_name='待校對題庫').fillna("")
                
                final_questions = []
                for idx, row in df.iterrows():
                    if str(row.get("題目", "")).strip() == "": continue
                    opts = {k: str(row.get(f"選項{k}", "")).strip() for k in ['A', 'B', 'C', 'D'] if str(row.get(f"選項{k}", "")).strip()}
                    q_num = str(row.get("題號", "0")).strip()
                    q_num = int(float(q_num)) if q_num.replace('.', '', 1).isdigit() else 0
                    
                    q = {
                        "question_number": q_num, "answer": str(row.get("正確答案", "")).strip(),
                        "explanation": str(row.get("解析", "")).strip(),
                        "tags": {"年份": str(row.get("年份", "")).strip(), "主題": str(row.get("主題 (下拉選單)", "")).strip()},
                        "question_text": str(row.get("題目", "")).strip(), "options": opts
                    }
                    final_questions.append(q)
                    
                if final_questions:
                    replace_images_in_dict(final_questions, img_db)
                    st.success(f"🎉 封裝成功！共匯入 {len(final_questions)} 題！")
                    json_str = json.dumps(final_questions, ensure_ascii=False, separators=(',', ':'))
                    st.download_button("📥 下載最終上線版 JSON", data=json_str, file_name=uploaded_excel.name.replace(".xlsx", "_最終上線版.json"), mime="application/json", type="primary", use_container_width=True)
            except Exception as e:
                st.error(f"錯誤：{e}")
