import streamlit as st
import docx
import json
import re
import pandas as pd
import io
import uuid
import zipfile
from PIL import Image

# ==========================================
# 內部輔助函數 (文字、圖片壓縮與標籤清洗)
# ==========================================
def compress_image_blob(blob):
    """智慧圖片壓縮引擎：自動縮圖並轉為高壓縮 JPEG"""
    try:
        img = Image.open(io.BytesIO(blob))
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.thumbnail((800, 800), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=75) 
        return out.getvalue()
    except Exception as e:
        return blob

def get_para_text_with_images(para, image_db):
    """提取文字，並留下 [IMG: 檔名.jpg] 的輕量化佔位符"""
    full_text = para.text  
    img_placeholders = ""
    for run in para.runs:
        try:
            blips = run._element.xpath('.//*[local-name()="blip"]')
            imagedatas = run._element.xpath('.//*[local-name()="imagedata"]')
            
            for element in blips + imagedatas:
                rId = element.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed') or \
                      element.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if rId:
                    part = para.part.related_parts[rId]
                    compressed_blob = compress_image_blob(part.blob)
                    
                    # 🌟 改為實體檔案命名法
                    img_filename = f"IMG_{uuid.uuid4().hex[:8]}.jpg"
                    image_db[img_filename] = compressed_blob
                    img_placeholders += f"\n[IMG: {img_filename}]\n"
        except:
            pass
    return full_text + img_placeholders

def normalize_line(text):
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    text = re.sub(r'[ \t\u3000]+', ' ', text)
    return text.strip()

def extract_and_remove_tags(text, q_dict):
    if not text: return text
    keywords = ["難度", "再現性", "主題", "分類", "單元", "章節"]
    for kw in keywords:
        pattern = r'[\"\'”’]?\s*(' + kw + r')\s*[:：]\s*([^\"\'”’\n]+)[\"\'”’]?'
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

def finalize_question(q_dict, topic_mapping):
    if not q_dict.get("explanation"):
        raw = q_dict.get("_raw_text", "")
        fb1 = re.search(r'\n[\"\'\,\.\-\s【\[<]*(解\s*答?\s*析)(?!度)[\s:：\]】>\"\',，]*(.*)', raw, re.IGNORECASE | re.DOTALL)
        if fb1:
            q_dict["explanation"] = fb1.group(2).strip()
            q_dict["_raw_text"] = raw[:fb1.start()].strip()
        else:
            fb2 = re.search(r'(解\s*答?\s*析)(?!度)[\s:：]+(.*)', raw, re.IGNORECASE | re.DOTALL)
            if fb2:
                q_dict["explanation"] = fb2.group(2).strip()
                q_dict["_raw_text"] = raw[:fb2.start()].strip(' ,"\'.\n\t\r【<[')

    _extract_options(q_dict)
    _extract_tags_from_all(q_dict)
    auto_categorize(q_dict, topic_mapping)

def create_zip_package(json_data, image_db, filename_prefix):
    """將 JSON 與所有實體圖片打包成 ZIP"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        json_str = json.dumps(json_data, ensure_ascii=False, separators=(',', ':'))
        zip_file.writestr(f"{filename_prefix}.json", json_str.encode('utf-8'))
        
        for img_name, img_bytes in image_db.items():
            zip_file.writestr(f"images/{img_name}", img_bytes)
    return zip_buffer.getvalue()

# ==========================================
# 網頁介面開始
# ==========================================
st.set_page_config(page_title="國考題庫轉檔系統", page_icon="⚙️", layout="wide")
st.title("⚙️ 國考題庫轉檔系統 (V18 圖文分離版)")
st.write("徹底解決系統 503 當機問題！產出的 ZIP 檔內含極輕量化 JSON 與獨立圖片資料夾。")

default_mapping = {
    "過敏反應": ["IgE", "過敏", "氣喘"],
    "腫瘤免疫": ["腫瘤", "癌症", "tumor"],
    "自體免疫": ["自體免疫", "紅斑性狼瘡", "風濕", "SLE"],
    "移植免疫": ["移植", "排斥", "GVHD", "MHC"],
    "先天性免疫": ["先天免疫", "巨噬細胞", "補體"],
    "細胞性免疫": ["T細胞", "CD4", "CD8", "T cell"],
    "體液性免疫": ["B細胞", "B cell", "抗體", "IgG"]
}

def parse_word_document(uploaded_file, topic_mapping):
    doc = docx.Document(uploaded_file)
    image_db = {}
    all_lines = []
    
    for block in doc.element.body.iterchildren():
        if block.tag.endswith('p'):
            para = docx.text.paragraph.Paragraph(block, doc)
            raw_text = get_para_text_with_images(para, image_db)
            for line in re.split(r'[\n\v]', raw_text):
                clean_line = normalize_line(line)
                if clean_line: all_lines.append(clean_line)
                
        elif block.tag.endswith('tbl'):
            table = docx.table.Table(block, doc)
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        raw_text = get_para_text_with_images(para, image_db)
                        for line in re.split(r'[\n\v]', raw_text):
                            clean_line = normalize_line(line)
                            if clean_line: all_lines.append(clean_line)

    questions = []
    current_year = "未知年份"
    current_topic = "未分類"
    current_q = None
    read_mode = "raw" 

    year_pattern = re.compile(r'(\d{2,4})\s*年')
    q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
    topic_pattern = re.compile(r'^(?:【([^】]+)】|(?:\w{2}[:：]\s*)(.+))$')
    exp_pattern = re.compile(r'^[\"\'\,\.\-\s【\[<]*解\s*答?\s*析(?!度)[\s:：\]】>\"\',，]*(.*)', re.IGNORECASE)

    for text in all_lines:
        t_match = topic_pattern.match(text)
        if t_match and not q_start_pattern.search(text) and not exp_pattern.match(text):
            if current_q:
                finalize_question(current_q, topic_mapping)
                questions.append(current_q)
                current_q = None
            current_topic = t_match.group(1) or t_match.group(2)
            continue

        year_match = year_pattern.search(text)
        if year_match and not q_start_pattern.search(text) and not exp_pattern.match(text): 
            if current_q:
                finalize_question(current_q, topic_mapping)
                questions.append(current_q)
                current_q = None
            current_year = text.replace('"', '').replace(',', '').strip()
            continue
            
        q_match = q_start_pattern.match(text)
        if q_match:
            if current_q:
                finalize_question(current_q, topic_mapping)
                questions.append(current_q)
            
            read_mode = "raw"
            ans = q_match.group('ans').strip().upper().replace('，', ',')
            q_num = int(q_match.group('num')) if q_match.group('num').isdigit() else 0
            
            current_q = {
                "question_number": q_num, "answer": ans, "explanation": "", 
                "tags": {"年份": current_year, "主題": current_topic}, 
                "_raw_text": q_match.group('text').strip()
            }
            continue
            
        exp_match = exp_pattern.match(text)
        if exp_match and current_q:
            read_mode = "explanation"
            if not current_q["explanation"]:
                current_q["explanation"] = exp_match.group(1).strip()
            else:
                current_q["explanation"] += "\n" + text
            continue
            
        if current_q:
            if read_mode == "explanation":
                current_q["explanation"] += "\n" + text
            else:
                current_q["_raw_text"] += "\n" + text

    if current_q:
        finalize_question(current_q, topic_mapping)
        questions.append(current_q)
        
    return questions, image_db

mapping_str_1 = st.text_area("關鍵字分類字典：", value=json.dumps(default_mapping, ensure_ascii=False, indent=4), height=150)
try: topic_mapping_1 = json.loads(mapping_str_1)
except: topic_mapping_1 = default_mapping

uploaded_word_1 = st.file_uploader("上傳 Word 題庫 (.docx)", type=["docx"])
if uploaded_word_1 and st.button("🚀 產出 ZIP 題庫包", type="primary", use_container_width=True):
    with st.spinner("正在執行圖文分離與打包作業..."):
        qs, img_db = parse_word_document(uploaded_word_1, topic_mapping_1)
        if qs:
            prefix = uploaded_word_1.name.replace(".docx", "")
            zip_bytes = create_zip_package(qs, img_db, prefix)
            
            st.success(f"成功解析 {len(qs)} 題！共抽取 {len(img_db)} 張圖片，已完美分離打包！")
            st.download_button(
                label="📦 下載 ZIP 題庫包 (含 JSON 與獨立 images 資料夾)", 
                data=zip_bytes, 
                file_name=f"{prefix}_題庫包.zip", 
                mime="application/zip", 
                type="primary", 
                use_container_width=True
            )
