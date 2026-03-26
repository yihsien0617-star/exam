import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import re
import json
import io
import base64
from PIL import Image

# --- 1. 抽取純文字與圖片引擎 (圖片壓縮升級版) ---
def extract_raw_text(file_stream):
    doc = docx.Document(file_stream)
    lines = []
    
    def process_paragraph(para):
        para_text = ""
        for run in para.runs:
            para_text += run.text
            for blip in run._element.xpath('.//*[local-name()="blip"]'):
                embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed:
                    try:
                        image_part = run.part.rels[embed].target_part
                        image_data = image_part.blob
                        img = Image.open(io.BytesIO(image_data))
                        if img.mode in ("RGBA", "P"):
                            img = img.convert("RGB")
                        img.thumbnail((800, 800), Image.Resampling.LANCZOS)
                        buffer = io.BytesIO()
                        img.save(buffer, format="JPEG", quality=75)
                        compressed_blob = buffer.getvalue()
                        b64_str = base64.b64encode(compressed_blob).decode('utf-8')
                        para_text += f"\n[IMAGE_BASE64:data:image/jpeg;base64,{b64_str}]\n"
                    except Exception as e:
                        pass
        
        if para_text.strip():
            for line in para_text.split('\n'):
                if line.strip():
                    lines.append(line.strip())

    for element in doc.element.body:
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            process_paragraph(para)
        elif element.tag.endswith('tbl'):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        process_paragraph(para)
    return lines

# --- 2. 標籤與解析淨化工具 ---
def extract_tags_and_clean(text, current_tags):
    diff_match = re.search(r'難\s*度[:：]\s*([^\(,"”]+)', text)
    if diff_match: current_tags["難度"] = diff_match.group(1).strip()
    rep_match = re.search(r'再\s*現\s*性[:：]\s*([^\(,"”]+)', text)
    if rep_match: current_tags["再現性"] = rep_match.group(1).strip()
        
    clean_exp = re.sub(r'[,"]*\s*難\s*度[:：][^"]+"?', '', text)
    clean_exp = re.sub(r'[,"]*\s*再\s*現\s*性[:：][^"]+"?', '', clean_exp)
    clean_exp = re.sub(r'\([^\)]*極低[^\)]*\)', '', clean_exp)
    clean_exp = re.sub(r'\([^\)]*非常簡單[^\)]*\)', '', clean_exp)
    return clean_exp.strip('", '), current_tags

# --- 3. 自動主題分類引擎 (排除圖片干擾版) ---
def auto_classify_topic(q_dict):
    text = q_dict["question_text"] + " " + q_dict.get("explanation", "")
    for opt in q_dict["options"].values(): text += " " + opt
    
    # 修正重點 1：先移除 Base64 圖片亂碼，避免干擾關鍵字比對
    text = re.sub(r'\[IMAGE_BASE64:[^\]]+\]', '', text)
    text = text.upper()
    
    topics = {
        "先天免疫與發炎反應": ["發炎", "白血球", "吞噬", "巨噬細胞", "MACROPHAGE", "NEUTROPHIL", "NK細胞", "自然殺手", "TLR", "先天免疫", "發燒", "C-REACTIVE", "CRP"],
        "補體系統": ["補體", "COMPLEMENT", "C3", "C4", "C5", "MAC", "古典途徑", "替代途徑", "凝集素途徑", "C1Q"],
        "抗體與免疫球蛋白": ["抗體", "免疫球蛋白", "IGG", "IGA", "IGM", "IGE", "IGD", "ISOTYPE", "輕鏈", "重鏈", "FAB", "FC", "ALLOTYPE", "IDIOTYPE"],
        "T細胞與細胞免疫": ["T細胞", "T CELL", "CD4", "CD8", "TH1", "TH2", "TREG", "胸腺", "細胞激素", "CYTOKINE", "IL-", "干擾素", "IFN", "穿孔素", "FAS"],
        "B細胞與體液免疫": ["B細胞", "B CELL", "漿細胞", "PLASMA CELL", "記憶B", "BCR"],
        "MHC與移植免疫": ["MHC", "HLA", "移植", "排斥", "組織相容", "GRAFT", "GVHD"],
        "過敏反應": ["過敏", "HYPERSENSITIVITY", "氣喘", "肥大細胞", "MAST CELL", "組織胺", "第一型", "第二型", "第三型", "第四型", "ARTHUS", "接觸性皮膚炎"],
        "自體免疫疾病": ["自體免疫", "AUTOIMMUNE", "ANA", "SLE", "紅斑性狼瘡", "類風濕", "RF", "重症肌無力", "橋本氏", "HASHIMOTO", "GRAVES", "SCLERODERMA", "硬皮症", "乾燥症"],
        "腫瘤免疫": ["腫瘤", "癌症", "TUMOR", "CANCER", "癌", "CEA", "AFP", "PSA", "腫瘤標記", "免疫檢查點", "PD-1", "CTLA-4", "CAR-T", "SIPULEUCEL"],
        "疫苗與預防接種": ["疫苗", "VACCINE", "佐劑", "ADJUVANT", "被動免疫", "主動免疫", "減毒", "類毒素", "TOXOID"],
        "免疫檢驗技術": ["ELISA", "流式細胞", "FLOW CYTOMETRY", "螢光", "沉澱", "凝集", "西方墨點", "WESTERN BLOT", "免疫分析", "RIA", "VDRL", "RPR", "免疫電泳"]
    }
    
    best_topic = "其他綜合"
    max_hits = 0
    for topic, keywords in topics.items():
        hits = sum(1 for kw in keywords if kw in text)
        if hits > max_hits:
            max_hits, best_topic = hits, topic
    return best_topic

# --- 4. 核心解析引擎 (智慧排錯升級版) ---
def parse_unified_format(lines):
    questions = []
    current_q = None
    current_year = "未分類" 
    last_opt = None
    
    # 修正重點 2：放寬年份抓取，不再強求必須有「醫檢師」三個字
    year_pattern = re.compile(r'(\d{2,4})\s*年.*?第?\s*[一二12]\s*次?')
    opt_pattern = re.compile(r'\(([A-E])\)\s*(.*?)(?=\s*\([A-E]\)|$)')
    
    for line in lines:
        clean_line = line.strip()
        if not clean_line: continue
        
        year_match = year_pattern.search(clean_line)
        if year_match and ("次" in clean_line or "醫檢" in clean_line or "解析" in clean_line or "試題" in clean_line):
            # 濾除多餘空白，確保年份格式一致 (如: 106年第二次)
            current_year = year_match.group(0).replace(" ", "")
            continue
            
        q_match_full = re.match(r'^\s*\(([^)]+)\)\s*(\d+)[\.、]\s*(.*)', clean_line)
        is_new_q = False
        ans, num, q_text = None, None, ""

        if q_match_full:
            ans, num_str, q_text = q_match_full.groups()
            num = int(num_str)
            is_new_q = True
        else:
            q_match_missing_ans = re.match(r'^\s*(\d+)[\.、]\s*(.*)', clean_line)
            if q_match_missing_ans:
                temp_num = int(q_match_missing_ans.group(1))
                if not current_q or temp_num == 1 or (current_q and temp_num == current_q["question_number"] + 1):
                    num = temp_num
                    q_text = q_match_missing_ans.group(2)
                    ans = ""
                    is_new_q = True

        if is_new_q:
            opt_matches = opt_pattern.findall(q_text)
            if opt_matches:
                q_text = re.split(r'\s*\([A-E]\)', q_text, 1)[0].strip()
                
            if current_q:
                current_q["explanation"] = current_q["explanation"].strip()
                current_q["tags"]["主題"] = auto_classify_topic(current_q)
                questions.append(current_q)
            
            # 初始化時就會帶入正確抓取到的 current_year
            current_q = {
                "question_number": num,
                "question_text": q_text,
                "answer": ans.strip() if ans else "未提供",
                "options": {},
                "explanation": "",
                "tags": {"年份": current_year}
            }
            last_opt = None
            
            if opt_matches:
                for opt_letter, opt_text in opt_matches:
                    current_q["options"][opt_letter] = opt_text.strip()
                    last_opt = opt_letter
            continue
            
        if not current_q: continue
            
        opt_matches = opt_pattern.findall(clean_line)
        if opt_matches and not current_q["explanation"]:
            prefix_text = re.split(r'\s*\([A-E]\)', clean_line, 1)[0].strip()
            if prefix_text:
                current_q["question_text"] += "\n" + prefix_text

            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
                last_opt = opt_letter
            continue
            
        # 修正重點 3：放寬解析開頭的辨識，容許空白
        if re.search(r'解\s*析\s*[:：]', clean_line):
            exp_text = re.sub(r'^.*?(?:解\s*析)\s*[:：]\s*', '', clean_line)
            clean_exp, updated_tags = extract_tags_and_clean(exp_text, current_q["tags"])
            current_q["tags"] = updated_tags
            current_q["explanation"] += clean_exp + "\n"
            continue
            
        if not current_q["options"] and not current_q["explanation"]:
            current_q["question_text"] += "\n" + clean_line
        elif current_q["options"] and not current_q["explanation"]:
            if last_opt and clean_line:
                current_q["options"][last_opt] += "\n" + clean_line
        elif current_q["explanation"]:
            clean_exp, updated_tags = extract_tags_and_clean(clean_line, current_q["tags"])
            current_q["tags"] = updated_tags
            if clean_exp:
                current_q["explanation"] += clean_exp + "\n"

    # 確保最後一題也有順利加上主題分類
    if current_q:
        current_q["explanation"] = current_q["explanation"].strip()
        current_q["tags"]["主題"] = auto_classify_topic(current_q)
        questions.append(current_q)
        
    return questions

# --- 5. 網頁介面設計 ---
st.set_page_config(page_title="國考題庫極速轉檔", page_icon="⚡", layout="wide")
st.title("⚡ 國考題庫：極速轉檔工具 (支援圖片壓縮版)")

col1, col2 = st.columns([1, 2])
with col1:
    uploaded_file = st.file_uploader("上傳含有圖片的 Word 檔案 (.docx)", type=['docx'])
    if uploaded_file is not None:
        with st.spinner('正在抽取圖文與壓縮圖片中，請稍候...'):
            try:
                file_stream = io.BytesIO(uploaded_file.read())
                lines = extract_raw_text(file_stream)
                parsed_data = parse_unified_format(lines)
                st.session_state['parsed_data'] = parsed_data
                st.session_state['file_name'] = uploaded_file.name
                st.success("✅ 解析完成！圖片已自動壓縮並封裝，年份與主題分類皆已恢復。")
            except Exception as e:
                st.error(f"❌ 發生錯誤：{e}")

    if 'parsed_data' in st.session_state:
        json_bytes = json.dumps(st.session_state['parsed_data'], ensure_ascii=False, indent=4).encode('utf-8')
        st.download_button(
            label="📥 下載含壓縮圖片之 JSON 題庫檔",
            data=json_bytes,
            file_name=st.session_state['file_name'].replace(".docx", "_輕量版.json"),
            mime="application/json",
            use_container_width=True
        )

with col2:
    st.subheader("🔍 解析結果預覽 (純文字檢視)")
    if 'parsed_data' in st.session_state:
        st.info("已啟用自動化圖片壓縮 (最大寬度 800px, JPEG 格式)。")
        st.json(st.session_state['parsed_data'][:2])
