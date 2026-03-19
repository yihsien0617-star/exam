import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import re
import json
import io

# --- 1. 抽取純文字引擎 ---
def extract_raw_text(file_stream):
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
                        for line in cell.text.split('\n'):
                            if line.strip():
                                lines.append(line.strip())
    return lines

# --- 2. 標籤與解析淨化工具 ---
def extract_tags_and_clean(text, current_tags):
    diff_match = re.search(r'難\s*度[:：]\s*([^\(,"”]+)', text)
    if diff_match:
        current_tags["難度"] = diff_match.group(1).strip()
        
    rep_match = re.search(r'再\s*現\s*性[:：]\s*([^\(,"”]+)', text)
    if rep_match:
        current_tags["再現性"] = rep_match.group(1).strip()
        
    clean_exp = re.sub(r'[,"]*\s*難\s*度[:：][^"]+"?', '', text)
    clean_exp = re.sub(r'[,"]*\s*再\s*現\s*性[:：][^"]+"?', '', clean_exp)
    clean_exp = re.sub(r'\([^\)]*極低[^\)]*\)', '', clean_exp)
    clean_exp = re.sub(r'\([^\)]*非常簡單[^\)]*\)', '', clean_exp)
    
    return clean_exp.strip('", '), current_tags

# --- 3. 新增：自動主題分類引擎 ---
def auto_classify_topic(q_dict):
    """根據題目與解析的關鍵字，自動將題目歸類到對應的免疫學主題"""
    text = q_dict["question_text"] + " " + q_dict.get("explanation", "")
    for opt in q_dict["options"].values():
        text += " " + opt
    
    text = text.upper()
    
    # 老師可以隨時在這裡新增或修改「單元名稱」與對應的「關鍵字」
    topics = {
        "先天免疫與發炎反應": ["發炎", "白血球", "吞噬", "巨噬細胞", "MACROPHAGE", "NEUTROPHIL", "NK細胞", "自然殺手", "TLR", "先天免疫", "發燒", "C-REACTIVE", "CRP", "急性期"],
        "補體系統": ["補體", "COMPLEMENT", "C3", "C4", "C5", "MAC", "古典途徑", "替代途徑", "凝集素途徑", "C1Q"],
        "抗體與免疫球蛋白": ["抗體", "免疫球蛋白", "IGG", "IGA", "IGM", "IGE", "IGD", "ISOTYPE", "輕鏈", "重鏈", "FAB", "FC", "ALLOTYPE", "IDIOTYPE"],
        "T細胞與細胞免疫": ["T細胞", "T CELL", "CD4", "CD8", "TH1", "TH2", "TREG", "胸腺", "細胞激素", "CYTOKINE", "IL-", "干擾素", "IFN", "穿孔素", "FAS"],
        "B細胞與體液免疫": ["B細胞", "B CELL", "漿細胞", "PLASMA CELL", "記憶B", "BCR"],
        "MHC與移植免疫": ["MHC", "HLA", "移植", "排斥", "組織相容", "GRAFT", "GVHD"],
        "過敏反應 (Hypersensitivity)": ["過敏", "HYPERSENSITIVITY", "氣喘", "肥大細胞", "MAST CELL", "組織胺", "第一型", "第二型", "第三型", "第四型", "ARTHUS", "接觸性皮膚炎"],
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
            max_hits = hits
            best_topic = topic
            
    return best_topic

# --- 4. 核心精準解析引擎 ---
def parse_unified_format(lines):
    questions = []
    current_q = None
    current_year = "未分類" 
    
    year_pattern = re.compile(r'(\d{3})\s*年\s*(第[一二]次)')
    q_pattern = re.compile(r'^\s*\(([A-E])\)\s*(\d+)[\.、]\s*(.*)')
    opt_pattern = re.compile(r'\(([A-E])\)\s*([^()]+?)(?=\([A-E]\)|$)')
    
    for line in lines:
        clean_line = line.strip()
        
        year_match = year_pattern.search(clean_line)
        if year_match and "醫檢師" in clean_line: 
            current_year = f"{year_match.group(1)}年{year_match.group(2)}"
            continue
            
        q_match = q_pattern.match(clean_line)
        if q_match:
            if current_q:
                current_q["explanation"] = current_q["explanation"].strip()
                # 存入前，讓系統自動判斷並寫入「主題」標籤
                current_q["tags"]["主題"] = auto_classify_topic(current_q)
                questions.append(current_q)
            
            ans, num, q_text = q_match.groups()
            current_q = {
                "question_number": int(num),
                "question_text": q_text.strip(),
                "answer": ans,
                "options": {},
                "explanation": "",
                "tags": {
                    "年份": current_year 
                }
            }
            continue
            
        if not current_q:
            continue
            
        opt_matches = opt_pattern.findall(clean_line)
        if opt_matches and not current_q["explanation"]:
            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
            continue
            
        if "解  析:" in clean_line or "解析:" in clean_line or "解析：" in clean_line:
            exp_text = re.sub(r'^.*?(?:解\s*析)[:：]\s*', '', clean_line)
            clean_exp, updated_tags = extract_tags_and_clean(exp_text, current_q["tags"])
            current_q["tags"] = updated_tags
            current_q["explanation"] += clean_exp + "\n"
            continue
            
        if not current_q["options"] and not current_q["explanation"]:
            current_q["question_text"] += "\n" + clean_line
        elif current_q["explanation"]:
            clean_exp, updated_tags = extract_tags_and_clean(clean_line, current_q["tags"])
            current_q["tags"] = updated_tags
            if clean_exp:
                current_q["explanation"] += clean_exp + "\n"

    if current_q:
        current_q["explanation"] = current_q["explanation"].strip()
        current_q["tags"]["主題"] = auto_classify_topic(current_q)
        questions.append(current_q)
        
    return questions

# --- 5. 網頁介面設計 ---
st.set_page_config(page_title="國考題庫極速轉檔", page_icon="⚡", layout="wide")

st.title("⚡ 國考題庫：極速轉檔工具 (年份與單元主題全自動版)")
st.markdown("上傳檔案後，系統不僅會偵測各章節的年份，還會自動依據專有名詞幫題目進行**「單元主題歸類」**！")

col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("📁 檔案上傳區")
    uploaded_file = st.file_uploader("上傳已統一格式的 Word 檔案 (.docx)", type=['docx'])
    
    if uploaded_file is not None:
        with st.spinner('正在極速解析與自動分類中...'):
            try:
                file_stream = io.BytesIO(uploaded_file.read())
                lines = extract_raw_text(file_stream)
                parsed_data = parse_unified_format(lines)
                
                st.session_state['parsed_data'] = parsed_data
                st.session_state['file_name'] = uploaded_file.name
                st.success(f"✅ 解析完成！共擷取 {len(parsed_data)} 道題目。")
                
            except Exception as e:
                st.error(f"❌ 發生錯誤：{e}")

    if 'parsed_data' in st.session_state:
        json_str = json.dumps(st.session_state['parsed_data'], ensure_ascii=False, indent=4)
        st.download_button(
            label="📥 下載完整 JSON 題庫檔",
            data=json_str,
            file_name=st.session_state['file_name'].replace(".docx", "_完整分類.json"),
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
