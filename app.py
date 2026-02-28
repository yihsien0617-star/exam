import streamlit as st
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
import re
import json
import io

# --- 核心解析引擎 (無差別攤平版) ---
def parse_exam_docx(file_stream):
    doc = docx.Document(file_stream)
    
    # 1. 將整份文件攤平成為一行一行的純文字陣列
    lines = []
    seen_tc = set() # 用來避免合併儲存格造成文字重複抓取
    
    for element in doc.element.body:
        # 如果是段落
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            if para.text.strip():
                # 遇到段落內自帶換行，強制切開成獨立行
                for line in para.text.split('\n'):
                    if line.strip(): lines.append(line.strip())
        # 如果是表格
        elif element.tag.endswith('tbl'):
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    # 過濾掉合併儲存格產生的重複物件
                    if cell._tc in seen_tc:
                        continue
                    seen_tc.add(cell._tc)
                    if cell.text.strip():
                        # 表格內的文字也強制切成獨立行
                        for line in cell.text.split('\n'):
                            if line.strip(): lines.append(line.strip())

    # 2. 開始逐行精準捕捉
    questions = []
    current_q = None
    
    # 正規表達式 (允許前面有任何空白或雜訊)
    q_pattern = re.compile(r'^\s*(?:\(([A-E])\)\s*)?(\d+)[\.、]\s*(.*)')
    opt_pattern = re.compile(r'\(([A-E])\)\s*([^()]+?)(?=\([A-E]\)|$)')
    
    for text in lines:
        # 建立一個去除所有空白的字串，用來抵抗排版不一的干擾
        clean_text = text.replace(" ", "").replace("　", "")
        
        # [步驟 A] 抓題目
        q_match = q_pattern.match(text)
        if q_match:
            if current_q:
                # 儲存上一題前，清掉解析尾巴多餘的空白
                current_q["explanation"] = current_q["explanation"].strip()
                questions.append(current_q)
            
            ans, num, q_text = q_match.groups()
            current_q = {
                "question_number": int(num),
                "question_text": q_text.strip(),
                "answer": ans if ans else "未提供",
                "options": {},
                "explanation": "",
                "tags": {}
            }
            continue
            
        if not current_q:
            continue
            
        # [步驟 B] 抓選項
        opt_matches = opt_pattern.findall(text)
        if opt_matches and not current_q["explanation"]:
            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
            continue
            
        # [步驟 C] 抓解析 (只要句子裡含有解析兩字就觸發)
        if "解析:" in clean_text or "解析：" in clean_text or clean_text.startswith("解析"):
            exp_text = re.sub(r'^.*?解\s*析\s*[:：]?\s*', '', text)
            current_q["explanation"] += exp_text + "\n"
            continue
            
        # [步驟 D] 抓標籤 - 難度
        if "難度:" in clean_text or "難度：" in clean_text or clean_text.startswith("難度"):
            diff_text = re.sub(r'^.*?難\s*度\s*[:：]?\s*', '', text)
            # 自動砍
