import streamlit as st
import docx
import re
import json
import io

# --- 新增：透視表格與段落的讀取工具 ---
def iter_block_items(doc):
    """按文件順序讀取 Word 檔內所有段落與表格內的文字"""
    for element in doc.element.body:
        if element.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(element, doc)
            if p.text.strip():
                yield p.text.strip()
        elif element.tag.endswith('tbl'):
            table = docx.table.Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if p.text.strip():
                            yield p.text.strip()

# --- 升級版：核心解析引擎 ---
def parse_exam_docx(file_stream):
    doc = docx.Document(file_stream)
    questions = []
    current_q = None
    
    # 寬容模式的正規表達式
    q_pattern = re.compile(r'^(?:\(([A-E])\)\s*)?(\d+)[\.、]\s*(.*)')
    opt_pattern = re.compile(r'\(([A-E])\)\s*([^()]+?)(?=\([A-E]\)|$)')
    
    # 改用我們自訂的 iter_block_items 來讀取，這樣就不會漏掉表格！
    for text in iter_block_items(doc):
            
        # 1. 抓取題目
        q_match = q_pattern.match(text)
        if q_match:
            if current_q:
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
            
        # 2. 抓取選項
        opt_matches = opt_pattern.findall(text)
        if opt_matches and current_q:
            for opt_letter, opt_text in opt_matches:
                current_q["options"][opt_letter] = opt_text.strip()
            continue
            
        # 3. 抓取解析與標籤
        if current_q:
            # 支援 "解  析:", "解析：", "【解析】" 等多種排版
            if re.match(r'^(?:【)?解\s*析(?:】)?\s*[:：]?', text):
                raw_exp = re.sub(r'^(?:【)?解\s*析(?:】)?\s*[:：]?\s*', '', text)
                
                # 處理標籤切割 (如難度、再現性)
                parts = re.split(r'","|",\s*"|"', raw_exp)
                current_q["explanation"] = parts[0].strip()
                
                for part in parts[1:]:
                    if ":" in part or "：" in part:
                        key_val = re.split(r'[:：]', part, 1)
                        if len(key_val) == 2:
                            k = key_val[0].replace(" ", "").strip()
                            v = key_val[1].split('(')[0].strip()
                            current_q["tags"][k] = v
                continue
                
            # 處理跨行文字 (選項還沒出來前算題目，解析出來後算解析的延伸)
            if not current_q["options"] and not current_q["explanation"]:
                current_q["question_text"] += "\n" + text
            elif current_q["explanation"] and not re.match(r'^(?:【)?解\s*析(?:】)?\s*[:：]?', text):
                current_q["explanation"] += "\n" + text

    # 記得收尾最後一題
    if current_q:
        questions.append(current_q)
        
    return questions
