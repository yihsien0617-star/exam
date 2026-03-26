import docx
import json
import re
import os

def normalize_text(text):
    """將文字標準化，解決排版不一的問題"""
    text = text.replace('（', '(').replace('）', ')')
    text = text.replace('：', ':')
    text = re.sub(r'[\s\u3000]+', ' ', text)
    return text.strip()

def convert_docx_to_json_v3(docx_path, json_path):
    print(f"🔄 開始讀取檔案：{docx_path} ...")
    
    try:
        doc = docx.Document(docx_path)
    except Exception as e:
        print(f"❌ 讀取 Word 檔案失敗: {e}")
        return

    questions = []
    current_year = "未知年份"
    current_q = None
    
    # 🌟 1. 標題終極匹配：只要開頭有「數字2~4碼 + 年」，不管後面寫什麼都抓！
    year_pattern = re.compile(r'(\d{2,4})\s*年')
    
    # 🌟 2. 題目終極匹配：無視前面的 [source] 或雜訊，直接找 (答案) + 數字 + 點
    # 例如： (D) 1. 或是 (皆對) 27. 都可以抓到
    q_start_pattern = re.compile(r'^.*?[\(]\s*(?P<ans>[A-Ea-e,皆全對送分]+)\s*[\)]\s*(?P<num>\d+)\s*[.、\s]\s*(?P<text>.*)')
    
    # 🌟 3. 解析終極匹配：無視開頭的引號 "、逗號 , 或是破折號
    exp_pattern = re.compile(r'^[\"\'\,\.\-\s]*解\s*析\s*[:\s](.*)', re.IGNORECASE)

    skipped_lines = []

    for para in doc.paragraphs:
        text = normalize_text(para.text)
        if not text:
            continue
            
        # --- 階段 A：判斷是否為年份標題 ---
        year_match = year_pattern.search(text)
        # 必須確保這行沒有題目的特徵 (例如沒有 (A) 1. 這種格式)，才認定為標題
        if year_match and not q_start_pattern.search(text): 
            # 抓取「年」前面的所有字 + 後面的所有字當作標題
            current_year = text.replace('"', '').replace(',', '').strip()
            print(f"📂 進入新區塊：{current_year}")
            continue
            
        # --- 階段 B：判斷是否為新題目 ---
        q_match = q_start_pattern.match(text)
        if q_match:
            if current_q:
                _extract_options_v3(current_q)
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
                    "主題": "未分類"
                },
                "_raw_text": q_text
            }
            continue
            
        # --- 階段 C：判斷是否為解析開頭 ---
        exp_match = exp_pattern.match(text)
        if exp_match and current_q:
            current_q["explanation"] = exp_match.group(1).strip()
            continue
            
        # --- 階段 D：多行內容延續 ---
        if current_q:
            # 處理藏在同一行中後段的解析 (無視前面的引號或逗號)
            hidden_exp_match = re.search(r'[\"\'\,\.\-\s]*解\s*析\s*[:\s]', text, re.IGNORECASE)
            if hidden_exp_match:
                split_idx = hidden_exp_match.start()
                q_part = text[:split_idx].strip()
                exp_part = text[hidden_exp_match.end():].strip()
                
                if q_part:
                    current_q["_raw_text"] += "\n" + q_part
                current_q["explanation"] += exp_part
                continue

            # 串接換行文字
            if current_q["explanation"]:
                current_q["explanation"] += "\n" + text
            else:
                current_q["_raw_text"] += "\n" + text
        else:
            if len(text) > 5:
                skipped_lines.append(f"[{current_year}] {text}")

    if current_q:
        _extract_options_v3(current_q)
        questions.append(current_q)

    # 輸出成極限壓縮版 JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, separators=(',', ':'))
        
    print(f"\n🎉 轉換大功告成！共擷取了 {len(questions)} 題。")
    
    if skipped_lines:
        report_path = json_path.replace(".json", "_漏題報告.txt")
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("⚠️ 以下是未被辨識的文字，請確認是否有漏網之魚：\n\n")
            for line in skipped_lines:
                f.write(line + "\n")
        print(f"📄 發現可能有無法辨識的雜訊，已輸出至：{report_path}")

def _extract_options_v3(q_dict):
    """將題目與選項分離"""
    raw = q_dict.pop("_raw_text", "")
    
    match_A = re.search(r'\(\s*[A]\s*\)', raw)
    if match_A:
        q_dict["question_text"] = raw[:match_A.start()].strip()
        opts_text = raw[match_A.start():]
        
        opt_pattern = re.compile(r'\(\s*(?P<key>[A-E])\s*\)\s*(?P<val>.*?)(?=(?:\(\s*[A-E]\s*\))|$)')
        options = {}
        for m in opt_pattern.finditer(opts_text):
            options[m.group('key')] = m.group('val').strip()
        q_dict["options"] = options
    else:
        q_dict["question_text"] = raw.strip()
        q_dict["options"] = {}

if __name__ == "__main__":
    input_docx = "臨床血清免疫學解析.docx"  
    output_json = "臨床血清免疫學_極限修復版.json"
    
    if os.path.exists(input_docx):
        convert_docx_to_json_v3(input_docx, output_json)
    else:
        print(f"⚠️ 找不到檔案 {input_docx}")
