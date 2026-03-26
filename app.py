import docx
import json
import re
import os

def normalize_text(text):
    """將文字標準化，解決排版不一的問題"""
    # 1. 轉半形括號
    text = text.replace('（', '(').replace('）', ')')
    # 2. 轉半形冒號
    text = text.replace('：', ':')
    # 3. 將多個空格(含全形空格)替換為單一半形空格
    text = re.sub(r'[\s\u3000]+', ' ', text)
    return text.strip()

def convert_docx_to_json_super_loose(docx_path, json_path):
    print(f"🔄 開始讀取檔案：{docx_path} ...")
    
    try:
        doc = docx.Document(docx_path)
    except Exception as e:
        print(f"❌ 讀取 Word 檔案失敗: {e}")
        return

    questions = []
    current_year = "未知年份"
    current_q = None
    
    # 🌟 1. 標題模糊匹配：只要有「數字+年+第+x+次」字眼就抓
    year_pattern = re.compile(r'(\d{3})\s*年\s*第\s*([一二])\s*次')
    
    # 🌟 2. 題號與答案【超級寬鬆匹配】：
    # 容許前方有任何 [source] 或 (Ans) 標記
    # 精準抓取類似 (A)30. 或 (Ans)30、 或 (A) 30. 的格式
    q_start_pattern = re.compile(r'^\s*.*?\(\s*(?P<ans>[A-Za-z,皆對送分]+)\s*\)\s*(?P<num>\d+)\s*[\.、\s]\s*(?P<text>.*)')
    
    # 🌟 3. 解析模糊匹配：容許「解 析 :」、「解 析:」、「解 析」開頭
    exp_pattern = re.compile(r'^解\s*析\s*[:\s](.*)', re.IGNORECASE)

    # 用來記錄跳過的行，方便除錯
    skipped_lines = []

    for para in doc.paragraphs:
        # 1. 先做文字標準化
        text = normalize_text(para.text)
        
        if not text:
            continue
            
        # --- 階段 A：判斷是否為年份標題 (模糊匹配) ---
        year_match = year_pattern.search(text)
        # 排除包含 (A) 的行，避免把題目當標題
        if year_match and "(A)" not in text and "()" not in text: 
            # 重新格式化年份，保持一致性，例如 "108年第1次"
            current_year = f"{year_match.group(1)}年第{year_match.group(2)}次"
            # 嘗試保留標題後方的完整文字，如 "免疫學及病毒學"
            title_suffix = text.split('次')[-1].strip()
            if title_suffix:
                current_year += " " + title_suffix
            print(f"📂 偵測到年份標題：{current_year}")
            continue
            
        # --- 階段 B：判斷是否為新題目 (超級寬鬆) ---
        q_match = q_start_pattern.match(text)
        if q_match:
            # 提交上一題
            if current_q:
                _extract_options_v2(current_q)
                questions.append(current_q)
            
            ans = q_match.group('ans').strip().upper()
            # 防呆：如果答案寫成了全形
            ans = ans.replace('，', ',')
            
            try:
                q_num = int(q_match.group('num'))
            except:
                q_num = 0 # 萬一題號不是數字
                
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
            
        # --- 階段 D：如果都不是，就是多行內容的延續 ---
        if current_q:
            # 防呆：如果「解析」藏在同一行的中後段
            if re.search(r'解\s*析\s*[:\s]', text, re.IGNORECASE):
                parts = re.split(r'解\s*析\s*[:\s]', text, maxsplit=1, flags=re.IGNORECASE)
                if len(parts) > 1:
                    if parts[0].strip():
                        current_q["_raw_text"] += "\n" + parts[0].strip()
                    current_q["explanation"] += parts[1].strip()
                    continue

            # 如果已經在讀解析了
            if current_q["explanation"]:
                current_q["explanation"] += "\n" + text
            else:
                # 否則就是題目或選項的換行
                current_q["_raw_text"] += "\n" + text
        else:
            # 記錄沒有被歸類的行，通常是雜訊或轉檔失敗的題目
            if len(text) > 5: # 忽略太短的雜訊
                skipped_lines.append(f"[{current_year}] {text}")

    # 迴圈結束，記得儲存最後一題
    if current_q:
        _extract_options_v2(current_q)
        questions.append(current_q)

    # 輸出成 JSON 檔案 (同時進行壓縮瘦身)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(questions, f, ensure_ascii=False, separators=(',', ':'))
        
    print(f"\n🎉 轉換完成！共擷取了 {len(questions)} 題。")
    print(f"💾 檔案已儲存至：{json_path}")

    # 將跳過的行輸出成報告，方便您檢查 Word 檔
    if skipped_lines:
        report_path = json_path.replace(".json", "_漏題檢查報告.txt")
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("⚠️ 以下是未被辨識為題目的文字，請檢查 Word 檔格式是否太過奇特：\n")
            f.write("--------------------------------------------------\n")
            for line in skipped_lines:
                f.write(line + "\n")
        print(f"📄 產生漏題檢查報告：{report_path} (請查看此檔案確認是否有題目漏抓)")


def _extract_options_v2(q_dict):
    """內部輔助函數：將題目與選項分離 (增強版)"""
    raw = q_dict.pop("_raw_text", "")
    
    # 尋找 (A) 的模糊匹配
    match_A = re.search(r'\(\s*A\s*\)', raw)
    
    if match_A:
        q_dict["question_text"] = raw[:match_A.start()].strip()
        opts_text = raw[match_A.start():]
        
        # 利用正規表達式抓取 (A)xxx (B)yyy
        opt_pattern = re.compile(r'\(\s*(?P<key>[A-E])\s*\)\s*(?P<val>.*?)(?=(?:\(\s*[A-E]\s*\))|$)')
        options = {}
        for m in opt_pattern.finditer(opts_text):
            options[m.group('key')] = m.group('val').strip()
            
        q_dict["options"] = options
    else:
        q_dict["question_text"] = raw.strip()
        q_dict["options"] = {}

# ==========================================
# 執行區塊
# ==========================================
if __name__ == "__main__":
    # 將您的 Word 檔放在同一層資料夾，並確認檔名正確
    input_docx = "臨床血清免疫學解析.docx"  # 請確認檔名
    output_json = "臨床血清免疫學_修復版.json" # 產生一個新檔名
    
    if os.path.exists(input_docx):
        convert_docx_to_json_super_loose(input_docx, output_json)
    else:
        print(f"⚠️ 找不到檔案 {input_docx}，請確認檔案是否放在同一個資料夾中！")
