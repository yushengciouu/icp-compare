import os
import glob
import re
import json
from bs4 import BeautifulSoup
import pandas as pd
from openai import OpenAI
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# 1. 初始化 OpenAI 客戶端，連接您的本地 vLLM (gemma-4:31B)
client = OpenAI(
    base_url="http://192.168.39.143:8002/v1",
    api_key="empty_api_key_for_vllm"
)

# 本地模型 ID 與併發限制線程數
MODEL_NAME = "gemma-4:31B"
MAX_WORKERS = 64

def clean_text(text):
    if text is None:
        return ""
    return str(text).strip()

def parse_xml_file(xml_path):
    """
    解析 XML 檔案，過濾出 percentage >= 75.0% 的所有 Result 對照紀錄。
    """
    print(f"正在讀取並解析 XML 檔案: {os.path.basename(xml_path)}...")
    with open(xml_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    soup = BeautifulSoup(content, 'xml')
    results = soup.find_all('Result')
    
    matched_pairs = []
    
    for r in results:
        condition = r.find('Condition')
        if not condition:
            continue
            
        condition_id = clean_text(condition.find('ConditionId').text if condition.find('ConditionId') else "")
        query_name = clean_text(condition.find('QueryName').text if condition.find('QueryName') else "")
        country = clean_text(condition.find('Country').text if condition.find('Country') else "")
        city = clean_text(condition.find('City').text if condition.find('City') else "")
        address = clean_text(condition.find('Address').text if condition.find('Address') else "")
        
        # 尋找底下的所有 Party 節點
        parties = r.find_all('Party')
        for p in parties:
            # 兼容大小寫 groupid / Groupid 等
            group_id_node = p.find(re.compile('^groupid$', re.I))
            percentage_node = p.find(re.compile('^percentage$', re.I))
            name_node = p.find(re.compile('^name$', re.I))
            addr_node = p.find(re.compile('^address$', re.I))
            content_node = p.find(re.compile('^content$', re.I))
            
            group_id = clean_text(group_id_node.text if group_id_node else "")
            pct_str = clean_text(percentage_node.text if percentage_node else "0")
            party_name = clean_text(name_node.text if name_node else "")
            party_address = clean_text(addr_node.text if addr_node else "")
            party_content = clean_text(content_node.text if content_node else "")
            
            try:
                percentage = float(pct_str)
            except ValueError:
                percentage = 0.0
                
            # 只篩選出命中率 >= 75% 的
            if percentage >= 75.0:
                matched_pairs.append({
                    "condition_id": condition_id,
                    "query_name": query_name,
                    "query_country": country,
                    "query_city": city,
                    "query_address": address,
                    "party_id": group_id,
                    "party_name": party_name,
                    "party_address": party_address,
                    "party_content": party_content,
                    "xml_percentage": percentage
                })
                
    return matched_pairs

def get_llm_judgment(pair):
    """
    呼叫本地 gemma-4:31B 進行深度語意比對。
    """
    prompt = f"""你是一名專業的進出口貿易合規審計專家（ICP/RPS Compliance Auditor）。
我們需要比對一筆「客戶/供應商查詢條件」與一筆「戰略性高科技貨品限制實體（黑名單）資料」，評估兩者是否為同一個實體、關係企業（如分公司、子公司、母公司、轉投資公司等），或是純粹的「字面重疊誤判（False Positive）」。

【查詢條件 (Condition)】
- 條件 ID: {pair['condition_id']}
- 查詢名稱 (QueryName): {pair['query_name']}
- 國家 (Country): {pair['query_country']}
- 城市 (City): {pair['query_city']}
- 查詢地址 (Address): {pair['query_address']}

【黑名單實體 (Party)】
- 限制實體 ID: {pair['party_id']}
- 限制實體名稱 (Name): {pair['party_name']}
- 限制實體地址 (Address): {pair['party_address']}
- 限制實體關聯詳細內容 (Content): {pair['party_content']}

請遵循以下思考框架進行細緻的語意分析：
1. **名稱主體分析 (Name Core Element)**：
   - 移除常見法律後綴詞（如 GMBH, CO. KG, INC, LTD, PTE LTD, LLC 等）。
   - 比對兩者核心名稱（例如 EBM-PAPST 是否與黑名單中的名稱實質關聯）。
   - 注意中英文對譯（例如 廈門算能科技 與 Xiamen Sophgo）。
2. **地址物理一致性與地理常識分析 (Address & Geography)**：
   - 是否在同一棟商業大樓或物流園區？如果是像 Software Park (軟體園)、Midview City、工業園區等，多個不相關的公司共用同一地址是常見的。這時若公司名稱不同，大概率是 False Positive。
   - 如果是「ShipTo (出貨地址)」比對，即使公司名字不同，若出貨的物理地址完全一致（特別是貨代倉庫或敏感地址），則仍具備高度的轉運合規風險。
3. **關聯性與別名解析 (Alias & Associations)**：
   - 檢查黑名單 Detail (Content) 中是否有列出 alias (別名) 或轉投資股權結構，看看查詢原名是否正是其別名之一。

請做出最終的合規審計判定：
- **High (同一實體)**: 核心名稱完全對上，且物理地址高度相符。
- **Medium (關聯企業)**: 核心名稱相同但位於不同國家/分支機構，或是物理地址完全一致但名稱看似不同（有出貨轉運風險）。
- **Low (低風險/懷疑)**: 有些微相似度，但無法判定。
- **False Positive (確定誤判)**: 字面相似或因共享園區大樓被篩出，但兩者名稱品牌毫不相干，且非同一實體。

請「僅」回覆以下標準的 JSON 格式（不要包含任何前後引導廢話或 markdown 的 ` ```json ` 標記，純 JSON 內容，確保可以被 json.loads 解析）：
{{
  "reasoning": "繁體中文的優雅流暢分析推理過程，列出名稱、地址及關聯性的具體論據（注意：切勿在 reasoning 的字串內部使用未經轉義的雙引號，若需用引號請使用單引號或是 \\\"）",
  "match_level": "High" / "Medium" / "Low" / "False Positive",
  "is_same_entity": true / false
}}
"""
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=1000
        )
        resp_text = response.choices[0].message.content.strip()
        
        # 嘗試解析 JSON 內容
        # 移除可能夾帶的 ```json  ``` 區塊
        clean_json_str = resp_text
        if "```" in resp_text:
            match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", resp_text, re.DOTALL)
            if match:
                clean_json_str = match.group(1)
            else:
                # 簡單清理
                clean_json_str = resp_text.replace("```json", "").replace("```", "").strip()
        
        # 硬性防禦機制：如果模型在 JSON 中使用了非法的未轉義雙引號
        # 這邊我們先把常見的 Markdown 程式碼區塊標記徹底移除
        clean_json_str = clean_json_str.strip()
        
        try:
            judgment = json.loads(clean_json_str)
        except json.JSONDecodeError:
            # 進一步容錯：很多時候是 reasoning 的雙引號或者是換行符號問題。
            # 我們嘗試用更寬鬆的正則表達式把 match_level 與 is_same_entity 撈出來
            match_level_found = "Uncertain"
            is_same_found = False
            reasoning_found = "解析失敗，改由正則提取"
            
            lvl_match = re.search(r'"match_level"\s*:\s*"([^"]+)"', clean_json_str, re.I)
            same_match = re.search(r'"is_same_entity"\s*:\s*(true|false)', clean_json_str, re.I)
            reason_match = re.search(r'"reasoning"\s*:\s*"(.+?)"\s*,\s*"(?:match_level|is_same_entity)"', clean_json_str, re.DOTALL | re.I)
            
            if lvl_match:
                match_level_found = lvl_match.group(1)
            if same_match:
                is_same_found = same_match.group(1).lower() == "true"
            if reason_match:
                reasoning_found = reason_match.group(1).replace('\\"', '"').replace('\n', ' ')
            else:
                # 如果無法定位 reasoning 的結尾，就用原本的內容
                reasoning_found = f"原始生成：{resp_text}"
                
            judgment = {
                "reasoning": reasoning_found,
                "match_level": match_level_found,
                "is_same_entity": is_same_found
            }
            
        return judgment
    except Exception as e:
        print(f"呼叫 LLM 發生錯誤: {e}. 原始回覆: {resp_text if 'resp_text' in locals() else ''}")
        return {
            "reasoning": f"解析 LLM 失敗：{str(e)}",
            "match_level": "Uncertain",
            "is_same_entity": False
        }

def process_single_pair(pair, idx, total_count):
    """
    包裝單筆比對任務，用於多執行緒併發處理。
    """
    judgment = get_llm_judgment(pair)
    item = {
        "來源檔案": pair["source_file"],
        "條件ID": pair["condition_id"],
        "查詢名稱": pair["query_name"],
        "查詢國家": pair["query_country"],
        "查詢城市": pair["query_city"],
        "查詢地址": pair["query_address"],
        "黑名單ID": pair["party_id"],
        "黑名單名稱": pair["party_name"],
        "黑名單地址": pair["party_address"],
        "黑名單完整資訊": pair["party_content"][:500] + "..." if len(pair["party_content"]) > 500 else pair["party_content"],
        "原XML命中率": f"{pair['xml_percentage']}%",
        "LLM研判等級": judgment.get("match_level", "Uncertain"),
        "LLM判定是否同實體": "是" if judgment.get("is_same_entity") else "否",
        "LLM分析推理理由": judgment.get("reasoning", "無描述")
    }
    return item

def main():
    xml_files = glob.glob("testfile/*_raw.xml")
    if not xml_files:
        print("在 testfile/ 目錄下找不到任何以 _raw.xml 結尾的檔案！")
        return
        
    print(f"找到 {len(xml_files)} 個 raw XML 檔案。開始進行高風險（>=75%）名單提取...")
    
    all_pairs = []
    for xml_path in xml_files:
        pairs = parse_xml_file(xml_path)
        file_name = os.path.basename(xml_path)
        for p in pairs:
            p["source_file"] = file_name
        all_pairs.extend(pairs)
        
    total_count = len(all_pairs)
    print(f"\n提取完成！在所有 XML 檔案中，共找到 {total_count} 筆命中率 >= 75.0% 的高風險候選資料。")
    
    if total_count == 0:
        print("非常安全！所有查詢結果中均無大於或等於 75% 的黑名單結果，無需進入 LLM 審判。")
        return
        
    print(f"現在啟動本地 LLM gemma-4:31B 進行平行（可配置 concurrent={MAX_WORKERS}）深度合規審查...")
    
    audit_results = []
    
    # 2. 多執行緒併發處理 (使用 ThreadPoolExecutor 並配置 tqdm 進度條)
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_single_pair, pair, idx, total_count): idx for idx, pair in enumerate(all_pairs, 1)}
        
        for future in tqdm(as_completed(futures), total=total_count, desc="LLM Auditing Progress", unit="pairs"):
            try:
                item = future.result()
                audit_results.append(item)
            except Exception as e:
                print(f"處理 Future 發生嚴重錯誤: {e}")
        
    # 3. 輸出成 Excel 報表 (每次產生包含時間戳記的全新檔案，完美避開 PermissionError)
    df = pd.DataFrame(audit_results)
    
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel = f"ICP_Audit_Report_{timestamp}.xlsx"
    
    try:
        df.to_excel(output_excel, index=False)
        print(f"\n✅ 審計完成！最新結果已成功輸出至全新檔案: {os.path.abspath(output_excel)}")
    except Exception as e:
        print(f"❌ 儲存 Excel 時發生非預期錯誤: {e}")
    
    # 4. 列出統計摘要
    print("\n========= 審計結果統計摘要 =========")
    summary_df = df["LLM研判等級"].value_counts().to_frame()
    print(summary_df)
    print("====================================")

if __name__ == "__main__":
    main()
