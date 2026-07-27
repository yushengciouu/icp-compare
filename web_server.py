import os
import glob
import re
import json
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
import pandas as pd

# 匯入現有的審計核心模組
from compare_audit import parse_xml_file, get_llm_judgment, generate_formatted_excel_report

BASE_DIR = Path(__file__).parent.resolve()
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"
TESTFILE_DIR = BASE_DIR / "testfile"

STATIC_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
TESTFILE_DIR.mkdir(exist_ok=True)

app = FastAPI(title="ICP-Compare Web 戰略出口合規審計平台")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# 狀態記錄（即時執行狀態）
AUDIT_STATUS = {
    "is_running": False,
    "progress": 0,
    "total": 0,
    "completed": 0,
    "current_file": "",
    "last_completed_at": None,
    "last_report": ""
}

def get_latest_excel_path() -> Optional[Path]:
    files = list(BASE_DIR.glob("ICP_Audit_Report_*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]
    if not files:
        return None
    files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return files[0]

def load_excel_records(excel_path: Path) -> List[Dict[str, Any]]:
    if not excel_path.exists():
        return []
    
    df = pd.read_excel(excel_path, sheet_name=0)
    # 將 NaN 轉為 None / 空字串
    records = df.fillna("").to_dict(orient="records")
    return records

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    html_path = TEMPLATES_DIR / "index.html"
    if not html_path.exists():
        raise HTTPException(status_code=404, detail="前端模板 index.html 尚未建立")
    return HTMLResponse(content=html_path.read_text(encoding="utf-8"))

@app.get("/api/audit/latest")
async def get_latest_audit_results():
    latest_file = get_latest_excel_path()
    if not latest_file:
        return JSONResponse({"status": "empty", "message": "尚未發現任何審計報表", "records": [], "stats": {}})
    
    records = load_excel_records(latest_file)
    
    # 統計數量
    total = len(records)
    high_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "High")
    medium_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "Medium")
    low_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "Low")
    fp_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "False Positive")
    
    auto_release_rate = round((fp_count / total * 100), 1) if total > 0 else 0.0
    
    stats = {
        "total": total,
        "high": high_count,
        "medium": medium_count,
        "low": low_count,
        "fp": fp_count,
        "auto_release_rate": auto_release_rate,
        "file_name": latest_file.name,
        "last_updated": datetime.fromtimestamp(latest_file.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    }
    
    return {
        "status": "success",
        "file_name": latest_file.name,
        "stats": stats,
        "records": records
    }

@app.get("/api/reports")
async def list_reports():
    files = list(BASE_DIR.glob("ICP_Audit_Report_*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]
    files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    
    report_list = []
    for f in files:
        report_list.append({
            "name": f.name,
            "size": f"{round(f.stat().st_size / 1024, 1)} KB",
            "modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        })
    return {"reports": report_list}

@app.get("/api/reports/download/{filename}")
async def download_report(filename: str):
    file_path = BASE_DIR / filename
    if not file_path.exists() or not file_path.name.endswith(".xlsx"):
        raise HTTPException(status_code=404, detail="檔案不存在")
    return FileResponse(path=file_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.post("/api/upload")
async def upload_xml_files(files: List[UploadFile] = File(...)):
    saved_files = []
    for file in files:
        if not file.filename.endswith(".xml"):
            continue
        target_path = TESTFILE_DIR / file.filename
        content = await file.read()
        target_path.write_bytes(content)
        saved_files.append(file.filename)
        
    return {"status": "success", "uploaded": saved_files, "message": f"成功上傳 {len(saved_files)} 個 XML 檔案"}

def run_background_audit():
    global AUDIT_STATUS
    AUDIT_STATUS["is_running"] = True
    AUDIT_STATUS["progress"] = 0
    
    xml_files = glob.glob(str(TESTFILE_DIR / "*_raw.xml"))
    if not xml_files:
        xml_files = glob.glob(str(TESTFILE_DIR / "*.xml"))
        
    all_pairs = []
    for xml_file in xml_files:
        pairs = parse_xml_file(xml_file)
        for p in pairs:
            p["source_file"] = os.path.basename(xml_file)
        all_pairs.extend(pairs)
        
    if not all_pairs:
        AUDIT_STATUS["is_running"] = False
        return
        
    AUDIT_STATUS["total"] = len(all_pairs)
    AUDIT_STATUS["completed"] = 0
    
    results = []
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from compare_audit import MAX_WORKERS
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_pair = {executor.submit(get_llm_judgment, pair): pair for pair in all_pairs}
        for future in as_completed(future_to_pair):
            pair = future_to_pair[future]
            judgment = future.result()
            
            row = {
                "來源檔案": pair.get("source_file", ""),
                "條件ID": pair.get("condition_id", ""),
                "查詢名稱": pair.get("query_name", ""),
                "查詢國家": pair.get("query_country", ""),
                "查詢城市": pair.get("query_city", ""),
                "查詢地址": pair.get("query_address", ""),
                "黑名單ID": pair.get("party_id", ""),
                "黑名單名稱": pair.get("party_name", ""),
                "黑名單地址": pair.get("party_address", ""),
                "黑名單完整資訊": pair.get("party_content", ""),
                "原XML命中率": f"{pair.get('xml_percentage', 0.0)}%",
                "LLM研判等級": judgment.get("match_level", "False Positive"),
                "風險判定依據": judgment.get("risk_factor", ""),
                "LLM分析推理理由": judgment.get("reasoning", "")
            }
            results.append(row)
            AUDIT_STATUS["completed"] += 1
            AUDIT_STATUS["progress"] = int((AUDIT_STATUS["completed"] / AUDIT_STATUS["total"]) * 100)
            
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_excel = BASE_DIR / f"ICP_Audit_Report_{timestamp}.xlsx"
    
    df = pd.DataFrame(results)
    df.to_excel(output_excel, sheet_name="全審計結果清單", index=False)
    generate_formatted_excel_report(str(output_excel))
    
    AUDIT_STATUS["is_running"] = False
    AUDIT_STATUS["last_completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    AUDIT_STATUS["last_report"] = output_excel.name

@app.post("/api/audit/run")
async def trigger_audit(background_tasks: BackgroundTasks):
    if AUDIT_STATUS["is_running"]:
        return {"status": "running", "message": "審計作業正在執行中，請稍候..."}
    
    background_tasks.add_task(run_background_audit)
    return {"status": "started", "message": "已啟動全自動 LLM 合規審計作業！"}

@app.get("/api/audit/status")
async def get_audit_status():
    return AUDIT_STATUS

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_server:app", host="127.0.0.1", port=8000, reload=True)
