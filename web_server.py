import os
import glob
import re
import json
import uuid
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Query
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
import pandas as pd

# 匯入現有的審計核心模組
from compare_audit import parse_xml_file, get_llm_judgment, generate_formatted_excel_report

BASE_DIR = Path(__file__).parent.resolve()
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"
JOBS_DIR = BASE_DIR / "jobs"

STATIC_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
JOBS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="ICP-Compare Web 戰略出口合規審計平台 (Multi-User Job Isolation)")
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# 多任務動態狀態字典: job_id -> { is_running, progress, total, completed, last_report, records, stats }
JOBS_REGISTRY: Dict[str, Dict[str, Any]] = {}

def get_job_dir(job_id: str) -> Path:
    job_path = JOBS_DIR / job_id
    job_path.mkdir(parents=True, exist_ok=True)
    (job_path / "uploads").mkdir(exist_ok=True)
    return job_path

def load_excel_records(excel_path: Path) -> List[Dict[str, Any]]:
    if not excel_path.exists():
        return []
    df = pd.read_excel(excel_path, sheet_name=0)
    return df.fillna("").to_dict(orient="records")

def calculate_stats(records: List[Dict[str, Any]], file_name: str, mod_time: str) -> Dict[str, Any]:
    total = len(records)
    high_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "High")
    medium_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "Medium")
    low_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "Low")
    fp_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "False Positive")
    auto_release_rate = round((fp_count / total * 100), 1) if total > 0 else 0.0
    return {
        "total": total,
        "high": high_count,
        "medium": medium_count,
        "low": low_count,
        "fp": fp_count,
        "auto_release_rate": auto_release_rate,
        "file_name": file_name,
        "last_updated": mod_time
    }

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    html_path = TEMPLATES_DIR / "index.html"
    if not html_path.exists():
        raise HTTPException(status_code=404, detail="前端模板 index.html 尚未建立")
    return HTMLResponse(content=html_path.read_text(encoding="utf-8"))

@app.post("/api/upload")
async def upload_xml_files(files: List[UploadFile] = File(...), job_id: Optional[str] = Query(None)):
    if not job_id:
        job_id = f"job_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    
    job_path = get_job_dir(job_id)
    upload_dir = job_path / "uploads"
    
    # 清空該 job 之前的上傳舊檔
    for old_file in upload_dir.glob("*.xml"):
        try:
            old_file.unlink()
        except Exception:
            pass
            
    saved_files = []
    for file in files:
        if not file.filename.endswith(".xml"):
            continue
        target_path = upload_dir / file.filename
        content = await file.read()
        target_path.write_bytes(content)
        saved_files.append(file.filename)
        
    JOBS_REGISTRY[job_id] = {
        "is_running": False,
        "progress": 0,
        "total": 0,
        "completed": 0,
        "last_completed_at": None,
        "last_report": "",
        "uploaded_files": saved_files
    }
    
    return {
        "status": "success",
        "job_id": job_id,
        "uploaded": saved_files,
        "message": f"成功上傳 {len(saved_files)} 個 XML 檔案至任務專屬目錄"
    }

def run_background_job_audit(job_id: str):
    if job_id not in JOBS_REGISTRY:
        JOBS_REGISTRY[job_id] = {}
        
    job_info = JOBS_REGISTRY[job_id]
    job_info["is_running"] = True
    job_info["progress"] = 0
    
    job_path = get_job_dir(job_id)
    upload_dir = job_path / "uploads"
    
    xml_files = glob.glob(str(upload_dir / "*_raw.xml"))
    if not xml_files:
        xml_files = glob.glob(str(upload_dir / "*.xml"))
        
    all_pairs = []
    for xml_file in xml_files:
        pairs = parse_xml_file(xml_file)
        for p in pairs:
            p["source_file"] = os.path.basename(xml_file)
        all_pairs.extend(pairs)
        
    if not all_pairs:
        job_info["is_running"] = False
        job_info["progress"] = 100
        return
        
    job_info["total"] = len(all_pairs)
    job_info["completed"] = 0
    
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
            job_info["completed"] += 1
            job_info["progress"] = int((job_info["completed"] / job_info["total"]) * 100)
            
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_name = f"ICP_Audit_Report_{timestamp}.xlsx"
    output_excel = job_path / report_name
    
    df = pd.DataFrame(results)
    df.to_excel(output_excel, sheet_name="全審計結果清單", index=False)
    generate_formatted_excel_report(str(output_excel))
    
    mod_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    job_info["is_running"] = False
    job_info["last_completed_at"] = mod_time
    job_info["last_report"] = report_name
    job_info["records"] = results
    job_info["stats"] = calculate_stats(results, report_name, mod_time)

@app.post("/api/audit/run")
async def trigger_audit(background_tasks: BackgroundTasks, job_id: str = Query(...)):
    job_path = get_job_dir(job_id)
    upload_dir = job_path / "uploads"
    if not list(upload_dir.glob("*.xml")):
        raise HTTPException(status_code=400, detail="該任務目錄下尚無 XML 報文檔案，請先上傳")
        
    job_info = JOBS_REGISTRY.get(job_id, {})
    if job_info.get("is_running"):
        return {"status": "running", "job_id": job_id, "message": "任務正在執行中..."}
        
    background_tasks.add_task(run_background_job_audit, job_id)
    return {"status": "started", "job_id": job_id, "message": "已啟動任務專屬審算作業"}

@app.get("/api/audit/status")
async def get_audit_status(job_id: str = Query(...)):
    job_info = JOBS_REGISTRY.get(job_id, {"is_running": False, "progress": 0, "total": 0, "completed": 0})
    return job_info

@app.get("/api/audit/results")
async def get_audit_results(job_id: Optional[str] = Query(None)):
    if not job_id or job_id not in JOBS_REGISTRY:
        return JSONResponse({"status": "empty", "message": "當前無選定的審計任務，請上傳檔案開始審核", "records": [], "stats": {}})
        
    job_info = JOBS_REGISTRY[job_id]
    if "records" not in job_info:
        # 試圖讀取 job 目錄下最新產出的 excel
        job_path = get_job_dir(job_id)
        reports = list(job_path.glob("ICP_Audit_Report_*.xlsx"))
        if reports:
            reports.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            latest_file = reports[0]
            records = load_excel_records(latest_file)
            mod_time = datetime.fromtimestamp(latest_file.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            job_info["records"] = records
            job_info["stats"] = calculate_stats(records, latest_file.name, mod_time)
        else:
            return JSONResponse({"status": "empty", "message": "該任務尚無審核結果，請上傳檔案並點擊「啟動全自動 LLM 合規審核」", "records": [], "stats": {}})
            
    return {
        "status": "success",
        "job_id": job_id,
        "stats": job_info.get("stats", {}),
        "records": job_info.get("records", [])
    }

@app.get("/api/reports/download")
async def download_report(job_id: str = Query(...)):
    job_path = get_job_dir(job_id)
    reports = list(job_path.glob("ICP_Audit_Report_*.xlsx"))
    if not reports:
        raise HTTPException(status_code=404, detail="找不到該任務的審計報表檔案")
    reports.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    file_path = reports[0]
    return FileResponse(path=file_path, filename=file_path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_server:app", host="127.0.0.1", port=8000, reload=True)
