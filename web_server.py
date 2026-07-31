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

def generate_html_report(records: List[Dict[str, Any]], stats: Dict[str, Any], output_path: Path):
    """生成單一獨立、可離線開啟的 HTML 視覺化合規審計報告"""
    cards_html = ""
    for idx, rec in enumerate(records, 1):
        level = str(rec.get("LLM研判等級", "")).strip()
        badge_class = "badge-high" if level == "High" else "badge-medium" if level == "Medium" else "badge-low" if level == "Low" else "badge-fp"
        badge_label = "🔴 High (同一實體/轉運風險)" if level == "High" else "🟠 Medium (關聯企業)" if level == "Medium" else "🟡 Low (疑慮)" if level == "Low" else "🟢 False Positive (確定誤判)"
        
        cards_html += f"""
        <div class="card">
            <div class="card-header">
                <div class="case-title">#{idx:02d} {rec.get('查詢名稱', '—')} ⇄ {rec.get('黑名單名稱', '—')}</div>
                <div>
                    <span class="pct">{rec.get('原XML命中率', '—')}</span>
                    <span class="badge {badge_class}">{badge_label}</span>
                </div>
            </div>
            <div class="grid">
                <div class="box">
                    <h4>查詢實體 (Condition)</h4>
                    <p><strong>條件ID:</strong> {rec.get('條件ID', '—')}</p>
                    <p><strong>名稱:</strong> {rec.get('查詢名稱', '—')}</p>
                    <p><strong>國家/城市:</strong> {rec.get('查詢國家', '—')} / {rec.get('查詢城市', '—')}</p>
                    <p><strong>地址:</strong> {rec.get('查詢地址', '—')}</p>
                </div>
                <div class="box">
                    <h4>黑名單限制實體 (Party)</h4>
                    <p><strong>黑名單ID:</strong> {rec.get('黑名單ID', '—')}</p>
                    <p><strong>名稱:</strong> {rec.get('黑名單名稱', '—')}</p>
                    <p><strong>地址:</strong> {rec.get('黑名單地址', '—')}</p>
                    <p><strong>完整資訊:</strong> {rec.get('黑名單完整資訊', '—')}</p>
                </div>
            </div>
            {f'<div class="risk-tag">風險判定依據：{rec.get("風險判定依據")}</div>' if rec.get("風險判定依據") else ''}
            <div class="reasoning">
                <strong>LLM 專家模型 Chain-of-Thought (CoT) 推理分析：</strong>
                <p>{rec.get('LLM分析推理理由', '—')}</p>
            </div>
        </div>
        """
        
    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>ICP-Compare 獨立視覺化審計報告 ({stats.get('file_name', '')})</title>
    <style>
        body {{ background: #0B0F19; color: #F1F5F9; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 24px; line-height: 1.5; }}
        .header {{ background: rgba(30,41,59,0.7); border: 1px solid rgba(255,255,255,0.1); padding: 20px 28px; border-radius: 12px; margin-bottom: 24px; }}
        .header h1 {{ margin: 0; font-size: 1.5rem; color: #3B82F6; }}
        .kpis {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 24px; }}
        .kpi {{ background: rgba(20,27,44,0.65); border: 1px solid rgba(255,255,255,0.08); padding: 16px; border-radius: 12px; }}
        .kpi span {{ font-size: 0.8rem; color: #94A3B8; }}
        .kpi h2 {{ margin: 4px 0 0 0; font-size: 1.6rem; }}
        .card {{ background: rgba(20,27,44,0.65); border: 1px solid rgba(255,255,255,0.08); padding: 20px; border-radius: 12px; margin-bottom: 20px; }}
        .card-header {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 12px; margin-bottom: 16px; }}
        .case-title {{ font-size: 1.1rem; font-weight: 700; color: #FFF; }}
        .badge {{ padding: 4px 10px; border-radius: 12px; font-size: 0.75rem; font-weight: 700; text-transform: uppercase; }}
        .badge-high {{ background: rgba(239, 68, 68, 0.15); color: #EF4444; border: 1px solid #EF4444; }}
        .badge-medium {{ background: rgba(245, 158, 11, 0.15); color: #F59E0B; border: 1px solid #F59E0B; }}
        .badge-low {{ background: rgba(234, 179, 8, 0.15); color: #EAB308; }}
        .badge-fp {{ background: rgba(16, 185, 129, 0.15); color: #10B981; border: 1px solid #10B981; }}
        .pct {{ background: rgba(255,255,255,0.1); padding: 3px 8px; border-radius: 4px; font-size: 0.8rem; margin-right: 8px; }}
        .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 14px; }}
        .box {{ background: rgba(15,23,42,0.5); padding: 14px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.05); font-size: 0.85rem; }}
        .box h4 {{ margin: 0 0 10px 0; color: #14B8A6; font-size: 0.9rem; }}
        .box p {{ margin: 4px 0; }}
        .risk-tag {{ background: rgba(245,158,11,0.15); color: #F59E0B; border: 1px solid rgba(245,158,11,0.3); padding: 4px 10px; border-radius: 6px; font-size: 0.8rem; margin-bottom: 14px; display: inline-block; }}
        .reasoning {{ background: rgba(30,41,59,0.5); border-left: 4px solid #3B82F6; padding: 14px; border-radius: 6px; font-size: 0.85rem; }}
        .reasoning strong {{ color: #60A5FA; display: block; margin-bottom: 6px; }}
        @media print {{ body {{ background: #FFF; color: #000; }} .card {{ page-break-inside: avoid; border: 1px solid #CCC; color: #000; background: #FFF; }} .box {{ background: #F8FAFC; color: #000; }} .reasoning {{ background: #EFF6FF; color: #000; }} }}
    </style>
</head>
<body>
    <div class="header">
        <h1>🛡️ ICP-Compare 出口合規審計報告 (HTML 視覺化版)</h1>
        <p style="margin: 6px 0 0 0; color: #94A3B8; font-size: 0.85rem;">生成時間: {stats.get('last_updated', '')} | 報告檔名: {stats.get('file_name', '')}</p>
    </div>
    
    <div class="kpis">
        <div class="kpi"><span>總掃描對數</span><h2>{stats.get('total', 0)}</h2></div>
        <div class="kpi"><span>自動放行率</span><h2 style="color: #10B981;">{stats.get('auto_release_rate', 0)}%</h2></div>
        <div class="kpi"><span>🔴 High 高風險</span><h2 style="color: #EF4444;">{stats.get('high', 0)}</h2></div>
        <div class="kpi"><span>🟠 Medium 關聯企業</span><h2 style="color: #F59E0B;">{stats.get('medium', 0)}</h2></div>
        <div class="kpi"><span>🟢 False Positive</span><h2 style="color: #10B981;">{stats.get('fp', 0)}</h2></div>
    </div>
    
    <div class="cards">
        {cards_html}
    </div>
</body>
</html>"""
    output_path.write_text(html_content, encoding="utf-8")

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

@app.get("/api/reports/download_html")
async def download_html_report(job_id: str = Query(...)):
    job_path = get_job_dir(job_id)
    job_info = JOBS_REGISTRY.get(job_id, {})
    records = job_info.get("records", [])
    
    if not records:
        reports = list(job_path.glob("ICP_Audit_Report_*.xlsx"))
        if reports:
            reports.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            latest_excel = reports[0]
            records = load_excel_records(latest_excel)
            mod_time = datetime.fromtimestamp(latest_excel.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            stats = calculate_stats(records, latest_excel.name, mod_time)
        else:
            raise HTTPException(status_code=404, detail="找不到該任務的審計紀錄，無法生成 HTML 報表")
    else:
        stats = job_info.get("stats", {})
        
    html_filename = f"ICP_Audit_Report_{job_id}.html"
    html_path = job_path / html_filename
    generate_html_report(records, stats, html_path)
    
    return FileResponse(path=html_path, filename=html_filename, media_type="text/html")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("web_server:app", host="127.0.0.1", port=8000, reload=True)
