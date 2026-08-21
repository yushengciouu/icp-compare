import os
import glob
import re
import json
import uuid
import shutil
import time
import zipfile
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime
from contextlib import asynccontextmanager

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Query
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook
import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler

# 匯入現有的審計核心模組
from compare_audit import parse_xml_file, get_llm_judgment, generate_formatted_excel_report

BASE_DIR = Path(__file__).parent.resolve()
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"
JOBS_DIR = BASE_DIR / "jobs"

STATIC_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
JOBS_DIR.mkdir(exist_ok=True)

def clean_expired_jobs(days: int = 7):
    """自動清理在 JOBS_DIR 中最後修改時間超過指定天數 (預設 7 天) 的舊任務資料夾"""
    if not JOBS_DIR.exists():
        return

    now = time.time()
    cutoff_seconds = days * 86400  # 1 天 = 86400 秒
    deleted_count = 0

    for job_folder in JOBS_DIR.iterdir():
        if job_folder.is_dir():
            try:
                mtime = job_folder.stat().st_mtime
                if (now - mtime) > cutoff_seconds:
                    shutil.rmtree(job_folder)
                    deleted_count += 1
                    print(f"[Auto Cleanup] 已自動刪除過期任務目錄: {job_folder.name}")
            except Exception as e:
                print(f"[Auto Cleanup Error] 刪除目錄 {job_folder.name} 失敗: {e}")

    if deleted_count > 0:
        print(f"[Auto Cleanup] 共自動清理了 {deleted_count} 個超過 {days} 天的舊任務目錄。")

@asynccontextmanager
async def lifespan(app: FastAPI):
    # 服務啟動時：開啟每週背景排程器
    scheduler = BackgroundScheduler()
    # 每週日凌晨 03:00 執行一次，清理 7 天前舊檔
    scheduler.add_job(clean_expired_jobs, 'cron', day_of_week='sun', hour=3, minute=0, args=[7])
    scheduler.start()
    print("[Scheduler] 每週檔案自動清理排程器已成功啟動 (每週日 03:00 清理過期任務檔案)")

    yield

    # 服務關閉時：停止排程器
    scheduler.shutdown()
    print("[Scheduler] 自動清理排程器已正常停止")

app = FastAPI(title="ICP-Compare Web 戰略出口合規審計平台 (Multi-User Job Isolation)", lifespan=lifespan)
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
    return {
        "total": total,
        "high": high_count,
        "medium": medium_count,
        "low": low_count,
        "file_name": file_name,
        "last_updated": mod_time
    }

def generate_html_report(records: List[Dict[str, Any]], stats: Dict[str, Any], output_path: Path):
    """生成單一獨立、可離線開啟且支援即時篩選與搜尋的 HTML 視覺化合規審計報告"""
    if not records:
        cards_html = """
        <div class="card" style="text-align: center; padding: 48px 20px;">
            <div style="font-size: 2.5rem; margin-bottom: 12px;">✅</div>
            <h3 style="color: #10B981; margin: 0 0 8px 0; font-size: 1.2rem;">本批次 XML 報文全數合格</h3>
            <p style="color: #94A3B8; font-size: 0.88rem; margin: 0;">經比對未檢出任何命中率 >= 75% 之限制黑名單疑慮實體。</p>
        </div>
        """
    else:
        cards_html = ""
        for idx, rec in enumerate(records, 1):
            level = str(rec.get("LLM研判等級", "")).strip()
            badge_class = "badge-high" if level == "High" else "badge-medium" if level == "Medium" else "badge-low"
            badge_label = "🔴 High (同一實體/轉運風險)" if level == "High" else "🟠 Medium (關聯企業)" if level == "Medium" else "🟡 Low (低風險/可放行)"
            
            search_text = (
                str(rec.get("查詢名稱", "")) +
                str(rec.get("黑名單名稱", "")) +
                str(rec.get("條件ID", "")) +
                str(rec.get("黑名單ID", "")) +
                str(rec.get("查詢地址", "")) +
                str(rec.get("黑名單地址", "")) +
                str(rec.get("查詢國家", ""))
            ).lower().replace('"', '&quot;')

            cards_html += f"""
            <div class="card" data-level="{level}" data-search="{search_text}">
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
                <div class="risk-tag">公司名稱：{rec.get("公司名稱比對", "—")} ｜ 地址比對：{rec.get("地址比對", "—")}</div>
                <div class="reasoning">
                    <strong>專家模型審計推理理由：</strong>
                    <p>{rec.get('LLM分析推理理由', '—')}</p>
                </div>
            </div>
            """
        
    high_count = int(stats.get('high', 0))
    medium_count = int(stats.get('medium', 0))
    high_kpi_class = "kpi kpi-warning-high" if high_count > 0 else "kpi"
    medium_kpi_class = "kpi kpi-warning-medium" if medium_count > 0 else "kpi"

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
        .kpi {{ background: rgba(20,27,44,0.65); border: 1px solid rgba(255,255,255,0.08); padding: 16px; border-radius: 12px; transition: all 0.3s; }}
        .kpi span {{ font-size: 0.8rem; color: #94A3B8; }}
        .kpi h2 {{ margin: 4px 0 0 0; font-size: 1.6rem; }}
        
        /* 🔴 High 非 0 時：強烈警示脈衝發光 (Pulsing Glow) */
        .kpi.kpi-warning-high {{
            border: 1px solid rgba(239, 68, 68, 0.85) !important;
            background: linear-gradient(135deg, rgba(239, 68, 68, 0.28) 0%, rgba(225, 29, 72, 0.12) 100%) !important;
            animation: warningPulseRed 2.4s infinite ease-in-out;
        }}
        /* 🟠 Medium 非 0 時：二審警示脈衝發光 (Pulsing Glow) */
        .kpi.kpi-warning-medium {{
            border: 1px solid rgba(245, 158, 11, 0.85) !important;
            background: linear-gradient(135deg, rgba(245, 158, 11, 0.28) 0%, rgba(217, 119, 6, 0.12) 100%) !important;
            animation: warningPulseAmber 2.4s infinite ease-in-out;
        }}
        @keyframes warningPulseRed {{
            0%, 100% {{
                box-shadow: 0 0 15px rgba(239, 68, 68, 0.35), inset 0 0 10px rgba(239, 68, 68, 0.15);
                border-color: rgba(248, 113, 113, 0.6);
            }}
            50% {{
                box-shadow: 0 0 35px rgba(239, 68, 68, 0.85), inset 0 0 22px rgba(239, 68, 68, 0.4);
                border-color: rgba(239, 68, 68, 1);
            }}
        }}
        @keyframes warningPulseAmber {{
            0%, 100% {{
                box-shadow: 0 0 15px rgba(245, 158, 11, 0.35), inset 0 0 10px rgba(245, 158, 11, 0.15);
                border-color: rgba(251, 191, 36, 0.6);
            }}
            50% {{
                box-shadow: 0 0 35px rgba(245, 158, 11, 0.85), inset 0 0 22px rgba(245, 158, 11, 0.4);
                border-color: rgba(245, 158, 11, 1);
            }}
        }}

        /* 篩選與搜尋列區塊 */
        .filter-section {{ display: flex; justify-content: space-between; align-items: center; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; background: rgba(20,27,44,0.65); border: 1px solid rgba(255,255,255,0.08); padding: 14px 20px; border-radius: 12px; }}
        .tabs-group {{ display: flex; gap: 8px; flex-wrap: wrap; }}
        .tab-btn {{ background: rgba(30,41,59,0.6); color: #94A3B8; border: 1px solid rgba(255,255,255,0.1); padding: 8px 16px; border-radius: 8px; cursor: pointer; font-size: 0.85rem; font-weight: 500; transition: all 0.2s; }}
        .tab-btn:hover {{ background: rgba(51,65,85,0.8); color: #FFF; }}
        .tab-btn.active {{ background: #3B82F6; color: #FFF; border-color: #60A5FA; font-weight: 600; box-shadow: 0 0 12px rgba(59,130,246,0.4); }}
        .tab-btn.tab-high.active {{ background: #EF4444; border-color: #F87171; box-shadow: 0 0 12px rgba(239,68,68,0.4); }}
        .tab-btn.tab-medium.active {{ background: #F59E0B; border-color: #FBBF24; box-shadow: 0 0 12px rgba(245,158,11,0.4); }}
        .tab-btn.tab-low.active {{ background: #EAB308; border-color: #FDE047; box-shadow: 0 0 12px rgba(234,179,8,0.4); }}
        .tab-btn.tab-fp.active {{ background: #10B981; border-color: #34D399; box-shadow: 0 0 12px rgba(16,185,129,0.4); }}
        .search-box input {{ background: rgba(15,23,42,0.8); border: 1px solid rgba(255,255,255,0.15); color: #FFF; padding: 8px 16px; border-radius: 8px; font-size: 0.85rem; width: 280px; outline: none; transition: all 0.2s; }}
        .search-box input:focus {{ border-color: #3B82F6; box-shadow: 0 0 8px rgba(59,130,246,0.3); }}
        
        .card {{ background: rgba(20,27,44,0.65); border: 1px solid rgba(255,255,255,0.08); padding: 20px; border-radius: 12px; margin-bottom: 20px; transition: all 0.2s; }}
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
        @media print {{ body {{ background: #FFF; color: #000; }} .filter-section {{ display: none; }} .card {{ page-break-inside: avoid; border: 1px solid #CCC; color: #000; background: #FFF; display: block !important; }} .box {{ background: #F8FAFC; color: #000; }} .reasoning {{ background: #EFF6FF; color: #000; }} }}
    </style>
</head>
<body>
    <div class="header">
        <h1>🛡️ ICP-Compare 出口合規審計報告 (HTML 視覺化獨立版)</h1>
        <p style="margin: 6px 0 0 0; color: #94A3B8; font-size: 0.85rem;">生成時間: {stats.get('last_updated', '')} | 報告檔名: {stats.get('file_name', '')}</p>
    </div>
    
    <div class="kpis">
        <div class="kpi"><span>總掃描對數</span><h2>{stats.get('total', 0)}</h2></div>
        <div class="{high_kpi_class}"><span>🔴 High 高風險</span><h2 style="color: #EF4444;">{high_count}</h2></div>
        <div class="{medium_kpi_class}"><span>🟠 Medium 關聯企業</span><h2 style="color: #F59E0B;">{medium_count}</h2></div>
        <div class="kpi"><span>🟡 Low 低風險/可放行</span><h2 style="color: #10B981;">{stats.get('low', 0)}</h2></div>
    </div>
    
    <section class="filter-section">
        <div class="tabs-group">
            <button class="tab-btn active" onclick="selectTab(this, 'ALL')">全部案件 ({stats.get('total', 0)})</button>
            <button class="tab-btn tab-high" onclick="selectTab(this, 'High')">🔴 High 高風險 ({stats.get('high', 0)})</button>
            <button class="tab-btn tab-medium" onclick="selectTab(this, 'Medium')">🟠 Medium 關聯企業 ({stats.get('medium', 0)})</button>
            <button class="tab-btn tab-low" onclick="selectTab(this, 'Low')">🟡 Low 低風險/可放行 ({stats.get('low', 0)})</button>
        </div>
        <div class="search-box">
            <input type="text" id="search-input" placeholder="搜尋實體名稱、條件 ID、地址或國家..." oninput="filterCards()">
        </div>
    </section>

    <div class="cards" id="cards-container">
        {cards_html}
    </div>

    <script>
        let currentFilter = 'ALL';
        function selectTab(btn, filter) {{
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            currentFilter = filter;
            filterCards();
        }}
        function filterCards() {{
            const query = (document.getElementById('search-input').value || '').toLowerCase().trim();
            const cards = document.querySelectorAll('.card');
            cards.forEach(card => {{
                const level = card.getAttribute('data-level') || '';
                const text = card.getAttribute('data-search') || '';
                const matchesFilter = (currentFilter === 'ALL') || (level === currentFilter);
                const matchesSearch = !query || text.includes(query);
                card.style.display = (matchesFilter && matchesSearch) ? 'block' : 'none';
            }});
        }}
    </script>
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
    
    # 清空該 job 之前的上傳舊檔與舊產出報表
    for old_file in upload_dir.glob("*.xml"):
        try:
            old_file.unlink()
        except Exception:
            pass
            
    for old_report in job_path.glob("ICP_Audit_Report_*.*"):
        try:
            old_report.unlink()
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
        
    job_info["total"] = len(all_pairs)
    job_info["completed"] = 0
    
    results = []
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from compare_audit import MAX_WORKERS
    
    if all_pairs:
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
                    "LLM研判等級": judgment.get("match_level", "Low"),
                    "公司名稱比對": judgment.get("name_match", "不同"),
                    "地址比對": judgment.get("address_match", "完全不同"),
                    "LLM分析推理理由": judgment.get("reasoning", "")
                }
                results.append(row)
                job_info["completed"] += 1
                job_info["progress"] = int((job_info["completed"] / job_info["total"]) * 100)
    else:
        job_info["progress"] = 100

    # 1. 產出「全檔案合併總 Excel 報表」 (Consolidated Report)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    columns = [
        "來源檔案", "條件ID", "查詢名稱", "查詢國家", "查詢城市", "查詢地址",
        "黑名單ID", "黑名單名稱", "黑名單地址", "黑名單完整資訊",
        "原XML命中率", "LLM研判等級", "公司名稱比對", "地址比對", "LLM分析推理理由"
    ]
    
    merged_report_name = f"ICP_Audit_Report_Consolidated_{timestamp}.xlsx"
    merged_excel_path = job_path / merged_report_name
    
    if results:
        df_merged = pd.DataFrame(results)
        df_merged.to_excel(merged_excel_path, sheet_name="全審計結果清單", index=False)
        generate_formatted_excel_report(str(merged_excel_path))
    else:
        source_names = ", ".join(os.path.basename(f) for f in xml_files) if xml_files else "—"
        pass_row = {col: "—" for col in columns}
        pass_row["來源檔案"] = source_names
        pass_row["LLM研判等級"] = "Pass (合格)"
        pass_row["LLM分析推理理由"] = "全數合格：本批次上傳之 XML 報文無任何命中率 >= 75% 之限制實體紀錄。"
        df_merged = pd.DataFrame([pass_row])
        df_merged.to_excel(merged_excel_path, sheet_name="全審計結果清單", index=False)

    # 2. 針對每一個上傳的 XML 檔案，個別產出專屬獨立 Excel 報表
    generated_individual_reports = []
    for xml_path_str in xml_files:
        xml_fname = os.path.basename(xml_path_str)
        stem = Path(xml_fname).stem.replace("_raw", "")
        file_report_name = f"ICP_Audit_Report_{stem}_{timestamp}.xlsx"
        output_excel = job_path / file_report_name
        
        file_results = [r for r in results if r.get("來源檔案") == xml_fname]
        if file_results:
            df = pd.DataFrame(file_results)
            df.to_excel(output_excel, sheet_name="全審計結果清單", index=False)
            generate_formatted_excel_report(str(output_excel))
        else:
            pass_row = {col: "—" for col in columns}
            pass_row["來源檔案"] = xml_fname
            pass_row["LLM研判等級"] = "Pass (合格)"
            pass_row["LLM分析推理理由"] = f"全數合格：本檔案 ({xml_fname}) 上傳之 XML 報文無任何命中率 >= 75% 之限制實體紀錄。"
            df = pd.DataFrame([pass_row])
            df.to_excel(output_excel, sheet_name="全審計結果清單", index=False)
            
        generated_individual_reports.append(file_report_name)
        
    # 3. 若上傳多個 XML 檔案，將個別報表打包為 ZIP 壓縮包
    if len(generated_individual_reports) > 1:
        zip_name = f"ICP_Audit_Reports_Individual_{timestamp}.zip"
        zip_path = job_path / zip_name
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for rep_name in generated_individual_reports:
                rep_file = job_path / rep_name
                if rep_file.exists():
                    zipf.write(rep_file, arcname=rep_name)
                    
    mod_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    job_info["is_running"] = False
    job_info["progress"] = 100
    job_info["last_completed_at"] = mod_time
    summary_filename = merged_report_name if len(xml_files) <= 1 else f"合併總表 + {len(generated_individual_reports)} 份個別報表"
    job_info["last_report"] = summary_filename
    job_info["records"] = results
    
    stats = calculate_stats(results, summary_filename, mod_time)
    stats["report_count"] = len(xml_files)
    stats["merged_report"] = merged_report_name
    stats["individual_reports"] = generated_individual_reports
    job_info["stats"] = stats

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
        reports = list(job_path.glob("ICP_Audit_Report_Consolidated_*.xlsx")) or list(job_path.glob("ICP_Audit_Report_*.xlsx"))
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
async def download_report(job_id: str = Query(...), mode: str = Query("merged")):
    job_path = get_job_dir(job_id)
    
    # 模式一：匯出各檔案個別獨立 Excel (ZIP 壓縮包)
    if mode in ["individual", "zip", "split"]:
        zip_files = list(job_path.glob("ICP_Audit_Reports_Individual_*.zip")) or list(job_path.glob("ICP_Audit_Reports_*.zip"))
        if zip_files:
            zip_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            return FileResponse(path=zip_files[0], filename=zip_files[0].name, media_type="application/zip")
            
        indiv_reports = [r for r in job_path.glob("ICP_Audit_Report_*.xlsx") if "Consolidated" not in r.name]
        if indiv_reports:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_path = job_path / f"ICP_Audit_Reports_Individual_{timestamp}.zip"
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for rep in indiv_reports:
                    zipf.write(rep, arcname=rep.name)
            return FileResponse(path=zip_path, filename=zip_path.name, media_type="application/zip")
            
    # 模式二：匯出全檔案合併總表 (Consolidated / Merged Excel)
    merged_reports = list(job_path.glob("ICP_Audit_Report_Consolidated_*.xlsx"))
    if merged_reports:
        merged_reports.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        file_path = merged_reports[0]
        return FileResponse(path=file_path, filename=file_path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    all_reports = list(job_path.glob("ICP_Audit_Report_*.xlsx"))
    if all_reports:
        all_reports.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        file_path = all_reports[0]
        return FileResponse(path=file_path, filename=file_path.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    raise HTTPException(status_code=404, detail="找不到該任務的審計報表檔案")

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
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", 8000))
    reload = os.getenv("RELOAD", "false").lower() == "true"
    uvicorn.run("web_server:app", host=host, port=port, reload=reload)
