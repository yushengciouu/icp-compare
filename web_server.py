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
REPORTS_DIR = BASE_DIR / "reports"

STATIC_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
TESTFILE_DIR.mkdir(exist_ok=True)
REPORTS_DIR.mkdir(exist_ok=True)

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
    files = list(REPORTS_DIR.glob("ICP_Audit_Report_*.xlsx")) + list(BASE_DIR.glob("ICP_Audit_Report_*.xlsx"))
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
    files = list(REPORTS_DIR.glob("ICP_Audit_Report_*.xlsx")) + list(BASE_DIR.glob("ICP_Audit_Report_*.xlsx"))
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
    file_path = REPORTS_DIR / filename
    if not file_path.exists():
        file_path = BASE_DIR / filename
    if not file_path.exists() or not file_path.name.endswith(".xlsx"):
        raise HTTPException(status_code=404, detail="檔案不存在")
    return FileResponse(path=file_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.post("/api/upload")
async def upload_xml_files(files: List[UploadFile] = File(...)):
    # 每次上傳新檔案時，先自動清空舊的 XML 檔案
    for old_file in TESTFILE_DIR.glob("*.xml"):
        try:
            old_file.unlink()
        except Exception as e:
            print(f"無法刪除舊檔案 {old_file.name}: {e}")

    saved_files = []
    for file in files:
        if not file.filename.endswith(".xml"):
            continue
        target_path = TESTFILE_DIR / file.filename
        content = await file.read()
        target_path.write_bytes(content)
        saved_files.append(file.filename)
        
    return {"status": "success", "uploaded": saved_files, "message": f"已自動清空舊檔並成功上傳 {len(saved_files)} 個 XML 檔案"}

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
    output_excel = REPORTS_DIR / f"ICP_Audit_Report_{timestamp}.xlsx"
    
    df = pd.DataFrame(results)
    df.to_excel(output_excel, sheet_name="全審計結果清單", index=False)
    generate_formatted_excel_report(str(output_excel))
    generate_standalone_html(output_excel)
    
    AUDIT_STATUS["is_running"] = False
    AUDIT_STATUS["last_completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    AUDIT_STATUS["last_report"] = output_excel.name

def generate_standalone_html(excel_path: Path) -> Path:
    html_output_path = excel_path.with_suffix(".html")
    records = load_excel_records(excel_path)
    
    total = len(records)
    high_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "High")
    medium_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "Medium")
    low_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "Low")
    fp_count = sum(1 for r in records if str(r.get("LLM研判等級", "")).strip() == "False Positive")
    auto_release_rate = round((fp_count / total * 100), 1) if total > 0 else 0.0

    css_content = (STATIC_DIR / "style.css").read_text(encoding="utf-8") if (STATIC_DIR / "style.css").exists() else ""
    
    records_json = json.dumps(records, ensure_ascii=False)
    
    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ICP-Compare 戰略出口合規審計報告 ({excel_path.stem})</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
{css_content}
    </style>
</head>
<body class="dark-theme">
    <div class="bg-glow blob-1"></div>
    <div class="bg-glow blob-2"></div>
    <div class="bg-glow blob-3"></div>

    <div class="app-container">
        <header class="navbar glass-card">
            <div class="brand">
                <div class="logo-icon"><i data-lucide="shield-check"></i></div>
                <div class="brand-text">
                    <h1>ICP-Compare <span>Interactive Report</span></h1>
                    <p>戰略性高科技貨品出口合規審計報告 (離線獨立版: {excel_path.name})</p>
                </div>
            </div>
        </header>

        <section class="kpi-grid">
            <div class="kpi-card glass-card border-total">
                <div class="kpi-icon total"><i data-lucide="file-search"></i></div>
                <div class="kpi-info">
                    <span class="kpi-label">總掃描對數</span>
                    <h2 class="kpi-value">{total}</h2>
                </div>
            </div>
            <div class="kpi-card glass-card border-release">
                <div class="kpi-icon release"><i data-lucide="zap"></i></div>
                <div class="kpi-info">
                    <span class="kpi-label">自動放行率 (誤判率)</span>
                    <h2 class="kpi-value highlight-green">{auto_release_rate}%</h2>
                </div>
            </div>
            <div class="kpi-card glass-card border-high {'has-warning-high' if high_count > 0 else ''}">
                <div class="kpi-icon high"><i data-lucide="alert-triangle"></i></div>
                <div class="kpi-info">
                    <span class="kpi-label">🔴 High 高風險 (攔截)</span>
                    <h2 class="kpi-value text-high">{high_count}</h2>
                </div>
            </div>
            <div class="kpi-card glass-card border-medium {'has-warning-medium' if medium_count > 0 else ''}">
                <div class="kpi-icon medium"><i data-lucide="git-fork"></i></div>
                <div class="kpi-info">
                    <span class="kpi-label">🟠 Medium 關聯企業 (二審)</span>
                    <h2 class="kpi-value text-medium">{medium_count}</h2>
                </div>
            </div>
            <div class="kpi-card glass-card border-fp">
                <div class="kpi-icon fp"><i data-lucide="check-circle-2"></i></div>
                <div class="kpi-info">
                    <span class="kpi-label">🟢 False Positive (確定誤判)</span>
                    <h2 class="kpi-value text-fp">{fp_count}</h2>
                </div>
            </div>
        </section>

        <section class="filter-section">
            <div class="tabs-group">
                <button class="tab-btn active" data-filter="ALL"><i data-lucide="layers"></i> 全部案件 ({total})</button>
                <button class="tab-btn tab-high" data-filter="High"><i data-lucide="shield-alert"></i> 🔴 High 高風險 ({high_count})</button>
                <button class="tab-btn tab-medium" data-filter="Medium"><i data-lucide="building"></i> 🟠 Medium 關聯企業 ({medium_count})</button>
                <button class="tab-btn tab-low" data-filter="Low"><i data-lucide="help-circle"></i> 🟡 Low 疑慮案件 ({low_count})</button>
                <button class="tab-btn tab-fp" data-filter="False Positive"><i data-lucide="shield-check"></i> 🟢 確定誤判放行 ({fp_count})</button>
            </div>
            <div class="search-box">
                <i data-lucide="search" class="search-icon"></i>
                <input type="text" id="search-input" placeholder="搜尋實體名稱、條件 ID、地址或國家...">
            </div>
        </section>

        <main class="cards-grid" id="cards-container"></main>
    </div>

    <script>
        const allRecords = {records_json};
        let currentFilter = "ALL";
        let currentSearch = "";

        document.addEventListener("DOMContentLoaded", () => {{
            lucide.createIcons();
            renderCards();
            setupEvents();
        }});

        function renderCards() {{
            const container = document.getElementById("cards-container");
            container.innerHTML = "";

            const filtered = allRecords.filter(rec => {{
                const level = String(rec["LLM研判等級"] || "").trim();
                const matchesFilter = (currentFilter === "ALL") || (level === currentFilter);
                const searchText = (
                    String(rec["查詢名稱"] || "") +
                    String(rec["黑名單名稱"] || "") +
                    String(rec["條件ID"] || "") +
                    String(rec["黑名單ID"] || "") +
                    String(rec["查詢地址"] || "") +
                    String(rec["黑名單地址"] || "")
                ).toLowerCase();
                const matchesSearch = !currentSearch || searchText.includes(currentSearch.toLowerCase());
                return matchesFilter && matchesSearch;
            }});

            if (filtered.length === 0) {{
                container.innerHTML = `<div class="loading-state"><i data-lucide="inbox" style="width: 48px; height: 48px; opacity: 0.5;"></i><p style="margin-top: 12px;">無符合條件的合規案件對照紀錄</p></div>`;
                lucide.createIcons();
                return;
            }}

            filtered.forEach((rec, idx) => {{
                const card = createCardElement(rec, idx + 1);
                container.appendChild(card);
            }});
            lucide.createIcons();
        }}

        function getBadgeClass(level) {{
            switch (level) {{
                case "High": return "badge-high";
                case "Medium": return "badge-medium";
                case "Low": return "badge-low";
                case "False Positive": return "badge-fp";
                default: return "badge-low";
            }}
        }}

        function getBadgeLabel(level) {{
            switch (level) {{
                case "High": return "🔴 High (同一實體/轉運風險)";
                case "Medium": return "🟠 Medium (關聯企業)";
                case "Low": return "🟡 Low (疑慮)";
                case "False Positive": return "🟢 False Positive (確定誤判)";
                default: return level;
            }}
        }}

        function createCardElement(rec, index) {{
            const card = document.createElement("article");
            card.className = "audit-card glass-card";
            const level = String(rec["LLM研判等級"] || "").trim();
            const queryName = rec["查詢名稱"] || "—";
            const watchName = rec["黑名單名稱"] || "—";
            const conditionId = rec["條件ID"] || "—";
            const partyId = rec["黑名單ID"] || "—";
            const pct = rec["原XML命中率"] || "—";
            const riskFactor = rec["風險判定依據"] || "";
            const reasoning = rec["LLM分析推理理由"] || "尚無推理資料";

            card.innerHTML = `
                <div class="card-header-bar">
                    <div class="case-title">
                        <span class="case-num">#${{String(index).padStart(2, '0')}}</span>
                        <span class="entity-pair-name">${{escapeHtml(queryName)}} <span>⇄</span> ${{escapeHtml(watchName)}}</span>
                    </div>
                    <div class="badges-group">
                        <span class="pct-badge"><i data-lucide="percent" style="width: 12px; display:inline;"></i> ${{pct}}</span>
                        <span class="badge ${{getBadgeClass(level)}}">${{getBadgeLabel(level)}}</span>
                    </div>
                </div>

                <div class="comparison-container">
                    <div class="entity-box">
                        <h4><i data-lucide="user-search"></i> 查詢實體 (Condition)</h4>
                        <div class="info-row"><strong>條件ID:</strong> ${{escapeHtml(conditionId)}}</div>
                        <div class="info-row"><strong>名稱:</strong> ${{escapeHtml(queryName)}}</div>
                        <div class="info-row"><strong>國家/城市:</strong> ${{escapeHtml(rec["查詢國家"] || "—")}} / ${{escapeHtml(rec["查詢城市"] || "—")}}</div>
                        <div class="info-row"><strong>地址:</strong> ${{escapeHtml(rec["查詢地址"] || "—")}}</div>
                    </div>
                    <div class="entity-box">
                        <h4><i data-lucide="shield-alert"></i> 黑名單限制實體 (Party)</h4>
                        <div class="info-row"><strong>黑名單ID:</strong> ${{escapeHtml(partyId)}}</div>
                        <div class="info-row"><strong>黑名單名稱:</strong> ${{escapeHtml(watchName)}}</div>
                        <div class="info-row"><strong>地址:</strong> ${{escapeHtml(rec["黑名單地址"] || "—")}}</div>
                        <div class="info-row"><strong>完整細節:</strong> ${{escapeHtml(rec["黑名單完整資訊"] || "—")}}</div>
                    </div>
                </div>

                ${{riskFactor ? `<div class="risk-factor-tag"><i data-lucide="target" style="width: 12px; display:inline;"></i> 風險判定依據：${{escapeHtml(riskFactor)}}</div>` : ''}}

                <div class="reasoning-box">
                    <div class="reasoning-header">
                        <i data-lucide="brain-circuit"></i> 專家模型審計推理理由：
                    </div>
                    <div class="reasoning-body">${{escapeHtml(reasoning)}}</div>
                </div>

                <div class="card-actions">
                    ${{level === 'High' ? `<div class="suggestion-tag suggestion-high"><i data-lucide="slash" style="width:15px;height:15px;"></i><span>建議處置：<strong>攔截凍結交易</strong></span></div>` : ''}}
                    ${{level === 'Medium' ? `<div class="suggestion-tag suggestion-medium"><i data-lucide="file-text" style="width:15px;height:15px;"></i><span>建議處置：<strong>人工二審</strong></span></div>` : ''}}
                    ${{(level === 'False Positive' || level === 'Low') ? `<div class="suggestion-tag suggestion-pass"><i data-lucide="check-circle" style="width:15px;height:15px;"></i><span>建議處置：<strong>自動放行</strong></span></div>` : ''}}
                </div>
            `;
            return card;
        }}

        function setupEvents() {{
            document.querySelectorAll(".tab-btn").forEach(btn => {{
                btn.addEventListener("click", () => {{
                    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
                    btn.classList.add("active");
                    currentFilter = btn.getAttribute("data-filter");
                    renderCards();
                }});
            }});

            const searchInput = document.getElementById("search-input");
            searchInput.addEventListener("input", (e) => {{
                currentSearch = e.target.value.trim();
                renderCards();
            }});
        }}

        function escapeHtml(text) {{
            if (!text) return "";
            return String(text)
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#039;");
        }}
    </script>
</body>
</html>
"""
    html_output_path.write_text(html_content, encoding="utf-8")
    return html_output_path

@app.get("/api/reports/download_html/{filename}")
async def download_html_report(filename: str):
    base_name = filename.rsplit(".", 1)[0]
    excel_file = REPORTS_DIR / f"{base_name}.xlsx"
    if not excel_file.exists():
        excel_file = BASE_DIR / f"{base_name}.xlsx"
    if not excel_file.exists():
        raise HTTPException(status_code=404, detail="對應的 Excel 審計報表不存在")
    
    html_file = generate_standalone_html(excel_file)
    return FileResponse(path=html_file, filename=html_file.name, media_type="text/html")

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
