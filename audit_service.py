import os
import re
import uuid
import zipfile
from pathlib import Path
from typing import List, Tuple, Dict, Any, Optional
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
from fastapi import APIRouter, UploadFile, File, Query, HTTPException
from fastapi.responses import FileResponse

from compare_audit import (
    MAX_WORKERS,
    group_xml_files,
    parse_batch_records,
    get_llm_judgment,
    generate_formatted_excel_report
)

COLUMNS = [
    "來源檔案", "條件ID", "客戶代號", "內部黑名單標記", "查詢名稱", "查詢國家", "查詢城市", "查詢地址",
    "黑名單ID", "黑名單名稱", "黑名單地址", "黑名單完整資訊",
    "原XML命中率", "LLM研判等級", "公司名稱比對", "地址比對", "LLM分析推理理由"
]

def run_audit_pipeline(
    xml_paths: List[str],
    output_dir: Path
) -> Tuple[Path, str, str]:
    """
    執行端到端合規審核管線：
    1. 解析並分組 XML 報文 (支援 Request/Response 配對)
    2. 提取疑似限制實體候選對象 (>=75%)
    3. 呼叫本地雙 H200 LLM 進行高併發語意研判
    4. 根據現有規則產出單一格式化 Excel 報表 (.xlsx)，不包成 ZIP 壓縮檔

    回傳: (檔案絕對路徑, 匯出檔案名稱, MIME類型)
    """
    if not xml_paths:
        raise ValueError("未提供任何 XML 檔案進行審核")

    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 1. 分組並解析 XML
    batches = group_xml_files(xml_paths)
    all_pairs = []
    for bkey, bdata in batches.items():
        pairs = parse_batch_records(bkey, bdata)
        all_pairs.extend(pairs)

    # 2. 進行 LLM 推論比對
    results = []
    if all_pairs:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            future_to_pair = {executor.submit(get_llm_judgment, pair): pair for pair in all_pairs}
            for future in as_completed(future_to_pair):
                pair = future_to_pair[future]
                judgment = future.result()
                row = {
                    "來源檔案": pair.get("source_file", ""),
                    "條件ID": pair.get("condition_id", ""),
                    "客戶代號": pair.get("customer_no", "—") if pair.get("customer_no") else "—",
                    "內部黑名單標記": pair.get("is_blacklisted", "—") if pair.get("is_blacklisted") else "—",
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

    # 3. 產出單一格式化 Excel 報表
    excel_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # 依現有命名規則：若單一業務批次則以該 batch_key 命名，否則以時間戳記命名
    if len(batches) == 1:
        batch_name = list(batches.keys())[0]
        report_filename = f"ICP_Audit_Report_{batch_name}_{timestamp}.xlsx"
    else:
        report_filename = f"ICP_Audit_Report_{timestamp}.xlsx"

    output_excel_path = output_dir / report_filename

    if results:
        df = pd.DataFrame(results)
        df.to_excel(output_excel_path, sheet_name="全審計結果清單", index=False)
        generate_formatted_excel_report(str(output_excel_path))
    else:
        source_names = ", ".join(batches.keys()) if batches else "—"
        pass_row = {col: "—" for col in COLUMNS}
        pass_row["來源檔案"] = source_names
        pass_row["LLM研判等級"] = "Pass (合格)"
        pass_row["LLM分析推理理由"] = "全數合格：本批次上傳之 XML 報文無任何命中率 >= 75% 之限制實體紀錄。"
        df = pd.DataFrame([pass_row])
        df.to_excel(output_excel_path, sheet_name="全審計結果清單", index=False)

    return output_excel_path, report_filename, excel_mime


# 建立獨立的 FastAPI 路由模組
router = APIRouter(prefix="/api/v1/audit", tags=["Audit API (一鍵合規審計直出)"])

BASE_JOBS_DIR = Path(__file__).parent.resolve() / "jobs"

@router.post(
    "/export",
    summary="上傳 XML 並直接下載合規審計 Excel 報表",
    description="外部自動化專用端點：支援上傳 XML 檔案（支援 Request 與 Response 配對），經本地 LLM 審核後直接輸出單一 Excel 報表 (.xlsx)。"
)
async def export_audit_report(
    files: List[UploadFile] = File(..., description="要上傳審查的一或多個 XML 檔案")
):
    # 建立獨立專屬暫存工作目錄
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_job_id = f"direct_{timestamp}_{uuid.uuid4().hex[:6]}"
    job_dir = BASE_JOBS_DIR / temp_job_id
    upload_dir = job_dir / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)

    xml_paths = []
    for file in files:
        if not file.filename.endswith(".xml"):
            continue
        save_path = upload_dir / file.filename
        content = await file.read()
        save_path.write_bytes(content)
        xml_paths.append(str(save_path))

    if not xml_paths:
        raise HTTPException(status_code=400, detail="請上傳至少一個副檔名為 .xml 的檔案")

    try:
        file_path, file_name, media_type = run_audit_pipeline(
            xml_paths=xml_paths,
            output_dir=job_dir
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"審計執行失敗: {str(e)}")

    if not file_path.exists():
        raise HTTPException(status_code=500, detail="報表生成失敗，產出檔案不存在")

    return FileResponse(
        path=file_path,
        filename=file_name,
        media_type=media_type
    )
