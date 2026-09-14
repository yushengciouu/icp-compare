import os
import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from audit_service import router as audit_router

app = FastAPI(
    title="ICP-Compare Audit API Server",
    description="戰略性高科技貨品出口實體管理名單 (RPS) 獨立合規審計 API 服務。支援單檔/多檔 XML 上傳，經雙 H200 本地 LLM 語意研判後，一鍵直接下載格式化 Excel 報表。",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# 允許跨域請求 (CORS)，方便外部前端系統或工具直接呼叫
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 掛載審計核心 API 路由
app.include_router(audit_router)

# 修復 FastAPI / Swagger UI 對多檔案上傳 (OpenAPI 3.1) 的相容性問題，確保顯示檔案選取按鈕
from fastapi.openapi.utils import get_openapi

def custom_openapi():
    if app.openapi_schema:
        return app.openapi_schema
    schema = get_openapi(
        title=app.title,
        version=app.version,
        description=app.description,
        routes=app.routes,
    )
    for s in schema.get("components", {}).get("schemas", {}).values():
        for p in s.get("properties", {}).values():
            if p.get("type") == "array" and p.get("items", {}).get("contentMediaType") == "application/octet-stream":
                p["items"] = {"type": "string", "format": "binary"}
            elif p.get("contentMediaType") == "application/octet-stream":
                p["format"] = "binary"
    app.openapi_schema = schema
    return app.openapi_schema

app.openapi = custom_openapi

@app.get("/", tags=["系統健康檢查"])
async def root():
    return JSONResponse({
        "service": "ICP-Compare Audit API Server",
        "status": "healthy",
        "endpoints": {
            "docs": "/docs",
            "export_api": "/api/v1/audit/export"
        },
        "description": "請至 /docs 查看 Swagger 互動文件，或直接發送 POST /api/v1/audit/export 進行審計直出"
    })

@app.get("/health", tags=["系統健康檢查"])
async def health():
    return {"status": "ok"}

if __name__ == "__main__":
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", 8001))
    reload = os.getenv("RELOAD", "false").lower() == "true"
    print(f"🚀 獨立 API 伺服器啟動中: http://{host}:{port}")
    print(f"📖 互動式 Swagger API 文件位址: http://localhost:{port}/docs")
    uvicorn.run("api_server:app", host=host, port=port, reload=reload)
