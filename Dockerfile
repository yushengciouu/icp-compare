FROM python:3.11-slim

WORKDIR /app

# 安裝系統基本依賴
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 從官方鏡像複製 uv
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uv/bin/uv

# 複製依賴設定並進行高速安裝
COPY pyproject.toml uv.lock /app/
RUN /uv/bin/uv sync --frozen --no-cache

# 複製專案主要程式碼
COPY . /app/

EXPOSE 8000

ENV HOST=0.0.0.0
ENV PORT=8000

CMD ["/uv/bin/uv", "run", "python", "web_server.py"]
