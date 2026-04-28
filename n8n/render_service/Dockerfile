FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8080

WORKDIR /app

# OS deps (python-docx needs nothing exotic; keep image slim)
RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Service + vendored pipeline code
COPY main.py ./
COPY render_proposal_package.py ./
COPY render_from_notion.py ./
COPY fee_memo.py ./
COPY scope_library.py ./

# Assets (contract template, comps cache)
COPY contract_template_merge_ready.docx ./
COPY comps_cache.json ./

EXPOSE 8080

# Cloud Run sets $PORT; bind uvicorn to it
CMD exec uvicorn main:app --host 0.0.0.0 --port ${PORT}
