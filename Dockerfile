# Dockerfile for Render
FROM python:3.11-slim

# System dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice tesseract-ocr ghostscript poppler-utils curl \
    && rm -rf /var/lib/apt/lists/*

# Optional: add more Tesseract languages (uncomment if needed)
# RUN apt-get update && apt-get install -y --no-install-recommends tesseract-ocr-hin tesseract-ocr-mar

WORKDIR /app
COPY . /app

# Python deps
WORKDIR /app/backend
RUN python -m venv /opt/venv && . /opt/venv/bin/activate && \
    pip install --upgrade pip && pip install -r requirements.txt

ENV PATH="/opt/venv/bin:$PATH"

# Render provides $PORT
CMD ["bash", "-lc", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
