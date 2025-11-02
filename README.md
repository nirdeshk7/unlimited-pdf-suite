# UnlimitedPDFSuite — Render Deploy (GitHub-Ready)

Self-hosted PDF toolkit (FastAPI + HTML/JS) with: Convert, Merge, Split, Rotate, N‑up, Booklet, OCR, Compress, Password Set/Unlock, Bookmarks, **PDF → DOCX (Beta)**.

## Deploy on Render (from GitHub)
1. Push this repo to GitHub.
2. On Render: **New → Blueprint** → select this repo (it reads `render.yaml`).
3. Wait for build & deploy. Open your service URL.

### Local test
```bash
cd backend
python3 -m venv .venv && source .venv/bin/activate
pip install --upgrade pip && pip install -r requirements.txt
uvicorn app:app --reload --port 8000
# Open http://localhost:8000
```

### Notes
- Large files: Render may have platform limits; for very large uploads, consider a reverse proxy with Nginx.
- OCR languages: install extra packs in Dockerfile if needed (e.g. `tesseract-ocr-hin`).
- Persistence: For keeping outputs, attach a Render Disk to `/app/backend/output`.
- Security: Unlock works only with correct password; no bypass/crack.
