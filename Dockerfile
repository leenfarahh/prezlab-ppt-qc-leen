# Prezlab PPT QC: Linux image for the Render demo/eval deployment.
# Renders slides via LibreOffice + poppler (no PowerPoint on Linux), persists
# to Supabase via DATABASE_URL. NOT for confidential client decks; the LAN
# box on Windows keeps client work (see DEPLOY.md).
FROM python:3.11-slim-bookworm

ENV PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1

# LibreOffice Impress (pptx -> pdf), poppler (pdf page -> png), and fonts:
# Liberation/DejaVu for Latin metric compatibility, Noto for Arabic shaping.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress \
        poppler-utils \
        fonts-liberation2 \
        fonts-dejavu-core \
        fonts-noto-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Only what the app needs at runtime (tests, fixtures, venv excluded via
# .dockerignore). qc/ imports spike/, so both ship.
COPY qc/ ./qc/
COPY spike/ ./spike/

EXPOSE 8000
# Render provides $PORT. Auth, renderer, DB, banner all come from env vars.
CMD ["sh", "-c", "uvicorn qc.web:app --host 0.0.0.0 --port ${PORT:-8000}"]
