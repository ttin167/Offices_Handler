# ==============================================================================
# Dockerfile — Antigravity Document Handler Enterprise Production Container (v4.0)
# Multi-engine Document Processing, Native Charts, OCR, and Technical Diagrams
# ==============================================================================

FROM python:3.11-slim-bookworm

# Configure environment
ENV PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    LANG=C.UTF-8 \
    LC_ALL=C.UTF-8 \
    PLANTUML_LIMIT_SIZE=16384 \
    PORT=8000

# Install system dependencies:
# - LibreOffice: Zero-COM PDF conversion backend
# - Tesseract OCR + Vietnamese & English models
# - Chromium & driver: Headless canvas diagram rendering
# - OpenJDK JRE: PlantUML diagram rendering
# - Core TrueType / OpenType fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    tesseract-ocr \
    tesseract-ocr-vie \
    tesseract-ocr-eng \
    default-jre-headless \
    chromium \
    fonts-liberation \
    fonts-dejavu-core \
    fontconfig \
    curl \
    git \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python requirements
COPY requirements.txt pyproject.toml /app/
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt && \
    pip install --no-cache-dir -e .

# Copy full application codebase
COPY . /app/

# Create runtime directories
RUN mkdir -p /app/diagram_assets /app/mail_merge_output /app/logs

# Healthcheck to verify Web Studio API
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:8000/api/doctor || exit 1

EXPOSE 8000

# Start Web Studio Server with SSE event streaming
CMD ["python", "web_api_server.py", "8000"]
