# ==============================================================================
# Dockerfile — Antigravity Document Handler & BA Enterprise Workstation (v4.6.0)
# Multi-Stage Build: Node.js Frontend Compiler + Python 3.11 Full Office Engine
# Optimized for Cloud Deployment (Render.com, Railway, AWS ECS, Fly.io, Self-hosted)
# ==============================================================================

# ------------------------------------------------------------------------------
# STAGE 1: Compile High-Fidelity Vite Frontend Bundle
# ------------------------------------------------------------------------------
FROM node:18-alpine AS frontend-builder
WORKDIR /frontend

# Install frontend dependencies
COPY package.json package-lock.json* ./
RUN npm ci || npm install

# Build production bundle with optimized chunking and CSS minification
COPY index.html vite.config.js* ./
COPY src/ ./src/
RUN npm run build

# ------------------------------------------------------------------------------
# STAGE 2: Enterprise Python Document Processing & Diagram Execution Runtime
# ------------------------------------------------------------------------------
FROM python:3.11-slim-bookworm

# Configure environment:
# - Cap JVM memory to 192MB with -Xmx192m to prevent 512MB RAM OOM kills on Render Free tier
# - Set PLANTUML_LIMIT_SIZE and UTF-8 charset inside JAVA_TOOL_OPTIONS
# - Set MAX_JOB_WORKERS=1 to serialize memory-intensive headless jobs (LibreOffice / Chromium)
# - Set MAX_UPLOAD_SIZE_MB=50 to enforce upload quota guard
# - Bind HOST to 0.0.0.0 for container networking
ENV PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    LANG=C.UTF-8 \
    LC_ALL=C.UTF-8 \
    HOST=0.0.0.0 \
    PORT=10000 \
    MAX_JOB_WORKERS=1 \
    MAX_UPLOAD_SIZE_MB=50 \
    CHROMIUM_PATH=/usr/bin/chromium \
    JAVA_TOOL_OPTIONS="-Xmx192m -DPLANTUML_LIMIT_SIZE=16384 -Dfile.encoding=UTF-8"

# Install system dependencies:
# - LibreOffice: Zero-COM PDF/DOCX/XLSX/PPTX headless conversion engine
# - Tesseract OCR + Vietnamese & English models for scan recovery
# - Chromium: Headless browser for 300+ DPI Mermaid, Draw.io mxGraph, Canvas diagram rendering
# - Graphviz (dot): Mandatory binary for PlantUML Class, C4, Use Case, Activity rendering
# - OpenJDK JRE 17 headless: PlantUML diagram compilation backend
# - Liberation, DejaVu, and Google Noto fonts: Full typography & Vietnamese diacritics support
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    tesseract-ocr \
    tesseract-ocr-vie \
    tesseract-ocr-eng \
    openjdk-17-jre-headless \
    graphviz \
    chromium \
    fonts-liberation \
    fonts-dejavu-core \
    fonts-noto-core \
    fonts-noto-extra \
    fontconfig \
    curl \
    git \
    && fc-cache -f \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python third-party dependencies first for Docker layer caching efficiency
COPY requirements.txt /app/
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt

# Copy application codebase and install package in editable mode
# Note: BA Enterprise Workstation & Academy (core/ba_hub, specs/ba_*, etc.) is 100% Local-First
# and strictly excluded from Docker / Render Cloud builds via .dockerignore and explicit pruning below.
COPY . /app/
RUN rm -rf /app/core/ba_hub \
           /app/cli/ba_commands.py \
           /app/cli/ba_print_commands.py \
           /app/server/routes/ba_academy_routes.py \
           /app/src/modules/ba_academy_workstation.js \
           /app/specs/ba_*.json \
           /app/labor_law_vault \
           /app/ba_academy_vault \
           /app/data/labor_law \
           /app/tests/test_ba_*.py \
           /app/tests/test_labor_law_*.py \
           /app/tests/test_regulatory_*.py || true
RUN pip install --no-cache-dir --no-deps -e .

# Copy compiled static frontend from Stage 1 into /app/dist
COPY --from=frontend-builder /frontend/dist /app/dist

# Ensure tools directory exists and plantuml.jar is present
RUN mkdir -p /app/tools && \
    if [ ! -f /app/tools/plantuml.jar ]; then \
        curl -sSL -o /app/tools/plantuml.jar https://github.com/plantuml/plantuml/releases/download/v1.2024.7/plantuml-1.2024.7.jar || true; \
    fi

# Create runtime and scratch directories
RUN mkdir -p /app/diagram_assets /app/mail_merge_output /app/logs /app/.web_uploads /app/.web_outputs

# Health check probe against API health endpoint
HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 CMD curl -f http://localhost:${PORT:-10000}/api/health || exit 1

EXPOSE 10000

# Start Web Studio Server (reads dynamic $PORT from cloud provider)
CMD ["python", "web_api_server.py"]
