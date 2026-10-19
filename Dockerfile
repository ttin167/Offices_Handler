# ==============================================================================
# Dockerfile — Antigravity Document Handler Enterprise Production Container (v4.1)
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
COPY public/ ./public/
RUN npm run build

# ------------------------------------------------------------------------------
# STAGE 2: Enterprise Python Document Processing & Diagram Execution Runtime
# ------------------------------------------------------------------------------
FROM python:3.11-slim-bookworm

# Configure environment:
# - Cap JVM memory to 192MB with -Xmx192m to prevent 512MB RAM OOM kills on Render Free tier
# - Set PLANTUML_LIMIT_SIZE and UTF-8 charset inside JAVA_TOOL_OPTIONS
# - Set MAX_JOB_WORKERS=1 to serialize memory-intensive headless jobs (LibreOffice / Chromium)
# - Bind HOST to 0.0.0.0 for container networking
ENV PYTHONUNBUFFERED=1 \
    DEBIAN_FRONTEND=noninteractive \
    LANG=C.UTF-8 \
    LC_ALL=C.UTF-8 \
    HOST=0.0.0.0 \
    PORT=8000 \
    MAX_JOB_WORKERS=1 \
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

# Install Python dependencies first for caching efficiency
COPY requirements.txt pyproject.toml /app/
RUN pip install --no-cache-dir --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt && \
    pip install --no-cache-dir -e .

# Copy application codebase
COPY . /app/

# Copy compiled static frontend from Stage 1 into /app/dist
COPY --from=frontend-builder /frontend/dist /app/dist

# Create runtime and scratch directories
RUN mkdir -p /app/diagram_assets /app/mail_merge_output /app/logs /app/.web_uploads /app/.web_outputs

# Health check probe against API health endpoint
HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD curl -f http://localhost:${PORT:-8000}/api/health || exit 1

EXPOSE 8000

# Start Web Studio Server (reads dynamic $PORT from cloud provider)
CMD ["python", "web_api_server.py"]
