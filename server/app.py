"""
web_api_server.py — High-Performance Zero-Dependency Python Web API Bridge.

Serves RESTful endpoints for the Universal Office & Presentation Web Studio.
Designed to be lightweight, cross-platform, and standard-library based (zero external pip web frameworks).
Exposes:
  - GET  /api/health
  - GET  /api/info
  - POST /api/upload
  - POST /api/convert
  - POST /api/docx-to-pptx
  - POST /api/pdf-to-pptx
  - POST /api/render-diagram
  - GET  /api/download/<file>
  - GET  /api/preview/<file>
"""

from __future__ import annotations

import http.server
import json
import os
import sys
import threading
import time
import urllib.parse
from typing import Any, Dict, List, Optional, Tuple, Union

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass

APP_VERSION = "4.6.0"

# Import core handlers
try:
    from core.converter.engine import convert_universal
    from core.pptx.docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx
    from core.pptx.writer import THEMES, write_pptx_from_spec
    from core.docx.writer import patch_docx_template, write_docx
    from core.docx.mail_merge import inspect_template_tokens
    from core.xlsx.writer import mutate_template_excel
    from core.xlsx.validator import validate_excel_sheet
    from core.xlsx.reader import read_xlsx
    from .db import record_history, get_recent_history, delete_history_item, clear_history
    from core.docx.math import parse_inline_math, add_math_to_paragraph
    from core.docx.presets import apply_preset_to_document, load_preset, list_available_presets, build_administrative_header
    from core.converter.pdf_toolkit import (
        merge_pdfs, split_pdf, rotate_and_clean_pdf, stamp_watermark_pdf, compress_pdf,
        redact_sensitive_pdf, protect_pdf, normalize_and_number_pdf, stamp_giap_lai_pdf,
        extract_pdf_tables_to_xlsx, render_pdf_thumbnails, reorganize_pdf_pages,
        convert_images_to_pdf, export_pdf_to_images_zip, sign_pdf_with_stamp, flatten_and_sanitize_pdf
    )
    from core.xlsx.writer import apply_worksheet_print_setup, UniversalExcelEngine
except (ImportError, ValueError):
    from core.converter.engine import convert_universal  # type: ignore
    from core.pptx.docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx  # type: ignore
    from core.pptx.writer import THEMES, write_pptx_from_spec  # type: ignore
    from core.docx.writer import patch_docx_template, write_docx  # type: ignore
    from core.docx.mail_merge import inspect_template_tokens  # type: ignore
    from core.xlsx.writer import mutate_template_excel  # type: ignore
    from core.xlsx.validator import validate_excel_sheet  # type: ignore
    from server.db import record_history, get_recent_history, delete_history_item, clear_history  # type: ignore
    from core.docx.math import parse_inline_math, add_math_to_paragraph  # type: ignore
    from core.docx.presets import apply_preset_to_document, load_preset, list_available_presets, build_administrative_header  # type: ignore
    from core.converter.pdf_toolkit import (  # type: ignore
        merge_pdfs, split_pdf, rotate_and_clean_pdf, stamp_watermark_pdf, compress_pdf,
        redact_sensitive_pdf, protect_pdf, normalize_and_number_pdf, stamp_giap_lai_pdf,
        extract_pdf_tables_to_xlsx, render_pdf_thumbnails, reorganize_pdf_pages,
        convert_images_to_pdf, export_pdf_to_images_zip, sign_pdf_with_stamp, flatten_and_sanitize_pdf
    )
    from core.xlsx.writer import apply_worksheet_print_setup, UniversalExcelEngine  # type: ignore

# Optional local-only module: Digital Yearbook & Photobook Studio (excluded on GitHub/Docker)
try:
    from core.yearbook import (
        YearbookLayoutEngine,
        YearbookPressEngine,
        YearbookHandoffPackager,
        PRESETS as YEARBOOK_PRESETS,
        SPREAD_TEMPLATES as YEARBOOK_TEMPLATES,
        get_default_mock_spec,
    )
    HAS_YEARBOOK = True
except (ImportError, ValueError):
    HAS_YEARBOOK = False
    YearbookLayoutEngine = None  # type: ignore
    YearbookPressEngine = None  # type: ignore
    YearbookHandoffPackager = None  # type: ignore
    YEARBOOK_PRESETS = {}  # type: ignore
    YEARBOOK_TEMPLATES = {}  # type: ignore
    def get_default_mock_spec(*args: Any, **kwargs: Any) -> Dict[str, Any]:  # type: ignore
        return {"title": "Yearbook Module Not Installed", "spreads": [], "pages": []}

try:
    from core.job_manager import global_job_manager
    from core.env_doctor import run_doctor
except (ImportError, ValueError):
    from core.job_manager import global_job_manager  # type: ignore
    from core.env_doctor import run_doctor  # type: ignore

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", 8000))
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST_DIR = os.path.join(BASE_DIR, "dist")
UPLOAD_DIR = os.path.join(BASE_DIR, ".web_uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, ".web_outputs")

# Security & Quota Environment Configurations
API_SECRET_KEY = os.environ.get("API_KEY") or os.environ.get("OFFICE_API_KEY") or ""
MAX_UPLOAD_SIZE_MB = int(os.environ.get("MAX_UPLOAD_SIZE_MB", 50))
MAX_UPLOAD_SIZE_BYTES = MAX_UPLOAD_SIZE_MB * 1024 * 1024
ALLOWED_ORIGINS = [o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "").split(",") if o.strip()]

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def cleanup_old_temp_files(max_age_seconds: int = 7200) -> int:
    """
    Purges temporary files in UPLOAD_DIR and OUTPUT_DIR older than max_age_seconds (default 2h).
    Protects Ephemeral Container Disk on memory-constrained cloud environments (Render Free 512MB).
    """
    now = time.time()
    removed_count = 0
    for target_dir in (UPLOAD_DIR, OUTPUT_DIR):
        if not os.path.isdir(target_dir):
            continue
        try:
            for fname in os.listdir(target_dir):
                fpath = os.path.join(target_dir, fname)
                if os.path.isfile(fpath):
                    try:
                        mtime = os.path.getmtime(fpath)
                        if now - mtime > max_age_seconds:
                            os.remove(fpath)
                            removed_count += 1
                    except OSError:
                        pass
        except OSError:
            pass
    return removed_count


def export_pptx_slides(pptx_path: str, output_dir: str, slides_arg: Any = "all") -> list[dict[str, Any]]:
    """Exports presentation slides to high-res PNG images using Win32 COM with PyMuPDF fallback."""
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(pptx_path))[0]
    exported: list[dict[str, Any]] = []

    # 1. Native Windows PowerPoint COM automation (Pixel-perfect 1920x1080)
    if sys.platform == "win32":
        try:
            import win32com.client
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            deck = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
            total = deck.Slides.Count

            if str(slides_arg).lower() == "all":
                target_indices = list(range(1, total + 1))
            elif isinstance(slides_arg, list):
                target_indices = [int(x) for x in slides_arg if str(x).isdigit()]
            else:
                target_indices = []
                for part in str(slides_arg).split(","):
                    part = part.strip()
                    if "-" in part:
                        try:
                            p_start, p_end = part.split("-", 1)
                            target_indices.extend(range(int(p_start.strip()), int(p_end.strip()) + 1))
                        except Exception:
                            pass
                    elif part.isdigit():
                        target_indices.append(int(part))

            for s_idx in target_indices:
                if 1 <= s_idx <= total:
                    fname = f"preview_{base_name}_slide_{s_idx:02d}.png"
                    out_png = os.path.join(output_dir, fname)
                    deck.Slides(s_idx).Export(out_png, "PNG", 1920, 1080)
                    exported.append({
                        "slide": s_idx,
                        "filename": fname,
                        "path": out_png,
                        "preview_url": f"/api/preview/{fname}"
                    })

            deck.Close()
            powerpoint.Quit()
        except Exception as com_err:
            print(f"[WARN] PowerPoint COM export failed or not available: {com_err}")

    # 2. Headless LibreOffice + PyMuPDF fallback
    if not exported:
        try:
            import tempfile
            import subprocess
            import pymupdf  # type: ignore

            temp_pdf_dir = tempfile.mkdtemp()
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", os.path.abspath(pptx_path), "--outdir", temp_pdf_dir],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
            )
            pdf_name = f"{base_name}.pdf"
            temp_pdf = os.path.join(temp_pdf_dir, pdf_name)
            if os.path.isfile(temp_pdf):
                doc = pymupdf.open(temp_pdf)
                total = len(doc)
                if str(slides_arg).lower() == "all":
                    target_indices = list(range(1, total + 1))
                elif isinstance(slides_arg, list):
                    target_indices = [int(x) for x in slides_arg if str(x).isdigit()]
                else:
                    target_indices = [int(x.strip()) for x in str(slides_arg).split(",") if x.strip().isdigit()]

                for s_idx in target_indices:
                    if 1 <= s_idx <= total:
                        page = doc[s_idx - 1]
                        pix = page.get_pixmap(dpi=150)
                        fname = f"preview_{base_name}_slide_{s_idx:02d}.png"
                        out_png = os.path.join(output_dir, fname)
                        pix.save(out_png)
                        exported.append({
                            "slide": s_idx,
                            "filename": fname,
                            "path": out_png,
                            "preview_url": f"/api/preview/{fname}"
                        })
                doc.close()
        except Exception as fb_err:
            print(f"[WARN] Headless fallback failed: {fb_err}")

    return exported


def resolve_file_path(path_str: Optional[str]) -> Optional[str]:
    """Safely resolves an input file path against allowed sandbox roots."""
    if not path_str or not isinstance(path_str, str):
        return None
    cleaned = path_str.strip()
    if not cleaned:
        return None

    import tempfile
    allowed_roots = [
        os.path.abspath(BASE_DIR),
        os.path.abspath(UPLOAD_DIR),
        os.path.abspath(OUTPUT_DIR),
        os.path.abspath(tempfile.gettempdir()),
    ]

    def _is_safe(candidate_abs: str) -> bool:
        norm_cand = os.path.normpath(candidate_abs)
        for r in allowed_roots:
            norm_r = os.path.normpath(r)
            try:
                if os.path.commonpath([norm_r, norm_cand]) == norm_r:
                    return True
            except (ValueError, OSError):
                continue
        return False

    # 1. Exact path or relative to current working directory
    if os.path.isfile(cleaned):
        abs_p = os.path.abspath(cleaned)
        if _is_safe(abs_p):
            return abs_p

    # 2. Check inside UPLOAD_DIR by basename
    safe_base = os.path.basename(cleaned)
    in_upload = os.path.join(UPLOAD_DIR, safe_base)
    if os.path.isfile(in_upload) and _is_safe(in_upload):
        return os.path.abspath(in_upload)

    # 3. Check inside OUTPUT_DIR by basename
    in_output = os.path.join(OUTPUT_DIR, safe_base)
    if os.path.isfile(in_output) and _is_safe(in_output):
        return os.path.abspath(in_output)

    # 3b. Check inside OUTPUT_DIR/pdf_previews by basename
    in_previews = os.path.join(OUTPUT_DIR, "pdf_previews", safe_base)
    if os.path.isfile(in_previews) and _is_safe(in_previews):
        return os.path.abspath(in_previews)

    # 4. Check relative to BASE_DIR
    in_base = os.path.join(BASE_DIR, cleaned)
    if os.path.isfile(in_base):
        abs_b = os.path.abspath(in_base)
        if _is_safe(abs_b):
            return abs_b

    return None


class OfficeApiHandler(http.server.BaseHTTPRequestHandler):
    server_version = "AntigravityOfficeStudio/4.0"

    def _send_cors_headers(self) -> None:
        client_origin = self.headers.get("Origin", "")
        if not ALLOWED_ORIGINS:
            # Default: allow localhost, 127.0.0.1, app://, or * for universal developer tools
            allowed_origin = client_origin if client_origin else "*"
            self.send_header("Access-Control-Allow-Origin", allowed_origin)
        else:
            if "*" in ALLOWED_ORIGINS:
                self.send_header("Access-Control-Allow-Origin", "*")
            elif client_origin in ALLOWED_ORIGINS:
                self.send_header("Access-Control-Allow-Origin", client_origin)
            elif client_origin and (
                client_origin.startswith("http://localhost:")
                or client_origin.startswith("http://127.0.0.1:")
                or client_origin.startswith("app://")
            ):
                self.send_header("Access-Control-Allow-Origin", client_origin)
            else:
                self.send_header("Access-Control-Allow-Origin", ALLOWED_ORIGINS[0])

        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With, X-API-Key")
        self.send_header("Vary", "Origin")

    def _check_auth(self, path: str) -> bool:
        """
        Validates API_SECRET_KEY if configured in environment.
        Exempts /api/health and static web assets.
        """
        if not API_SECRET_KEY:
            return True
        if path == "/api/health" or not path.startswith("/api/"):
            return True

        # Check Authorization header (Bearer token)
        auth_hdr = self.headers.get("Authorization", "").strip()
        if auth_hdr:
            token = auth_hdr[7:].strip() if auth_hdr.lower().startswith("bearer ") else auth_hdr
            if token == API_SECRET_KEY:
                return True

        # Check X-API-Key header
        x_key = self.headers.get("X-API-Key", "").strip()
        if x_key == API_SECRET_KEY:
            return True

        # Check query string (?api_key=...)
        parsed = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed.query)
        q_key = qs.get("api_key", [None])[0]
        if q_key == API_SECRET_KEY:
            return True

        self._send_json({
            "success": False,
            "error": "Unauthorized: Missing or invalid API key. Set 'Authorization: Bearer <API_KEY>'."
        }, status=401)
        return False

    def do_OPTIONS(self) -> None:
        self.send_response(200)
        self._send_cors_headers()
        self.end_headers()

    def do_HEAD(self) -> None:
        """Handles HTTP HEAD requests from cloud load balancers and health probes."""
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        if path == "/api/health" or path in ("/", "/index.html"):
            self.send_response(200)
            self._send_cors_headers()
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
        else:
            self.do_GET()

    def _send_json(self, data: Any, status: int = 200) -> None:
        encoded = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self._send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def _read_body_json(self) -> Dict[str, Any]:
        content_length = int(self.headers.get("Content-Length", 0))
        if content_length <= 0:
            return {}
        if content_length > MAX_UPLOAD_SIZE_BYTES:
            raise ValueError(f"Payload Too Large: {content_length} bytes exceeds maximum allowed {MAX_UPLOAD_SIZE_MB}MB")
        raw_body = self.rfile.read(content_length).decode("utf-8", errors="replace")
        return json.loads(raw_body) if raw_body else {}

    def do_GET(self) -> None:
        try:
            parsed_url = urllib.parse.urlparse(self.path)
            path = parsed_url.path

            if not self._check_auth(path):
                return

            if path == "/api/health":
                self._send_json({
                    "status": "online",
                    "healthy": True,
                    "version": APP_VERSION,
                    "service": "Antigravity Office & Presentation Studio",
                    "timestamp": time.time(),
                })

            elif path == "/api/info":
                self._send_json({
                    "name": "Antigravity Office Studio",
                    "version": APP_VERSION,
                    "capabilities": [
                        "docx-patch",
                        "xlsx-mutate",
                        "xlsx-validate",
                        "xlsx-read",
                        "pptx-create",
                        "pptx-build",
                        "pptx-preview",
                        "in-place-preservation",
                        "diagram-render",
                        "universal-convert",
                        "env-doctor",
                        "sse-job-queue",
                        "pdf-tools",
                        "pdf-redact",
                        "pdf-protect",
                        "pdf-giap-lai",
                        "pdf-normalize-a4",
                        "pdf-table-extract",
                        "pdf-organizer",
                        "yearbook-studio",
                        "photobook-press",
                    ],
                    "themes": list(THEMES.keys()),
                    "formats": [".docx", ".pptx", ".pdf", ".md", ".xlsx"],
                    "invariants": [
                        "PPTX_INV_01_TO_08",
                        "ERR_DOCX_001_TO_006",
                        "ERR_XLSX_001_TO_007",
                        "E1_TO_E14"
                    ],
                    "upload_dir": UPLOAD_DIR,
                    "output_dir": OUTPUT_DIR,
                })

            elif path == "/api/doctor":
                self._send_json(run_doctor())

            elif path == "/api/presets":
                try:
                    presets = list_available_presets()
                except Exception:
                    presets = ["nd30", "academic_thesis", "corporate_report"]
                self._send_json({"success": True, "presets": presets})

            elif path.startswith("/api/jobs/"):
                parts = path.strip("/").split("/")
                if len(parts) >= 3 and parts[2] == "stream":
                    job_id = parts[1]
                    self._handle_job_stream(job_id)
                    return
                else:
                    job_id = parts[1]
                    job = global_job_manager.get_job(job_id)
                    if not job:
                        self._send_json({"error": f"Job not found: {job_id}"}, status=404)
                    else:
                        self._send_json(job)
                    return

            elif path == "/api/jobs":
                jobs = global_job_manager.list_jobs(limit=50)
                self._send_json({"jobs": jobs, "total": len(jobs)})
                return

            elif path == "/api/history":
                items = get_recent_history(limit=50)
                self._send_json({"items": items, "total": len(items)})

            elif path == "/api/yearbook/mock-album":
                self._send_json({"success": True, "spec": get_default_mock_spec()})
                return

            elif path == "/api/yearbook/presets":
                self._send_json({
                    "success": True,
                    "presets": YEARBOOK_PRESETS,
                    "templates": YEARBOOK_TEMPLATES,
                })
                return

            elif path == "/api/xlsx/formulas":
                query_params = urllib.parse.parse_qs(parsed_url.query)
                cat = query_params.get("category", ["all"])[0]
                q = query_params.get("search", [None])[0]
                try:
                    from core.xlsx.formula_playbook import get_formula_catalog
                except (ImportError, ValueError):
                    try:
                        from formula_playbook import get_formula_catalog  # type: ignore
                    except (ImportError, ValueError):
                        self._send_json({"error": "Failed to import formula playbook"}, status=500)
                        return
                items = get_formula_catalog(category=cat, search=q)
                self._send_json({"success": True, "count": len(items), "formulas": items})
                return

            elif path == "/api/xlsx/formulas/export":
                try:
                    from core.xlsx.formula_playbook import generate_formula_playbook_xlsx
                except (ImportError, ValueError):
                    try:
                        from formula_playbook import generate_formula_playbook_xlsx  # type: ignore
                    except (ImportError, ValueError):
                        self._send_json({"error": "Failed to import formula playbook"}, status=500)
                        return
                export_filename = "Excel_Formula_Playbook.xlsx"
                dest = os.path.join(OUTPUT_DIR, export_filename)
                ok, res = generate_formula_playbook_xlsx(dest)
                if not ok:
                    self._send_json({"success": False, "error": res}, status=500)
                    return
                try:
                    record_history("xlsx-playbook", export_filename, export_filename, res)
                except Exception:
                    pass
                self._serve_file(export_filename, as_attachment=True)
                return

            elif path.startswith("/api/download/"):
                filename = urllib.parse.unquote(path[len("/api/download/"):])
                self._serve_file(filename, as_attachment=True)

            elif path.startswith("/api/preview/"):
                filename = urllib.parse.unquote(path[len("/api/preview/"):])
                self._serve_file(filename, as_attachment=False)

            elif path.startswith("/api/"):
                self._send_json({"error": f"Endpoint not found: {path}"}, status=404)

            else:
                self._serve_static_file(path)
        except Exception as exc:
            import traceback
            traceback.print_exc()
            self._send_json({"error": f"Internal server error: {str(exc)}"}, status=500)

    def do_DELETE(self) -> None:
        try:
            parsed_url = urllib.parse.urlparse(self.path)
            path = parsed_url.path

            if not self._check_auth(path):
                return

            if path.startswith("/api/history/"):
                id_str = path[len("/api/history/"):]
                if id_str.isdigit():
                    ok = delete_history_item(int(id_str))
                    self._send_json({"success": ok})
                    return
                self._send_json({"error": "Invalid history ID"}, status=400)
                return

            elif path == "/api/history":
                ok = clear_history()
                self._send_json({"success": ok})
                return

            self._send_json({"error": f"Endpoint not found: {path}"}, status=404)
        except Exception as exc:
            self._send_json({"error": f"Internal server error: {str(exc)}"}, status=500)

    def do_POST(self) -> None:
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        if not self._check_auth(path):
            return

        content_type = self.headers.get("Content-Type", "")

        try:
            # Endpoint 0: Asynchronous Job Submission (SSE Streaming Queue)
            if path == "/api/jobs/submit":
                payload = self._read_body_json()
                action = payload.get("action", "convert")
                job_payload = payload.get("payload", payload)
                job_id = self._dispatch_job_submission(action, job_payload)
                self._send_json({
                    "success": True,
                    "job_id": job_id,
                    "status": "pending",
                    "stream_url": f"/api/jobs/{job_id}/stream",
                    "status_url": f"/api/jobs/{job_id}/status",
                })
                return

            # Endpoint 1: File Upload (multipart or json base64)
            elif path == "/api/upload":
                self._handle_upload(content_type)

            # Endpoint 2: Universal Document Conversion
            elif path == "/api/convert":
                payload = self._read_body_json()
                src_param = (
                    payload.get("source_path")
                    or payload.get("input_path")
                    or payload.get("filepath")
                    or payload.get("file_path")
                    or payload.get("docx_path")
                    or payload.get("pdf_path")
                    or payload.get("filename")
                )
                src = resolve_file_path(src_param)
                dest_ext = payload.get("target_ext", ".docx")
                pages = payload.get("page_range")

                if not src:
                    self._send_json({"success": False, "error": f"Source file not found: {src_param}"}, status=400)
                    return

                base_name = os.path.splitext(os.path.basename(src))[0]
                target_path = os.path.join(OUTPUT_DIR, f"{base_name}_converted{dest_ext}")

                ok, res, elap = convert_universal(
                    source_path=src,
                    target_path=target_path,
                    target_ext=dest_ext,
                    page_range=pages
                )

                if not ok:
                    self._send_json({
                        "success": False,
                        "error": res,
                        "result_path": res,
                        "elapsed": elap
                    }, status=500)
                    return

                out_name = os.path.basename(res)
                quoted_out = urllib.parse.quote(out_name)
                try:
                    record_history("convert", os.path.basename(src), out_name, res)
                except Exception:
                    pass

                self._send_json({
                    "success": True,
                    "result_path": res,
                    "filename": out_name,
                    "output_file": out_name,
                    "download_url": f"/api/download/{quoted_out}",
                    "preview_url": f"/api/preview/{quoted_out}",
                    "elapsed": elap
                })

            # Endpoint 3: Word to PowerPoint Presentation
            elif path == "/api/docx-to-pptx":
                payload = self._read_body_json()
                src_param = (
                    payload.get("docx_path")
                    or payload.get("source_path")
                    or payload.get("input_path")
                    or payload.get("filepath")
                    or payload.get("file_path")
                    or payload.get("filename")
                )
                src = resolve_file_path(src_param)
                theme = payload.get("theme", "thesis_blue")
                deck_scope = payload.get("deck_scope") or payload.get("scope") or "auto"
                tmpl_param = payload.get("template_path") or payload.get("template_pptx")
                template = resolve_file_path(tmpl_param) if tmpl_param else None

                if not src:
                    self._send_json({"success": False, "error": f"Word file not found: {src_param}"}, status=400)
                    return

                base_name = os.path.splitext(os.path.basename(src))[0]
                out_pptx = os.path.join(OUTPUT_DIR, f"{base_name}_presentation.pptx")
                out_spec = os.path.join(OUTPUT_DIR, f"{base_name}_slide_spec.json")

                t0 = time.time()
                final_pptx, final_spec = convert_docx_to_pptx(
                    docx_path=src,
                    output_pptx=out_pptx,
                    spec_out=out_spec,
                    theme=theme,
                    template_path=template,
                    deck_scope=deck_scope,
                )

                # Load generated spec for instant frontend UI preview
                spec_data = {}
                if os.path.isfile(final_spec):
                    with open(final_spec, "r", encoding="utf-8") as f:
                        spec_data = json.load(f)

                out_name = os.path.basename(final_pptx)
                spec_name = os.path.basename(final_spec)
                try:
                    record_history("docx_to_pptx", os.path.basename(src), out_name, final_pptx)
                except Exception:
                    pass

                self._send_json({
                    "success": True,
                    "pptx_path": final_pptx,
                    "spec_path": final_spec,
                    "filename": out_name,
                    "output_file": out_name,
                    "pptx_file": out_name,
                    "spec_file": spec_name,
                    "download_url": f"/api/download/{out_name}",
                    "spec": spec_data,
                    "slides_count": len(spec_data.get("slides", [])),
                    "elapsed": time.time() - t0
                })

            # Endpoint 4: PDF to PowerPoint Presentation
            elif path == "/api/pdf-to-pptx":
                payload = self._read_body_json()
                src_param = (
                    payload.get("pdf_path")
                    or payload.get("source_path")
                    or payload.get("input_path")
                    or payload.get("filepath")
                    or payload.get("file_path")
                    or payload.get("filename")
                )
                src = resolve_file_path(src_param)
                theme = payload.get("theme", "thesis_blue")
                deck_scope = payload.get("deck_scope") or payload.get("scope") or "auto"
                pages = payload.get("pages")
                tmpl_param = payload.get("template_path") or payload.get("template_pptx")
                template = resolve_file_path(tmpl_param) if tmpl_param else None

                if not src:
                    self._send_json({"success": False, "error": f"PDF file not found: {src_param}"}, status=400)
                    return

                base_name = os.path.splitext(os.path.basename(src))[0]
                out_pptx = os.path.join(OUTPUT_DIR, f"{base_name}_presentation.pptx")
                out_spec = os.path.join(OUTPUT_DIR, f"{base_name}_slide_spec.json")

                t0 = time.time()
                final_pptx, final_spec = convert_pdf_to_pptx(
                    pdf_path=src,
                    output_pptx=out_pptx,
                    spec_out=out_spec,
                    theme=theme,
                    template_path=template,
                    page_range=pages,
                    deck_scope=deck_scope,
                )


                spec_data = {}
                if os.path.isfile(final_spec):
                    with open(final_spec, "r", encoding="utf-8") as f:
                        spec_data = json.load(f)

                out_name = os.path.basename(final_pptx)
                spec_name = os.path.basename(final_spec)
                try:
                    record_history("pdf_to_pptx", os.path.basename(src), out_name, final_pptx)
                except Exception:
                    pass

                self._send_json({
                    "success": True,
                    "pptx_path": final_pptx,
                    "spec_path": final_spec,
                    "filename": out_name,
                    "output_file": out_name,
                    "pptx_file": out_name,
                    "spec_file": spec_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "spec": spec_data,
                    "slides_count": len(spec_data.get("slides", [])),
                    "elapsed": time.time() - t0
                })

            # Endpoint 5: Technical Diagram Render
            elif path == "/api/render-diagram":
                payload = self._read_body_json()
                engine = payload.get("engine", "mermaid")
                spec_content = payload.get("spec") or payload.get("code") or ""

                base_time = int(time.time() * 1000)
                out_png = os.path.join(OUTPUT_DIR, f"diagram_{base_time}.png")

                if engine == "mermaid":
                    try:
                        from core.diagrams.legacy.mermaid import render_mermaid_to_png
                    except (ImportError, ValueError):
                        from core.diagrams.legacy.mermaid import render_mermaid_to_png  # type: ignore

                    if isinstance(spec_content, dict):
                        m_spec = dict(spec_content)
                        m_spec.setdefault("output_path", out_png)
                    else:
                        m_spec = {
                            "mermaid_code": spec_content,
                            "code": spec_content,
                            "output_path": out_png,
                            "scale": payload.get("scale", 2),
                        }
                    render_mermaid_to_png(m_spec)
                elif engine == "plantuml":
                    try:
                        from core.diagrams.legacy.plantuml import render_plantuml_to_png
                    except (ImportError, ValueError):
                        from core.diagrams.legacy.plantuml import render_plantuml_to_png  # type: ignore

                    if isinstance(spec_content, dict):
                        p_spec = dict(spec_content)
                        p_spec.setdefault("output_path", out_png)
                    else:
                        p_spec = {
                            "plantuml_code": spec_content,
                            "code": spec_content,
                            "output_path": out_png,
                            "scale": payload.get("scale", 2),
                        }
                    render_plantuml_to_png(p_spec)
                elif engine == "drawio":
                    try:
                        from core.diagrams.mxgraph import render_diagram, render_mxgraph_to_png
                    except (ImportError, ValueError):
                        from core.diagrams.mxgraph import render_diagram, render_mxgraph_to_png  # type: ignore

                    if isinstance(spec_content, str) and spec_content.strip().startswith("{"):
                        try:
                            d_spec = json.loads(spec_content)
                        except Exception:
                            d_spec = {}
                        d_spec["output_path"] = out_png
                        d_spec["scale"] = int(payload.get("scale", 2))
                        render_diagram(d_spec)
                    else:
                        if isinstance(spec_content, dict):
                            xml_code = spec_content.get("xml_content") or spec_content.get("code") or ""
                        else:
                            xml_code = str(spec_content)
                        render_mxgraph_to_png(xml_code, out_png, scale=int(payload.get("scale", 2)))
                elif engine in ("canvas", "spec"):
                    try:
                        from core.diagrams.legacy.canvas import PrecisionDiagram
                    except (ImportError, ValueError):
                        from legacy_engines.spec_diagram_engine import PrecisionDiagram  # type: ignore

                    if isinstance(spec_content, str):
                        try:
                            c_spec = json.loads(spec_content)
                        except Exception:
                            c_spec = {"nodes": [], "edges": []}
                    else:
                        c_spec = dict(spec_content)
                    scale = int(payload.get("scale", 2))
                    diag = PrecisionDiagram.from_spec(c_spec)
                    diag.render_to_png(out_png, scale=scale)

                out_name = os.path.basename(out_png)
                quoted_name = urllib.parse.quote(out_name)
                try:
                    record_history("diagram_render", f"{engine}_diagram", out_name, out_png)
                except Exception:
                    pass

                self._send_json({
                    "success": True,
                    "image_path": out_png,
                    "filename": out_name,
                    "output_file": out_name,
                    "image_file": out_name,
                    "download_url": f"/api/download/{quoted_name}",
                    "preview_url": f"/api/preview/{quoted_name}"
                })

            # Endpoint 5.5: Inspect DOCX Template Tokens / Placeholders
            elif path == "/api/docx/inspect-tokens":
                payload = self._read_body_json()
                src_param = (
                    payload.get("template_path")
                    or payload.get("source_path")
                    or payload.get("input_path")
                    or payload.get("filepath")
                    or payload.get("file_path")
                )
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"Template DOCX file not found: {src_param}"}, status=400)
                    return

                try:
                    try:
                        from core.docx.mail_merge import inspect_template_tokens as _fn_tokens
                    except (ImportError, ValueError):
                        from core.docx.mail_merge import inspect_template_tokens as _fn_tokens  # type: ignore
                    res = _fn_tokens(src)
                    self._send_json({
                        "success": True,
                        "template_file": res.get("template_file"),
                        "variables": res.get("variables", []),
                        "conditionals": res.get("conditionals", []),
                        "blocks": res.get("blocks", []),
                        "total_unique_tokens": res.get("total_unique_tokens", 0),
                    })
                except Exception as ex:
                    self._send_json({"success": False, "error": str(ex)}, status=500)

            # Endpoint 6: In-Place DOCX Template Patching
            elif path == "/api/docx/patch":
                payload = self._read_body_json()
                src_param = (
                    payload.get("template_path")
                    or payload.get("source_path")
                    or payload.get("input_path")
                    or payload.get("filepath")
                    or payload.get("file_path")
                )
                src = resolve_file_path(src_param)
                text_rep = payload.get("text_replacements") or payload.get("replacements") or (payload.get("patch_spec") or {}).get("text_replacements", {})
                block_rep = payload.get("block_replacements") or (payload.get("patch_spec") or {}).get("block_replacements", {})
                custom_name = payload.get("output_name")

                if isinstance(text_rep, dict):
                    norm: Dict[str, str] = {}
                    for k, v in text_rep.items():
                        k_str = str(k).strip()
                        v_str = str(v)
                        norm[k_str] = v_str
                        if not k_str.startswith("{{") and not k_str.endswith("}}"):
                            norm[f"{{{{{k_str}}}}}"] = v_str
                    text_rep = norm

                if not src:
                    self._send_json({"success": False, "error": f"Template DOCX file not found: {src_param}"}, status=400)
                    return

                base_name = os.path.splitext(os.path.basename(src))[0]
                out_name = custom_name or f"{base_name}_patched.docx"
                out_path = os.path.join(OUTPUT_DIR, out_name)

                t0 = time.time()
                try:
                    res_path = patch_docx_template(src, out_path, text_rep, block_rep)
                    elap = time.time() - t0
                    bname = os.path.basename(res_path)
                    qbname = urllib.parse.quote(bname)
                    self._send_json({
                        "success": True,
                        "result_path": res_path,
                        "filename": bname,
                        "output_file": bname,
                        "download_url": f"/api/download/{qbname}",
                        "preview_url": f"/api/preview/{qbname}",
                        "elapsed": elap
                    })
                except Exception as ex:
                    self._send_json({"success": False, "error": str(ex)}, status=500)

            # Endpoint 7: Dynamic Excel Template Mutation
            elif path == "/api/xlsx/mutate":
                payload = self._read_body_json()
                src_param = (
                    payload.get("template_path")
                    or payload.get("source_path")
                    or payload.get("input_path")
                    or payload.get("filepath")
                    or payload.get("file_path")
                )
                src = resolve_file_path(src_param)
                spec = payload.get("spec", {})
                custom_name = payload.get("output_name")

                if not src:
                    self._send_json({"success": False, "error": f"Template XLSX file not found: {src_param}"}, status=400)
                    return

                base_name = os.path.splitext(os.path.basename(src))[0]
                out_name = custom_name or f"{base_name}_mutated.xlsx"
                out_path = os.path.join(OUTPUT_DIR, out_name)

                t0 = time.time()
                try:
                    res_path = mutate_template_excel(src, out_path, spec)
                    elap = time.time() - t0
                    bname = os.path.basename(res_path)
                    qbname = urllib.parse.quote(bname)
                    self._send_json({
                        "success": True,
                        "result_path": res_path,
                        "filename": bname,
                        "output_file": bname,
                        "download_url": f"/api/download/{qbname}",
                        "elapsed": elap
                    })
                except Exception as ex:
                    self._send_json({"success": False, "error": str(ex)}, status=500)

            # Endpoint 8: Excel 12 Quality Gates Diff & Validator
            elif path == "/api/xlsx/validate":
                payload = self._read_body_json()
                src_param = (
                    payload.get("file_path")
                    or payload.get("filepath")
                    or payload.get("template_path")
                    or payload.get("source_path")
                    or payload.get("xlsx_path")
                    or payload.get("filename")
                )
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"Excel file not found: {src_param}"}, status=400)
                    return

                target_sheet = payload.get("target_sheet") or payload.get("sheet_name")
                if target_sheet and str(target_sheet).strip().lower() in ("none", "null", ""):
                    target_sheet = None
                ref_sheet = payload.get("ref_sheet")
                if ref_sheet and str(ref_sheet).strip().lower() in ("none", "null", ""):
                    ref_sheet = None

                t0 = time.time()
                try:
                    res = validate_excel_sheet(
                        src,
                        target_sheet=target_sheet,
                        ref_sheet=ref_sheet,
                        verbose=False,
                    )
                    elap = time.time() - t0
                    self._send_json({
                        "success": True,
                        "passed": res.passed,
                        "target_sheet": res.target_sheet,
                        "ref_sheet": res.ref_sheet,
                        "gates_evaluated": res.gates_evaluated,
                        "gates_failed": res.gates_failed,
                        "diffs": res.diffs,
                        "diff_count": len(res.diffs),
                        "elapsed": elap,
                    })
                except Exception as ex:
                    self._send_json({"success": False, "error": str(ex)}, status=500)

            # Endpoint 9: Excel Headless JSON Snapshot Inspector
            elif path == "/api/xlsx/read":
                payload = self._read_body_json()
                src_param = (
                    payload.get("file_path")
                    or payload.get("filepath")
                    or payload.get("template_path")
                    or payload.get("source_path")
                    or payload.get("xlsx_path")
                    or payload.get("filename")
                )
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"Excel file not found: {src_param}"}, status=400)
                    return

                sheet_name = payload.get("sheet_name") or payload.get("sheet")
                if sheet_name and str(sheet_name).strip().lower() in ("none", "null", ""):
                    sheet_name = None
                engine = payload.get("engine", "auto")
                cell_range = payload.get("cell_range")
                if cell_range and str(cell_range).strip().lower() in ("none", "null", ""):
                    cell_range = None

                t0 = time.time()
                try:
                    snapshot = read_xlsx(
                        xlsx_path=src,
                        sheet_name=sheet_name,
                        cell_range=cell_range,
                        engine=engine,
                    )
                    elap = time.time() - t0
                    sheet_names = [s.get("sheet_name") for s in snapshot.get("sheets", []) if s.get("sheet_name")]
                    self._send_json({
                        "success": True,
                        "engine": engine,
                        "snapshot": snapshot,
                        "sheet_names": sheet_names,
                        "elapsed": elap,
                    })
                except Exception as ex:
                    self._send_json({"success": False, "error": str(ex)}, status=500)

            # Endpoint 10: Declarative Slide Spec to PowerPoint Build
            elif path == "/api/pptx/build":
                payload = self._read_body_json()
                spec = payload.get("spec")
                if not spec:
                    self._send_json({"success": False, "error": "Missing 'spec' in request body"}, status=400)
                    return

                if isinstance(spec, str):
                    resolved_spec_path = resolve_file_path(spec)
                    if resolved_spec_path and os.path.isfile(resolved_spec_path):
                        with open(resolved_spec_path, "r", encoding="utf-8") as f:
                            spec = json.load(f)
                    else:
                        try:
                            spec = json.loads(spec)
                        except Exception:
                            self._send_json({"success": False, "error": "Invalid JSON spec string"}, status=400)
                            return

                theme = payload.get("theme") or spec.get("theme", "thesis_blue")
                tmpl_param = payload.get("template_path") or payload.get("template_pptx") or spec.get("template")
                template = resolve_file_path(tmpl_param) if tmpl_param else None

                custom_name = payload.get("output_name")
                base_time = int(time.time() * 1000)
                out_name = custom_name or f"presentation_{base_time}.pptx"
                out_path = os.path.join(OUTPUT_DIR, out_name)

                t0 = time.time()
                try:
                    res_path = write_pptx_from_spec(
                        spec=spec,
                        output_path=out_path,
                        theme_name=theme,
                        template_path=template,
                    )
                    elap = time.time() - t0
                    fname = os.path.basename(res_path)
                    qfname = urllib.parse.quote(fname)
                    self._send_json({
                        "success": True,
                        "pptx_path": res_path,
                        "filename": fname,
                        "output_file": fname,
                        "download_url": f"/api/download/{qfname}",
                        "slides_count": len(spec.get("slides", [])),
                        "elapsed": elap,
                    })
                except Exception as ex:
                    self._send_json({"success": False, "error": str(ex)}, status=500)

            # Endpoint 11: Export PowerPoint Slide PNG Previews
            elif path == "/api/pptx/preview":
                payload = self._read_body_json()
                src_param = (
                    payload.get("pptx_path")
                    or payload.get("file_path")
                    or payload.get("filepath")
                    or payload.get("filename")
                    or payload.get("output_file")
                )
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"PowerPoint file not found: {src_param}"}, status=400)
                    return

                slides_arg = payload.get("slides", "all")
                t0 = time.time()
                try:
                    previews = export_pptx_slides(src, OUTPUT_DIR, slides_arg=slides_arg)
                    elap = time.time() - t0
                    self._send_json({
                        "success": True,
                        "previews": previews,
                        "total_previews": len(previews),
                        "elapsed": elap,
                    })
                except Exception as ex:
                    self._send_json({"success": False, "error": str(ex)}, status=500)

            # Endpoint 12: Math & LaTeX to DOCX Export
            elif path == "/api/math/export":
                payload = self._read_body_json()
                latex_text = payload.get("latex", "").strip()
                doc_title = payload.get("title", "Công Thức Toán Học").strip()
                if not latex_text:
                    self._send_json({"success": False, "error": "Thiếu nội dung công thức toán LaTeX (latex field)"}, status=400)
                    return

                t0 = time.time()
                ts = int(time.time())
                out_name = f"Math_Export_{ts}.docx"
                out_path = os.path.join(OUTPUT_DIR, out_name)

                doc_ast = {
                    "title": doc_title,
                    "body": [
                        {
                            "type": "paragraph",
                            "style_name": "Heading 1",
                            "runs": [{"text": doc_title, "bold": True, "font_size": 16.0}]
                        },
                        {
                            "type": "paragraph",
                            "style_name": "Normal",
                            "runs": [{"text": "Biểu thức toán học được trích xuất tự động qua Native OMML & OpenXML Math Engine:"}]
                        }
                    ]
                }
                for line in latex_text.splitlines():
                    s = line.strip()
                    if s:
                        doc_ast["body"].append({
                            "type": "paragraph",
                            "text": s
                        })

                write_docx(doc_ast, out_path)
                elap = round(time.time() - t0, 3)
                quoted_out = urllib.parse.quote(out_name)
                try:
                    record_history("math_export", doc_title, out_name, out_path)
                except Exception:
                    pass

                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "output_file": out_name,
                    "download_url": f"/api/download/{quoted_out}",
                    "elapsed": elap
                })

            # Endpoint 13: Parse Math LaTeX Tokens
            elif path == "/api/math/parse":
                payload = self._read_body_json()
                latex_text = payload.get("latex", "")
                tokens = parse_inline_math(latex_text)
                self._send_json({
                    "success": True,
                    "tokens": [
                        {
                            "text": t.text,
                            "is_italic": t.is_italic,
                            "is_bold": t.is_bold,
                            "is_subscript": t.is_subscript,
                            "is_superscript": t.is_superscript,
                            "scale_size": t.scale_size
                        } for t in tokens
                    ]
                })

            # Endpoint 14: Apply Document Presets (ND30, Academic, Corporate)
            elif path == "/api/docx/preset":
                payload = self._read_body_json()
                src_param = payload.get("file_path") or payload.get("filepath") or payload.get("source_path") or payload.get("filename")
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"File not found: {src_param}"}, status=400)
                    return
                preset_name = payload.get("preset", "nd30")
                add_header = bool(payload.get("add_header", False))
                out_param = payload.get("output_path")

                from docx import Document
                doc = Document(src)
                preset_res = apply_preset_to_document(doc, preset_name)

                if add_header:
                    build_administrative_header(
                        doc,
                        agency_parent=payload.get("agency_parent", "BỘ THÔNG TIN VÀ TRUYỀN THÔNG"),
                        agency_name=payload.get("agency_name", "CỤC CÔNG NGHỆ THÔNG TIN"),
                        document_no=payload.get("document_no", "Số: 01/QĐ-CNTT"),
                        location_date=payload.get("location_date", "Hà Nội, ngày 18 tháng 10 năm 2026"),
                        insert_at_top=True,
                    )

                base_name = os.path.splitext(os.path.basename(src))[0]
                resolved_out = resolve_file_path(out_param) if out_param else None
                actual_target = resolved_out or os.path.join(OUTPUT_DIR, f"{base_name}_{preset_name}.docx")
                doc.save(actual_target)
                out_name = os.path.basename(actual_target)
                quoted_out = urllib.parse.quote(out_name)
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "output_file": out_name,
                    "download_url": f"/api/download/{quoted_out}",
                    "details": preset_res
                })

            # Endpoint 15: PDF Pro & Everyday Toolkit
            elif path == "/api/pdf/tools":
                payload = self._read_body_json()
                action = payload.get("action", "").lower().strip()
                input_param = payload.get("input_path") or payload.get("filepath") or payload.get("file_path") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None

                if action == "merge":
                    inputs: List[str] = [rf for p in payload.get("inputs", []) if (rf := resolve_file_path(p)) is not None]
                    if not inputs:
                        self._send_json({"success": False, "error": "No valid input PDF files specified for merge"}, status=400)
                        return
                    out_path = os.path.join(OUTPUT_DIR, payload.get("output_name", "merged_document.pdf"))
                    res_path = merge_pdfs(inputs, out_path)
                    out_name = os.path.basename(res_path)
                    self._send_json({
                        "success": True,
                        "action": "merge",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}"
                    })
                    return

                elif action in ("images_to_pdf", "images-to-pdf", "img2pdf"):
                    raw_images = payload.get("images") or payload.get("inputs") or ([src] if src else [])
                    resolved_images: List[str] = [rf for p in raw_images if (rf := resolve_file_path(p)) is not None]
                    if not resolved_images:
                        self._send_json({"success": False, "error": "No valid input image files specified for images_to_pdf"}, status=400)
                        return
                    out_name = payload.get("output_name", f"images_bundle_{int(time.time())}.pdf")
                    out_file = os.path.join(OUTPUT_DIR, out_name)
                    details = convert_images_to_pdf(
                        resolved_images,
                        out_file,
                        page_size=payload.get("page_size", "A4"),
                        orientation=payload.get("orientation", "auto"),
                        margin_pt=float(payload.get("margin_pt", 36.0)),
                    )
                    self._send_json({
                        "success": True,
                        "action": "images_to_pdf",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                if not src:
                    self._send_json({"success": False, "error": f"Input PDF not found: {input_param}"}, status=400)
                    return

                base_name = os.path.splitext(os.path.basename(src))[0]

                if action == "split":
                    out_dir = os.path.join(OUTPUT_DIR, f"{base_name}_split")
                    ranges = payload.get("ranges")
                    files = split_pdf(src, out_dir, page_ranges=ranges)
                    self._send_json({
                        "success": True,
                        "action": "split",
                        "generated_files": [os.path.basename(f) for f in files],
                        "count": len(files)
                    })
                    return

                elif action in ("rotate", "clean"):
                    angle = int(payload.get("angle", 90 if action == "rotate" else 0))
                    remove_blanks = bool(payload.get("remove_blanks", True if action == "clean" else False))
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_{action}.pdf")
                    details = rotate_and_clean_pdf(src, out_file, rotation=angle, remove_blank=remove_blanks)
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": action,
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action == "watermark":
                    text = payload.get("watermark_text", "CONFIDENTIAL")
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_watermarked.pdf")
                    res_path = stamp_watermark_pdf(src, out_file, watermark_text=text)
                    out_name = os.path.basename(res_path)
                    self._send_json({
                        "success": True,
                        "action": "watermark",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}"
                    })
                    return

                elif action == "compress":
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_compressed.pdf")
                    details = compress_pdf(src, out_file)
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "compress",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action == "redact":
                    patterns = payload.get("patterns")
                    custom_terms = payload.get("custom_terms") or payload.get("terms")
                    if isinstance(custom_terms, str):
                        custom_terms = [t.strip() for t in custom_terms.split(",") if t.strip()]
                    rect_list = payload.get("rect_list") or payload.get("rects")
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_redacted.pdf")
                    details = redact_sensitive_pdf(src, out_file, patterns=patterns, custom_terms=custom_terms, rect_list=rect_list)
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "redact",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action == "protect":
                    u_pwd = payload.get("user_password") or payload.get("user_pwd")
                    o_pwd = payload.get("owner_password") or payload.get("owner_pwd")
                    allow_print = bool(payload.get("allow_print", True))
                    allow_copy = bool(payload.get("allow_copy", False))
                    allow_modify = bool(payload.get("allow_modify", False))
                    allow_annotate = bool(payload.get("allow_annotate", False))
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_protected.pdf")
                    details = protect_pdf(
                        src,
                        out_file,
                        user_password=u_pwd,
                        owner_password=o_pwd,
                        allow_print=allow_print,
                        allow_copy=allow_copy,
                        allow_modify=allow_modify,
                        allow_annotate=allow_annotate,
                    )
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "protect",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action in ("normalize_a4", "normalize", "bates"):
                    orientation = payload.get("orientation", "auto")
                    margin = float(payload.get("margin", 24.0))
                    bates_prefix = payload.get("bates_prefix") or payload.get("bates")
                    bates_pattern = payload.get("bates_pattern", "{prefix} | Trang {page:02d}/{total:02d}")
                    start_number = int(payload.get("start_number", 1))
                    position = payload.get("position", "bottom-right")
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_normalized_a4.pdf")
                    details = normalize_and_number_pdf(
                        src,
                        out_file,
                        orientation=orientation,
                        margin=margin,
                        bates_prefix=bates_prefix,
                        bates_pattern=bates_pattern,
                        start_number=start_number,
                        position=position,
                    )
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "normalize_a4",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action in ("giap_lai", "giap-lai"):
                    stamp_param = payload.get("stamp_image_path") or payload.get("stamp_path") or payload.get("stamp_filename") or payload.get("stamp_file")
                    stamp_src = resolve_file_path(stamp_param) if stamp_param else None
                    if not stamp_src:
                        self._send_json({"success": False, "error": f"Stamp image file not found: {stamp_param}"}, status=400)
                        return
                    side = payload.get("side", "right")
                    stamp_diameter = float(payload.get("stamp_diameter", 120.0))
                    overlap_pct = float(payload.get("overlap_pct", 8.0))
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_giap_lai.pdf")
                    details = stamp_giap_lai_pdf(
                        src,
                        out_file,
                        stamp_src,
                        side=side,
                        stamp_diameter=stamp_diameter,
                        overlap_pct=overlap_pct,
                    )
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "giap_lai",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action in ("extract_tables", "table_extract", "tables_to_xlsx"):
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_tables.xlsx")
                    details = extract_pdf_tables_to_xlsx(src, out_file)
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "extract_tables",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action in ("reorganize", "reorder"):
                    pages_plan = payload.get("pages_plan") or payload.get("plan") or []
                    if not pages_plan:
                        self._send_json({"success": False, "error": "pages_plan is required for reorganize"}, status=400)
                        return
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_reorganized.pdf")
                    details = reorganize_pdf_pages(src, out_file, pages_plan)
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "reorganize",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action in ("thumbnails", "preview_pages"):
                    max_pages = int(payload.get("max_pages", 100))
                    dpi = int(payload.get("dpi", 120))
                    thumbs_dir = os.path.join(OUTPUT_DIR, "pdf_previews")
                    thumbs = render_pdf_thumbnails(src, thumbs_dir, max_pages=max_pages, dpi=dpi)
                    self._send_json({
                        "success": True,
                        "filename": os.path.basename(src),
                        "total_pages": len(thumbs),
                        "pages": thumbs
                    })
                    return

                elif action in ("export_images", "export_images_zip", "to_images", "pdf_to_images"):
                    dpi = int(payload.get("dpi", 300))
                    img_fmt = payload.get("format", "png")
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_images.zip")
                    details = export_pdf_to_images_zip(src, out_file, dpi=dpi, img_format=img_fmt)
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "export_images_zip",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action in ("sign", "sign_pdf", "stamp_sign"):
                    signatures = payload.get("signatures")
                    if not signatures:
                        sig_obj = {}
                        if "image_base64" in payload:
                            sig_obj["image_base64"] = payload["image_base64"]
                        elif "image_path" in payload:
                            sig_obj["image_path"] = resolve_file_path(payload["image_path"])
                        sig_obj["page"] = payload.get("page", "last")
                        if "pages" in payload:
                            sig_obj["pages"] = payload["pages"]
                        sig_obj["x_pct"] = float(payload.get("x_pct", 0.65))
                        sig_obj["y_pct"] = float(payload.get("y_pct", 0.75))
                        sig_obj["w_pct"] = float(payload.get("w_pct", 0.25))
                        sig_obj["h_pct"] = float(payload.get("h_pct", 0.15))
                        sig_obj["signer_name"] = payload.get("signer_name", "")
                        sig_obj["signer_title"] = payload.get("signer_title", "")
                        sig_obj["signer_date"] = payload.get("signer_date", "")
                        signatures = [sig_obj]
                    else:
                        for s in signatures:
                            if "image_path" in s and s["image_path"]:
                                s["image_path"] = resolve_file_path(s["image_path"])

                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_signed.pdf")
                    details = sign_pdf_with_stamp(src, out_file, signatures)
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "sign_pdf",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                elif action in ("flatten", "sanitize", "flatten_and_sanitize", "flatten_pdf"):
                    flatten_vector = bool(payload.get("flatten_vector", True))
                    strict_raster = bool(payload.get("strict_raster", payload.get("strict", False)))
                    sanitize_metadata = bool(payload.get("sanitize_metadata", True))
                    out_file = os.path.join(OUTPUT_DIR, f"{base_name}_flattened.pdf")
                    details = flatten_and_sanitize_pdf(
                        src,
                        out_file,
                        flatten_vector=flatten_vector,
                        strict_raster=strict_raster,
                        sanitize_metadata=sanitize_metadata
                    )
                    out_name = os.path.basename(out_file)
                    self._send_json({
                        "success": True,
                        "action": "flatten_and_sanitize",
                        "filename": out_name,
                        "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                        "details": details
                    })
                    return

                else:
                    self._send_json({"success": False, "error": f"Unknown pdf-tools action: '{action}'"}, status=400)
                    return

            # Endpoint 15b: Dedicated PDF Page Thumbnails for Visual Page Organizer
            elif path == "/api/pdf/pages/preview":
                payload = self._read_body_json()
                input_param = payload.get("input_path") or payload.get("filepath") or payload.get("file_path") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"PDF file not found: {input_param}"}, status=400)
                    return
                thumbs_dir = os.path.join(OUTPUT_DIR, "pdf_previews")
                thumbs = render_pdf_thumbnails(src, thumbs_dir, max_pages=int(payload.get("max_pages", 100)), dpi=int(payload.get("dpi", 120)))
                self._send_json({
                    "success": True,
                    "filename": os.path.basename(src),
                    "total_pages": len(thumbs),
                    "pages": thumbs
                })
                return

            # Endpoint 15c: Dedicated Images to PDF Endpoint
            elif path == "/api/pdf/images-to-pdf":
                payload = self._read_body_json()
                raw_images = payload.get("images") or payload.get("inputs") or []
                resolved_images = [rf for p in raw_images if (rf := resolve_file_path(p)) is not None]
                if not resolved_images:
                    self._send_json({"success": False, "error": "No valid input images found"}, status=400)
                    return
                out_name = payload.get("output_name") or f"images_bundle_{int(time.time())}.pdf"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                details = convert_images_to_pdf(
                    resolved_images,
                    out_file,
                    page_size=payload.get("page_size", "A4"),
                    orientation=payload.get("orientation", "auto"),
                    margin_pt=float(payload.get("margin_pt", 36.0)),
                )
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            # Endpoint 15d: Dedicated PDF to Images ZIP (300 DPI)
            elif path == "/api/pdf/export-images-zip":
                payload = self._read_body_json()
                input_param = payload.get("input_path") or payload.get("filepath") or payload.get("file_path") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"PDF file not found: {input_param}"}, status=400)
                    return
                base_name = os.path.splitext(os.path.basename(src))[0]
                out_name = f"{base_name}_images.zip"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                details = export_pdf_to_images_zip(src, out_file, dpi=int(payload.get("dpi", 300)), img_format=payload.get("format", "png"))
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            # Endpoint 15e: Dedicated Sign & Annotate PDF
            elif path == "/api/pdf/sign":
                payload = self._read_body_json()
                input_param = payload.get("input_path") or payload.get("filepath") or payload.get("file_path") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"PDF file not found: {input_param}"}, status=400)
                    return
                base_name = os.path.splitext(os.path.basename(src))[0]
                signatures = payload.get("signatures")
                if not signatures:
                    sig_obj = {}
                    if "image_base64" in payload:
                        sig_obj["image_base64"] = payload["image_base64"]
                    elif "image_path" in payload:
                        sig_obj["image_path"] = resolve_file_path(payload["image_path"])
                    sig_obj["page"] = payload.get("page", "last")
                    if "pages" in payload:
                        sig_obj["pages"] = payload["pages"]
                    sig_obj["x_pct"] = float(payload.get("x_pct", 0.65))
                    sig_obj["y_pct"] = float(payload.get("y_pct", 0.75))
                    sig_obj["w_pct"] = float(payload.get("w_pct", 0.25))
                    sig_obj["h_pct"] = float(payload.get("h_pct", 0.15))
                    sig_obj["signer_name"] = payload.get("signer_name", "")
                    sig_obj["signer_title"] = payload.get("signer_title", "")
                    sig_obj["signer_date"] = payload.get("signer_date", "")
                    signatures = [sig_obj]
                else:
                    for s in signatures:
                        if "image_path" in s and s["image_path"]:
                            s["image_path"] = resolve_file_path(s["image_path"])

                out_name = f"{base_name}_signed.pdf"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                details = sign_pdf_with_stamp(src, out_file, signatures)
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            # Endpoint 15f: Dedicated Flatten & Sanitize PDF
            elif path == "/api/pdf/sanitize":
                payload = self._read_body_json()
                input_param = payload.get("input_path") or payload.get("filepath") or payload.get("file_path") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"PDF file not found: {input_param}"}, status=400)
                    return
                base_name = os.path.splitext(os.path.basename(src))[0]
                flatten_vector = bool(payload.get("flatten_vector", True))
                strict_raster = bool(payload.get("strict_raster", payload.get("strict", False)))
                sanitize_metadata = bool(payload.get("sanitize_metadata", True))
                out_name = f"{base_name}_flattened.pdf"
                out_file = os.path.join(OUTPUT_DIR, out_name)
                details = flatten_and_sanitize_pdf(
                    src,
                    out_file,
                    flatten_vector=flatten_vector,
                    strict_raster=strict_raster,
                    sanitize_metadata=sanitize_metadata
                )
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            # Endpoint 16: Excel Print-Ready Page Setup (ERR_XLSX_008)
            elif path == "/api/xlsx/print-setup":
                payload = self._read_body_json()
                src_param = payload.get("file_path") or payload.get("filepath") or payload.get("source_path") or payload.get("filename")
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"Excel file not found: {src_param}"}, status=400)
                    return
                sheet_name = payload.get("sheet_name")
                orientation = payload.get("orientation")

                engine = UniversalExcelEngine(src)
                res = engine.apply_print_setup(sheet_name=sheet_name, orientation=orientation)
                engine.save(src)
                out_name = os.path.basename(src)
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "details": res
                })

            # Endpoint 17: Mail Merge Inspection (Live Token Mapping)
            elif path in ("/api/mail-merge/inspect", "/api/mail_merge/inspect"):
                payload = self._read_body_json()
                tpl_param = payload.get("template") or payload.get("template_path") or payload.get("filepath")
                tpl = resolve_file_path(tpl_param)
                if not tpl:
                    self._send_json({"success": False, "error": f"Template file not found: {tpl_param}"}, status=400)
                    return
                try:
                    from core.docx.mail_merge import inspect_template_tokens
                except (ImportError, ValueError):
                    from core.docx.mail_merge import inspect_template_tokens  # type: ignore
                tokens = inspect_template_tokens(tpl)
                self._send_json({"success": True, "template": os.path.basename(tpl), "tokens": tokens})
                return

            # Endpoint 18: Execute Batch Mail Merge Studio
            elif path in ("/api/mail-merge", "/api/mail_merge"):
                payload = self._read_body_json()
                tpl_param = payload.get("template") or payload.get("template_path")
                tpl = resolve_file_path(tpl_param)
                if not tpl:
                    self._send_json({"success": False, "error": f"Template file not found: {tpl_param}"}, status=400)
                    return

                data_param = payload.get("data") or payload.get("data_source")
                data_source = data_param
                if isinstance(data_param, str):
                    resolved_data = resolve_file_path(data_param)
                    if resolved_data:
                        data_source = resolved_data

                if not data_source:
                    self._send_json({"success": False, "error": "data or data_source is required"}, status=400)
                    return

                out_dir = payload.get("output_dir") or os.path.join(OUTPUT_DIR, "generated_mail_merge")
                pattern = payload.get("pattern")
                zip_name = payload.get("zip")
                zip_path = os.path.join(OUTPUT_DIR, zip_name) if zip_name else None

                try:
                    from core.docx.mail_merge import execute_mail_merge
                except (ImportError, ValueError):
                    from core.docx.mail_merge import execute_mail_merge  # type: ignore

                res = execute_mail_merge(
                    template_path=tpl,
                    data_source=data_source,
                    output_dir=out_dir,
                    filename_pattern=pattern,
                    zip_output=zip_path,
                )
                if res.get("zip_path"):
                    z_name = os.path.basename(res["zip_path"])
                    res["download_url"] = f"/api/download/{urllib.parse.quote(z_name)}"
                self._send_json(res)
                return

            # Endpoint 19: Legal Redline Diff (Track Changes & Comments)
            elif path in ("/api/docx/diff", "/api/docx-diff"):
                payload = self._read_body_json()
                doc_a_param = payload.get("doc_a") or payload.get("doc_a_path") or payload.get("original")
                doc_b_param = payload.get("doc_b") or payload.get("doc_b_path") or payload.get("revised")
                doc_a = resolve_file_path(doc_a_param)
                doc_b = resolve_file_path(doc_b_param)

                if not doc_a:
                    self._send_json({"success": False, "error": f"Original document not found: {doc_a_param}"}, status=400)
                    return
                if not doc_b:
                    self._send_json({"success": False, "error": f"Revised document not found: {doc_b_param}"}, status=400)
                    return

                author = payload.get("author", "AI Reviewer")
                out_name = payload.get("output_name") or f"{os.path.splitext(os.path.basename(doc_b))[0]}_redline.docx"
                out_path = os.path.join(OUTPUT_DIR, out_name)

                try:
                    from core.docx.diff_engine import compare_docx_documents
                except (ImportError, ValueError):
                    from core.docx.diff_engine import compare_docx_documents  # type: ignore

                res = compare_docx_documents(doc_a, doc_b, out_path, author=author)
                qname = urllib.parse.quote(out_name)
                res["download_url"] = f"/api/download/{qname}"
                res["filename"] = out_name
                self._send_json(res)
                return

            # Endpoint 20: Vietnamese OCR Pipeline (Scan & Red-Stamp Filter)
            elif path in ("/api/ocr/scan", "/api/ocr-scan"):
                payload = self._read_body_json()
                input_param = payload.get("file_path") or payload.get("filepath") or payload.get("input_path") or payload.get("filename")
                src = resolve_file_path(input_param)
                if not src:
                    self._send_json({"success": False, "error": f"Input file not found: {input_param}"}, status=400)
                    return

                lang = payload.get("lang", "vie+eng")
                filter_stamp = bool(payload.get("filter_red_stamp", True))
                deskew_flag = bool(payload.get("deskew", True))
                out_format = payload.get("format", "md").lower()

                base_name = os.path.splitext(os.path.basename(src))[0]
                out_ext = ".docx" if out_format == "docx" else ".md"
                out_name = f"{base_name}_ocr{out_ext}"
                out_path = os.path.join(OUTPUT_DIR, out_name)

                try:
                    from core.converter.ocr_pipeline import VietnameseOcrPipeline
                except (ImportError, ValueError):
                    from core.converter.ocr_pipeline import VietnameseOcrPipeline  # type: ignore

                pipeline = VietnameseOcrPipeline(lang=lang, filter_red_stamp=filter_stamp, deskew=deskew_flag)
                res = pipeline.process_document(src, output_path=out_path)
                qname = urllib.parse.quote(out_name)
                res["download_url"] = f"/api/download/{qname}"
                res["filename"] = out_name
                summary = {k: v for k, v in res.items() if k != "pages"}
                self._send_json(summary)
                return

            # Endpoint 21: Unified QA Diagnostic & Audit Pipeline
            elif path in ("/api/qa/audit", "/api/qa/validate"):
                payload = self._read_body_json()
                src_param = payload.get("file_path") or payload.get("filepath") or payload.get("filename")
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"Target document not found: {src_param}"}, status=400)
                    return

                repair_flag = bool(payload.get("repair", False))
                ext = os.path.splitext(src)[1].lower()

                try:
                    from tools.unified_qa_diagnostic import (
                        ExcelDiagnosticEngine,
                        DocxDiagnosticEngine,
                        PptxDiagnosticEngine,
                        DiagramAndConversionEngine,
                    )
                    WordDiagnosticEngine = DocxDiagnosticEngine
                    PowerPointDiagnosticEngine = PptxDiagnosticEngine
                    DiagramDiagnosticEngine = DiagramAndConversionEngine
                except (ImportError, ValueError):
                    import importlib.util
                    qa_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools", "unified_qa_diagnostic.py")
                    spec = importlib.util.spec_from_file_location("unified_qa_diagnostic", qa_path)
                    if spec is not None and spec.loader is not None:
                        mod = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(mod)
                        ExcelDiagnosticEngine = getattr(mod, "ExcelDiagnosticEngine", None)
                        WordDiagnosticEngine = getattr(mod, "DocxDiagnosticEngine", None) or getattr(mod, "WordDiagnosticEngine", None)
                        PowerPointDiagnosticEngine = getattr(mod, "PptxDiagnosticEngine", None) or getattr(mod, "PowerPointDiagnosticEngine", None)
                        DiagramDiagnosticEngine = getattr(mod, "DiagramAndConversionEngine", None) or getattr(mod, "DiagramDiagnosticEngine", None)
                    else:
                        ExcelDiagnosticEngine = None
                        WordDiagnosticEngine = None
                        PowerPointDiagnosticEngine = None
                        DiagramDiagnosticEngine = None

                t0 = time.time()
                issues = []
                metrics = {}
                repaired = False
                repair_actions = []
                out_name = None

                if ext in (".xlsx", ".xlsm"):
                    if ExcelDiagnosticEngine is None:
                        self._send_json({"success": False, "error": "Excel diagnostic engine could not be loaded"}, status=500)
                        return
                    issues, metrics = ExcelDiagnosticEngine.diagnose(src)
                    if repair_flag:
                        out_name = f"repaired_{os.path.basename(src)}"
                        out_path = os.path.join(OUTPUT_DIR, out_name)
                        repaired, repair_actions = ExcelDiagnosticEngine.repair(src, out_path)
                elif ext in (".docx", ".doc"):
                    if WordDiagnosticEngine is None:
                        self._send_json({"success": False, "error": "Word diagnostic engine could not be loaded"}, status=500)
                        return
                    issues, metrics = WordDiagnosticEngine.diagnose(src)
                    if repair_flag:
                        out_name = f"repaired_{os.path.basename(src)}"
                        out_path = os.path.join(OUTPUT_DIR, out_name)
                        repaired, repair_actions = WordDiagnosticEngine.repair(src, out_path)
                elif ext in (".pptx", ".potx"):
                    if PowerPointDiagnosticEngine is None:
                        self._send_json({"success": False, "error": "PowerPoint diagnostic engine could not be loaded"}, status=500)
                        return
                    issues, metrics = PowerPointDiagnosticEngine.diagnose(src)
                    if repair_flag:
                        out_name = f"repaired_{os.path.basename(src)}"
                        out_path = os.path.join(OUTPUT_DIR, out_name)
                        repaired, repair_actions = PowerPointDiagnosticEngine.repair(src, out_path)
                elif ext in (".drawio", ".xml", ".svg"):
                    if DiagramDiagnosticEngine is None:
                        self._send_json({"success": False, "error": "Diagram diagnostic engine could not be loaded"}, status=500)
                        return
                    issues, metrics = DiagramDiagnosticEngine.diagnose_diagram_file(src)
                elif ext == ".pdf":
                    # PDF can be scanned by Security Engine
                    pass
                else:
                    self._send_json({"success": False, "error": f"Unsupported format for QA audit: '{ext}'"}, status=400)
                    return

                # Gate 13: Enterprise Document Security Scan (if requested or by default)
                security_flag = bool(payload.get("security", False))
                deep_cdr_flag = bool(payload.get("deep_cdr", True))
                redact_pii_flag = bool(payload.get("redact_pii", False))
                if security_flag:
                    try:
                        from tools.unified_qa_diagnostic import DocumentSecurityDiagnosticEngine
                        sec_issues, sec_metrics = DocumentSecurityDiagnosticEngine.diagnose(src)
                        issues.extend(sec_issues)
                        metrics["security"] = sec_metrics
                        if repair_flag and (sec_issues or deep_cdr_flag or redact_pii_flag):
                            if not out_name:
                                out_name = f"repaired_{os.path.basename(src)}"
                                out_path = os.path.join(OUTPUT_DIR, out_name)
                            from core.security.sanitizer import sanitize_document
                            san_res = sanitize_document(
                                out_path if (out_path and os.path.exists(out_path)) else src,
                                output_path=out_path,
                                deep_cdr=deep_cdr_flag,
                                redact_pii=redact_pii_flag,
                            )
                            if san_res.get("success"):
                                repaired = True
                                repair_actions.extend(san_res.get("actions_taken", []))
                                metrics["cdr_metrics"] = san_res.get("metrics", {})
                    except Exception as e:
                        metrics["security_error"] = str(e)

                elap = round(time.time() - t0, 3)
                critical_count = sum(1 for i in issues if getattr(i, "severity", "").upper() == "CRITICAL")
                warning_count = sum(1 for i in issues if getattr(i, "severity", "").upper() == "WARNING")

                res = {
                    "success": True,
                    "target_file": os.path.basename(src),
                    "format": ext,
                    "elapsed": elap,
                    "total_issues": len(issues),
                    "critical_count": critical_count,
                    "warning_count": warning_count,
                    "passed": critical_count == 0,
                    "metrics": metrics,
                    "issues": [i.to_dict() if hasattr(i, "to_dict") else vars(i) for i in issues],
                    "repaired": repaired,
                    "repair_actions": repair_actions,
                    "security_scanned": security_flag,
                }
                if repaired and out_name:
                    res["download_url"] = f"/api/download/{urllib.parse.quote(out_name)}"
                    res["repaired_file"] = out_name

                self._send_json(res)
                return

            # Endpoint 21b: Document Security Audit & Sanitizer (Cloudflare Harness)
            elif path == "/api/security/scan":
                payload = self._read_body_json()
                src_param = payload.get("file_path") or payload.get("filepath") or payload.get("filename")
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"Target document not found: {src_param}"}, status=400)
                    return

                try:
                    from core.security.document_auditor import DocumentSecurityAuditor
                    auditor = DocumentSecurityAuditor()
                    report = auditor.audit_file(src)
                    self._send_json({"success": True, "report": report.to_dict()})
                except Exception as e:
                    self._send_json({"success": False, "error": str(e)}, status=500)
                return

            elif path == "/api/security/sanitize":
                payload = self._read_body_json()
                src_param = payload.get("file_path") or payload.get("filepath") or payload.get("filename")
                src = resolve_file_path(src_param)
                if not src:
                    self._send_json({"success": False, "error": f"Target document not found: {src_param}"}, status=400)
                    return

                defuse = bool(payload.get("defuse_formulas", True))
                strip_meta = bool(payload.get("strip_metadata", True))
                unhide = bool(payload.get("unhide_sheets", True))
                deep_cdr = bool(payload.get("deep_cdr", True))
                redact_pii = bool(payload.get("redact_pii", False))
                pii_types = payload.get("pii_types")

                base_name = os.path.basename(src)
                out_name = f"sanitized_{base_name}"
                out_path = os.path.join(OUTPUT_DIR, out_name)

                try:
                    from core.security.sanitizer import DocumentSanitizer
                    sanitizer = DocumentSanitizer()
                    res = sanitizer.sanitize(
                        src,
                        output_path=out_path,
                        defuse_formulas=defuse,
                        strip_metadata=strip_meta,
                        unhide_sheets=unhide,
                        deep_cdr=deep_cdr,
                        redact_pii=redact_pii,
                        pii_types=pii_types,
                    )
                    res["download_url"] = f"/api/download/{urllib.parse.quote(out_name)}"
                    res["sanitized_file"] = out_name
                    self._send_json(res)
                except Exception as e:
                    self._send_json({"success": False, "error": str(e)}, status=500)
                return

            # Endpoint 22: Digital Yearbook & Photobook Studio
            elif path == "/api/yearbook/auto-layout":
                payload = self._read_body_json()
                raw_photos = payload.get("photos", [])
                resolved_photos: list[str] = []
                for p in raw_photos:
                    p_str = p.get("filepath") or p.get("photo") or p.get("url") if isinstance(p, dict) else p
                    res_p = resolve_file_path(p_str) if isinstance(p_str, str) else None
                    if res_p:
                        resolved_photos.append(res_p)

                title = payload.get("title") or payload.get("album_title") or "KỶ YẾU TỐT NGHIỆP 2026"
                subtitle = payload.get("subtitle") or payload.get("album_subtitle") or "Hành Trình Rực Rỡ & Những Năm Tháng Thanh Xuân"
                author = payload.get("author") or payload.get("album_author") or "Biện Trung Tín"
                degree = payload.get("degree") or payload.get("album_degree") or "Cử Nhân Kỹ Thuật Phần Mềm"
                date_str = payload.get("date") or payload.get("album_date") or "Tháng 11 / 2026"
                preset = payload.get("preset") or payload.get("page_size") or "photobook_square_25x25"
                bleed_mm = float(payload.get("bleed_mm", 3.0))

                spec = YearbookLayoutEngine.auto_layout(
                    photo_paths=resolved_photos,
                    title=title,
                    subtitle=subtitle,
                    author=author,
                    degree=degree,
                    date_str=date_str,
                    preset_key=preset,
                    bleed_mm=bleed_mm,
                )
                if isinstance(spec, dict):
                    if "spreads" not in spec and "pages" in spec:
                        spec["spreads"] = spec["pages"]
                    if "page_size" not in spec and "preset" in spec:
                        spec["page_size"] = spec["preset"]
                self._send_json({"success": True, "spec": spec})
                return

            elif path == "/api/yearbook/render-preview":
                payload = self._read_body_json()
                spec = payload.get("spec")
                if not spec or not isinstance(spec, dict):
                    self._send_json({"success": False, "error": "spec object is required"}, status=400)
                    return

                spread_idx = int(payload.get("spread_index", 0))
                dpi = int(payload.get("dpi", 150))
                thumb_fname = f"yearbook_spread_{spread_idx:02d}_{int(time.time()*1000)}.png"
                thumb_path = os.path.join(OUTPUT_DIR, thumb_fname)

                YearbookPressEngine.render_spread_thumbnail_png(spec, spread_idx, thumb_path, dpi=dpi)
                self._send_json({
                    "success": True,
                    "filename": thumb_fname,
                    "preview_url": f"/api/preview/{thumb_fname}",
                    "spread_index": spread_idx,
                })
                return

            elif path == "/api/yearbook/export":
                payload = self._read_body_json()
                spec = payload.get("spec")
                if not spec or not isinstance(spec, dict):
                    self._send_json({"success": False, "error": "spec object is required"}, status=400)
                    return

                export_type = payload.get("export_type", "handoff_zip").lower()
                import re
                clean_title = re.sub(r'[^\w\d\-_]', '_', spec.get("title", "Yearbook"))
                t0 = time.time()

                if export_type in ("print_pdf", "print_ready", "pdf_print"):
                    out_name = f"{clean_title}_Print_Ready_300DPI.pdf"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    YearbookPressEngine.export_print_ready_pdf(spec, out_path)
                elif export_type in ("spread_pdf", "spread_preview", "pdf_spread"):
                    out_name = f"{clean_title}_Spread_Preview.pdf"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    YearbookPressEngine.export_spread_preview_pdf(spec, out_path)
                elif export_type in ("digital_html", "standalone_html", "html"):
                    out_name = f"{clean_title}_Interactive_Digital_Yearbook.html"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    YearbookHandoffPackager.export_standalone_html(spec, out_path)
                else:
                    out_name = f"{clean_title}_Designer_Handoff_Bundle.zip"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    YearbookHandoffPackager.export_handoff_bundle(spec, out_path)

                elap = round(time.time() - t0, 2)
                file_sz = os.path.getsize(out_path) if os.path.isfile(out_path) else 0
                try:
                    record_history("yearbook_export", f"{clean_title} ({export_type})", out_name, out_path, file_sz)
                except Exception:
                    pass

                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "file_size_bytes": file_sz,
                    "elapsed": elap,
                    "export_type": export_type,
                })
                return

            # Endpoint 23: v4.6.0 Enterprise Upgrade Suite Endpoints
            elif path == "/api/pdf/crypto-sign":
                from core.converter.pdf_toolkit import crypto_sign_pdf
                payload = self._read_body_json()
                input_param = payload.get("input_path") or payload.get("file_path") or payload.get("filepath") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"PDF file not found: {input_param}"}, status=400)
                    return
                base_name = os.path.splitext(os.path.basename(src))[0]
                out_name = f"{base_name}_crypto_signed.pdf"
                out_path = os.path.join(OUTPUT_DIR, out_name)
                details = crypto_sign_pdf(
                    src,
                    out_path,
                    signer_name=payload.get("signer_name", "Antigravity Authority"),
                    reason=payload.get("reason", "Phê duyệt & Xác thực tài liệu (PAdES/PKCS#7)"),
                    location=payload.get("location", "Việt Nam"),
                    contact_info=payload.get("contact_info", ""),
                    pfx_path=resolve_file_path(payload.get("pfx_path")) if payload.get("pfx_path") else None,
                    pfx_password=payload.get("pfx_password"),
                    stamp_image=resolve_file_path(payload.get("stamp_image")) if payload.get("stamp_image") else None,
                    page=payload.get("page", "last"),
                )
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            elif path in ("/api/pdf/pdfa", "/api/pdf/to-pdfa"):
                from core.converter.pdf_toolkit import convert_to_pdfa
                payload = self._read_body_json()
                input_param = payload.get("input_path") or payload.get("file_path") or payload.get("filepath") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"PDF file not found: {input_param}"}, status=400)
                    return
                base_name = os.path.splitext(os.path.basename(src))[0]
                part_num = int(payload.get("part", 2))
                conf_lvl = str(payload.get("conformance", "B")).upper()
                out_name = f"{base_name}_pdfa{part_num}{conf_lvl.lower()}.pdf"
                out_path = os.path.join(OUTPUT_DIR, out_name)
                details = convert_to_pdfa(src, out_path, part=part_num, conformance=conf_lvl)
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            elif path == "/api/xlsx/to-brief":
                from core.xlsx.reporting import excel_to_brief
                payload = self._read_body_json()
                input_param = payload.get("xlsx_path") or payload.get("file_path") or payload.get("filepath") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"Excel file not found: {input_param}"}, status=400)
                    return
                fmt = str(payload.get("target_format", payload.get("format", "docx"))).lower().lstrip(".")
                base_name = os.path.splitext(os.path.basename(src))[0]
                out_name = f"{base_name}_brief.{fmt}"
                out_path = os.path.join(OUTPUT_DIR, out_name)
                details = excel_to_brief(
                    src,
                    output_path=out_path,
                    target_format=fmt,
                    title=payload.get("title"),
                    include_kpi_summary=bool(payload.get("include_kpi_summary", True)),
                )
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            elif path == "/api/xlsx/consolidate":
                from core.xlsx.reporting import consolidate_workbooks
                payload = self._read_body_json()
                raw_inputs = payload.get("inputs") or payload.get("input_paths") or []
                resolved = [rf for p in raw_inputs if (rf := resolve_file_path(p)) is not None]
                if not resolved:
                    self._send_json({"success": False, "error": "No valid input Excel files found"}, status=400)
                    return
                out_name = payload.get("output_name") or f"master_consolidated_{int(time.time())}.xlsx"
                out_path = os.path.join(OUTPUT_DIR, out_name)
                details = consolidate_workbooks(
                    resolved,
                    out_path,
                    mode=str(payload.get("mode", "append_rows")),
                    strict_schema=bool(payload.get("strict_schema", True)),
                )
                self._send_json({
                    "success": True,
                    "filename": out_name,
                    "download_url": f"/api/download/{urllib.parse.quote(out_name)}",
                    "details": details,
                })
                return

            elif path == "/api/pipeline/run":
                from core.pipeline import run_pipeline
                payload = self._read_body_json()
                input_param = payload.get("input_path") or payload.get("file_path") or payload.get("filepath") or payload.get("filename")
                src = resolve_file_path(input_param) if input_param else None
                if not src:
                    self._send_json({"success": False, "error": f"Input file not found: {input_param}"}, status=400)
                    return
                spec = payload.get("pipeline_spec") or payload.get("spec")
                if not spec:
                    self._send_json({"success": False, "error": "pipeline_spec is required"}, status=400)
                    return
                res = run_pipeline(src, spec)
                out_file = res.get("output_file", "")
                out_name = os.path.basename(out_file)
                if out_file and os.path.dirname(os.path.abspath(out_file)) != os.path.abspath(OUTPUT_DIR):
                    dest_in_out = os.path.join(OUTPUT_DIR, out_name)
                    shutil.copy2(out_file, dest_in_out)
                res["filename"] = out_name
                res["download_url"] = f"/api/download/{urllib.parse.quote(out_name)}"
                self._send_json(res)
                return

            elif path == "/api/diagrams/from-code":
                from core.diagrams.code_to_diagram import generate_diagram_from_code
                payload = self._read_body_json()
                src_code = payload.get("source_code") or payload.get("source_path") or payload.get("code")
                if not src_code:
                    self._send_json({"success": False, "error": "source_code or source_path is required"}, status=400)
                    return
                resolved_src = resolve_file_path(src_code) if isinstance(src_code, str) and len(src_code) < 260 else None
                actual_src = resolved_src or src_code
                ts = int(time.time() * 1000)
                png_name = f"code_diag_{ts}.png"
                drawio_name = f"code_diag_{ts}.drawio"
                svg_name = f"code_diag_{ts}.svg"
                res = generate_diagram_from_code(
                    source_path=actual_src,
                    diagram_type=str(payload.get("diagram_type", "auto")),
                    output_png_path=os.path.join(OUTPUT_DIR, png_name),
                    output_drawio_path=os.path.join(OUTPUT_DIR, drawio_name),
                    output_svg_path=os.path.join(OUTPUT_DIR, svg_name),
                )
                res["png_download_url"] = f"/api/download/{urllib.parse.quote(png_name)}"
                res["drawio_download_url"] = f"/api/download/{urllib.parse.quote(drawio_name)}"
                res["svg_download_url"] = f"/api/download/{urllib.parse.quote(svg_name)}"
                self._send_json(res)
                return

            else:
                self._send_json({"error": f"POST endpoint not found: {path}"}, status=404)

        except ValueError as exc:
            err_msg = str(exc)
            status_code = 413 if "Payload Too Large" in err_msg else 400
            self._send_json({"success": False, "error": err_msg}, status=status_code)
        except Exception as exc:
            import traceback
            traceback.print_exc()
            self._send_json({"success": False, "error": str(exc)}, status=500)

    def _handle_upload(self, content_type: str) -> None:
        """Handles binary file upload via multipart/form-data or application/octet-stream."""
        import base64
        import re

        try:
            cleanup_old_temp_files(max_age_seconds=7200)
        except Exception:
            pass

        content_length = int(self.headers.get("Content-Length", 0))
        if content_length > MAX_UPLOAD_SIZE_BYTES:
            self._send_json({
                "success": False,
                "error": f"Payload Too Large: {content_length} bytes exceeds maximum allowed {MAX_UPLOAD_SIZE_MB}MB"
            }, status=413)
            return

        if "multipart/form-data" in content_type:
            boundary = content_type.split("boundary=")[-1].strip()
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)

            parts = body.split(b"--" + boundary.encode())
            saved_files = []

            for part in parts:
                if b'filename="' in part:
                    header_part, file_bytes = part.split(b"\r\n\r\n", 1)
                    # Strip trailing \r\n
                    file_bytes = file_bytes.rstrip(b"\r\n")

                    match = re.search(r'filename="([^"]+)"', header_part.decode("utf-8", errors="replace"))
                    fname = match.group(1) if match else f"upload_{int(time.time())}.bin"
                    fname = os.path.basename(fname)
                    saved_path = os.path.join(UPLOAD_DIR, fname)
                    try:
                        with open(saved_path, "wb") as f:
                            f.write(file_bytes)
                    except (PermissionError, OSError):
                        stem, ext = os.path.splitext(fname)
                        fname = f"{stem}_{int(time.time()*1000)}{ext}"
                        saved_path = os.path.join(UPLOAD_DIR, fname)
                        with open(saved_path, "wb") as f:
                            f.write(file_bytes)

                    saved_files.append({
                        "file_path": os.path.abspath(saved_path),
                        "filename": fname,
                        "size": len(file_bytes),
                        "preview_url": f"/api/preview/{urllib.parse.quote(fname)}"
                    })

            first_item = saved_files[0] if saved_files else None
            first_path = first_item["file_path"] if first_item else None
            self._send_json({
                "success": True,
                "filepath": first_path,
                "file_path": first_path,
                "files": saved_files,
                "first_file": first_item
            })
        else:
            # Fallback for base64 JSON payload
            payload = self._read_body_json()
            fname = os.path.basename(payload.get("filename", f"upload_{int(time.time())}.bin"))
            b64_content = payload.get("data", "")
            file_bytes = base64.b64decode(b64_content)
            saved_path = os.path.join(UPLOAD_DIR, fname)
            try:
                with open(saved_path, "wb") as f:
                    f.write(file_bytes)
            except (PermissionError, OSError):
                stem, ext = os.path.splitext(fname)
                fname = f"{stem}_{int(time.time()*1000)}{ext}"
                saved_path = os.path.join(UPLOAD_DIR, fname)
                with open(saved_path, "wb") as f:
                    f.write(file_bytes)

            saved_abs = os.path.abspath(saved_path)
            self._send_json({
                "success": True,
                "filepath": saved_abs,
                "file_path": saved_abs,
                "filename": fname,
                "output_file": fname,
                "size": len(file_bytes),
                "preview_url": f"/api/preview/{urllib.parse.quote(fname)}"
            })

    def _serve_file(self, filename: str, as_attachment: bool = False) -> None:
        """Safely streams a file strictly from either UPLOAD_DIR or OUTPUT_DIR, preventing path traversal."""
        import re
        safe_fname = os.path.basename(filename)
        if not safe_fname or safe_fname in (".", ".."):
            self._send_json({"error": "Invalid file path"}, status=400)
            return

        candidate_1 = os.path.join(OUTPUT_DIR, safe_fname)
        candidate_2 = os.path.join(UPLOAD_DIR, safe_fname)
        candidate_3 = os.path.join(OUTPUT_DIR, "pdf_previews", safe_fname)

        target_file = None
        if os.path.isfile(candidate_1):
            target_file = candidate_1
        elif os.path.isfile(candidate_2):
            target_file = candidate_2
        elif os.path.isfile(candidate_3):
            target_file = candidate_3

        if not target_file:
            self._send_json({"error": f"File not found: {safe_fname}"}, status=404)
            return

        # Double check containment inside allowed directories
        allowed_dirs = [os.path.abspath(OUTPUT_DIR), os.path.abspath(UPLOAD_DIR), os.path.abspath(os.path.join(OUTPUT_DIR, "pdf_previews"))]
        target_abs = os.path.abspath(target_file)
        if not any(os.path.commonpath([d, target_abs]) == d for d in allowed_dirs):
            self._send_json({"error": "Forbidden path access"}, status=403)
            return

        ext = os.path.splitext(safe_fname)[1].lower()
        mime_types = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".svg": "image/svg+xml",
            ".pdf": "application/pdf",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".json": "application/json",
            ".md": "text/markdown; charset=utf-8",
            ".txt": "text/plain; charset=utf-8",
            ".html": "text/html; charset=utf-8",
            ".zip": "application/zip",
        }
        mime = mime_types.get(ext, "application/octet-stream")

        with open(target_file, "rb") as f:
            content = f.read()

        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        if as_attachment:
            safe_ascii = re.sub(r'[^\x20-\x7E]', '_', safe_fname).replace('"', '')
            encoded_fn = urllib.parse.quote(safe_fname, encoding='utf-8')
            self.send_header("Content-Disposition", f'attachment; filename="{safe_ascii}"; filename*=UTF-8\'\'{encoded_fn}')
        self.end_headers()
        self.wfile.write(content)

    def _serve_static_file(self, rel_path: str) -> None:
        """
        Safely streams static web frontend files (index.html, assets/*, etc.).
        Prioritizes the production Vite build directory ('dist/'), then falls back
        to the root workspace directory, and implements SPA fallback to dist/index.html.
        Strictly blocks source code (.py), dotfiles (.env, .git), and backend configs.
        """
        clean_rel = rel_path.lstrip("/").replace("\\", "/")
        if not clean_rel or clean_rel == "":
            clean_rel = "index.html"

        norm_rel = os.path.normpath(clean_rel)
        if norm_rel.startswith("..") or os.path.isabs(norm_rel):
            self._send_json({"error": "Forbidden path traversal attempt"}, status=403)
            return

        # Security check: Block sensitive dotfiles and backend source files
        base_name = os.path.basename(norm_rel).lower()
        ext = os.path.splitext(norm_rel)[1].lower()

        FORBIDDEN_FILES = {".env", ".gitignore", ".dockerignore", "dockerfile", "requirements.txt", "pyproject.toml"}
        FORBIDDEN_EXTENSIONS = {".py", ".pyc", ".env", ".git", ".log", ".bak", ".tmp", ".bat", ".sh", ".cmd", ".jar"}
        ALLOWED_STATIC_EXTENSIONS = {
            ".html", ".css", ".js", ".mjs", ".json", ".png", ".jpg", ".jpeg",
            ".svg", ".ico", ".woff", ".woff2", ".ttf", ".webp", ".txt", ".map"
        }

        if base_name.startswith(".") or base_name in FORBIDDEN_FILES or ext in FORBIDDEN_EXTENSIONS:
            self._send_json({"error": "Forbidden: Access to sensitive file is prohibited"}, status=403)
            return

        if ext and ext not in ALLOWED_STATIC_EXTENSIONS:
            self._send_json({"error": "Forbidden: Unsupported static file type"}, status=403)
            return

        target_file = None

        # 1. First priority: Production build in dist/
        if os.path.isdir(DIST_DIR):
            dist_candidate = os.path.join(DIST_DIR, norm_rel)
            if os.path.isfile(dist_candidate):
                target_file = dist_candidate
            elif norm_rel == "index.html" and os.path.isfile(os.path.join(DIST_DIR, "index.html")):
                target_file = os.path.join(DIST_DIR, "index.html")

        # 2. Second priority: Root workspace (development files) - only if in allowed static extensions
        if not target_file:
            base_candidate = os.path.join(BASE_DIR, norm_rel)
            if os.path.isfile(base_candidate) and ext in ALLOWED_STATIC_EXTENSIONS:
                target_file = base_candidate

        # 3. Third priority: Single Page Application (SPA) fallback to dist/index.html
        if not target_file and os.path.isdir(DIST_DIR) and not norm_rel.startswith("api"):
            spa_index = os.path.join(DIST_DIR, "index.html")
            if os.path.isfile(spa_index):
                target_file = spa_index

        if not target_file or not os.path.isfile(target_file):
            self._send_json({"error": f"Static asset not found: {clean_rel}"}, status=404)
            return

        # Ensure target_file is within BASE_DIR or DIST_DIR
        target_abs = os.path.abspath(target_file)
        if os.path.commonpath([os.path.abspath(BASE_DIR), target_abs]) != os.path.abspath(BASE_DIR):
            self._send_json({"error": "Forbidden path access"}, status=403)
            return

        target_ext = os.path.splitext(target_file)[1].lower()
        mime_types = {
            ".html": "text/html; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".mjs": "application/javascript; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".json": "application/json; charset=utf-8",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".svg": "image/svg+xml",
            ".ico": "image/x-icon",
            ".woff": "font/woff",
            ".woff2": "font/woff2",
            ".ttf": "font/ttf",
            ".webp": "image/webp",
        }
        mime = mime_types.get(target_ext, "application/octet-stream")

        with open(target_file, "rb") as f:
            content = f.read()

        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def _handle_job_stream(self, job_id: str) -> None:
        """Streams Server-Sent Events (SSE) for background job progress tracking."""
        job = global_job_manager.get_job(job_id)
        if not job:
            self._send_json({"error": f"Job not found: {job_id}"}, status=404)
            return

        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache, no-transform")
        self.send_header("Connection", "keep-alive")
        self.send_header("X-Accel-Buffering", "no")
        self.end_headers()

        initial_data = f"data: {json.dumps(job, ensure_ascii=False)}\n\n".encode("utf-8")
        try:
            self.wfile.write(initial_data)
            self.wfile.flush()
        except Exception:
            return

        if job["status"] in ("completed", "failed"):
            return

        import queue
        q = global_job_manager.subscribe(job_id)
        if not q:
            return

        try:
            while True:
                try:
                    event = q.get(timeout=25)
                    data_str = f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
                    self.wfile.write(data_str.encode("utf-8"))
                    self.wfile.flush()
                    if event.get("status") in ("completed", "failed"):
                        break
                except queue.Empty:
                    self.wfile.write(b": keep-alive\n\n")
                    self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError, Exception):
            pass
        finally:
            global_job_manager.unsubscribe(job_id, q)

    def _dispatch_job_submission(self, action: str, p: Dict[str, Any]) -> str:
        """Dispatches job payload to thread pool executor with progress reporting."""
        def _runner(payload: Dict[str, Any], cb: Any) -> Any:
            act = action.lower()
            if act in ("convert", "universal-convert"):
                src_param = (
                    payload.get("source_path")
                    or payload.get("input_path")
                    or payload.get("filepath")
                    or payload.get("file_path")
                    or payload.get("filename")
                )
                src = resolve_file_path(src_param)
                if not src:
                    raise FileNotFoundError(f"Source file not found: {src_param}")
                dest_ext = payload.get("target_ext", ".docx")
                pages = payload.get("page_range")
                base_name = os.path.splitext(os.path.basename(src))[0]
                target_path = os.path.join(OUTPUT_DIR, f"{base_name}_converted{dest_ext}")
                ok, res, elap = convert_universal(
                    source_path=src,
                    target_path=target_path,
                    target_ext=dest_ext,
                    page_range=pages,
                    progress_callback=cb,
                )
                if not ok:
                    raise RuntimeError(res)
                out_name = os.path.basename(res)
                file_sz = os.path.getsize(res) if os.path.isfile(res) else 0
                record_history("convert", os.path.basename(src), out_name, res, file_sz)
                return {
                    "filename": out_name,
                    "download_url": f"/api/download/{out_name}",
                    "preview_url": f"/api/preview/{out_name}",
                    "elapsed": elap,
                }

            elif act in ("docx-to-pptx", "pptx-build"):
                src_param = payload.get("docx_path") or payload.get("source_path") or payload.get("filename")
                src = resolve_file_path(src_param)
                if not src:
                    raise FileNotFoundError(f"File not found: {src_param}")
                theme = payload.get("theme", "corporate_blue")
                base_name = os.path.splitext(os.path.basename(src))[0]
                target_pptx = os.path.join(OUTPUT_DIR, f"{base_name}_presentation.pptx")
                actual_path, _ = convert_docx_to_pptx(src, target_pptx, theme=theme, progress_callback=cb)
                preview_list = export_pptx_slides(actual_path, os.path.join(OUTPUT_DIR, "previews"))
                out_name = os.path.basename(actual_path)
                file_sz = os.path.getsize(actual_path) if os.path.isfile(actual_path) else 0
                record_history("docx_to_pptx", os.path.basename(src), out_name, actual_path, file_sz)
                return {
                    "filename": out_name,
                    "download_url": f"/api/download/{out_name}",
                    "slides_count": len(preview_list),
                    "previews": preview_list,
                }

            elif act == "pdf-to-pptx":
                src_param = payload.get("pdf_path") or payload.get("source_path") or payload.get("filename")
                src = resolve_file_path(src_param)
                if not src:
                    raise FileNotFoundError(f"File not found: {src_param}")
                theme = payload.get("theme", "corporate_blue")
                pages = payload.get("page_range")
                base_name = os.path.splitext(os.path.basename(src))[0]
                target_pptx = os.path.join(OUTPUT_DIR, f"{base_name}_pdf_presentation.pptx")
                actual_path, _ = convert_pdf_to_pptx(src, target_pptx, theme=theme, page_range=pages, progress_callback=cb)
                preview_list = export_pptx_slides(actual_path, os.path.join(OUTPUT_DIR, "previews"))
                out_name = os.path.basename(actual_path)
                file_sz = os.path.getsize(actual_path) if os.path.isfile(actual_path) else 0
                record_history("pdf_to_pptx", os.path.basename(src), out_name, actual_path, file_sz)
                return {
                    "filename": out_name,
                    "download_url": f"/api/download/{out_name}",
                    "slides_count": len(preview_list),
                    "previews": preview_list,
                }

            elif act in ("render-diagram", "diagram-render"):
                code = payload.get("code") or payload.get("mermaid_code") or payload.get("plantuml_code") or payload.get("xml_content") or ""
                lang = payload.get("engine") or payload.get("type", "mermaid")
                try:
                    from core.diagrams.bridge import render_diagram_code_to_png
                except (ImportError, ValueError):
                    from core.diagrams.bridge import render_diagram_code_to_png

                cb(20.0, f"Đang khởi động engine {lang}...")
                diag_res = render_diagram_code_to_png(code, lang, scale=payload.get("scale", 2))
                if not diag_res or not diag_res.get("png_path"):
                    raise RuntimeError(f"Lỗi render sơ đồ {lang}")
                png_path = diag_res["png_path"]
                out_name = os.path.basename(png_path)
                dest_path = os.path.join(OUTPUT_DIR, out_name)
                import shutil
                shutil.copy2(png_path, dest_path)
                cb(100.0, "Hoàn tất render sơ đồ!")
                file_sz = os.path.getsize(dest_path) if os.path.isfile(dest_path) else 0
                record_history("diagram_render", f"diag_{lang}", out_name, dest_path, file_sz)
                return {
                    "filename": out_name,
                    "download_url": f"/api/download/{out_name}",
                    "preview_url": f"/api/preview/{out_name}",
                    "dimensions_px": diag_res["dimensions_px"],
                }

            elif act in ("yearbook_export", "yearbook-export"):
                spec = payload.get("spec")
                if not spec:
                    raise ValueError("spec is required for yearbook export")
                cb(10.0, "Đang khởi tạo động cơ xuất bản kỷ yếu...")
                export_type = payload.get("export_type", "handoff_zip").lower()
                import re
                clean_title = re.sub(r'[^\w\d\-_]', '_', spec.get("title", "Yearbook"))
                if export_type in ("print_pdf", "print_ready", "pdf_print"):
                    out_name = f"{clean_title}_Print_Ready_300DPI.pdf"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    cb(30.0, "Đang kết xuất PDF đơn trang chuẩn in 300 DPI...")
                    YearbookPressEngine.export_print_ready_pdf(spec, out_path)
                elif export_type in ("spread_pdf", "spread_preview", "pdf_spread"):
                    out_name = f"{clean_title}_Spread_Preview.pdf"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    cb(30.0, "Đang kết xuất PDF xem trước mở phẳng 2 trang...")
                    YearbookPressEngine.export_spread_preview_pdf(spec, out_path)
                else:
                    out_name = f"{clean_title}_Designer_Handoff_Bundle.zip"
                    out_path = os.path.join(OUTPUT_DIR, out_name)
                    cb(30.0, "Đang đóng gói hồ sơ Designer Handoff 5-trong-1...")
                    YearbookHandoffPackager.export_handoff_bundle(spec, out_path)

                cb(100.0, "Đã hoàn tất xuất bản Kỷ yếu!")
                file_sz = os.path.getsize(out_path) if os.path.isfile(out_path) else 0
                record_history("yearbook_export", f"{clean_title} ({export_type})", out_name, out_path, file_sz)
                return {
                    "filename": out_name,
                    "download_url": f"/api/download/{out_name}",
                    "file_size": file_sz,
                }

            else:
                raise ValueError(f"Hành động không hợp lệ: {action}")

        return global_job_manager.submit_job(action, p, _runner)


StudioAPIHandler = OfficeApiHandler


def is_server_already_running(port: int = PORT, host: str = "127.0.0.1") -> bool:
    import urllib.request
    try:
        check_host = "127.0.0.1" if host in ("0.0.0.0", "") else host
        req = urllib.request.Request(f"http://{check_host}:{port}/api/health", headers={"User-Agent": "Antigravity-Probe"})
        with urllib.request.urlopen(req, timeout=1.5) as res:
            if res.status == 200:
                return True
    except Exception:
        pass
    return False


def run_server(port: Optional[int] = None, host: Optional[str] = None) -> None:
    if port is None:
        port = int(os.environ.get("PORT", PORT))
    if host is None:
        host = os.environ.get("HOST", HOST)

    check_host = "127.0.0.1" if host in ("0.0.0.0", "") else host
    if is_server_already_running(port, check_host):
        print(f"✅ [Office API Bridge] Server is already running on http://{check_host}:{port} (healthy).")
        try:
            while is_server_already_running(port, check_host):
                time.sleep(2)
        except KeyboardInterrupt:
            return

    server_address = (host, port)
    # Use ThreadingHTTPServer for non-blocking concurrent requests
    handler_class = getattr(http.server, "ThreadingHTTPServer", http.server.HTTPServer)
    httpd = handler_class(server_address, OfficeApiHandler)
    display_host = "127.0.0.1" if host in ("0.0.0.0", "") else host
    print(f"🚀 [Office API Bridge] Running on http://{display_host}:{port} (bound to {host}:{port}, pid {os.getpid()})")
    print(f"   Uploads directory: {UPLOAD_DIR}")
    print(f"   Outputs directory: {OUTPUT_DIR}")
    if os.path.isdir(DIST_DIR):
        print(f"   Static bundle: Serving compiled Vite assets from {DIST_DIR}")
    srv_thread = threading.Thread(target=httpd.serve_forever, name="OfficeApiBridgeThread", daemon=True)
    srv_thread.start()
    try:
        while srv_thread.is_alive():
            time.sleep(0.25)
    except KeyboardInterrupt:
        print("\n[Office API Bridge] Shutting down gracefully.")
    finally:
        try:
            httpd.shutdown()
            httpd.server_close()
        except Exception:
            pass


if __name__ == "__main__":
    port_arg = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else int(os.environ.get("PORT", PORT))
    host_arg = sys.argv[2] if len(sys.argv) > 2 else os.environ.get("HOST", HOST)
    run_server(port_arg, host_arg)
