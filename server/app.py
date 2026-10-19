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

# Import core handlers
try:
    from core.converter.engine import convert_universal
    from core.pptx.docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx
    from core.pptx.writer import THEMES, write_pptx_from_spec
    from core.docx.writer import patch_docx_template, write_docx
    from core.xlsx.writer import mutate_template_excel
    from core.xlsx.validator import validate_excel_sheet
    from core.xlsx.reader import read_xlsx
    from .db import record_history, get_recent_history, delete_history_item, clear_history
    from core.docx.math import parse_inline_math, add_math_to_paragraph
    from core.docx.presets import apply_preset_to_document, load_preset, list_available_presets, build_administrative_header
    from core.converter.pdf_toolkit import merge_pdfs, split_pdf, rotate_and_clean_pdf, stamp_watermark_pdf, compress_pdf
    from core.xlsx.writer import apply_worksheet_print_setup, UniversalExcelEngine
except (ImportError, ValueError):
    from converter_engine import convert_universal  # type: ignore
    from docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx  # type: ignore
    from pptx_writer import THEMES, write_pptx_from_spec  # type: ignore
    from docx_writer import patch_docx_template, write_docx  # type: ignore
    from xlsx_writer import mutate_template_excel  # type: ignore
    from xlsx_validator import validate_excel_sheet  # type: ignore
    from db_manager import record_history, get_recent_history, delete_history_item, clear_history  # type: ignore
    from docx_math import parse_inline_math, add_math_to_paragraph  # type: ignore
    from docx_presets import apply_preset_to_document, load_preset, list_available_presets, build_administrative_header  # type: ignore
    from pdf_toolkit import merge_pdfs, split_pdf, rotate_and_clean_pdf, stamp_watermark_pdf, compress_pdf  # type: ignore
    from xlsx_writer import apply_worksheet_print_setup, UniversalExcelEngine  # type: ignore

try:
    from core.job_manager import global_job_manager
    from core.env_doctor import run_doctor
except (ImportError, ValueError):
    from job_manager import global_job_manager  # type: ignore
    from env_doctor import run_doctor  # type: ignore

PORT = int(os.environ.get("PORT", 8000))
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOAD_DIR = os.path.join(BASE_DIR, ".web_uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, ".web_outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


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
                target_indices = [int(x.strip()) for x in str(slides_arg).split(",") if x.strip().isdigit()]

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
    """Safely resolves an input file path against filesystem, UPLOAD_DIR, OUTPUT_DIR, and CWD."""
    if not path_str or not isinstance(path_str, str):
        return None
    cleaned = path_str.strip()
    if not cleaned:
        return None
    # 1. Exact path or relative to current working directory
    if os.path.isfile(cleaned):
        return os.path.abspath(cleaned)
    # 2. Check inside UPLOAD_DIR by basename
    in_upload = os.path.join(UPLOAD_DIR, os.path.basename(cleaned))
    if os.path.isfile(in_upload):
        return os.path.abspath(in_upload)
    # 3. Check inside OUTPUT_DIR by basename
    in_output = os.path.join(OUTPUT_DIR, os.path.basename(cleaned))
    if os.path.isfile(in_output):
        return os.path.abspath(in_output)
    # 4. Check relative to BASE_DIR
    in_base = os.path.join(BASE_DIR, cleaned)
    if os.path.isfile(in_base):
        return os.path.abspath(in_base)
    return None


class OfficeApiHandler(http.server.BaseHTTPRequestHandler):
    server_version = "AntigravityOfficeStudio/3.4"

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With")

    def do_OPTIONS(self) -> None:
        self.send_response(200)
        self._send_cors_headers()
        self.end_headers()

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
        raw_body = self.rfile.read(content_length).decode("utf-8", errors="replace")
        return json.loads(raw_body) if raw_body else {}

    def do_GET(self) -> None:
        try:
            parsed_url = urllib.parse.urlparse(self.path)
            path = parsed_url.path

            if path == "/api/health":
                self._send_json({
                    "status": "online",
                    "healthy": True,
                    "version": "3.4.0",
                    "service": "Antigravity Office & Presentation Studio",
                    "timestamp": time.time(),
                })

            elif path == "/api/info":
                self._send_json({
                    "name": "Antigravity Office Studio",
                    "version": "3.4.0",
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
                    template_path=template
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
                    page_range=pages
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
                        from mermaid_renderer import render_mermaid_to_png  # type: ignore

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
                        from plantuml_renderer import render_plantuml_to_png  # type: ignore

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
                        from mxgraph_engine import render_diagram, render_mxgraph_to_png  # type: ignore

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
                text_rep = payload.get("text_replacements", {})
                block_rep = payload.get("block_replacements", {})
                custom_name = payload.get("output_name")

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
                    self._send_json({
                        "success": True,
                        "engine": engine,
                        "snapshot": snapshot,
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

            # Endpoint 15: PDF Everyday Toolkit (merge, split, rotate, clean, watermark, compress)
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

                else:
                    self._send_json({"success": False, "error": f"Unknown pdf-tools action: '{action}'"}, status=400)
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
                    from docx_mail_merge import inspect_template_tokens  # type: ignore
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
                    from docx_mail_merge import execute_mail_merge  # type: ignore

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
                    from docx_diff_engine import compare_docx_documents  # type: ignore

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
                    from ocr_pipeline import VietnameseOcrPipeline  # type: ignore

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
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)
                    ExcelDiagnosticEngine = getattr(mod, "ExcelDiagnosticEngine", None)
                    WordDiagnosticEngine = getattr(mod, "DocxDiagnosticEngine", None) or getattr(mod, "WordDiagnosticEngine", None)
                    PowerPointDiagnosticEngine = getattr(mod, "PptxDiagnosticEngine", None) or getattr(mod, "PowerPointDiagnosticEngine", None)
                    DiagramDiagnosticEngine = getattr(mod, "DiagramAndConversionEngine", None) or getattr(mod, "DiagramDiagnosticEngine", None)

                t0 = time.time()
                issues = []
                metrics = {}
                repaired = False
                repair_actions = []
                out_name = None

                if ext in (".xlsx", ".xlsm"):
                    issues, metrics = ExcelDiagnosticEngine.diagnose(src)
                    if repair_flag:
                        out_name = f"repaired_{os.path.basename(src)}"
                        out_path = os.path.join(OUTPUT_DIR, out_name)
                        repaired, repair_actions = ExcelDiagnosticEngine.repair(src, out_path)
                elif ext in (".docx", ".doc"):
                    issues, metrics = WordDiagnosticEngine.diagnose(src)
                    if repair_flag:
                        out_name = f"repaired_{os.path.basename(src)}"
                        out_path = os.path.join(OUTPUT_DIR, out_name)
                        repaired, repair_actions = WordDiagnosticEngine.repair(src, out_path)
                elif ext in (".pptx", ".potx"):
                    issues, metrics = PowerPointDiagnosticEngine.diagnose(src)
                    if repair_flag:
                        out_name = f"repaired_{os.path.basename(src)}"
                        out_path = os.path.join(OUTPUT_DIR, out_name)
                        repaired, repair_actions = PowerPointDiagnosticEngine.repair(src, out_path)
                elif ext in (".drawio", ".xml", ".svg"):
                    issues, metrics = DiagramDiagnosticEngine.diagnose_diagram_file(src)
                else:
                    self._send_json({"success": False, "error": f"Unsupported format for QA audit: '{ext}'"}, status=400)
                    return

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
                }
                if repaired and out_name:
                    res["download_url"] = f"/api/download/{urllib.parse.quote(out_name)}"
                    res["repaired_file"] = out_name

                self._send_json(res)
                return

            else:
                self._send_json({"error": f"POST endpoint not found: {path}"}, status=404)

        except Exception as exc:
            import traceback
            traceback.print_exc()
            self._send_json({"success": False, "error": str(exc)}, status=500)

    def _handle_upload(self, content_type: str) -> None:
        """Handles binary file upload via multipart/form-data or application/octet-stream."""
        import base64
        import re

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
        """Safely streams a file from either UPLOAD_DIR or OUTPUT_DIR."""
        import re
        candidate_1 = os.path.join(OUTPUT_DIR, filename)
        candidate_2 = os.path.join(UPLOAD_DIR, filename)

        target_file = None
        if os.path.isfile(candidate_1):
            target_file = candidate_1
        elif os.path.isfile(candidate_2):
            target_file = candidate_2

        if not target_file:
            self._send_json({"error": f"File not found: {filename}"}, status=404)
            return

        ext = os.path.splitext(filename)[1].lower()
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
        }
        mime = mime_types.get(ext, "application/octet-stream")

        with open(target_file, "rb") as f:
            content = f.read()

        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        if as_attachment:
            safe_ascii = re.sub(r'[^\x20-\x7E]', '_', filename).replace('"', '')
            encoded_fn = urllib.parse.quote(filename, encoding='utf-8')
            self.send_header("Content-Disposition", f'attachment; filename="{safe_ascii}"; filename*=UTF-8\'\'{encoded_fn}')
        self.end_headers()
        self.wfile.write(content)

    def _serve_static_file(self, rel_path: str) -> None:
        """Safely streams static web frontend files (index.html, src/*, assets/*) from workspace."""
        clean_rel = rel_path.lstrip("/").replace("\\", "/")
        if not clean_rel or clean_rel == "":
            clean_rel = "index.html"

        norm_rel = os.path.normpath(clean_rel)
        if norm_rel.startswith("..") or os.path.isabs(norm_rel):
            self._send_json({"error": "Forbidden path traversal attempt"}, status=403)
            return

        target_file = os.path.join(BASE_DIR, norm_rel)
        if not os.path.isfile(target_file):
            self._send_json({"error": f"Static asset not found: {clean_rel}"}, status=404)
            return

        ext = os.path.splitext(target_file)[1].lower()
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
        }
        mime = mime_types.get(ext, "application/octet-stream")

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
                    from diagram_bridge import render_diagram_code_to_png

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

            else:
                raise ValueError(f"Hành động không hợp lệ: {action}")

        return global_job_manager.submit_job(action, p, _runner)


def is_server_already_running(port: int) -> bool:
    import urllib.request
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/api/health", headers={"User-Agent": "Antigravity-Probe"})
        with urllib.request.urlopen(req, timeout=1.5) as res:
            if res.status == 200:
                return True
    except Exception:
        pass
    return False


def run_server(port: int = PORT) -> None:
    if is_server_already_running(port):
        print(f"✅ [Office API Bridge] Server is already running on http://127.0.0.1:{port} (healthy).")
        try:
            while is_server_already_running(port):
                time.sleep(2)
        except KeyboardInterrupt:
            return

    server_address = ("127.0.0.1", port)
    # Use ThreadingHTTPServer for non-blocking concurrent requests
    handler_class = getattr(http.server, "ThreadingHTTPServer", http.server.HTTPServer)
    httpd = handler_class(server_address, OfficeApiHandler)
    print(f"🚀 [Office API Bridge] Running on http://127.0.0.1:{port} (pid {os.getpid()})")
    print(f"   Uploads directory: {UPLOAD_DIR}")
    print(f"   Outputs directory: {OUTPUT_DIR}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[Office API Bridge] Shutting down gracefully.")
        httpd.server_close()


if __name__ == "__main__":
    port_arg = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else PORT
    run_server(port_arg)
