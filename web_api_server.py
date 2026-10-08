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
from typing import Any, Dict, Optional, Tuple

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
    from .converter_engine import convert_universal
    from .docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx
    from .pptx_writer import THEMES, write_pptx_from_spec
    from .docx_writer import patch_docx_template
    from .xlsx_writer import mutate_template_excel
    from .xlsx_validator import validate_excel_sheet
    from .xlsx_reader import read_xlsx
except (ImportError, ValueError):
    from converter_engine import convert_universal  # type: ignore
    from docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx  # type: ignore
    from pptx_writer import THEMES, write_pptx_from_spec  # type: ignore
    from docx_writer import patch_docx_template  # type: ignore
    from xlsx_writer import mutate_template_excel  # type: ignore
    from xlsx_validator import validate_excel_sheet  # type: ignore
    from xlsx_reader import read_xlsx  # type: ignore

PORT = int(os.environ.get("PORT", 8000))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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
    server_version = "AntigravityOfficeStudio/3.2"

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
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
                    "version": "3.3.0",
                    "service": "Antigravity Office & Presentation Studio",
                    "timestamp": time.time(),
                })

            elif path == "/api/info":
                self._send_json({
                    "name": "Antigravity Office Studio",
                    "version": "3.3.0",
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
                        "universal-convert"
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

    def do_POST(self) -> None:
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        content_type = self.headers.get("Content-Type", "")

        try:
            # Endpoint 1: File Upload (multipart or json base64)
            if path == "/api/upload":
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
                        from .mermaid_renderer import render_mermaid_to_png
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
                        from .plantuml_renderer import render_plantuml_to_png
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
                        from .mxgraph_engine import render_mxgraph_to_png
                    except (ImportError, ValueError):
                        from mxgraph_engine import render_mxgraph_to_png  # type: ignore

                    if isinstance(spec_content, dict):
                        xml_code = spec_content.get("xml_content") or spec_content.get("code") or ""
                    else:
                        xml_code = str(spec_content)
                    render_mxgraph_to_png(xml_code, out_png, scale=int(payload.get("scale", 2)))

                out_name = os.path.basename(out_png)
                quoted_name = urllib.parse.quote(out_name)
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
