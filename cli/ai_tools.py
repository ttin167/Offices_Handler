"""
ai_tools_cli.py — CLI toolkit for AI-driven document editing.

Provides structured read/write/diff operations on DOCX and XLSX files,
outputting JSON snapshots that AI (Antigravity IDE) can read and modify.

Subcommands:
  docx-read       — Read DOCX → JSON snapshot
  docx-write      — Write JSON snapshot → DOCX (with optional style template)
  docx-diff       — Compare two JSON snapshots → before/after diff report
  docx-inject     — Surgically inject paragraphs, runs, and images into DOCX
  mermaid-render  — Render Mermaid code to high-res PNG and inject into DOCX
  plantuml-render — Render PlantUML code to high-res PNG and inject into DOCX
  diagram-render  — Unified multi-engine dispatcher (canvas | mermaid | plantuml)
  spec-render     — Render SVG Canvas Precision Diagram to high-res PNG
  diagram-editor  — Interactive Canvas Editor for Precision Diagram JSON specs
  docx-to-pptx    — Synthesize 16:9 PowerPoint presentation deck from Word .docx
  pptx-build      — Compile 16:9 PowerPoint deck from structured JSON spec (or pptx-create)
  pptx-preview    — Export presentation slides to high-res PNG previews for QA
  pptx-from-md    — Compile Markdown slide deck ('---' delimiter) to .pptx
  xlsx-read       — Read XLSX → JSON snapshot (auto-hybrid, headless openpyxl fallback)
  xlsx-write      — Write JSON snapshot → XLSX
  xlsx-mutate     — Mutate Excel template in-place (row/col expansion, style cloning, chart relink)
  xlsx-validate   — Validate Excel sheet against reference template across 12 Quality Gates (Invariant E6)
  doctor          — Run Zero-Setup Environment Doctor diagnostic suite (v4.0)

  python ai_tools_cli.py doctor
  python ai_tools_cli.py doctor --json
  python ai_tools_cli.py pptx-preview deck.pptx -o ./preview_slides --slides 1,5,11,15
  python ai_tools_cli.py docx-to-pptx document.docx -o presentation.pptx --theme thesis_blue
  python ai_tools_cli.py diagram-render --spec spec.json -o out.png
  python ai_tools_cli.py plantuml-render erd_spec.json -o erd.png
  python ai_tools_cli.py docx-read file.docx -o snap.json
  python ai_tools_cli.py docx-write snap.json --template file.docx -o out.docx
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import sys
import time
from typing import Any, Dict, List, cast

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# docx-read
# ---------------------------------------------------------------------------

def cmd_docx_read(args: argparse.Namespace) -> int:
    try:
        from core.docx.reader import read_docx_to_json
    except (ImportError, ValueError):
        from core.docx.reader import read_docx_to_json

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] File not found: {src}", file=sys.stderr)
        return 1

    t0 = time.time()
    try:
        result = read_docx_to_json(src, output_path=args.output)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    if args.output:
        print(f"[OK] Snapshot written to: {result}  ({elapsed:.2f}s)")
    else:
        print(result)
    return 0


# ---------------------------------------------------------------------------
# docx-write
# ---------------------------------------------------------------------------

def cmd_docx_write(args: argparse.Namespace) -> int:
    try:
        from core.docx.writer import write_docx_from_json_file
    except (ImportError, ValueError):
        from core.docx.writer import write_docx_from_json_file

    json_path = os.path.abspath(args.snapshot)
    if not os.path.isfile(json_path):
        print(f"[ERROR] Snapshot not found: {json_path}", file=sys.stderr)
        return 1

    template_path = os.path.abspath(args.template) if args.template else None
    if template_path and not os.path.isfile(template_path):
        print(f"[ERROR] Template DOCX not found: {template_path}", file=sys.stderr)
        return 1

    output_path = args.output or os.path.splitext(json_path)[0] + "_written.docx"
    output_path = os.path.abspath(output_path)

    t0 = time.time()
    try:
        result = write_docx_from_json_file(json_path, output_path, template_path)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] DOCX written to: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# docx-inject
# ---------------------------------------------------------------------------

def cmd_docx_inject(args: argparse.Namespace) -> int:
    try:
        from core.docx.writer import inject_content_into_docx
    except (ImportError, ValueError):
        from core.docx.writer import inject_content_into_docx

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Injection spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    out_path = os.path.abspath(args.output) if args.output else src

    t0 = time.time()
    try:
        result = inject_content_into_docx(src, spec, out_path)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] Content injected into: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# docx-patch
# ---------------------------------------------------------------------------

def cmd_docx_patch(args: argparse.Namespace) -> int:
    try:
        try:
            from core.docx.writer import patch_docx_template
        except (ImportError, ValueError):
            from core.docx.writer import patch_docx_template
    except (ImportError, ValueError):
        from core.docx.writer import patch_docx_template  # type: ignore

    template_path = os.path.abspath(args.template)
    if not os.path.isfile(template_path):
        print(f"[ERROR] Template DOCX not found: {template_path}", file=sys.stderr)
        return 1

    patch_path = os.path.abspath(args.patch)
    if not os.path.isfile(patch_path):
        print(f"[ERROR] Patch JSON spec not found: {patch_path}", file=sys.stderr)
        return 1

    with open(patch_path, "r", encoding="utf-8") as f:
        patch_spec = json.load(f)

    text_rep = patch_spec.get("text_replacements") or patch_spec.get("texts") or {}
    block_rep = patch_spec.get("block_replacements") or patch_spec.get("blocks") or {}

    out_path = os.path.abspath(args.output) if args.output else os.path.splitext(template_path)[0] + "_patched.docx"

    t0 = time.time()
    try:
        result = patch_docx_template(template_path, out_path, text_rep, block_rep)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] Patched DOCX written to: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# docx-page-numbers
# ---------------------------------------------------------------------------

def cmd_docx_page_numbers(args: argparse.Namespace) -> int:
    import docx
    try:
        try:
            from core.docx.advanced_engine import inject_dynamic_page_numbers, safe_save_docx
        except (ImportError, ValueError):
            from core.docx.advanced_engine import inject_dynamic_page_numbers, safe_save_docx
    except (ImportError, ValueError):
        from core.docx.advanced_engine import inject_dynamic_page_numbers, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        inject_dynamic_page_numbers(doc, font_name=args.font, font_size_pt=args.size)
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Dynamic page numbers injected into: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-caption-fix
# ---------------------------------------------------------------------------

def cmd_docx_caption_fix(args: argparse.Namespace) -> int:
    import docx
    try:
        try:
            from core.docx.advanced_engine import format_figure_captions, safe_save_docx
        except (ImportError, ValueError):
            from core.docx.advanced_engine import format_figure_captions, safe_save_docx
    except (ImportError, ValueError):
        from core.docx.advanced_engine import format_figure_captions, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        count = format_figure_captions(doc, font_name=args.font, font_size_pt=args.size)
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Standardized {count} captions in: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-sanitize-icons
# ---------------------------------------------------------------------------

def cmd_docx_sanitize_icons(args: argparse.Namespace) -> int:
    import docx
    try:
        try:
            from core.docx.advanced_engine import sanitize_emojis_and_symbols, safe_save_docx
        except (ImportError, ValueError):
            from core.docx.advanced_engine import sanitize_emojis_and_symbols, safe_save_docx
    except (ImportError, ValueError):
        from core.docx.advanced_engine import sanitize_emojis_and_symbols, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        count = sanitize_emojis_and_symbols(doc)
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Sanitized {count} elements (emojis/icons) in: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-sanitize-vietnamese
# ---------------------------------------------------------------------------

def cmd_docx_sanitize_vietnamese(args: argparse.Namespace) -> int:
    import docx
    try:
        from core.docx.advanced_engine import sanitize_vietnamese_typography, safe_save_docx
    except (ImportError, ValueError):
        from core.docx.advanced_engine import sanitize_vietnamese_typography, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        count = sanitize_vietnamese_typography(doc)
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Sanitized {count} text runs (Vietnamese NFC, spacing, En-dash) in: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-apply-preset
# ---------------------------------------------------------------------------

def cmd_docx_apply_preset(args: argparse.Namespace) -> int:
    import docx
    try:
        from core.docx.presets import apply_preset_to_document, build_administrative_header
        from core.docx.advanced_engine import safe_save_docx
    except (ImportError, ValueError):
        from core.docx.presets import apply_preset_to_document, build_administrative_header  # type: ignore
        from core.docx.advanced_engine import safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        apply_preset_to_document(doc, preset=args.preset)
        if getattr(args, "add_header", False):
            build_administrative_header(
                doc,
                agency_parent=getattr(args, "agency_parent", "BỘ THÔNG TIN VÀ TRUYỀN THÔNG"),
                agency_name=getattr(args, "agency_name", "CỤC CÔNG NGHỆ THÔNG TIN"),
                document_no=getattr(args, "doc_no", "Số: 01/QĐ-CNTT"),
                location_date=getattr(args, "date", "Hà Nội, ngày 18 tháng 10 năm 2026"),
                insert_at_top=True,
            )
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Applied preset '{args.preset}' to: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pdf-tools
# ---------------------------------------------------------------------------

def cmd_pdf_tools(args: argparse.Namespace) -> int:
    try:
        from core.converter.pdf_toolkit import (
            merge_pdfs,
            split_pdf,
            rotate_and_clean_pdf,
            stamp_watermark_pdf,
            compress_pdf,
            redact_sensitive_pdf,
            protect_pdf,
            normalize_and_number_pdf,
            stamp_giap_lai_pdf,
            extract_pdf_tables_to_xlsx,
            reorganize_pdf_pages,
            convert_images_to_pdf,
            export_pdf_to_images_zip,
            sign_pdf_with_stamp,
            flatten_and_sanitize_pdf,
        )
    except (ImportError, ValueError):
        from core.converter.pdf_toolkit import (  # type: ignore
            merge_pdfs,
            split_pdf,
            rotate_and_clean_pdf,
            stamp_watermark_pdf,
            compress_pdf,
            redact_sensitive_pdf,
            protect_pdf,
            normalize_and_number_pdf,
            stamp_giap_lai_pdf,
            extract_pdf_tables_to_xlsx,
            reorganize_pdf_pages,
            convert_images_to_pdf,
            export_pdf_to_images_zip,
            sign_pdf_with_stamp,
            flatten_and_sanitize_pdf,
        )

    action = args.action
    try:
        if action == "merge":
            inputs = args.inputs or ([args.input] if args.input else [])
            if len(inputs) < 2:
                print("[ERROR] 'merge' action requires at least 2 input files (--inputs file1.pdf file2.pdf)", file=sys.stderr)
                return 1
            out = args.output or "merged.pdf"
            saved = merge_pdfs(inputs, out)
            print(f"[OK] Merged {len(inputs)} PDFs into: {saved}")
            return 0

        elif action == "split":
            if not args.input:
                print("[ERROR] 'split' action requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out_dir = args.output or "./split_output"
            files = split_pdf(args.input, out_dir, page_ranges=args.ranges)
            print(f"[OK] Split into {len(files)} files in: {out_dir}")
            return 0

        elif action in ("rotate", "clean"):
            if not args.input:
                print(f"[ERROR] '{action}' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or args.input
            res = rotate_and_clean_pdf(
                args.input,
                out,
                rotation=getattr(args, "angle", 0) if action == "rotate" else 0,
                remove_blank=getattr(args, "remove_blank", False) or action == "clean",
            )
            print(f"[OK] Processed PDF: {res}")
            return 0

        elif action == "watermark":
            if not args.input:
                print("[ERROR] 'watermark' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or args.input
            text = getattr(args, "text", "BẢN QUYỀN")
            saved = stamp_watermark_pdf(args.input, out, watermark_text=text, opacity=getattr(args, "opacity", 0.25))
            print(f"[OK] Applied watermark to: {saved}")
            return 0

        elif action == "compress":
            if not args.input:
                print("[ERROR] 'compress' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or args.input
            res = compress_pdf(args.input, out)
            print(f"[OK] Compressed PDF ({res['reduction_pct']}% reduction): {res['output_path']}")
            return 0

        elif action == "redact":
            if not args.input:
                print("[ERROR] 'redact' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_redacted.pdf"
            pats = [p.strip() for p in args.patterns.split(",")] if getattr(args, "patterns", None) else None
            trms = [t.strip() for t in args.terms.split(",")] if getattr(args, "terms", None) else None
            res = redact_sensitive_pdf(args.input, out, patterns=pats, custom_terms=trms)
            print(f"[OK] Redacted sensitive data ({res['total_redactions']} redactions on {res['pages_affected_count']} pages): {res['output_path']}")
            return 0

        elif action == "protect":
            if not args.input:
                print("[ERROR] 'protect' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_protected.pdf"
            u_pwd = getattr(args, "user_pwd", None)
            o_pwd = getattr(args, "owner_pwd", None)
            allow_print = not getattr(args, "no_print", False)
            allow_copy = getattr(args, "allow_copy", False)
            allow_modify = getattr(args, "allow_modify", False)
            res = protect_pdf(
                args.input,
                out,
                user_password=u_pwd,
                owner_password=o_pwd,
                allow_print=allow_print,
                allow_copy=allow_copy,
                allow_modify=allow_modify,
            )
            print(f"[OK] Protected PDF with AES-256 encryption: {res['output_path']}")
            return 0

        elif action in ("normalize-a4", "normalize_a4"):
            if not args.input:
                print("[ERROR] 'normalize-a4' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_a4.pdf"
            res = normalize_and_number_pdf(
                args.input,
                out,
                orientation=getattr(args, "orientation", "auto"),
                bates_prefix=getattr(args, "bates_prefix", None),
            )
            print(f"[OK] Normalized PDF pages to standard A4 (total {res['total_pages']} pages): {res['output_path']}")
            return 0

        elif action in ("giap-lai", "giap_lai"):
            if not args.input:
                print("[ERROR] 'giap-lai' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            stamp_file = getattr(args, "stamp", None)
            if not stamp_file or not os.path.isfile(stamp_file):
                print(f"[ERROR] 'giap-lai' requires a valid stamp image (--stamp image.png)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_giap_lai.pdf"
            res = stamp_giap_lai_pdf(
                args.input,
                out,
                stamp_file,
                side=getattr(args, "side", "right"),
            )
            print(f"[OK] Applied Giáp Lai stamp across {res['stamp_slices']} pages: {res['output_path']}")
            return 0

        elif action in ("extract-tables", "extract_tables"):
            if not args.input:
                print("[ERROR] 'extract-tables' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_tables.xlsx"
            res = extract_pdf_tables_to_xlsx(args.input, out)
            print(f"[OK] Extracted {res['tables_extracted']} tables ({res['total_rows']} rows) to Excel: {res['output_path']}")
            return 0

        elif action in ("images-to-pdf", "images_to_pdf", "img2pdf"):
            inputs = args.inputs or ([args.input] if args.input else [])
            if not inputs:
                print("[ERROR] 'images-to-pdf' requires at least one image input (--inputs img1.png img2.jpg)", file=sys.stderr)
                return 1
            out = args.output or "output_images.pdf"
            res = convert_images_to_pdf(
                inputs,
                out,
                page_size=getattr(args, "page_size", "A4"),
                orientation=getattr(args, "orientation", "auto"),
                margin_pt=float(getattr(args, "margin", 36.0)),
            )
            print(f"[OK] Converted {res['total_images']} images into PDF ({res['total_pages']} pages): {res['output_path']}")
            return 0

        elif action in ("export-images-zip", "export_images_zip", "to-images", "pdf2img"):
            if not args.input:
                print("[ERROR] 'export-images-zip' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_images.zip"
            res = export_pdf_to_images_zip(
                args.input,
                out,
                dpi=int(getattr(args, "dpi", 300)),
                img_format=getattr(args, "format", "png"),
            )
            print(f"[OK] Exported {res['total_pages']} pages ({res['dpi']} DPI) to ZIP: {res['output_zip']}")
            return 0

        elif action in ("sign", "sign-pdf", "sign_pdf"):
            if not args.input:
                print("[ERROR] 'sign' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            stamp_file = getattr(args, "stamp", None) or getattr(args, "signature", None)
            if not stamp_file or not os.path.isfile(stamp_file):
                print("[ERROR] 'sign' requires a valid signature image (--stamp/--signature img.png)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_signed.pdf"
            page_num = int(getattr(args, "page", 1))
            x_pct = float(getattr(args, "x_pct", 0.65))
            y_pct = float(getattr(args, "y_pct", 0.75))
            w_pct = float(getattr(args, "w_pct", 0.25))
            h_pct = float(getattr(args, "h_pct", 0.15))
            res = sign_pdf_with_stamp(
                args.input,
                out,
                [{"page": page_num, "image_path": stamp_file, "x_pct": x_pct, "y_pct": y_pct, "w_pct": w_pct, "h_pct": h_pct}],
            )
            print(f"[OK] Applied signature/stamp to page {page_num}: {res['output_path']}")
            return 0

        elif action in ("flatten", "sanitize", "flatten-sanitize"):
            if not args.input:
                print("[ERROR] 'flatten' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_flattened.pdf"
            strict = getattr(args, "strict", False)
            res = flatten_and_sanitize_pdf(
                args.input,
                out,
                flatten_vector=True,
                strict_raster=strict,
                sanitize_metadata=True,
            )
            print(f"[OK] Flattened & sanitized PDF (mode: {res['mode']}): {res['output_path']}")
            return 0

        elif action == "reorganize":
            if not args.input:
                print("[ERROR] 'reorganize' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_reorganized.pdf"
            plan_str = getattr(args, "plan", None)
            if plan_str:
                import json
                try:
                    plan = json.loads(plan_str)
                except Exception:
                    # Parse simple list like 1,3,2 or 1:90,2:0
                    plan: List[Dict[str, Any]] = []
                    for seg in plan_str.split(","):
                        seg = seg.strip()
                        if not seg:
                            continue
                        if seg.lower() == "blank":
                            plan.append({"is_blank": True})
                        elif ":" in seg:
                            p, r = seg.split(":", 1)
                            plan.append({"page": int(p), "rotation": int(r)})
                        else:
                            plan.append({"page": int(seg), "rotation": 0})
            else:
                print("[ERROR] 'reorganize' requires a plan (--plan '1,3,blank,2' or JSON)", file=sys.stderr)
                return 1

            res = reorganize_pdf_pages(args.input, out, plan)
            print(f"[OK] Reorganized PDF ({res['output_pages']} output pages): {res['output_path']}")
            return 0

        else:
            print(f"[ERROR] Unknown action: {action}", file=sys.stderr)
            return 1
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-mail-merge (v4.0)
# ---------------------------------------------------------------------------

def cmd_docx_mail_merge(args: argparse.Namespace) -> int:
    try:
        from core.docx.mail_merge import execute_mail_merge, inspect_template_tokens
    except (ImportError, ValueError):
        from core.docx.mail_merge import execute_mail_merge, inspect_template_tokens  # type: ignore

    template_path = args.template
    if not os.path.isfile(template_path):
        print(f"[ERROR] Template file not found: {template_path}", file=sys.stderr)
        return 1

    if args.inspect:
        tokens = inspect_template_tokens(template_path)
        print(json.dumps({"template": template_path, "tokens": tokens}, indent=2, ensure_ascii=False))
        return 0

    if not args.data:
        print("[ERROR] --data <file.xlsx|data.json> is required unless --inspect is specified", file=sys.stderr)
        return 1

    out_dir = args.output or "./mail_merge_output"
    res = execute_mail_merge(
        template_path=template_path,
        data_source=args.data,
        output_dir=out_dir,
        filename_pattern=args.pattern,
        zip_output=args.zip,
    )
    print(json.dumps(res, indent=2, ensure_ascii=False))
    return 0 if res.get("success") else 1


# ---------------------------------------------------------------------------
# docx-diff (v4.0 Redline Track Changes)
# ---------------------------------------------------------------------------

def cmd_docx_diff(args: argparse.Namespace) -> int:
    param_a = getattr(args, "doc_a", None) or getattr(args, "before", None)
    param_b = getattr(args, "doc_b", None) or getattr(args, "after", None)

    if not param_a or not param_b:
        print("[ERROR] Both original (doc_a/before) and revised (doc_b/after) file paths are required.", file=sys.stderr)
        return 1

    doc_a = os.path.abspath(param_a)
    doc_b = os.path.abspath(param_b)
    if not os.path.isfile(doc_a):
        print(f"[ERROR] Original file not found: {doc_a}", file=sys.stderr)
        return 1
    if not os.path.isfile(doc_b):
        print(f"[ERROR] Revised file not found: {doc_b}", file=sys.stderr)
        return 1

    ext_a = os.path.splitext(doc_a)[1].lower()
    ext_b = os.path.splitext(doc_b)[1].lower()

    # JSON snapshot diff mode
    if ext_a == ".json" and ext_b == ".json":
        import difflib
        with open(doc_a, "r", encoding="utf-8") as f:
            snap_a = json.load(f)
        with open(doc_b, "r", encoding="utf-8") as f:
            snap_b = json.load(f)
        lines_a = json.dumps(snap_a, indent=2, ensure_ascii=False).splitlines()
        lines_b = json.dumps(snap_b, indent=2, ensure_ascii=False).splitlines()
        diff = list(difflib.unified_diff(lines_a, lines_b, fromfile=doc_a, tofile=doc_b, lineterm=""))
        out_content = "\n".join(diff)
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(out_content)
            print(f"[OK] Diff written to: {args.output}")
        else:
            print(out_content)
        return 0

    # DOCX Redline Diff mode
    try:
        from core.docx.diff_engine import compare_docx_documents
    except (ImportError, ValueError):
        from core.docx.diff_engine import compare_docx_documents  # type: ignore

    out_path = args.output or os.path.splitext(doc_b)[0] + "_redline.docx"
    author = getattr(args, "author", "AI Reviewer") or "AI Reviewer"

    t0 = time.time()
    try:
        res = compare_docx_documents(doc_a, doc_b, out_path, author=author)
        elapsed = time.time() - t0
        res["elapsed_seconds"] = round(elapsed, 2)
        print(json.dumps(res, indent=2, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"[ERROR] Redline diff failed: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# ocr-scan (v4.0 Vietnamese OCR Pipeline)
# ---------------------------------------------------------------------------

def cmd_ocr_scan(args: argparse.Namespace) -> int:
    try:
        from core.converter.ocr_pipeline import VietnameseOcrPipeline
    except (ImportError, ValueError):
        from core.converter.ocr_pipeline import VietnameseOcrPipeline  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Input file not found: {src}", file=sys.stderr)
        return 1

    pipeline = VietnameseOcrPipeline(
        lang=args.lang,
        filter_red_stamp=not args.no_red_stamp,
        deskew=not args.no_deskew,
    )
    t0 = time.time()
    try:
        res = pipeline.process_document(src, output_path=args.output)
        elapsed = time.time() - t0
        res["elapsed_seconds"] = round(elapsed, 2)
        summary = {k: v for k, v in res.items() if k != "pages"}
        print(json.dumps(summary, indent=2, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"[ERROR] OCR pipeline failed: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-translate
# ---------------------------------------------------------------------------

def cmd_docx_translate(args: argparse.Namespace) -> int:
    import docx
    try:
        try:
            from core.docx.advanced_engine import translate_vn_to_en_protected, safe_save_docx
        except (ImportError, ValueError):
            from core.docx.advanced_engine import translate_vn_to_en_protected, safe_save_docx
    except (ImportError, ValueError):
        from core.docx.advanced_engine import translate_vn_to_en_protected, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        translated_count = 0
        for p in doc.paragraphs:
            old_t = p.text
            new_t = translate_vn_to_en_protected(old_t)
            if new_t != old_t:
                p.text = new_t
                translated_count += 1
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        old_t = p.text
                        new_t = translate_vn_to_en_protected(old_t)
                        if new_t != old_t:
                            p.text = new_t
                            translated_count += 1
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Translated {translated_count} blocks in: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# mermaid-render
# ---------------------------------------------------------------------------

def cmd_mermaid_render(args: argparse.Namespace) -> int:
    try:
        try:
            from core.diagrams.legacy.mermaid import render_mermaid_to_png
        except (ImportError, ValueError):
            from core.diagrams.legacy.mermaid import render_mermaid_to_png
    except (ImportError, ValueError):
        from core.diagrams.legacy.mermaid import render_mermaid_to_png  # type: ignore

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Mermaid spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    # CLI flag overrides
    if args.output:
        spec["output_path"] = args.output
    if args.inject:
        spec["inject_into"] = args.inject
    if args.theme:
        spec["theme_preset"] = args.theme
    if args.scale:
        spec["scale"] = args.scale

    base_dir = os.path.dirname(spec_path)
    t0 = time.time()
    try:
        result = render_mermaid_to_png(spec, base_dir=base_dir)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    png_path = result["png_path"]
    w_cm = result["width_cm"]
    h_cm = result["height_cm"]
    w_px, h_px = result["dimensions_px"]
    f_size = result["file_size_bytes"]

    print(f"[OK] Diagram rendered: {png_path}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Doc Layout: {w_cm}cm x {h_cm}cm | Size: {f_size} bytes")

    if result.get("injected"):
        print(f"[OK] Injected into DOCX: {result.get('docx_path')}")

    return 0


# ---------------------------------------------------------------------------
# plantuml-render
# ---------------------------------------------------------------------------

def cmd_plantuml_render(args: argparse.Namespace) -> int:
    try:
        try:
            from core.diagrams.legacy.plantuml import render_plantuml_to_png
        except (ImportError, ValueError):
            from core.diagrams.legacy.plantuml import render_plantuml_to_png
    except (ImportError, ValueError):
        from core.diagrams.legacy.plantuml import render_plantuml_to_png  # type: ignore

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] PlantUML spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    if args.output:
        spec["output_path"] = args.output
    if args.inject:
        spec["inject_into"] = args.inject
    if args.dpi:
        spec["dpi"] = args.dpi

    base_dir = os.path.dirname(spec_path)
    t0 = time.time()
    try:
        result = render_plantuml_to_png(spec, base_dir=base_dir)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    png_path = result["png_path"]
    w_cm = result["width_cm"]
    h_cm = result["height_cm"]
    w_px, h_px = result["dimensions_px"]
    f_size = result["file_size_bytes"]

    print(f"[OK] PlantUML diagram rendered: {png_path}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Doc Layout: {w_cm}cm x {h_cm}cm | Size: {f_size} bytes")

    if result.get("injected"):
        print(f"[OK] Injected into DOCX: {result.get('docx_path')}")

    return 0


# ---------------------------------------------------------------------------
# diagram-render (Unified Multi-Engine Technical Diagram Dispatcher)
# ---------------------------------------------------------------------------

def cmd_diagram_render(args: argparse.Namespace) -> int:
    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Spec file not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    if args.output:
        spec["output_path"] = args.output
    if args.inject:
        spec["inject_into"] = args.inject

    engine = str(spec.get("engine", "mxgraph")).strip().lower()
    base_dir = os.path.dirname(spec_path)
    out_png = args.output or spec.get("output_path") or os.path.splitext(spec_path)[0] + ".png"
    out_png = os.path.abspath(out_png)
    spec["output_path"] = out_png

    t0 = time.time()
    try:
        if engine in ("mxgraph", "drawio"):
            try:
                from core.diagrams.mxgraph import render_diagram
            except (ImportError, ValueError):
                from core.diagrams.mxgraph import render_diagram  # type: ignore
            if args.scale is not None:
                spec["scale"] = args.scale
            result = render_diagram(spec, base_dir=base_dir)
        elif engine == "canvas":
            try:
                from core.diagrams.legacy.canvas import PrecisionDiagram
            except (ImportError, ValueError):
                from core.diagrams.legacy.canvas import PrecisionDiagram  # type: ignore
            scale = args.scale if args.scale is not None else spec.get("scale", 3)
            diag = PrecisionDiagram.from_spec(spec, base_dir=base_dir)
            result = diag.render_to_png(out_png, scale=scale)
        elif engine == "mermaid":
            try:
                from core.diagrams.legacy.mermaid import render_mermaid_to_png
            except (ImportError, ValueError):
                from core.diagrams.legacy.mermaid import render_mermaid_to_png  # type: ignore
            result = render_mermaid_to_png(spec, base_dir=base_dir)
        elif engine == "plantuml":
            try:
                from core.diagrams.legacy.plantuml import render_plantuml_to_png
            except (ImportError, ValueError):
                from core.diagrams.legacy.plantuml import render_plantuml_to_png  # type: ignore
            result = render_plantuml_to_png(spec, base_dir=base_dir)
        else:
            raise ValueError(f"Unknown diagram engine: '{engine}'. Supported: 'mxgraph', 'drawio', 'canvas', 'mermaid', 'plantuml'")
    except Exception as exc:
        print(f"[ERROR] Diagram rendering failed ({engine}): {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    w_px, h_px = result["dimensions_px"]
    f_size = result["file_size_bytes"]
    print(f"[OK] {engine.upper()} diagram rendered: {result['png_path']}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Size: {f_size} bytes")

    # Windows Trap 4: Sau khi render ra result["png_path"] thành công:
    # Auto-inject vào Word nếu spec có inject_into và chưa được inject
    docx_target = spec.get("inject_into")
    if docx_target and not result.get("injected"):
        if not os.path.isabs(docx_target) and base_dir:
            docx_target = os.path.abspath(os.path.join(base_dir, docx_target))
        if os.path.exists(docx_target):
            try:
                from core.docx.writer import inject_diagram_into_docx
            except (ImportError, ValueError):
                from core.docx.writer import inject_diagram_into_docx  # type: ignore
            inject_diagram_into_docx(
                docx_path=docx_target,
                png_path=result["png_path"],
                heading=spec.get("target_heading"),
                placeholder=spec.get("placeholder"),
                caption=spec.get("caption_template") or spec.get("caption"),
                width_cm=spec.get("width_cm", 14.0),
                max_height_cm=spec.get("max_height_cm", 20.0),
            )
            print(f"[OK] Injected diagram into {docx_target}")
        else:
            print(f"[WARN] Target DOCX for injection not found: {docx_target}")
    elif result.get("injected"):
        print(f"[OK] Injected diagram into {docx_target}")

    return 0


# ---------------------------------------------------------------------------
# spec-render (Precision Technical & Architecture Diagram Engine)
# ---------------------------------------------------------------------------

def cmd_spec_render(args: argparse.Namespace) -> int:
    try:
        from core.diagrams.legacy.canvas import PrecisionDiagram
    except (ImportError, ValueError):
        from core.diagrams.legacy.canvas import PrecisionDiagram

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Spec file not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec_dict = json.load(f)

    out_png = args.output or spec_dict.get("output_path") or os.path.splitext(spec_path)[0] + ".png"
    out_png = os.path.abspath(out_png)
    scale = args.scale if args.scale is not None else spec_dict.get("scale", 3)

    t0 = time.time()
    try:
        diag = PrecisionDiagram.from_spec(spec_dict)
        res = diag.render_to_png(out_png, scale=scale)
    except Exception as exc:
        print(f"[ERROR] Precision diagram rendering failed: {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    w_px, h_px = res["dimensions_px"]
    f_size = res["file_size_bytes"]
    ar = res["aspect_ratio"]

    print(f"[OK] Precision Diagram rendered: {res['png_path']}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Aspect Ratio: {ar}:1 | Size: {f_size} bytes | Scale: {scale}x")
    return 0


def cmd_diagram_editor(args: argparse.Namespace) -> int:
    try:
        try:
            from core.diagrams.editor import DiagramEditorApp
        except (ImportError, ValueError):
            from diagram_editor import DiagramEditorApp
    except (ImportError, ValueError):
        from diagram_editor import DiagramEditorApp  # type: ignore
    app = DiagramEditorApp(spec_path=args.spec)
    app.run()
    return 0

def _flatten_body(body: list[dict]) -> dict[int, dict]:
    """Return {element_index: element} mapping."""
    return {el.get("element_index", i): el for i, el in enumerate(body)}


def cmd_docx_diff(args: argparse.Namespace) -> int:
    before_path = os.path.abspath(args.before)
    after_path  = os.path.abspath(args.after)

    for p in (before_path, after_path):
        if not os.path.isfile(p):
            print(f"[ERROR] File not found: {p}", file=sys.stderr)
            return 1

    with open(before_path, "r", encoding="utf-8") as f:
        before_snap = json.load(f)
    with open(after_path, "r", encoding="utf-8") as f:
        after_snap = json.load(f)

    before_body = _flatten_body(before_snap.get("body", []))
    after_body  = _flatten_body(after_snap.get("body", []))

    all_indices = sorted(set(before_body.keys()) | set(after_body.keys()))

    changes = []
    for idx in all_indices:
        b = before_body.get(idx)
        a = after_body.get(idx)

        if b == a:
            continue  # No change

        entry: dict = {"element_index": idx}
        if b is None:
            entry["change_type"] = "added"
            entry["after"] = a
        elif a is None:
            entry["change_type"] = "removed"
            entry["before"] = b
        else:
            entry["change_type"] = "modified"
            entry["before"] = b
            entry["after"] = a

        changes.append(entry)

    diff_report = {
        "before_file": before_snap.get("source_file", before_path),
        "after_file":  after_snap.get("source_file", after_path),
        "total_before_elements": len(before_body),
        "total_after_elements":  len(after_body),
        "total_changes": len(changes),
        "changes": changes,
    }

    json_str = json.dumps(diff_report, ensure_ascii=False, indent=2)

    if args.output:
        out_path = os.path.abspath(args.output)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(json_str)
        print(f"[OK] Diff report written to: {out_path}  ({len(changes)} changes)")
    else:
        print(json_str)

    return 0


# ---------------------------------------------------------------------------
# xlsx-read
# ---------------------------------------------------------------------------

def cmd_xlsx_read(args: argparse.Namespace) -> int:
    try:
        from core.xlsx.reader import read_xlsx_to_json
    except (ImportError, ValueError):
        from core.xlsx.reader import read_xlsx_to_json

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] File not found: {src}", file=sys.stderr)
        return 1

    t0 = time.time()
    try:
        result = read_xlsx_to_json(
            src,
            output_path=args.output,
            sheet_name=args.sheet,
            cell_range=args.range,
            engine=getattr(args, "engine", "auto"),
        )
    except ImportError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 3
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    if args.output:
        print(f"[OK] Snapshot written to: {result}  ({elapsed:.2f}s)")
    else:
        print(result)
    return 0


# ---------------------------------------------------------------------------
# xlsx-write
# ---------------------------------------------------------------------------

def cmd_xlsx_write(args: argparse.Namespace) -> int:
    try:
        from core.xlsx.writer import write_xlsx_from_json_file
    except (ImportError, ValueError):
        from core.xlsx.writer import write_xlsx_from_json_file

    json_path = os.path.abspath(args.snapshot)
    if not os.path.isfile(json_path):
        print(f"[ERROR] Snapshot not found: {json_path}", file=sys.stderr)
        return 1

    template_path = os.path.abspath(args.template) if args.template else None
    if template_path and not os.path.isfile(template_path):
        print(f"[ERROR] Template XLSX not found: {template_path}", file=sys.stderr)
        return 1

    output_path = args.output or os.path.splitext(json_path)[0] + "_written.xlsx"
    output_path = os.path.abspath(output_path)

    t0 = time.time()
    try:
        result = write_xlsx_from_json_file(json_path, output_path, template_path)
    except ImportError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 3
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] XLSX written to: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# xlsx-mutate
# ---------------------------------------------------------------------------

def cmd_xlsx_mutate(args: argparse.Namespace) -> int:
    try:
        try:
            from core.xlsx.writer import mutate_template_excel
        except (ImportError, ValueError):
            from core.xlsx.writer import mutate_template_excel
    except (ImportError, ValueError):
        from core.xlsx.writer import mutate_template_excel  # type: ignore

    template_path = os.path.abspath(args.template)
    if not os.path.isfile(template_path):
        print(f"[ERROR] Template XLSX not found: {template_path}", file=sys.stderr)
        return 1

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Mutation JSON spec not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    out_path = os.path.abspath(args.output) if args.output else os.path.splitext(template_path)[0] + "_mutated.xlsx"

    t0 = time.time()
    try:
        result = mutate_template_excel(template_path, out_path, spec)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] Mutated XLSX written to: {result}  ({elapsed:.2f}s)")

    if getattr(args, "validate", False):
        try:
            from core.xlsx.validator import validate_excel_sheet
        except (ImportError, ValueError):
            from core.xlsx.validator import validate_excel_sheet  # type: ignore

        sheets_spec = spec.get("sheets", {})
        for sheet_name in sheets_spec:
            try:
                res = validate_excel_sheet(
                    result,
                    target_sheet=sheet_name,
                    ref_sheet=getattr(args, "ref_sheet", None),
                    verbose=True,
                )
                if not res.passed:
                    print(f"[WARN] Invariant E6: {len(res.diffs)} format discrepancy(ies) on sheet '{sheet_name}'.", file=sys.stderr)
            except Exception as v_exc:
                print(f"[WARN] Validation skipped for sheet '{sheet_name}': {v_exc}")

    return 0


# ---------------------------------------------------------------------------
# xlsx-validate
# ---------------------------------------------------------------------------

def cmd_xlsx_validate(args: argparse.Namespace) -> int:
    try:
        try:
            from core.xlsx.validator import validate_excel_sheet
        except (ImportError, ValueError):
            from core.xlsx.validator import validate_excel_sheet
    except (ImportError, ValueError):
        from core.xlsx.validator import validate_excel_sheet  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Workbook not found: {src}", file=sys.stderr)
        return 2

    t0 = time.time()
    try:
        res = validate_excel_sheet(
            src,
            target_sheet=args.target,
            ref_sheet=args.ref,
            verbose=True,
        )
        elapsed = time.time() - t0
        if res.passed:
            print(f"[OK] Sheet '{args.target}' passed all 12 Quality Gates! ({elapsed:.2f}s)")
            return 0
        else:
            print(f"[ERROR] Sheet '{args.target}' failed {res.gates_failed} Quality Gate(s) ({elapsed:.2f}s)", file=sys.stderr)
            return 1
    except Exception as exc:
        print(f"[ERROR] Validation failed: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-create
# ---------------------------------------------------------------------------

def cmd_pptx_create(args: argparse.Namespace) -> int:
    try:
        from core.pptx.writer import write_pptx_from_spec
    except (ImportError, ValueError):
        from core.pptx.writer import write_pptx_from_spec  # type: ignore

    spec_raw = getattr(args, "spec_opt", None) or getattr(args, "spec", None)
    if not spec_raw:
        print("[ERROR] No slide spec JSON provided. Use --spec <file.json> or positional argument.", file=sys.stderr)
        return 1

    spec_path = os.path.abspath(spec_raw)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Slide spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    out_path = args.output or spec.get("output_path") or os.path.splitext(spec_path)[0] + ".pptx"
    out_path = os.path.abspath(out_path)

    tpl = args.template or spec.get("template")
    tpl_path = os.path.abspath(tpl) if tpl else None

    try:
        result = write_pptx_from_spec(
            spec=spec,
            output_path=out_path,
            theme_name=args.theme or spec.get("theme", "thesis_blue"),
            template_path=tpl_path,
        )
        print(f"[OK] PowerPoint presentation created: {result}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-from-md
# ---------------------------------------------------------------------------

def cmd_pptx_from_md(args: argparse.Namespace) -> int:
    try:
        try:
            from core.pptx.engine import compile_markdown_file_to_pptx
        except (ImportError, ValueError):
            from core.pptx.engine import compile_markdown_file_to_pptx
    except (ImportError, ValueError):
        from core.pptx.engine import compile_markdown_file_to_pptx  # type: ignore

    md_path = os.path.abspath(args.file)
    if not os.path.isfile(md_path):
        print(f"[ERROR] Markdown file not found: {md_path}", file=sys.stderr)
        return 1

    out_path = args.output or os.path.splitext(md_path)[0] + ".pptx"
    out_path = os.path.abspath(out_path)

    try:
        result = compile_markdown_file_to_pptx(
            md_file_path=md_path,
            output_pptx_path=out_path,
            theme_name=args.theme,
            template_path=os.path.abspath(args.template) if args.template else None,
        )
        print(f"[OK] Compiled Markdown slide deck to: {result}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-to-pptx
# ---------------------------------------------------------------------------

def cmd_docx_to_pptx(args: argparse.Namespace) -> int:
    try:
        try:
            from core.pptx.docx_to_pptx import convert_docx_to_pptx
        except (ImportError, ValueError):
            from core.pptx.docx_to_pptx import convert_docx_to_pptx
    except (ImportError, ValueError):
        from core.pptx.docx_to_pptx import convert_docx_to_pptx  # type: ignore

    docx_path = os.path.abspath(args.file)
    if not os.path.isfile(docx_path):
        print(f"[ERROR] Source DOCX not found: {docx_path}", file=sys.stderr)
        return 1

    try:
        pptx_res, spec_res = convert_docx_to_pptx(
            docx_path=docx_path,
            output_pptx=args.output,
            spec_out=args.spec_out,
            theme=args.theme,
            template_path=os.path.abspath(args.template) if args.template else None,
            deck_scope=getattr(args, "scope", "auto"),
        )
        print(f"[OK] Successfully synthesized presentation: {pptx_res}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to convert DOCX to PPTX: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pdf-to-pptx
# ---------------------------------------------------------------------------

def cmd_pdf_to_pptx(args: argparse.Namespace) -> int:
    try:
        try:
            from core.pptx.docx_to_pptx import convert_pdf_to_pptx
        except (ImportError, ValueError):
            from core.pptx.docx_to_pptx import convert_pdf_to_pptx
    except (ImportError, ValueError):
        from core.pptx.docx_to_pptx import convert_pdf_to_pptx  # type: ignore

    pdf_path = os.path.abspath(args.file)
    if not os.path.isfile(pdf_path):
        print(f"[ERROR] Source PDF not found: {pdf_path}", file=sys.stderr)
        return 1

    try:
        pptx_res, spec_res = convert_pdf_to_pptx(
            pdf_path=pdf_path,
            output_pptx=args.output,
            spec_out=args.spec_out,
            theme=args.theme,
            template_path=os.path.abspath(args.template) if args.template else None,
            page_range=getattr(args, "pages", None),
            deck_scope=getattr(args, "scope", "auto"),
        )

        print(f"[OK] Successfully synthesized presentation from PDF: {pptx_res}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to convert PDF to PPTX: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-insert-diagram
# ---------------------------------------------------------------------------

def cmd_pptx_insert_diagram(args: argparse.Namespace) -> int:
    try:
        try:
            from core.pptx.engine import insert_diagram_into_pptx
        except (ImportError, ValueError):
            from core.pptx.engine import insert_diagram_into_pptx
    except (ImportError, ValueError):
        from core.pptx.engine import insert_diagram_into_pptx  # type: ignore

    pptx_path = os.path.abspath(args.file)
    img_path = os.path.abspath(args.image)
    if not os.path.isfile(pptx_path):
        print(f"[ERROR] Target PPTX not found: {pptx_path}", file=sys.stderr)
        return 1
    if not os.path.isfile(img_path):
        print(f"[ERROR] Diagram image not found: {img_path}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else pptx_path

    try:
        result = insert_diagram_into_pptx(
            pptx_path=pptx_path,
            image_path=img_path,
            slide_index=args.slide,
            title=args.title,
            caption=args.caption,
            output_path=out_path,
            theme_name=args.theme,
        )
        print(f"[OK] Diagram successfully inserted into: {result}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-inspect
# ---------------------------------------------------------------------------

def cmd_pptx_inspect(args: argparse.Namespace) -> int:
    try:
        try:
            from core.pptx.reader import read_pptx_to_json, read_pptx
        except (ImportError, ValueError):
            from core.pptx.reader import read_pptx_to_json, read_pptx
    except (ImportError, ValueError):
        from core.pptx.reader import read_pptx_to_json, read_pptx  # type: ignore

    pptx_path = os.path.abspath(args.file)
    if not os.path.isfile(pptx_path):
        print(f"[ERROR] PPTX not found: {pptx_path}", file=sys.stderr)
        return 1

    try:
        if args.output:
            out_json = read_pptx_to_json(pptx_path, args.output)
            print(f"[OK] Presentation AST written to: {out_json}")
        else:
            ast = read_pptx(pptx_path)
            print(json.dumps(ast, indent=2, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-preview
# ---------------------------------------------------------------------------

def cmd_pptx_preview(args: argparse.Namespace) -> int:
    """Exports PowerPoint presentation slides to high-res PNG previews."""
    import tempfile
    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Presentation file not found: {src}", file=sys.stderr)
        return 1

    out_dir = os.path.abspath(args.output_dir or os.path.join(os.path.dirname(src), "preview_slides"))
    os.makedirs(out_dir, exist_ok=True)

    slides_arg = getattr(args, "slides", "all") or "all"
    slides_arg = slides_arg.strip()
    dpi = getattr(args, "dpi", 150) or 150
    t0 = time.time()

    exported_files = []

    # 1. Native Windows PowerPoint COM automation (Pixel-perfect 1920x1080)
    if sys.platform == "win32":
        try:
            import win32com.client
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            deck = powerpoint.Presentations.Open(src, WithWindow=False)
            total = deck.Slides.Count

            if slides_arg.lower() == "all":
                target_indices = list(range(1, total + 1))
            else:
                target_indices = [int(x.strip()) for x in slides_arg.split(",") if x.strip().isdigit()]

            for s_idx in target_indices:
                if 1 <= s_idx <= total:
                    out_png = os.path.join(out_dir, f"slide_{s_idx:02d}_preview.png")
                    deck.Slides(s_idx).Export(out_png, "PNG", 1920, 1080)
                    exported_files.append(out_png)

            deck.Close()
            powerpoint.Quit()
        except Exception as com_err:
            print(f"[WARN] PowerPoint COM export unavailable or encountered an error: {com_err}", file=sys.stderr)

    # 2. Headless LibreOffice / PyMuPDF fallback
    if not exported_files:
        try:
            import subprocess, pymupdf
            temp_pdf_dir = tempfile.mkdtemp()
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", src, "--outdir", temp_pdf_dir],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
            )
            pdf_name = os.path.splitext(os.path.basename(src))[0] + ".pdf"
            temp_pdf = os.path.join(temp_pdf_dir, pdf_name)
            if os.path.isfile(temp_pdf):
                doc = pymupdf.open(temp_pdf)
                total = len(doc)
                if slides_arg.lower() == "all":
                    target_indices = list(range(1, total + 1))
                else:
                    target_indices = [int(x.strip()) for x in slides_arg.split(",") if x.strip().isdigit()]

                for s_idx in target_indices:
                    if 1 <= s_idx <= total:
                        page = doc[s_idx - 1]
                        pix = page.get_pixmap(dpi=dpi)
                        out_png = os.path.join(out_dir, f"slide_{s_idx:02d}_preview.png")
                        pix.save(out_png)
                        exported_files.append(out_png)
                doc.close()
        except Exception as fb_err:
            print(f"[WARN] Headless fallback failed: {fb_err}", file=sys.stderr)

    elapsed = time.time() - t0
    if exported_files:
        print(f"[OK] Exported {len(exported_files)} slide previews to: {out_dir}  ({elapsed:.2f}s)")
        for f in exported_files[:5]:
            print(f"     - {os.path.basename(f)}")
        if len(exported_files) > 5:
            print(f"     ... and {len(exported_files) - 5} more.")
        return 0
    else:
        print("[ERROR] Could not export previews. Ensure Microsoft PowerPoint (Windows) or LibreOffice is installed.", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# felo-generate
# ---------------------------------------------------------------------------

def cmd_felo_generate(args: argparse.Namespace) -> int:
    skill_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".agents", "skills", "felo-slides", "scripts")
    if not os.path.isdir(skill_dir):
        skill_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "skills", "felo-slides", "scripts")

    adapter_path = os.path.join(skill_dir, "felo_adapter.py")
    if not os.path.isfile(adapter_path):
        print(f"[ERROR] Felo adapter not found at: {adapter_path}", file=sys.stderr)
        return 1

    sys.path.insert(0, skill_dir)
    try:
        import felo_adapter  # type: ignore
        res = felo_adapter.run_felo_ppt_task(
            query=args.query,
            file_path=args.file,
            theme_id=args.theme,
            task_id=args.task_id,
        )
        print("\n=== FELO PRESENTATION GENERATION RESULT ===")
        print(json.dumps(res, indent=2, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# doctor
# ---------------------------------------------------------------------------

def cmd_doctor(args: argparse.Namespace) -> int:
    try:
        from core.env_doctor import run_doctor, format_doctor_report
    except (ImportError, ValueError):
        from core.env_doctor import run_doctor, format_doctor_report  # type: ignore

    report = run_doctor()
    if getattr(args, "json", False):
        print(json.dumps(report, indent=2, ensure_ascii=False))
    else:
        print(format_doctor_report(report))
    return 0 if report.get("status") == "healthy" else 1


# ---------------------------------------------------------------------------
# qa-audit (Unified QA Diagnostic)
# ---------------------------------------------------------------------------

def cmd_qa_audit(args: argparse.Namespace) -> int:
    try:
        from tools.unified_qa_diagnostic import run_pipeline
    except ImportError:
        import importlib.util
        spec_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools", "unified_qa_diagnostic.py")
        if not os.path.exists(spec_path):
            spec_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tools", "unified_qa_diagnostic.py")
        spec = importlib.util.spec_from_file_location("unified_qa_diagnostic", spec_path)
        if spec and spec.loader:
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            run_pipeline = getattr(mod, "run_pipeline")
        else:
            raise ImportError(f"Could not load unified_qa_diagnostic from {spec_path}")

    return run_pipeline(args)


# ---------------------------------------------------------------------------
# doc-security (Cloudflare Harness Document Security & Sanitizer)
# ---------------------------------------------------------------------------

def cmd_doc_security(args: argparse.Namespace) -> int:
    try:
        from core.security.document_auditor import DocumentSecurityAuditor
        from core.security.sanitizer import DocumentSanitizer
    except ImportError:
        import sys
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if base_dir not in sys.path:
            sys.path.insert(0, base_dir)
        from core.security.document_auditor import DocumentSecurityAuditor
        from core.security.sanitizer import DocumentSanitizer

    auditor = DocumentSecurityAuditor()
    report = auditor.audit_file(args.file)

    if getattr(args, "sanitize", False):
        sanitizer = DocumentSanitizer()
        out_p = getattr(args, "output", None)
        san_res = sanitizer.sanitize(args.file, output_path=out_p)
        print(f"[SECURITY SANITIZE] Cleaned file saved to: {san_res.get('output_file')}")
        for act in san_res.get("actions_taken", []):
            print(f"  + {act}")

    if getattr(args, "format", "terminal") == "json":
        import json
        out_str = json.dumps(report.to_dict(), indent=2, ensure_ascii=False)
        if getattr(args, "output", None) and not getattr(args, "sanitize", False):
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(out_str)
            print(f"Security report saved to: {args.output}")
        else:
            print(out_str)
    else:
        print("=" * 80)
        print("       ENTERPRISE DOCUMENT SECURITY AUDIT (CLOUDFLARE HARNESS)")
        print("=" * 80)
        print(f"Target File: {report.target}")
        print(f"Timestamp:   {report.timestamp}")
        print(f"Status:      {'100% CLEAN' if report.is_clean else 'THREATS DETECTED'}")
        print(f"Counts:      Critical: {report.critical_count} | Warning: {report.warning_count} | Info: {report.info_count}")
        print("-" * 80)
        for cat in ("confirmed", "needs_validation", "rejected"):
            items = report.findings.get(cat, [])
            if items:
                print(f"[{cat.upper()} ({len(items)})]")
                for item in items:
                    sev = item.get("severity", "-")
                    code = item.get("invariant_code", "-")
                    title = item.get("title", "")
                    evidence = item.get("evidence", "")
                    print(f"  * [{sev}] {code} - {title}")
                    if evidence:
                        print(f"    Evidence: {evidence}")
                    if item.get("fix_proposal"):
                        print(f"    Fix: {item.get('fix_proposal')}")
                print("-" * 80)

    return 2 if report.critical_count > 0 else (1 if report.warning_count > 0 else 0)


# ---------------------------------------------------------------------------
# md-toc & convert
# ---------------------------------------------------------------------------

def cmd_md_toc(args: argparse.Namespace) -> int:
    md_path = os.path.abspath(args.file)
    if not os.path.isfile(md_path):
        print(f"[ERROR] Markdown file not found: {md_path}", file=sys.stderr)
        return 1

    try:
        from core.converter.markdown import generate_markdown_toc, update_markdown_toc_in_file
    except (ImportError, ValueError):
        from core.converter.markdown import generate_markdown_toc, update_markdown_toc_in_file  # type: ignore

    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    max_lvl = getattr(args, "max_level", 3) or 3
    min_lvl = getattr(args, "min_level", 2) or 2
    toc_text = generate_markdown_toc(content, min_level=min_lvl, max_level=max_lvl)

    if getattr(args, "check", False):
        print("[INFO] Generated Table of Contents preview:")
        print(toc_text)
        return 0

    if getattr(args, "in_place", False):
        ok = update_markdown_toc_in_file(md_path, max_level=max_lvl)
        if ok:
            print(f"[OK] Successfully updated interactive Table of Contents in: {md_path}")
            return 0
        else:
            print(f"[ERROR] Could not update TOC in: {md_path}", file=sys.stderr)
            return 1

    if getattr(args, "output", None):
        out_path = os.path.abspath(args.output)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(toc_text)
        print(f"[OK] Table of Contents written to: {out_path}")
        return 0

    print(toc_text)
    return 0


def cmd_convert(args: argparse.Namespace) -> int:
    try:
        from core.converter.engine import convert_universal
    except (ImportError, ValueError):
        from core.converter.engine import convert_universal  # type: ignore

    source_path = os.path.abspath(args.input)
    if not os.path.isfile(source_path):
        print(f"[ERROR] Source file not found: {source_path}", file=sys.stderr)
        return 1

    target_path = getattr(args, "output", None)
    if not target_path:
        t_ext = f".{args.target}" if getattr(args, "target", None) else ".docx"
        target_path = os.path.splitext(source_path)[0] + t_ext
    else:
        target_path = os.path.abspath(target_path)

    ok, out_p, elapsed = convert_universal(
        source_path,
        target_path,
        style_path=getattr(args, "style", None),
        page_range=getattr(args, "pages", None),
        deck_scope=getattr(args, "scope", "auto"),
    )
    if ok:
        print(f"[OK] Conversion completed in {elapsed:.2f}s: {out_p}")
        return 0
    else:
        print(f"[ERROR] Conversion failed: {out_p}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# Atomic Office Tools (CLI Commands)
# ---------------------------------------------------------------------------

def cmd_xlsx_set(args: argparse.Namespace) -> int:
    try:
        from core.xlsx.writer import UniversalExcelEngine
    except (ImportError, ValueError):
        from core.xlsx.writer import UniversalExcelEngine  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Excel file not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        engine = UniversalExcelEngine(src)
        sheet_name = args.sheet
        if sheet_name and sheet_name in engine.wb.sheetnames:
            ws = engine.wb[sheet_name]
        else:
            ws = engine.wb.active or engine.wb.worksheets[0]
        cell_coord = args.cell.strip().upper()
        engine.safe_set_cell(ws, cell_coord, value=args.value, formula=args.formula)
        engine.save(out_path)
        sheet_title = getattr(ws, "title", "Sheet")
        print(f"[OK] Cell {cell_coord} updated in '{sheet_title}' -> {out_path}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to set cell: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_insert_rows(args: argparse.Namespace) -> int:
    try:
        from core.xlsx.writer import UniversalExcelEngine
    except (ImportError, ValueError):
        from core.xlsx.writer import UniversalExcelEngine  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Excel file not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        engine = UniversalExcelEngine(src)
        sheet_name = args.sheet
        if not sheet_name or sheet_name not in engine.wb.sheetnames:
            sheet_name = engine.wb.sheetnames[0]
        proto_row = int(args.proto) if args.proto is not None else None
        created = engine.expand_table_rows(
            sheet_name=sheet_name,
            start_row=int(args.start_row),
            count=int(args.count),
            prototype_row=proto_row,
            shift_formulas=not getattr(args, "no_shift_formulas", False),
        )
        engine.save(out_path)
        print(f"[OK] Inserted {len(created)} row(s) starting at row {args.start_row} in sheet '{sheet_name}' -> {out_path}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to insert rows: {exc}", file=sys.stderr)
        return 2


def cmd_docx_tokens(args: argparse.Namespace) -> int:
    try:
        from core.docx.mail_merge import inspect_template_tokens
    except (ImportError, ValueError):
        try:
            from mail_merge import inspect_template_tokens  # type: ignore
        except (ImportError, ValueError):
            from core.docx.writer import inspect_template_tokens  # type: ignore

    src = os.path.abspath(args.template)
    if not os.path.isfile(src):
        print(f"[ERROR] Word template file not found: {src}", file=sys.stderr)
        return 1

    try:
        res = inspect_template_tokens(src)
        if getattr(args, "json", False):
            print(json.dumps(res, indent=2, ensure_ascii=False))
        else:
            print(f"Template: {os.path.basename(src)}")
            print(f"Total Placeholders: {res.get('total_placeholders', 0)}")
            vars_ = res.get("variables", [])
            if vars_:
                print(f"Variables ({len(vars_)}):")
                for v in vars_:
                    print(f"  - {{{{ {v} }}}}")
            conds = res.get("conditionals", [])
            if conds:
                print(f"Conditionals ({len(conds)}):")
                for c in conds:
                    print(f"  - {c}")
            blocks = res.get("blocks", [])
            if blocks:
                print(f"Blocks ({len(blocks)}):")
                for b in blocks:
                    print(f"  - {b}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to inspect docx tokens: {exc}", file=sys.stderr)
        return 2


def cmd_docx_replace(args: argparse.Namespace) -> int:
    try:
        from core.docx.writer import patch_docx_template
    except (ImportError, ValueError):
        from core.docx.writer import patch_docx_template  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Word document not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src

    raw_data = args.data
    replacements: dict = {}
    if os.path.isfile(raw_data):
        with open(raw_data, "r", encoding="utf-8") as f:
            replacements = json.load(f)
    else:
        try:
            replacements = json.loads(raw_data)
        except Exception:
            print(f"[ERROR] --data must be a valid JSON string or path to a JSON file", file=sys.stderr)
            return 1

    if not isinstance(replacements, dict):
        print(f"[ERROR] Replacements data must be a JSON object (dictionary)", file=sys.stderr)
        return 1

    # Normalize token keys: support both "KEY" and "{{KEY}}"
    normalized_rep: dict[str, str] = {}
    for k, v in replacements.items():
        k_str = str(k).strip()
        v_str = str(v)
        normalized_rep[k_str] = v_str
        if not k_str.startswith("{{") and not k_str.endswith("}}"):
            normalized_rep[f"{{{{{k_str}}}}}"] = v_str

    try:
        res_path = patch_docx_template(src, out_path, text_replacements=normalized_rep)
        print(f"[OK] Replaced {len(replacements)} token(s) in -> {res_path}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to replace docx text: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_formulas(args: argparse.Namespace) -> int:
    try:
        from core.xlsx.formula_playbook import get_formula_catalog, generate_formula_playbook_xlsx
    except (ImportError, ValueError):
        try:
            from formula_playbook import get_formula_catalog, generate_formula_playbook_xlsx  # type: ignore
        except (ImportError, ValueError):
            print("[ERROR] Failed to import core.xlsx.formula_playbook", file=sys.stderr)
            return 1

    if getattr(args, "export", None):
        dest = os.path.abspath(args.export)
        ok, res = generate_formula_playbook_xlsx(dest)
        if ok:
            print(f"[OK] Exported Excel Formula Playbook -> {res}")
            return 0
        else:
            print(f"[ERROR] Failed to export formula playbook: {res}", file=sys.stderr)
            return 1

    cat = getattr(args, "category", "all")
    search = getattr(args, "search", None)
    items = get_formula_catalog(category=cat, search=search)

    if getattr(args, "json", False):
        print(json.dumps(items, ensure_ascii=False, indent=2))
        return 0

    print("=" * 80)
    print(f" EXCEL FORMULA PLAYBOOK & CHEAT SHEET ({len(items)} functions)")
    if cat != "all":
        print(f" Category: {cat}")
    if search:
        print(f" Search: {search}")
    print("=" * 80)

    for idx, item in enumerate(items, start=1):
        print(f"\n[{idx:02d}] {item['name']}  |  {item['category_vi']}")
        print(f"  Cú Pháp:   {item['template']}")
        print(f"  Ví Dụ:     {item['example']}")
        print(f"  Khi Nào:   {item['when_to_use']}")
        print(f"  Tại Sao:   {item['why_choose']}")
        print("-" * 80)

    return 0


def cmd_yearbook_layout(args: argparse.Namespace) -> int:
    try:
        from core.yearbook import YearbookLayoutEngine, get_default_mock_spec
    except (ImportError, ValueError):
        print("[ERROR] Failed to import core.yearbook", file=sys.stderr)
        return 1

    photo_inputs = getattr(args, "photos", []) or []
    collected_photos: list[str] = []
    for item in photo_inputs:
        if os.path.isdir(item):
            for fname in sorted(os.listdir(item)):
                fpath = os.path.join(item, fname)
                if os.path.isfile(fpath) and os.path.splitext(fname)[1].lower() in (".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff"):
                    collected_photos.append(os.path.abspath(fpath))
        elif os.path.isfile(item):
            collected_photos.append(os.path.abspath(item))

    title = getattr(args, "title", "KỶ YẾU TỐT NGHIỆP 2026")
    subtitle = getattr(args, "subtitle", "Hành Trình Rực Rỡ & Những Năm Tháng Thanh Xuân")
    author = getattr(args, "author", "Biện Trung Tín")
    degree = getattr(args, "degree", "Cử Nhân Kỹ Thuật Phần Mềm")
    date_str = getattr(args, "date", "Tháng 11 / 2026")
    preset = getattr(args, "preset", "photobook_square_25x25")
    bleed_mm = float(getattr(args, "bleed", 3.0))

    if not collected_photos and getattr(args, "mock", False):
        spec = get_default_mock_spec(
            title=title,
            subtitle=subtitle,
            author=author,
            degree=degree,
            date_str=date_str,
            preset_key=preset,
            bleed_mm=bleed_mm
        )
    else:
        spec = YearbookLayoutEngine.auto_layout(
            photo_paths=collected_photos,
            title=title,
            subtitle=subtitle,
            author=author,
            degree=degree,
            date_str=date_str,
            preset_key=preset,
            bleed_mm=bleed_mm,
        )

    out_path = args.output
    if out_path:
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(spec, f, ensure_ascii=False, indent=2)
        print(f"[OK] Generated Yearbook spec with {len(spec.get('pages', []))} spread(s) -> {out_path}")
    else:
        print(json.dumps(spec, ensure_ascii=False, indent=2))
    return 0


def cmd_yearbook_build(args: argparse.Namespace) -> int:
    try:
        from core.yearbook import YearbookPressEngine, YearbookHandoffPackager
    except (ImportError, ValueError):
        print("[ERROR] Failed to import core.yearbook", file=sys.stderr)
        return 1

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Spec file not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    out_path = args.output
    if not out_path:
        out_path = f"{os.path.splitext(spec_path)[0]}_bundle.zip"

    export_type = "handoff_zip"
    if getattr(args, "print_ready", False):
        export_type = "print_pdf"
    elif getattr(args, "spread_preview", False):
        export_type = "spread_pdf"
    elif getattr(args, "export_type", None):
        export_type = args.export_type.lower()

    if export_type == "print_pdf":
        actual_out = YearbookPressEngine.export_print_ready_pdf(spec, out_path)
        print(f"[OK] Rendered 300 DPI Print-Ready PDF -> {actual_out}")
    elif export_type == "spread_pdf":
        actual_out = YearbookPressEngine.export_spread_preview_pdf(spec, out_path)
        print(f"[OK] Rendered Spread Preview PDF -> {actual_out}")
    else:
        actual_out = YearbookHandoffPackager.export_handoff_bundle(spec, out_path)
        print(f"[OK] Assembled Designer Handoff Bundle ZIP -> {actual_out}")
    return 0


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ai_tools_cli",
        description=(
            "AI Document Editing Toolkit — Read/write DOCX and XLSX as JSON snapshots.\n"
            "Designed for Antigravity IDE to edit non-source-code files directly."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Read DOCX to JSON snapshot:
  python ai_tools_cli.py docx-read report.docx -o report_snap.json

  # Write modified snapshot back to DOCX (with original as style anchor):
  python ai_tools_cli.py docx-write report_snap.json --template report.docx -o report_v2.docx

  # Compare two snapshots (before AI edit vs after):
  python ai_tools_cli.py docx-diff before.json after.json -o diff.json

  # Read XLSX (specific sheet + range):
  python ai_tools_cli.py xlsx-read data.xlsx --sheet "Q1 Report" --range "A1:H50" -o snap.json

  # Write XLSX from snapshot (with original as template):
  python ai_tools_cli.py xlsx-write snap.json --template data.xlsx -o data_v2.xlsx
        """,
    )

    sub = parser.add_subparsers(dest="command", required=True)

    # --- docx-read ---
    p_dr = sub.add_parser("docx-read", help="Read DOCX → JSON snapshot")
    p_dr.add_argument("file", help="Path to source .docx file")
    p_dr.add_argument("-o", "--output", default=None,
                      help="Output JSON file path (prints to stdout if omitted)")
    p_dr.set_defaults(func=cmd_docx_read)

    # --- docx-write ---
    p_dw = sub.add_parser("docx-write", help="Write JSON snapshot → DOCX")
    p_dw.add_argument("snapshot", help="Path to .json snapshot file")
    p_dw.add_argument("--template", default=None,
                      help="Path to original .docx to use as style anchor (recommended)")
    p_dw.add_argument("-o", "--output", default=None,
                      help="Output .docx path (default: <snapshot>_written.docx)")
    p_dw.set_defaults(func=cmd_docx_write)

    # --- docx-inject ---
    p_di = sub.add_parser(
        "docx-inject",
        help="Surgically inject paragraphs, runs, and images into an existing DOCX without rewriting",
    )
    p_di.add_argument("file", help="Path to source/target .docx file")
    p_di.add_argument("spec", help="Path to injection spec .json file")
    p_di.add_argument(
        "-o", "--output", default=None,
        help="Output .docx path (defaults to in-place overwrite of source file)",
    )
    p_di.set_defaults(func=cmd_docx_inject)

    # --- docx-patch ---
    p_dp = sub.add_parser(
        "docx-patch",
        help="In-place template patching for DOCX (text placeholders and block insertion without wiping layout)",
    )
    p_dp.add_argument("template", help="Path to source .docx template file")
    p_dp.add_argument("patch", help="Path to patch spec .json file")
    p_dp.add_argument("-o", "--output", default=None, help="Output .docx path (default: <template>_patched.docx)")
    p_dp.set_defaults(func=cmd_docx_patch)

    # --- mermaid-render ---
    p_mr = sub.add_parser(
        "mermaid-render",
        help="Render Mermaid diagram from spec JSON to high-res PNG and optionally inject into DOCX",
    )
    p_mr.add_argument("spec", help="Path to mermaid diagram spec .json file")
    p_mr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_mr.add_argument("--inject", default=None,
                      help="Override target .docx to inject into")
    p_mr.add_argument("--theme", default=None,
                      help="Override theme preset ('clean_modern', 'enterprise', or json path)")
    p_mr.add_argument("--scale", type=int, default=None,
                      help="Override scale factor (default: 2)")
    p_mr.set_defaults(func=cmd_mermaid_render)

    # --- plantuml-render ---
    p_pr = sub.add_parser(
        "plantuml-render",
        help="Render PlantUML diagram from spec JSON to high-res PNG and optionally inject into DOCX",
    )
    p_pr.add_argument("spec", help="Path to plantuml diagram spec .json file")
    p_pr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_pr.add_argument("--inject", default=None,
                      help="Override target .docx to inject into")
    p_pr.add_argument("--dpi", type=int, default=300,
                      help="Override DPI (default: 300)")
    p_pr.set_defaults(func=cmd_plantuml_render)

    # --- diagram-render (Unified Dispatcher) ---
    p_dr = sub.add_parser(
        "diagram-render",
        help="Unified multi-engine diagram dispatcher (reads 'engine': 'canvas' | 'mermaid' | 'plantuml')",
    )
    p_dr.add_argument("spec", help="Path to unified diagram spec .json file")
    p_dr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_dr.add_argument("--inject", default=None,
                      help="Override target .docx to inject into")
    p_dr.add_argument("--scale", type=int, default=None,
                      help="Override scale factor for canvas engine (default: 3)")
    p_dr.set_defaults(func=cmd_diagram_render)

    # --- spec-render ---
    p_sr = sub.add_parser(
        "spec-render",
        help="Render Precision Technical & Architecture diagram from spec JSON to high-res PNG (Chromium 300+ DPI)",
    )
    p_sr.add_argument("spec", help="Path to precision diagram spec .json file")
    p_sr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_sr.add_argument("--scale", type=int, default=None,
                      help="Override scale factor (default: 3)")
    p_sr.set_defaults(func=cmd_spec_render)

    # --- diagram-editor ---
    p_de = sub.add_parser(
        "diagram-editor",
        help="Launch Interactive Canvas Editor for Precision Diagram JSON specs",
    )
    p_de.add_argument("spec", nargs="?", default=None, help="Path to precision diagram spec .json file (optional)")
    p_de.set_defaults(func=cmd_diagram_editor)


    # --- docx-page-numbers ---
    p_dpn = sub.add_parser("docx-page-numbers", help="Inject dynamic OpenXML page numbering into footers")
    p_dpn.add_argument("file", help="Path to .docx file")
    p_dpn.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dpn.add_argument("--font", default="Times New Roman", help="Font family (default: Times New Roman)")
    p_dpn.add_argument("--size", type=float, default=10.0, help="Font size in pt (default: 10.0)")
    p_dpn.set_defaults(func=cmd_docx_page_numbers)

    # --- docx-caption-fix ---
    p_dcf = sub.add_parser("docx-caption-fix", help="Normalize figure/table captions with En-dash and styling")
    p_dcf.add_argument("file", help="Path to .docx file")
    p_dcf.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dcf.add_argument("--font", default="Times New Roman", help="Font family (default: Times New Roman)")
    p_dcf.add_argument("--size", type=float, default=10.5, help="Font size in pt (default: 10.5)")
    p_dcf.set_defaults(func=cmd_docx_caption_fix)

    # --- docx-sanitize-icons ---
    p_dsi = sub.add_parser("docx-sanitize-icons", help="Remove emojis and special symbols from Word docx")
    p_dsi.add_argument("file", help="Path to .docx file")
    p_dsi.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dsi.set_defaults(func=cmd_docx_sanitize_icons)

    # --- docx-translate ---
    p_dt = sub.add_parser("docx-translate", help="Translate document while strictly protecting quoted terms")
    p_dt.add_argument("file", help="Path to .docx file")
    p_dt.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dt.add_argument("--sl", default="vi", help="Source language (default: vi)")
    p_dt.add_argument("--tl", default="en", help="Target language (default: en)")
    p_dt.set_defaults(func=cmd_docx_translate)

    # --- docx-diff (v4.0 Legal Redline & Snapshot Diff) ---
    p_dd = sub.add_parser(
        "docx-diff",
        help="Compare two Word documents (.docx) or JSON snapshots and generate diff report / track changes (ERR_DOCX_008)"
    )
    p_dd.add_argument("before", help="Path to original Word .docx document or 'before' JSON snapshot")
    p_dd.add_argument("after", help="Path to revised Word .docx document or 'after' JSON snapshot")
    p_dd.add_argument("-o", "--output", default=None, help="Output diff path (.docx or .json)")
    p_dd.add_argument("--author", default="AI Reviewer", help="Reviewer author name for track changes")
    p_dd.set_defaults(func=cmd_docx_diff)

    # --- xlsx-read ---
    p_xr = sub.add_parser("xlsx-read", help="Read XLSX → JSON snapshot (auto-hybrid, headless openpyxl fallback)")
    p_xr.add_argument("file", help="Path to source .xlsx file")
    p_xr.add_argument("--sheet", default=None,
                      help="Sheet name to export (exports all sheets if omitted)")
    p_xr.add_argument("--range", default=None,
                      help="Cell range to limit export, e.g. 'A1:H20' (requires --sheet)")
    p_xr.add_argument("--engine", default="auto", choices=["auto", "openpyxl", "xlwings"],
                      help="Reading engine: auto (default), openpyxl (headless), or xlwings (requires Excel)")
    p_xr.add_argument("-o", "--output", default=None,
                      help="Output JSON file path (prints to stdout if omitted)")
    p_xr.set_defaults(func=cmd_xlsx_read)

    # --- xlsx-write ---
    p_xw = sub.add_parser("xlsx-write", help="Write JSON snapshot → XLSX")
    p_xw.add_argument("snapshot", help="Path to .json snapshot file")
    p_xw.add_argument("--template", default=None,
                      help="Path to original .xlsx to use as base (optional)")
    p_xw.add_argument("-o", "--output", default=None,
                      help="Output .xlsx path (default: <snapshot>_written.xlsx)")
    p_xw.set_defaults(func=cmd_xlsx_write)

    # --- xlsx-mutate ---
    p_xm = sub.add_parser(
        "xlsx-mutate",
        help="Mutate Excel template in-place with row/col expansion, style cloning, and formula shifting",
    )
    p_xm.add_argument("template", help="Path to base .xlsx template file")
    p_xm.add_argument("spec", help="Path to mutation spec .json file")
    p_xm.add_argument("-o", "--output", default=None, help="Output .xlsx path (default: <template>_mutated.xlsx)")
    p_xm.add_argument("--validate", action="store_true", help="Automatically run 12 Quality Gates diff after mutating (Invariant E6)")
    p_xm.add_argument("--ref-sheet", default=None, help="Reference sheet for validation (default: auto-discover Example/Template)")
    p_xm.set_defaults(func=cmd_xlsx_mutate)

    # --- xlsx-validate ---
    p_xv = sub.add_parser(
        "xlsx-validate",
        help="Validate an Excel sheet against in-template reference sheet across 12 Quality Gates (Invariant E6)",
    )
    p_xv.add_argument("file", help="Path to .xlsx workbook")
    p_xv.add_argument("--target", required=True, help="Target sheet name to validate")
    p_xv.add_argument("--ref", default=None, help="Reference sheet name (default: auto-discover Example/Template)")
    p_xv.set_defaults(func=cmd_xlsx_validate)

    # --- pptx-create / pptx-build ---
    for cmd_name in ["pptx-create", "pptx-build"]:
        p_pc = sub.add_parser(cmd_name, help="Generate 16:9 PowerPoint deck from structured JSON spec")
        p_pc.add_argument("spec", nargs="?", default=None, help="Path to slide spec .json file")
        p_pc.add_argument("-s", "--spec", dest="spec_opt", default=None, help="Path to slide spec .json file (flag option)")
        p_pc.add_argument("-o", "--output", default=None, help="Output .pptx path")
        p_pc.add_argument("--theme", default="thesis_blue",
                          choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                          help="Visual theme (default: thesis_blue)")
        p_pc.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
        p_pc.set_defaults(func=cmd_pptx_create)

    # --- pptx-from-md ---
    p_pm = sub.add_parser("pptx-from-md", help="Compile Markdown slide deck ('---' delimiter) to .pptx")
    p_pm.add_argument("file", help="Path to markdown slides .md file")
    p_pm.add_argument("-o", "--output", default=None, help="Output .pptx path")
    p_pm.add_argument("--theme", default="thesis_blue",
                      choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                      help="Visual theme (default: thesis_blue)")
    p_pm.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
    p_pm.set_defaults(func=cmd_pptx_from_md)

    # --- docx-to-pptx ---
    p_dtp = sub.add_parser("docx-to-pptx", help="Synthesize 16:9 PowerPoint presentation deck from Word .docx document")
    p_dtp.add_argument("file", help="Path to source Word .docx file")
    p_dtp.add_argument("-o", "--output", default=None, help="Output .pptx path")
    p_dtp.add_argument("--spec-out", default=None, help="Path to save intermediate slide_spec.json")
    p_dtp.add_argument("--theme", default="thesis_blue",
                       choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                       help="Visual theme (default: thesis_blue)")
    p_dtp.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
    p_dtp.add_argument("--scope", default="auto",
                       choices=["auto", "executive", "standard", "comprehensive"],
                       help="Target presentation deck scope (default: auto)")
    p_dtp.set_defaults(func=cmd_docx_to_pptx)

    # --- pdf-to-pptx ---
    p_ptp = sub.add_parser("pdf-to-pptx", help="Synthesize 16:9 PowerPoint presentation deck directly from PDF document")
    p_ptp.add_argument("file", help="Path to source PDF file")
    p_ptp.add_argument("-o", "--output", default=None, help="Output .pptx path")
    p_ptp.add_argument("--spec-out", default=None, help="Path to save intermediate slide_spec.json")
    p_ptp.add_argument("--pages", default=None, help="Page range to extract, e.g. '1-10' or '1,3,5'")
    p_ptp.add_argument("--theme", default="thesis_blue",
                       choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                       help="Visual theme (default: thesis_blue)")
    p_ptp.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
    p_ptp.add_argument("--scope", default="auto",
                       choices=["auto", "executive", "standard", "comprehensive"],
                       help="Target presentation deck scope (default: auto)")
    p_ptp.set_defaults(func=cmd_pdf_to_pptx)


    # --- pptx-insert-diagram ---
    p_pid = sub.add_parser("pptx-insert-diagram", help="Insert technical diagram image into a PowerPoint slide")
    p_pid.add_argument("file", help="Path to .pptx file")
    p_pid.add_argument("image", help="Path to diagram image file (.png/.jpg)")
    p_pid.add_argument("--slide", type=int, default=None, help="Slide number (1-based); appends new slide if omitted")
    p_pid.add_argument("--title", default=None, help="Slide title")
    p_pid.add_argument("--caption", default=None, help="Figure caption underneath diagram")
    p_pid.add_argument("-o", "--output", default=None, help="Output .pptx path (default: in-place safe save)")
    p_pid.add_argument("--theme", default="corporate_blue", help="Visual theme for newly created slide")
    p_pid.set_defaults(func=cmd_pptx_insert_diagram)

    # --- pptx-inspect ---
    p_pi = sub.add_parser("pptx-inspect", help="Extract PowerPoint structure, slides, text and tables into JSON AST")
    p_pi.add_argument("file", help="Path to .pptx file")
    p_pi.add_argument("-o", "--output", default=None, help="Output JSON path (prints to stdout if omitted)")
    p_pi.set_defaults(func=cmd_pptx_inspect)

    # --- pptx-preview ---
    p_pp = sub.add_parser("pptx-preview", help="Export PowerPoint slides as high-res PNG previews for QA")
    p_pp.add_argument("file", help="Path to .pptx presentation file")
    p_pp.add_argument("-o", "--output-dir", default=None, help="Directory to save preview PNGs (default: <dir>/preview_slides)")
    p_pp.add_argument("--slides", default="all", help="Comma-separated slide indices, e.g. '1,5,11,15', or 'all' (default: all)")
    p_pp.add_argument("--dpi", type=int, default=150, help="Image resolution DPI (default: 150)")
    p_pp.set_defaults(func=cmd_pptx_preview)

    # --- felo-generate ---
    p_fg = sub.add_parser("felo-generate", help="Generate AI presentation via Felo Open Platform cloud service")
    p_fg.add_argument("query", help="Presentation prompt / topic")
    p_fg.add_argument("--file", default=None, help="Optional context file or image to upload")
    p_fg.add_argument("--theme", default=None, help="Felo PPT theme ID")
    p_fg.add_argument("--task-id", default=None, help="Resume polling an existing task")
    p_fg.set_defaults(func=cmd_felo_generate)

    # --- doctor ---
    p_doc = sub.add_parser("doctor", help="Run Zero-Setup Environment Doctor diagnostic suite")
    p_doc.add_argument("--json", action="store_true", help="Output machine-readable JSON format")
    p_doc.set_defaults(func=cmd_doctor)

    # --- docx-apply-preset ---
    p_dap = sub.add_parser("docx-apply-preset", help="Apply document preset (nd30, academic_thesis, corporate_report) to DOCX")
    p_dap.add_argument("file", help="Path to .docx file")
    p_dap.add_argument("--preset", default="nd30", help="Preset name (default: nd30)")
    p_dap.add_argument("--add-header", action="store_true", help="Insert administrative 2-column header (ND30)")
    p_dap.add_argument("--agency-parent", default="BỘ THÔNG TIN VÀ TRUYỀN THÔNG", help="Parent agency name")
    p_dap.add_argument("--agency-name", default="CỤC CÔNG NGHỆ THÔNG TIN", help="Agency name")
    p_dap.add_argument("--doc-no", default="Số: 01/QĐ-CNTT", help="Document number")
    p_dap.add_argument("--date", default="Hà Nội, ngày 18 tháng 10 năm 2026", help="Date & location string")
    p_dap.add_argument("-o", "--output", default=None, help="Output .docx path")
    p_dap.set_defaults(func=cmd_docx_apply_preset)

    # --- docx-sanitize-vietnamese ---
    p_dsv = sub.add_parser("docx-sanitize-vietnamese", help="Sanitize Vietnamese typography (Unicode NFC, punctuation spacing, En-dash) in DOCX")
    p_dsv.add_argument("file", help="Path to .docx file")
    p_dsv.add_argument("-o", "--output", default=None, help="Output .docx path")
    p_dsv.set_defaults(func=cmd_docx_sanitize_vietnamese)

    # --- pdf-tools (Everyday & Pro Toolkit) ---
    p_pdft = sub.add_parser("pdf-tools", help="PDF Pro & Everyday Toolkit: merge, split, rotate, watermark, compress, redact, protect, normalize-a4, giap-lai, extract-tables, images-to-pdf, export-images-zip, sign, flatten, reorganize")
    p_pdft.add_argument("action", choices=["merge", "split", "rotate", "clean", "watermark", "compress", "redact", "protect", "normalize-a4", "giap-lai", "extract-tables", "images-to-pdf", "export-images-zip", "sign", "flatten", "reorganize"], help="Action to perform")
    p_pdft.add_argument("-i", "--input", default=None, help="Input PDF file path")
    p_pdft.add_argument("--inputs", nargs="+", default=[], help="Multiple input PDF/image paths (for merge or images-to-pdf)")
    p_pdft.add_argument("-o", "--output", default=None, help="Output PDF path, Excel path, ZIP path, or directory")
    p_pdft.add_argument("--ranges", default=None, help="Page ranges for split (e.g. '1-3, 4-5')")
    p_pdft.add_argument("--angle", type=int, default=90, help="Rotation angle in degrees (90, 180, 270)")
    p_pdft.add_argument("--text", default="BẢN QUYỀN", help="Watermark text")
    p_pdft.add_argument("--opacity", type=float, default=0.25, help="Watermark opacity (0.0 - 1.0)")
    p_pdft.add_argument("--remove-blank", action="store_true", help="Remove blank pages when rotating or cleaning")
    p_pdft.add_argument("--patterns", default="cccd,phone,email", help="Comma-separated sensitive patterns for redaction (e.g. 'cccd,phone,email,bank_card,salary')")
    p_pdft.add_argument("--terms", default=None, help="Comma-separated custom keywords to redact")
    p_pdft.add_argument("--user-pwd", default=None, help="User password required to open protected PDF")
    p_pdft.add_argument("--owner-pwd", default=None, help="Owner password to manage protected PDF permissions")
    p_pdft.add_argument("--no-print", action="store_true", help="Disallow printing for protected PDF")
    p_pdft.add_argument("--allow-copy", action="store_true", help="Allow text copying for protected PDF")
    p_pdft.add_argument("--allow-modify", action="store_true", help="Allow modifications for protected PDF")
    p_pdft.add_argument("--orientation", default="auto", choices=["auto", "portrait", "landscape"], help="Orientation for A4 normalization / images-to-pdf")
    p_pdft.add_argument("--page-size", default="A4", help="Page size for images-to-pdf (default: A4)")
    p_pdft.add_argument("--margin", type=float, default=36.0, help="Margin padding in points for images-to-pdf (default: 36)")
    p_pdft.add_argument("--bates-prefix", default=None, help="Prefix for Bates Numbering (e.g. 'HD-2026')")
    p_pdft.add_argument("--stamp", default=None, help="Path to stamp or signature image")
    p_pdft.add_argument("--signature", default=None, help="Alias for --stamp")
    p_pdft.add_argument("--side", default="right", choices=["right", "left"], help="Side for Giáp Lai stamp (right or left)")
    p_pdft.add_argument("--page", type=int, default=1, help="Target page number (1-based) for sign/stamp")
    p_pdft.add_argument("--x-pct", type=float, default=0.65, help="X coordinate percentage (0.0 to 1.0)")
    p_pdft.add_argument("--y-pct", type=float, default=0.75, help="Y coordinate percentage (0.0 to 1.0)")
    p_pdft.add_argument("--w-pct", type=float, default=0.25, help="Width percentage (0.0 to 1.0)")
    p_pdft.add_argument("--h-pct", type=float, default=0.15, help="Height percentage (0.0 to 1.0)")
    p_pdft.add_argument("--strict", action="store_true", help="Strict raster flattening (render pages to static images)")
    p_pdft.add_argument("--dpi", type=int, default=300, help="Image export DPI (default: 300)")
    p_pdft.add_argument("--format", default="png", choices=["png", "jpg", "jpeg"], help="Image export format")
    p_pdft.add_argument("--plan", default=None, help="Reorganization plan string (e.g. '1,3,blank,2' or JSON)")
    p_pdft.set_defaults(func=cmd_pdf_tools)

    # --- docx-mail-merge (v4.0) ---
    p_mm = sub.add_parser("docx-mail-merge", help="Batch Mail Merge Studio: Word template + Excel/JSON dataset ➔ batch DOCX / ZIP")
    p_mm.add_argument("template", help="Path to Word .docx template with {{Token}} fields")
    p_mm.add_argument("-d", "--data", default=None, help="Path to data source (.xlsx or .json)")
    p_mm.add_argument("-o", "--output", default=None, help="Output directory path (default: ./mail_merge_output)")
    p_mm.add_argument("--pattern", default=None, help="Filename pattern, e.g. 'HDLD_{{MaNV}}_{{HoTen}}.docx'")
    p_mm.add_argument("--zip", default=None, help="Optional ZIP bundle output path, e.g. 'contracts.zip'")
    p_mm.add_argument("--inspect", action="store_true", help="Inspect and list all {{Tokens}} in template without merging")
    p_mm.set_defaults(func=cmd_docx_mail_merge)


    # --- ocr-scan (v4.0) ---
    p_ocr = sub.add_parser("ocr-scan", help="Vietnamese OCR Pipeline: extract text & tables from scanned PDFs/images (ERR_CONV_005)")
    p_ocr.add_argument("file", help="Path to scanned PDF or image file (.png, .jpg, etc.)")
    p_ocr.add_argument("-o", "--output", default=None, help="Output path (.md, .txt, or .docx)")
    p_ocr.add_argument("--lang", default="vie+eng", help="OCR languages (default: vie+eng)")
    p_ocr.add_argument("--no-red-stamp", action="store_true", help="Disable HSV red-stamp filter")
    p_ocr.add_argument("--no-deskew", action="store_true", help="Disable automatic deskewing")
    p_ocr.set_defaults(func=cmd_ocr_scan)

    # --- qa-audit (v4.0) ---
    p_qa = sub.add_parser("qa-audit", help="Run Enterprise Unified QA Diagnostic (DOCX, XLSX, PPTX, Draw.io, Scans)")
    p_qa.add_argument("--xlsx", default=None, help="Path to Excel workbook (.xlsx)")
    p_qa.add_argument("--docx", default=None, help="Path to Word document (.docx)")
    p_qa.add_argument("--pptx", default=None, help="Path to PowerPoint presentation (.pptx)")
    p_qa.add_argument("--diagram", default=None, help="Path to Draw.io / vector diagram (.drawio, .xml, .png)")
    p_qa.add_argument("--all", default=None, help="Directory containing documents to recursively scan")
    p_qa.add_argument("--mode", choices=["diagnose", "audit-fix"], default="diagnose",
                      help="Operation mode: 'diagnose' (read-only audit) or 'audit-fix' (safe repair)")
    p_qa.add_argument("--in-place", action="store_true",
                      help="Allow in-place repair (creates a .bak file before modifying)")
    p_qa.add_argument("--format", choices=["terminal", "json", "sarif"], default="terminal",
                      help="Output report format")
    p_qa.add_argument("-o", "--output", default=None, help="Optional output file path for the report/log")
    p_qa.add_argument("--security", action="store_true", help="Enable Gate 13: Enterprise Document Security & PII Audit (Cloudflare Harness)")
    p_qa.add_argument("--sanitize", action="store_true", help="Sanitize detected security threats and PII")
    p_qa.set_defaults(func=cmd_qa_audit)

    # --- doc-security (Cloudflare Harness Document Security & Sanitizer) ---
    p_sec = sub.add_parser("doc-security", help="Audit & sanitize document security threats, DDE/formulas, and PII (Cloudflare Harness)")
    p_sec.add_argument("file", help="Path to document to audit (.docx, .xlsx, .pptx, .pdf, .csv, .drawio)")
    p_sec.add_argument("--sanitize", action="store_true", help="Sanitize and defuse detected security threats")
    p_sec.add_argument("-o", "--output", default=None, help="Output path for sanitized file or JSON report")
    p_sec.add_argument("--format", choices=["terminal", "json"], default="terminal", help="Report format")
    p_sec.set_defaults(func=cmd_doc_security)

    # --- convert (Unified Converter Dispatcher) ---
    p_conv = sub.add_parser("convert", help="Universal multi-format document conversion (PDF ↔ DOCX ↔ MD ↔ PPTX ↔ XLSX)")
    p_conv.add_argument("input", help="Source file path (.pdf, .docx, .md)")
    p_conv.add_argument("-o", "--output", default=None, help="Output file path (.docx, .pdf, .md, .pptx, .xlsx)")
    p_conv.add_argument("-t", "--target", choices=["docx", "pdf", "md", "pptx", "xlsx"], default=None, help="Target format extension")
    p_conv.add_argument("--style", default=None, help="Path to style YAML stylesheet (for Markdown/HTML conversions)")
    p_conv.add_argument("--pages", default=None, help="Page range (e.g. '1-5, 8')")
    p_conv.add_argument("--scope", choices=["auto", "executive", "standard", "comprehensive"], default="auto", help="Presentation deck scope (for PPTX target)")
    p_conv.set_defaults(func=cmd_convert)

    # --- md-toc (Interactive Markdown Table of Contents) ---
    p_toc = sub.add_parser("md-toc", help="Generate or update interactive Table of Contents for Markdown (.md) documents")
    p_toc.add_argument("file", help="Path to Markdown (.md) file")
    p_toc.add_argument("--max-level", type=int, default=3, help="Maximum heading level (default: 3)")
    p_toc.add_argument("--min-level", type=int, default=2, help="Minimum heading level (default: 2)")
    p_toc.add_argument("--in-place", action="store_true", help="Update TOC directly in the file")
    p_toc.add_argument("--check", action="store_true", help="Print generated TOC preview without modifying file")
    p_toc.add_argument("-o", "--output", default=None, help="Save TOC to separate file")
    p_toc.set_defaults(func=cmd_md_toc)

    # --- xlsx-set (Atomic Excel Cell Update) ---
    p_xs = sub.add_parser("xlsx-set", help="Atomic Excel tool: set cell value or formula with merged-cell protection")
    p_xs.add_argument("file", help="Path to .xlsx file")
    p_xs.add_argument("--sheet", default=None, help="Worksheet name (default: active sheet)")
    p_xs.add_argument("--cell", required=True, help="Cell coordinate, e.g. 'A1', 'C7'")
    p_xs.add_argument("--value", default=None, help="Cell value to set")
    p_xs.add_argument("--formula", default=None, help="Excel formula string (e.g. '=SUM(C2:C6)')")
    p_xs.add_argument("-o", "--output", default=None, help="Output .xlsx path (default: in-place safe save)")
    p_xs.set_defaults(func=cmd_xlsx_set)

    # --- xlsx-insert-rows (Atomic Excel Row Expansion) ---
    p_xi = sub.add_parser("xlsx-insert-rows", help="Atomic Excel tool: insert blank or prototype-cloned rows with formula shifting")
    p_xi.add_argument("file", help="Path to .xlsx file")
    p_xi.add_argument("--sheet", default=None, help="Worksheet name (default: active sheet)")
    p_xi.add_argument("--start-row", type=int, required=True, help="Start row index (1-based)")
    p_xi.add_argument("--count", type=int, default=1, help="Number of rows to insert (default: 1)")
    p_xi.add_argument("--proto", type=int, default=None, help="Prototype row index to clone styles/formulas from")
    p_xi.add_argument("--no-shift-formulas", action="store_true", help="Disable automatic formula range shifting")
    p_xi.add_argument("-o", "--output", default=None, help="Output .xlsx path (default: in-place safe save)")
    p_xi.set_defaults(func=cmd_xlsx_insert_rows)

    # --- docx-tokens (Atomic Word Placeholders Inspection) ---
    p_dtk = sub.add_parser("docx-tokens", help="Atomic Word tool: inspect template placeholders and variables")
    p_dtk.add_argument("template", help="Path to .docx template file")
    p_dtk.add_argument("--json", action="store_true", help="Output machine-readable JSON format")
    p_dtk.set_defaults(func=cmd_docx_tokens)

    # --- docx-replace (Atomic Word Text Replacement) ---
    p_drp = sub.add_parser("docx-replace", help="Atomic Word tool: replace text tokens preserving run styles")
    p_drp.add_argument("file", help="Path to .docx file")
    p_drp.add_argument("--data", required=True, help="Path to JSON file or raw JSON string with token replacements")
    p_drp.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_drp.set_defaults(func=cmd_docx_replace)

    # --- xlsx-formulas (Excel Formula Playbook & Cheat Sheet) ---
    p_xf = sub.add_parser("xlsx-formulas", help="Browse, filter, and export the Excel Formula Playbook & Cheat Sheet")
    p_xf.add_argument("--category", choices=["all", "lookup", "aggregation", "text", "logic", "dynamic", "financial"], default="all", help="Filter by formula category")
    p_xf.add_argument("--search", default=None, help="Search by function name, keywords, or use cases")
    p_xf.add_argument("--export", default=None, help="Export dual-sheet cheat sheet Excel file (.xlsx)")
    p_xf.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    p_xf.set_defaults(func=cmd_xlsx_formulas)

    # --- yearbook-layout (Digital Yearbook Auto-Layout) ---
    p_yl = sub.add_parser("yearbook-layout", help="Generate declarative Yearbook Project Spec JSON from photos directory or list")
    p_yl.add_argument("photos", nargs="*", default=[], help="Image file paths or directory of photos")
    p_yl.add_argument("-o", "--output", default=None, help="Output spec .json path")
    p_yl.add_argument("--title", default="KỶ YẾU TỐT NGHIỆP 2026", help="Album title")
    p_yl.add_argument("--subtitle", default="Hành Trình Rực Rỡ & Những Năm Tháng Thanh Xuân", help="Album subtitle")
    p_yl.add_argument("--author", default="Biện Trung Tín", help="Owner / Graduate name")
    p_yl.add_argument("--degree", default="Cử Nhân Kỹ Thuật Phần Mềm", help="Academic degree / major")
    p_yl.add_argument("--date", default="Tháng 11 / 2026", help="Graduation date string")
    p_yl.add_argument("--preset", default="photobook_square_25x25",
                      choices=["photobook_square_25x25", "a4_landscape", "a4_portrait", "magazine_b5"],
                      help="Photobook dimension preset")
    p_yl.add_argument("--bleed", type=float, default=3.0, help="Pre-press bleed margin in mm (default: 3.0)")
    p_yl.add_argument("--mock", action="store_true", help="Generate complete mock demonstration album if no photos given")
    p_yl.set_defaults(func=cmd_yearbook_layout)

    # --- yearbook-build (Pre-press PDF & Handoff Bundle Packager) ---
    p_yb = sub.add_parser("yearbook-build", help="Build print-ready PDF (300 DPI, Bleed, Crop Marks) or Designer Handoff ZIP")
    p_yb.add_argument("spec", help="Path to yearbook project spec .json file")
    p_yb.add_argument("-o", "--output", default=None, help="Output destination (.zip or .pdf)")
    p_yb.add_argument("--print-ready", action="store_true", help="Export single-page 300 DPI print-ready PDF with Bleed & Crop Marks")
    p_yb.add_argument("--spread-preview", action="store_true", help="Export 2-page spread preview PDF for client inspection")
    p_yb.add_argument("--handoff", action="store_true", help="Export full 5-in-1 Designer Handoff Bundle (.zip)")
    p_yb.add_argument("--export-type", choices=["handoff_zip", "print_pdf", "spread_pdf"], default="handoff_zip", help="Explicit export type")
    p_yb.set_defaults(func=cmd_yearbook_build)

    from cli.v46_commands import register_v46_subcommands
    register_v46_subcommands(sub, p_pdft, p_sec, cmd_pdf_tools, cmd_doc_security)

    return parser


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


def main_spec_render() -> int:
    parser = argparse.ArgumentParser(description="Render Precision Technical Diagram from spec JSON to PNG")
    parser.add_argument("spec", help="Path to diagram spec JSON file")
    parser.add_argument("-o", "--output", default=None, help="Output PNG path")
    parser.add_argument("--scale", type=float, default=3.0, help="Resolution scale factor (default: 3.0)")
    args = parser.parse_args()
    return cmd_spec_render(args)


if __name__ == "__main__":
    sys.exit(main())
