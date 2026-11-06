"""
cli/print_commands.py — Enterprise Silent Printing & Wireframe CLI Dispatcher (Office Core).

Pure Office Document & Technical Diagram capabilities:
1. Enterprise Printing:
   - print-list: Discover and inspect system physical & virtual printers
   - print-doc: Silent document printing with ink-saver, booklet optimizer, margin guard
2. UI Wireframe Mockup:
   - wireframe-render: Render PlantUML Salt UI Wireframes to 300+ DPI PNG/SVG
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any, Dict


def cmd_wireframe_render(args: argparse.Namespace) -> int:
    from core.diagrams.salt_wireframe import SaltWireframeEngine

    if getattr(args, "list_templates", False):
        templates = SaltWireframeEngine.get_templates()
        print("=" * 70)
        print("  DANH MỤC 6 MẪU WIREFRAME NGHIỆP VỤ CHUẨN (BA WIREFRAME TEMPLATES)")
        print("=" * 70)
        for idx, (tid, tinfo) in enumerate(templates.items(), start=1):
            print(f"  {idx}. {tid:<16} | [{tinfo['category']}] {tinfo['title']}")
            print(f"     Mô tả: {tinfo['description']}")
        print("=" * 70)
        return 0

    src = getattr(args, "spec", None) or getattr(args, "template", None)
    if not src:
        print("[ERROR] Vui lòng truyền đường dẫn file spec, mã Salt hoặc cờ --template <tên_mẫu>.", file=sys.stderr)
        return 1

    fmt = getattr(args, "format", "png").lower()
    scale = float(getattr(args, "scale", 2.0))
    inject_doc = getattr(args, "inject_doc", None)

    try:
        if os.path.isfile(src):
            with open(src, "r", encoding="utf-8") as f:
                content = f.read().strip()
            try:
                spec_data: Any = json.loads(content)
            except Exception:
                spec_data = content
        else:
            spec_data = src

        if inject_doc:
            heading = getattr(args, "heading", None)
            caption = getattr(args, "caption", None) or "Hình: Phác thảo giao diện màn hình (UI Wireframe Mockup)"
            out_docx = args.output if (args.output and args.output.lower().endswith(".docx")) else None
            res_doc = SaltWireframeEngine.inject_wireframe_into_docx(
                docx_path=inject_doc,
                spec_or_code=spec_data,
                output_path=out_docx,
                heading=heading,
                caption=caption,
            )
            print(
                f"[OK] Wireframe injected into Word DOCX ({res_doc['width_cm']}x{res_doc['height_cm']} cm, "
                f"margin guard <= {res_doc['usable_margin_cm']} cm) -> {res_doc['docx_path']}"
            )
            return 0

        if fmt == "svg" or (args.output and args.output.lower().endswith(".svg")):
            output_svg = args.output or "output/wireframe.svg"
            res_svg = SaltWireframeEngine.render_salt_to_svg(spec_data, output_path=output_svg)
            print(
                f"[OK] Wireframe SVG rendered ({res_svg['dimensions_px'][0]}x{res_svg['dimensions_px'][1]} px) -> {res_svg['svg_path']}"
            )
            return 0

        if fmt == "both":
            base_out = args.output or "output/wireframe.png"
            stem = os.path.splitext(base_out)[0]
            svg_out = stem + ".svg"
            png_out = stem + ".png"
            res_svg = SaltWireframeEngine.render_salt_to_svg(spec_data, output_path=svg_out)
            res_png = SaltWireframeEngine.render_salt_to_png(spec_data, png_out, scale=scale)
            print(
                f"[OK] Wireframe Dual-Rendered ({res_png['dimensions_px'][0]}x{res_png['dimensions_px'][1]} px) -> "
                f"SVG: {res_svg['svg_path']} | PNG: {res_png['png_path']}"
            )
            return 0

        output_path = args.output or "output/wireframe.png"
        res = SaltWireframeEngine.render_salt_to_png(spec_data, output_path, scale=scale)
        print(f"[OK] Wireframe rendered successfully ({res['dimensions_px'][0]}x{res['dimensions_px'][1]} px) -> {res['png_path']}")
        return 0
    except Exception as exc:
        print(f"[ERROR] wireframe-render failed: {exc}", file=sys.stderr)
        return 2


def cmd_print_list(args: argparse.Namespace) -> int:
    from core.printing import PrinterManager

    try:
        printers = PrinterManager.list_printers()
        if getattr(args, "json", False):
            print(json.dumps(printers, ensure_ascii=False, indent=2))
        else:
            print("=" * 70)
            print("  DANH SÁCH MÁY IN HỆ THỐNG (SYSTEM HARDWARE & VIRTUAL PRINTERS)")
            print("=" * 70)
            for idx, p in enumerate(printers, start=1):
                star = " [DEFAULT]" if p.get("is_default") else ""
                v_tag = " (Virtual)" if p.get("is_virtual") else " (Hardware)"
                print(f"  {idx}. {p['name']}{star}{v_tag}")
                if p.get("description"):
                    print(f"     Mô tả: {p['description']}")
            print("=" * 70)
        return 0
    except Exception as exc:
        print(f"[ERROR] print-list failed: {exc}", file=sys.stderr)
        return 2


def cmd_print_doc(args: argparse.Namespace) -> int:
    from core.printing import PrinterManager

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] File to print not found: {src}", file=sys.stderr)
        return 1

    try:
        opt_options = {
            "grayscale": getattr(args, "grayscale", False),
            "vector_grayscale": not getattr(args, "raster_grayscale", False),
            "dpi": int(getattr(args, "dpi", 300)),
            "n_up": int(getattr(args, "n_up", 1)),
            "booklet": getattr(args, "booklet", False),
            "creep_compensation": getattr(args, "creep", False),
            "paper_thickness_pt": float(getattr(args, "paper_thickness", 0.28)),
            "watermark": getattr(args, "watermark", None),
            "margin_guard_mm": float(getattr(args, "margin_guard", 5.0)),
        }
        res = PrinterManager.print_file(
            file_path=src,
            printer_name=getattr(args, "printer", None),
            copies=int(getattr(args, "copies", 1)),
            optimizer_options=opt_options,
            tray_id=getattr(args, "tray", None),
            duplex=getattr(args, "duplex", None),
        )
        print(f"[OK] Print Job Dispatched: {res['job_id']} -> Printer: {res['printer']} ({res['status']})")
        return 0
    except Exception as exc:
        print(f"[ERROR] print-doc failed: {exc}", file=sys.stderr)
        return 2


def register_print_subcommands(sub: argparse._SubParsersAction) -> None:
    """Registers pure Office Printing & Wireframe subcommands."""

    # 1. wireframe-render
    p_wf = sub.add_parser("wireframe-render", help="Render PlantUML Salt UI Wireframe mockup to SVG/PNG or inject into Word")
    p_wf.add_argument("spec", nargs="?", default=None, help="Path to wireframe JSON spec file, raw Salt code, or template ID")
    p_wf.add_argument("-o", "--output", default=None, help="Output file path (.png, .svg, or .docx)")
    p_wf.add_argument("-f", "--format", default="png", choices=["png", "svg", "both"], help="Output format: png, svg, or both")
    p_wf.add_argument(
        "-t",
        "--template",
        default=None,
        choices=["login_auth", "crud_table", "wizard_form", "modal_confirm", "mobile_screen", "dashboard_kpi"],
        help="Use one of 6 built-in BA wireframe templates",
    )
    p_wf.add_argument("--list-templates", action="store_true", help="List all 6 built-in BA wireframe templates")
    p_wf.add_argument("--inject-doc", default=None, help="Target Word (.docx) file to inject the wireframe into (ERR_DOCX_005)")
    p_wf.add_argument("--heading", default=None, help="Target heading in Word document")
    p_wf.add_argument("--caption", default=None, help="Figure caption text in Word document")
    p_wf.add_argument("--scale", type=float, default=2.0, help="Scale factor (default: 2.0)")
    p_wf.set_defaults(func=cmd_wireframe_render)

    # 2. print-list
    p_plist = sub.add_parser("print-list", help="Discover and list available system physical & virtual printers")
    p_plist.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    p_plist.set_defaults(func=cmd_print_list)

    # 3. print-doc
    p_pdoc = sub.add_parser("print-doc", help="Silent document printing with vector grayscale, booklet creep & DEVMODE tray")
    p_pdoc.add_argument("file", help="Path to document (.pdf, .docx, .xlsx)")
    p_pdoc.add_argument("--printer", default=None, help="Target printer name (default: system default printer)")
    p_pdoc.add_argument("--copies", type=int, default=1, help="Number of copies (default: 1)")
    p_pdoc.add_argument("--grayscale", action="store_true", help="Convert pages to grayscale (toner saver)")
    p_pdoc.add_argument("--raster-grayscale", action="store_true", help="Use rasterized grayscale instead of native vector /DeviceGray")
    p_pdoc.add_argument("--dpi", type=int, default=300, choices=[150, 300, 600], help="Raster fallback DPI (default: 300)")
    p_pdoc.add_argument("--n-up", type=int, default=1, choices=[1, 2, 4], help="N-Up pages per sheet (1, 2, 4)")
    p_pdoc.add_argument("--booklet", action="store_true", help="Arrange pages in 2-sided saddle-stitch booklet order")
    p_pdoc.add_argument("--creep", action="store_true", help="Enable Booklet Shingling Creep compensation for inner sheets")
    p_pdoc.add_argument("--paper-thickness", type=float, default=0.28, help="Paper thickness in pt for creep calculation (default: 0.28pt ~ 80gsm)")
    p_pdoc.add_argument("--tray", type=int, default=None, help="Hardware paper tray bin ID (validated via DeviceCapabilities DC_BINS Guard G2)")
    p_pdoc.add_argument("--duplex", default=None, choices=["simplex", "vertical", "horizontal"], help="Hardware duplex mode")
    p_pdoc.add_argument("--watermark", default=None, help="Security watermark text across pages")
    p_pdoc.add_argument("--margin-guard", type=float, default=5.0, help="Safety margin in mm (ERR_PRINT_001)")
    p_pdoc.set_defaults(func=cmd_print_doc)

