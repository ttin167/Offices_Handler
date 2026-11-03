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

    src = args.spec
    output_path = args.output or "output/wireframe.png"
    scale = float(getattr(args, "scale", 2.0))

    try:
        if os.path.isfile(src):
            with open(src, "r", encoding="utf-8") as f:
                content = f.read().strip()
            try:
                spec_data = json.loads(content)
            except Exception:
                spec_data = content
        else:
            spec_data = src

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
            "n_up": int(getattr(args, "n_up", 1)),
            "booklet": getattr(args, "booklet", False),
            "watermark": getattr(args, "watermark", None),
            "margin_guard_mm": float(getattr(args, "margin_guard", 5.0)),
        }
        res = PrinterManager.print_file(
            file_path=src,
            printer_name=getattr(args, "printer", None),
            copies=int(getattr(args, "copies", 1)),
            optimizer_options=opt_options
        )
        print(f"[OK] Print Job Dispatched: {res['job_id']} -> Printer: {res['printer']} ({res['status']})")
        return 0
    except Exception as exc:
        print(f"[ERROR] print-doc failed: {exc}", file=sys.stderr)
        return 2


def register_print_subcommands(sub: argparse._SubParsersAction) -> None:
    """Registers pure Office Printing & Wireframe subcommands."""

    # 1. wireframe-render
    p_wf = sub.add_parser("wireframe-render", help="Render PlantUML Salt UI Wireframe mockup to 300+ DPI PNG")
    p_wf.add_argument("spec", help="Path to wireframe JSON spec file or raw Salt code")
    p_wf.add_argument("-o", "--output", default=None, help="Output PNG path")
    p_wf.add_argument("--scale", type=float, default=2.0, help="Scale factor (default: 2.0)")
    p_wf.set_defaults(func=cmd_wireframe_render)

    # 2. print-list
    p_plist = sub.add_parser("print-list", help="Discover and list available system physical & virtual printers")
    p_plist.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    p_plist.set_defaults(func=cmd_print_list)

    # 3. print-doc
    p_pdoc = sub.add_parser("print-doc", help="Silent document printing with ink-saver & booklet optimizer")
    p_pdoc.add_argument("file", help="Path to document (.pdf, .docx, .xlsx)")
    p_pdoc.add_argument("--printer", default=None, help="Target printer name (default: system default printer)")
    p_pdoc.add_argument("--copies", type=int, default=1, help="Number of copies (default: 1)")
    p_pdoc.add_argument("--grayscale", action="store_true", help="Convert pages to grayscale (toner saver)")
    p_pdoc.add_argument("--n-up", type=int, default=1, choices=[1, 2, 4], help="N-Up pages per sheet (1, 2, 4)")
    p_pdoc.add_argument("--booklet", action="store_true", help="Arrange pages in 2-sided booklet order")
    p_pdoc.add_argument("--watermark", default=None, help="Security watermark text across pages")
    p_pdoc.add_argument("--margin-guard", type=float, default=5.0, help="Safety margin in mm (ERR_PRINT_001)")
    p_pdoc.set_defaults(func=cmd_print_doc)

