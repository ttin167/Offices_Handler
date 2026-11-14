"""
cli/print_commands.py — Enterprise Silent Printing CLI Dispatcher (Office Core).

Pure Office Document Printing capabilities:
- print-list: Discover and inspect system physical & virtual printers
- print-status: Inspect printer hardware status & spooler queue
- print-doc: Silent document printing with ink-saver, booklet optimizer, margin guard
"""

from __future__ import annotations

import argparse
import json
import os
import sys


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


def cmd_print_status(args: argparse.Namespace) -> int:
    from core.printing import PrinterManager

    try:
        target = getattr(args, "printer", None)
        status_info = PrinterManager.get_printer_status(target)
        if getattr(args, "json", False):
            print(json.dumps(status_info, ensure_ascii=False, indent=2))
        else:
            print("=" * 70)
            print(f"  TRẠNG THÁI MÁY IN: {status_info['printer']}")
            print("=" * 70)
            print(f"  Trạng thái:     {status_info['status']} ({status_info['message']})")
            print(f"  Hàng đợi:       {status_info['c_jobs']} lệnh in")
            print(f"  Hỗ trợ Duplex:  {'Có' if status_info.get('supports_duplex') else 'Không'}")
            if status_info.get("hardware_warning"):
                print(f"  [CẢNH BÁO]:     {status_info['hardware_warning']}")
            if status_info.get("jobs"):
                print("  Danh sách tác vụ đang xếp hàng:")
                for j in status_info["jobs"]:
                    print(f"    - #{j.get('job_id')}: {j.get('document')} ({j.get('pages_printed')}/{j.get('total_pages')} trang)")
            print("=" * 70)
        return 0
    except Exception as exc:
        print(f"[ERROR] print-status failed: {exc}", file=sys.stderr)
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
            "page_range": getattr(args, "page_range", None),
            "page_size": getattr(args, "page_size", "a4"),
            "scaling": getattr(args, "scaling", "fit"),
        }
        res = PrinterManager.print_file(
            file_path=src,
            printer_name=getattr(args, "printer", None),
            copies=int(getattr(args, "copies", 1)),
            optimizer_options=opt_options,
            tray_id=getattr(args, "tray", None),
            duplex=getattr(args, "duplex", None),
            page_range=getattr(args, "page_range", None),
            page_size=getattr(args, "page_size", "a4"),
            scaling=getattr(args, "scaling", "fit"),
        )
        print(f"[OK] Print Job Dispatched: {res['job_id']} -> Printer: {res['printer']} ({res['status']})")
        if res.get("hardware_warning"):
            print(f"[WARNING] Hardware Alert: {res['hardware_warning']}")
        return 0
    except Exception as exc:
        print(f"[ERROR] print-doc failed: {exc}", file=sys.stderr)
        return 2


def register_print_subcommands(sub: argparse._SubParsersAction) -> None:
    """Registers pure Office Printing subcommands."""

    # 1. print-list
    p_plist = sub.add_parser("print-list", help="Discover and list available system physical & virtual printers")
    p_plist.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    p_plist.set_defaults(func=cmd_print_list)

    # 3. print-status
    p_pstat = sub.add_parser("print-status", help="Inspect printer hardware status, paper out, jam, and spooler queue")
    p_pstat.add_argument("--printer", default=None, help="Target printer name (default: system default printer)")
    p_pstat.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    p_pstat.set_defaults(func=cmd_print_status)

    # 4. print-doc
    p_pdoc = sub.add_parser("print-doc", help="Silent document printing with vector grayscale, booklet creep & DEVMODE tray")
    p_pdoc.add_argument("file", help="Path to document (.pdf, .docx, .xlsx)")
    p_pdoc.add_argument("--printer", default=None, help="Target printer name (default: system default printer)")
    p_pdoc.add_argument("--copies", type=int, default=1, help="Number of copies (default: 1)")
    p_pdoc.add_argument("--page-range", default=None, help="Page range expression (e.g. '1-5, 8, 11-20' or 'all')")
    p_pdoc.add_argument("--page-size", default="a4", choices=["a4", "letter", "a3"], help="Target paper size (default: a4)")
    p_pdoc.add_argument("--scaling", default="fit", choices=["fit", "actual"], help="Page scaling mode (default: fit)")
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

