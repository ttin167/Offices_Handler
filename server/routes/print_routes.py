# -*- coding: utf-8 -*-
"""
server.routes.print_routes — Enterprise Hardware & Virtual Printing API Endpoints.
"""

from __future__ import annotations

import os
import sys
from typing import Any

from .base import router, send_json, send_error, read_json_body


@router.get("/api/print/printers")
def handle_list_printers(handler: Any) -> None:
    """Lists all available hardware and virtual spooler printers."""
    try:
        from core.printing.printer_manager import PrinterManager
        printers = PrinterManager.list_printers()
        default_printer = PrinterManager.get_default_printer()
        send_json(handler, {
            "success": True,
            "printers": printers,
            "default_printer": default_printer,
            "total": len(printers),
        })
    except Exception as e:
        send_error(handler, f"Lỗi liệt kê danh sách máy in: {e}", status=500)


@router.get("/api/print/status")
def handle_printer_status(handler: Any) -> None:
    """Queries hardware status, paper out, paper jam and queue for a printer."""
    try:
        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(getattr(handler, "path", ""))
        qs = parse_qs(parsed.query)
        printer_param = qs.get("printer", [None])[0]

        from core.printing.printer_manager import PrinterManager
        status_info = PrinterManager.get_printer_status(printer_param)
        send_json(handler, {
            "success": True,
            **status_info
        })
    except Exception as e:
        send_error(handler, f"Lỗi truy vấn trạng thái máy in: {e}", status=500)


@router.post("/api/print/job")
def handle_print_job(handler: Any) -> None:
    """Submits silent print job with print optimization chain."""
    body = read_json_body(handler)
    file_path = body.get("file_path") or body.get("input_path") or body.get("file") or body.get("filepath")
    if not file_path:
        send_error(handler, "Thiếu tham số bắt buộc: file_path", status=400)
        return

    # Check file exists in uploads or workspace
    if not os.path.isfile(file_path):
        # Check in uploads dir
        uploads_dir = os.path.abspath(os.path.join(os.getcwd(), "uploads"))
        candidate = os.path.join(uploads_dir, os.path.basename(file_path))
        if os.path.isfile(candidate):
            file_path = candidate
        else:
            send_error(handler, f"Tệp không tồn tại: {file_path}", status=404)
            return

    printer_name = body.get("printer_name") or body.get("printer")
    copies = int(body.get("copies", 1))
    job_title = body.get("job_title")
    tray_id = body.get("tray_id")
    duplex = body.get("duplex")
    page_range = body.get("page_range")
    page_size = body.get("page_size")
    scaling = body.get("scaling", "fit")

    optimizer_options = {
        "grayscale": body.get("grayscale", False),
        "vector_grayscale": body.get("vector_grayscale", True),
        "dpi": int(body.get("dpi", 300)),
        "n_up": int(body.get("n_up", 1)),
        "booklet": body.get("booklet", False),
        "creep_compensation": body.get("creep_compensation", True),
        "paper_thickness_pt": float(body.get("paper_thickness_pt", 0.28)),
        "watermark": body.get("watermark"),
        "margin_guard_mm": float(body.get("margin_guard_mm", 5.0)),
        "page_range": page_range,
        "page_size": page_size,
        "scaling": scaling,
    }

    try:
        from core.printing.printer_manager import PrinterManager
        res = PrinterManager.print_file(
            file_path=file_path,
            printer_name=printer_name,
            copies=copies,
            job_title=job_title,
            silent=True,
            optimizer_options=optimizer_options,
            tray_id=tray_id,
            duplex=duplex,
            page_range=page_range,
            page_size=page_size,
            scaling=scaling,
        )
        send_json(handler, res)
    except Exception as e:
        send_error(handler, f"Lỗi thực hiện lệnh in: {e}", status=500)

