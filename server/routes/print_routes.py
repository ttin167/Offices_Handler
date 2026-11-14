# -*- coding: utf-8 -*-
"""
server.routes.print_routes — Enterprise Hardware & Virtual Printing API Endpoints.
"""

from __future__ import annotations

from typing import Any

from .base import router, send_json, send_error, read_json_body, resolve_body_file


def _parse_num(body: dict[str, Any], key: str, default: int | float, cast: type) -> Any:
    raw = body.get(key)
    return cast(default if raw is None or raw == "" else raw)


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
    raw_path, file_path = resolve_body_file(
        body, ("file_path", "input_path", "file", "filepath", "filename")
    )
    if not raw_path:
        send_error(handler, "Thiếu tham số bắt buộc: file_path", status=400)
        return
    if not file_path:
        send_error(handler, f"Tệp không tồn tại hoặc nằm ngoài phạm vi an toàn: {raw_path}", status=404)
        return

    try:
        printer_name = body.get("printer_name") or body.get("printer")
        copies = _parse_num(body, "copies", 1, int)
        job_title = body.get("job_title")
        tray_id = body.get("tray_id")
        duplex = body.get("duplex")
        page_range = body.get("page_range")
        page_size = body.get("page_size")
        scaling = body.get("scaling", "fit")

        optimizer_options = {
            "grayscale": body.get("grayscale", False),
            "vector_grayscale": body.get("vector_grayscale", True),
            "dpi": _parse_num(body, "dpi", 300, int),
            "n_up": _parse_num(body, "n_up", 1, int),
            "booklet": body.get("booklet", False),
            "creep_compensation": body.get("creep_compensation", True),
            "paper_thickness_pt": _parse_num(body, "paper_thickness_pt", 0.28, float),
            "watermark": body.get("watermark"),
            "margin_guard_mm": _parse_num(body, "margin_guard_mm", 5.0, float),
            "page_range": page_range,
            "page_size": page_size,
            "scaling": scaling,
        }
    except (TypeError, ValueError) as exc:
        send_error(handler, f"Tham số in không hợp lệ: {exc}", status=400)
        return

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

