# -*- coding: utf-8 -*-
"""
server.routes.doc_routes — Document AST Inspection & Mail Merge Dataset Handlers.
"""

from __future__ import annotations

import os
from typing import Any

from .base import router, send_json, send_error, read_json_body, resolve_body_file


@router.post("/api/docx/read")
def handle_docx_read(handler: Any) -> None:
    """Reads a DOCX file and returns its structured JSON AST snapshot."""
    body = read_json_body(handler)
    raw_path, resolved = resolve_body_file(
        body, ("file_path", "filepath", "source_path", "path", "filename")
    )
    if not resolved:
        send_error(handler, f"Tệp DOCX không tồn tại hoặc không hợp lệ: {raw_path}", status=404)
        return

    try:
        from core.docx.reader import read_docx
        snapshot = read_docx(resolved)
        send_json(handler, {
            "success": True,
            "filename": os.path.basename(resolved),
            "source_file": resolved,
            "snapshot": snapshot,
        })
    except Exception as e:
        send_error(handler, f"Lỗi đọc tài liệu DOCX: {str(e)}", status=500)


@router.post("/api/mail-merge/inspect-data")
def handle_mail_merge_inspect_data(handler: Any) -> None:
    """Inspects a dataset file (XLSX, CSV, JSON) and returns column headers and sample records."""
    body = read_json_body(handler)
    raw_path, resolved = resolve_body_file(
        body, ("data", "data_source", "file_path", "filepath", "filename")
    )
    if not resolved:
        send_error(handler, f"Tệp dữ liệu không tồn tại: {raw_path}", status=404)
        return

    sheet_name = body.get("sheet_name") or body.get("sheet")
    if sheet_name and str(sheet_name).strip().lower() in ("none", "null", ""):
        sheet_name = None

    try:
        from core.docx.mail_merge import inspect_data_columns
        result = inspect_data_columns(resolved, sheet_name=sheet_name)
        result["filename"] = os.path.basename(resolved)
        send_json(handler, result)
    except Exception as e:
        send_error(handler, f"Lỗi phân tích tệp dữ liệu: {str(e)}", status=500)


@router.post("/api/mail-merge/inspect-dataset")
def handle_mail_merge_inspect_dataset_alias(handler: Any) -> None:
    """Alias for /api/mail-merge/inspect-data."""
    handle_mail_merge_inspect_data(handler)

