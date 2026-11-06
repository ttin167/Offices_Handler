# -*- coding: utf-8 -*-
"""
server.routes.doc_routes — Document AST Inspection & Mail Merge Dataset Handlers.
"""

from __future__ import annotations

import os
import tempfile
from typing import Any, Optional

from .base import router, send_json, send_error, read_json_body


def _safe_resolve(path_str: Optional[str]) -> Optional[str]:
    """Safely resolves an input file path against allowed sandbox roots."""
    if not path_str or not isinstance(path_str, str):
        return None
    cleaned = path_str.strip()
    if not cleaned:
        return None

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    allowed_roots = [
        base_dir,
        os.path.join(base_dir, "uploads"),
        os.path.join(base_dir, "output"),
        tempfile.gettempdir(),
    ]

    candidates = [
        cleaned,
        os.path.join(base_dir, cleaned),
        os.path.join(base_dir, "uploads", cleaned),
        os.path.join(base_dir, "uploads", os.path.basename(cleaned)),
        os.path.join(base_dir, "output", os.path.basename(cleaned)),
    ]

    for cand in candidates:
        abs_c = os.path.abspath(cand)
        if os.path.isfile(abs_c):
            norm_cand = os.path.normpath(abs_c)
            for r in allowed_roots:
                norm_r = os.path.normpath(r)
                try:
                    if os.path.commonpath([norm_r, norm_cand]) == norm_r:
                        return abs_c
                except (ValueError, OSError):
                    pass
    return None


@router.post("/api/docx/read")
def handle_docx_read(handler: Any) -> None:
    """Reads a DOCX file and returns its structured JSON AST snapshot."""
    body = read_json_body(handler)
    raw_path = (
        body.get("file_path")
        or body.get("filepath")
        or body.get("source_path")
        or body.get("path")
        or body.get("filename")
    )
    resolved = _safe_resolve(raw_path)
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
    raw_path = (
        body.get("data")
        or body.get("data_source")
        or body.get("file_path")
        or body.get("filepath")
        or body.get("filename")
    )
    resolved = _safe_resolve(raw_path)
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

