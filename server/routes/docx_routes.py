# -*- coding: utf-8 -*-
"""
server.routes.docx_routes — Pure Office Word Mastery REST API Endpoints.
"""

from __future__ import annotations

from typing import Any, List

from .base import (
    build_output_filepath,
    download_url_for,
    read_json_body,
    resolve_workspace_file,
    router,
    send_error,
    send_json,
)


def _resolve_doc_input(body: dict) -> str | None:
    for key in ("file_path", "filepath", "filename", "input_path"):
        raw = body.get(key)
        if raw:
            resolved = resolve_workspace_file(str(raw))
            if resolved:
                return resolved
    return None


@router.post("/api/docx/polish")
def handle_docx_polish(handler: Any) -> None:
    """Apply Decree 30/2020/NĐ-CP or Executive Word formatting, cleanup, and Page X/Y."""
    try:
        from core.docx.toolkit import polish_docx_document

        body = read_json_body(handler)
        src = _resolve_doc_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Word (.docx) hợp lệ.", status=400)
            return

        preset = str(body.get("preset", "nd30")).strip() or "nd30"
        out_name, out_path = build_output_filepath(src, f"polished_{preset}", ext=".docx")
        res = polish_docx_document(
            input_path=src,
            output_path=out_path,
            preset=preset,
            add_page_numbers=bool(body.get("add_page_numbers", True)),
            sanitize_typography=bool(body.get("sanitize_typography", True)),
            fix_tables=bool(body.get("fix_tables", True)),
            normalize_captions=bool(body.get("normalize_captions", False)),
            admin_header=body.get("admin_header") if isinstance(body.get("admin_header"), dict) else None,
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi chuẩn hóa văn bản Word: {exc}", status=500)


@router.post("/api/docx/merge")
def handle_docx_merge(handler: Any) -> None:
    """Merge multiple Word .docx files sequentially while preserving media/tables."""
    try:
        from core.docx.toolkit import merge_docx_files

        body = read_json_body(handler)
        raw_inputs = body.get("inputs") or body.get("input_paths") or []
        if not isinstance(raw_inputs, list) or len(raw_inputs) < 2:
            send_error(handler, "Cần chọn ít nhất 2 tệp Word (.docx) để ghép nối.", status=400)
            return

        resolved_paths: List[str] = []
        for item in raw_inputs:
            p = resolve_workspace_file(str(item))
            if p:
                resolved_paths.append(p)
        if len(resolved_paths) < 2:
            send_error(handler, "Không tìm thấy đủ tệp Word hợp lệ trên hệ thống.", status=400)
            return

        out_name, out_path = build_output_filepath(resolved_paths[0], "merged", ext=".docx")
        res = merge_docx_files(
            input_paths=resolved_paths,
            output_path=out_path,
            add_page_breaks=bool(body.get("add_page_breaks", True)),
            normalize_page_numbers=bool(body.get("normalize_page_numbers", True)),
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi gộp nhiều file Word: {exc}", status=500)


@router.post("/api/docx/split")
def handle_docx_split(handler: Any) -> None:
    """Split a master Word document by Heading level into chapter .docx files (.zip)."""
    try:
        from core.docx.toolkit import split_docx_by_heading

        body = read_json_body(handler)
        src = _resolve_doc_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Word (.docx) cần tách.", status=400)
            return

        heading_level = int(body.get("heading_level", 1))
        out_name, out_path = build_output_filepath(src, f"split_h{heading_level}", ext=".zip")
        res = split_docx_by_heading(
            input_path=src,
            output_path=out_path,
            heading_level=heading_level,
            as_zip=True,
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi tách file Word theo Heading: {exc}", status=500)


@router.post("/api/docx/watermark-hf")
def handle_docx_watermark_hf(handler: Any) -> None:
    """Apply diagonal watermark and custom header/footer + Page X/Y to Word document."""
    try:
        from core.docx.toolkit import apply_docx_watermark_and_hf

        body = read_json_body(handler)
        src = _resolve_doc_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Word (.docx) hợp lệ.", status=400)
            return

        out_name, out_path = build_output_filepath(src, "branded", ext=".docx")
        res = apply_docx_watermark_and_hf(
            input_path=src,
            output_path=out_path,
            watermark_text=body.get("watermark_text") or None,
            header_text=body.get("header_text") or None,
            footer_text=body.get("footer_text") or None,
            add_page_numbers=bool(body.get("add_page_numbers", True)),
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi đóng watermark & Header/Footer Word: {exc}", status=500)

