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
    resolve_body_file,
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
        writing_style = str(body.get("writing_style") or "").strip() or None
        writing_style_mode = str(body.get("writing_style_mode") or "safe").strip() or "safe"
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
            writing_style=writing_style,
            writing_style_mode=writing_style_mode,
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


@router.get("/api/docx/writing-styles")
def handle_list_writing_styles(handler: Any) -> None:
    """List available Vietnamese/enterprise prose registers and phrase templates."""
    try:
        from core.writing_style import PHRASE_TEMPLATES, list_writing_styles

        send_json(
            handler,
            {
                "success": True,
                "styles": list_writing_styles(),
                "phrase_templates": PHRASE_TEMPLATES,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi tải danh mục văn phong: {exc}", status=500)


@router.post("/api/docx/writing-style/analyze")
def handle_analyze_writing_style(handler: Any) -> None:
    """Analyze plain text or a workspace DOCX/MD/TXT document against a prose register."""
    try:
        from core.writing_style import analyze_docx_writing_style, analyze_writing_style

        body = read_json_body(handler)
        style_id = str(body.get("style_id") or body.get("writing_style") or "hanh_chinh").strip()
        raw_text = body.get("text")
        raw_path, resolved = resolve_body_file(body)

        if raw_path:
            if not resolved:
                send_error(handler, f"Không tìm thấy tệp văn bản hợp lệ: {raw_path}", status=404)
                return
            if resolved.lower().endswith(".docx"):
                report = analyze_docx_writing_style(resolved, style_id=style_id)
                send_json(handler, {"success": True, **report})
                return
            with open(resolved, "r", encoding="utf-8", errors="replace") as f:
                raw_text = f.read()

        if raw_text is None or not str(raw_text).strip():
            send_error(handler, "Cần cung cấp 'text' hoặc đường dẫn tệp (.docx/.md/.txt) để phân tích.", status=400)
            return

        result = analyze_writing_style(str(raw_text), style_id=style_id).to_dict()
        send_json(handler, {"success": True, **result})
    except ValueError as exc:
        send_error(handler, str(exc), status=400)
    except Exception as exc:
        send_error(handler, f"Lỗi phân tích văn phong: {exc}", status=500)


@router.post("/api/docx/writing-style/refine")
def handle_refine_writing_style(handler: Any) -> None:
    """Refine prose style on plain text or a DOCX file (run-preserving OpenXML)."""
    try:
        from core.writing_style import apply_writing_style_to_docx, resolve_style_id, rewrite_with_style
        from .copilot_routes import _is_cloud_ai_allowed

        body = read_json_body(handler)
        style_id = resolve_style_id(str(body.get("style_id") or body.get("writing_style") or "hanh_chinh"))
        mode = str(body.get("mode") or "safe").strip().lower()
        preserve_quotes = bool(body.get("preserve_quotes", True))
        raw_path, resolved = resolve_body_file(body)

        if raw_path:
            if not resolved:
                send_error(handler, f"Không tìm thấy tệp văn bản hợp lệ: {raw_path}", status=404)
                return
            if resolved.lower().endswith(".docx"):
                out_name, out_path = build_output_filepath(resolved, f"styled_{style_id}", ext=".docx")
                res = apply_writing_style_to_docx(
                    input_path=resolved,
                    output_path=out_path,
                    style_id=style_id,
                    mode=mode,
                    preserve_quotes=preserve_quotes,
                    include_headings=bool(body.get("include_headings", False)),
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
                return
            with open(resolved, "r", encoding="utf-8", errors="replace") as f:
                raw_text = f.read()
        else:
            raw_text = body.get("text")

        if raw_text is None or not str(raw_text).strip():
            send_error(handler, "Cần cung cấp 'text' hoặc đường dẫn tệp (.docx/.md/.txt) để chuẩn hóa.", status=400)
            return

        use_ai = bool(body.get("use_ai", False))
        requested_provider = str(body.get("provider") or body.get("copilot_mode") or "local").lower()
        copilot_mode = requested_provider if _is_cloud_ai_allowed() else "local"

        res = rewrite_with_style(
            text=str(raw_text),
            style_id=style_id,
            mode=mode,
            use_ai=use_ai,
            copilot_mode=copilot_mode,
            document_type=body.get("doc_type") or body.get("document_type"),
            instructions=body.get("instructions"),
            preserve_quotes=preserve_quotes,
        )
        send_json(handler, res.to_dict())
    except ValueError as exc:
        send_error(handler, str(exc), status=400)
    except Exception as exc:
        send_error(handler, f"Lỗi chuẩn hóa văn phong: {exc}", status=500)

