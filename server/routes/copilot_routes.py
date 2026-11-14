# -*- coding: utf-8 -*-
"""
server.routes.copilot_routes — REST API Endpoints for Universal RAG & Hybrid AI Copilot Workspace.
"""

from __future__ import annotations

import os
from typing import Any

from .base import router, send_json, send_error, read_json_body, resolve_body_file


def _is_cloud_ai_allowed() -> bool:
    """In Pure Local security mode, Cloud AI calls remain locked off unless explicitly opted in."""
    return os.environ.get("ANTIGRAVITY_ALLOW_CLOUD_AI", "0").strip().lower() in ("1", "true", "yes")


@router.get("/api/copilot/status")
def handle_copilot_status(handler: Any) -> None:
    """Returns Hybrid LLM connector status and Universal RAG index statistics."""
    try:
        from core.copilot.copilot_provider import CopilotOrchestrator
        from core.copilot.universal_rag import UniversalRAGEngine

        cloud_allowed = _is_cloud_ai_allowed()
        orch = CopilotOrchestrator()
        rag = UniversalRAGEngine()
        status_payload = orch.status()
        status_payload["local_only_enforced"] = not cloud_allowed
        if not cloud_allowed:
            status_payload["default_mode"] = "local"
            if isinstance(status_payload.get("cloud"), dict):
                status_payload["cloud"]["online"] = False
                status_payload["cloud"]["blocked_by_local_policy"] = True

        send_json(
            handler,
            {
                "success": True,
                "copilot": status_payload,
                "rag": rag.get_stats(),
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi kiểm tra trạng thái Copilot: {exc}", status=500)


@router.post("/api/copilot/index")
def handle_copilot_index(handler: Any) -> None:
    """Indexes a document or raw text into the Universal FTS5 RAG Knowledge Base."""
    body = read_json_body(handler)
    raw_path, resolved = resolve_body_file(
        body, ("file_path", "filepath", "path", "filename")
    )
    raw_text = body.get("content") or body.get("text")

    try:
        from core.copilot.universal_rag import UniversalRAGEngine

        rag = UniversalRAGEngine()
        if raw_path:
            if not resolved:
                send_error(handler, f"Không tìm thấy tệp để lập chỉ mục RAG: {raw_path}", status=404)
                return
            res = rag.index_file(resolved)
            send_json(handler, res)
            return

        if raw_text:
            doc_id = str(body.get("doc_id") or "DOC-INLINE-001")
            fname = str(body.get("title") or "Inline_Note.md")
            res = rag.index_text(doc_id=doc_id, filename=fname, content=str(raw_text))
            send_json(handler, res)
            return

        send_error(handler, "Cần truyền 'file_path' hoặc 'content' để lập chỉ mục RAG.", status=400)
    except Exception as exc:
        send_error(handler, f"Lỗi lập chỉ mục Universal RAG: {exc}", status=500)


@router.post("/api/copilot/search")
def handle_copilot_search(handler: Any) -> None:
    """Searches Universal RAG FTS5 index."""
    body = read_json_body(handler)
    query = str(body.get("query") or body.get("q") or "").strip()
    if not query:
        send_error(handler, "Tham số 'query' không được để trống.", status=400)
        return

    top_k = int(body.get("top_k") or body.get("limit") or 5)
    file_type = body.get("file_type")

    try:
        from core.copilot.universal_rag import UniversalRAGEngine

        rag = UniversalRAGEngine()
        chunks = rag.search(query=query, top_k=top_k, file_type=file_type)
        send_json(
            handler,
            {
                "success": True,
                "query": query,
                "total": len(chunks),
                "results": chunks,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi tìm kiếm Universal RAG: {exc}", status=500)


@router.post("/api/copilot/chat")
def handle_copilot_chat(handler: Any) -> None:
    """Executes RAG-augmented chat (enforces Local Ollama / Offline FTS5 RAG by default)."""
    body = read_json_body(handler)
    prompt = str(body.get("prompt") or body.get("message") or body.get("query") or "").strip()
    if not prompt:
        send_error(handler, "Tham số 'prompt' không được để trống.", status=400)
        return

    requested_mode = str(body.get("provider") or body.get("mode") or "local").lower()
    mode = requested_mode if _is_cloud_ai_allowed() else "local"
    top_k = int(body.get("top_k") or 5)
    _, active_resolved = resolve_body_file(body, ("active_file", "filepath", "file_path", "filename"))

    try:
        from core.copilot.copilot_provider import CopilotOrchestrator
        from core.copilot.universal_rag import UniversalRAGEngine

        rag = UniversalRAGEngine()
        if active_resolved:
            try:
                rag.index_file(active_resolved)
            except Exception:
                pass

        chunks = rag.search(query=prompt, top_k=top_k)
        orch = CopilotOrchestrator()
        resp = orch.chat(prompt=prompt, context_chunks=chunks, mode=mode)
        send_json(handler, resp.to_dict())
    except Exception as exc:
        send_error(handler, f"Lỗi xử lý AI Copilot Chat: {exc}", status=500)

