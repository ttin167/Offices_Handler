# -*- coding: utf-8 -*-
"""
Tests for Phase 4: Universal RAG (SQLite FTS5 Semantic Chunking) & Hybrid AI Copilot Provider.
"""

from __future__ import annotations

import os
import tempfile
from typing import Any, Dict

import fitz
import openpyxl
import pytest
from docx import Document

pytest.importorskip("core.copilot")

from core.copilot.copilot_provider import (
    CopilotOrchestrator,
    GoogleGeminiProvider,
    LocalOllamaProvider,
    infer_action_cards,
)
from core.copilot.universal_rag import UniversalRAGEngine
from server.routes.base import Router


def test_universal_rag_multi_format_indexing_and_search() -> None:
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmpdir:
        db_path = os.path.join(tmpdir, "rag_test.db")
        rag = UniversalRAGEngine(db_path=db_path)

        try:
            # 1. Index a DOCX document
            docx_path = os.path.join(tmpdir, "Quy_Che_Luong.docx")
            doc = Document()
            doc.add_heading("Quy Chế Lương Thưởng 3P", level=1)
            doc.add_paragraph("Mức phụ cấp ăn trưa tối đa là 730.000 VNĐ/tháng được miễn thuế TNCN.")
            doc.save(docx_path)
            r_docx = rag.index_file(docx_path)
            assert r_docx["success"] is True
            assert r_docx["chunks_indexed"] >= 1

            # 2. Index an XLSX workbook
            xlsx_path = os.path.join(tmpdir, "Bang_Cong.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "TongHop"
            ws.append(["MaNV", "HoTen", "PhuCapDocHai"])
            ws.append(["NV001", "Nguyen Van An", "Chế độ độc hại ca đêm 30%"])
            wb.save(xlsx_path)
            wb.close()
            r_xlsx = rag.index_file(xlsx_path)
            assert r_xlsx["success"] is True
            assert r_xlsx["chunks_indexed"] >= 1

            # 3. Index a PDF file
            pdf_path = os.path.join(tmpdir, "Huong_Dan_In.pdf")
            pdf = fitz.open()
            page = pdf.new_page()
            page.insert_text((72, 72), "Chế độ bình trang Booklet Creep Shingling bù trượt gáy sách.")
            pdf.save(pdf_path)
            pdf.close()
            r_pdf = rag.index_file(pdf_path)
            assert r_pdf["success"] is True
            assert r_pdf["chunks_indexed"] >= 1

            # 4. Search with unaccented Vietnamese (unicode61 remove_diacritics 2)
            hits_lunch = rag.search("phu cap an trua mien thue")
            assert len(hits_lunch) >= 1
            assert hits_lunch[0]["filename"] == "Quy_Che_Luong.docx"

            hits_booklet = rag.search("booklet creep shingling", file_type="pdf")
            assert len(hits_booklet) >= 1
            assert hits_booklet[0]["file_type"] == "pdf"

            stats = rag.get_stats()
            assert stats["total_documents"] == 3
            assert stats["total_chunks"] >= 3
        finally:
            rag.close()


def test_copilot_hybrid_orchestrator_and_action_cards() -> None:
    orch = CopilotOrchestrator(
        default_mode="hybrid",
        local_provider=LocalOllamaProvider(base_url="http://127.0.0.1:59999/v1", timeout=0.3),
        cloud_provider=GoogleGeminiProvider(api_key="", timeout=0.3),
    )

    chunks = [
        {
            "doc_id": "DOC-001",
            "filename": "Huong_Dan_Excel_Pivot.md",
            "section_title": "Mục 5: Tổng hợp Pivot & Chuẩn hóa bảng lương",
            "content": "Sử dụng Excel Power Suite để chuẩn hóa họ tên, SĐT/CCCD và xuất Pivot Chart tự động.",
        }
    ]

    resp = orch.chat(
        prompt="Hướng dẫn làm sạch dữ liệu bảng lương Excel, tạo Pivot và trộn thư Word",
        context_chunks=chunks,
        mode="local",
    )
    assert resp.success is True
    assert resp.fallback_used is True
    assert "Huong_Dan_Excel_Pivot.md" in resp.answer
    assert len(resp.action_cards) >= 1
    assert any(c.command_type == "switch_tab" for c in resp.action_cards)
    card_tabs = {c.payload.get("tab") for c in resp.action_cards}
    assert "ba" not in card_tabs
    assert "wireframe" not in card_tabs
    assert "law" not in card_tabs

    # Test cloud provider routing in hybrid mode with large context
    cloud_prov = GoogleGeminiProvider(api_key="dummy-key-test", timeout=0.2)
    orch_cloud = CopilotOrchestrator(default_mode="hybrid", cloud_provider=cloud_prov)
    selected = orch_cloud.get_provider(mode="hybrid", context_length=25000)
    assert selected.provider_name == "cloud_gemini"


def test_copilot_routes_registered_and_dispatch() -> None:
    router = Router.get_instance()
    assert ("GET", "/api/copilot/status") in router._exact_routes
    assert ("POST", "/api/copilot/index") in router._exact_routes
    assert ("POST", "/api/copilot/search") in router._exact_routes
    assert ("POST", "/api/copilot/chat") in router._exact_routes

    class DummyHandler:
        def __init__(self, body_dict: Dict[str, Any]) -> None:
            self._body = body_dict
            self.response_data: Any = None
            self.status_code: int = 200

        def _read_body_json(self) -> Dict[str, Any]:
            return self._body

        def _send_json(self, data: Any, status: int = 200) -> None:
            self.response_data = data
            self.status_code = status

    h = DummyHandler({"prompt": "Hướng dẫn trộn thư mail merge hàng loạt"})
    dispatched = router.dispatch(h, "POST", "/api/copilot/chat")
    assert dispatched is True
    assert h.status_code == 200
    assert h.response_data["success"] is True
    assert len(h.response_data["action_cards"]) >= 1

