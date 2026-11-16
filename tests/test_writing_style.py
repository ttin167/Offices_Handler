# -*- coding: utf-8 -*-
"""
tests/test_writing_style.py — Unit & Integration Tests for Core Writing Style Engine (Văn phong / Hành văn).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from docx import Document

from cli.ai_tools import build_parser
from core.copilot.copilot_provider import CopilotOrchestrator, LocalOllamaProvider, infer_action_cards
from core.docx.toolkit import polish_docx_document
from core.writing_style import (
    analyze_docx_writing_style,
    analyze_writing_style,
    apply_writing_style_to_docx,
    get_phrase_templates,
    list_writing_styles,
    refine_writing_style,
    resolve_style_id,
    rewrite_with_style,
)
import server.routes.docx_routes  # noqa: F401
from server.routes.base import Router


def test_style_registry_aliases_and_templates() -> None:
    styles = list_writing_styles()
    ids = {s["style_id"] for s in styles}
    assert ids == {"hanh_chinh", "hoc_thuat", "doanh_nghiep", "phap_ly", "ky_thuat"}

    assert resolve_style_id("nd30") == "hanh_chinh"
    assert resolve_style_id("academic") == "hoc_thuat"
    assert resolve_style_id("executive") == "doanh_nghiep"
    assert resolve_style_id("legal") == "phap_ly"
    assert resolve_style_id("srs") == "ky_thuat"

    tpl_all = get_phrase_templates("hanh_chinh")
    assert "cong_van" in tpl_all["available_doc_types"]
    tpl_cv = get_phrase_templates("hanh_chinh", doc_type="cong_van")
    assert "Căn cứ" in tpl_cv["template"]["basis"]


def test_analyze_and_refine_safe_vs_strict_with_protected_spans() -> None:
    raw = (
        'Trong thời gian vừa qua mình thấy kết quả rất là ok nhé. '
        'Cái này ko đc thiếu ở `ko_token` hay {{MaNV}} và "ko đc sửa trong ngoặc" tại https://example.com/ko. '
        'Ngoài ra k có chuyện 10k hay điểm k ở Washington DC, và và tự mình phải làm cái này cùng vs đơn vị.'
    )

    analysis = analyze_writing_style(raw, style_id="hanh_chinh")
    assert analysis.suitability_score < 80
    categories = {f.category for f in analysis.findings}
    assert {"shorthand", "colloquial", "wordiness", "pronoun", "redundancy"}.issubset(categories)

    # Safe mode keeps informal pronouns unless strict, while fixing shorthand/colloquial/wordiness/dups
    safe_res = refine_writing_style(raw, style_id="hanh_chinh", mode="safe", preserve_quotes=True)
    assert safe_res.changed is True
    assert "`ko_token`" in safe_res.refined_text
    assert "{{MaNV}}" in safe_res.refined_text
    assert '"ko đc sửa trong ngoặc"' in safe_res.refined_text
    assert "https://example.com/ko" in safe_res.refined_text
    assert "10k" in safe_res.refined_text
    assert "điểm k" in safe_res.refined_text
    assert "Washington DC" in safe_res.refined_text
    assert "tự mình" in safe_res.refined_text
    assert "và và" not in safe_res.refined_text
    assert "không được thiếu" in safe_res.refined_text
    assert "không có chuyện" in safe_res.refined_text
    assert "cùng với đơn vị" in safe_res.refined_text
    assert "đạt yêu cầu" in safe_res.refined_text
    assert "nhé" not in safe_res.refined_text

    # Strict mode also neutralizes informal standalone pronouns while keeping "tự mình"
    strict_res = refine_writing_style(raw, style_id="hoc_thuat", mode="strict", preserve_quotes=True)
    assert "nghiên cứu này" in strict_res.refined_text
    assert "tự mình" in strict_res.refined_text
    assert strict_res.analysis_after is not None
    assert strict_res.analysis_after["suitability_score"] > safe_res.analysis_before["suitability_score"]


def test_legal_and_technical_vagueness_and_long_sentence_detection() -> None:
    legal_text = (
        "Bên B phải giao hàng sớm nhất có thể và hợp lý tùy tình hình thực tế. "
        + " ".join(["điều khoản"] * 52)
        + "."
    )
    res = analyze_writing_style(legal_text, style_id="phap_ly")
    rule_ids = {f.rule_id for f in res.findings}
    assert "STYLE_VAGUE_SOONEST" in rule_ids
    assert "STYLE_VAGUE_REASONABLE" in rule_ids
    assert "STYLE_LONG_SENTENCE" in rule_ids


def test_docx_writing_style_preserves_runs_and_polish_hook(tmp_path: Path) -> None:
    src = tmp_path / "draft.docx"
    doc = Document()
    doc.add_heading("Đề xuất dự án ko sửa heading", level=1)

    p = doc.add_paragraph()
    r_bold = p.add_run("Cần kiểm tra ")
    r_bold.bold = True
    r_plain = p.add_run("cái này ko đc bỏ sót nhé.")

    tbl = doc.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = "Hạng mục"
    tbl.cell(0, 1).text = "Ghi chú"
    tbl.cell(1, 0).text = "Báo cáo"
    tbl.cell(1, 1).text = "Trong thời gian vừa qua đã xong rồi."
    doc.save(str(src))

    out_styled = tmp_path / "styled.docx"
    res = apply_writing_style_to_docx(str(src), output_path=str(out_styled), style_id="hanh_chinh", mode="safe")
    assert res["ok"] is True
    assert out_styled.is_file()
    assert res["edits_count"] >= 3

    styled_doc = Document(str(out_styled))
    assert styled_doc.paragraphs[0].text == "Đề xuất dự án ko sửa heading"
    body_p = styled_doc.paragraphs[1]
    assert body_p.runs[0].bold is True
    assert "không được" in body_p.text
    assert "nhé" not in body_p.text
    assert "trong thời gian qua" in styled_doc.tables[0].cell(1, 1).text.lower()

    diag = analyze_docx_writing_style(str(out_styled), style_id="hanh_chinh")
    assert diag["suitability_score"] >= 90

    # Hook inside polish_docx_document
    out_polished = tmp_path / "polished_styled.docx"
    pol_res = polish_docx_document(
        input_path=str(src),
        output_path=str(out_polished),
        preset="nd30",
        writing_style="doanh_nghiep",
        writing_style_mode="safe",
    )
    assert pol_res["ok"] is True
    assert "writing_style_report" in pol_res
    assert pol_res["writing_style_report"]["style_id"] == "doanh_nghiep"


def test_copilot_style_fallback_and_action_cards(monkeypatch: Any) -> None:
    monkeypatch.setenv("OLLAMA_BASE_URL", "http://127.0.0.1:59998/v1")
    res = rewrite_with_style(
        text="Mình thấy cái này ko đc ổn nhé.",
        style_id="hanh_chinh",
        mode="strict",
        use_ai=True,
        copilot_mode="local",
    )
    assert res.success is True
    assert res.fallback_used is True
    assert "không được" in res.refined_text
    assert "đơn vị" in res.refined_text.lower()

    orch = CopilotOrchestrator(
        default_mode="local",
        local_provider=LocalOllamaProvider(base_url="http://127.0.0.1:59998/v1", timeout=0.2),
    )
    orch_res = orch.rewrite_style("Cái này ok nhé.", style_id="doanh_nghiep", use_ai=True)
    assert orch_res["success"] is True
    assert "thống nhất" in orch_res["refined_text"]

    cards = infer_action_cards("Giúp tôi chuẩn hóa văn phong hành văn theo Nghị định 30")
    assert any(c.action_id == "act_writing_style" for c in cards)


def test_rest_routes_and_cli_docx_style(tmp_path: Path, capsys: Any) -> None:
    router = Router.get_instance()
    assert ("GET", "/api/docx/writing-styles") in router._exact_routes
    assert ("POST", "/api/docx/writing-style/analyze") in router._exact_routes
    assert ("POST", "/api/docx/writing-style/refine") in router._exact_routes

    class DummyHandler:
        def __init__(self, body_dict: Dict[str, Any] | None = None) -> None:
            self._body = body_dict or {}
            self.response_data: Any = None
            self.status_code: int = 200

        def _read_body_json(self) -> Dict[str, Any]:
            return self._body

        def _send_json(self, data: Any, status: int = 200) -> None:
            self.response_data = data
            self.status_code = status

    h_list = DummyHandler()
    assert router.dispatch(h_list, "GET", "/api/docx/writing-styles") is True
    assert h_list.status_code == 200
    assert len(h_list.response_data["styles"]) == 5

    h_refine = DummyHandler(
        {"text": "Nội dung này ko đc thiếu nhé.", "style_id": "hanh_chinh", "mode": "safe"}
    )
    assert router.dispatch(h_refine, "POST", "/api/docx/writing-style/refine") is True
    assert h_refine.status_code == 200
    assert "không được" in h_refine.response_data["refined_text"]

    # CLI test
    parser = build_parser()
    out_txt = tmp_path / "refined.txt"
    args = parser.parse_args(
        [
            "docx-style",
            "refine",
            "--text",
            "Báo cáo cái này ko đc chậm nhé.",
            "--style",
            "doanh_nghiep",
            "-o",
            str(out_txt),
            "--json",
        ]
    )
    rc = args.func(args)
    assert rc == 0
    captured = json.loads(capsys.readouterr().out)
    assert captured["success"] is True
    assert out_txt.is_file()
    assert "không được" in out_txt.read_text(encoding="utf-8")
