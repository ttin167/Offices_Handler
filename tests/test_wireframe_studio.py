"""
tests/test_wireframe_studio.py — Automated Test Suite for Rapid Low-Fi UI Wireframe Studio (Option A).

Covers:
1. Pure Vector SVG rendering (-tsvg) with UTF-8 Vietnamese text & transparent monochrome skinparams.
2. Catalog of 6 Enterprise BA Wireframe Templates (login_auth, crud_table, wizard_form, modal_confirm, mobile_screen, dashboard_kpi).
3. One-Way JSON-to-Salt generator & direct Salt code rendering.
4. 1-Click Word (.docx) OpenXML injection enforcing Printable Margin Guard (ERR_DOCX_005 <= 14.0 cm),
   Last Paragraph Rule (ERR_DOCX_001), and temp file cleanup (ERR_PIPE_001).
5. REST API endpoints (/api/diagram/wireframe/templates, /render, /inject-doc).
6. MCP Server tool (ba_wireframe_render) & CLI (wireframe-render).
"""

from __future__ import annotations

import glob
import http.server
import json
import os
import subprocess
import sys
import tempfile
import threading
import urllib.request
from docx import Document

from core.diagrams.salt_wireframe import SaltWireframeEngine, WIREFRAME_TEMPLATES
from server.app import OfficeApiHandler
from server.mcp_server import handle_rpc_request


def test_render_salt_to_svg_vietnamese_and_skinparams():
    """1. Verify pure vector SVG export preserves Vietnamese diacritics and injects clean skinparams."""
    spec = {
        "title": "Cổng Xác Thực Giao Dịch Ngân Hàng Số",
        "elements": [
            {"type": "input", "label": "Tài khoản nhận:", "value": "190088889999"},
            {"type": "checkbox", "label": "Xác nhận sinh trắc học khuôn mặt", "checked": True},
            {"type": "button_group", "buttons": ["Xác Nhận Chuyển Tiền", "Hủy Bỏ"]},
        ],
    }
    with tempfile.TemporaryDirectory() as tmp:
        out_svg = os.path.join(tmp, "auth_wireframe.svg")
        res = SaltWireframeEngine.render_salt_to_svg(spec, output_path=out_svg)

        assert res["success"] is True
        assert os.path.isfile(out_svg)
        assert "<svg" in res["svg"]
        assert "Cổng Xác Thực Giao Dịch Ngân Hàng Số" in res["svg"]
        assert "Xác Nhận Chuyển Tiền" in res["svg"]
        assert "skinparam BackgroundColor transparent" in res["salt_code"]
        assert "skinparam monochrome true" in res["salt_code"]
        assert res["width"] > 100
        assert res["height"] > 40


def test_wireframe_templates_catalog_completeness():
    """2. Verify all 6 enterprise BA wireframe templates exist and compile cleanly to Salt & SVG."""
    expected_ids = {
        "login_auth",
        "crud_table",
        "wizard_form",
        "modal_confirm",
        "mobile_screen",
        "dashboard_kpi",
    }
    catalog = SaltWireframeEngine.get_templates()
    assert set(catalog.keys()) == expected_ids

    for tid, tpl in catalog.items():
        assert tpl["id"] == tid
        assert tpl["title"]
        assert tpl["category"]
        assert isinstance(tpl["json_spec"], dict)
        assert "@startsalt" in tpl["salt_code"]

        # Verify JSON spec one-way generator produces valid Salt code
        generated_salt = SaltWireframeEngine.spec_to_salt_code(tpl["json_spec"])
        assert "@startsalt" in generated_salt
        assert "skinparam BackgroundColor transparent" in generated_salt

    # Render one template by ID directly to SVG
    res_svg = SaltWireframeEngine.render_salt_to_svg("crud_table")
    assert res_svg["success"] is True
    assert "REQ-001" in res_svg["svg"]


def test_inject_wireframe_into_docx_and_cell_rule():
    """3. Verify 1-Click Word DOCX injection embeds picture, adds caption, and cleans up temp files."""
    with tempfile.TemporaryDirectory() as tmp:
        doc_path = os.path.join(tmp, "sample_brd.docx")
        doc = Document()
        doc.add_heading("1. Đặc Tả Giao Diện Người Dùng", level=1)
        tbl = doc.add_table(rows=1, cols=1)
        tbl.cell(0, 0).text = "Bảng thông tin yêu cầu"
        doc.save(doc_path)

        res = SaltWireframeEngine.inject_wireframe_into_docx(
            docx_path=doc_path,
            spec_or_code="login_auth",
            heading="1. Đặc Tả Giao Diện Người Dùng",
            caption="Hình 1.1: Màn hình đăng nhập SSO & 2FA",
        )

        assert res["success"] is True
        assert os.path.isfile(res["docx_path"])
        assert res["width_cm"] <= 14.0
        assert res["height_cm"] > 1.0

        # Verify picture and caption exist inside the saved Word document
        reloaded = Document(doc_path)
        assert len(reloaded.inline_shapes) == 1
        all_text = "\n".join(p.text for p in reloaded.paragraphs)
        assert "Hình 1.1: Màn hình đăng nhập SSO & 2FA" in all_text

        # Verify no leftover temp wireframe PNGs in temp dir (ERR_PIPE_001)
        leftover = glob.glob(os.path.join(tempfile.gettempdir(), "step_wireframe_temp_*.png"))
        assert len(leftover) == 0


def test_margin_guard_auto_scale_err_docx_005():
    """4. Verify Printable Margin Guard (ERR_DOCX_005) caps image width <= 14.0 cm even if requested larger."""
    with tempfile.TemporaryDirectory() as tmp:
        out_docx = os.path.join(tmp, "margin_guarded.docx")
        res = SaltWireframeEngine.inject_wireframe_into_docx(
            docx_path=out_docx,
            spec_or_code="dashboard_kpi",
            heading="Dashboard Quản Trị",
            caption="Hình 2.1: Dashboard KPI",
            max_width_cm=28.0,  # Intentionally oversized request
        )
        assert res["success"] is True
        assert res["width_cm"] <= 14.0
        assert res["width_cm"] <= res["usable_margin_cm"]

        doc = Document(out_docx)
        shape = doc.inline_shapes[0]
        width_cm_actual = shape.width.cm
        assert width_cm_actual <= 14.05


def test_rest_api_wireframe_endpoints():
    """5 & 6 & 7. Verify REST API endpoints (/templates, /render, /inject-doc)."""
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), OfficeApiHandler)
    port = server.server_port
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    try:
        base_url = f"http://127.0.0.1:{port}"

        # A. GET /api/diagram/wireframe/templates
        with urllib.request.urlopen(f"{base_url}/api/diagram/wireframe/templates", timeout=10) as resp:
            assert resp.status == 200
            tpl_data = json.loads(resp.read().decode("utf-8"))
            assert tpl_data["success"] is True
            assert "login_auth" in tpl_data["templates"]
            assert "dashboard_kpi" in tpl_data["templates"]

        # B. POST /api/diagram/wireframe/render (format="both")
        render_payload = json.dumps({
            "template": "wizard_form",
            "format": "both",
        }).encode("utf-8")
        req_render = urllib.request.Request(
            f"{base_url}/api/diagram/wireframe/render",
            data=render_payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req_render, timeout=20) as resp:
            assert resp.status == 200
            r_data = json.loads(resp.read().decode("utf-8"))
            assert r_data["success"] is True
            assert "<svg" in r_data["svg"]
            assert r_data["svg_url"].endswith(".svg")
            assert r_data["png_url"].endswith(".png")

        # C. POST /api/diagram/wireframe/inject-doc
        inject_payload = json.dumps({
            "template": "mobile_screen",
            "heading": "Giao Diện Mobile App",
            "caption": "Hình 3: Màn hình Ví điện tử trên Mobile",
        }).encode("utf-8")
        req_inject = urllib.request.Request(
            f"{base_url}/api/diagram/wireframe/inject-doc",
            data=inject_payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req_inject, timeout=20) as resp:
            assert resp.status == 200
            i_data = json.loads(resp.read().decode("utf-8"))
            assert i_data["success"] is True
            assert i_data["filename"].endswith(".docx")
            assert i_data["width_cm"] <= 14.0
    finally:
        server.shutdown()
        server.server_close()


def test_mcp_ba_wireframe_render_tool():
    """8. Verify MCP tool ba_wireframe_render supports template and SVG format."""
    with tempfile.TemporaryDirectory() as tmp:
        out_svg = os.path.join(tmp, "mcp_wireframe.svg")
        rpc_req = {
            "jsonrpc": "2.0",
            "id": 99,
            "method": "tools/call",
            "params": {
                "name": "ba_wireframe_render",
                "arguments": {
                    "template": "modal_confirm",
                    "format": "svg",
                    "output_png_path": out_svg,
                },
            },
        }
        rpc_resp = handle_rpc_request(rpc_req)
        assert rpc_resp is not None
        assert rpc_resp["result"]["isError"] is False
        payload = json.loads(rpc_resp["result"]["content"][0]["text"])
        assert payload["success"] is True
        assert "<svg" in payload["svg"]
        assert os.path.isfile(out_svg)


def test_cli_wireframe_render_svg_and_inject():
    """9. Verify CLI subcommand wireframe-render with --template, --format svg, and --inject-doc."""
    with tempfile.TemporaryDirectory() as tmp:
        out_svg = os.path.join(tmp, "cli_out.svg")
        cmd_svg = [
            sys.executable,
            "-X",
            "utf8",
            "ai_tools_cli.py",
            "wireframe-render",
            "--template",
            "login_auth",
            "--format",
            "svg",
            "-o",
            out_svg,
        ]
        proc = subprocess.run(cmd_svg, capture_output=True, text=True, encoding="utf-8", timeout=30)
        assert proc.returncode == 0
        assert os.path.isfile(out_svg)

        out_docx = os.path.join(tmp, "cli_injected.docx")
        cmd_doc = [
            sys.executable,
            "-X",
            "utf8",
            "ai_tools_cli.py",
            "wireframe-render",
            "--template",
            "crud_table",
            "--inject-doc",
            out_docx,
            "--heading",
            "Màn Hình Quản Trị Danh Mục",
            "--caption",
            "Hình 4: Bảng CRUD",
        ]
        proc_doc = subprocess.run(cmd_doc, capture_output=True, text=True, encoding="utf-8", timeout=30)
        assert proc_doc.returncode == 0
        assert os.path.isfile(out_docx)

