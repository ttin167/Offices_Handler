# -*- coding: utf-8 -*-
"""
tests/test_modular_server_routes.py — Test Suite for Modular Router & Routes (v5.0).
"""

import json
from unittest.mock import MagicMock
import pytest

from server.routes import Router, router, send_json, send_error, read_json_body


def test_router_exact_match():
    r = Router()
    called = []

    def mock_handler(req):
        called.append("exact_match")

    r.register("GET", "/api/v5/test", mock_handler)

    assert r.dispatch(None, "GET", "/api/v5/test") is True
    assert called == ["exact_match"]

    # Wrong method or path
    assert r.dispatch(None, "POST", "/api/v5/test") is False
    assert r.dispatch(None, "GET", "/api/v5/other") is False


def test_router_parametric_match():
    r = Router()
    captured_params = {}

    def param_handler(req, item_id, action):
        captured_params["id"] = item_id
        captured_params["action"] = action

    r.register("POST", "/api/items/<item_id>/<action>", param_handler)

    assert r.dispatch(None, "POST", "/api/items/1234/execute") is True
    assert captured_params == {"id": "1234", "action": "execute"}


def test_system_routes_dispatch():
    r = Router.get_instance()

    mock_req = MagicMock()
    mock_req._send_json = MagicMock()

    # /api/health
    dispatched = r.dispatch(mock_req, "GET", "/api/health")
    assert dispatched is True
    mock_req._send_json.assert_called_once()
    args, kwargs = mock_req._send_json.call_args
    assert args[0]["status"] == "healthy"
    assert args[0]["version"] == "5.0.0"


def test_print_routes_dispatch():
    r = Router.get_instance()

    mock_req = MagicMock()
    mock_req._send_json = MagicMock()

    # /api/print/printers
    dispatched = r.dispatch(mock_req, "GET", "/api/print/printers")
    assert dispatched is True
    mock_req._send_json.assert_called_once()
    args, kwargs = mock_req._send_json.call_args
    assert args[0]["success"] is True
    assert "printers" in args[0]


def test_pure_office_xlsx_and_docx_routes_registered_and_validate() -> None:
    r = Router.get_instance()

    expected_post_routes = [
        "/api/xlsx/power-clean",
        "/api/xlsx/transform-columns",
        "/api/xlsx/pivot-chart",
        "/api/xlsx/xlookup-join",
        "/api/xlsx/split-by-column",
        "/api/xlsx/diff",
        "/api/xlsx/eval",
        "/api/xlsx/tracker/init",
        "/api/xlsx/tracker/add",
        "/api/xlsx/tracker/sync-daily",
        "/api/docx/polish",
        "/api/docx/merge",
        "/api/docx/split",
        "/api/docx/watermark-hf",
    ]
    assert ("GET", "/api/xlsx/tracker/inspect") in r._exact_routes
    for route in expected_post_routes:
        assert ("POST", route) in r._exact_routes

    # Verify forbidden BA/Law/Wireframe prefixes are completely gone
    for method, route in r._exact_routes:
        assert not route.startswith("/api/ba/")
        assert not route.startswith("/api/law/")
        assert not route.startswith("/api/wireframe/")

    # Invalid/missing file inputs should fail cleanly with status 400 JSON error
    for route in ("/api/xlsx/power-clean", "/api/docx/polish"):
        mock_req = MagicMock()
        mock_req._read_body_json = MagicMock(return_value={})
        mock_req._send_json = MagicMock()
        assert r.dispatch(mock_req, "POST", route) is True
        mock_req._send_json.assert_called_once()
        args, kwargs = mock_req._send_json.call_args
        assert kwargs.get("status", args[1] if len(args) > 1 else None) == 400 or args[0].get("error")


def test_resolve_workspace_file_web_dirs_and_traversal_guard() -> None:
    import os
    import uuid
    from server.routes.base import get_output_dir, get_upload_dir, resolve_workspace_file

    uid = uuid.uuid4().hex[:8]
    up_name = f"_test_up_{uid}.txt"
    out_name = f"_test_out_{uid}.txt"
    up_path = os.path.join(get_upload_dir(), up_name)
    out_path = os.path.join(get_output_dir(), out_name)

    try:
        with open(up_path, "w", encoding="utf-8") as f:
            f.write("upload")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("output")

        assert resolve_workspace_file(up_name) == os.path.abspath(up_path)
        assert resolve_workspace_file(out_name) == os.path.abspath(out_path)
        assert resolve_workspace_file(f"../.web_uploads/{up_name}") is None
        outside_same_base = os.path.join(os.path.abspath(os.sep), "outside_sandbox_dir", up_name)
        assert resolve_workspace_file(outside_same_base) is None
    finally:
        for p in (up_path, out_path):
            if os.path.exists(p):
                os.remove(p)


def test_doc_copilot_and_print_routes_resolve_web_outputs() -> None:
    import os
    import uuid
    from unittest.mock import patch
    from docx import Document
    from server.routes.base import Router, get_output_dir

    r = Router.get_instance()
    uid = uuid.uuid4().hex[:8]
    out_dir = get_output_dir()
    docx_name = f"_test_preview_{uid}.docx"
    csv_name = f"_test_merge_{uid}.csv"
    pdf_name = f"_test_print_{uid}.pdf"
    docx_path = os.path.join(out_dir, docx_name)
    csv_path = os.path.join(out_dir, csv_name)
    pdf_path = os.path.join(out_dir, pdf_name)

    def _dispatch(route: str, body: dict):
        req = MagicMock()
        req._read_body_json = MagicMock(return_value=body)
        req._send_json = MagicMock()
        assert r.dispatch(req, "POST", route) is True
        args, kwargs = req._send_json.call_args
        return args[0], kwargs.get("status", args[1] if len(args) > 1 else 200)

    try:
        d = Document()
        d.add_paragraph("Preview snapshot test")
        d.save(docx_path)

        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("HoTen,PhongBan\nNguyen An,IT\n")

        with open(pdf_path, "wb") as f:
            f.write(b"%PDF-1.4\n%%EOF\n")

        res_docx, status_docx = _dispatch("/api/docx/read", {"filename": docx_name})
        assert status_docx == 200
        assert res_docx["success"] is True
        assert res_docx["source_file"] == os.path.abspath(docx_path)

        res_csv, status_csv = _dispatch("/api/mail-merge/inspect-data", {"filename": csv_name})
        assert status_csv == 200
        assert res_csv["success"] is True
        assert "HoTen" in res_csv["columns"]

        with patch("core.copilot.universal_rag.UniversalRAGEngine.index_file", return_value={"success": True}) as mock_idx:
            res_idx, status_idx = _dispatch("/api/copilot/index", {"filename": csv_name})
            assert status_idx == 200
            assert res_idx["success"] is True
            mock_idx.assert_called_once_with(os.path.abspath(csv_path))

        with patch("core.printing.printer_manager.PrinterManager.print_file", return_value={"success": True}) as mock_print:
            res_pr, status_pr = _dispatch("/api/print/job", {"filename": pdf_name, "copies": None})
            assert status_pr == 200
            assert res_pr["success"] is True
            assert mock_print.call_args.kwargs["file_path"] == os.path.abspath(pdf_path)
            assert mock_print.call_args.kwargs["copies"] == 1

            res_bad_num, status_bad_num = _dispatch("/api/print/job", {"filename": pdf_name, "copies": "abc"})
            assert status_bad_num == 400
            assert res_bad_num["success"] is False

            res_trav, status_trav = _dispatch("/api/print/job", {"file_path": f"../.web_outputs/{pdf_name}"})
            assert status_trav == 404
            assert res_trav["success"] is False
            mock_print.assert_called_once()
    finally:
        for p in (docx_path, csv_path, pdf_path):
            if os.path.exists(p):
                os.remove(p)


def test_app_delete_dispatch_and_modular_error_propagation() -> None:
    import io
    from server.app import OfficeApiHandler, MAX_UPLOAD_SIZE_BYTES
    from server.routes.base import Router

    class _LocalMockHandler(OfficeApiHandler):
        def __init__(self, body_bytes: bytes, headers: dict, path: str, method: str) -> None:
            self.rfile = io.BytesIO(body_bytes)
            self.wfile = io.BytesIO()
            self.path = path
            self.command = method
            self.requestline = f"{method} {path} HTTP/1.1"
            self.headers = headers
            self.client_address = ("127.0.0.1", 55000)
            self._headers_buffer = []
            self.status_code = 0

        def send_response(self, code, message=None):
            self.status_code = code

        def send_header(self, keyword, value):
            self._headers_buffer.append((keyword, value))

        def end_headers(self):
            pass

    # 1. DELETE /api/health must dispatch DELETE (not GET) and return 404
    h_del = _LocalMockHandler(b"", {}, "/api/health", "DELETE")
    h_del.do_DELETE()
    assert h_del.status_code == 404

    # 2. Malformed JSON / oversized payload in modular POST route must return 400 / 413 (not 404 fallback)
    h_bad_json = _LocalMockHandler(b"{bad", {"Content-Length": "4"}, "/api/docx/read", "POST")
    h_bad_json.do_POST()
    assert h_bad_json.status_code == 400

    h_big = _LocalMockHandler(b"", {"Content-Length": str(MAX_UPLOAD_SIZE_BYTES + 10)}, "/api/docx/read", "POST")
    h_big.do_POST()
    assert h_big.status_code == 413

    # 3. Unhandled exception inside modular route must propagate to 500 instead of silent fallback
    router_inst = Router.get_instance()
    test_route = "/api/_unit_test_boom"

    def _boom(_handler):
        raise RuntimeError("modular route boom")

    router_inst.register("POST", test_route, _boom)
    try:
        h_boom = _LocalMockHandler(b"{}", {"Content-Length": "2"}, test_route, "POST")
        h_boom.do_POST()
        assert h_boom.status_code == 500
        payload = json.loads(h_boom.wfile.getvalue().decode("utf-8"))
        assert payload["success"] is False
        assert "modular route boom" in payload["error"]
    finally:
        router_inst._exact_routes.pop(("POST", test_route), None)


