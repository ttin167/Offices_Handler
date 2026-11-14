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


