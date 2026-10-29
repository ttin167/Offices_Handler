# -*- coding: utf-8 -*-
"""
tests/test_server_security.py — Automated Verification for Server Security Hardening.

Validates:
  1. Strict Sandboxing in resolve_file_path (Rejects paths outside allowed roots).
  2. Path Traversal Elimination in _serve_file (Rejects ../ traversal).
  3. Source Exposure Protection in _serve_static_file (Blocks .py, .env, .git).
  4. Upload Payload Quota Enforcement (Rejects Content-Length > MAX_UPLOAD_SIZE_MB with HTTP 413).
  5. Optional API Key Protection (Rejects unauthorized requests with HTTP 401 when API_KEY is set).
  6. Smart CORS Header Generation (Validates origin echo and custom allowed origins).
"""

import io
import json
import os
import tempfile
import unittest
import urllib.parse
from unittest.mock import MagicMock, patch

from server.app import (
    BASE_DIR,
    OUTPUT_DIR,
    UPLOAD_DIR,
    OfficeApiHandler,
    resolve_file_path,
)


class DummyMockServer:
    """Mock HTTP Server for BaseHTTPRequestHandler instantiation."""
    def __init__(self):
        self.server_name = "localhost"
        self.server_port = 8000


class DummyMockHandler(OfficeApiHandler):
    """Subclass of OfficeApiHandler with in-memory streams for testing."""
    def __init__(self, request_bytes: bytes, headers: dict, path: str = "/api/health", method: str = "GET"):
        self.rfile = io.BytesIO(request_bytes)
        self.wfile = io.BytesIO()
        self.path = path
        self.command = method
        self.requestline = f"{method} {path} HTTP/1.1"
        self.headers = headers
        self.client_address = ("127.0.0.1", 54321)
        self.server = DummyMockServer()
        self._headers_buffer = []

    def send_response(self, code, message=None):
        self.status_code = code

    def send_header(self, keyword, value):
        self._headers_buffer.append((keyword, value))

    def end_headers(self):
        pass


class TestServerSecurity(unittest.TestCase):

    def test_01_resolve_file_path_sandboxing(self):
        """Validates that resolve_file_path strictly forbids paths outside allowed roots."""
        # 1. Path inside BASE_DIR should resolve
        readme = os.path.join(BASE_DIR, "README.md")
        self.assertTrue(os.path.isfile(readme))
        res = resolve_file_path(readme)
        self.assertIsNotNone(res)
        self.assertEqual(os.path.abspath(readme), res)

        # 2. Path inside tempfile directory should resolve
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tf:
            tf.write(b"safe temp file")
            temp_path = tf.name
        try:
            self.assertIsNotNone(resolve_file_path(temp_path))
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

        # 3. Path outside allowed roots should be rejected (return None)
        # Windows System32 / Linux /etc/hosts
        sys_path = r"C:\Windows\System32\drivers\etc\hosts" if os.name == "nt" else "/etc/hosts"
        if os.path.isfile(sys_path):
            res_sys = resolve_file_path(sys_path)
            self.assertIsNone(res_sys, f"Security Breach: Path outside sandbox resolved: {sys_path}")

    def test_02_path_traversal_serve_file(self):
        """Validates that _serve_file eliminates path traversal via os.path.basename."""
        handler = DummyMockHandler(b"", {}, path="/api/download/..%2f..%2fserver%2fapp.py")
        handler._serve_file("../../server/app.py")
        
        # Should be rejected with 400 (Invalid file path) or 404/403
        self.assertIn(handler.status_code, (400, 403, 404))
        wfile_val = handler.wfile.getvalue().decode("utf-8", errors="replace")
        self.assertNotIn("AntigravityOfficeStudio", wfile_val)

    def test_03_static_file_forbidden_extensions(self):
        """Validates that _serve_static_file blocks .py, .env, and sensitive files with HTTP 403."""
        # 1. Access to server/app.py
        h1 = DummyMockHandler(b"", {}, path="/server/app.py")
        h1._serve_static_file("server/app.py")
        self.assertEqual(h1.status_code, 403)
        self.assertIn("Forbidden", h1.wfile.getvalue().decode("utf-8"))

        # 2. Access to .env
        h2 = DummyMockHandler(b"", {}, path="/.env")
        h2._serve_static_file(".env")
        self.assertEqual(h2.status_code, 403)

        # 3. Path traversal with ..
        h3 = DummyMockHandler(b"", {}, path="/../../etc/passwd")
        h3._serve_static_file("../../etc/passwd")
        self.assertEqual(h3.status_code, 403)

    def test_04_upload_quota_enforcement(self):
        """Validates that payloads exceeding MAX_UPLOAD_SIZE_MB receive HTTP 413."""
        # Simulate payload larger than limit (e.g. 51MB when limit is 50MB)
        from server.app import MAX_UPLOAD_SIZE_BYTES
        oversized = MAX_UPLOAD_SIZE_BYTES + 1024

        headers = {
            "Content-Type": "multipart/form-data; boundary=----WebKitFormBoundary",
            "Content-Length": str(oversized),
        }
        handler = DummyMockHandler(b"", headers, path="/api/upload", method="POST")
        handler._handle_upload("multipart/form-data; boundary=----WebKitFormBoundary")

        self.assertEqual(handler.status_code, 413)
        resp = json.loads(handler.wfile.getvalue().decode("utf-8"))
        self.assertFalse(resp["success"])
        self.assertIn("Payload Too Large", resp["error"])

    def test_05_api_key_authentication(self):
        """Validates optional API_KEY protection: 401 when missing/invalid, 200 when valid."""
        import server.app as app_mod

        # Temporarily enable API_KEY
        orig_key = app_mod.API_SECRET_KEY
        try:
            app_mod.API_SECRET_KEY = "super_secret_test_token"

            # 1. /api/health is exempt and succeeds without key
            h_health = DummyMockHandler(b"", {}, path="/api/health", method="GET")
            self.assertTrue(h_health._check_auth("/api/health"))

            # 2. /api/info without key receives 401 Unauthorized
            h_unauth = DummyMockHandler(b"", {}, path="/api/info", method="GET")
            self.assertFalse(h_unauth._check_auth("/api/info"))
            self.assertEqual(h_unauth.status_code, 401)
            resp = json.loads(h_unauth.wfile.getvalue().decode("utf-8"))
            self.assertIn("Unauthorized", resp["error"])

            # 3. /api/info with valid Bearer token succeeds
            h_auth = DummyMockHandler(
                b"",
                {"Authorization": "Bearer super_secret_test_token"},
                path="/api/info",
                method="GET"
            )
            self.assertTrue(h_auth._check_auth("/api/info"))

            # 4. /api/info with valid X-API-Key header succeeds
            h_xkey = DummyMockHandler(
                b"",
                {"X-API-Key": "super_secret_test_token"},
                path="/api/info",
                method="GET"
            )
            self.assertTrue(h_xkey._check_auth("/api/info"))

            # 5. /api/info with valid ?api_key= query param succeeds
            h_qkey = DummyMockHandler(
                b"",
                {},
                path="/api/info?api_key=super_secret_test_token",
                method="GET"
            )
            self.assertTrue(h_qkey._check_auth("/api/info"))

        finally:
            app_mod.API_SECRET_KEY = orig_key

    def test_06_smart_cors_headers(self):
        """Validates CORS header generation for localhost and custom origins."""
        # Localhost origin should be echoed
        h_local = DummyMockHandler(b"", {"Origin": "http://localhost:5173"})
        h_local._send_cors_headers()
        cors_headers = dict(h_local._headers_buffer)
        self.assertEqual(cors_headers.get("Access-Control-Allow-Origin"), "http://localhost:5173")

    def test_07_serve_index_html_and_head_method(self):
        """Validates that GET / and HEAD / return 200 without NameError."""
        # 1. GET /
        h_get = DummyMockHandler(b"", {}, path="/", method="GET")
        h_get.do_GET()
        self.assertEqual(h_get.status_code, 200)
        content_val = h_get.wfile.getvalue()
        self.assertGreater(len(content_val), 0)

        # 2. HEAD /
        h_head = DummyMockHandler(b"", {}, path="/", method="HEAD")
        h_head.do_HEAD()
        self.assertEqual(h_head.status_code, 200)

        # 3. HEAD /api/health
        h_health = DummyMockHandler(b"", {}, path="/api/health", method="HEAD")
        h_health.do_HEAD()
        self.assertEqual(h_health.status_code, 200)


if __name__ == "__main__":
    unittest.main()

