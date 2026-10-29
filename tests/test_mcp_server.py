"""
tests/test_mcp_server.py — Unit Tests for Model Context Protocol (MCP) Server.

Tests:
1. JSON-RPC 2.0 initialize request / response.
2. ping request / response.
3. notifications/initialized handling.
4. tools/list schema verification (all 6 registered tools).
5. tools/call execution for env_doctor.
6. Error handling for unknown tools.
7. Error handling for unknown methods.
8. Root facade re-export parity.
"""

from __future__ import annotations

import json
import os
import unittest
from server.mcp_server import (
    handle_rpc_request,
    TOOL_DEFINITIONS,
    SERVER_NAME,
    SERVER_VERSION,
)
import mcp_server


class TestMCPServer(unittest.TestCase):
    def test_01_initialize(self):
        req = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "test-client", "version": "1.0.0"}
            }
        }
        res = handle_rpc_request(req)
        self.assertIsNotNone(res)
        self.assertEqual(res["id"], 1)
        self.assertEqual(res["jsonrpc"], "2.0")
        self.assertIn("result", res)
        self.assertEqual(res["result"]["serverInfo"]["name"], SERVER_NAME)
        self.assertEqual(res["result"]["serverInfo"]["version"], SERVER_VERSION)

    def test_02_ping(self):
        req = {"jsonrpc": "2.0", "id": 42, "method": "ping"}
        res = handle_rpc_request(req)
        self.assertIsNotNone(res)
        self.assertEqual(res["id"], 42)
        self.assertEqual(res["result"], {})

    def test_03_notification_initialized(self):
        req = {"jsonrpc": "2.0", "method": "notifications/initialized"}
        res = handle_rpc_request(req)
        self.assertIsNone(res)

    def test_04_tools_list_registers_six_tools(self):
        req = {"jsonrpc": "2.0", "id": 100, "method": "tools/list"}
        res = handle_rpc_request(req)
        self.assertIsNotNone(res)
        tools = res["result"]["tools"]
        tool_names = [t["name"] for t in tools]

        expected_tools = [
            "doc_convert",
            "doc_inspect",
            "doc_insert_diagram",
            "diagram_render",
            "pptx_synthesize",
            "env_doctor",
            "docx_apply_preset",
            "pdf_quick_tools",
            "doc_mail_merge",
            "doc_redline_diff",
            "doc_ocr_scan",
            "xlsx_set_cell",
            "xlsx_set_range",
            "xlsx_insert_rows",
            "xlsx_find_anchor",
            "docx_inspect_placeholders",
            "docx_replace_text",
            "docx_insert_paragraph",
            "docx_insert_table",
            "xlsx_formula_helper",
            "yearbook_auto_layout",
            "yearbook_export",
            "doc_security_audit",
            "pdf_crypto_sign_and_pdfa",
            "xlsx_to_brief_report",
            "xlsx_consolidate_workbooks",
            "doc_run_pipeline",
            "diagram_from_code",
        ]
        self.assertEqual(len(tools), 28)
        for expected in expected_tools:
            self.assertIn(expected, tool_names)

        # Validate JSON Schema of tools
        for tool in tools:
            self.assertIn("name", tool)
            self.assertIn("description", tool)
            self.assertIn("inputSchema", tool)
            self.assertEqual(tool["inputSchema"]["type"], "object")

    def test_05_tools_call_env_doctor(self):
        req = {
            "jsonrpc": "2.0",
            "id": 101,
            "method": "tools/call",
            "params": {
                "name": "env_doctor",
                "arguments": {}
            }
        }
        res = handle_rpc_request(req)
        self.assertIsNotNone(res)
        self.assertEqual(res["id"], 101)
        self.assertIn("result", res)
        self.assertFalse(res["result"]["isError"])
        content = res["result"]["content"]
        self.assertEqual(len(content), 1)
        self.assertEqual(content[0]["type"], "text")
        parsed = json.loads(content[0]["text"])
        self.assertIn("status", parsed)
        self.assertIn("components", parsed)

    def test_06_tools_call_unknown_tool(self):
        req = {
            "jsonrpc": "2.0",
            "id": 999,
            "method": "tools/call",
            "params": {
                "name": "non_existent_tool_xyz",
                "arguments": {}
            }
        }
        res = handle_rpc_request(req)
        self.assertIsNotNone(res)
        self.assertEqual(res["id"], 999)
        self.assertIn("error", res)
        self.assertEqual(res["error"]["code"], -32601)

    def test_07_unknown_method(self):
        req = {"jsonrpc": "2.0", "id": 555, "method": "unknown/arbitrary_method"}
        res = handle_rpc_request(req)
        self.assertIsNotNone(res)
        self.assertEqual(res["id"], 555)
        self.assertIn("error", res)
        self.assertEqual(res["error"]["code"], -32601)

    def test_08_root_facade_parity(self):
        self.assertEqual(mcp_server.handle_rpc_request, handle_rpc_request)
        self.assertEqual(mcp_server.TOOL_DEFINITIONS, TOOL_DEFINITIONS)

    def test_09_tools_call_doc_redline_diff(self):
        import tempfile
        from docx import Document
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f_a, \
             tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f_b, \
             tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f_out:
            p_a, p_b, p_out = f_a.name, f_b.name, f_out.name

        try:
            d_a = Document()
            d_a.add_paragraph("Original text A.")
            d_a.save(p_a)

            d_b = Document()
            d_b.add_paragraph("Revised text B.")
            d_b.save(p_b)

            req = {
                "jsonrpc": "2.0",
                "id": 102,
                "method": "tools/call",
                "params": {
                    "name": "doc_redline_diff",
                    "arguments": {
                        "doc_a_path": p_a,
                        "doc_b_path": p_b,
                        "output_path": p_out,
                        "author": "AI Test Auditor"
                    }
                }
            }
            res = handle_rpc_request(req)
            self.assertIsNotNone(res)
            self.assertEqual(res["id"], 102)
            self.assertIn("result", res)
            self.assertFalse(res["result"]["isError"])
            data = json.loads(res["result"]["content"][0]["text"])
            self.assertIn("insertions_count", data)
            self.assertIn("deletions_count", data)
        finally:
            for p in [p_a, p_b, p_out]:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass

    def test_10_tools_call_doc_security_audit(self):
        import tempfile
        with tempfile.NamedTemporaryFile(mode="w", suffix=".csv", delete=False, encoding="utf-8") as f:
            f.write("Name,Formula\nTest,=cmd|' /C calc'!A1\n")
            csv_path = f.name

        try:
            req = {
                "jsonrpc": "2.0",
                "id": 103,
                "method": "tools/call",
                "params": {
                    "name": "doc_security_audit",
                    "arguments": {
                        "file_path": csv_path,
                        "sanitize": True
                    }
                }
            }
            res = handle_rpc_request(req)
            self.assertIsNotNone(res)
            self.assertEqual(res["id"], 103)
            self.assertIn("result", res)
            self.assertFalse(res["result"]["isError"])
            data = json.loads(res["result"]["content"][0]["text"])
            self.assertTrue(data.get("success"))
            self.assertFalse(data.get("is_clean"))
            self.assertGreaterEqual(data.get("critical_count", 0), 1)
            self.assertIn("sanitization", data)
            san_path = data["sanitization"].get("output_path")
            if san_path and os.path.exists(san_path):
                os.remove(san_path)
        finally:
            if os.path.exists(csv_path):
                os.remove(csv_path)


if __name__ == "__main__":
    unittest.main()
