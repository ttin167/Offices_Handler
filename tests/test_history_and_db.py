"""
Unit tests for db_manager.py and web_api_server.py history & math endpoints.
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import threading
import time
import urllib.request
from http.server import HTTPServer

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import db_manager
from web_api_server import OfficeApiHandler


def test_db_manager_crud():
    """Validates low-level SQLite storage CRUD operations in isolation."""
    # Direct test with record and retrieve
    task_type = "test_unit"
    src_file = "sample_test.docx"
    out_file = "output_test.docx"
    
    # Create temp dummy file to verify file_size calculation
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(b"Mock binary content 1234567890")
        tmp_path = tmp.name

    try:
        row_id = db_manager.record_history(task_type, src_file, out_file, tmp_path)
        assert row_id > 0, f"Expected valid inserted row ID, got {row_id}"

        # Fetch history
        items = db_manager.get_recent_history(limit=10)
        assert len(items) > 0, "Expected non-empty history items"
        found = any(item["id"] == row_id and item["output_filename"] == out_file for item in items)
        assert found, "Inserted item not found in history"

        # Delete single item
        del_ok = db_manager.delete_history_item(row_id)
        assert del_ok is True, "Failed to delete single history item"

        items_after = db_manager.get_recent_history(limit=10)
        assert not any(item["id"] == row_id for item in items_after), "Deleted item still present"

    finally:
        if os.path.isfile(tmp_path):
            os.remove(tmp_path)


def test_web_api_history_and_math():
    """Validates /api/history, /api/math/parse, and /api/math/export endpoints."""
    test_port = 8991
    server = HTTPServer(('127.0.0.1', test_port), OfficeApiHandler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()
    time.sleep(0.4)

    base_url = f"http://127.0.0.1:{test_port}"

    try:
        # 1. GET /api/history
        with urllib.request.urlopen(f"{base_url}/api/history") as res:
            assert res.status == 200
            data = json.loads(res.read().decode('utf-8'))
            assert "items" in data
            assert isinstance(data["items"], list)

        # 2. POST /api/math/parse
        parse_payload = json.dumps({"latex": "v = \\frac{\\Delta s}{\\Delta t} + \\sqrt{x}"}).encode('utf-8')
        req_parse = urllib.request.Request(
            f"{base_url}/api/math/parse",
            data=parse_payload,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_parse) as res:
            assert res.status == 200
            p_data = json.loads(res.read().decode('utf-8'))
            assert p_data["success"] is True
            assert len(p_data["tokens"]) > 0

        # 3. POST /api/math/export
        export_payload = json.dumps({
            "title": "Báo Cáo Kiểm Thử Tự Động",
            "latex": "Phương trình động học:\n$$v(t) = v_{ref} + a \\cdot t$$\nVới gia tốc $a \\approx 9.8\\text{ m/s}^{2}$."
        }).encode('utf-8')
        req_export = urllib.request.Request(
            f"{base_url}/api/math/export",
            data=export_payload,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_export) as res:
            assert res.status == 200
            exp_data = json.loads(res.read().decode('utf-8'))
            assert exp_data["success"] is True
            assert "download_url" in exp_data
            out_filename = exp_data["filename"]
            assert out_filename.endswith(".docx")

        # 4. Verify History recorded the math export
        with urllib.request.urlopen(f"{base_url}/api/history") as res:
            assert res.status == 200
            data = json.loads(res.read().decode('utf-8'))
            items = data["items"]
            matching = [it for it in items if it["output_filename"] == out_filename]
            assert len(matching) > 0, "Exported DOCX was not automatically logged to history"
            target_id = matching[0]["id"]

        # 5. Download the exported file to ensure it is valid
        with urllib.request.urlopen(f"{base_url}/api/download/{out_filename}") as res:
            assert res.status == 200
            file_bytes = res.read()
            assert len(file_bytes) > 1000, "Exported DOCX is unexpectedly small"
            assert file_bytes[:2] == b"PK", "Exported DOCX does not have valid ZIP header"

        # 6. DELETE /api/history/<id>
        req_del = urllib.request.Request(
            f"{base_url}/api/history/{target_id}",
            headers={"Content-Type": "application/json"},
            method="DELETE"
        )
        with urllib.request.urlopen(req_del) as res:
            assert res.status == 200
            d_res = json.loads(res.read().decode('utf-8'))
            assert d_res.get("success") is True

    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    print("Testing db_manager CRUD...")
    test_db_manager_crud()
    print("  -> PASSED!")
    print("Testing web_api history & math endpoints...")
    test_web_api_history_and_math()
    print("  -> PASSED!")
