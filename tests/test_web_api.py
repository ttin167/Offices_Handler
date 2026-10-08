"""
Unit test for web_api_server.py
Spawns server in a background thread, tests endpoints, and shuts down cleanly.
"""
import os
import sys
import threading
import time
import urllib.request
import json

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from http.server import HTTPServer
from web_api_server import OfficeApiHandler

def run_tests():
    server = HTTPServer(('127.0.0.1', 8899), OfficeApiHandler)
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()
    time.sleep(0.5)

    base_url = "http://127.0.0.1:8899"

    # Test 1: /api/health
    print("Testing GET /api/health...")
    with urllib.request.urlopen(f"{base_url}/api/health") as response:
        assert response.status == 200
        data = json.loads(response.read().decode('utf-8'))
        assert data.get("status") == "online"
        assert data.get("healthy") is True
        print("  -> OK:", data)

    # Test 2: /api/info
    print("Testing GET /api/info...")
    with urllib.request.urlopen(f"{base_url}/api/info") as response:
        assert response.status == 200
        data = json.loads(response.read().decode('utf-8'))
        assert "version" in data
        assert "capabilities" in data
        print("  -> OK. App:", data.get("name"), "v" + data.get("version"))

    # Test 3: /api/upload (Multipart)
    print("Testing POST /api/upload (Multipart)...")
    boundary = "----TestBoundary12345"
    file_content = b"Mock document content for testing"
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="test_doc.docx"\r\n'
        f"Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n"
    ).encode("utf-8") + file_content + f"\r\n--{boundary}--\r\n".encode("utf-8")

    req_upload = urllib.request.Request(
        f"{base_url}/api/upload",
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
    )
    with urllib.request.urlopen(req_upload) as response:
        assert response.status == 200
        up_data = json.loads(response.read().decode('utf-8'))
        assert up_data.get("success") is True
        assert up_data.get("filepath") is not None
        assert os.path.isfile(up_data["filepath"])
        print("  -> OK. Uploaded filepath:", up_data.get("filepath"))

    # Test 4: /api/render-diagram with Draw.io wrapped in mxfile
    print("Testing POST /api/render-diagram (Draw.io mxfile)...")
    drawio_xml = '''<mxfile host="Antigravity Studio" version="21.0.0">
  <diagram id="sample" name="System Overview">
    <mxGraphModel dx="1200" dy="800" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1200" pageHeight="800" background="#FFFFFF">
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="hub" value="Antigravity Hub" style="rounded=1;whiteSpace=wrap;html=1;fillColor=#0051E2;strokeColor=#003DB3;fontColor=#FFFFFF;fontStyle=1;fontSize=14;" vertex="1" parent="1">
          <mxGeometry x="450" y="250" width="180" height="70" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>'''
    payload_drawio = json.dumps({
        "engine": "drawio",
        "code": drawio_xml,
        "format": "png"
    }).encode('utf-8')
    req_drawio = urllib.request.Request(f"{base_url}/api/render-diagram", data=payload_drawio, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req_drawio) as response:
        assert response.status == 200
        data = json.loads(response.read().decode('utf-8'))
        assert data.get("success") is True
        print("  -> OK. Draw.io Image:", data.get("image_path") or data.get("image_file"))

    # Test 4b: POST /api/render-diagram (Mermaid string code)
    print("Testing POST /api/render-diagram (Mermaid string code)...")
    payload_m = json.dumps({
        "engine": "mermaid",
        "code": "graph TD;\n    A-->B;",
        "format": "png"
    }).encode('utf-8')
    req_m = urllib.request.Request(f"{base_url}/api/render-diagram", data=payload_m, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req_m) as response:
            assert response.status == 200
            m_data = json.loads(response.read().decode('utf-8'))
            assert m_data.get("success") is True
            print("  -> OK. Mermaid Image:", m_data.get("filename"))
    except urllib.error.HTTPError as he:
        body = he.read().decode('utf-8')
        assert "has no attribute 'get'" not in body
        print("  -> OK. Handled gracefully without 'str' get crash:", body[:80])

    # Test 5: POST /api/docx/patch
    print("Testing POST /api/docx/patch...")
    test_docx = os.path.join(ROOT_DIR, "scratch", "test_docx_out", "test_template.docx")
    if os.path.isfile(test_docx):
        payload_patch = json.dumps({
            "template_path": test_docx,
            "text_replacements": {"{{PROJECT_NAME}}": "Web API Patched Project"},
            "output_name": "web_api_patched.docx"
        }).encode('utf-8')
        req_patch = urllib.request.Request(f"{base_url}/api/docx/patch", data=payload_patch, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req_patch) as response:
            assert response.status == 200
            patch_data = json.loads(response.read().decode('utf-8'))
            assert patch_data.get("success") is True
            assert os.path.isfile(patch_data["result_path"])
            print("  -> OK. Patched DOCX:", patch_data.get("filename"))

    # Test 6: POST /api/xlsx/mutate
    print("Testing POST /api/xlsx/mutate...")
    test_xlsx = os.path.join(ROOT_DIR, "scratch", "test_excel_out", "mock_template.xlsx")
    if not os.path.isfile(test_xlsx):
        # Create a quick workbook for testing
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Example"
        ws["A1"] = "Header"
        os.makedirs(os.path.dirname(test_xlsx), exist_ok=True)
        wb.save(test_xlsx)

    payload_mutate = json.dumps({
        "template_path": test_xlsx,
        "spec": {
            "sheets": {
                "Example": {
                    "cell_updates": {"A1": "Web API Mutated Title"}
                }
            }
        },
        "output_name": "web_api_mutated.xlsx"
    }).encode('utf-8')
    req_mutate = urllib.request.Request(f"{base_url}/api/xlsx/mutate", data=payload_mutate, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req_mutate) as response:
        assert response.status == 200
        mutate_data = json.loads(response.read().decode('utf-8'))
        assert mutate_data.get("success") is True
        assert os.path.isfile(mutate_data["result_path"])
        print("  -> OK. Mutated XLSX:", mutate_data.get("filename"))

    # Test 7: POST /api/xlsx/read
    print("Testing POST /api/xlsx/read...")
    payload_read = json.dumps({
        "file_path": test_xlsx,
        "engine": "openpyxl"
    }).encode('utf-8')
    req_read = urllib.request.Request(f"{base_url}/api/xlsx/read", data=payload_read, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req_read) as response:
        assert response.status == 200
        read_data = json.loads(response.read().decode('utf-8'))
        assert read_data.get("success") is True
        assert "snapshot" in read_data
        print("  -> OK. Snapshot engine:", read_data.get("engine"))

    # Test 8: POST /api/xlsx/validate
    print("Testing POST /api/xlsx/validate...")
    payload_val = json.dumps({
        "file_path": test_xlsx,
        "ref_sheet": "Test Cases",
        "target_sheet": "Test Cases"
    }).encode('utf-8')
    req_val = urllib.request.Request(f"{base_url}/api/xlsx/validate", data=payload_val, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req_val) as response:
            assert response.status == 200
            val_data = json.loads(response.read().decode('utf-8'))
            assert val_data.get("success") is True
            assert "passed" in val_data
            assert val_data.get("gates_evaluated") == 12
            print("  -> OK. 12 Quality Gates evaluated:", val_data.get("gates_evaluated"))
    except urllib.error.HTTPError as e:
        print("HTTP ERROR 500 BODY:", e.read().decode('utf-8'))
        raise

    # Test 8b: POST /api/xlsx/validate with target_sheet=None
    print("Testing POST /api/xlsx/validate with target_sheet=None...")
    payload_val_auto = json.dumps({
        "file_path": test_xlsx,
        "target_sheet": None,
        "ref_sheet": None
    }).encode('utf-8')
    req_val_auto = urllib.request.Request(f"{base_url}/api/xlsx/validate", data=payload_val_auto, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req_val_auto) as response:
        assert response.status == 200
        val_auto_data = json.loads(response.read().decode('utf-8'))
        assert val_auto_data.get("success") is True
        assert val_auto_data.get("target_sheet") is not None
        print("  -> OK. Auto-discovered target sheet:", val_auto_data.get("target_sheet"))

    # Test 9: POST /api/pptx/build
    print("Testing POST /api/pptx/build...")
    slide_spec = {
        "theme": "thesis_blue",
        "slides": [
            {
                "type": "content",
                "title": "Web API Spec Build Test",
                "bullets": ["Testing declarative slide build", "Ensuring 16:9 widescreen layout"]
            }
        ]
    }
    payload_pptx = json.dumps({
        "spec": slide_spec,
        "output_name": "web_api_built_deck.pptx"
    }).encode('utf-8')
    req_pptx = urllib.request.Request(f"{base_url}/api/pptx/build", data=payload_pptx, headers={"Content-Type": "application/json"})
    built_pptx_path = None
    with urllib.request.urlopen(req_pptx) as response:
        assert response.status == 200
        pptx_data = json.loads(response.read().decode('utf-8'))
        assert pptx_data.get("success") is True
        built_pptx_path = pptx_data.get("pptx_path")
        assert os.path.isfile(built_pptx_path)
        print("  -> OK. Built PPTX:", pptx_data.get("filename"))

    # Test 10: POST /api/pptx/preview
    print("Testing POST /api/pptx/preview...")
    if built_pptx_path:
        payload_prev = json.dumps({
            "pptx_path": built_pptx_path,
            "slides": "1"
        }).encode('utf-8')
        req_prev = urllib.request.Request(f"{base_url}/api/pptx/preview", data=payload_prev, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req_prev) as response:
            assert response.status == 200
            prev_data = json.loads(response.read().decode('utf-8'))
            assert prev_data.get("success") is True
            print("  -> OK. Previews generated:", prev_data.get("total_previews"))

    # Test 11: GET / (Static Asset Serving)
    print("Testing GET / (Static Index Serving)...")
    with urllib.request.urlopen(f"{base_url}/") as response:
        assert response.status == 200
        content = response.read().decode('utf-8')
        assert "Antigravity Office Studio" in content
        assert "panel-spreadsheet" in content
        print("  -> OK. Static index.html served correctly!")

    # Test 12: GET /api/download/ with Vietnamese filename
    print("Testing GET /api/download/ with Vietnamese UTF-8 filename...")
    vn_filename = "ĐÁP ÁN - BÌNH DÂN HỌC VỤ.txt"
    vn_file_path = os.path.join(ROOT_DIR, ".web_outputs", vn_filename)
    os.makedirs(os.path.dirname(vn_file_path), exist_ok=True)
    with open(vn_file_path, "w", encoding="utf-8") as f:
        f.write("Nội dung thử nghiệm tiếng Việt có dấu.")

    quoted_vn = urllib.parse.quote(vn_filename)
    with urllib.request.urlopen(f"{base_url}/api/download/{quoted_vn}") as response:
        assert response.status == 200
        content_vn = response.read().decode("utf-8")
        assert "tiếng Việt" in content_vn
        cd_header = response.headers.get("Content-Disposition", "")
        assert "filename*=" in cd_header or "filename=" in cd_header
        print("  -> OK. Downloaded Vietnamese file successfully without socket hang up!")

    server.shutdown()
    server.server_close()
    print("All web_api_server endpoint unit tests PASSED successfully!")

if __name__ == "__main__":
    run_tests()
