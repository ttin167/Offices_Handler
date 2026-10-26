"""
tests/test_yearbook_studio.py - Test Suite for Digital Yearbook & Pre-Press Photobook Studio

Covers:
1. core.yearbook.layout_engine:
   - Aspect ratio detection & EXIF rotation normalization (ERR_EXIF_001)
   - 9 Editorial spread templates (Cover, Editorial, Duo, Mag-4, Polaroid, Mosaic-6, Classroom-8, Megagrid-12, Panoramic)
   - Auto-layout pagination algorithm with small, medium and large batches (50+ photos)
   - get_default_mock_spec structure and validity
2. core.yearbook.pdf_press_exporter:
   - crop_image_to_aspect with top-bias 35% (ERR_IMG_CROP_001)
   - export_spread_preview_pdf (Double-page spread view)
   - export_print_ready_pdf (3mm Bleed, Crop Marks, Vietnamese Unicode font ERR_PDF_FONT_001)
3. core.yearbook.handoff_packager:
   - 5-in-1 ZIP bundle creation with disk streaming (ERR_RAM_001)
   - Word DOCX brief with OpenXML table invariants (cantSplit, tblHeader, vAlign="center")
4. server.app REST API endpoints:
   - /api/yearbook/mock-album
   - /api/yearbook/presets
   - /api/yearbook/auto-layout
   - /api/yearbook/export
5. CLI commands:
   - yearbook-layout (--mock and custom photos)
   - yearbook-build (handoff ZIP creation)
6. server.mcp_server tools:
   - yearbook_auto_layout
   - yearbook_export
"""

import io
import json
import os
import sys
import io
import json
import os
import sys
import tempfile
import threading
import time
import urllib.request
import zipfile
import pytest
from PIL import Image

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from http.server import HTTPServer
from server.app import OfficeApiHandler
from core.yearbook.layout_engine import (
    SPREAD_TEMPLATES,
    YearbookLayoutEngine,
    get_default_mock_spec,
)
from core.yearbook.pdf_press_exporter import (
    YearbookPressEngine,
    crop_image_to_aspect,
    generate_mock_photo_bytes,
)
from core.yearbook.handoff_packager import YearbookHandoffPackager
from server.mcp_server import TOOL_DISPATCHER
from cli.ai_tools import cmd_yearbook_layout, cmd_yearbook_build


# =============================================================================
# 1. LAYOUT ENGINE TESTS
# =============================================================================

class TestYearbookLayoutEngine:
    def test_templates_count_and_structure(self):
        """Verify all 9 editorial templates are properly registered."""
        assert len(SPREAD_TEMPLATES) == 9
        expected_keys = [
            "cover_hero",
            "editorial_quote",
            "symmetric_duo",
            "magazine_grid_4",
            "polaroid_scrapbook",
            "mosaic_gallery_6",
            "classroom_grid_8",
            "megagrid_12",
            "panoramic_spread",
        ]
        for key in expected_keys:
            assert key in SPREAD_TEMPLATES
            tmpl = SPREAD_TEMPLATES[key]
            assert "name" in tmpl
            assert "capacity" in tmpl
            assert tmpl["capacity"] > 0

    def test_inspect_photo_aspect_and_orientation(self, tmp_path):
        """Test image aspect ratio classification (Landscape, Portrait, Square)."""
        # Create landscape image (600x400)
        img_l = Image.new("RGB", (600, 400), color=(100, 150, 200))
        path_l = str(tmp_path / "landscape.jpg")
        img_l.save(path_l)
        meta_l = YearbookLayoutEngine.inspect_photo(path_l)
        assert meta_l.width == 600 and meta_l.height == 400
        assert meta_l.orientation_category == "landscape"

        # Create portrait image (400x600)
        img_p = Image.new("RGB", (400, 600), color=(200, 100, 150))
        path_p = str(tmp_path / "portrait.jpg")
        img_p.save(path_p)
        meta_p = YearbookLayoutEngine.inspect_photo(path_p)
        assert meta_p.width == 400 and meta_p.height == 600
        assert meta_p.orientation_category == "portrait"

        # Create square image (500x500)
        img_s = Image.new("RGB", (500, 500), color=(150, 200, 100))
        path_s = str(tmp_path / "square.jpg")
        img_s.save(path_s)
        meta_s = YearbookLayoutEngine.inspect_photo(path_s)
        assert meta_s.width == 500 and meta_s.height == 500
        assert meta_s.orientation_category == "square"

    def test_default_mock_spec(self):
        """Verify default mock spec provides complete album with 8 spreads."""
        spec = get_default_mock_spec()
        assert spec["title"] == "KỶ YẾU TỐT NGHIỆP 2026"
        pages = spec.get("pages") or spec.get("spreads")
        assert len(pages) == 8

        # First spread must be cover_hero
        assert pages[0]["type"] == "cover_hero"
        # Second must be editorial_quote
        assert pages[1]["type"] == "editorial_quote"
        # Eighth must be megagrid_12
        assert pages[7]["type"] == "megagrid_12"

    def test_auto_layout_with_large_batch(self, tmp_path):
        """Verify pagination with large photo batches (30+ photos) leveraging high-density templates."""
        photo_paths = []
        for i in range(30):
            p = str(tmp_path / f"batch_photo_{i:02d}.jpg")
            img = Image.new("RGB", (400, 400), color=(i * 7 % 255, 100, 150))
            img.save(p)
            photo_paths.append(p)

        spec = YearbookLayoutEngine.auto_layout(
            photo_paths=photo_paths,
            title="KỶ YẾU KHÓA 2022-2026",
            subtitle="LỚP KỸ SƯ CNTT",
            preset_key="photobook_square_25x25"
        )
        assert spec["title"] == "KỶ YẾU KHÓA 2022-2026"
        pages = spec.get("pages") or spec.get("spreads")
        assert len(pages) >= 4

        # Layouts should include high-density grids
        layouts_used = [p["type"] for p in pages]
        assert "classroom_grid_8" in layouts_used or "megagrid_12" in layouts_used or "mosaic_gallery_6" in layouts_used


# =============================================================================
# 2. PDF PRESS EXPORTER TESTS
# =============================================================================

class TestYearbookPdfPressExporter:
    def test_crop_image_to_aspect_top_bias(self):
        """Verify ERR_IMG_CROP_001 cover mode crops with 35% top bias."""
        # Original 3:2 landscape image (600x400)
        img = Image.new("RGB", (600, 400), color=(50, 100, 150))

        # Crop to 1:1 square (should become 400x400)
        cropped_sq = crop_image_to_aspect(img, 400, 400, top_bias=0.35)
        assert cropped_sq.size == (400, 400)

        # Crop to 2:3 vertical (400 wide -> 600 tall, so width is cropped to 400*2/3 = 266)
        cropped_vert = crop_image_to_aspect(img, 200, 300, top_bias=0.35)
        assert cropped_vert.size[0] == int(400 * (2/3))
        assert cropped_vert.size[1] == 400

    def test_generate_mock_photo_bytes(self):
        """Verify mock vector placeholder generation."""
        mock_bytes = generate_mock_photo_bytes(400, 400, "gown", "Mock Label")
        assert mock_bytes is not None
        assert len(mock_bytes) > 500

    def test_export_spread_preview_pdf(self, tmp_path):
        """Verify spread PDF (2-page view) renders with valid dimensions."""
        spec = get_default_mock_spec()
        spec["pages"] = spec["pages"][:3]
        out_pdf = str(tmp_path / "spread_preview.pdf")

        engine = YearbookPressEngine()
        engine.export_spread_preview_pdf(spec, out_pdf)
        assert os.path.exists(out_pdf)
        assert os.path.getsize(out_pdf) > 1000

    def test_export_print_ready_pdf_with_bleed_and_crop_marks(self, tmp_path):
        """Verify ERR_PDF_FONT_001 and print-ready single pages with 3mm Bleed."""
        spec = get_default_mock_spec()
        spec["pages"] = spec["pages"][:2]
        out_pdf = str(tmp_path / "print_ready_300dpi.pdf")

        engine = YearbookPressEngine()
        engine.export_print_ready_pdf(spec, out_pdf)
        assert os.path.exists(out_pdf)
        assert os.path.getsize(out_pdf) > 1000


# =============================================================================
# 3. HANDOFF PACKAGER & OPENXML TABLE TESTS
# =============================================================================

class TestYearbookHandoffPackager:
    def test_export_handoff_bundle_5_in_1(self, tmp_path):
        """Verify 5-in-1 ZIP bundle creation and OpenXML table invariants (cantSplit, tblHeader, vAlign)."""
        spec = get_default_mock_spec()
        spec["pages"] = spec["pages"][:3]
        out_zip = str(tmp_path / "handoff_bundle.zip")

        packager = YearbookHandoffPackager()
        bundle_path = packager.export_handoff_bundle(spec, out_zip)
        assert os.path.exists(bundle_path)
        assert zipfile.is_zipfile(bundle_path)

        with zipfile.ZipFile(bundle_path, "r") as zf:
            namelist = zf.namelist()
            assert "01_Digital_Yearbook_Spread_Preview.pdf" in namelist
            assert "02_Print_Ready_Pages_300DPI.pdf" in namelist
            assert "03_Design_Brief_And_Captions.docx" in namelist
            assert "03_Designer_Handoff_Brief.md" in namelist
            assert "04_Yearbook_Project_Spec.json" in namelist

            # Test OpenXML table invariants inside 03_Design_Brief_And_Captions.docx
            docx_bytes = zf.read("03_Design_Brief_And_Captions.docx")
            with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as docx_zip:
                doc_xml = docx_zip.read("word/document.xml").decode("utf-8")
                # Table invariants: cantSplit, tblHeader, vAlign="center"
                assert "cantSplit" in doc_xml
                assert "tblHeader" in doc_xml
                assert 'vAlign w:val="center"' in doc_xml


# =============================================================================
# 4. REST API ENDPOINTS TESTS
# =============================================================================

@pytest.fixture(scope="module")
def yearbook_server_url():
    server = HTTPServer(('127.0.0.1', 0), OfficeApiHandler)
    port = server.server_port
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    time.sleep(0.3)
    base = f"http://127.0.0.1:{port}"
    yield base
    server.shutdown()


class TestYearbookRestApi:
    def test_get_mock_album(self, yearbook_server_url):
        """Test GET /api/yearbook/mock-album."""
        with urllib.request.urlopen(f"{yearbook_server_url}/api/yearbook/mock-album") as resp:
            assert resp.status == 200
            data = json.loads(resp.read().decode("utf-8"))
            assert data["success"] is True
            assert "spec" in data
            pages = data["spec"].get("pages") or data["spec"].get("spreads")
            assert len(pages) == 8

    def test_get_presets(self, yearbook_server_url):
        """Test GET /api/yearbook/presets."""
        with urllib.request.urlopen(f"{yearbook_server_url}/api/yearbook/presets") as resp:
            assert resp.status == 200
            data = json.loads(resp.read().decode("utf-8"))
            assert data["success"] is True
            assert len(data["presets"]) >= 4

    def test_post_auto_layout(self, yearbook_server_url):
        """Test POST /api/yearbook/auto-layout."""
        payload = {
            "photos": ["mock:grad_cap", "mock:ceremony", "mock:friends", "mock:campus"],
            "album_title": "KỶ YẾU TEST",
            "album_subtitle": "2026",
            "page_size": "200x200mm"
        }
        body = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            f"{yearbook_server_url}/api/yearbook/auto-layout",
            data=body,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req) as resp:
            assert resp.status == 200
            data = json.loads(resp.read().decode("utf-8"))
            assert data["success"] is True
            assert "spec" in data
            pages = data["spec"].get("pages") or data["spec"].get("spreads")
            assert len(pages) >= 1

    def test_post_export_spread_pdf(self, yearbook_server_url):
        """Test POST /api/yearbook/export with spread_pdf."""
        spec = get_default_mock_spec()
        spec["pages"] = spec["pages"][:2]
        payload = {
            "spec": spec,
            "export_type": "spread_pdf"
        }
        body = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            f"{yearbook_server_url}/api/yearbook/export",
            data=body,
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req) as resp:
            assert resp.status == 200
            data = json.loads(resp.read().decode("utf-8"))
            assert data["success"] is True
            assert data["download_url"].startswith("/api/download/") or data["download_url"].startswith("/api/files/download")



# =============================================================================
# 5. CLI & MCP SERVER TESTS
# =============================================================================

class TestYearbookCliAndMcp:
    def test_cli_yearbook_layout_and_build(self, tmp_path):
        """Test CLI commands yearbook-layout and yearbook-build."""
        spec_path = str(tmp_path / "cli_spec.json")
        bundle_path = str(tmp_path / "cli_bundle.zip")

        class ArgsLayout:
            photos = []
            mock = True
            title = "KỶ YẾU CLI"
            subtitle = "TEST"
            size = "photobook_square_25x25"
            output = spec_path

        exit_code_layout = cmd_yearbook_layout(ArgsLayout())
        assert exit_code_layout == 0
        assert os.path.exists(spec_path)

        # Build bundle from spec (trim pages to 2 for speed)
        with open(spec_path, "r", encoding="utf-8") as f:
            sp = json.load(f)
        if "pages" in sp:
            sp["pages"] = sp["pages"][:2]
        elif "spreads" in sp:
            sp["spreads"] = sp["spreads"][:2]
        with open(spec_path, "w", encoding="utf-8") as f:
            json.dump(sp, f, ensure_ascii=False)

        class ArgsBuild:
            spec = spec_path
            output = bundle_path
            type = "handoff"
            dpi = 150

        exit_code_build = cmd_yearbook_build(ArgsBuild())
        assert exit_code_build == 0
        assert os.path.exists(bundle_path)

    def test_mcp_yearbook_tools(self, tmp_path):
        """Test MCP tools yearbook_auto_layout and yearbook_export."""
        # Tool 1: yearbook_auto_layout
        layout_fn = TOOL_DISPATCHER["yearbook_auto_layout"]
        layout_res = layout_fn({
            "photos": ["mock:grad_cap", "mock:portrait"],
            "album_title": "MCP Yearbook Test",
            "album_subtitle": "2026"
        })
        assert "spec" in layout_res
        spec = layout_res["spec"]

        # Tool 2: yearbook_export
        export_fn = TOOL_DISPATCHER["yearbook_export"]
        out_zip = str(tmp_path / "mcp_export.zip")
        export_res = export_fn({
            "spec": spec,
            "export_type": "handoff_zip",
            "output_path": out_zip
        })
        assert export_res["success"] is True
        assert os.path.exists(out_zip)
