"""
Unit test suite for Native PowerPoint Engine (pptx_writer, pptx_reader, pptx_engine).
Validates PPTX_INV_01 .. PPTX_INV_08 invariants.
"""

from __future__ import annotations

import json
import os
import tempfile
import unittest

from PIL import Image

import pptx
from pptx import Presentation

from pptx_writer import (
    write_pptx_from_spec,
    init_presentation,
    SLIDE_WIDTH_IN,
    SLIDE_HEIGHT_IN,
    THEMES,
)
from pptx_reader import read_pptx, read_pptx_to_json
from pptx_engine import (
    compile_markdown_file_to_pptx,
    parse_markdown_to_slides_spec,
    insert_diagram_into_pptx,
)


class TestPowerPointEngine(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.dpath = self.temp_dir.name

        # Create a sample test diagram PNG image
        self.test_img = os.path.join(self.dpath, "sample_diagram.png")
        im = Image.new("RGB", (1200, 600), color=(30, 41, 59))
        im.save(self.test_img)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_widescreen_dimensions_inv_01(self):
        """Validates PPTX_INV_01: Strict 16:9 Widescreen dimensions."""
        prs = init_presentation()
        self.assertAlmostEqual(prs.slide_width.inches, SLIDE_WIDTH_IN, places=2)
        self.assertAlmostEqual(prs.slide_height.inches, SLIDE_HEIGHT_IN, places=2)
        ratio = prs.slide_width.inches / prs.slide_height.inches
        self.assertAlmostEqual(ratio, 16.0 / 9.0, places=2)

    def test_write_and_read_spec_deck(self):
        """Validates JSON Spec compilation and AST extraction via pptx_reader."""
        spec = {
            "theme": "corporate_blue",
            "slides": [
                {
                    "type": "title",
                    "title": "Smart Mart Architecture",
                    "subtitle": "High-Fidelity Specification",
                    "presenter": "Antigravity Team",
                    "date": "2026-10-01",
                },
                {
                    "type": "content",
                    "title": "Core Modules Overview",
                    "subtitle": "Decoupled Architecture",
                    "bullets": [
                        ("High-speed OpenXML Word engine", 0),
                        ("Full HD mxGraph Draw.io renderer", 0),
                        ("Nested sub-bullet point item", 1),
                    ],
                },
                {
                    "type": "two_column",
                    "title": "Comparison Analysis",
                    "left_bullets": ["Local Offline Mode", "Zero Cloud Latency"],
                    "right_bullets": ["Felo AI Cloud Mode", "Rapid Ideation"],
                },
                {
                    "type": "diagram",
                    "title": "System Topology Diagram",
                    "image_path": self.test_img,
                    "caption": "Figure 1: High-level System Topology (300 DPI)",
                },
                {
                    "type": "kpi_cards",
                    "title": "Performance Metrics",
                    "cards": [
                        {"title": "Throughput", "value": "1.2k/s", "description": "Requests processed"},
                        {"title": "Uptime", "value": "99.99%", "description": "SLA guarantee"},
                        {"title": "Accuracy", "value": "100%", "description": "Zero XML corruption"},
                    ],
                },
                {
                    "type": "table",
                    "title": "Benchmark Results",
                    "headers": ["Engine", "Output Format", "Speed (ms)", "Fidelity"],
                    "rows": [
                        ["docx_writer", ".docx", 45, "100%"],
                        ["mxgraph_engine", ".drawio", 120, "Full HD"],
                        ["pptx_writer", ".pptx", 35, "16:9 Native"],
                    ],
                },
            ],
        }

        out_pptx = os.path.join(self.dpath, "full_deck.pptx")
        saved = write_pptx_from_spec(spec, out_pptx, theme_name="corporate_blue")
        self.assertTrue(os.path.isfile(saved))

        # Inspect with pptx_reader
        ast = read_pptx(saved)
        self.assertEqual(ast["total_slides"], 6)
        self.assertEqual(ast["dimensions"]["aspect_ratio"], "16:9")

        # Verify Slide 1
        s1 = ast["slides"][0]
        self.assertIn("Smart Mart Architecture", s1["title"])

        # Verify Slide 6 (Table)
        s6 = ast["slides"][5]
        tbl_shape = next((s for s in s6["shapes"] if "table" in s), None)
        self.assertIsNotNone(tbl_shape)
        self.assertEqual(tbl_shape["table"]["num_rows"], 4)
        self.assertEqual(tbl_shape["table"]["num_cols"], 4)

    def test_markdown_slide_compiler(self):
        """Validates Markdown to PowerPoint slide compilation."""
        md_content = f"""# Master Platform Deck
Subtitle: Next-Gen Enterprise Engineering

---
## Technical Highlights
- Pure Native OpenXML Generation
- Multi-Engine Draw.io and PlantUML
  - Nested architectural layer
- Automated 12-Gate Quality Diff

---
## Component Diagram
![Architecture Blueprint]({self.test_img})

---
## Key Data
| Metric | Q1 | Q2 |
| Active Users | 10k | 25k |
| Latency | 5ms | 4ms |
"""
        md_path = os.path.join(self.dpath, "slides.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        out_pptx = os.path.join(self.dpath, "md_deck.pptx")
        compiled = compile_markdown_file_to_pptx(md_path, out_pptx, theme_name="modern_dark")
        self.assertTrue(os.path.isfile(compiled))

        ast = read_pptx(compiled)
        self.assertEqual(ast["total_slides"], 4)
        self.assertEqual(ast["slides"][0]["title"], "Master Platform Deck")
        self.assertEqual(ast["slides"][1]["title"], "Technical Highlights")
        self.assertEqual(ast["slides"][2]["title"], "Component Diagram")

    def test_insert_diagram_into_existing_pptx(self):
        """Validates in-place diagram injection into an existing PPTX presentation."""
        # Create minimal base presentation
        base_pptx = os.path.join(self.dpath, "base.pptx")
        prs = init_presentation()
        prs.slides.add_slide(prs.slide_layouts[6])
        prs.save(base_pptx)

        # Insert diagram
        updated = insert_diagram_into_pptx(
            pptx_path=base_pptx,
            image_path=self.test_img,
            title="Appended Diagram Slide",
            caption="Figure 2: In-place injection test",
        )
        self.assertTrue(os.path.isfile(updated))

        ast = read_pptx(updated)
        self.assertEqual(ast["total_slides"], 2)
        s2 = ast["slides"][1]
        self.assertEqual(s2["title"], "Appended Diagram Slide")


if __name__ == "__main__":
    unittest.main()
