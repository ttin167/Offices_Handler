"""
Unit test suite for Advanced Technical PowerPoint Layouts and CLI Integration.
Validates:
- add_split_diagram_slide (Card left + Diagram right)
- add_dual_diagram_slide (Card left + Dual stacked diagrams)
- add_metrics_summary_slide (N-column cards + bottom summary banner)
- Markdown compiler auto-categorization to split_diagram and dual_diagram
- ai_tools_cli pptx-build and pptx-preview commands
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest

from PIL import Image
from pptx import Presentation

from pptx_writer import (
    init_presentation,
    add_split_diagram_slide,
    add_dual_diagram_slide,
    add_metrics_summary_slide,
    write_pptx_from_spec,
    safe_save_pptx,
)
from pptx_reader import read_pptx
from pptx_engine import (
    parse_markdown_to_slides_spec,
    compile_markdown_file_to_pptx,
)


class TestAdvancedPowerPointLayouts(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.dpath = self.temp_dir.name

        # Create two sample dummy diagram PNGs
        self.img1 = os.path.join(self.dpath, "diagram1.png")
        im1 = Image.new("RGB", (800, 450), color=(15, 41, 74))
        im1.save(self.img1)

        self.img2 = os.path.join(self.dpath, "diagram2.png")
        im2 = Image.new("RGB", (800, 450), color=(0, 81, 226))
        im2.save(self.img2)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_split_diagram_slide(self):
        """Validates add_split_diagram_slide layout and prefix bullet parsing."""
        prs = init_presentation()
        slide = add_split_diagram_slide(
            prs=prs,
            title="Kinematic Model & Instantaneous Rotation Center",
            bullets=[
                ("Kinematic Model", "Unicycle model with Instantaneous Rotation Center along axle."),
                ("Forward Kinematics", "Linear speed v = r(w_R + w_L)/2; Angular speed w = r(w_R - w_L)/b."),
                ("Wheelbase", "Effective track wheelbase b = 216 mm."),
            ],
            image_path=self.img1,
            caption="Figure 5.12: Instantaneous Rotation Center coordinate frames",
            banner="Chapter 5. Controller Design",
            theme_name="thesis_blue",
        )
        self.assertIsNotNone(slide)
        out_path = os.path.join(self.dpath, "split_test.pptx")
        saved = safe_save_pptx(prs, out_path)
        self.assertTrue(os.path.isfile(saved))

        # Inspect AST
        ast = read_pptx(saved)
        self.assertEqual(ast["total_slides"], 1)
        s1 = ast["slides"][0]
        all_text = " ".join(
            p["text"] for sh in s1["shapes"] if "paragraphs" in sh for p in sh["paragraphs"]
        )
        self.assertIn("Kinematic Model", all_text)
        self.assertIn("Chapter 5", all_text)
        self.assertGreaterEqual(len(s1["shapes"]), 3)

    def test_dual_diagram_slide(self):
        """Validates add_dual_diagram_slide with stacked comparative figures."""
        prs = init_presentation()
        slide = add_dual_diagram_slide(
            prs=prs,
            title="Locomotion & Drive Layout Selection",
            bullets=[
                ("Structure 1 (Selected)", "Two coaxial driving wheels + 1 front caster."),
                ("Structure 2 (Rejected)", "Two driving wheels + 5 passive support rollers."),
            ],
            img1_path=self.img1,
            cap1="Figure 1.4b: Structure 1 — Selected",
            img2_path=self.img2,
            cap2="Figure 1.6b: Structure 2 — Rejected",
            banner="Chapter 2. Conceptual Design",
            orientation="vertical",
            theme_name="corporate_blue",
        )
        self.assertIsNotNone(slide)
        out_path = os.path.join(self.dpath, "dual_test.pptx")
        saved = safe_save_pptx(prs, out_path)
        self.assertTrue(os.path.isfile(saved))

        ast = read_pptx(saved)
        self.assertEqual(ast["total_slides"], 1)

    def test_metrics_summary_slide(self):
        """Validates add_metrics_summary_slide 3-column cards and bottom banner."""
        prs = init_presentation()
        cards = [
            {
                "title": "Left Motor PI",
                "metrics": [
                    ("Proportional Gain", "kP = 0.301"),
                    ("Integral Gain", "kI = 12.745"),
                    ("Settling Time", "0.14 s"),
                ]
            },
            {
                "title": "Right Motor PI",
                "metrics": [
                    ("Proportional Gain", "kP = 0.323"),
                    ("Integral Gain", "kI = 11.755"),
                    ("Settling Time", "0.14 s"),
                ]
            },
            {
                "title": "Heading PID",
                "metrics": [
                    ("Proportional Gain", "kP = 3.000"),
                    ("Integral Gain", "kI = 2.182"),
                    ("Derivative Gain", "kD = 0.01147"),
                ]
            }
        ]
        summary_bullets = [
            ("PWM Deadzone", "15% minimum duty cycle applied to overcome static friction."),
            ("Stability", "All closed-loop poles lie strictly in the open left-half s-plane."),
        ]
        slide = add_metrics_summary_slide(
            prs=prs,
            title="PI & PID Controller Tuning Parameters",
            cards=cards,
            summary_bullets=summary_bullets,
            banner="Chapter 5. Controller Design",
            theme_name="thesis_blue",
        )
        self.assertIsNotNone(slide)
        out_path = os.path.join(self.dpath, "metrics_test.pptx")
        saved = safe_save_pptx(prs, out_path)
        self.assertTrue(os.path.isfile(saved))

        ast = read_pptx(saved)
        self.assertEqual(ast["total_slides"], 1)

    def test_spec_json_compilation_all_new_types(self):
        """Validates write_pptx_from_spec with split_diagram, dual_diagram, and metrics_summary."""
        spec = {
            "theme": "thesis_blue",
            "slides": [
                {
                    "type": "split_diagram",
                    "title": "Differential Drive Kinematics",
                    "banner": "Chapter 5. Controller Design",
                    "bullets": [
                        {"prefix": "Forward Kinematics", "text": "Linear speed v = r(w_R + w_L)/2"},
                        {"prefix": "Inverse Kinematics", "text": "Right wheel w_R = (2v + bw)/(2r)"},
                    ],
                    "image_path": self.img1,
                    "caption": "Figure 5.12: Kinematics",
                },
                {
                    "type": "dual_diagram",
                    "title": "Cleaning Mechanism Selection",
                    "banner": "Chapter 2. Conceptual Design",
                    "bullets": [
                        "Alternative 1: Direct suction",
                        "Alternative 2: Single roller brush",
                    ],
                    "img1_path": self.img1,
                    "cap1": "Figure 1.8a: Direct Suction",
                    "img2_path": self.img2,
                    "cap2": "Figure 1.8b: Roller Brush",
                },
                {
                    "type": "metrics_summary",
                    "title": "Tuning & Closed-Loop Responses",
                    "cards": [
                        {"title": "Motor A", "metrics": [("kP", "0.301"), ("kI", "12.745")]},
                        {"title": "Motor B", "metrics": [("kP", "0.323"), ("kI", "11.755")]},
                    ],
                    "summary_bullets": [
                        ("Stability Guarantee", "Asymptotically stable via Routh-Hurwitz criterion.")
                    ]
                }
            ]
        }
        out_path = os.path.join(self.dpath, "compiled_spec.pptx")
        saved = write_pptx_from_spec(spec, out_path)
        self.assertTrue(os.path.isfile(saved))

        ast = read_pptx(saved)
        self.assertEqual(ast["total_slides"], 3)

    def test_markdown_auto_categorization(self):
        """Validates Markdown auto-detecting split_diagram and dual_diagram slides."""
        md_text = f"""# Title Slide
---
## Slide With One Image
- **Item 1:** Feature description
- **Item 2:** Second feature
![Architecture Diagram]({self.img1})
---
## Slide With Dual Images
- **Option A:** Direct comparison
- **Option B:** Secondary mode
![Figure Left]({self.img1})
![Figure Right]({self.img2})
"""
        spec = parse_markdown_to_slides_spec(md_text)
        slides = spec["slides"]
        self.assertEqual(len(slides), 3)
        self.assertEqual(slides[0]["type"], "title")
        self.assertEqual(slides[1]["type"], "split_diagram")
        self.assertEqual(slides[2]["type"], "dual_diagram")

    def test_cli_pptx_build_and_preview(self):
        """Validates CLI pptx-build and pptx-preview commands."""
        spec_path = os.path.join(self.dpath, "cli_spec.json")
        out_pptx = os.path.join(self.dpath, "cli_output.pptx")
        preview_dir = os.path.join(self.dpath, "cli_previews")

        spec_data = {
            "theme": "thesis_blue",
            "slides": [
                {
                    "type": "split_diagram",
                    "title": "CLI Test Split Slide",
                    "bullets": [("CLI Bullet", "Testing command line dispatcher.")],
                    "image_path": self.img1,
                    "caption": "Figure: CLI Verification",
                }
            ]
        }
        with open(spec_path, "w", encoding="utf-8") as f:
            json.dump(spec_data, f)

        # 1. Run ai_tools_cli.py pptx-build
        cmd_build = [sys.executable, "ai_tools_cli.py", "pptx-build", spec_path, "-o", out_pptx]
        res_build = subprocess.run(cmd_build, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        self.assertEqual(res_build.returncode, 0, f"pptx-build failed: {res_build.stderr}")
        self.assertTrue(os.path.isfile(out_pptx))

        # 2. Run ai_tools_cli.py pptx-preview
        cmd_preview = [sys.executable, "ai_tools_cli.py", "pptx-preview", out_pptx, "-o", preview_dir, "--slides", "1"]
        res_preview = subprocess.run(cmd_preview, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        self.assertEqual(res_preview.returncode, 0, f"pptx-preview failed: {res_preview.stderr}")
        expected_png = os.path.join(preview_dir, "slide_01_preview.png")
        self.assertTrue(os.path.isfile(expected_png), f"Preview PNG not found at: {expected_png}")


if __name__ == "__main__":
    unittest.main()
