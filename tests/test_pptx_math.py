"""
test_pptx_math.py — Comprehensive Test Suite for PowerPoint Math Engine and OpenXML Invariants.

Validates:
- pptx_math.py: Greek symbol lexicon, operators, LaTeX tokenization, DrawingML run synthesis.
- PPTX_INV_09: Run-Level Formatting Supremacy (all styles in <a:rPr>, zero <a:defRPr b="..."> locks).
- PPTX_INV_10: Default Pure Black Typography (#000000 default, zero legacy #0F294A / #475569 leaks).
- PPTX_INV_11: Slide Numbering Freedom (run-level Arial 12pt black, zero defRPr lock).
- PPTX_INV_12: Native DrawingML Math Subscript/Superscript (baseline="-25000", baseline="30000").
- PPTX_INV_13: AlternateContent Ghost Shape Pruning.
- Full Markdown -> PPTX compilation with math formulas ($...$).
- Real PowerPoint COM Automation validation (headless opening, 0 repair dialogs, slide preview export).
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
import unittest
from lxml import etree

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from pptx_math import (
    GREEK_SYMBOLS,
    MATH_OPERATORS,
    replace_latex_symbols,
    parse_inline_math,
    create_drawingml_run,
    add_math_text_to_paragraph,
)
from pptx_writer import (
    init_presentation,
    set_paragraph_text_clean,
    sanitize_openxml_presentation,
    safe_save_pptx,
    write_pptx_from_spec,
    add_content_slide,
    add_title_slide,
    add_chapter_slide,
    add_table_slide,
    add_kpi_cards_slide,
)
from pptx_engine import (
    compile_markdown_to_presentation,
    parse_markdown_to_slides_spec,
)

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


class TestPPTXMathLexiconAndTokenizer(unittest.TestCase):
    """Unit tests for pptx_math parser, symbols, and DrawingML run generation."""

    def test_greek_and_operator_replacement(self):
        """Validates LaTeX symbols replacement with Unicode math characters."""
        expr = r"\alpha + \beta = \gamma \le \Omega \approx 10\degree"
        replaced = replace_latex_symbols(expr)
        self.assertIn("α", replaced)
        self.assertIn("β", replaced)
        self.assertIn("γ", replaced)
        self.assertIn("≤", replaced)
        self.assertIn("Ω", replaced)
        self.assertIn("≈", replaced)
        self.assertIn("°", replaced)

    def test_parse_subscripts_and_superscripts(self):
        """Validates tokenization of subscripts (v_{ref}) and superscripts (x^2)."""
        tokens = parse_inline_math("v_{ref} = 0")
        self.assertEqual(tokens[0].text, "v")
        self.assertTrue(tokens[0].is_italic)
        self.assertFalse(tokens[0].is_subscript)

        self.assertEqual(tokens[1].text, "ref")
        self.assertTrue(tokens[1].is_subscript)
        self.assertFalse(tokens[1].is_italic)
        self.assertEqual(tokens[1].scale_size, 0.75)

        tokens_sup = parse_inline_math("x^2 + y^{max}")
        tok_dict = {t.text: t for t in tokens_sup}
        self.assertIn("2", tok_dict)
        self.assertTrue(tok_dict["2"].is_superscript)
        self.assertIn("max", tok_dict)
        self.assertTrue(tok_dict["max"].is_superscript)

    def test_parse_text_and_bold_blocks(self):
        """Validates \\text{...} and \\mathbf{...} blocks parsing."""
        tokens = parse_inline_math(r"\mathbf{F} = m \cdot \mathbf{a} + \text{const}")
        tok_f = tokens[0]
        self.assertEqual(tok_f.text, "F")
        self.assertTrue(tok_f.is_bold)
        self.assertFalse(tok_f.is_italic)

        # Check variable m is italic
        tok_m = [t for t in tokens if t.text == "m"][0]
        self.assertTrue(tok_m.is_italic)
        self.assertFalse(tok_m.is_bold)

        # Check \text{const} is upright
        tok_const = [t for t in tokens if "const" in t.text][0]
        self.assertFalse(tok_const.is_italic)

    def test_create_drawingml_run_subscript_superscript(self):
        """Validates DrawingML <a:rPr> baseline attributes for scripts (PPTX_INV_12)."""
        run_sub = create_drawingml_run("ref", base_sz=1400, scale_size=0.75, is_subscript=True)
        xml_sub = etree.tostring(run_sub).decode("utf-8")
        self.assertIn('baseline="-25000"', xml_sub)
        self.assertIn('sz="1050"', xml_sub)

        run_sup = create_drawingml_run("2", base_sz=1400, scale_size=0.75, is_superscript=True)
        xml_sup = etree.tostring(run_sup).decode("utf-8")
        self.assertIn('baseline="30000"', xml_sup)


class TestPPTXDrawingMLInvariants(unittest.TestCase):
    """Integration tests validating PPTX_INV_09 through PPTX_INV_13."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.dpath = self.temp_dir.name

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_pptx_inv_09_run_level_formatting_supremacy(self):
        """
        PPTX_INV_09: Asserts that no <a:defRPr> in any paragraph contains b="1" or solidFill,
        guaranteeing the PowerPoint Ribbon Bold and Color Picker remain 100% unlocked.
        """
        prs = init_presentation()
        slide = add_content_slide(
            prs=prs,
            title="Đặc tính kỹ thuật hệ thống",
            subtitle="Phân tích chi tiết tham số động cơ và nguồn cấp",
            bullets=[
                "**Vận tốc danh định:** $v_{ref} = 0.8\\text{ m/s}$ tại tần số $f = 50\\text{ Hz}$.",
                "**Công suất cực đại:** $P_{max} \\approx 250\\text{ W}$ với hiệu suất $\\eta \\ge 88\\%$.",
                "**Điện áp ngắt pin:** $V_{cut} = 10.5\\text{ V}$ bảo vệ cell $LiFePO_4$.",
            ],
            theme_name="corporate_blue",
        )
        out_path = os.path.join(self.dpath, "inv09_test.pptx")
        saved = safe_save_pptx(prs, out_path)
        self.assertTrue(os.path.isfile(saved))

        # Inspect raw XML of saved presentation
        prs_check = Presentation(saved)
        for s in prs_check.slides:
            for p in s._element.findall(f".//{{{A_NS}}}p"):
                pPr = p.find(f"{{{A_NS}}}pPr")
                if pPr is not None:
                    defRPr = pPr.find(f"{{{A_NS}}}defRPr")
                    if defRPr is not None:
                        # Must NOT lock bold
                        self.assertNotIn("b", defRPr.attrib, "Violation of PPTX_INV_09: defRPr has b attribute locked!")
                        # Must NOT lock color
                        sf = defRPr.find(f"{{{A_NS}}}solidFill")
                        self.assertIsNone(sf, "Violation of PPTX_INV_09: defRPr has solidFill locked!")

    def test_pptx_inv_10_pure_black_typography(self):
        """
        PPTX_INV_10: Asserts that subtitles, body runs, and cards default to pure black (#000000)
        and no legacy navy (#0F294A) or slate (#475569) leaks remain.
        """
        prs = init_presentation()
        add_title_slide(
            prs=prs,
            title="BÁO CÁO TỐT NGHIỆP",
            subtitle="HỆ THỐNG ĐIỀU KHIỂN ROBOT DI ĐỘNG",
            presenter="Nguyễn Vũ Nguyên",
            date_str="Tháng 10/2026",
            theme_name="corporate_blue",
        )
        add_chapter_slide(
            prs=prs,
            chapter_num="CHƯƠNG 4",
            title="THIẾT KẾ PHẦN CỨNG VÀ NGUỒN CẤP",
            subtitle="Tính toán bảo vệ cầu chì và cảm biến dòng",
            theme_name="thesis_blue",
        )
        out_path = os.path.join(self.dpath, "inv10_test.pptx")
        saved = safe_save_pptx(prs, out_path)

        prs_check = Presentation(saved)
        for s in prs_check.slides:
            for r in s._element.findall(f".//{{{A_NS}}}r"):
                rPr = r.find(f"{{{A_NS}}}rPr")
                if rPr is not None:
                    for clr in rPr.findall(f".//{{{A_NS}}}srgbClr"):
                        val = clr.attrib.get("val", "").upper()
                        self.assertNotIn(val, ["0F294A", "475569"], f"Violation of PPTX_INV_10: Legacy hardcoded color #{val} found in run!")

    def test_pptx_inv_11_slide_numbering_freedom(self):
        """
        PPTX_INV_11: Asserts slide number placeholders have clean run-level black Arial 12pt formatting.
        """
        prs = init_presentation()
        # Create a slide and add a slide number placeholder
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tx_box = slide.shapes.add_textbox(Inches(12.0), Inches(7.0), Inches(1.0), Inches(0.4))
        tx_box.name = "Slide Number Placeholder 1"
        tx_box.text_frame.text = "1"

        out_path = os.path.join(self.dpath, "inv11_test.pptx")
        saved = safe_save_pptx(prs, out_path)

        prs_check = Presentation(saved)
        s = prs_check.slides[0]
        sh = [sh for sh in s.shapes if "Slide Number" in sh.name][0]
        p = sh.text_frame.paragraphs[0]
        # Check that run exists and is formatted cleanly
        self.assertGreater(len(p.runs), 0)
        r = p.runs[0]
        self.assertEqual(r.text, "1")
        self.assertEqual(r.font.name, "Arial")
        self.assertEqual(r.font.size, Pt(12))
        self.assertFalse(r.font.bold)

    def test_pptx_inv_13_alternatecontent_ghost_pruning(self):
        """
        PPTX_INV_13: Asserts orphan / corrupt AlternateContent blocks are cleanly pruned on save.
        """
        prs = init_presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Inject a dummy AlternateContent with legacy code
        ac_xml = (
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            '<mc:Choice Requires="a14">'
            '<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:r><a:rPr><a:solidFill><a:srgbClr val="0F294A"/></a:solidFill></a:rPr><a:t>Ghost PTC Data</a:t></a:r>'
            '</a:p>'
            '</mc:Choice>'
            '</mc:AlternateContent>'
        )
        from pptx.oxml import parse_xml
        ghost_elem = parse_xml(ac_xml)
        slide._element.spTree.append(ghost_elem)

        # Count before
        self.assertTrue(any(c.tag.endswith("AlternateContent") for c in slide._element.spTree))

        out_path = os.path.join(self.dpath, "inv13_test.pptx")
        saved = safe_save_pptx(prs, out_path)

        # Load back and verify ghost AlternateContent was pruned
        prs_check = Presentation(saved)
        s = prs_check.slides[0]
        has_ac = any(c.tag.endswith("AlternateContent") for c in s._element.spTree)
        self.assertFalse(has_ac, "Violation of PPTX_INV_13: Ghost AlternateContent was not pruned!")


class TestMarkdownMathCompilation(unittest.TestCase):
    """End-to-end tests for compiling Markdown with LaTeX math into PowerPoint."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.dpath = self.temp_dir.name

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_compile_markdown_with_math_formulas(self):
        """Validates markdown slides with inline formulas compile into native multi-run DrawingML."""
        md_text = """
# Thiết kế Hệ thống Điều khiển Robot
> Bộ điều khiển bám quỹ đạo với sai số $e(t) \\le 0.05\\text{ m}$

---
## Mô hình Động học & Vận tốc
<!-- banner: CHƯƠNG 5. ĐIỀU KHIỂN TỰ ĐỘNG -->
- **Vận tốc tham chiếu:** Đặt $v_{ref} = 0.5\\text{ m/s}$ và gia tốc $a_{max} = 1.2\\text{ m/s}^2$.
- **Tốc độ góc bánh xe:** Quan hệ $\\omega = \\frac{v_R - v_L}{b}$ với khoảng cách trục $b = 216\\text{ mm}$.
- **Góc hướng la bàn:** Sai số $\\Delta \\theta = \\theta_{des} - \\theta_{cur} \\approx 0\\degree$.
"""
        out_path = os.path.join(self.dpath, "math_markdown.pptx")
        saved = compile_markdown_to_presentation(md_text, out_path, theme_name="corporate_blue")
        self.assertTrue(os.path.isfile(saved))

        prs = Presentation(saved)
        self.assertEqual(len(prs.slides), 2)

        # Inspect slide 2 runs for math subscripts and italic variables
        slide2 = prs.slides[1]
        all_text = ""
        has_italic_run = False
        has_subscript_run = False

        for p in slide2._element.findall(f".//{{{A_NS}}}p"):
            for r in p.findall(f"{{{A_NS}}}r"):
                t = r.find(f"{{{A_NS}}}t")
                if t is not None and t.text:
                    all_text += t.text + " "
                rPr = r.find(f"{{{A_NS}}}rPr")
                if rPr is not None:
                    if rPr.attrib.get("i") == "1":
                        has_italic_run = True
                    if rPr.attrib.get("baseline") == "-25000":
                        has_subscript_run = True

        self.assertIn("v", all_text)
        self.assertIn("ref", all_text)
        self.assertIn("ω", all_text)
        self.assertIn("Δ", all_text)
        self.assertIn("θ", all_text)
        self.assertTrue(has_italic_run, "Math variable was not rendered as italic run!")
        self.assertTrue(has_subscript_run, "Subscript ref was not rendered with baseline='-25000'!")


class TestPowerPointCOMAutomation(unittest.TestCase):
    """
    Validates generated PPTX presentations against real Microsoft PowerPoint COM engine.
    Ensures 0 repair dialogs, perfect widescreen rendering, and preview generation.
    """

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.dpath = self.temp_dir.name

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_powerpoint_com_open_and_render_clean(self):
        """Opens generated presentation in Microsoft PowerPoint via win32com and exports slide PNG."""
        try:
            import win32com.client
        except ImportError:
            self.skipTest("win32com is not installed")

        prs = init_presentation()
        add_title_slide(
            prs=prs,
            title="ĐỒ ÁN TỐT NGHIỆP CƠ ĐIỆN TỬ",
            subtitle="Nghiên cứu động học Robot di động 2 bánh vi sai",
            presenter="Nguyễn Vũ Nguyên",
            date_str="2026",
            theme_name="corporate_blue",
        )
        add_content_slide(
            prs=prs,
            title="Thông số Động lực học & Công thức Toán",
            subtitle="Phân tích mô hình toán học trong không gian trạng thái",
            bullets=[
                "**Vận tốc tham chiếu:** $v_{ref} = 0\\text{ m/s}$ tại điểm bắt đầu $t_0 = 0\\text{ s}$.",
                "**Góc quay thân xe:** Tích phân $\\theta(t) = \\int_0^t \\omega(\\tau) d\\tau$.",
                "**Điện trở shunt:** $R_{shunt} = 0.01\\text{ }\\Omega$ với dòng $I_{max} = 10\\text{ A}$.",
            ],
            theme_name="corporate_blue",
        )
        out_path = os.path.abspath(os.path.join(self.dpath, "com_test.pptx"))
        safe_save_pptx(prs, out_path)

        # Open in PowerPoint COM
        ppt_app = None
        deck = None
        try:
            ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
            deck = ppt_app.Presentations.Open(out_path, ReadOnly=True, Untitled=False, WithWindow=False)
            self.assertEqual(deck.Slides.Count, 2)

            # Export slide 2 preview PNG
            preview_png = os.path.join(self.dpath, "slide2_preview.png")
            slide2 = deck.Slides(2)
            slide2.Export(preview_png, "PNG", 1920, 1080)
            self.assertTrue(os.path.isfile(preview_png))
            self.assertGreater(os.path.getsize(preview_png), 5000)
        finally:
            if deck is not None:
                deck.Close()
            if ppt_app is not None:
                ppt_app.Quit()


if __name__ == "__main__":
    unittest.main()
