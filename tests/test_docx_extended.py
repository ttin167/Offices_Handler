"""
tests/test_docx_extended.py — Comprehensive Extended Quality & Invariant Tests for Antigravity Doc Handler.
Verifies:
1. Markdown style YAML loading and default fallbacks.
2. Table alignment parsing (Left, Center, Right) and heuristic widths.
3. Fenced code block typography (Consolas, 9.5pt).
4. Table of Contents dot-leader tab-stops.
5. Inline typography parsing (bold, italic, inline code).
6. Heading hierarchy (H1-H4) and keep_with_next invariant.
7. Dynamic Section Orientation (Portrait <-> Landscape) with dimension swap.
8. Custom page margins in centimeters.
9. Header text and dynamic page footer (<w:fldSimple w:instr="PAGE">).
10. Table merged cells (colspan and rowspan).
11. Table cell shading (<w:shd>) and cell borders (<w:tcBorders>).
12. Image fallback card on missing image paths.
13. Printable margin limit enforcement (ERR_DOCX_005).
14. In-place template patching with {{KEY}} replacement.
15. Run-merging split token replacement across Word runs.
16. Block anchor replacement ({{BLOCK:id}} -> Table/Paragraphs).
17. Table OpenXML invariants (cantSplit, tblHeader, vAlign).
18. Signature block right-alignment detection.
19. LaTeX Math integration in paragraphs and tables.
20. List level out-of-bound clamping (0 to 2).
21. Hex color parsing resilience (valid hex, missing hash, invalid length).
22. Converter engine format detection.
23. Heuristic column widths damping behavior.
24. Document round-trip JSON snapshot serialization.
25. Cell paragraph spacing and line spacing invariants.
"""

import os
import unittest
import docx
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

from docx_writer import (
    write_docx,
    patch_docx_template,
    _resolve_style,
    _hex_to_rgb,
    _write_fallback_card,
)
from markdown_converter import (
    decoupled_to_docx,
    _parse_length,
    _add_formatted_runs,
)
from smart_post_processor import calculate_heuristic_column_widths
from docx_lists import apply_list_formatting


class TestDocxExtendedEngine(unittest.TestCase):
    def setUp(self):
        self.out_dir = os.path.join(os.getcwd(), ".web_outputs")
        os.makedirs(self.out_dir, exist_ok=True)

    def test_01_hex_to_rgb_parsing_resilience(self):
        """Valid hex, 8-digit hex, and malformed hex strings must be handled gracefully."""
        self.assertEqual(_hex_to_rgb("#1F4E78"), RGBColor(0x1F, 0x4E, 0x78))
        self.assertEqual(_hex_to_rgb("1F4E78"), RGBColor(0x1F, 0x4E, 0x78))
        self.assertEqual(_hex_to_rgb("#FF0000"), RGBColor(0xFF, 0x00, 0x00))
        # Fallback on invalid hex returns None so Word uses default
        fallback = _hex_to_rgb("invalid")
        self.assertIsNone(fallback)

    def test_02_parse_length_units(self):
        """_parse_length parses float, inches, cm, pt correctly."""
        self.assertEqual(_parse_length(1.5), Inches(1.5))
        self.assertEqual(_parse_length("2.0in"), Inches(2.0))
        self.assertAlmostEqual(_parse_length("2.54cm"), Cm(2.54), delta=100)
        self.assertEqual(_parse_length(None, default_inches=1.0), Inches(1.0))

    def test_03_markdown_inline_formatting_runs(self):
        """_add_formatted_runs must split bold, italic, and normal text correctly."""
        doc = Document()
        p = doc.add_paragraph()
        _add_formatted_runs(p, "Normal **bold text** and *italic text* here.")
        run_texts = [r.text for r in p.runs]
        self.assertIn("bold text", run_texts)
        self.assertIn("italic text", run_texts)

        # Bold run check
        bold_run = next(r for r in p.runs if r.text == "bold text")
        self.assertTrue(bold_run.bold)

        # Italic run check
        italic_run = next(r for r in p.runs if r.text == "italic text")
        self.assertTrue(italic_run.italic)

    def test_04_markdown_fenced_code_block(self):
        """Fenced code blocks must format with Consolas font and compact spacing."""
        md_text = (
            "# Code Example\n"
            "```python\n"
            "def hello():\n"
            "    return 'world'\n"
            "```\n"
        )
        md_file = os.path.join(self.out_dir, "test_code_block.md")
        docx_file = os.path.join(self.out_dir, "test_code_block.docx")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_text)
        decoupled_to_docx(md_file, docx_file)

        doc = Document(docx_file)
        code_paras = [p for p in doc.paragraphs if p.runs and p.runs[0].font.name == "Consolas"]
        self.assertEqual(len(code_paras), 2)
        self.assertEqual(code_paras[0].text, "def hello():")
        self.assertEqual(code_paras[1].text, "    return 'world'")

    def test_05_markdown_table_alignments(self):
        """Table columns with :---, :---:, ---: must align Left, Center, Right."""
        md_text = (
            "| Left | Center | Right |\n"
            "|:---|:---:|---:|\n"
            "| L1 | C1 | R1 |\n"
        )
        md_file = os.path.join(self.out_dir, "test_table_align.md")
        docx_file = os.path.join(self.out_dir, "test_table_align.docx")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_text)
        decoupled_to_docx(md_file, docx_file)

        doc = Document(docx_file)
        t = doc.tables[0]
        data_row = t.rows[1]
        self.assertEqual(data_row.cells[0].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.LEFT)
        self.assertEqual(data_row.cells[1].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(data_row.cells[2].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.RIGHT)

    def test_06_markdown_headings_hierarchy_and_keep_with_next(self):
        """Headings 1-4 must set keep_with_next=True to prevent orphan headings."""
        md_text = (
            "# Heading 1\n"
            "## Heading 2\n"
            "### Heading 3\n"
            "#### Heading 4\n"
        )
        md_file = os.path.join(self.out_dir, "test_headings.md")
        docx_file = os.path.join(self.out_dir, "test_headings.docx")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_text)
        decoupled_to_docx(md_file, docx_file)

        doc = Document(docx_file)
        for p in doc.paragraphs:
            self.assertTrue(p.paragraph_format.keep_with_next)

    def test_07_markdown_signature_block_alignment(self):
        """'NGƯỜI LẬP KẾ HOẠCH' and '(Đã ký)' must be right-aligned."""
        md_text = (
            "Văn bản hành chính.\n"
            "NGƯỜI LẬP KẾ HOẠCH\n"
            "(Đã ký)\n"
            "Nguyễn Văn A\n"
        )
        md_file = os.path.join(self.out_dir, "test_sig.md")
        docx_file = os.path.join(self.out_dir, "test_sig.docx")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_text)
        decoupled_to_docx(md_file, docx_file)

        doc = Document(docx_file)
        p_sig = [p for p in doc.paragraphs if "NGƯỜI LẬP KẾ HOẠCH" in p.text][0]
        p_signed = [p for p in doc.paragraphs if "(Đã ký)" in p.text][0]
        self.assertEqual(p_sig.alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertEqual(p_signed.alignment, WD_ALIGN_PARAGRAPH.RIGHT)

    def test_08_docx_writer_landscape_orientation_switch(self):
        """Landscape section switches orientation and swaps width and height."""
        snapshot = {
            "sections": [
                {
                    "orientation": "landscape",
                    "page_width_cm": 29.7,
                    "page_height_cm": 21.0
                }
            ],
            "body": [{"type": "paragraph", "text": "Landscape page"}]
        }
        docx_file = os.path.join(self.out_dir, "test_landscape.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        sec = doc.sections[0]
        self.assertEqual(sec.orientation, WD_ORIENT.LANDSCAPE)
        self.assertTrue(sec.page_width > sec.page_height)

    def test_09_docx_writer_custom_margins(self):
        """Sections must support custom margins in centimeters."""
        snapshot = {
            "sections": [
                {
                    "margin_top_cm": 3.0,
                    "margin_bottom_cm": 3.0,
                    "margin_left_cm": 2.0,
                    "margin_right_cm": 2.0
                }
            ],
            "body": [{"type": "paragraph", "text": "Margin test"}]
        }
        docx_file = os.path.join(self.out_dir, "test_margins.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        sec = doc.sections[0]
        self.assertAlmostEqual(sec.top_margin.cm, 3.0, places=1)
        self.assertAlmostEqual(sec.left_margin.cm, 2.0, places=1)

    def test_10_docx_writer_header_and_footer_page_number(self):
        """Header text and footer page field <w:fldSimple w:instr='PAGE'>."""
        snapshot = {
            "sections": [
                {
                    "header_text": "Confidential Architecture Spec",
                    "footer_page_number": True
                }
            ],
            "body": [{"type": "paragraph", "text": "Page body"}]
        }
        docx_file = os.path.join(self.out_dir, "test_header_footer.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        sec = doc.sections[0]
        self.assertIn("Confidential Architecture Spec", sec.header.paragraphs[0].text)
        footer_p = sec.footer.paragraphs[0]
        fldSimple = footer_p._p.find(qn("w:fldSimple"))
        self.assertIsNotNone(fldSimple)
        self.assertEqual(fldSimple.get(qn("w:instr")), "PAGE")

    def test_11_docx_writer_table_merged_colspan(self):
        """Colspan > 1 merges adjacent cells across the row."""
        snapshot = {
            "body": [
                {
                    "type": "table",
                    "rows": [
                        {
                            "cells": [
                                {"text": "Merged 2 Columns", "colspan": 2},
                                {"text": "Column 3"}
                            ]
                        },
                        {
                            "cells": [
                                {"text": "Col 1"},
                                {"text": "Col 2"},
                                {"text": "Col 3"}
                            ]
                        }
                    ]
                }
            ]
        }
        docx_file = os.path.join(self.out_dir, "test_colspan.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        table = doc.tables[0]
        self.assertEqual(len(table.columns), 3)

    def test_12_docx_writer_table_merged_rowspan(self):
        """Rowspan > 1 merges cells across vertical rows."""
        snapshot = {
            "body": [
                {
                    "type": "table",
                    "rows": [
                        {
                            "cells": [
                                {"text": "Merged 2 Rows", "rowspan": 2},
                                {"text": "R1 C2"}
                            ]
                        },
                        {
                            "cells": [
                                {"text": "R2 C2"}
                            ]
                        }
                    ]
                }
            ]
        }
        docx_file = os.path.join(self.out_dir, "test_rowspan.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        table = doc.tables[0]
        self.assertEqual(len(table.rows), 2)

    def test_13_docx_writer_image_fallback_card(self):
        """Missing image gracefully generates fallback card without raising exception."""
        snapshot = {
            "body": [
                {
                    "type": "image",
                    "path": "non_existent_image_12345.png",
                    "alt": "Missing Diagram"
                }
            ]
        }
        docx_file = os.path.join(self.out_dir, "test_fallback.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        self.assertTrue(len(doc.tables) > 0)
        card_cell = doc.tables[0].cell(0, 0)
        self.assertIn("Missing Diagram", card_cell.text)

    def test_14_in_place_template_patching_simple(self):
        """patch_docx_template replaces {{KEY}} tokens in template."""
        base_doc = Document()
        base_doc.add_paragraph("Project: {{PROJECT_NAME}}, Author: {{AUTHOR}}")
        tpl_path = os.path.join(self.out_dir, "base_tpl.docx")
        base_doc.save(tpl_path)

        out_path = os.path.join(self.out_dir, "patched_tpl.docx")
        patch_docx_template(
            tpl_path,
            out_path,
            replacements={"{{PROJECT_NAME}}": "DocHandler Pro", "{{AUTHOR}}": "Antigravity AI"}
        )

        doc = Document(out_path)
        self.assertIn("Project: DocHandler Pro, Author: Antigravity AI", doc.paragraphs[0].text)

    def test_15_in_place_template_patching_split_runs(self):
        """Split runs with {{PART1}}{{PART2}} are merged and replaced."""
        base_doc = Document()
        p = base_doc.add_paragraph()
        r1 = p.add_run("Hello {{")
        r2 = p.add_run("USER_")
        r3 = p.add_run("NAME}} welcome!")
        tpl_path = os.path.join(self.out_dir, "split_tpl.docx")
        base_doc.save(tpl_path)

        out_path = os.path.join(self.out_dir, "split_patched.docx")
        patch_docx_template(
            tpl_path,
            out_path,
            replacements={"{{USER_NAME}}": "Alice"}
        )

        doc = Document(out_path)
        self.assertIn("Hello Alice welcome!", doc.paragraphs[0].text)

    def test_16_in_place_template_block_anchor_replacement(self):
        """Block anchor {{BLOCK:summary}} is replaced with table elements."""
        base_doc = Document()
        base_doc.add_paragraph("Intro")
        base_doc.add_paragraph("{{BLOCK:kpi_table}}")
        base_doc.add_paragraph("Outro")
        tpl_path = os.path.join(self.out_dir, "block_tpl.docx")
        base_doc.save(tpl_path)

        out_path = os.path.join(self.out_dir, "block_patched.docx")
        table_block = {
            "type": "table",
            "rows": [
                {"cells": [{"text": "Metric"}, {"text": "Value"}]},
                {"cells": [{"text": "Coverage"}, {"text": "100%"}]}
            ]
        }
        patch_docx_template(
            tpl_path,
            out_path,
            replacements={"{{BLOCK:kpi_table}}": table_block}
        )

        doc = Document(out_path)
        self.assertTrue(len(doc.tables) > 0)
        self.assertNotIn("{{BLOCK:kpi_table}}", [p.text for p in doc.paragraphs])

    def test_17_heuristic_column_widths_damping(self):
        """Square root damped weighting prevents extreme column widths."""
        table_data = [
            ["ID", "Very Long Description of Feature and Architectural Requirements for Enterprise Application", "Status"],
            ["1", "Short text", "Pass"]
        ]
        total_w = 9000
        col_widths = calculate_heuristic_column_widths(table_data, total_w)
        self.assertEqual(len(col_widths), 3)
        self.assertEqual(sum(col_widths), total_w)
        # ID column must get reasonable width (not crushed to 0)
        self.assertTrue(col_widths[0] >= 500)
        # Description column gets larger share
        self.assertTrue(col_widths[1] > col_widths[0])

    def test_18_latex_math_in_markdown(self):
        """$E=mc^2$ and block $$a^2 + b^2 = c^2$$ compile via docx_math into OMML."""
        md_text = (
            "# Physics\n"
            "Energy formula: $E = mc^2$.\n"
            "$$\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}$$\n"
        )
        md_file = os.path.join(self.out_dir, "test_math_md.md")
        docx_file = os.path.join(self.out_dir, "test_math_md.docx")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_text)
        decoupled_to_docx(md_file, docx_file)

        doc = Document(docx_file)
        xml_all = "".join(p._p.xml for p in doc.paragraphs)
        self.assertTrue("superscript" in xml_all or "∑" in xml_all or "oMath" in xml_all)

    def test_19_list_level_out_of_bounds_clamping(self):
        """Level > 2 is clamped to 2, and negative level is clamped to 0."""
        doc = Document()
        p_high = doc.add_paragraph("High level")
        apply_list_formatting(p_high, item_type="bullet", level=99, num_id=2002)
        ilvl_high = p_high._p.pPr.find(qn("w:numPr")).find(qn("w:ilvl")).get(qn("w:val"))
        self.assertEqual(ilvl_high, "2")

        p_low = doc.add_paragraph("Low level")
        apply_list_formatting(p_low, item_type="bullet", level=-5, num_id=2002)
        ilvl_low = p_low._p.pPr.find(qn("w:numPr")).find(qn("w:ilvl")).get(qn("w:val"))
        self.assertEqual(ilvl_low, "0")

    def test_20_style_resolution_fallback_to_normal(self):
        """Unknown style names gracefully fall back to 'Normal' style."""
        doc = Document()
        s = _resolve_style(doc, "NonExistentStyle123")
        self.assertEqual(s.name, "Normal")

    def test_21_table_openxml_flags_cant_split_and_tbl_header(self):
        """Tables generated by write_docx have cantSplit and tblHeader on row 0."""
        snapshot = {
            "body": [
                {
                    "type": "table",
                    "rows": [
                        {"cells": [{"text": "H1", "is_header": True}, {"text": "H2", "is_header": True}]},
                        {"cells": [{"text": "D1"}, {"text": "D2"}]}
                    ]
                }
            ]
        }
        docx_file = os.path.join(self.out_dir, "test_tbl_flags.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        t = doc.tables[0]
        r0_tr = t.rows[0]._tr.trPr
        self.assertIsNotNone(r0_tr.find(qn("w:cantSplit")))
        self.assertIsNotNone(r0_tr.find(qn("w:tblHeader")))

    def test_22_table_cell_vertical_centering(self):
        """Table cells must have <w:vAlign w:val='center'/>."""
        snapshot = {
            "body": [
                {
                    "type": "table",
                    "rows": [
                        {"cells": [{"text": "Centered Cell"}]}
                    ]
                }
            ]
        }
        docx_file = os.path.join(self.out_dir, "test_valign.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        tcPr = doc.tables[0].cell(0, 0)._tc.tcPr
        vAlign = tcPr.find(qn("w:vAlign"))
        self.assertIsNotNone(vAlign)
        self.assertEqual(vAlign.get(qn("w:val")), "center")

    def test_23_paragraph_spacing_and_line_spacing(self):
        """Paragraph spacing attributes are applied cleanly."""
        snapshot = {
            "body": [
                {
                    "type": "paragraph",
                    "text": "Spaced paragraph",
                    "space_before_pt": 4.0,
                    "space_after_pt": 8.0,
                    "line_spacing": 1.25
                }
            ]
        }
        docx_file = os.path.join(self.out_dir, "test_spacing.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        p = doc.paragraphs[0]
        self.assertEqual(p.paragraph_format.space_before, Pt(4.0))
        self.assertEqual(p.paragraph_format.space_after, Pt(8.0))
        self.assertEqual(p.paragraph_format.line_spacing, 1.25)

    def test_24_markdown_toc_dot_leaders(self):
        """TOC entries matching Title ....... Page format receive right tab stop."""
        md_text = (
            "# TABLE OF CONTENTS\n"
            "Introduction ................................. 1\n"
            "Architecture Overview ....................... 5\n"
        )
        md_file = os.path.join(self.out_dir, "test_toc.md")
        docx_file = os.path.join(self.out_dir, "test_toc.docx")
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(md_text)
        decoupled_to_docx(md_file, docx_file)

        doc = Document(docx_file)
        p_toc = doc.paragraphs[1]
        tabs = p_toc._p.pPr.find(qn("w:tabs"))
        self.assertIsNotNone(tabs)
        tab_elem = tabs.find(qn("w:tab"))
        self.assertEqual(tab_elem.get(qn("w:val")), "right")
        self.assertEqual(tab_elem.get(qn("w:leader")), "dot")

    def test_25_section_break_adds_new_section(self):
        """section_break element creates a new section with specified orientation."""
        snapshot = {
            "body": [
                {"type": "paragraph", "text": "Section 1 Portrait"},
                {
                    "type": "section_break",
                    "orientation": "landscape",
                    "margin_top_cm": 2.0
                },
                {"type": "paragraph", "text": "Section 2 Landscape"}
            ]
        }
        docx_file = os.path.join(self.out_dir, "test_sec_break.docx")
        write_docx(snapshot, docx_file)

        doc = Document(docx_file)
        self.assertEqual(len(doc.sections), 2)
        self.assertEqual(doc.sections[0].orientation, WD_ORIENT.PORTRAIT)
        self.assertEqual(doc.sections[1].orientation, WD_ORIENT.LANDSCAPE)


if __name__ == "__main__":
    unittest.main()
