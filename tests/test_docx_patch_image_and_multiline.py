# -*- coding: utf-8 -*-
"""
tests/test_docx_patch_image_and_multiline.py — Tests for Rich-Media Image & Multi-line Patching in DOCX Quick Patch.
"""

import os
import tempfile
from PIL import Image
import docx
from docx import Document
from core.docx.writer import patch_docx_template


def test_docx_patch_multiline_and_images():
    with tempfile.TemporaryDirectory() as tmp_dir:
        tpl_path = os.path.join(tmp_dir, "template.docx")
        out_path = os.path.join(tmp_dir, "patched.docx")
        img_path = os.path.join(tmp_dir, "logo.png")

        # 1. Create a dummy test image
        img = Image.new("RGB", (200, 100), color=(73, 109, 137))
        img.save(img_path)

        # 2. Create template DOCX
        doc = Document()
        p1 = doc.add_paragraph("Thong tin: {{ADDRESS}}")
        p2 = doc.add_paragraph("Logo cong ty: {{IMAGE:LOGO}}")
        table = doc.add_table(rows=1, cols=2)
        table.rows[0].cells[0].text = "Chu ky:"
        table.rows[0].cells[1].text = "{{IMAGE:SIGNATURE}}"
        doc.save(tpl_path)

        # 3. Apply patch with multi-line string and image replacements
        replacements = {
            "{{ADDRESS}}": "So 123 Duong Le Loi\nQuan 1, TP. Ho Chi Minh\nViet Nam",
            "{{IMAGE:LOGO}}": {"type": "image", "path": img_path, "width_in": 2.0},
            "{{IMAGE:SIGNATURE}}": img_path,
        }

        actual_out = patch_docx_template(tpl_path, out_path, replacements=replacements)
        assert os.path.isfile(actual_out)

        # 4. Verify patched document structure
        doc_res = Document(actual_out)

        # Check multi-line: paragraph xml should contain <w:br/>
        p_xml = doc_res.paragraphs[0]._p.xml
        assert "<w:br" in p_xml
        assert "Quan 1" in doc_res.paragraphs[0].text

        # Check inline images
        # There should be exactly 2 inline shapes (one in paragraph 2, one in table cell)
        assert len(doc_res.inline_shapes) == 2
