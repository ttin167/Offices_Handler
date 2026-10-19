# -*- coding: utf-8 -*-
"""
test_unified_qa_diagnostic.py — Unit tests for Enterprise Unified QA Diagnostic Engine.
Validates PPTX, DOCX, XLSX, Diagram and Scan image diagnostic & repair capabilities.
"""

import os
import tempfile
import xml.etree.ElementTree as ET
import pytest
from PIL import Image

import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn as pptx_qn
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from tools.unified_qa_diagnostic import (
    DocxDiagnosticEngine,
    PptxDiagnosticEngine,
    DiagramAndConversionEngine,
    ExcelDiagnosticEngine,
    run_pipeline,
)


def test_pptx_diagnostic_and_repair(tmp_path):
    """Verifies PPTX aspect ratio, word wrap, and defRPr repair."""
    prs = Presentation()
    # Set to 4:3 (10 x 7.5 in) to trigger PPTX_INV_01
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Add text shape without word wrap (PPTX_INV_04)
    tx_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
    tf = tx_box.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = "This is a very long sample text without word wrap enabled for QA test."

    # Add illegal bold attribute to <a:defRPr> to trigger PPTX_INV_09
    p_xml = p._p
    pPr = p_xml.get_or_add_pPr()
    defRPr = pPr.get_or_add_defRPr()
    defRPr.attrib["b"] = "1"

    test_pptx = os.path.join(tmp_path, "test_qa.pptx")
    prs.save(test_pptx)

    # 1. Diagnose
    issues, metrics = PptxDiagnosticEngine.diagnose(test_pptx)
    codes = [i.code for i in issues]
    assert "PPTX_INV_01" in codes  # Non-widescreen 4:3
    assert "PPTX_INV_04" in codes  # Missing word_wrap
    assert "PPTX_INV_09" in codes  # Formatting on <a:defRPr>
    assert metrics["widescreen_16_9"] is False

    # 2. Repair
    repaired_pptx = os.path.join(tmp_path, "test_qa_repaired.pptx")
    success, actions = PptxDiagnosticEngine.repair(test_pptx, repaired_pptx)
    assert success is True
    assert any("Enabled word_wrap" in act for act in actions)
    assert any("Pruned illegal formatting from <a:defRPr>" in act for act in actions)

    # 3. Verify repair
    issues_after, _ = PptxDiagnosticEngine.diagnose(repaired_pptx)
    codes_after = [i.code for i in issues_after]
    assert "PPTX_INV_04" not in codes_after
    assert "PPTX_INV_09" not in codes_after


def test_docx_diagnostic_and_repair(tmp_path):
    """Verifies DOCX table cell last paragraph rule (ERR_DOCX_001) and repair."""
    from docx.oxml.ns import qn
    doc = Document()
    doc.add_heading("QA Test Document", level=1)
    tbl = doc.add_table(rows=12, cols=2)
    tbl.cell(0, 0).text = "Header 1"
    tbl.cell(0, 1).text = "Header 2"
    tbl.cell(1, 0).text = "Data 1"

    # Intentionally strip <w:p> from cell(1, 1) to trigger ERR_DOCX_001
    bad_tc = tbl.cell(1, 1)._tc
    for p_elem in list(bad_tc.findall(qn("w:p"))):
        bad_tc.remove(p_elem)
    # Also remove cantSplit from rows
    for row in tbl.rows:
        trPr = row._tr.find(parse_xml(f"<w:trPr {nsdecls('w')}/>").tag)
        if trPr is not None:
            cant = trPr.find(parse_xml(f"<w:cantSplit {nsdecls('w')}/>").tag)
            if cant is not None:
                trPr.remove(cant)

    test_docx = os.path.join(tmp_path, "test_qa.docx")
    doc.save(test_docx)

    # 1. Diagnose
    issues, metrics = DocxDiagnosticEngine.diagnose(test_docx)
    codes = [i.code for i in issues]
    assert "ERR_DOCX_001" in codes  # missing <w:p> in table cell
    assert "ERR_DOCX_003" in codes  # missing cantSplit

    # 2. Repair
    repaired_docx = os.path.join(tmp_path, "test_qa_repaired.docx")
    success, actions = DocxDiagnosticEngine.repair(test_docx, repaired_docx)
    assert success is True
    assert any("cantSplit" in act for act in actions)

    # 3. Verify repair
    repaired_doc = Document(repaired_docx)
    assert len(repaired_doc.tables[0].rows[0]._tr.xpath(".//w:cantSplit")) > 0


def test_drawio_diagram_diagnostic(tmp_path):
    """Verifies Draw.io MX_INV_01, MX_INV_02, and MX_INV_08 quality gates."""
    # 1. Test MX_INV_01: UserObject wrapping raw PlantUML / Mermaid blob
    bad_drawio_content = """<mxfile host="Electron">
      <diagram name="Page-1">
        <mxGraphModel>
          <root>
            <mxCell id="0"/>
            <mxCell id="1" parent="0"/>
            <UserObject plantUmlData="skinparam monochrome true" id="2">
              <mxCell style="shape=rect" vertex="1" parent="1"/>
            </UserObject>
            <mxCell id="3" edge="1" parent="1" source="2" target="2" value="Calls API" style="edgeStyle=orthogonalEdgeStyle;"/>
          </root>
        </mxGraphModel>
      </diagram>
    </mxfile>"""
    bad_file = os.path.join(tmp_path, "bad.drawio")
    with open(bad_file, "w", encoding="utf-8") as f:
        f.write(bad_drawio_content)

    issues, metrics = DiagramAndConversionEngine.diagnose_diagram_file(bad_file)
    codes = [i.code for i in issues]
    assert "MX_INV_01" in codes
    assert "MX_INV_08" in codes  # Edge with label lacks labelBackgroundColor

    # 2. Test MX_INV_02: Corrupted XML
    corrupt_file = os.path.join(tmp_path, "corrupt.drawio")
    with open(corrupt_file, "w", encoding="utf-8") as f:
        f.write("<mxfile><diagram><unclosed>")

    issues_c, _ = DiagramAndConversionEngine.diagnose_diagram_file(corrupt_file)
    codes_c = [i.code for i in issues_c]
    assert "MX_INV_02" in codes_c


def test_scan_image_diagnostic(tmp_path):
    """Verifies ERR_CONV_005 red stamp saturation detection."""
    # Create an image with a solid red stamp rectangle (> 5% red pixels)
    img = Image.new("RGB", (200, 200), color=(255, 255, 255))
    for x in range(50, 150):
        for y in range(50, 150):
            img.putpixel((x, y), (220, 20, 20))  # High saturation red

    img_path = os.path.join(tmp_path, "test_stamp.png")
    img.save(img_path)

    issues, metrics = DiagramAndConversionEngine.diagnose_scan_image(img_path)
    codes = [i.code for i in issues]
    assert "ERR_CONV_005" in codes
    assert metrics["has_red_stamp"] is True


def test_unified_qa_pipeline_execution(tmp_path):
    """Verifies unified QA CLI pipeline runs with JSON output and returns proper code."""
    doc = Document()
    doc.add_paragraph("Clean paragraph")
    docx_path = os.path.join(tmp_path, "clean.docx")
    doc.save(docx_path)

    out_json = os.path.join(tmp_path, "qa_report.json")
    code = run_pipeline(["--docx", docx_path, "--format", "json", "--output", out_json])
    assert code in (0, 1)
    assert os.path.isfile(out_json)
