"""
tests/test_pdf_renderer.py — Unit test suite for Universal Multi-Backend PDF Renderer.
"""

from __future__ import annotations

import os
import pytest
from core.converter.pdf_renderer import (
    UniversalPdfRenderer,
    convert_document_to_pdf,
    find_libreoffice_binary,
    is_win32_com_available,
)


def test_universal_pdf_renderer_initialization():
    renderer = UniversalPdfRenderer()
    backends = renderer.get_available_backends()
    assert isinstance(backends, list)
    assert len(backends) >= 1
    assert "fallback" in backends


def test_detect_active_backend():
    renderer = UniversalPdfRenderer()
    active = renderer.detect_active_backend()
    assert active in ["win32com", "libreoffice", "fallback"]


def test_prefer_libreoffice_priority():
    renderer_default = UniversalPdfRenderer(prefer_libreoffice=False)
    renderer_libre = UniversalPdfRenderer(prefer_libreoffice=True)

    backends_libre = renderer_libre.get_available_backends()
    # When prefer_libreoffice is True and libreoffice binary is mock/found, it should precede win32com
    if renderer_libre.libreoffice_binary:
        assert backends_libre[0] == "libreoffice"


def test_convert_to_pdf_nonexistent_source_raises_error():
    renderer = UniversalPdfRenderer()
    with pytest.raises(FileNotFoundError):
        renderer.convert_docx_to_pdf("non_existent_file_12345.docx")

    with pytest.raises(FileNotFoundError):
        renderer.convert_pptx_to_pdf("non_existent_file_12345.pptx")

    with pytest.raises(FileNotFoundError):
        renderer.convert_xlsx_to_pdf("non_existent_file_12345.xlsx")


def test_convert_to_pdf_unsupported_format_raises_error(tmp_path):
    dummy_file = str(tmp_path / "test.unknown_ext")
    with open(dummy_file, "w") as f:
        f.write("content")

    renderer = UniversalPdfRenderer()
    with pytest.raises(ValueError):
        renderer.convert_to_pdf(dummy_file)


def test_convert_pdf_to_pdf_passthrough(tmp_path):
    pdf_in = str(tmp_path / "original.pdf")
    pdf_out = str(tmp_path / "copied.pdf")
    with open(pdf_in, "wb") as f:
        f.write(b"%PDF-1.4 mock content")

    renderer = UniversalPdfRenderer()
    res = renderer.convert_to_pdf(pdf_in, pdf_out)
    assert res == pdf_out
    assert os.path.exists(pdf_out)
    with open(pdf_out, "rb") as f:
        assert f.read() == b"%PDF-1.4 mock content"


def test_convert_xlsx_to_pdf_dispatches_excel_com_without_libreoffice(tmp_path, monkeypatch):
    src = tmp_path / "book.xlsx"
    src.write_bytes(b"placeholder")
    out_pdf = tmp_path / "book.pdf"

    monkeypatch.setattr("sys.platform", "win32")
    monkeypatch.setattr("core.converter.pdf_renderer.find_libreoffice_binary", lambda: None)
    monkeypatch.setattr(
        "core.converter.pdf_renderer.is_win32_com_available",
        lambda app="Word.Application": app == "Excel.Application",
    )

    called = []
    renderer = UniversalPdfRenderer(libreoffice_path=None)
    renderer.libreoffice_binary = None

    def fake_excel_com(xlsx_path, target_pdf):
        called.append((xlsx_path, target_pdf))
        with open(target_pdf, "wb") as f:
            f.write(b"%PDF-1.4 excel com mock")
        return target_pdf

    monkeypatch.setattr(renderer, "_convert_xlsx_win32com", fake_excel_com)
    res = renderer.convert_xlsx_to_pdf(str(src), str(out_pdf))
    assert res == str(out_pdf)
    assert len(called) == 1
    assert out_pdf.exists()


def test_real_excel_com_xlsx_to_pdf_and_xls_roundtrip_if_available(tmp_path):
    import sys
    if sys.platform != "win32" or not is_win32_com_available("Excel.Application"):
        pytest.skip("Microsoft Excel COM unavailable on host")

    import openpyxl
    import pymupdf as fitz

    src_xlsx = str(tmp_path / "com_test.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BaoCao"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 16
    ws["A1"] = "Hang muc"
    ws["B1"] = "Gia tri"
    ws["A2"] = "Doanh thu Q1"
    ws["B2"] = 1500
    ws["A3"] = "Tong"
    ws["B3"] = "=B2*2"
    wb.save(src_xlsx)
    wb.close()

    out_pdf = str(tmp_path / "com_test.pdf")
    renderer = UniversalPdfRenderer()
    res_pdf = renderer.convert_xlsx_to_pdf(src_xlsx, out_pdf)
    assert os.path.isfile(res_pdf)
    doc = fitz.open(res_pdf)
    assert len(doc) >= 1
    assert "Doanh thu Q1" in doc[0].get_text()
    assert "3000" in doc[0].get_text()
    doc.close()


def test_legacy_xls_to_xlsx_and_universal_dispatch_mocked(tmp_path, monkeypatch):
    from core.converter.engine import convert_universal
    import openpyxl

    fake_xls = tmp_path / "legacy.xls"
    fake_xls.write_bytes(b"mock-biff-xls")
    out_xlsx = tmp_path / "upgraded.xlsx"

    def mock_xls_to_xlsx(xls_path, target_xlsx=None, prefer_libreoffice=False):
        dest = target_xlsx or str(tmp_path / "legacy.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["A1"] = "MaHang"
        ws["B1"] = "SoTien"
        ws["A2"] = "SP01"
        ws["B2"] = 5000
        wb.save(dest)
        wb.close()
        return dest

    monkeypatch.setattr("core.converter.pdf_renderer.convert_xls_to_xlsx", mock_xls_to_xlsx)
    ok, res, _elap = convert_universal(str(fake_xls), str(out_xlsx))
    assert ok is True
    assert os.path.isfile(res)
    wb_chk = openpyxl.load_workbook(res)
    assert wb_chk["Sheet1"]["B2"].value == 5000
    wb_chk.close()

    out_md = tmp_path / "upgraded.md"
    ok_md, res_md, _ = convert_universal(str(fake_xls), str(out_md))
    assert ok_md is True
    assert "SP01" in open(res_md, encoding="utf-8").read()
