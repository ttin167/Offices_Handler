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
