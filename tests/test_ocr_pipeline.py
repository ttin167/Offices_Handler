"""
tests/test_ocr_pipeline.py — Unit test suite for Vietnamese OCR Pipeline & Scan Recovery (ERR_CONV_005).
"""

from __future__ import annotations

import os
import tempfile
import numpy as np
import pytest
from PIL import Image, ImageDraw

from core.converter.ocr_pipeline import (
    is_scanned_pdf,
    filter_red_stamp,
    detect_skew_angle,
    deskew_image,
    run_tesseract_ocr,
    find_tesseract_binary,
    VietnameseOcrPipeline,
)


def test_filter_red_stamp_neutralizes_red_and_preserves_black():
    # Create 100x100 white canvas
    img = Image.new("RGB", (100, 100), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Draw black horizontal bar (representing table line or text)
    draw.rectangle([10, 45, 90, 55], fill=(0, 0, 0))

    # Draw bright red circle (representing official red stamp)
    draw.ellipse([40, 20, 80, 80], fill=(230, 20, 20))

    # Run filter_red_stamp
    cleaned = filter_red_stamp(img)
    arr_cleaned = np.array(cleaned)

    # Point inside black bar but outside red circle (e.g. x=20, y=50): should remain pure black
    assert tuple(arr_cleaned[50, 20]) == (0, 0, 0)

    # Point inside red circle area (e.g. x=60, y=30): should be converted to white background
    assert tuple(arr_cleaned[30, 60]) == (255, 255, 255)


def test_filter_red_stamp_preserves_grayscale_and_blue():
    # Create image with blue, gray, and dark colors
    img = Image.new("RGB", (50, 50), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.rectangle([0, 0, 20, 20], fill=(0, 0, 255))      # Pure blue
    draw.rectangle([25, 25, 45, 45], fill=(128, 128, 128)) # Neutral gray

    cleaned = filter_red_stamp(img)
    arr = np.array(cleaned)
    assert tuple(arr[10, 10]) == (0, 0, 255)
    assert tuple(arr[35, 35]) == (128, 128, 128)


def test_is_scanned_pdf_digital_document(tmp_path):
    try:
        import pymupdf
    except ImportError:
        pytest.skip("PyMuPDF not installed")

    # Create a vector PDF with paragraphs
    pdf_path = str(tmp_path / "digital_doc.pdf")
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nNghị định số 30/2020/NĐ-CP về công tác văn thư."
    page.insert_text((50, 100), text, fontsize=12)
    doc.save(pdf_path)
    doc.close()

    is_scanned, stats = is_scanned_pdf(pdf_path)
    assert is_scanned is False
    assert stats["avg_text_length"] > 20


def test_is_scanned_pdf_raster_scan(tmp_path):
    try:
        import pymupdf
    except ImportError:
        pytest.skip("PyMuPDF not installed")

    # Create a scanned PDF containing a raster image and no digital text
    img_path = str(tmp_path / "scan_page.png")
    img = Image.new("RGB", (600, 800), color=(245, 245, 245))
    draw = ImageDraw.Draw(img)
    draw.rectangle([50, 50, 550, 750], outline=(0, 0, 0), width=2)
    img.save(img_path)

    pdf_path = str(tmp_path / "scanned_doc.pdf")
    doc = pymupdf.open()
    page = doc.new_page(width=600, height=800)
    # Insert image covering 90% of page
    page.insert_image(pymupdf.Rect(30, 40, 570, 760), filename=img_path)
    doc.save(pdf_path)
    doc.close()

    is_scanned, stats = is_scanned_pdf(pdf_path)
    assert is_scanned is True
    assert stats["avg_text_length"] < 20
    assert stats["avg_raster_ratio"] >= 0.70


def test_detect_skew_angle_horizontal_lines():
    img = Image.new("RGB", (400, 400), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    for y in range(50, 350, 30):
        draw.line([(30, y), (370, y)], fill=(0, 0, 0), width=4)

    angle = detect_skew_angle(img)
    # Horizontal lines should have skew angle ~ 0.0
    assert abs(angle) < 1.0


def test_deskew_image_straighten():
    img = Image.new("RGB", (200, 200), color=(255, 255, 255))
    straight, angle = deskew_image(img)
    assert straight.size[0] >= 200
    assert straight.size[1] >= 200


def test_run_tesseract_ocr_graceful_fallback_missing_binary():
    img = Image.new("RGB", (100, 100), color=(255, 255, 255))
    res = run_tesseract_ocr(img, tesseract_cmd="/non/existent/path/to/tesseract")
    assert res["success"] is False
    assert "not found" in res["error"].lower()


def test_vietnamese_ocr_pipeline_document_image(tmp_path):
    img_path = str(tmp_path / "receipt_sample.png")
    img = Image.new("RGB", (300, 200), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    draw.text((20, 30), "HOA DON BAN HANG", fill=(0, 0, 0))
    # Stamp
    draw.ellipse((150, 50, 250, 150), fill=(220, 20, 20))
    img.save(img_path)

    out_md = str(tmp_path / "receipt_ocr.md")
    pipeline = VietnameseOcrPipeline(lang="vie+eng", filter_red_stamp=True)
    res = pipeline.process_document(img_path, output_path=out_md)

    assert res["total_pages"] == 1
    assert os.path.exists(out_md)


def test_ocr_fallback_without_tesseract_and_hybrid_pdf(tmp_path, monkeypatch):
    import sys
    from PIL import ImageFont
    from core.converter.ocr_pipeline import run_ocr_with_fallback

    monkeypatch.setattr("core.converter.ocr_pipeline.find_tesseract_binary", lambda: None)

    img = Image.new("RGB", (600, 120), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 40)
    except Exception:
        font = ImageFont.load_default()
    draw.text((20, 30), "INVOICE 2026 TOTAL 1500", fill=(0, 0, 0), font=font)

    res = run_ocr_with_fallback(img, lang="eng", tesseract_cmd=None, backend="auto")
    if sys.platform == "win32":
        assert res["success"] is True
        assert res["backend"] in ("windows_winrt", "rapidocr", "tesseract")
        assert "2026" in res["text"]
    else:
        assert "backend" in res

    # Hybrid PDF text-layer recovery fallback when raster OCR yields empty output
    import pymupdf

    pdf_path = str(tmp_path / "hybrid_layer.pdf")
    doc = pymupdf.open()
    pg = doc.new_page(width=595, height=842)
    pg.insert_text((72, 72), "Embedded Digital Layer Text 2026", fontsize=12)
    doc.save(pdf_path)
    doc.close()

    monkeypatch.setattr(
        "core.converter.ocr_pipeline.run_ocr_with_fallback",
        lambda *a, **kw: {"success": False, "text": "", "lang_used": "vie+eng", "backend": "none", "error": "disabled"},
    )
    pipeline = VietnameseOcrPipeline(lang="vie+eng", filter_red_stamp=False, deskew=False)
    doc_res = pipeline.process_document(pdf_path)
    assert doc_res["total_pages"] == 1
    assert doc_res["pages"][0]["success"] is True
    assert doc_res["pages"][0]["backend"] == "pdf_text_layer"
    assert "Embedded Digital Layer Text 2026" in doc_res["full_text"]

