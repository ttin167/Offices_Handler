"""
core.converter — Universal multi-format conversion dispatchers (PDF ↔ DOCX ↔ MD ↔ PPTX).
"""

from __future__ import annotations

from .engine import convert_file, parse_page_range, get_pdf_page_count
from .markdown import (
    markdown_to_docx,
    markdown_to_pdf,
)
from .pdf_toolkit import (
    merge_pdfs,
    split_pdf,
    rotate_and_clean_pdf,
    stamp_watermark_pdf,
    compress_pdf,
)
from .ocr_pipeline import (
    is_scanned_pdf,
    filter_red_stamp,
    detect_skew_angle,
    deskew_image,
    run_tesseract_ocr,
    find_tesseract_binary,
    VietnameseOcrPipeline,
)
from .pdf_renderer import (
    UniversalPdfRenderer,
    convert_document_to_pdf,
    find_libreoffice_binary,
)

__all__ = [
    "convert_file",
    "parse_page_range",
    "get_pdf_page_count",
    "markdown_to_docx",
    "markdown_to_pdf",
    "merge_pdfs",
    "split_pdf",
    "rotate_and_clean_pdf",
    "stamp_watermark_pdf",
    "compress_pdf",
    "is_scanned_pdf",
    "filter_red_stamp",
    "detect_skew_angle",
    "deskew_image",
    "run_tesseract_ocr",
    "find_tesseract_binary",
    "VietnameseOcrPipeline",
    "UniversalPdfRenderer",
    "convert_document_to_pdf",
    "find_libreoffice_binary",
]
