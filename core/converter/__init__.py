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
    SENSITIVE_PATTERNS,
    compress_pdf,
    convert_images_to_pdf,
    export_pdf_to_images_zip,
    extract_pdf_tables_to_xlsx,
    flatten_and_sanitize_pdf,
    merge_pdfs,
    normalize_and_number_pdf,
    protect_pdf,
    redact_sensitive_pdf,
    render_pdf_thumbnails,
    reorganize_pdf_pages,
    rotate_and_clean_pdf,
    sign_pdf_with_stamp,
    split_pdf,
    stamp_giap_lai_pdf,
    stamp_watermark_pdf,
    crypto_sign_pdf,
    verify_pdf_crypto_signature,
    convert_to_pdfa,
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
    "SENSITIVE_PATTERNS",
    "merge_pdfs",
    "split_pdf",
    "rotate_and_clean_pdf",
    "stamp_watermark_pdf",
    "compress_pdf",
    "redact_sensitive_pdf",
    "protect_pdf",
    "normalize_and_number_pdf",
    "stamp_giap_lai_pdf",
    "extract_pdf_tables_to_xlsx",
    "render_pdf_thumbnails",
    "reorganize_pdf_pages",
    "convert_images_to_pdf",
    "export_pdf_to_images_zip",
    "sign_pdf_with_stamp",
    "flatten_and_sanitize_pdf",
    "crypto_sign_pdf",
    "verify_pdf_crypto_signature",
    "convert_to_pdfa",
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
