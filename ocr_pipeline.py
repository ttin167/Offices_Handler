"""
ocr_pipeline.py — Root Facade for Vietnamese OCR Pipeline & Scan Recovery (v4.0).
Provides dual-fallback import to core/converter/ocr_pipeline.py.
"""

from __future__ import annotations

import sys

try:
    from core.converter.ocr_pipeline import (
        is_scanned_pdf,
        filter_red_stamp,
        detect_skew_angle,
        deskew_image,
        run_tesseract_ocr,
        find_tesseract_binary,
        VietnameseOcrPipeline,
    )
except ImportError:
    import os
    sys.path.insert(0, os.path.dirname(__file__))
    from core.converter.ocr_pipeline import (  # type: ignore
        is_scanned_pdf,
        filter_red_stamp,
        detect_skew_angle,
        deskew_image,
        run_tesseract_ocr,
        find_tesseract_binary,
        VietnameseOcrPipeline,
    )

__all__ = [
    "is_scanned_pdf",
    "filter_red_stamp",
    "detect_skew_angle",
    "deskew_image",
    "run_tesseract_ocr",
    "find_tesseract_binary",
    "VietnameseOcrPipeline",
]
