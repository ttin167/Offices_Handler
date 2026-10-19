"""
docx_diff_engine.py — Root Facade for Enterprise Legal Redline Diff Engine (v4.0).
Provides dual-fallback import to core/docx/diff_engine.py.
"""

from __future__ import annotations

import sys

try:
    from core.docx.diff_engine import (
        RedlineDiffEngine,
        TrackChangeIdGenerator,
        tokenize_text,
        diff_token_sequences,
        create_ins_element,
        create_del_element,
        create_equal_run,
        compare_docx_documents,
    )
except ImportError:
    # Local fallback if core is not in sys.path
    import os
    sys.path.insert(0, os.path.dirname(__file__))
    from core.docx.diff_engine import (  # type: ignore
        RedlineDiffEngine,
        TrackChangeIdGenerator,
        tokenize_text,
        diff_token_sequences,
        create_ins_element,
        create_del_element,
        create_equal_run,
        compare_docx_documents,
    )

__all__ = [
    "RedlineDiffEngine",
    "TrackChangeIdGenerator",
    "tokenize_text",
    "diff_token_sequences",
    "create_ins_element",
    "create_del_element",
    "create_equal_run",
    "compare_docx_documents",
]
