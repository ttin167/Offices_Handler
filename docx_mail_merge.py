"""
docx_mail_merge.py — Root Facade for Enterprise Batch Mail Merge Studio.

Thin Facade pointing to core.docx.mail_merge for strict Layered Modular Architecture.
"""
from __future__ import annotations

from core.docx.mail_merge import (
    inspect_template_tokens,
    evaluate_condition,
    resolve_conditional_text,
    read_data_source,
    execute_mail_merge,
)

__all__ = [
    "inspect_template_tokens",
    "evaluate_condition",
    "resolve_conditional_text",
    "read_data_source",
    "execute_mail_merge",
]
