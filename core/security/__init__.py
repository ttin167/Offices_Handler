# -*- coding: utf-8 -*-
"""
core/security/ — Enterprise Document Security & Sanitizer Engine.
Implements Cloudflare-seeded defensive security audit standards (ERR_SEC_001–005).
"""

from core.security.document_auditor import (
    DocumentSecurityAuditor,
    SecurityAuditReport,
    SecurityFinding,
    audit_document_security,
)
from core.security.sanitizer import (
    DocumentSanitizer,
    MACRO_TO_STANDARD_CONTENT_TYPES,
    redact_pii_text,
    sanitize_document,
)

__all__ = [
    "DocumentSecurityAuditor",
    "DocumentSanitizer",
    "MACRO_TO_STANDARD_CONTENT_TYPES",
    "SecurityAuditReport",
    "SecurityFinding",
    "audit_document_security",
    "redact_pii_text",
    "sanitize_document",
]
