# -*- coding: utf-8 -*-
"""
core/security/sanitizer.py — Enterprise Document Sanitizer, Deep CDR & Smart PII Redactor (v4.6).

Provides non-destructive safe repairs and Content Disarm & Reconstruction (CDR) for office documents:
  - Defuses Formula / DDE Injection by prepending a single apostrophe (') to formula triggers (ERR_SEC_002).
  - Scrubs PII metadata (Author, LastModifiedBy, revision tracking) from docProps/core.xml (ERR_SEC_004).
  - Restores veryHidden and hidden sheets to normal visibility (ERR_SEC_004).
  - Deep CDR (ERR_SEC_001, ERR_SEC_003, ERR_SEC_006 / Edge Case #1):
      * Strips vbaProject.bin, vbaData.xml, xl/macrosheets/*, and malicious OLE PE ('MZ') binaries.
      * Cleans orphan <Relationship> entries in *.rels and <Override> entries in [Content_Types].xml.
      * Automatically promotes macroEnabled Content-Types to standard OpenXML MIME types so Word/Excel
        opens disarmed .docm/.xlsm/.pptm packages without corruption warnings.
      * Disarms PDF active content (/JavaScript, /JS, /Launch, /OpenAction, /AA, /RichMedia, /EmbeddedFiles).
  - Vietnamese PII Smart Redactor (ERR_SEC_007 / Edge Case #2):
      * Validates 12-digit CCCD province code (001..096) and century/gender digit (0..3).
      * Validates Vietnamese mobile carrier prefixes and excludes currency/percentage numbers.
      * Enforces Contextual Anchoring ('STK', 'Số tài khoản', 'Account No') for bank account numbers
        to prevent false-positive redaction of financial report amounts.
"""

from __future__ import annotations

import csv
import io
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional, Set, Tuple


DANGEROUS_FORMULA_RE = re.compile(
    r"^[=+\-@].*?(cmd|powershell|mshta|cscript|wscript|rundll32|certutil|bitsadmin|"
    r"regsvr32|curl|wget|bash|sh|calc|iex|invoke-|webservice|hyperlink\s*\(.*?://)",
    re.IGNORECASE,
)

# Guard #1 (ERR_SEC_006): Macro-Enabled to Standard OpenXML MIME Promotion Map
MACRO_TO_STANDARD_CONTENT_TYPES: Dict[str, str] = {
    "application/vnd.ms-word.document.macroEnabled.main+xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    ),
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
    ),
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    ),
    "application/vnd.ms-excel.template.macroEnabled.main+xml": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
    ),
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    ),
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml": (
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
    ),
}

# Guard #2 (ERR_SEC_007): Precision Vietnamese PII Regex with Financial False-Positive Protection
CCCD_12_RE = re.compile(
    r"(?<![\d])\b(0(?:0[1-9]|[1-8]\d|9[0-6])[0-3]\d{2})(\d{3})(\d{3})\b"
    r"(?![\d]|\s*(?:(?:VND|VNĐ|USD|EUR|đ|đồng)(?!\w)|%))",
    re.IGNORECASE,
)

VN_PHONE_RE = re.compile(
    r"(?<![\d])((?:\+84|0)(?:3[2-9]|5[2689]|7[06-9]|8[1-9]|9[0-9])\d)(\d{3})(\d{3})\b"
    r"(?![\d]|\s*(?:(?:VND|VNĐ|USD|EUR|đ|đồng)(?!\w)|%))",
    re.IGNORECASE,
)

BANK_ACCOUNT_CONTEXT_RE = re.compile(
    r"(?i)((?:STK|Số\s*tài\s*khoản|Tài\s*khoản\s*số|Tài\s*khoản\s*NH|Account\s*No\.?|Acc\s*No\.?|A/C)"
    r"\s*[:\-]?\s*)(\d{3,4})[\d\s\-]{3,12}(\d{3})\b"
)

MST_CONTEXT_RE = re.compile(
    r"(?i)((?:MST|Mã\s*số\s*thuế|Tax\s*ID)\s*[:\-]?\s*)(\d{3})\d{4}(\d{3}(?:-\d{3})?)\b"
)

EMAIL_PII_RE = re.compile(
    r"\b([a-zA-Z0-9_.+-])([a-zA-Z0-9_.+-]*)(@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)\b"
)


class RedactionCounts(dict):
    """
    Hybrid dict/int counter returned by `redact_pii_text`:
    - Supports `counts["cccd"]`, `counts["phone"]`, `counts["bank_account"]`, `counts["mst"]`, `counts["email"]`, `counts["total"]`
    - Also supports integer comparisons (`counts > 0`, `int(counts)`, `total + counts`).
    """

    def __int__(self) -> int:
        return int(self.get("total", 0))

    def __gt__(self, other: Any) -> bool:
        if isinstance(other, (int, float)):
            return int(self.get("total", 0)) > other
        return False

    def __ge__(self, other: Any) -> bool:
        if isinstance(other, (int, float)):
            return int(self.get("total", 0)) >= other
        return False

    def __lt__(self, other: Any) -> bool:
        if isinstance(other, (int, float)):
            return int(self.get("total", 0)) < other
        return False

    def __eq__(self, other: Any) -> bool:
        if isinstance(other, (int, float)):
            return int(self.get("total", 0)) == other
        return super().__eq__(other)

    def __add__(self, other: Any) -> int:
        if isinstance(other, (int, float)):
            return int(self.get("total", 0)) + int(other)
        if isinstance(other, dict):
            return int(self.get("total", 0)) + int(other.get("total", 0))
        return int(self.get("total", 0))

    def __radd__(self, other: Any) -> int:
        return self.__add__(other)


def redact_pii_text(
    text: str,
    pii_types: Optional[List[str]] = None,
) -> Tuple[str, RedactionCounts]:
    """
    Redacts Vietnamese & international PII in plain text while strictly preserving
    financial numbers, currency amounts, and general multi-digit identifiers (ERR_SEC_007).
    Returns (redacted_text, counts_dict).
    """
    counts = RedactionCounts({"cccd": 0, "phone": 0, "bank_account": 0, "mst": 0, "email": 0, "total": 0})
    if not text:
        return text, counts

    enabled = {t.strip().lower() for t in (pii_types or ["cccd", "phone", "bank_account", "mst", "email"])}
    out = text

    # 1. Contextual Bank Account Redaction (run first so anchored numbers are handled in context)
    if "bank_account" in enabled or "bank" in enabled:
        def _sub_bank(m: re.Match) -> str:
            counts["bank_account"] += 1
            counts["total"] += 1
            return f"{m.group(1)}[REDACTED_BANK_ACC]"
        out = BANK_ACCOUNT_CONTEXT_RE.sub(_sub_bank, out)

    # 2. Contextual Tax ID (MST) Redaction
    if "mst" in enabled or "tax" in enabled:
        def _sub_mst(m: re.Match) -> str:
            counts["mst"] += 1
            counts["total"] += 1
            return f"{m.group(1)}[REDACTED_MST]"
        out = MST_CONTEXT_RE.sub(_sub_mst, out)

    # 3. 12-Digit Vietnamese CCCD (Province 001..096 + Century/Gender 0..3)
    if "cccd" in enabled:
        def _sub_cccd(m: re.Match) -> str:
            counts["cccd"] += 1
            counts["total"] += 1
            return "[REDACTED_CCCD]"
        out = CCCD_12_RE.sub(_sub_cccd, out)

    # 4. Vietnamese Mobile Phone Numbers (excluding currency suffixes)
    if "phone" in enabled:
        def _sub_phone(m: re.Match) -> str:
            counts["phone"] += 1
            counts["total"] += 1
            return "[REDACTED_PHONE]"
        out = VN_PHONE_RE.sub(_sub_phone, out)

    # 5. Email Addresses
    if "email" in enabled:
        def _sub_email(m: re.Match) -> str:
            counts["email"] += 1
            counts["total"] += 1
            return "[REDACTED_EMAIL]"
        out = EMAIL_PII_RE.sub(_sub_email, out)

    return out, counts


def _is_dangerous_ole_or_macro_entry(name: str, data: bytes) -> bool:
    """Checks whether an OpenXML zip member is a macro stream or PE executable OLE object."""
    low = name.lower().replace("\\", "/")
    if low.endswith("vbaproject.bin") or low.endswith("vbadata.xml") or "/macrosheets/" in low:
        return True
    if "embeddings/" in low and (low.endswith(".bin") or low.endswith(".exe") or low.endswith(".scr")):
        # Check for MZ header inside first 512 bytes or embedded PE signature
        if b"MZ" in data[:512] or b"This program cannot be run in DOS mode" in data[:2048]:
            return True
    return False


class DocumentSanitizer:
    """Non-destructive document sanitizer, Deep CDR engine, and Smart PII Redactor."""

    def sanitize(
        self,
        file_path: str,
        output_path: Optional[str] = None,
        defuse_formulas: bool = True,
        strip_metadata: bool = True,
        unhide_sheets: bool = True,
        deep_cdr: bool = True,
        redact_pii: bool = False,
        pii_types: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        abs_path = os.path.abspath(file_path)
        if not os.path.isfile(abs_path):
            return {"success": False, "error": f"File not found: {abs_path}"}

        if not output_path:
            base, ext = os.path.splitext(abs_path)
            output_path = f"{base}_sanitized{ext}"
        else:
            output_path = os.path.abspath(output_path)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        ext = os.path.splitext(abs_path)[1].lower()

        actions_taken: List[str] = []
        metrics: Dict[str, Any] = {
            "defused_formulas": 0,
            "unhidden_sheets": 0,
            "stripped_metadata_fields": 0,
            "removed_macros_and_ole": 0,
            "macros_stripped": 0,
            "ole_objects_stripped": 0,
            "promoted_content_types": 0,
            "orphan_rels_cleaned": 0,
            "disarmed_pdf_objects": 0,
            "redacted_pii_count": 0,
            "output_path": output_path,
        }

        if ext in (".docx", ".xlsx", ".pptx", ".docm", ".xlsm", ".pptm"):
            self._sanitize_openxml(
                abs_path,
                output_path,
                defuse_formulas=defuse_formulas,
                strip_metadata=strip_metadata,
                unhide_sheets=unhide_sheets,
                deep_cdr=deep_cdr,
                redact_pii=redact_pii,
                pii_types=pii_types,
                actions_taken=actions_taken,
                metrics=metrics,
            )
        elif ext == ".csv":
            self._sanitize_csv(
                abs_path,
                output_path,
                defuse_formulas=defuse_formulas,
                redact_pii=redact_pii,
                pii_types=pii_types,
                actions_taken=actions_taken,
                metrics=metrics,
            )
        elif ext == ".pdf":
            self._sanitize_pdf(
                abs_path,
                output_path,
                deep_cdr=deep_cdr,
                strip_metadata=strip_metadata,
                redact_pii=redact_pii,
                pii_types=pii_types,
                actions_taken=actions_taken,
                metrics=metrics,
            )
        else:
            shutil.copy2(abs_path, output_path)
            actions_taken.append("Preserved file as static binary copy.")

        return {
            "success": True,
            "input_file": abs_path,
            "output_file": output_path,
            "output_path": output_path,
            "actions_taken": actions_taken,
            "metrics": metrics,
            "macros_stripped": metrics.get("macros_stripped", 0),
            "ole_objects_stripped": metrics.get("ole_objects_stripped", 0),
            "ole_Needs_stripped": metrics.get("ole_objects_stripped", 0),
            "content_types_promoted": metrics.get("promoted_content_types", 0),
            "orphan_rels_cleaned": metrics.get("orphan_rels_cleaned", 0),
            "redacted_pii_count": metrics.get("redacted_pii_count", 0),
        }

    def _sanitize_openxml(
        self,
        src_path: str,
        dst_path: str,
        defuse_formulas: bool,
        strip_metadata: bool,
        unhide_sheets: bool,
        deep_cdr: bool,
        redact_pii: bool,
        pii_types: Optional[List[str]],
        actions_taken: List[str],
        metrics: Dict[str, Any],
    ) -> None:
        """Process OpenXML zip stream in-memory with Deep CDR and Smart PII Redaction."""
        removed_entries: Set[str] = set()

        with zipfile.ZipFile(src_path, "r") as src_zip:
            # First pass: identify dangerous macro / OLE streams to strip when deep_cdr=True
            if deep_cdr:
                for item in src_zip.infolist():
                    name = item.filename
                    data = src_zip.read(name)
                    if _is_dangerous_ole_or_macro_entry(name, data):
                        norm_entry = name.replace("\\", "/")
                        removed_entries.add(norm_entry)
                        metrics["removed_macros_and_ole"] += 1
                        low_entry = norm_entry.lower()
                        if "vbaproject.bin" in low_entry or "vbadata.xml" in low_entry or "/macrosheets/" in low_entry:
                            metrics["macros_stripped"] = metrics.get("macros_stripped", 0) + 1
                        else:
                            metrics["ole_objects_stripped"] = metrics.get("ole_objects_stripped", 0) + 1
                        actions_taken.append(f"Deep CDR: Stripped active payload '{name}'.")

            with zipfile.ZipFile(dst_path, "w", compression=zipfile.ZIP_DEFLATED) as dst_zip:
                for item in src_zip.infolist():
                    name = item.filename
                    norm_name = name.replace("\\", "/")
                    if norm_name in removed_entries:
                        continue

                    data = src_zip.read(name)

                    # Guard #1 (ERR_SEC_006): Update [Content_Types].xml when disarming macros
                    if deep_cdr and norm_name == "[Content_Types].xml":
                        data, promoted = self._clean_and_promote_content_types(data, removed_entries)
                        if promoted > 0:
                            metrics["promoted_content_types"] += promoted
                            actions_taken.append(
                                f"Deep CDR: Promoted {promoted} macroEnabled Content-Type(s) to Standard OpenXML."
                            )

                    # Clean orphan relationships in *.rels pointing to removed macros/OLE
                    elif deep_cdr and norm_name.endswith(".rels") and removed_entries:
                        data, rels_cleaned = self._clean_rels_xml(data, removed_entries)
                        if rels_cleaned > 0:
                            metrics["orphan_rels_cleaned"] = metrics.get("orphan_rels_cleaned", 0) + rels_cleaned

                    # 1. Defuse Excel formulas & optionally redact PII in sheet*.xml
                    elif norm_name.startswith("xl/worksheets/sheet") and norm_name.endswith(".xml"):
                        if defuse_formulas:
                            data, count = self._defuse_sheet_formulas(data)
                            if count > 0:
                                metrics["defused_formulas"] += count
                                actions_taken.append(f"Defused {count} formula injection(s) in '{name}'.")
                        if redact_pii:
                            data, p_cnt = self._redact_xml_text_nodes(data, pii_types)
                            if p_cnt > 0:
                                metrics["redacted_pii_count"] += p_cnt

                    # Redact PII in sharedStrings.xml or Word document.xml
                    elif redact_pii and (
                        norm_name == "xl/sharedStrings.xml"
                        or (norm_name.startswith("word/") and norm_name.endswith(".xml"))
                    ):
                        data, p_cnt = self._redact_xml_text_nodes(data, pii_types)
                        if p_cnt > 0:
                            metrics["redacted_pii_count"] += p_cnt
                            actions_taken.append(f"Redacted {p_cnt} PII token(s) in '{name}'.")

                    # 2. Unhide sheets in xl/workbook.xml
                    elif unhide_sheets and norm_name == "xl/workbook.xml":
                        data, count = self._unhide_workbook_sheets(data)
                        if count > 0:
                            metrics["unhidden_sheets"] += count
                            actions_taken.append(f"Restored {count} hidden/veryHidden sheet(s) to visible.")

                    # 3. Strip PII from docProps/core.xml
                    elif strip_metadata and norm_name == "docProps/core.xml":
                        data, count = self._strip_core_metadata(data)
                        if count > 0:
                            metrics["stripped_metadata_fields"] += count
                            actions_taken.append(f"Scrubbed {count} PII field(s) from 'docProps/core.xml'.")

                    dst_zip.writestr(item, data)

    def _clean_and_promote_content_types(
        self, xml_bytes: bytes, removed_entries: Set[str]
    ) -> Tuple[bytes, int]:
        """
        Guard #1 (ERR_SEC_006): Removes <Override> tags for stripped macro/OLE parts
        and promotes macroEnabled ContentType values to standard OpenXML MIME types.
        """
        promoted_count = 0
        try:
            root = ET.fromstring(xml_bytes)
            modified = False
            removed_basenames = {os.path.basename(p).lower() for p in removed_entries}

            for elem in list(root):
                part_name = elem.attrib.get("PartName", "").lstrip("/")
                part_low = part_name.lower()
                if (
                    part_name in removed_entries
                    or os.path.basename(part_low) in removed_basenames
                    or "vbaproject.bin" in part_low
                    or "vbadata.xml" in part_low
                    or "/macrosheets/" in part_low
                ):
                    root.remove(elem)
                    modified = True
                    continue

                ctype = elem.attrib.get("ContentType", "")
                if ctype in MACRO_TO_STANDARD_CONTENT_TYPES:
                    elem.set("ContentType", MACRO_TO_STANDARD_CONTENT_TYPES[ctype])
                    promoted_count += 1
                    modified = True

            if modified:
                return ET.tostring(root, encoding="utf-8", xml_declaration=True), promoted_count
        except Exception:
            pass
        return xml_bytes, promoted_count

    def _clean_rels_xml(self, xml_bytes: bytes, removed_entries: Set[str]) -> Tuple[bytes, int]:
        """Removes <Relationship> nodes referencing stripped vbaProject or dangerous OLE binaries."""
        cleaned_count = 0
        try:
            root = ET.fromstring(xml_bytes)
            modified = False
            removed_basenames = {os.path.basename(p).lower() for p in removed_entries}

            for elem in list(root):
                target = elem.attrib.get("Target", "").replace("\\", "/")
                rel_type = elem.attrib.get("Type", "").lower()
                target_base = os.path.basename(target).lower()
                if (
                    target_base in removed_basenames
                    or "vbaproject" in rel_type
                    or "vbadata" in rel_type
                    or "macrosheet" in rel_type
                ):
                    root.remove(elem)
                    cleaned_count += 1
                    modified = True

            if modified:
                return ET.tostring(root, encoding="utf-8", xml_declaration=True), cleaned_count
        except Exception:
            pass
        return xml_bytes, cleaned_count

    def _redact_xml_text_nodes(
        self, xml_bytes: bytes, pii_types: Optional[List[str]]
    ) -> Tuple[bytes, int]:
        """Redacts PII inside OpenXML text elements (<w:t>, <t>, <v>) while preserving formatting runs."""
        count = 0
        try:
            root = ET.fromstring(xml_bytes)
            modified = False
            for elem in root.iter():
                tag_local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if tag_local in ("t", "v") and elem.text:
                    new_text, c = redact_pii_text(elem.text, pii_types=pii_types)
                    if c > 0:
                        elem.text = new_text
                        count += c
                        modified = True
            if modified:
                return ET.tostring(root, encoding="utf-8", xml_declaration=True), count
        except Exception:
            pass
        return xml_bytes, count

    def _defuse_sheet_formulas(self, xml_bytes: bytes) -> Tuple[bytes, int]:
        """Neutralize DDE formula calls inside worksheet XML by escaping them."""
        count = 0
        try:
            root = ET.fromstring(xml_bytes)
            modified = False
            for elem in root.iter():
                if elem.tag.endswith("f") and elem.text:
                    formula = elem.text.strip()
                    norm_f = formula if formula.startswith(("=", "+", "-", "@")) else ("=" + formula)
                    if DANGEROUS_FORMULA_RE.search(norm_f):
                        elem.text = f"'{formula}"
                        count += 1
                        modified = True
            if modified:
                return ET.tostring(root, encoding="utf-8", xml_declaration=True), count
        except Exception:
            pass
        return xml_bytes, count

    def _unhide_workbook_sheets(self, xml_bytes: bytes) -> Tuple[bytes, int]:
        """Convert state='hidden' or state='veryHidden' to visible in workbook.xml."""
        count = 0
        try:
            root = ET.fromstring(xml_bytes)
            modified = False
            for elem in root.iter():
                if elem.tag.endswith("sheet"):
                    state = elem.attrib.get("state", "").lower()
                    if state in ("hidden", "veryhidden"):
                        elem.attrib.pop("state", None)
                        count += 1
                        modified = True
            if modified:
                return ET.tostring(root, encoding="utf-8", xml_declaration=True), count
        except Exception:
            pass
        return xml_bytes, count

    def _strip_core_metadata(self, xml_bytes: bytes) -> Tuple[bytes, int]:
        """Reset creator and lastModifiedBy in docProps/core.xml."""
        count = 0
        try:
            root = ET.fromstring(xml_bytes)
            modified = False
            for elem in root.iter():
                tag = elem.tag.lower()
                if "creator" in tag and elem.text:
                    elem.text = "Antigravity Document Studio"
                    count += 1
                    modified = True
                elif "lastmodifiedby" in tag and elem.text:
                    elem.text = "Antigravity Document Studio"
                    count += 1
                    modified = True
            if modified:
                return ET.tostring(root, encoding="utf-8", xml_declaration=True), count
        except Exception:
            pass
        return xml_bytes, count

    def _sanitize_csv(
        self,
        src_path: str,
        dst_path: str,
        defuse_formulas: bool,
        redact_pii: bool,
        pii_types: Optional[List[str]],
        actions_taken: List[str],
        metrics: Dict[str, Any],
    ) -> None:
        """Escape dangerous CSV cell formulas and optionally redact PII."""
        formula_count = 0
        pii_count = 0
        rows = []
        with open(src_path, "r", encoding="utf-8", errors="ignore") as f:
            reader = csv.reader(f)
            for row in reader:
                sanitized_row = []
                for cell in row:
                    val = cell
                    if defuse_formulas and DANGEROUS_FORMULA_RE.search(val.strip()):
                        val = f"'{val}"
                        formula_count += 1
                    if redact_pii:
                        val, c = redact_pii_text(val, pii_types=pii_types)
                        pii_count += c
                    sanitized_row.append(val)
                rows.append(sanitized_row)

        with open(dst_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(rows)

        if formula_count > 0:
            metrics["defused_formulas"] += formula_count
            actions_taken.append(f"Defused {formula_count} formula injection(s) in CSV.")
        if pii_count > 0:
            metrics["redacted_pii_count"] += pii_count
            actions_taken.append(f"Redacted {pii_count} PII item(s) in CSV.")

    def _sanitize_pdf(
        self,
        src_path: str,
        dst_path: str,
        deep_cdr: bool,
        strip_metadata: bool,
        redact_pii: bool,
        pii_types: Optional[List[str]],
        actions_taken: List[str],
        metrics: Dict[str, Any],
    ) -> None:
        """Disarms PDF active scripts (/JavaScript, /Launch, /OpenAction) and optionally burns PII."""
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz  # type: ignore

        doc = fitz.open(src_path)
        disarmed = 0
        redacted_total = 0

        try:
            if deep_cdr:
                dangerous_keys = ("JavaScript", "JS", "Launch", "OpenAction", "AA", "RichMedia", "EmbeddedFiles")
                for xref in range(1, doc.xref_length()):
                    try:
                        obj_str = doc.xref_object(xref, compressed=False)
                    except Exception:
                        continue
                    for dkey in dangerous_keys:
                        if f"/{dkey}" in obj_str:
                            try:
                                doc.xref_set_key(xref, dkey, "null")
                                disarmed += 1
                            except Exception:
                                pass
                if disarmed > 0:
                    metrics["disarmed_pdf_objects"] += disarmed
                    actions_taken.append(f"Deep CDR: Disarmed {disarmed} active PDF trigger(s) (/JavaScript, /Launch).")

            if redact_pii:
                active_regexes = [CCCD_12_RE, VN_PHONE_RE, BANK_ACCOUNT_CONTEXT_RE, MST_CONTEXT_RE, EMAIL_PII_RE]
                for page in doc:
                    page_text = page.get_text("text") or ""
                    page_hits = 0
                    for rgx in active_regexes:
                        for m in rgx.finditer(page_text):
                            matched_str = m.group(0).strip()
                            if not matched_str:
                                continue
                            quads = page.search_for(matched_str)
                            for q in quads:
                                page.add_redact_annot(q, text="[REDACTED]", fill=(0.1, 0.1, 0.1))
                                page_hits += 1
                    if page_hits > 0:
                        page.apply_redactions()
                        redacted_total += page_hits
                if redacted_total > 0:
                    metrics["redacted_pii_count"] += redacted_total
                    actions_taken.append(f"True Redaction: Permanently burned {redacted_total} PII region(s) in PDF.")

            if strip_metadata:
                doc.set_metadata({})
                try:
                    doc.del_xml_metadata()
                except Exception:
                    pass
                metrics["stripped_metadata_fields"] += 1
                actions_taken.append("Scrubbed PDF Info & XMP metadata.")

            doc.save(dst_path, garbage=4, deflate=True)
        finally:
            doc.close()


def sanitize_document(
    file_path: str,
    output_path: Optional[str] = None,
    defuse_formulas: bool = True,
    strip_metadata: bool = True,
    unhide_sheets: bool = True,
    deep_cdr: bool = True,
    redact_pii: bool = False,
    pii_types: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Public helper function to sanitize document threats with Deep CDR and Smart PII Redaction."""
    sanitizer = DocumentSanitizer()
    return sanitizer.sanitize(
        file_path,
        output_path=output_path,
        defuse_formulas=defuse_formulas,
        strip_metadata=strip_metadata,
        unhide_sheets=unhide_sheets,
        deep_cdr=deep_cdr,
        redact_pii=redact_pii,
        pii_types=pii_types,
    )
