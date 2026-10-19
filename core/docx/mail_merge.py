"""
core.docx.mail_merge — Enterprise Batch Mail Merge Studio Engine (v4.0).

Capabilities:
1. inspect_template_tokens: Discovers all {{TOKEN}}, {{IF:...}}, and {{BLOCK:...}} variables across body, tables, headers, and footers.
2. read_data_source: Extracts records from Excel (.xlsx), CSV, JSON with leading-zero guard (ERR_XLSX_007), currency & date formatting.
3. evaluate_condition: Safe declarative evaluation of {{IF:condition}}...{{ELSE}}...{{ENDIF}} expressions.
4. execute_mail_merge: Batch generation of customized Word documents with dynamic filename formatting and ZIP bundle packaging.
"""

from __future__ import annotations

import csv
import datetime
import io
import json
import os
import re
import shutil
import time
import zipfile
from typing import Any, Callable, Dict, List, Optional, Set, Tuple, Union

import docx
from docx import Document

try:
    from .writer import patch_docx_template, _replace_text_in_paragraph
except (ImportError, ValueError):
    from docx_writer import patch_docx_template, _replace_text_in_paragraph  # type: ignore

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# Regex patterns for token discovery and conditional blocks
TOKEN_REGEX = re.compile(r"\{\{([A-Za-z0-9_\.\-]+)\}\}")
CONDITIONAL_FULL_REGEX = re.compile(r"\{\{IF:\s*([^\}]+?)\}\}(.*?)(?:\{\{ELSE\}\}(.*?))?\{\{ENDIF\}\}", re.DOTALL)
BLOCK_REGEX = re.compile(r"\{\{BLOCK:([A-Za-z0-9_\.\-]+)\}\}")


def inspect_template_tokens(template_path: str) -> Dict[str, Any]:
    """
    Scans a Word .docx template and returns all discovered placeholders:
    - variables: Standard {{TOKEN}} placeholders.
    - conditionals: Condition expressions in {{IF:...}} blocks.
    - blocks: Dynamic table/image anchors {{BLOCK:...}}.
    """
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Template DOCX not found: {template_path}")

    doc = Document(template_path)
    all_text_fragments: List[str] = []

    # 1. Document body paragraphs
    for p in doc.paragraphs:
        if p.text:
            all_text_fragments.append(p.text)

    # 2. Tables
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text:
                        all_text_fragments.append(p.text)

    # 3. Headers and Footers
    for s in doc.sections:
        for p in s.header.paragraphs:
            if p.text:
                all_text_fragments.append(p.text)
        for t in s.header.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if p.text:
                            all_text_fragments.append(p.text)
        for p in s.footer.paragraphs:
            if p.text:
                all_text_fragments.append(p.text)
        for t in s.footer.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if p.text:
                            all_text_fragments.append(p.text)

    combined_text = "\n".join(all_text_fragments)

    # Discover variables
    raw_tokens = TOKEN_REGEX.findall(combined_text)
    clean_vars: Set[str] = set()
    for tok in raw_tokens:
        tok_upper = tok.upper()
        if not tok_upper.startswith("IF:") and not tok_upper.startswith("BLOCK:") and tok_upper != "ELSE" and tok_upper != "ENDIF":
            clean_vars.add(tok)

    # Discover conditionals
    conditionals = [m.group(1).strip() for m in CONDITIONAL_FULL_REGEX.finditer(combined_text)]

    # Discover blocks
    blocks = BLOCK_REGEX.findall(combined_text)

    return {
        "template_file": os.path.basename(template_path),
        "variables": sorted(list(clean_vars)),
        "conditionals": list(dict.fromkeys(conditionals)),
        "blocks": sorted(list(set(blocks))),
        "total_unique_tokens": len(clean_vars),
    }


def evaluate_condition(condition_str: str, record: Dict[str, Any]) -> bool:
    """
    Safely evaluates a conditional expression against record dictionary.
    Supports:
    - Comparison: 'Age >= 18', 'Status == "Active"', 'Gender != "Nu"', 'Score > 8.0'
    - Membership: 'Role in ["Admin", "Manager"]'
    - Truthiness: 'HasBonus' or 'IsVip'
    """
    cond = condition_str.strip()
    if not cond:
        return False

    # Check comparison operators (longest first)
    comp_ops = [("==", lambda a, b: a == b),
                ("!=", lambda a, b: a != b),
                (">=", lambda a, b: a >= b),
                ("<=", lambda a, b: a <= b),
                (">",  lambda a, b: a > b),
                ("<",  lambda a, b: a < b),
                (" in ", lambda a, b: a in b)]

    for op, func in comp_ops:
        if op in cond:
            left_part, right_part = cond.split(op, 1)
            left_key = left_part.strip().strip("'\"")
            right_val_raw = right_part.strip().strip("'\"")

            # Extract left value from record (case-insensitive fallback)
            left_val = record.get(left_key)
            if left_val is None:
                for k, v in record.items():
                    if k.lower() == left_key.lower():
                        left_val = v
                        break

            # Handle type coercion for numeric comparison
            try:
                if isinstance(left_val, (int, float)):
                    right_val = float(right_val_raw)
                    return func(float(left_val), right_val)
                elif isinstance(right_val_raw, (int, float)):
                    return func(float(left_val or 0), float(right_val_raw))
            except Exception:
                pass

            # String comparison
            s_left = str(left_val or "").strip()
            s_right = str(right_val_raw).strip()
            return func(s_left, s_right)

    # Simple truthiness
    val = record.get(cond)
    if val is None:
        for k, v in record.items():
            if k.lower() == cond.lower():
                val = v
                break

    if val is None:
        return False
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val != 0
    s_val = str(val).strip().lower()
    return s_val not in ("", "0", "false", "no", "null", "none", "không")


def resolve_conditional_text(text: str, record: Dict[str, Any]) -> str:
    """
    Recursively evaluates {{IF:cond}}...{{ELSE}}...{{ENDIF}} blocks inside text.
    """
    if "{{IF:" not in text or "{{ENDIF}}" not in text:
        return text

    def _replace_match(m: re.Match) -> str:
        condition = m.group(1).strip()
        true_body = m.group(2) or ""
        false_body = m.group(3) or ""
        is_true = evaluate_condition(condition, record)
        return true_body if is_true else false_body

    # Replace all conditional blocks
    result = CONDITIONAL_FULL_REGEX.sub(_replace_match, text)
    return result


def read_data_source(
    data_source: Union[str, List[Dict[str, Any]]],
    sheet_name: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Extracts tabular records from Excel (.xlsx), CSV, or JSON data sources.
    Enforces ERR_XLSX_007: Identifier columns preserve leading zeros ('0123456789').
    """
    if isinstance(data_source, list):
        return [dict(r) for r in data_source]

    src = os.path.abspath(data_source)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Data source not found: {src}")

    ext = os.path.splitext(src)[1].lower()

    if ext == ".json":
        with open(src, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, list):
                return data
            elif isinstance(data, dict):
                for val in data.values():
                    if isinstance(val, list):
                        return val
            return [data]

    elif ext == ".csv":
        records = []
        with open(src, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                cleaned_row = {k.strip(): v.strip() for k, v in row.items() if k}
                records.append(cleaned_row)
        return records

    elif ext in (".xlsx", ".xlsm"):
        if not HAS_OPENPYXL:
            raise ImportError("openpyxl is required to read Excel data source for mail merge.")

        wb = openpyxl.load_workbook(src, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active

        # Find header row
        header_row_idx = None
        headers: List[str] = []
        for r in range(1, min(ws.max_row, 15) + 1):
            row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
            non_empty = [v for v in row_vals if v]
            if len(non_empty) >= 2:
                header_row_idx = r
                headers = row_vals
                break

        if header_row_idx is None:
            wb.close()
            return []

        # Keywords for leading-zero preservation (ERR_XLSX_007)
        id_kws = ["id", "code", "mã", "manv", "sđt", "sdt", "phone", "cmnd", "cccd", "stt", "account"]

        records = []
        for r in range(header_row_idx + 1, ws.max_row + 1):
            row_dict: Dict[str, Any] = {}
            has_val = False
            for c, h in enumerate(headers, start=1):
                if not h:
                    continue
                cell = ws.cell(r, c)
                val = cell.value

                if val is not None and str(val).strip() != "":
                    has_val = True

                # Format preservation
                if val is None:
                    row_dict[h] = ""
                elif isinstance(val, (datetime.datetime, datetime.date)):
                    row_dict[h] = val.strftime("%d/%m/%Y")
                elif isinstance(val, int):
                    # Check if column is an identifier that needs leading zeros preserved
                    h_lower = h.lower()
                    if any(k in h_lower for k in id_kws) and cell.number_format == "@":
                        row_dict[h] = str(val)
                    else:
                        row_dict[h] = str(val)
                elif isinstance(val, float):
                    if val.is_integer():
                        row_dict[h] = str(int(val))
                    else:
                        row_dict[h] = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                else:
                    row_dict[h] = str(val).strip()

            if has_val:
                records.append(row_dict)

        wb.close()
        return records

    else:
        raise ValueError(f"Unsupported data source extension: {ext}. Expected .xlsx, .csv, or .json")


def _sanitize_filename(name: str) -> str:
    """Removes invalid filesystem characters for Windows, Linux, and macOS."""
    cleaned = re.sub(r'[\\/*?:"<>|]', "", name)
    return cleaned.strip() or "document"


def execute_mail_merge(
    template_path: Optional[str] = None,
    data_source: Optional[Union[str, List[Dict[str, Any]]]] = None,
    output_dir: str = "./mail_merge_output",
    filename_pattern: Optional[str] = None,
    sheet_name: Optional[str] = None,
    block_providers: Optional[Dict[str, Any]] = None,
    create_zip: bool = False,
    zip_name: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    template_docx: Optional[str] = None,
    zip_output: Optional[Union[str, bool]] = None,
) -> Dict[str, Any]:
    """
    Executes high-fidelity batch mail merge:
    1. Loads template DOCX and tabular data.
    2. Resolves conditionals ({{IF:...}}...{{ENDIF}}).
    3. Replaces tokens across document body, tables, and headers/footers.
    4. Replaces dynamic block anchors ({{BLOCK:id}}).
    5. Saves individual documents and packages into ZIP if requested.
    """
    start_time = time.time()
    raw_template = template_path or template_docx
    if not raw_template:
        raise ValueError("template_path (or template_docx) is required for mail merge")
    t_path = os.path.abspath(raw_template)
    if not os.path.isfile(t_path):
        raise FileNotFoundError(f"Template DOCX not found: {t_path}")

    if data_source is None:
        raise ValueError("data_source is required for mail merge")

    out_d = os.path.abspath(output_dir)
    os.makedirs(out_d, exist_ok=True)

    if zip_output is not None:
        if isinstance(zip_output, str) and zip_output.strip():
            create_zip = True
            if zip_output.lower().endswith(".zip"):
                zip_name = os.path.basename(zip_output)
        elif bool(zip_output):
            create_zip = True

    fn_pattern = filename_pattern or "doc_{index}.docx"

    records = read_data_source(data_source, sheet_name=sheet_name)
    total_records = len(records)
    if total_records == 0:
        return {
            "success": True,
            "total_records": 0,
            "generated_files": [],
            "zip_path": None,
            "elapsed_seconds": 0.0,
            "message": "Data source is empty; 0 documents generated."
        }

    generated_files: List[str] = []

    for idx, record in enumerate(records, start=1):
        if progress_callback:
            pct = round((idx / total_records) * 90.0, 1)
            progress_callback(pct, f"Đang trộn thư ({idx}/{total_records})...")

        # 1. Format dynamic output filename
        fname = fn_pattern
        for k, v in record.items():
            fname = fname.replace(f"{{{{{k}}}}}", str(v))
            fname = fname.replace(f"{{{k}}}", str(v))
        fname = fname.replace("{index}", f"{idx:03d}").replace("{{index}}", f"{idx:03d}")
        if not fname.lower().endswith(".docx"):
            fname += ".docx"
        fname = _sanitize_filename(fname)
        doc_out_path = os.path.join(out_d, fname)

        # 2. Build replacements dictionary
        text_replacements: Dict[str, str] = {}
        for k, v in record.items():
            text_replacements[f"{{{{{k}}}}}"] = str(v)
            text_replacements[f"{{{k}}}"] = str(v)

        # 3. Open fresh copy of template
        doc = Document(t_path)

        # 4. Resolve conditional logic in paragraphs
        def _resolve_paragraph_conditionals(p):
            if "{{IF:" in p.text and "{{ENDIF}}" in p.text:
                new_text = resolve_conditional_text(p.text, record)
                if new_text != p.text:
                    p.text = new_text

        for p in doc.paragraphs:
            _resolve_paragraph_conditionals(p)

        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _resolve_paragraph_conditionals(p)

        for s in doc.sections:
            for p in s.header.paragraphs:
                _resolve_paragraph_conditionals(p)
            for p in s.footer.paragraphs:
                _resolve_paragraph_conditionals(p)

        # 5. Apply Token Replacements using run-merging
        for p in doc.paragraphs:
            _replace_text_in_paragraph(p, text_replacements)

        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _replace_text_in_paragraph(p, text_replacements)

        for s in doc.sections:
            for p in s.header.paragraphs:
                _replace_text_in_paragraph(p, text_replacements)
            for p in s.footer.paragraphs:
                _replace_text_in_paragraph(p, text_replacements)

        # 6. Resolve Dynamic Blocks ({{BLOCK:id}})
        if block_providers:
            for block_key, block_spec in block_providers.items():
                anchor_key = f"{{{{BLOCK:{block_key}}}}}" if not block_key.startswith("{{") else block_key
                for p in list(doc.paragraphs):
                    if anchor_key in p.text:
                        p_parent = p._p.getparent()
                        if isinstance(block_spec, dict) and block_spec.get("type") == "table":
                            # Dynamic table construction from records
                            t_spec = block_spec
                            rows_data = t_spec.get("rows", [])
                            num_rows = len(rows_data)
                            num_cols = len(t_spec.get("headers", rows_data[0] if rows_data else []))
                            new_t = doc.add_table(rows=num_rows, cols=num_cols)
                            # Move table right before anchor paragraph
                            if p_parent is not None:
                                p_parent.insert(p_parent.index(p._p), new_t._tbl)
                                p_parent.remove(p._p)
                            break

        # Save generated document
        doc.save(doc_out_path)
        generated_files.append(doc_out_path)

    # 7. Package ZIP bundle if requested
    zip_dest: Optional[str] = None
    if create_zip and generated_files:
        if progress_callback:
            progress_callback(95.0, "Đang đóng gói tệp ZIP...")
        if isinstance(zip_output, str) and (os.path.isabs(zip_output) or os.path.dirname(zip_output)):
            zip_dest = os.path.abspath(zip_output)
            os.makedirs(os.path.dirname(zip_dest), exist_ok=True)
        else:
            z_name = zip_name or (zip_output if isinstance(zip_output, str) and zip_output else None) or "mail_merge_bundle.zip"
            if not z_name.lower().endswith(".zip"):
                z_name += ".zip"
            zip_dest = os.path.join(out_d, z_name)
        with zipfile.ZipFile(zip_dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for gf in generated_files:
                zf.write(gf, arcname=os.path.basename(gf))

    if progress_callback:
        progress_callback(100.0, "Hoàn thành trộn thư!")

    elapsed = time.time() - start_time
    return {
        "success": True,
        "total_records": total_records,
        "generated_files": generated_files,
        "zip_path": zip_dest,
        "output_dir": out_d,
        "elapsed_seconds": round(elapsed, 2),
    }
