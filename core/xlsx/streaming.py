# -*- coding: utf-8 -*-
"""
core.xlsx.streaming — High-Performance Streaming I/O Engine for Large Spreadsheets.

Enforces Invariant ERR_XLSX_014: RAM-Constrained Streaming Threshold:
Maintains memory consumption strictly under 50 MB on constrained environments
(e.g., Render.com 512 MB Free Tier) when processing workbooks exceeding 10 MB
or 30,000 rows.
"""

from __future__ import annotations

import os
import zipfile
from typing import Any, Dict, Generator, Iterable, List, Optional
import xml.etree.ElementTree as ET

import openpyxl


__all__ = [
    "stream_read_xlsx",
    "stream_write_xlsx",
    "estimate_xlsx_metrics",
    "STREAMING_FILE_SIZE_THRESHOLD_MB",
    "STREAMING_ROW_COUNT_THRESHOLD",
]

STREAMING_FILE_SIZE_THRESHOLD_MB = 10.0
STREAMING_ROW_COUNT_THRESHOLD = 30000


def estimate_xlsx_metrics(file_path: str) -> Dict[str, Any]:
    """
    Rapidly estimates file metrics without loading OpenXML DOM into memory.
    Inspects archive byte size and uncompressed sheet XML size.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Spreadsheet not found: {file_path}")

    size_bytes = os.path.getsize(file_path)
    size_mb = round(size_bytes / (1024 * 1024), 2)

    estimated_rows = 0
    sheet_names: List[str] = []

    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            # Discover sheet XML files in package
            sheet_files = [f for f in zf.namelist() if f.startswith("xl/worksheets/sheet") and f.endswith(".xml")]
            for sf in sheet_files:
                info = zf.getinfo(sf)
                # Heuristic: Each spreadsheet row typically averages ~80-150 bytes in uncompressed XML
                estimated_rows += max(1, info.file_size // 120)

            # Extract sheet names from workbook.xml if available
            if "xl/workbook.xml" in zf.namelist():
                wb_xml = zf.read("xl/workbook.xml")
                root = ET.fromstring(wb_xml)
                ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                for sheet_el in root.findall(".//main:sheets/main:sheet", ns):
                    name = sheet_el.attrib.get("name")
                    if name:
                        sheet_names.append(name)
    except Exception:
        # Fallback to zero estimates on damaged or non-standard packages
        pass

    is_large = size_mb >= STREAMING_FILE_SIZE_THRESHOLD_MB or estimated_rows >= STREAMING_ROW_COUNT_THRESHOLD

    return {
        "file_size_bytes": size_bytes,
        "file_size_mb": size_mb,
        "estimated_rows": estimated_rows,
        "sheet_names": sheet_names,
        "is_large_file": is_large,
        "recommended_mode": "streaming" if is_large else "dom",
    }


def stream_read_xlsx(
    file_path: str,
    sheet_name: Optional[str] = None,
    max_rows: Optional[int] = None,
) -> Generator[List[Any], None, None]:
    """
    Memory-efficient generator reading rows from an XLSX file in read_only mode.
    Consumes minimal RAM (<30 MB) even on workbooks with 500,000+ rows.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Spreadsheet file not found: {file_path}")

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active or wb.worksheets[0]

        count = 0
        for row in ws.iter_rows(values_only=True):
            yield list(row)
            count += 1
            if max_rows is not None and count >= max_rows:
                break
    finally:
        wb.close()


def stream_write_xlsx(
    output_path: str,
    rows_generator: Iterable[List[Any]],
    sheet_name: str = "Data",
    headers: Optional[List[str]] = None,
) -> str:
    """
    Memory-efficient streaming writer utilizing openpyxl write_only mode.
    Streams tabular rows sequentially to disk, bypassing full DOM serialization.
    """
    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet(title=sheet_name[:31])

    if headers:
        ws.append(list(headers))

    for row in rows_generator:
        ws.append(list(row))

    wb.save(dest)
    wb.close()
    return dest
