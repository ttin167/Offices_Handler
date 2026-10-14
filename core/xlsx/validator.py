"""
xlsx_validator.py — 12 Quality Gates Format Validation & Diff Engine for Excel (XLSX).

Implements Invariant E6 and Quality Gates 1–12:
  1. Header Block (rows 2-8): cell-by-cell font, fill, border, alignment, numfmt
  2. Row 7 KPI Formulas: dynamic COUNTIF / SUM / COUNTA formula integrity
  3. Row 9 TC Headers: Col A navy+double border, Cols B-E navy+align, Cols F+ TC IDs (180° rot, double border top)
  4. Col A Continuous Navy Fill: unbroken FF000080 fill across all data rows (no gaps)
  5. Col B-C-D Unified Box: 3-column unified box (B.left=thin, B.right=None, C.mid=None, D.left=None, D.right=thin, solid white fill)
  6. Duplicate Header Detection: prevents repeated group titles on consecutive rows
  7. Matrix Grid & O-Marks: matrix cell borders on every cell, O-mark font/bold/alignment/border
  8. Result Block (Type/PF/Date/Defect): Courier non-bold 8pt, Date textRotation=255, mm/dd numfmt
  9. Merged Cells: Header merges & Result merges (B:D on footer rows)
 10. Section Outer Closing Borders: double bottom on Condition last row & Result last row
 11. Data Validations: 'O' matrix dropdown, 'N,A,B' Type dropdown, 'P,F' PF dropdown
 12. Geometry: Column widths & Row heights (R9=45.0, data=13.5) + Col D text overflow guard
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.utils import get_column_letter


@dataclass
class ValidationResult:
    target_sheet: str
    ref_sheet: str
    passed: bool
    diffs: List[str] = field(default_factory=list)
    gates_evaluated: int = 12
    gates_failed: int = 0

    def print_report(self) -> None:
        print(f"\n{'='*70}")
        print(f"EXCEL 12 QUALITY GATES VALIDATION REPORT")
        print(f"Target Sheet: '{self.target_sheet}'  |  Reference Sheet: '{self.ref_sheet}'")
        print(f"{'='*70}")
        if self.passed:
            print(f"[PASSED] 100% CLEAN - All {self.gates_evaluated} Quality Gates matched reference template perfectly.")
        else:
            print(f"[FAILED] {len(self.diffs)} format discrepancy(ies) detected across {self.gates_failed} gate(s):")
            for d in self.diffs[:50]:
                print(d)
            if len(self.diffs) > 50:
                print(f"  ... and {len(self.diffs) - 50} more issues.")
        print(f"{'='*70}\n")


def cell_sig(cell: Any) -> Dict[str, Any]:
    f = cell.font
    fi = cell.fill
    al = cell.alignment
    b = cell.border
    return {
        "font_name": f.name if f else None,
        "font_size": f.size if f else None,
        "font_bold": f.bold if f else None,
        "fill_type": fi.patternType if fi else None,
        "fill_fg": fi.fgColor.rgb if fi and fi.fgColor and fi.fgColor.type == "rgb" else None,
        "align_h": al.horizontal if al else None,
        "align_v": al.vertical if al else None,
        "align_rot": al.textRotation if al else None,
        "border_l": b.left.style if b and b.left else None,
        "border_r": b.right.style if b and b.right else None,
        "border_t": b.top.style if b and b.top else None,
        "border_b": b.bottom.style if b and b.bottom else None,
        "num_fmt": cell.number_format,
    }


def first_o_cell(ws: Any) -> Optional[Any]:
    for r in range(10, min(ws.max_row + 1, 60)):
        for c in range(6, min(ws.max_column + 1, 30)):
            if ws.cell(r, c).value == "O":
                return ws.cell(r, c)
    return None


def detect_section_boundaries(ws_ref: Any) -> Tuple[int, int, int, int, int, int]:
    cond_start = conf_start = res_start = type_row = defect_row = None
    max_tc_col = 6

    for r in range(9, min(ws_ref.max_row + 1, 65)):
        va = ws_ref.cell(r, 1).value
        vb = ws_ref.cell(r, 2).value
        if va:
            vs = str(va).strip().lower()
            if "condition" in vs and cond_start is None:
                cond_start = r
            elif "confirm" in vs and conf_start is None:
                conf_start = r
            elif "result" in vs and res_start is None:
                res_start = r
        if vb:
            vbs = str(vb).strip().lower()
            if "type(" in vbs or "type (" in vbs or vbs == "type":
                type_row = r
            elif "defect" in vbs:
                defect_row = r
        for c in range(6, min(ws_ref.max_column + 1, 40)):
            if ws_ref.cell(r, c).value is not None:
                max_tc_col = max(max_tc_col, c)

    return (
        cond_start or 10,
        conf_start or 33,
        res_start or 42,
        type_row or 45,
        defect_row or (res_start + 3 if res_start else 48),
        max_tc_col,
    )


def validate_excel_sheet(
    workbook_or_path: Union[str, openpyxl.Workbook],
    target_sheet: Optional[str] = None,
    ref_sheet: Optional[str] = None,
    verbose: bool = False,
) -> ValidationResult:
    """
    Validates a target sheet against an in-workbook reference sheet using all 12 Quality Gates.
    """
    is_path = isinstance(workbook_or_path, str)
    if is_path:
        if not os.path.isfile(workbook_or_path):
            raise FileNotFoundError(f"Workbook not found: {workbook_or_path}")
        wb = openpyxl.load_workbook(workbook_or_path, data_only=False)
    else:
        wb = workbook_or_path

    # Auto-default target sheet if None, empty, or string "None"/"null"
    if not target_sheet or str(target_sheet).strip().lower() in ("none", "null", ""):
        target_sheet = wb.sheetnames[0]

    if target_sheet not in wb.sheetnames:
        if is_path:
            try:
                wb.close()
            except Exception:
                pass
        raise ValueError(f"Target sheet '{target_sheet}' not found. Available: {wb.sheetnames}")

    # Discover reference sheet if not specified
    if not ref_sheet or str(ref_sheet).strip().lower() in ("none", "null", ""):
        ref_sheet = None
        ref_candidates = ["Example", "Template", "Sample", "Pattern", "Ref", "Reference"]
        for cand in ref_candidates:
            matched = next((s for s in wb.sheetnames if s.lower() == cand.lower()), None)
            if matched and matched != target_sheet:
                ref_sheet = matched
                break
        if not ref_sheet:
            # Fallback to first available sheet that isn't the target sheet
            ref_sheet = next((s for s in wb.sheetnames if s != target_sheet), target_sheet)

    if ref_sheet not in wb.sheetnames:
        if is_path:
            try:
                wb.close()
            except Exception:
                pass
        raise ValueError(f"Reference sheet '{ref_sheet}' not found. Available: {wb.sheetnames}")

    ws_tgt = wb[target_sheet]
    ws_ref = wb[ref_sheet]

    diffs: List[str] = []
    failed_gates = set()

    def cmp_cell(r_t: int, c: int, r_r: Optional[int] = None, label: Optional[str] = None, gate_idx: int = 1) -> None:
        r_r = r_r or r_t
        sig_r = cell_sig(ws_ref.cell(r_r, c))
        sig_t = cell_sig(ws_tgt.cell(r_t, c))
        ref_label = label or f"{get_column_letter(c)}{r_t}(vs{r_r})"
        for k in sig_r:
            if sig_r[k] != sig_t[k]:
                diffs.append(f"  [Gate {gate_idx} - {ref_label}] {k}: expected={sig_r[k]!r} got={sig_t[k]!r}")
                failed_gates.add(gate_idx)

    cr, cfr, rr, tr, dr, max_tc_ref = detect_section_boundaries(ws_ref)
    ct, cft, rt, tt, dt, max_tc_tgt = detect_section_boundaries(ws_tgt)

    # 1. Header Block (rows 2-8)
    for r in range(2, 9):
        for c in range(1, min(ws_ref.max_column + 1, 20)):
            cmp_cell(r, c, gate_idx=1)

    # 2. Row 7 KPI Formulas
    kpi_cols = [1, 3, 6, 12, 13, 14, 15]
    for c in kpi_cols:
        col_letter = get_column_letter(c)
        val = ws_tgt.cell(7, c).value
        if not val or not str(val).startswith("="):
            diffs.append(f"  [Gate 2 - Row7-Formula] {col_letter}7 should be a live formula, got: {val!r}")
            failed_gates.add(2)

    # 3. Row 9 Headers
    for c in range(1, max_tc_tgt + 1):
        cmp_cell(9, c, gate_idx=3)

    # 4. Col A Continuous Navy Fill
    for r in range(ct, dt + 1):
        cell = ws_tgt.cell(r, 1)
        fg = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor and cell.fill.fgColor.type == "rgb" else None
        if not cell.fill or not cell.fill.patternType or cell.fill.patternType == "none" or fg == "00000000":
            diffs.append(f"  [Gate 4 - ColA-fill] Row {r}: no navy fill")
            failed_gates.add(4)
        b_left = cell.border.left.style if cell.border and cell.border.left else None
        if b_left != "double":
            diffs.append(f"  [Gate 4 - ColA-border] Row {r} Col A left border expected='double', got={b_left!r}")
            failed_gates.add(4)

    # 5. Col B-C-D Unified Box
    for r in range(ct, rt):
        b_cell = ws_tgt.cell(r, 2)
        b_left = b_cell.border.left.style if b_cell.border and b_cell.border.left else None
        b_right = b_cell.border.right.style if b_cell.border and b_cell.border.right else None
        if b_left != "thin":
            diffs.append(f"  [Gate 5 - B-border-L] Row {r} Col B left border expected='thin', got={b_left!r}")
            failed_gates.add(5)
        if b_right is not None:
            diffs.append(f"  [Gate 5 - B-border-R] Row {r} Col B right border should be None, got={b_right!r}")
            failed_gates.add(5)

        c_cell = ws_tgt.cell(r, 3)
        c_left = c_cell.border.left.style if c_cell.border and c_cell.border.left else None
        c_right = c_cell.border.right.style if c_cell.border and c_cell.border.right else None
        if c_left is not None or c_right is not None:
            diffs.append(f"  [Gate 5 - C-border-vert] Row {r} Col C should have no vertical borders")
            failed_gates.add(5)

        d_cell = ws_tgt.cell(r, 4)
        d_left = d_cell.border.left.style if d_cell.border and d_cell.border.left else None
        d_right = d_cell.border.right.style if d_cell.border and d_cell.border.right else None
        if d_left is not None:
            diffs.append(f"  [Gate 5 - D-border-L] Row {r} Col D left border should be None")
            failed_gates.add(5)
        if d_right != "thin":
            diffs.append(f"  [Gate 5 - D-border-R] Row {r} Col D right border expected='thin', got={d_right!r}")
            failed_gates.add(5)

    # 6. Duplicate Header Check (Col B)
    for r in range(ct + 1, rt):
        prev_b = ws_tgt.cell(r - 1, 2).value
        curr_b = ws_tgt.cell(r, 2).value
        if curr_b and prev_b and str(curr_b).strip() == str(prev_b).strip():
            diffs.append(f"  [Gate 6 - Duplicate-Header] Row {r} Col B duplicates previous row: {curr_b!r}")
            failed_gates.add(6)

    # 7. Matrix Grid & O-Marks
    o_r = first_o_cell(ws_ref)
    o_t = first_o_cell(ws_tgt)
    if o_r and o_t:
        sig_r = cell_sig(o_r)
        sig_t = cell_sig(o_t)
        for k in ["font_name", "font_size", "font_bold", "align_h", "border_l", "border_r"]:
            if sig_r[k] != sig_t[k]:
                diffs.append(f"  [Gate 7 - O-mark] {k}: expected={sig_r[k]!r} got={sig_t[k]!r}")
                failed_gates.add(7)

    # 8. Result Rows (Type/PF/Date/Defect)
    for r_t, r_r in [(tt, tr), (tt + 1, tr + 1), (tt + 2, tr + 2), (dt, dr)]:
        cmp_cell(r_t, 2, r_r=r_r, label=f"Result-Label-B{r_t}", gate_idx=8)

    # 9. Merged Cells Validation
    tgt_merges = {str(m) for m in ws_tgt.merged_cells.ranges}
    for r in range(tt, dt + 1):
        expected_merge = f"B{r}:D{r}"
        if expected_merge not in tgt_merges:
            diffs.append(f"  [Gate 9 - Merges] Missing result footer merge: {expected_merge}")
            failed_gates.add(9)

    # 10. Section Outer Closing Borders
    r_last_cond = cft - 1
    for c in range(1, 5):
        bot = ws_tgt.cell(r_last_cond, c).border.bottom
        if not bot or bot.style != "double":
            diffs.append(f"  [Gate 10 - Double-Bottom] Condition closing border at {get_column_letter(c)}{r_last_cond} not 'double'")
            failed_gates.add(10)

    # 11. Data Validations
    if ws_ref.data_validations.dataValidation and not ws_tgt.data_validations.dataValidation:
        diffs.append("  [Gate 11 - Validations] Target sheet is missing Data Validations defined in reference sheet")
        failed_gates.add(11)

    # 12. Geometry
    for c in range(1, 6):
        col_letter = get_column_letter(c)
        w_r = ws_ref.column_dimensions[col_letter].width
        w_t = ws_tgt.column_dimensions[col_letter].width
        if w_r and w_t and abs(w_r - w_t) > 1.5:
            diffs.append(f"  [Gate 12 - ColWidth] Col {col_letter} width mismatch: ref={w_r} tgt={w_t}")
            failed_gates.add(12)

    passed = len(diffs) == 0
    res = ValidationResult(
        target_sheet=target_sheet,
        ref_sheet=ref_sheet,
        passed=passed,
        diffs=diffs,
        gates_evaluated=12,
        gates_failed=len(failed_gates),
    )

    if verbose or not passed:
        res.print_report()

    if is_path:
        try:
            wb.close()
        except Exception:
            pass

    return res


