# -*- coding: utf-8 -*-
"""
tests/test_xlsx_power_suite.py — Unit Tests for Pillar 1: Excel Power Suite & 57-Formula Playbook.
"""

from __future__ import annotations

import datetime
import zipfile
from pathlib import Path

import openpyxl
import pytest

from core.xlsx.formula_playbook import FORMULA_CATALOG, get_formula_catalog
from core.xlsx.power_tools import (
    build_pivot_summary_and_chart,
    clean_and_standardize_workbook,
    diff_workbooks_visual,
    split_workbook_by_column,
    transform_workbook_columns,
    xlookup_join_workbooks,
)


def _make_sample_hr_workbook(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "NhanSu"
    ws.append(["MaNV", "HoVaTen", "SoDienThoai", "CCCD", "NgayVaoLam", "PhongBan", "Luong"])
    ws.append(["NV001", "  nguyễn   văn   an  ", "0912345678", "001200012345", "ngày 05 tháng 03 năm 2024", "Kế toán", 15000000])
    ws.append(["NV001", "nguyễn văn an", "0912345678", "001200012345", "05/03/2024", "Kế toán", 15000000])
    ws.append(["NV002", "Lê   Thị  Minh Anh", "0987654321", "079199000111", "14/11/2026", "Kỹ thuật", 22000000])
    ws.append(["NV003", "trần bình", "0369112233", "031201009988", "01-02-2025", "Kế toán", 18000000])
    wb.save(str(path))
    wb.close()


def test_power_clean_dedup_dates_and_leading_zero_preservation(tmp_path: Path) -> None:
    src = tmp_path / "raw_hr.xlsx"
    out = tmp_path / "cleaned_hr.xlsx"
    _make_sample_hr_workbook(src)

    res = clean_and_standardize_workbook(
        input_path=str(src),
        output_path=str(out),
        dedup_columns=["MaNV"],
        normalize_dates=True,
        normalize_phones_and_ids=True,
        trim_whitespace=True,
        case_mode="title",
    )

    assert res["ok"] is True
    assert out.is_file()
    assert res["duplicates_removed"] == 1
    assert res["rows_processed"] == 4
    assert res["ids_protected"] >= 3

    wb = openpyxl.load_workbook(str(out))
    ws = wb["NhanSu"]
    assert ws.max_row == 4  # Header + 3 unique rows
    assert ws.cell(2, 2).value == "Nguyễn Văn An"
    assert ws.cell(2, 3).value == "0912345678"
    assert ws.cell(2, 3).number_format == "@"
    assert ws.cell(2, 4).value == "001200012345"
    assert ws.cell(2, 4).number_format == "@"
    assert isinstance(ws.cell(2, 5).value, (datetime.date, datetime.datetime))
    assert ws.cell(2, 5).value.strftime("%d/%m/%Y") == "05/03/2024"
    assert ws.cell(2, 5).number_format == "dd/mm/yyyy"
    assert ws.cell(4, 5).value.strftime("%d/%m/%Y") == "01/02/2025"
    wb.close()


def test_transform_columns_split_vn_fullname_and_merge(tmp_path: Path) -> None:
    src = tmp_path / "names.xlsx"
    split_out = tmp_path / "names_split.xlsx"
    merge_out = tmp_path / "names_merged.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DS"
    ws.append(["MaNV", "HoVaTen", "PhongBan"])
    ws.append(["NV01", "Nguyễn Văn An", "Hành chính"])
    ws.append(["NV02", "Lê Thị Minh Anh", "Kỹ thuật"])
    ws.append(["NV03", "Trần Bình", "Kế toán"])
    wb.save(str(src))
    wb.close()

    res_split = transform_workbook_columns(
        input_path=str(src),
        mode="split_vn_fullname",
        source_columns=["HoVaTen"],
        output_path=str(split_out),
    )
    assert res_split["ok"] is True
    assert res_split["rows_transformed"] == 3
    assert len(res_split["created_columns"]) == 3

    wb_s = openpyxl.load_workbook(str(split_out))
    ws_s = wb_s["DS"]
    headers = [ws_s.cell(1, c).value for c in range(1, ws_s.max_column + 1)]
    assert "Họ" in headers and "Tên đệm" in headers and "Tên" in headers
    ho_col = headers.index("Họ") + 1
    dem_col = headers.index("Tên đệm") + 1
    ten_col = headers.index("Tên") + 1
    assert (ws_s.cell(2, ho_col).value, ws_s.cell(2, dem_col).value, ws_s.cell(2, ten_col).value) == ("Nguyễn", "Văn", "An")
    assert (ws_s.cell(3, ho_col).value, ws_s.cell(3, dem_col).value, ws_s.cell(3, ten_col).value) == ("Lê", "Thị Minh", "Anh")
    assert (ws_s.cell(4, ho_col).value, ws_s.cell(4, dem_col).value or "", ws_s.cell(4, ten_col).value) == ("Trần", "", "Bình")
    wb_s.close()

    res_merge = transform_workbook_columns(
        input_path=str(src),
        mode="merge_columns",
        source_columns=["MaNV", "PhongBan"],
        output_path=str(merge_out),
        joiner=" - ",
        new_column_name="Ma_Phong",
    )
    assert res_merge["ok"] is True
    wb_m = openpyxl.load_workbook(str(merge_out))
    ws_m = wb_m["DS"]
    assert ws_m.cell(1, ws_m.max_column).value == "Ma_Phong"
    assert ws_m.cell(2, ws_m.max_column).value == "NV01 - Hành chính"
    wb_m.close()


def test_pivot_summary_and_chart_creation(tmp_path: Path) -> None:
    src = tmp_path / "payroll.xlsx"
    out = tmp_path / "payroll_pivot.xlsx"
    _make_sample_hr_workbook(src)

    res = build_pivot_summary_and_chart(
        input_path=str(src),
        group_by_col="PhongBan",
        value_cols=["Luong"],
        agg_func="SUM",
        chart_type="bar",
        output_path=str(out),
    )

    assert res["ok"] is True
    assert res["summary_sheet"] == "Pivot_Summary"
    assert res["groups_count"] == 2

    wb = openpyxl.load_workbook(str(out))
    assert "Pivot_Summary" in wb.sheetnames
    ws_p = wb["Pivot_Summary"]
    assert len(ws_p._charts) >= 1
    # Ensure live dynamic SUMIFS / COUNTIFS formulas are generated (Invariant E3)
    formulas = [
        str(ws_p.cell(r, c).value or "")
        for r in range(1, ws_p.max_row + 1)
        for c in range(1, ws_p.max_column + 1)
    ]
    assert any("SUMIFS(" in f or "COUNTIFS(" in f for f in formulas)
    wb.close()


def test_xlookup_join_same_workbook_and_cross_workbook(tmp_path: Path) -> None:
    wb_both = openpyxl.Workbook()
    ws_left = wb_both.active
    ws_left.title = "DonHang"
    ws_left.append(["MaNV", "SoTien"])
    ws_left.append(["NV01", 500000])
    ws_left.append(["NV99", 200000])

    ws_right = wb_both.create_sheet("DanhMuc")
    ws_right.append(["MaNV", "HoTen", "PhongBan"])
    ws_right.append(["NV01", "Nguyễn Văn An", "Kinh doanh"])
    same_path = tmp_path / "join_same.xlsx"
    out_same = tmp_path / "join_same_out.xlsx"
    wb_both.save(str(same_path))
    wb_both.close()

    res_same = xlookup_join_workbooks(
        left_path=str(same_path),
        right_path=None,
        left_key="MaNV",
        right_key="MaNV",
        return_cols=["HoTen", "PhongBan"],
        output_path=str(out_same),
        left_sheet="DonHang",
        right_sheet="DanhMuc",
        not_found_value="Chưa gán",
        formula_mode="auto",
    )
    assert res_same["ok"] is True
    assert res_same["matched_rows"] == 1
    assert res_same["unmatched_rows"] == 1
    assert res_same["live_formulas"] is True

    wb_check = openpyxl.load_workbook(str(out_same))
    ws_res = wb_check["DonHang"]
    assert "INDEX(" in str(ws_res.cell(2, 3).value)
    wb_check.close()

    # Cross-workbook join
    right_ext = tmp_path / "right_master.xlsx"
    wb_r = openpyxl.Workbook()
    ws_re = wb_r.active
    ws_re.append(["MaNV", "Email"])
    ws_re.append(["NV01", "an.nv@company.vn"])
    wb_r.save(str(right_ext))
    wb_r.close()

    out_cross = tmp_path / "join_cross_out.xlsx"
    res_cross = xlookup_join_workbooks(
        left_path=str(same_path),
        right_path=str(right_ext),
        left_key="MaNV",
        right_key="MaNV",
        return_cols=["Email"],
        output_path=str(out_cross),
        left_sheet="DonHang",
        not_found_value="N/A",
    )
    assert res_cross["ok"] is True
    assert res_cross["live_formulas"] is False
    wb_c = openpyxl.load_workbook(str(out_cross))
    ws_c = wb_c["DonHang"]
    assert ws_c.cell(2, 3).value == "an.nv@company.vn"
    assert ws_c.cell(3, 3).value == "N/A"
    wb_c.close()


def test_split_workbook_by_column_and_visual_diff(tmp_path: Path) -> None:
    src = tmp_path / "master.xlsx"
    _make_sample_hr_workbook(src)

    split_zip = tmp_path / "dept_split.zip"
    res_split = split_workbook_by_column(
        input_path=str(src),
        split_col="PhongBan",
        output_path=str(split_zip),
        as_zip=True,
    )
    assert res_split["ok"] is True
    assert res_split["total_files"] == 2
    assert split_zip.is_file()
    with zipfile.ZipFile(str(split_zip), "r") as zf:
        names = zf.namelist()
        assert len(names) == 2
        assert all(n.endswith(".xlsx") for n in names)

    # Visual Diff between two versions
    v1 = tmp_path / "v1.xlsx"
    v2 = tmp_path / "v2.xlsx"
    diff_out = tmp_path / "diff_report.xlsx"

    wb1 = openpyxl.Workbook()
    ws1 = wb1.active
    ws1.title = "Data"
    ws1.append(["Ma", "GiaTri"])
    ws1.append(["A1", 100])
    wb1.save(str(v1))
    wb1.close()

    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "Data"
    ws2.append(["Ma", "GiaTri"])
    ws2.append(["A1", 150])  # Modified
    ws2.append(["A2", 200])  # Added
    wb2.save(str(v2))
    wb2.close()

    res_diff = diff_workbooks_visual(str(v1), str(v2), output_path=str(diff_out))
    assert res_diff["ok"] is True
    assert res_diff["modified_cells"] >= 1
    assert res_diff["added_cells"] >= 1
    wb_d = openpyxl.load_workbook(str(diff_out))
    assert "_Diff_Summary" in wb_d.sheetnames
    wb_d.close()


def test_formula_playbook_expanded_catalog() -> None:
    assert len(FORMULA_CATALOG) >= 52
    all_items = get_formula_catalog()
    assert len(all_items) == len(FORMULA_CATALOG)
    cat_ids = {c["category"] for c in all_items}
    assert {"aggregation", "logic", "text", "lookup", "dynamic", "financial"}.issubset(cat_ids)
