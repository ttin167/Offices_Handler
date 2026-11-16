# -*- coding: utf-8 -*-
"""
tests/test_xlsx_formula_evaluator.py — Comprehensive Test Suite for AST Offline Formula Engine.
"""

import os
import tempfile
import pytest
import openpyxl

from core.xlsx.formula_evaluator import (
    ASTFormulaEvaluator,
    ExcelFunctions,
    tokenize_excel_formula,
    matches_excel_criteria,
)


def test_tokenizer_and_criteria():
    tokens = tokenize_excel_formula("=SUM(A1:A10) + IF(B1 > 5, 10, 0)")
    assert len(tokens) > 5
    assert tokens[0] == ('FUNC', 'SUM')

    # Test criteria matching
    assert matches_excel_criteria(10, 10) is True
    assert matches_excel_criteria(15, ">10") is True
    assert matches_excel_criteria(5, "<=5") is True
    assert matches_excel_criteria("Hà Nội", "Hà*") is True
    assert matches_excel_criteria("Sales Dept", "*Dept") is True
    assert matches_excel_criteria(10, "<>20") is True
    assert matches_excel_criteria(20, "<>20") is False


def test_math_and_statistics():
    evaluator = ASTFormulaEvaluator()
    evaluator.set_cell_value("Sheet1", "A1", 10)
    evaluator.set_cell_value("Sheet1", "A2", 20)
    evaluator.set_cell_value("Sheet1", "A3", 30)

    assert evaluator.evaluate_formula("=SUM(A1:A3)") == 60.0
    assert evaluator.evaluate_formula("=AVERAGE(A1:A3)") == 20.0
    assert evaluator.evaluate_formula("=COUNT(A1:A3)") == 3
    assert evaluator.evaluate_formula("=MIN(A1:A3)") == 10.0
    assert evaluator.evaluate_formula("=MAX(A1:A3)") == 30.0
    assert evaluator.evaluate_formula("=ROUND(123.456, 2)") == 123.46
    assert evaluator.evaluate_formula("=INT(99.9)") == 99
    assert evaluator.evaluate_formula("=ABS(-45)") == 45.0
    assert evaluator.evaluate_formula("=MOD(10, 3)") == 1.0


def test_conditional_aggregation():
    evaluator = ASTFormulaEvaluator()
    # Departments
    evaluator.set_cell_value("Sheet1", "A1", "IT")
    evaluator.set_cell_value("Sheet1", "A2", "HR")
    evaluator.set_cell_value("Sheet1", "A3", "IT")
    evaluator.set_cell_value("Sheet1", "A4", "Sales")
    # Salaries
    evaluator.set_cell_value("Sheet1", "B1", 1500)
    evaluator.set_cell_value("Sheet1", "B2", 1200)
    evaluator.set_cell_value("Sheet1", "B3", 1800)
    evaluator.set_cell_value("Sheet1", "B4", 2000)

    # SUMIF IT salary
    assert evaluator.evaluate_formula('=SUMIF(A1:A4, "IT", B1:B4)') == 3300.0
    # COUNTIF IT employees
    assert evaluator.evaluate_formula('=COUNTIF(A1:A4, "IT")') == 2
    # AVERAGEIF IT
    assert evaluator.evaluate_formula('=AVERAGEIF(A1:A4, "IT", B1:B4)') == 1650.0

    # COUNTIFS with multiple conditions
    assert evaluator.evaluate_formula('=COUNTIFS(A1:A4, "IT", B1:B4, ">1600")') == 1


def test_logical_and_text_functions():
    evaluator = ASTFormulaEvaluator()
    evaluator.set_cell_value("Sheet1", "A1", 85)

    assert evaluator.evaluate_formula('=IF(A1 >= 80, "Distinction", "Pass")') == "Distinction"
    assert evaluator.evaluate_formula('=IFS(A1 >= 90, "A", A1 >= 80, "B", A1 >= 70, "C")') == "B"
    assert evaluator.evaluate_formula('=AND(A1 > 50, A1 < 100)') is True
    assert evaluator.evaluate_formula('=OR(A1 < 50, A1 > 80)') is True
    assert evaluator.evaluate_formula('=NOT(A1 < 50)') is True
    assert evaluator.evaluate_formula('=IFERROR(10 / 0, "Error Occurred")') == "Error Occurred"

    # Text functions
    evaluator.set_cell_value("Sheet1", "T1", "Hello")
    evaluator.set_cell_value("Sheet1", "T2", "World")
    assert evaluator.evaluate_formula('=TEXTJOIN(" ", TRUE, T1:T2)') == "Hello World"
    assert evaluator.evaluate_formula('=UPPER("vietnam")') == "VIETNAM"
    assert evaluator.evaluate_formula('=LOWER("ANTIGRAVITY")') == "antigravity"
    assert evaluator.evaluate_formula('=TRIM("   Clean Text   ")') == "Clean Text"
    assert evaluator.evaluate_formula('=LEN("12345")') == 5


def test_lookups_xlookup_vlookup():
    evaluator = ASTFormulaEvaluator()
    # IDs
    evaluator.set_cell_value("Sheet1", "A1", "EMP01")
    evaluator.set_cell_value("Sheet1", "A2", "EMP02")
    evaluator.set_cell_value("Sheet1", "A3", "EMP03")
    # Names
    evaluator.set_cell_value("Sheet1", "B1", "Nguyễn Văn A")
    evaluator.set_cell_value("Sheet1", "B2", "Trần Thị B")
    evaluator.set_cell_value("Sheet1", "B3", "Lê Văn C")

    # XLOOKUP
    res_x = evaluator.evaluate_formula('=XLOOKUP("EMP02", A1:A3, B1:B3, "Not Found")')
    assert res_x == "Trần Thị B"

    res_x_not_found = evaluator.evaluate_formula('=XLOOKUP("EMP99", A1:A3, B1:B3, "Not Found")')
    assert res_x_not_found == "Not Found"

    # INDEX & MATCH
    assert evaluator.evaluate_formula('=MATCH("EMP03", A1:A3, 0)') == 3
    assert evaluator.evaluate_formula('=INDEX(B1:B3, 2, 1)') == "Trần Thị B"


def test_cross_sheet_and_arithmetic_operators():
    evaluator = ASTFormulaEvaluator()
    evaluator.set_cell_value("DataSheet", "A1", 50)
    evaluator.set_cell_value("DataSheet", "A2", 2)

    assert evaluator.evaluate_formula("=DataSheet!A1 * DataSheet!A2 + 10") == 110.0
    assert evaluator.evaluate_formula("=DataSheet!A1 ^ 2") == 2500.0


def test_openxml_populate_cached_values_and_data_only():
    with tempfile.TemporaryDirectory() as tmp_dir:
        file_path = os.path.join(tmp_dir, "test_eval_output.xlsx")

        # Create workbook with static numbers and formulas
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Revenue"
        ws["A1"] = 100
        ws["A2"] = 250
        ws["A3"] = 150
        ws["A4"] = "=SUM(A1:A3)"
        ws["B1"] = '=IF(A4 > 400, "Target Reached", "Below Target")'

        wb.save(file_path)
        wb.close()

        # Prior to evaluation: openpyxl data_only=True returns None for formula cells
        wb_before = openpyxl.load_workbook(file_path, data_only=True)
        assert wb_before["Revenue"]["A4"].value is None
        assert wb_before["Revenue"]["B1"].value is None
        wb_before.close()

        # Execute populate_cached_values
        patched_path = ASTFormulaEvaluator.populate_cached_values(file_path)
        assert os.path.isfile(patched_path)

        # After evaluation: data_only=True immediately yields exact numbers and text!
        wb_after = openpyxl.load_workbook(patched_path, data_only=True)
        assert wb_after["Revenue"]["A4"].value == 500
        assert wb_after["Revenue"]["B1"].value == "Target Reached"
        wb_after.close()

        # Non data_only=False must still preserve the original formula text!
        wb_formulas = openpyxl.load_workbook(patched_path, data_only=False)
        assert wb_formulas["Revenue"]["A4"].value == "=SUM(A1:A3)"
        assert wb_formulas["Revenue"]["B1"].value == '=IF(A4 > 400, "Target Reached", "Below Target")'
        wb_formulas.close()


def test_circular_reference_unary_op_and_temp_cleanup():
    evaluator = ASTFormulaEvaluator()
    evaluator.set_cell_value("Sheet1", "A1", "=B1")
    evaluator.set_cell_value("Sheet1", "B1", "=A1")
    evaluator.set_cell_value("Sheet1", "C1", 25)
    evaluator.set_cell_value("Sheet1", "D1", None)
    evaluator.set_cell_value("Sheet1", "E1", "15")

    assert evaluator.get_cell_value("Sheet1", "A1") == "#REF!"
    assert evaluator.evaluate_formula("=-C1") == -25.0
    assert evaluator.evaluate_formula("=-D1") == 0.0
    assert evaluator.evaluate_formula("=-E1 + 5") == -10.0

    with tempfile.TemporaryDirectory() as tmp_dir:
        wb_path = os.path.join(tmp_dir, "unary.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = 40
        ws["A2"] = "=-A1"
        wb.save(wb_path)
        wb.close()

        ASTFormulaEvaluator.populate_cached_values(wb_path)
        leftover_eval_dirs = [n for n in os.listdir(tmp_dir) if n.startswith(".ag_eval_")]
        assert leftover_eval_dirs == []

        wb_val = openpyxl.load_workbook(wb_path, data_only=True)
        assert wb_val["Data"]["A2"].value == -40
        wb_val.close()


def test_p2_full_catalog_functions_and_dynamic_scopes(tmp_path):
    import datetime

    ev = ASTFormulaEvaluator()
    # Grid for lookup, aggregation, sorting, financial & date tests
    ev.set_cell_value("Sheet1", "A1", "Dept")
    ev.set_cell_value("Sheet1", "B1", "Salary")
    ev.set_cell_value("Sheet1", "C1", "Score")

    rows = [
        ("IT", 1500, 92),
        ("HR", 1200, 80),
        ("IT", 2000, 95),
        ("Sales", 1800, 88),
    ]
    for idx, (dept, sal, sc) in enumerate(rows, start=2):
        ev.set_cell_value("Sheet1", f"A{idx}", dept)
        ev.set_cell_value("Sheet1", f"B{idx}", sal)
        ev.set_cell_value("Sheet1", f"C{idx}", sc)

    # Aggregation & rounding & statistics (including dotted function names)
    assert ev.evaluate_formula("=ROUNDUP(1234, -2)") == 1300.0
    assert ev.evaluate_formula("=ROUNDDOWN(1299, -2)") == 1200.0
    assert ev.evaluate_formula("=SUMPRODUCT((A2:A5=\"IT\")*(B2:B5))") == 3500.0
    assert ev.evaluate_formula("=SUBTOTAL(9, B2:B5)") == 6500.0
    ev.set_cell_value("Sheet1", "D2", 100)
    ev.set_cell_value("Sheet1", "D3", "#N/A")
    ev.set_cell_value("Sheet1", "D4", 300)
    assert ev.evaluate_formula("=AGGREGATE(9, 6, D2:D4)") == 400.0
    assert ev.evaluate_formula("=RANK.EQ(2000, B2:B5)") == 1
    assert ev.evaluate_formula("=PERCENTILE.INC(B2:B5, 0.5)") == 1650.0

    # Text & Info & Logic
    assert ev.evaluate_formula('=LEFT("Antigravity", 4)') == "Anti"
    assert ev.evaluate_formula('=RIGHT("Antigravity", 4)') == "vity"
    assert ev.evaluate_formula('=MID("Antigravity", 5, 4)') == "grav"
    assert ev.evaluate_formula('=SUBSTITUTE("A-B-C", "-", "/", 2)') == "A-B/C"
    assert ev.evaluate_formula('=PROPER("nguyễn văn an")') == "Nguyễn Văn An"
    assert ev.evaluate_formula('=EXACT("VN", "vn")') is False
    assert ev.evaluate_formula('=TEXTBEFORE("EMP-2026-HN", "-", 2)') == "EMP-2026"
    assert ev.evaluate_formula('=TEXTAFTER("EMP-2026-HN", "-", 1)') == "2026-HN"
    assert ev.evaluate_formula('=NUMBERVALUE("1.250,50", ",", ".")') == 1250.5
    assert ev.evaluate_formula('=10 + 5 & " đ"') == "15 đ"
    assert ev.evaluate_formula('=TEXT(1500000, "#,##0")') == "1,500,000"
    assert ev.evaluate_formula('=IFNA(#N/A, "Missing")') == "Missing"
    assert ev.evaluate_formula('=SWITCH("B", "A", 1, "B", 2, 0)') == 2
    assert ev.evaluate_formula('=CHOOSE(2, "Low", "Mid", "High")') == "Mid"
    assert ev.evaluate_formula("=ISNUMBER(TRUE)") is False
    assert ev.evaluate_formula("=ISNUMBER(1500)") is True
    assert ev.evaluate_formula('=ISBLANK(Z99)') is True

    # Lookup modes & dynamic references (XLOOKUP -1/1, VLOOKUP range, HLOOKUP, XMATCH, OFFSET, INDIRECT)
    ev.set_cell_value("Sheet1", "F1", 0)
    ev.set_cell_value("Sheet1", "F2", 100)
    ev.set_cell_value("Sheet1", "F3", 500)
    ev.set_cell_value("Sheet1", "G1", "Tier 1")
    ev.set_cell_value("Sheet1", "G2", "Tier 2")
    ev.set_cell_value("Sheet1", "G3", "Tier 3")
    assert ev.evaluate_formula('=XLOOKUP(250, F1:F3, G1:G3, "None", -1)') == "Tier 2"
    assert ev.evaluate_formula('=XLOOKUP(250, F1:F3, G1:G3, "None", 1)') == "Tier 3"
    assert ev.evaluate_formula("=VLOOKUP(250, F1:G3, 2, TRUE)") == "Tier 2"
    assert ev.evaluate_formula("=HLOOKUP(\"Salary\", A1:C5, 3, FALSE)") == 1200
    assert ev.evaluate_formula("=XMATCH(250, F1:F3, -1)") == 2
    assert ev.evaluate_formula('=INDIRECT("G3")') == "Tier 3"
    assert ev.evaluate_formula("=OFFSET(F1, 2, 1)") == "Tier 3"
    assert ev.evaluate_formula('=HYPERLINK("https://example.com", "Docs")') == "Docs"

    # Dynamic Arrays, LET, LAMBDA
    assert ev.evaluate_formula("=COUNTA(UNIQUE(A2:A5))") == 3
    assert ev.evaluate_formula("=INDEX(SORT(B2:B5, 1, -1), 1, 1)") == 2000
    assert ev.evaluate_formula('=SUM(FILTER(B2:B5, A2:A5="IT"))') == 3500.0
    assert ev.evaluate_formula("=SUM(SEQUENCE(3, 1, 10, 5))") == 45.0
    assert ev.evaluate_formula("=SUM(TAKE(SORT(B2:B5, 1, -1), 2))") == 3800.0
    assert ev.evaluate_formula("=LET(x, 10, y, x * 2, x + y)") == 30.0
    assert ev.evaluate_formula("=LET(fn, LAMBDA(a, b, a * 2 + b), fn(5, 3))") == 13.0

    # Dates & Financial
    assert ev.evaluate_formula('=EOMONTH("2024-02-10", 0)') == datetime.date(2024, 2, 29)
    assert ev.evaluate_formula('=EDATE("2026-01-31", 1)') == datetime.date(2026, 2, 28)
    assert ev.evaluate_formula('=DATEDIF("2020-03-15", "2026-05-20", "Y")') == 6
    assert ev.evaluate_formula('=NETWORKDAYS.INTL("2026-11-16", "2026-11-22", 1)') == 5
    assert ev.evaluate_formula('=WORKDAY.INTL("2026-11-16", 5, 1)') == datetime.date(2026, 11, 23)
    assert abs(ev.evaluate_formula("=PMT(0.01, 12, -1000)") - 88.84878868) < 1e-4
    assert abs(ev.evaluate_formula("=NPV(0.1, 100, 100)") - 173.553719) < 1e-4
    assert abs(ev.evaluate_formula("=FV(0, 10, -100, -500)") - 1500.0) < 1e-6

    # Cached <v> OpenXML write with LET + FILTER + RANK.EQ + EOMONTH
    wb_path = str(tmp_path / "p2_cached.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KPI"
    ws["A1"] = 100
    ws["A2"] = 300
    ws["A3"] = 200
    ws["B1"] = "=RANK.EQ(A2, A1:A3)"
    ws["B2"] = "=LET(bonus, 50, SUM(A1:A3) + bonus)"
    ws["B3"] = '=EOMONTH("2026-02-05", 0)'
    wb.save(wb_path)
    wb.close()

    ASTFormulaEvaluator.populate_cached_values(wb_path)
    wb_cached = openpyxl.load_workbook(wb_path, data_only=True)
    assert wb_cached["KPI"]["B1"].value == 1
    assert wb_cached["KPI"]["B2"].value == 650
    assert wb_cached["KPI"]["B3"].value is not None
    wb_cached.close()


