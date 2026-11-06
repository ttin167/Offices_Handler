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

