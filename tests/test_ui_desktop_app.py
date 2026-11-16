"""Unit tests for P1.4 Desktop Tkinter UI improvements (headless-safe)."""

from __future__ import annotations

import os
import tkinter as tk
from docx import Document
import openpyxl
import pytest

from ui.desktop_app import UniversalDocStudioApp


@pytest.fixture
def tk_app():
    try:
        root = tk.Tk()
    except tk.TclError as exc:
        pytest.skip(f"Tk display unavailable: {exc}")
    root.withdraw()
    app = UniversalDocStudioApp(root)
    try:
        yield app
    finally:
        root.destroy()


def test_desktop_app_supports_xls_xlsx_routing_and_quick_actions(tmp_path, monkeypatch, tk_app):
    # 1. File extension routing in Single Conversion tab
    tk_app.set_single_file("old_report.xls")
    assert tk_app._get_selected_target_ext() == ".xlsx"
    assert tk_app.single_out_var.get() == "old_report.xlsx"

    tk_app.set_single_file("modern_report.xlsx")
    assert tk_app._get_selected_target_ext() == ".pdf"

    # Suppress blocking dialogs
    infos = []
    monkeypatch.setattr("ui.desktop_app.messagebox.showinfo", lambda title, msg: infos.append((title, msg)))
    monkeypatch.setattr("ui.desktop_app.messagebox.askyesno", lambda title, msg: True)

    # 2. Tab 4 (Enterprise/Pure Office) should not hit empty batch error
    tk_app.notebook.select(3)
    tk_app._start_conversion()
    assert any("Pure Office" in t for t, _ in infos)

    # 3. Quick Excel AST Eval
    wb_path = str(tmp_path / "calc.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 10
    ws["A2"] = 25
    ws["A3"] = "=SUM(A1:A2)"
    wb.save(wb_path)
    wb.close()

    tk_app.ent_src_var.set(wb_path)
    tk_app._run_quick_excel_eval()
    assert os.path.isfile(tk_app.last_output_path)
    wb_eval = openpyxl.load_workbook(tk_app.last_output_path, data_only=True)
    assert wb_eval.active["A3"].value == 35
    wb_eval.close()

    # 4. Quick Word NĐ 30/2020 Polish
    docx_path = str(tmp_path / "sample.docx")
    doc = Document()
    doc.add_paragraph("Noi dung hanh chinh")
    doc.save(docx_path)

    tk_app.ent_src_var.set(docx_path)
    tk_app._run_quick_word_polish()
    assert os.path.isfile(tk_app.last_output_path)

    # 5. Quick Print via Virtual Spooler mock (isolate output/spool into tmp_path)
    import core.printing.printer_manager as pm_mod
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(pm_mod.PrinterManager, "get_default_printer", classmethod(lambda cls: "Virtual Spooler (PDF Buffer)"))
    pdf_path = str(tmp_path / "to_print.pdf")
    import pymupdf as fitz
    pdf_doc = fitz.open()
    pdf_doc.new_page()
    pdf_doc.save(pdf_path)
    pdf_doc.close()

    tk_app.ent_src_var.set(pdf_path)
    tk_app._run_quick_print()
    assert any("gửi lệnh in" in t.lower() for t, _ in infos)
