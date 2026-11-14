# -*- coding: utf-8 -*-
"""
tests/test_xlsx_tracker.py — Unit Tests for Pure Office Work & Daily Task Tracker.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

from core.xlsx.tracker import OfficeWorkTracker
from tools import project_tracker


def test_office_work_tracker_lifecycle_and_daily_sync(tmp_path: Path) -> None:
    proj_file = tmp_path / "project_tracking.xlsx"
    task_file = tmp_path / "task_tracking.xlsx"

    # Pre-seed task_tracking.xlsx with legacy date sheet '14-11'
    wb_t = openpyxl.Workbook()
    ws_t = wb_t.active
    ws_t.title = "14-11"
    ws_t.append(["Báo cáo Công việc Hàng ngày - 14/11"])
    ws_t.append(["STT", "Bộ phận", "Nội dung công việc", "Trạng thái", "Ghi chú"])
    wb_t.save(str(task_file))
    wb_t.close()

    tracker = OfficeWorkTracker(project_file=proj_file, task_file=task_file)

    init_res = tracker.init_project_tracking(force=True)
    assert init_res["ok"] is True
    assert proj_file.is_file()

    # Add a Pure Office tracking record dated 14/11/2026
    add_res = tracker.add_entry(
        module="Hành chính - Kế toán",
        code="ACC-2026-01",
        name="Đối chiếu bảng lương tháng 11",
        change_type="Tạo mới",
        rules="Tổng hợp SUMIFS theo phòng ban, kiểm tra định dạng CCCD text '@'.",
        ui="Xuất Excel sạch + Pivot Chart",
        status="Đã duyệt",
        notes="Sẵn sàng trình ký",
        date_str="14/11/2026",
    )
    assert add_res["ok"] is True
    assert add_res["stt"] == 3
    assert add_res["code"] == "ACC-2026-01"

    entries = tracker.list_entries(date_filter="14/11/2026")
    assert any(e["code"] == "ACC-2026-01" for e in entries)

    sync_res = tracker.sync_daily_report(
        date_str="14/11/2026",
        apply_to_task_tracking=True,
    )
    assert sync_res["ok"] is True
    assert sync_res["entries_matched"] >= 1
    assert sync_res["injected_to_task_file"] is True
    assert sync_res["target_sheet"] == "14-11"

    # Verify rows actually written into '14-11'
    wb_t_check = openpyxl.load_workbook(str(task_file))
    ws_1411 = wb_t_check["14-11"]
    descriptions = [str(ws_1411.cell(r, 3).value or "") for r in range(3, ws_1411.max_row + 1)]
    assert any("ACC-2026-01" in d for d in descriptions)
    wb_t_check.close()

    meta = tracker.inspect_workbooks()
    assert meta["ok"] is True
    assert meta["project_exists"] is True
    assert meta["task_exists"] is True
    assert "14-11" in meta["task_sheets"]
    assert meta["total_entries"] >= 3


def test_project_tracker_module_delegates_to_core_tracker() -> None:
    assert hasattr(project_tracker, "OfficeWorkTracker")
    assert project_tracker.OfficeWorkTracker is OfficeWorkTracker

