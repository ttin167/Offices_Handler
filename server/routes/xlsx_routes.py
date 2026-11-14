# -*- coding: utf-8 -*-
"""
server.routes.xlsx_routes — Pure Office Excel Power Suite & Tracker REST API Endpoints.
"""

from __future__ import annotations

from typing import Any, List, Optional

from .base import (
    build_output_filepath,
    download_url_for,
    read_json_body,
    resolve_workspace_file,
    router,
    send_error,
    send_json,
)


def _split_col_list(raw: Any) -> List[str]:
    """Normalize either a comma-separated string or list of strings into a clean list."""
    if not raw:
        return []
    if isinstance(raw, str):
        return [item.strip() for item in raw.split(",") if item.strip()]
    if isinstance(raw, (list, tuple)):
        return [str(item).strip() for item in raw if str(item).strip()]
    return []


def _resolve_required_input(body: dict, keys: tuple[str, ...] = ("file_path", "filepath", "filename", "input_path")) -> Optional[str]:
    """Extract and sandbox-resolve the first matching input file field from body."""
    for key in keys:
        val = body.get(key)
        if val:
            resolved = resolve_workspace_file(str(val))
            if resolved:
                return resolved
    return None


@router.post("/api/xlsx/power-clean")
def handle_xlsx_power_clean(handler: Any) -> None:
    """Standardize spaces, VN dates, phone/CCCD '@' text formats, and remove duplicates."""
    try:
        from core.xlsx.power_tools import clean_and_standardize_workbook

        body = read_json_body(handler)
        src = _resolve_required_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Excel nguồn hợp lệ.", status=400)
            return

        out_name, out_path = build_output_filepath(src, "cleaned", ext=".xlsx")
        dedup_cols = _split_col_list(body.get("dedup_columns")) or None
        res = clean_and_standardize_workbook(
            input_path=src,
            output_path=out_path,
            sheet_name=body.get("sheet_name") or None,
            dedup_columns=dedup_cols,
            normalize_dates=bool(body.get("normalize_dates", True)),
            normalize_phones_and_ids=bool(body.get("normalize_phones_and_ids", True)),
            trim_whitespace=bool(body.get("trim_whitespace", True)),
            case_mode=str(body.get("case_mode", "preserve")),
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi chuẩn hóa dữ liệu Excel: {exc}", status=500)


@router.post("/api/xlsx/transform-columns")
def handle_xlsx_transform_columns(handler: Any) -> None:
    """Split Vietnamese full names, split by delimiter, or merge multiple columns."""
    try:
        from core.xlsx.power_tools import transform_workbook_columns

        body = read_json_body(handler)
        src = _resolve_required_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Excel nguồn hợp lệ.", status=400)
            return

        mode = str(body.get("mode", "split_vn_fullname")).strip()
        src_cols = _split_col_list(body.get("source_columns") or body.get("source_column"))
        if not src_cols:
            send_error(handler, "Thiếu danh sách cột nguồn (`source_columns`).", status=400)
            return

        target_cols = _split_col_list(body.get("target_columns")) or None
        out_name, out_path = build_output_filepath(src, f"cols_{mode}", ext=".xlsx")
        res = transform_workbook_columns(
            input_path=src,
            mode=mode,
            source_columns=src_cols,
            output_path=out_path,
            sheet_name=body.get("sheet_name") or None,
            delimiter=str(body.get("delimiter", " ")),
            joiner=str(body.get("joiner", " ")),
            target_columns=target_cols,
            new_column_name=body.get("new_column_name") or None,
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi tách/gộp cột Excel: {exc}", status=500)


@router.post("/api/xlsx/pivot-chart")
def handle_xlsx_pivot_chart(handler: Any) -> None:
    """Generate a live Pivot_Summary sheet with embedded native Bar/Line/Pie chart."""
    try:
        from core.xlsx.power_tools import build_pivot_summary_and_chart

        body = read_json_body(handler)
        src = _resolve_required_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Excel nguồn hợp lệ.", status=400)
            return

        group_by = str(body.get("group_by_col", "")).strip()
        if not group_by:
            send_error(handler, "Thiếu cột phân nhóm Pivot (`group_by_col`).", status=400)
            return

        val_cols = _split_col_list(body.get("value_cols")) or None
        out_name, out_path = build_output_filepath(src, "pivot_chart", ext=".xlsx")
        res = build_pivot_summary_and_chart(
            input_path=src,
            group_by_col=group_by,
            value_cols=val_cols,
            agg_func=str(body.get("agg_func", "SUM")),
            chart_type=str(body.get("chart_type", "bar")),
            output_path=out_path,
            source_sheet=body.get("source_sheet") or None,
            summary_sheet_name=str(body.get("summary_sheet_name", "Pivot_Summary")),
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi tạo Pivot & Biểu đồ: {exc}", status=500)


@router.post("/api/xlsx/xlookup-join")
def handle_xlsx_xlookup_join(handler: Any) -> None:
    """No-code XLOOKUP join between two sheets or two workbooks."""
    try:
        from core.xlsx.power_tools import xlookup_join_workbooks

        body = read_json_body(handler)
        left_src = _resolve_required_input(body, keys=("left_path", "file_path", "filepath"))
        if not left_src:
            send_error(handler, "Thiếu hoặc không tìm thấy bảng chính (`left_path`).", status=400)
            return

        right_raw = body.get("right_path")
        right_src = resolve_workspace_file(str(right_raw)) if right_raw else left_src
        if not right_src:
            send_error(handler, "Không tìm thấy bảng tra cứu (`right_path`).", status=400)
            return

        left_key = str(body.get("left_key", "")).strip()
        right_key = str(body.get("right_key", "")).strip() or left_key
        ret_cols = _split_col_list(body.get("return_cols"))
        if not left_key or not ret_cols:
            send_error(handler, "Thiếu khóa dò tìm (`left_key`) hoặc cột cần lấy (`return_cols`).", status=400)
            return

        out_name, out_path = build_output_filepath(left_src, "joined", ext=".xlsx")
        res = xlookup_join_workbooks(
            left_path=left_src,
            right_path=right_src,
            left_key=left_key,
            right_key=right_key,
            return_cols=ret_cols,
            output_path=out_path,
            left_sheet=body.get("left_sheet") or None,
            right_sheet=body.get("right_sheet") or None,
            not_found_value=str(body.get("not_found_value", "N/A")),
            formula_mode=str(body.get("formula_mode", "auto")),
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi nối bảng XLOOKUP Join: {exc}", status=500)


@router.post("/api/xlsx/split-by-column")
def handle_xlsx_split_by_column(handler: Any) -> None:
    """Split a master Excel sheet by unique column values into separate XLSX files (ZIP)."""
    try:
        from core.xlsx.power_tools import split_workbook_by_column

        body = read_json_body(handler)
        src = _resolve_required_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Excel nguồn.", status=400)
            return

        split_col = str(body.get("split_col", "")).strip()
        if not split_col:
            send_error(handler, "Thiếu tên cột phân tách (`split_col`).", status=400)
            return

        out_name, out_path = build_output_filepath(src, "split", ext=".zip")
        res = split_workbook_by_column(
            input_path=src,
            split_col=split_col,
            output_path=out_path,
            sheet_name=body.get("sheet_name") or None,
            as_zip=True,
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi tách workbook theo cột: {exc}", status=500)


@router.post("/api/xlsx/diff")
def handle_xlsx_diff(handler: Any) -> None:
    """Visual cell-by-cell Excel Diff with Amber/Emerald/Rose highlighting and _Diff_Summary."""
    try:
        from core.xlsx.power_tools import diff_workbooks_visual

        body = read_json_body(handler)
        v1 = _resolve_required_input(body, keys=("v1_path", "file1_path"))
        v2 = _resolve_required_input(body, keys=("v2_path", "file2_path"))
        if not v1 or not v2:
            send_error(handler, "Cần cung cấp đầy đủ 2 tệp Excel v1 và v2 hợp lệ.", status=400)
            return

        out_name, out_path = build_output_filepath(v2, "diff", ext=".xlsx")
        sheets = _split_col_list(body.get("sheets")) or None
        res = diff_workbooks_visual(
            v1_path=v1,
            v2_path=v2,
            output_path=out_path,
            sheets=sheets,
        )
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "download_url": download_url_for(out_name),
                **res,
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi đối chiếu Excel Diff: {exc}", status=500)


@router.post("/api/xlsx/eval")
def handle_xlsx_eval(handler: Any) -> None:
    """Evaluate Excel formulas offline via ASTFormulaEvaluator and cache <v> values."""
    try:
        from core.xlsx.formula_evaluator import ASTFormulaEvaluator

        body = read_json_body(handler)
        src = _resolve_required_input(body)
        if not src:
            send_error(handler, "Thiếu hoặc không tìm thấy tệp Excel cần tính công thức.", status=400)
            return

        out_name, out_path = build_output_filepath(src, "evaluated", ext=".xlsx")
        final_path = ASTFormulaEvaluator.populate_cached_values(src, output_path=out_path)
        send_json(
            handler,
            {
                "success": True,
                "filename": out_name,
                "output_path": final_path,
                "download_url": download_url_for(out_name),
            },
        )
    except Exception as exc:
        send_error(handler, f"Lỗi tính toán công thức offline (AST Eval): {exc}", status=500)


@router.get("/api/xlsx/tracker/inspect")
def handle_xlsx_tracker_inspect(handler: Any) -> None:
    """Inspect work tracking & daily task workbooks status and recent entries."""
    try:
        from core.xlsx.tracker import OfficeWorkTracker

        res = OfficeWorkTracker().inspect_workbooks()
        send_json(handler, {"success": True, **res})
    except Exception as exc:
        send_error(handler, f"Lỗi đọc sổ theo dõi công việc: {exc}", status=500)


@router.post("/api/xlsx/tracker/init")
def handle_xlsx_tracker_init(handler: Any) -> None:
    """Initialize fresh project_tracking.xlsx workbook if absent (or force overwrite)."""
    try:
        from core.xlsx.tracker import OfficeWorkTracker

        body = read_json_body(handler)
        res = OfficeWorkTracker().init_project_tracking(force=bool(body.get("force", False)))
        send_json(handler, {"success": True, **res})
    except Exception as exc:
        send_error(handler, f"Lỗi khởi tạo sổ theo dõi: {exc}", status=500)


@router.post("/api/xlsx/tracker/add")
def handle_xlsx_tracker_add(handler: Any) -> None:
    """Append a work record to project_tracking.xlsx."""
    try:
        from core.xlsx.tracker import OfficeWorkTracker

        body = read_json_body(handler)
        module = str(body.get("module", "")).strip()
        code = str(body.get("code", "")).strip()
        name = str(body.get("name", "")).strip()
        rules = str(body.get("rules", "")).strip()
        if not (module and code and name and rules):
            send_error(handler, "Cần nhập đủ Bộ phận (module), Mã (code), Tên việc (name) và Nội dung (rules).", status=400)
            return

        res = OfficeWorkTracker().add_entry(
            module=module,
            code=code,
            name=name,
            change_type=str(body.get("change_type", "Tạo mới")),
            rules=rules,
            ui=str(body.get("ui", "")),
            status=str(body.get("status", "Đang soạn thảo")),
            notes=str(body.get("notes", "")),
            date_str=body.get("date_str") or None,
        )
        send_json(handler, {"success": True, **res})
    except Exception as exc:
        send_error(handler, f"Lỗi thêm dòng nhật ký: {exc}", status=500)


@router.post("/api/xlsx/tracker/sync-daily")
def handle_xlsx_tracker_sync_daily(handler: Any) -> None:
    """Synthesize daily task summary and optionally inject into task_tracking.xlsx."""
    try:
        from core.xlsx.tracker import OfficeWorkTracker

        body = read_json_body(handler)
        res = OfficeWorkTracker().sync_daily_report(
            date_str=body.get("date_str") or None,
            apply_to_task_tracking=bool(body.get("apply_to_task_tracking", False)),
        )
        send_json(handler, {"success": True, **res})
    except Exception as exc:
        send_error(handler, f"Lỗi đồng bộ báo cáo ngày: {exc}", status=500)

