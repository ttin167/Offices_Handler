"""
cli/office_power_commands.py — Pure Office Excel & Word Power Suite CLI Subcommands.

Exposes 11 high-frequency everyday Office IT operations:
- Excel: xlsx-power-clean, xlsx-transform-cols, xlsx-pivot-chart,
  xlsx-xlookup-join, xlsx-split-col, xlsx-diff, xlsx-track
- Word: docx-polish, docx-merge, docx-split, docx-watermark
"""

from __future__ import annotations

import argparse
import json
import sys
from typing import Any, Dict


def _print_result(res: Dict[str, Any], as_json: bool) -> int:
    if as_json:
        print(json.dumps(res, ensure_ascii=False, indent=2, default=str))
    else:
        for k, v in res.items():
            if isinstance(v, (list, dict)) and len(str(v)) > 140:
                print(f"  {k}: ({len(v)} items)")
            else:
                print(f"  {k}: {v}")
    return 0


def cmd_xlsx_power_clean(args: argparse.Namespace) -> int:
    from core.xlsx.power_tools import clean_and_standardize_workbook

    try:
        dedup = [c.strip() for c in args.dedup.split(",") if c.strip()] if args.dedup else None
        res = clean_and_standardize_workbook(
            input_path=args.file,
            output_path=args.output,
            sheet_name=args.sheet,
            dedup_columns=dedup,
            normalize_dates=not args.no_dates,
            normalize_phones_and_ids=not args.no_ids,
            trim_whitespace=not args.no_trim,
            case_mode=args.case_mode,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] xlsx-power-clean failed: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_transform_cols(args: argparse.Namespace) -> int:
    from core.xlsx.power_tools import transform_workbook_columns

    try:
        src_cols = [c.strip() for c in args.columns.split(",") if c.strip()]
        tgt_cols = [c.strip() for c in args.target_columns.split(",") if c.strip()] if args.target_columns else None
        res = transform_workbook_columns(
            input_path=args.file,
            mode=args.mode,
            source_columns=src_cols,
            output_path=args.output,
            sheet_name=args.sheet,
            delimiter=args.delimiter,
            joiner=args.delimiter,
            target_columns=tgt_cols,
            new_column_name=args.new_column,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] xlsx-transform-cols failed: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_pivot_chart(args: argparse.Namespace) -> int:
    from core.xlsx.power_tools import build_pivot_summary_and_chart

    try:
        val_cols = [c.strip() for c in args.values.split(",") if c.strip()] if args.values else None
        res = build_pivot_summary_and_chart(
            input_path=args.file,
            group_by_col=args.group_by,
            value_cols=val_cols,
            agg_func=args.agg,
            chart_type=args.chart,
            output_path=args.output,
            source_sheet=args.sheet,
            summary_sheet_name=args.summary_sheet,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] xlsx-pivot-chart failed: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_xlookup_join(args: argparse.Namespace) -> int:
    from core.xlsx.power_tools import xlookup_join_workbooks

    try:
        ret_cols = [c.strip() for c in args.return_cols.split(",") if c.strip()]
        res = xlookup_join_workbooks(
            left_path=args.left_file,
            right_path=args.right_file,
            left_key=args.left_key,
            right_key=args.right_key or args.left_key,
            return_cols=ret_cols,
            output_path=args.output,
            left_sheet=args.left_sheet,
            right_sheet=args.right_sheet,
            not_found_value=args.not_found,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] xlsx-xlookup-join failed: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_split_col(args: argparse.Namespace) -> int:
    from core.xlsx.power_tools import split_workbook_by_column

    try:
        res = split_workbook_by_column(
            input_path=args.file,
            split_col=args.column,
            output_path=args.output,
            sheet_name=args.sheet,
            as_zip=not args.no_zip,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] xlsx-split-col failed: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_diff(args: argparse.Namespace) -> int:
    from core.xlsx.power_tools import diff_workbooks_visual

    try:
        sheets = [s.strip() for s in args.sheets.split(",") if s.strip()] if args.sheets else None
        res = diff_workbooks_visual(
            v1_path=args.v1,
            v2_path=args.v2,
            output_path=args.output,
            sheets=sheets,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] xlsx-diff failed: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_track(args: argparse.Namespace) -> int:
    from core.xlsx.tracker import OfficeWorkTracker

    try:
        tracker = OfficeWorkTracker(project_file=args.project_file, task_file=args.task_file)
        if args.action == "init":
            res = tracker.init_project_tracking(force=args.force)
        elif args.action == "add":
            if not (args.module and args.code and args.name and args.rules):
                raise ValueError("add requires --module, --code, --name, --rules")
            res = tracker.add_entry(
                module=args.module,
                code=args.code,
                name=args.name,
                change_type=args.change_type,
                rules=args.rules,
                ui=args.ui or "",
                status=args.status,
                notes=args.notes or "",
                date_str=args.date,
            )
        elif args.action == "sync-daily":
            res = tracker.sync_daily_report(date_str=args.date, apply_to_task_tracking=args.apply)
        else:
            res = {"ok": True, "entries": tracker.list_entries(date_filter=args.date, module_filter=args.module)}
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] xlsx-track failed: {exc}", file=sys.stderr)
        return 2


def cmd_docx_polish(args: argparse.Namespace) -> int:
    from core.docx.toolkit import polish_docx_document

    try:
        res = polish_docx_document(
            input_path=args.file,
            output_path=args.output,
            preset=args.preset,
            add_page_numbers=not args.no_page_numbers,
            sanitize_typography=not args.no_typo,
            fix_tables=not args.no_fix_tables,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] docx-polish failed: {exc}", file=sys.stderr)
        return 2


def cmd_docx_merge(args: argparse.Namespace) -> int:
    from core.docx.toolkit import merge_docx_files

    try:
        res = merge_docx_files(
            input_paths=args.files,
            output_path=args.output,
            add_page_breaks=not args.no_page_breaks,
            normalize_page_numbers=not args.no_page_numbers,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] docx-merge failed: {exc}", file=sys.stderr)
        return 2


def cmd_docx_split(args: argparse.Namespace) -> int:
    from core.docx.toolkit import split_docx_by_heading

    try:
        res = split_docx_by_heading(
            input_path=args.file,
            output_path=args.output,
            heading_level=args.heading_level,
            as_zip=not args.no_zip,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] docx-split failed: {exc}", file=sys.stderr)
        return 2


def cmd_docx_watermark(args: argparse.Namespace) -> int:
    from core.docx.toolkit import apply_docx_watermark_and_hf

    try:
        res = apply_docx_watermark_and_hf(
            input_path=args.file,
            output_path=args.output,
            watermark_text=args.text,
            header_text=args.header,
            footer_text=args.footer,
            add_page_numbers=not args.no_page_numbers,
        )
        return _print_result(res, args.json)
    except Exception as exc:
        print(f"[ERROR] docx-watermark failed: {exc}", file=sys.stderr)
        return 2


def register_office_power_subcommands(sub: argparse._SubParsersAction) -> None:
    """Register Pure Office Excel & Word Power Suite subcommands."""
    p_clean = sub.add_parser("xlsx-power-clean", help="Clean spaces, standardize VN dates, protect phones/IDs '@', dedup rows")
    p_clean.add_argument("file", help="Input Excel .xlsx file")
    p_clean.add_argument("-o", "--output", default=None)
    p_clean.add_argument("--sheet", default=None)
    p_clean.add_argument("--dedup", default=None, help="Comma-separated columns for dedup, or '*'")
    p_clean.add_argument("--case-mode", default="preserve", choices=["preserve", "vn_name", "upper", "lower"])
    p_clean.add_argument("--no-dates", action="store_true")
    p_clean.add_argument("--no-ids", action="store_true")
    p_clean.add_argument("--no-trim", action="store_true")
    p_clean.add_argument("--json", action="store_true")
    p_clean.set_defaults(func=cmd_xlsx_power_clean)

    p_tr = sub.add_parser("xlsx-transform-cols", help="Split VN full names, split by delimiter, or merge Excel columns")
    p_tr.add_argument("file")
    p_tr.add_argument("--mode", required=True, choices=["split_vn_fullname", "split_delimiter", "merge_columns"])
    p_tr.add_argument("--columns", required=True, help="Comma-separated source column names/letters")
    p_tr.add_argument("--delimiter", default=" ")
    p_tr.add_argument("--target-columns", default=None)
    p_tr.add_argument("--new-column", default=None)
    p_tr.add_argument("--sheet", default=None)
    p_tr.add_argument("-o", "--output", default=None)
    p_tr.add_argument("--json", action="store_true")
    p_tr.set_defaults(func=cmd_xlsx_transform_cols)

    p_piv = sub.add_parser("xlsx-pivot-chart", help="Create live formula Pivot_Summary sheet + embedded native chart")
    p_piv.add_argument("file")
    p_piv.add_argument("--group-by", required=True)
    p_piv.add_argument("--values", default=None)
    p_piv.add_argument("--agg", default="SUM", choices=["SUM", "COUNT", "AVERAGE"])
    p_piv.add_argument("--chart", default="bar", choices=["bar", "line", "pie"])
    p_piv.add_argument("--sheet", default=None)
    p_piv.add_argument("--summary-sheet", default="Pivot_Summary")
    p_piv.add_argument("-o", "--output", default=None)
    p_piv.add_argument("--json", action="store_true")
    p_piv.set_defaults(func=cmd_xlsx_pivot_chart)

    p_xlk = sub.add_parser("xlsx-xlookup-join", help="No-code XLOOKUP join between sheets or workbooks")
    p_xlk.add_argument("left_file")
    p_xlk.add_argument("--right-file", default=None)
    p_xlk.add_argument("--left-key", required=True)
    p_xlk.add_argument("--right-key", default=None)
    p_xlk.add_argument("--return-cols", required=True)
    p_xlk.add_argument("--left-sheet", default=None)
    p_xlk.add_argument("--right-sheet", default=None)
    p_xlk.add_argument("--not-found", default="N/A")
    p_xlk.add_argument("-o", "--output", default=None)
    p_xlk.add_argument("--json", action="store_true")
    p_xlk.set_defaults(func=cmd_xlsx_xlookup_join)

    p_spl = sub.add_parser("xlsx-split-col", help="Split Excel workbook by unique column values into separate files / ZIP")
    p_spl.add_argument("file")
    p_spl.add_argument("--column", required=True)
    p_spl.add_argument("--sheet", default=None)
    p_spl.add_argument("-o", "--output", default=None)
    p_spl.add_argument("--no-zip", action="store_true")
    p_spl.add_argument("--json", action="store_true")
    p_spl.set_defaults(func=cmd_xlsx_split_col)

    p_diff = sub.add_parser("xlsx-diff", help="Visual cell-by-cell Excel comparison with Comments and _Diff_Summary")
    p_diff.add_argument("v1")
    p_diff.add_argument("v2")
    p_diff.add_argument("--sheets", default=None)
    p_diff.add_argument("-o", "--output", default=None)
    p_diff.add_argument("--json", action="store_true")
    p_diff.set_defaults(func=cmd_xlsx_diff)

    p_trk = sub.add_parser("xlsx-track", help="Manage Office Work Log and Daily Task Tracker workbooks")
    p_trk.add_argument("action", choices=["init", "add", "list", "sync-daily"])
    p_trk.add_argument("--project-file", default="project_tracking.xlsx")
    p_trk.add_argument("--task-file", default="task_tracking.xlsx")
    p_trk.add_argument("--force", action="store_true")
    p_trk.add_argument("--module", default=None)
    p_trk.add_argument("--code", default=None)
    p_trk.add_argument("--name", default=None)
    p_trk.add_argument("--change-type", default="Tạo mới")
    p_trk.add_argument("--rules", default=None)
    p_trk.add_argument("--ui", default="")
    p_trk.add_argument("--status", default="Đang soạn thảo")
    p_trk.add_argument("--notes", default="")
    p_trk.add_argument("--date", default=None)
    p_trk.add_argument("--apply", action="store_true")
    p_trk.add_argument("--json", action="store_true")
    p_trk.set_defaults(func=cmd_xlsx_track)

    p_dpol = sub.add_parser("docx-polish", help="Polish Word DOCX with NĐ 30 / Executive presets, table fixes & Page X/Y")
    p_dpol.add_argument("file")
    p_dpol.add_argument("-o", "--output", default=None)
    p_dpol.add_argument("--preset", default="nd30")
    p_dpol.add_argument("--no-page-numbers", action="store_true")
    p_dpol.add_argument("--no-typo", action="store_true")
    p_dpol.add_argument("--no-fix-tables", action="store_true")
    p_dpol.add_argument("--json", action="store_true")
    p_dpol.set_defaults(func=cmd_docx_polish)

    p_dmrg = sub.add_parser("docx-merge", help="Merge multiple Word .docx files preserving media and tables")
    p_dmrg.add_argument("files", nargs="+")
    p_dmrg.add_argument("-o", "--output", required=True)
    p_dmrg.add_argument("--no-page-breaks", action="store_true")
    p_dmrg.add_argument("--no-page-numbers", action="store_true")
    p_dmrg.add_argument("--json", action="store_true")
    p_dmrg.set_defaults(func=cmd_docx_merge)

    p_dspl = sub.add_parser("docx-split", help="Split a Word .docx file by Heading level into chapter files / ZIP")
    p_dspl.add_argument("file")
    p_dspl.add_argument("-o", "--output", default=None)
    p_dspl.add_argument("--heading-level", type=int, default=1)
    p_dspl.add_argument("--no-zip", action="store_true")
    p_dspl.add_argument("--json", action="store_true")
    p_dspl.set_defaults(func=cmd_docx_split)

    p_dwm = sub.add_parser("docx-watermark", help="Apply diagonal VML watermark, Header/Footer, and Page X/Y to DOCX")
    p_dwm.add_argument("file")
    p_dwm.add_argument("-o", "--output", default=None)
    p_dwm.add_argument("--text", default="BẢN NHÁP")
    p_dwm.add_argument("--header", default=None)
    p_dwm.add_argument("--footer", default=None)
    p_dwm.add_argument("--no-page-numbers", action="store_true")
    p_dwm.add_argument("--json", action="store_true")
    p_dwm.set_defaults(func=cmd_docx_watermark)

