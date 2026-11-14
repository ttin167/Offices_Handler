"""
tools/project_tracker.py — CLI wrapper for Office Work & Daily Task Tracker.

Delegates all workbook generation, schema validation, and daily report syncing
to `core.xlsx.tracker.OfficeWorkTracker` so there is zero duplicated logic.
"""

from __future__ import annotations

import argparse
import io
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Ensure UTF-8 printing on Windows consoles
if sys.stdout and hasattr(sys.stdout, "buffer") and sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from core.xlsx.tracker import (  # noqa: E402
    DEFAULT_PROJECT_FILE as PROJECT_FILE,
    DEFAULT_TASK_FILE as TASK_FILE,
    OfficeWorkTracker,
)


def init_project_tracking(file_path: Path = PROJECT_FILE, force: bool = False) -> Path:
    """Create or re-initialize the master work log workbook."""
    res = OfficeWorkTracker(project_file=file_path).init_project_tracking(force=force)
    print(f"[{'OK' if res['created'] else 'INFO'}] {res['message']} -> {res['path']}")
    return Path(res["path"])


def add_project_entry(
    module: str,
    code: str,
    name: str,
    change_type: str,
    rules: str,
    ui: str = "",
    status: str = "Đang soạn thảo",
    notes: str = "",
    date_str: Optional[str] = None,
    file_path: Path = PROJECT_FILE,
) -> int:
    """Append a new tracked work entry and return its STT."""
    res = OfficeWorkTracker(project_file=file_path).add_entry(
        module=module,
        code=code,
        name=name,
        change_type=change_type,
        rules=rules,
        ui=ui,
        status=status,
        notes=notes,
        date_str=date_str,
    )
    print(f"[OK] Added STT {res['stt']} [{code}] '{name}' to row {res['row']} in {file_path.name}")
    return int(res["stt"])


def list_project_entries(
    date_filter: Optional[str] = None,
    module_filter: Optional[str] = None,
    status_filter: Optional[str] = None,
    file_path: Path = PROJECT_FILE,
    as_json: bool = False,
) -> List[Dict[str, Any]]:
    """List work log records with optional filtering."""
    records = OfficeWorkTracker(project_file=file_path).list_entries(
        date_filter=date_filter,
        module_filter=module_filter,
        status_filter=status_filter,
    )
    if as_json:
        print(json.dumps(records, ensure_ascii=False, indent=2))
        return records

    print(f"\n=== OFFICE WORK TRACKING LOG ({len(records)} records) ===")
    for item in records:
        print(
            f"[{item['stt']}] {item['date']} | {item['module']} | "
            f"{item['code']} - {item['name']} | {item['type']} | {item['status']}"
        )
        print(f"    Rules: {item['rules']}")
        if item["ui"]:
            print(f"    Output/UI: {item['ui']}")
        if item["notes"]:
            print(f"    Notes: {item['notes']}")
    print("=" * 60 + "\n")
    return records


def sync_to_daily_report(
    date_str: Optional[str] = None,
    apply_to_task_tracking: bool = False,
    project_file: Path = PROJECT_FILE,
    task_file: Path = TASK_FILE,
) -> None:
    """Synthesize daily tasks from work log and optionally inject into task_tracking.xlsx."""
    res = OfficeWorkTracker(project_file=project_file, task_file=task_file).sync_daily_report(
        date_str=date_str,
        apply_to_task_tracking=apply_to_task_tracking,
    )
    print(f"\n=== TỔNG HỢP BÁO CÁO NGÀY ({res['date']}) ===")
    for item in res["tasks"]:
        print(f"  {item['stt']}. [{item['module']}] {item['description']} -> {item['status']}")
    if res["injected_to_task_file"]:
        print(f"[OK] Injected {len(res['tasks'])} tasks into sheet '{res['target_sheet']}' of {task_file.name}")


def main() -> None:
    """CLI entrypoint for tools/project_tracker.py."""
    parser = argparse.ArgumentParser(
        description="Office Work & Daily Task Tracker CLI (OpenPyXL Zero-COM)"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    p_init = subparsers.add_parser("init", help="Initialize fresh workbook if not exists")
    p_init.add_argument("--force", action="store_true", help="Overwrite existing file")
    p_init.add_argument("--file", type=str, default=str(PROJECT_FILE), help="Target file path")

    p_add = subparsers.add_parser("add", help="Add a new work change record")
    p_add.add_argument("--module", required=True, help="Module / Bộ phận")
    p_add.add_argument("--code", required=True, help="Mã đầu việc / Tài liệu")
    p_add.add_argument("--name", required=True, help="Tên đầu việc / Biểu mẫu")
    p_add.add_argument("--type", required=True, choices=["Tạo mới", "Bổ sung", "Chỉnh sửa", "Hủy bỏ"], help="Phân loại thay đổi")
    p_add.add_argument("--rules", required=True, help="Quy tắc & nội dung chính")
    p_add.add_argument("--ui", default="", help="Đầu ra / Biểu mẫu liên quan")
    p_add.add_argument("--status", default="Đang soạn thảo", choices=["Đang soạn thảo", "Chờ Review", "Đã duyệt", "Đã chốt"], help="Trạng thái")
    p_add.add_argument("--notes", default="", help="Ghi chú")
    p_add.add_argument("--date", default=None, help="Ngày (DD/MM/YYYY), mặc định hôm nay")
    p_add.add_argument("--file", type=str, default=str(PROJECT_FILE), help="Target file path")

    p_list = subparsers.add_parser("list", help="List recorded entries")
    p_list.add_argument("--date", default=None, help="Filter by date (DD/MM/YYYY)")
    p_list.add_argument("--module", default=None, help="Filter by module substring")
    p_list.add_argument("--status", default=None, help="Filter by status")
    p_list.add_argument("--json", action="store_true", help="Output as JSON")
    p_list.add_argument("--file", type=str, default=str(PROJECT_FILE), help="Target file path")

    p_sync = subparsers.add_parser("sync-daily", help="Synthesize daily report from recorded entries")
    p_sync.add_argument("--date", default=None, help="Target date (DD/MM/YYYY), defaults to today")
    p_sync.add_argument("--apply", action="store_true", help="Auto-inject tasks into task_tracking.xlsx")
    p_sync.add_argument("--project-file", type=str, default=str(PROJECT_FILE), help="Source work tracking file")
    p_sync.add_argument("--task-file", type=str, default=str(TASK_FILE), help="Destination task tracking file")

    args = parser.parse_args()
    if args.command == "init":
        init_project_tracking(Path(args.file), force=args.force)
    elif args.command == "add":
        add_project_entry(
            module=args.module,
            code=args.code,
            name=args.name,
            change_type=args.type,
            rules=args.rules,
            ui=args.ui,
            status=args.status,
            notes=args.notes,
            date_str=args.date,
            file_path=Path(args.file),
        )
    elif args.command == "list":
        list_project_entries(
            date_filter=args.date,
            module_filter=args.module,
            status_filter=args.status,
            file_path=Path(args.file),
            as_json=args.json,
        )
    elif args.command == "sync-daily":
        sync_to_daily_report(
            date_str=args.date,
            apply_to_task_tracking=args.apply,
            project_file=Path(args.project_file),
            task_file=Path(args.task_file),
        )


if __name__ == "__main__":
    main()
