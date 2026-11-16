# -*- coding: utf-8 -*-
"""
tests/test_com_watchdog.py — Test Suite for Win32 COM Watchdog & Isolation.
"""

import time
import pytest

from core.converter.com_watchdog import ComProcessWatchdog


def test_watchdog_normal_execution():
    def fast_task(x: int, y: int) -> int:
        return x + y

    res = ComProcessWatchdog.run_with_timeout(
        fast_task, args=(10, 25), timeout_sec=3.0, app_type="word"
    )
    assert res == 35


def test_watchdog_exception_propagation():
    def failing_task():
        raise ValueError("Simulated COM Failure")

    with pytest.raises(ValueError, match="Simulated COM Failure"):
        ComProcessWatchdog.run_with_timeout(
            failing_task, timeout_sec=3.0, app_type="word"
        )


def test_watchdog_timeout_trigger():
    def hung_task():
        time.sleep(5.0)
        return "Finished"

    with pytest.raises(TimeoutError, match="hung and timed out"):
        ComProcessWatchdog.run_with_timeout(
            hung_task, timeout_sec=0.5, app_type="word"
        )


def test_watchdog_terminate_processes_safe():
    count = ComProcessWatchdog.terminate_hung_com_processes("all")
    assert isinstance(count, int)


def test_watchdog_automation_only_preserves_interactive_office(monkeypatch):
    monkeypatch.setattr("sys.platform", "win32")
    monkeypatch.setattr(
        ComProcessWatchdog,
        "_list_office_processes",
        classmethod(
            lambda cls, names=None: [
                (1001, "WINWORD.EXE", '"C:\\Office\\WINWORD.EXE" "hop_dong_quan_trong.docx"'),
                (2002, "EXCEL.EXE", '"C:\\Office\\EXCEL.EXE" /automation -Embedding'),
                (3003, "EXCEL.EXE", '"C:\\Office\\EXCEL.EXE" "bao_cao_tai_chinh.xlsx"'),
            ]
        ),
    )
    killed = []
    monkeypatch.setattr(
        ComProcessWatchdog,
        "_kill_pid",
        classmethod(lambda cls, pid: killed.append(pid) or True),
    )
    count = ComProcessWatchdog.terminate_hung_com_processes("all", automation_only=True)
    assert count == 1
    assert killed == [2002]


def test_watchdog_timeout_kills_only_tracked_worker_pid(monkeypatch):
    monkeypatch.setattr("sys.platform", "win32")
    monkeypatch.setattr(
        ComProcessWatchdog,
        "_list_office_processes",
        classmethod(lambda cls, names=None: [(9999, "WINWORD.EXE", '"C:\\Office\\WINWORD.EXE" "user_doc.docx"')]),
    )
    killed = []
    monkeypatch.setattr(
        ComProcessWatchdog,
        "_kill_pid",
        classmethod(lambda cls, pid: killed.append(pid) or True),
    )

    def hung_com_worker():
        ComProcessWatchdog.register_current_pid(4444)
        time.sleep(2.0)

    with pytest.raises(TimeoutError):
        ComProcessWatchdog.run_with_timeout(hung_com_worker, timeout_sec=0.3, app_type="word")

    assert killed == [4444]

