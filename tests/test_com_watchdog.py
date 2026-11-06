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

