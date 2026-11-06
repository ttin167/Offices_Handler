# -*- coding: utf-8 -*-
"""
tests/test_persistent_job_queue.py — Test Suite for Persistent SQLite Task Queue (v5.0).
"""

import os
import sqlite3
import tempfile
import time
import pytest

from server.job_queue import PersistentJobQueue


def test_sqlite_wal_mode_and_initialization():
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        db_path = os.path.join(tmp, "test_queue.db")
        queue = PersistentJobQueue(db_path=db_path)

        # Verify PRAGMA journal_mode is WAL
        conn = sqlite3.connect(db_path)
        cur = conn.execute("PRAGMA journal_mode;")
        mode = cur.fetchone()[0]
        assert mode.lower() == "wal"

        # Verify table exists
        cur = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='system_job_queue';")
        assert cur.fetchone() is not None
        conn.close()
        queue.close()


def test_job_submission_and_progress_tracking():
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        db_path = os.path.join(tmp, "test_queue.db")
        queue = PersistentJobQueue(db_path=db_path)

        def sample_runner(payload, progress_cb):
            progress_cb(50.0, "Đang xử lý 50%...")
            time.sleep(0.05)
            return {"output": "SUCCESS", "records_processed": 100}

        job_id = queue.submit_job(
            action="data_conversion",
            payload={"input": "test.pdf", "format": "docx"},
            runner_fn=sample_runner
        )
        assert job_id.startswith("job_")

        # Wait for worker thread to complete
        for _ in range(50):
            job = queue.get_job(job_id)
            if job and job["status"] in ("completed", "failed"):
                break
            time.sleep(0.05)

        job_done = queue.get_job(job_id)
        assert job_done is not None
        assert job_done["status"] == "completed"
        assert job_done["progress"] == 100.0
        assert job_done["result"] == {"output": "SUCCESS", "records_processed": 100}
        assert job_done["error"] is None
        queue.close()


def test_crash_recovery_sweeper():
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        db_path = os.path.join(tmp, "test_recovery.db")

        # 1. Simulate an interrupted job left in 'running' state from previous crash
        conn = sqlite3.connect(db_path)
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.execute("""
            CREATE TABLE system_job_queue (
                job_id TEXT PRIMARY KEY,
                action TEXT NOT NULL,
                status TEXT NOT NULL,
                progress REAL DEFAULT 0.0,
                message TEXT,
                payload_json TEXT,
                result_json TEXT,
                error_msg TEXT,
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL
            );
        """)
        now = time.time()
        conn.execute("""
            INSERT INTO system_job_queue
            (job_id, action, status, progress, message, created_at, updated_at)
            VALUES ('job_crashed_1', 'ocr_scan', 'running', 45.0, 'Đang nhận dạng trang 25...', ?, ?);
        """, (now, now))
        conn.commit()
        conn.close()

        # 2. Start PersistentJobQueue - must trigger recover_stale_jobs automatically
        queue = PersistentJobQueue(db_path=db_path)

        job = queue.get_job("job_crashed_1")
        assert job is not None
        assert job["status"] == "failed"
        assert job["error"] == "SERVER_RESTARTED"
        assert "Crash Recovery" in job["message"]
        queue.close()


def test_sse_subscriber_events():
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        db_path = os.path.join(tmp, "test_sse.db")
        queue = PersistentJobQueue(db_path=db_path)

        def runner(payload, progress_cb):
            progress_cb(30.0, "Bước 1")
            progress_cb(75.0, "Bước 2")
            return "DONE"

        job_id = queue.submit_job("test_action", {}, runner_fn=runner)
        sub_queue = queue.subscribe(job_id)

        # Collect events from subscription queue
        events = []
        for _ in range(50):
            try:
                ev = sub_queue.get(timeout=0.1)
                events.append(ev)
                if ev.get("status") == "completed":
                    break
            except Exception:
                pass

        assert len(events) >= 2
        statuses = [e["status"] for e in events]
        assert "completed" in statuses
        queue.close()
