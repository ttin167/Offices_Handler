# -*- coding: utf-8 -*-
"""
server.job_queue — Persistent SQLite Task Queue & Crash Recovery Engine.

Conforms to Antigravity v5.0 Enterprise Architecture:
- Guard G1: PRAGMA journal_mode = WAL; PRAGMA busy_timeout = 5000;
- Exponential backoff retry on database locks.
- Automatic Crash Recovery Sweeper (recover_stale_jobs) on server startup.
- Thread-safe submission, progress tracking, and SSE event streaming.
"""

from __future__ import annotations

import concurrent.futures
import json
import os
import queue
import sqlite3
import sys
import threading
import time
import uuid
from dataclasses import asdict, dataclass, field
from typing import Any, Callable, Dict, List, Optional, Tuple


@dataclass
class JobRecord:
    job_id: str
    action: str
    status: str  # pending, running, completed, failed
    progress: float = 0.0
    message: str = ""
    payload: Dict[str, Any] = field(default_factory=dict)
    result: Optional[Any] = None
    error: Optional[str] = None
    created_at: float = field(default_factory=time.time)
    updated_at: float = field(default_factory=time.time)


class PersistentJobQueue:
    """Manages asynchronous background jobs with persistent SQLite WAL backing."""

    _instance: Optional[PersistentJobQueue] = None

    def __init__(self, db_path: Optional[str] = None, max_workers: Optional[int] = None) -> None:
        if db_path is None:
            db_path = os.path.abspath(os.path.join(os.getcwd(), "history.db"))
        self.db_path = db_path

        if max_workers is None:
            try:
                max_workers = max(1, int(os.environ.get("MAX_JOB_WORKERS", 4)))
            except (ValueError, TypeError):
                max_workers = 4

        self._max_workers = max_workers
        self._executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=max_workers,
            thread_name_prefix="PersistentJobWorker"
        )
        self._lock = threading.Lock()
        self._subscribers: Dict[str, List[queue.Queue]] = {}

        self._init_database()
        self.recover_stale_jobs()

    @classmethod
    def get_instance(cls, db_path: Optional[str] = None) -> PersistentJobQueue:
        if cls._instance is None:
            cls._instance = PersistentJobQueue(db_path=db_path)
        return cls._instance

    @classmethod
    def reset(cls) -> None:
        if cls._instance is not None:
            cls._instance.close()
            cls._instance = None

    def close(self) -> None:
        """Shuts down executor and checkpoints WAL database cleanly."""
        try:
            self._executor.shutdown(wait=False)
        except Exception:
            pass
        try:
            with sqlite3.connect(self.db_path, timeout=5.0) as conn:
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
        except Exception:
            pass

    def _get_connection(self) -> sqlite3.Connection:
        """
        Creates SQLite connection with Guard G1:
        PRAGMA journal_mode = WAL; PRAGMA busy_timeout = 5000;
        """
        conn = sqlite3.connect(self.db_path, timeout=10.0)
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.execute("PRAGMA busy_timeout = 5000;")
        conn.row_factory = sqlite3.Row
        return conn

    def _execute_with_retry(self, query: str, params: Tuple[Any, ...] = (), max_retries: int = 5) -> Any:
        """Executes SQL write query with exponential backoff retry on database locks."""
        delay = 0.05
        for attempt in range(max_retries):
            try:
                with self._get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(query, params)
                    conn.commit()
                    return cursor
            except sqlite3.OperationalError as e:
                if "locked" in str(e).lower() and attempt < max_retries - 1:
                    time.sleep(delay)
                    delay *= 2
                else:
                    raise

    def _init_database(self) -> None:
        """Creates table and indices if they do not exist."""
        os.makedirs(os.path.dirname(os.path.abspath(self.db_path)), exist_ok=True)
        with self._get_connection() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS system_job_queue (
                    job_id TEXT PRIMARY KEY,
                    action TEXT NOT NULL,
                    status TEXT NOT NULL,
                    progress REAL NOT NULL DEFAULT 0.0,
                    message TEXT,
                    payload_json TEXT,
                    result_json TEXT,
                    error_msg TEXT,
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL
                );
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_job_queue_status
                ON system_job_queue(status);
            """)
            conn.commit()

    def recover_stale_jobs(self) -> int:
        """
        Crash Recovery Sweeper:
        Recovers jobs that were interrupted when server crashed or restarted.
        """
        now = time.time()
        msg = "Tác vụ bị gián đoạn do máy chủ khởi động lại (Crash Recovery). Đã tự động thu hồi."
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE system_job_queue
                SET status = 'failed',
                    message = ?,
                    error_msg = 'SERVER_RESTARTED',
                    updated_at = ?
                WHERE status IN ('running', 'pending');
            """, (msg, now))
            conn.commit()
            return cursor.rowcount

    def submit_job(
        self,
        action: str,
        payload: Dict[str, Any],
        runner_fn: Optional[Callable[[Dict[str, Any], Callable[[float, str], None]], Any]] = None,
    ) -> str:
        """Schedules a persistent background job and returns unique job_id."""
        job_id = f"job_{int(time.time())}_{uuid.uuid4().hex[:8]}"
        now = time.time()
        initial_msg = "Tác vụ đã được lưu vào hàng đợi CSDL bền bỉ..."

        query = """
            INSERT INTO system_job_queue
            (job_id, action, status, progress, message, payload_json, created_at, updated_at)
            VALUES (?, ?, 'pending', 0.0, ?, ?, ?, ?);
        """
        self._execute_with_retry(
            query, (job_id, action, initial_msg, json.dumps(payload, ensure_ascii=False), now, now)
        )

        with self._lock:
            self._subscribers[job_id] = []

        if runner_fn is not None:
            def _worker() -> None:
                self.update_progress(job_id, 5.0, "Bắt đầu thực thi tác vụ nền...", status="running")

                def _progress_cb(pct: float, msg: str) -> None:
                    self.update_progress(job_id, pct, msg)

                try:
                    res = runner_fn(payload, _progress_cb)
                    self.update_progress(job_id, 100.0, "Tác vụ hoàn tất thành công!", status="completed", result=res)
                except Exception as exc:
                    self.update_progress(job_id, 100.0, f"Lỗi tác vụ: {exc}", status="failed", error=str(exc))

            self._executor.submit(_worker)

        return job_id

    def update_progress(
        self,
        job_id: str,
        progress: float,
        message: str,
        status: Optional[str] = None,
        result: Optional[Any] = None,
        error: Optional[str] = None,
    ) -> None:
        """Updates job progress and broadcasts SSE event."""
        now = time.time()
        pct = round(float(progress), 1)

        set_clauses = ["progress = ?", "message = ?", "updated_at = ?"]
        params: List[Any] = [pct, message, now]

        if status is not None:
            set_clauses.append("status = ?")
            params.append(status)
        if result is not None:
            set_clauses.append("result_json = ?")
            params.append(json.dumps(result, ensure_ascii=False))
        if error is not None:
            set_clauses.append("error_msg = ?")
            params.append(error)

        params.append(job_id)
        query = f"UPDATE system_job_queue SET {', '.join(set_clauses)} WHERE job_id = ?"
        self._execute_with_retry(query, tuple(params))

        # Broadcast to in-memory SSE subscribers
        job_data = self.get_job(job_id)
        if job_data:
            with self._lock:
                subscribers = list(self._subscribers.get(job_id, []))
            for q in subscribers:
                try:
                    q.put_nowait(job_data)
                except queue.Full:
                    pass

    def get_job(self, job_id: str) -> Optional[Dict[str, Any]]:
        """Retrieves job record from database."""
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM system_job_queue WHERE job_id = ?", (job_id,)
            ).fetchone()
            if not row:
                return None
            res_val = json.loads(row["result_json"]) if row["result_json"] else None
            payload_val = json.loads(row["payload_json"]) if row["payload_json"] else {}
            return {
                "job_id": row["job_id"],
                "action": row["action"],
                "status": row["status"],
                "progress": row["progress"],
                "message": row["message"],
                "payload": payload_val,
                "result": res_val,
                "error": row["error_msg"],
                "created_at": row["created_at"],
                "updated_at": row["updated_at"],
            }

    def list_jobs(self, limit: int = 50, status: Optional[str] = None) -> List[Dict[str, Any]]:
        """Lists recent jobs with optional status filter."""
        with self._get_connection() as conn:
            if status:
                rows = conn.execute(
                    "SELECT * FROM system_job_queue WHERE status = ? ORDER BY created_at DESC LIMIT ?",
                    (status, limit)
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM system_job_queue ORDER BY created_at DESC LIMIT ?",
                    (limit,)
                ).fetchall()

            result = []
            for r in rows:
                result.append({
                    "job_id": r["job_id"],
                    "action": r["action"],
                    "status": r["status"],
                    "progress": r["progress"],
                    "message": r["message"],
                    "error": r["error_msg"],
                    "created_at": r["created_at"],
                    "updated_at": r["updated_at"],
                })
            return result

    def subscribe(self, job_id: str) -> queue.Queue:
        """Subscribes to live SSE events for a specific job."""
        q: queue.Queue = queue.Queue(maxsize=100)
        with self._lock:
            if job_id not in self._subscribers:
                self._subscribers[job_id] = []
            self._subscribers[job_id].append(q)

        # Push current state immediately if job exists
        current = self.get_job(job_id)
        if current:
            try:
                q.put_nowait(current)
            except queue.Full:
                pass
        return q

    def unsubscribe(self, job_id: str, q: queue.Queue) -> None:
        """Removes a subscriber queue."""
        with self._lock:
            if job_id in self._subscribers:
                try:
                    self._subscribers[job_id].remove(q)
                except ValueError:
                    pass
