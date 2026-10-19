"""
core.job_manager — Thread-Safe Background Job Queue & SSE Event Dispatcher (v4.0).

Provides asynchronous task execution, thread-safe progress tracking,
and Server-Sent Events (SSE) streaming for long-running document operations.
"""

from __future__ import annotations

import concurrent.futures
import json
import queue
import threading
import time
import os
import uuid
from typing import Any, Callable, Dict, List, Optional


class JobManager:
    """
    Manages asynchronous background jobs with thread-safe progress tracking
    and multi-subscriber SSE event queues.
    """

    def __init__(self, max_workers: Optional[int] = None):
        if max_workers is None:
            try:
                max_workers = max(1, int(os.environ.get("MAX_JOB_WORKERS", 4)))
            except (ValueError, TypeError):
                max_workers = 4
        self._max_workers = max_workers
        self._executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=max_workers,
            thread_name_prefix="JobWorker"
        )
        self._lock = threading.Lock()
        self._jobs: Dict[str, Dict[str, Any]] = {}
        self._subscribers: Dict[str, List[queue.Queue]] = {}

    def submit_job(
        self,
        action: str,
        payload: Dict[str, Any],
        runner_fn: Callable[[Dict[str, Any], Callable[[float, str], None]], Any]
    ) -> str:
        """
        Creates and schedules a background job. Returns unique job_id.
        """
        job_id = f"job_{int(time.time())}_{uuid.uuid4().hex[:8]}"
        now = time.time()

        with self._lock:
            self._jobs[job_id] = {
                "job_id": job_id,
                "action": action,
                "status": "pending",
                "progress": 0.0,
                "message": "Công việc đã được đưa vào hàng đợi...",
                "payload": payload,
                "result": None,
                "error": None,
                "created_at": now,
                "updated_at": now,
            }
            self._subscribers[job_id] = []

        def _execute_task():
            self._set_status(job_id, "running", progress=5.0, message="Bắt đầu xử lý tác vụ...")

            def _progress_cb(pct: float, msg: str):
                self._update_progress(job_id, pct, msg)

            try:
                result = runner_fn(payload, _progress_cb)
                self._set_status(job_id, "completed", progress=100.0, message="Hoàn tất thành công!", result=result)
            except Exception as exc:
                self._set_status(job_id, "failed", progress=100.0, message=f"Lỗi: {exc}", error=str(exc))

        self._executor.submit(_execute_task)
        return job_id

    def _update_progress(self, job_id: str, progress: float, message: str) -> None:
        """Update progress percent and broadcast SSE event."""
        with self._lock:
            job = self._jobs.get(job_id)
            if not job or job["status"] in ("completed", "failed"):
                return
            job["progress"] = round(float(progress), 1)
            job["message"] = message
            job["updated_at"] = time.time()
            snapshot = dict(job)
            subscribers = list(self._subscribers.get(job_id, []))

        # Broadcast outside of lock
        for q in subscribers:
            try:
                q.put_nowait(snapshot)
            except queue.Full:
                pass

    def _set_status(
        self,
        job_id: str,
        status: str,
        progress: float = 0.0,
        message: str = "",
        result: Any = None,
        error: Optional[str] = None
    ) -> None:
        """Update terminal status and broadcast final event."""
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return
            job["status"] = status
            job["progress"] = round(float(progress), 1)
            job["message"] = message
            if result is not None:
                job["result"] = result
            if error is not None:
                job["error"] = error
            job["updated_at"] = time.time()
            snapshot = dict(job)
            subscribers = list(self._subscribers.get(job_id, []))

        for q in subscribers:
            try:
                q.put_nowait(snapshot)
            except queue.Full:
                pass

    def get_job(self, job_id: str) -> Optional[Dict[str, Any]]:
        """Return a snapshot dictionary of job status, or None."""
        with self._lock:
            job = self._jobs.get(job_id)
            return dict(job) if job else None

    def list_jobs(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Return recent jobs list sorted by creation time descending."""
        with self._lock:
            jobs_list = [dict(j) for j in self._jobs.values()]
        jobs_list.sort(key=lambda x: x["created_at"], reverse=True)
        return jobs_list[:limit]

    def subscribe(self, job_id: str) -> Optional[queue.Queue]:
        """Subscribe to live SSE events for a job."""
        with self._lock:
            if job_id not in self._jobs:
                return None
            q: queue.Queue = queue.Queue(maxsize=100)
            self._subscribers.setdefault(job_id, []).append(q)
            return q

    def unsubscribe(self, job_id: str, q: queue.Queue) -> None:
        """Unsubscribe and cleanup subscriber queue."""
        with self._lock:
            if job_id in self._subscribers and q in self._subscribers[job_id]:
                self._subscribers[job_id].remove(q)

    def shutdown(self, wait: bool = True) -> None:
        """Shutdown underlying ThreadPoolExecutor."""
        self._executor.shutdown(wait=wait)


# Global Singleton Instance
global_job_manager = JobManager()

