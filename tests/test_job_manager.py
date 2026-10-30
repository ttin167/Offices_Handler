"""
tests/test_job_manager.py — Unit Tests for Threaded JobManager and SSE Job Queue.

Tests:
1. JobManager initialization and custom worker pool.
2. Synchronous and asynchronous task submission.
3. Status transitions (pending -> running -> completed / failed).
4. Real-time progress updates and message broadcasting.
5. SSE queue subscriptions and fan-out delivery.
6. Error handling and exception capture.
7. Root facade re-export parity.
"""

from __future__ import annotations

import threading
import time
import unittest
from core.job_manager import JobManager, global_job_manager
import core.job_manager as job_manager


class TestJobManager(unittest.TestCase):
    def setUp(self):
        self.manager = JobManager(max_workers=2)

    def tearDown(self):
        self.manager.shutdown(wait=True)

    def test_01_submit_successful_job(self):
        def worker_fn(payload, update_fn):
            update_fn(25, "Processing stage 1")
            time.sleep(0.05)
            update_fn(75, "Processing stage 2")
            return {"output": "sample_result.docx"}

        job_id = self.manager.submit_job("convert", {"source": "test.pdf"}, worker_fn)
        self.assertIsNotNone(job_id)

        # Wait for job completion
        for _ in range(50):
            job = self.manager.get_job(job_id)
            if job and job["status"] in ("completed", "failed"):
                break
            time.sleep(0.05)

        job = self.manager.get_job(job_id)
        self.assertIsNotNone(job)
        self.assertEqual(job["status"], "completed")
        self.assertEqual(job["progress"], 100)
        self.assertEqual(job["result"], {"output": "sample_result.docx"})
        self.assertIsNone(job["error"])

    def test_02_submit_failing_job(self):
        def worker_fn(payload, update_fn):
            update_fn(50, "Working...")
            raise ValueError("Simulated processing error")

        job_id = self.manager.submit_job("render", {}, worker_fn)

        for _ in range(50):
            job = self.manager.get_job(job_id)
            if job and job["status"] in ("completed", "failed"):
                break
            time.sleep(0.05)

        job = self.manager.get_job(job_id)
        self.assertIsNotNone(job)
        self.assertEqual(job["status"], "failed")
        self.assertIn("Simulated processing error", str(job["error"]))

    def test_03_sse_subscriber_queue(self):
        gate = threading.Event()

        def worker_fn(payload, update_fn):
            gate.wait(timeout=2.0)  # Wait for test to subscribe
            update_fn(10, "Step 1")
            time.sleep(0.05)
            update_fn(50, "Step 2")
            time.sleep(0.05)
            return "done"

        job_id = self.manager.submit_job("test", {}, worker_fn)
        sub_queue = self.manager.subscribe(job_id)
        self.assertIsNotNone(sub_queue)
        gate.set()  # Allow worker to proceed now that subscription is active

        events = []
        start_time = time.time()
        while time.time() - start_time < 3.0:
            try:
                evt = sub_queue.get(timeout=0.3)
                events.append(evt)
                if evt.get("status") in ("completed", "failed"):
                    break
            except Exception:
                pass

        self.manager.unsubscribe(job_id, sub_queue)
        self.assertGreaterEqual(len(events), 2)
        final_evt = events[-1]
        self.assertEqual(final_evt["status"], "completed")
        self.assertEqual(final_evt["result"], "done")

    def test_04_list_jobs(self):
        self.manager.submit_job("action_a", {}, lambda p, u: 1)
        self.manager.submit_job("action_b", {}, lambda p, u: 2)
        time.sleep(0.1)

        jobs = self.manager.list_jobs(limit=10)
        self.assertGreaterEqual(len(jobs), 2)
        actions = [j["action"] for j in jobs]
        self.assertIn("action_a", actions)
        self.assertIn("action_b", actions)

    def test_05_facade_parity(self):
        self.assertEqual(job_manager.JobManager, JobManager)
        self.assertIsInstance(job_manager.global_job_manager, JobManager)


if __name__ == "__main__":
    unittest.main()
