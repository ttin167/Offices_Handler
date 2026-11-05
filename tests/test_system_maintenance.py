"""
tests.test_system_maintenance — Automated Test Suite for System Maintenance Suite.

Verifies all 7 pillars of periodic workspace maintenance:
1. Zombie Lock & Temp Files Purge (ERR_CONV_004).
2. Orphaned Headless Process Reaper & Safety Guard.
3. SQLite Database Compaction (WAL Checkpoint, VACUUM, FTS5).
4. Bytecode & Test Cache Sanitation (__pycache__).
5. Log Rotation & Output Pruning Safety.
6. Windows MPO Stuttering Diagnostic & UAC Guard.
7. Vitals Telemetry, Dry-Run Mode & CLI / JSON Dispatch.
"""

from __future__ import annotations

import json
import os
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import unittest

from core.system_maintenance import SystemMaintenanceEngine, MaintenanceReport


class TestSystemMaintenanceSuite(unittest.TestCase):
    """Unit and integration tests for SystemMaintenanceEngine."""

    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="test_maint_workspace_")
        self.engine = SystemMaintenanceEngine(workspace_root=self.test_dir)

    def tearDown(self):
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_01_vitals(self):
        """Pillar 7: System Vitals should report platform, disk, and python info."""
        vitals = self.engine.get_system_vitals()
        self.assertEqual(vitals["os_platform"], sys.platform)
        self.assertIn("python_version", vitals)
        self.assertIn("disk_total_gb", vitals)
        self.assertIn("disk_free_gb", vitals)
        self.assertGreater(vitals["disk_total_gb"], 0)

    def test_02_clean_zombie_locks_mock(self):
        """Pillar 1: Lock files (~$*.docx, .~lock.*#) should be detected and purged."""
        lock1 = os.path.join(self.test_dir, "~$mock_document.docx")
        lock2 = os.path.join(self.test_dir, ".~lock.mock_sheet.xlsx#")
        with open(lock1, "w", encoding="utf-8") as f:
            f.write("mock lock content")
        with open(lock2, "w", encoding="utf-8") as f:
            f.write("mock lock content")

        # Dry run should detect but not delete
        res_dry = self.engine.clean_zombie_locks(dry_run=True)
        self.assertEqual(res_dry["found_count"], 2)
        self.assertTrue(os.path.exists(lock1))
        self.assertTrue(os.path.exists(lock2))

        # Live run should delete
        res_live = self.engine.clean_zombie_locks(dry_run=False)
        self.assertEqual(res_live["removed_count"], 2)
        self.assertFalse(os.path.exists(lock1))
        self.assertFalse(os.path.exists(lock2))

    def test_03_clean_temp_artifacts_mock(self):
        """Pillar 1: Pipeline temp files (step_*_temp.*) should be purged."""
        temp_file = os.path.join(self.test_dir, "step_1_temp.docx")
        temp_swap = os.path.join(self.test_dir, "output.crswap")
        with open(temp_file, "w", encoding="utf-8") as f:
            f.write("temporary data")
        with open(temp_swap, "w", encoding="utf-8") as f:
            f.write("swap data")

        res = self.engine.clean_temp_artifacts(dry_run=False, max_age_hours=0.0)
        self.assertGreaterEqual(res["removed_files_count"], 2)
        self.assertFalse(os.path.exists(temp_file))
        self.assertFalse(os.path.exists(temp_swap))

    def test_04_bytecode_cache_mock(self):
        """Pillar 4: __pycache__ and *.pyc files should be purged."""
        sub = os.path.join(self.test_dir, "submodule")
        os.makedirs(sub, exist_ok=True)
        pycache_dir = os.path.join(sub, "__pycache__")
        os.makedirs(pycache_dir, exist_ok=True)
        pyc_file = os.path.join(pycache_dir, "helper.cpython-314.pyc")
        with open(pyc_file, "wb") as f:
            f.write(b"bytecode")

        res_dry = self.engine.clean_bytecode_and_test_cache(dry_run=True)
        self.assertGreaterEqual(res_dry["pycache_dirs_count"], 1)
        self.assertTrue(os.path.exists(pycache_dir))

        res_live = self.engine.clean_bytecode_and_test_cache(dry_run=False)
        self.assertGreaterEqual(res_live["pycache_dirs_count"], 1)
        self.assertFalse(os.path.exists(pycache_dir))

    def test_05_database_compaction_mock(self):
        """Pillar 3: Database vacuum and quick_check on SQLite database."""
        web_out = os.path.join(self.test_dir, ".web_outputs")
        os.makedirs(web_out, exist_ok=True)
        db_path = os.path.join(web_out, "studio_history.db")

        # Create table and insert 500 records then delete 450
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        cur.execute("CREATE TABLE mock_jobs (id INTEGER PRIMARY KEY, data TEXT);")
        for i in range(500):
            cur.execute("INSERT INTO mock_jobs (data) VALUES (?);", (f"data_payload_{i}" * 50,))
        conn.commit()
        cur.execute("DELETE FROM mock_jobs WHERE id > 50;")
        conn.commit()
        conn.close()

        size_before = os.path.getsize(db_path)
        res = self.engine.vacuum_and_optimize_databases(dry_run=False)
        self.assertEqual(res["databases_processed"], 1)
        rep = res["reports"][0]
        self.assertEqual(rep["integrity_status"], "OK")
        self.assertLessEqual(rep["size_after_bytes"], size_before)

    def test_06_database_compaction_live_dry_run(self):
        """Pillar 3: Running vacuum on real workspace databases in dry-run mode."""
        live_engine = SystemMaintenanceEngine()
        res = live_engine.vacuum_and_optimize_databases(dry_run=True)
        self.assertGreaterEqual(res["databases_processed"], 1)
        for rep in res["reports"]:
            self.assertEqual(rep["integrity_status"], "OK")

    def test_07_mpo_check(self):
        """Pillar 6: Windows DWM MPO check should execute safely and return valid status."""
        mpo = self.engine.check_and_diagnose_mpo(apply_fix=False)
        self.assertIn("supported", mpo)
        self.assertIn("status", mpo)
        self.assertIn("recommendation", mpo)
        if sys.platform == "win32":
            self.assertTrue(mpo["supported"])
            self.assertIn(mpo["status"], ["OPTIMAL_DISABLED", "WARNING_MPO_ENABLED"])

    def test_08_run_all_dry_run(self):
        """Pillar 7: run_all in dry_run mode should populate all report sections."""
        report = self.engine.run_all(dry_run=True)
        self.assertTrue(report.dry_run)
        self.assertIsInstance(report.vitals, dict)
        self.assertGreaterEqual(report.execution_time_sec, 0.0)
        d = report.to_dict()
        self.assertIn("total_bytes_freed", d)
        self.assertIn("mpo_status", d)

    def test_09_cli_dispatch_dry_run(self):
        """CLI: 'python ai_tools_cli.py system-maintenance --dry-run' should exit with 0."""
        cmd = [sys.executable, "ai_tools_cli.py", "system-maintenance", "--dry-run"]
        res = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
        self.assertEqual(res.returncode, 0, f"CLI error: {res.stderr}")
        self.assertIn("ANTIGRAVITY SYSTEM MAINTENANCE SUITE", res.stdout)
        self.assertIn("DRY-RUN PREVIEW", res.stdout)

    def test_10_cli_dispatch_json(self):
        """CLI: 'python ai_tools_cli.py system-maintenance --dry-run --json' should produce valid JSON."""
        cmd = [sys.executable, "ai_tools_cli.py", "system-maintenance", "--dry-run", "--json"]
        res = subprocess.run(cmd, capture_output=True, text=True, errors="replace")
        self.assertEqual(res.returncode, 0, f"CLI error: {res.stderr}")
        data = json.loads(res.stdout)
        self.assertTrue(data.get("dry_run"))
        self.assertIn("vitals", data)
        self.assertIn("databases_vacuumed", data)

    def test_11_selective_flags(self):
        """CLI: Selective flags should run only chosen pillars."""
        report = self.engine.run_all(
            dry_run=True,
            clean_locks=True,
            clean_temp=False,
            reap_processes=False,
            vacuum_db=False,
            clear_cache=False,
            rotate_logs=False,
            prune_outputs=False,
            check_mpo=False,
        )
        self.assertEqual(len(report.databases_vacuumed), 0)
        self.assertEqual(report.processes_scanned, 0)

    def test_12_outputs_prune_safety(self):
        """Pillar 5: Prune outputs should preserve studio_history.db and recent files."""
        web_out = os.path.join(self.test_dir, ".web_outputs")
        os.makedirs(web_out, exist_ok=True)
        db_file = os.path.join(web_out, "studio_history.db")
        recent_file = os.path.join(web_out, "recent_output.pdf")
        with open(db_file, "w", encoding="utf-8") as f:
            f.write("db")
        with open(recent_file, "w", encoding="utf-8") as f:
            f.write("recent")

        res = self.engine.prune_web_outputs(dry_run=False, max_age_days=7)
        self.assertTrue(os.path.exists(db_file), "Database file was deleted!")
        self.assertTrue(os.path.exists(recent_file), "Recent output was deleted!")

    def test_13_process_scan_safety(self):
        """Pillar 2: Process reaper in dry-run mode should not terminate anything."""
        res = self.engine.reap_orphaned_processes(dry_run=True)
        self.assertGreaterEqual(res["processes_scanned"], 1)

    def test_14_log_rotation_mock(self):
        """Pillar 5: Large log files should be rotated."""
        big_log = os.path.join(self.test_dir, "test_large.log")
        with open(big_log, "wb") as f:
            f.write(b"0" * (1024 * 1024 + 100))  # 1MB + 100 bytes

        res = self.engine.rotate_and_prune_logs(dry_run=False, max_size_mb=1.0)
        self.assertEqual(res["rotated_count"], 1)
        self.assertTrue(os.path.exists(big_log + ".old"))
        # Original should now be truncated to 0
        self.assertEqual(os.path.getsize(big_log), 0)

    def test_15_uac_guard_mpo(self):
        """Pillar 6: UAC elevation check for MPO fix should not crash."""
        mpo = self.engine.check_and_diagnose_mpo(apply_fix=True)
        self.assertIsInstance(mpo, dict)
        if sys.platform == "win32" and not mpo.get("is_admin"):
            self.assertIn("recommendation", mpo)


if __name__ == "__main__":
    unittest.main()

