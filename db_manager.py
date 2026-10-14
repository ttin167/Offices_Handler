"""
db_manager.py — Lightweight Zero-Config SQLite History & Task Storage.

Manages persistent document conversion history, job records, and output downloads
for the Antigravity Office Studio Web & Desktop interface.
Zero external pip dependencies (uses built-in sqlite3).
"""

from __future__ import annotations

import os
import sqlite3
import time
from typing import Any, Dict, List, Optional

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, ".web_outputs")
DB_PATH = os.path.join(OUTPUT_DIR, "studio_history.db")


def get_db_path() -> str:
    """Returns absolute path to SQLite database, ensuring directory exists."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    return DB_PATH


def get_db_connection() -> sqlite3.Connection:
    """Creates a thread-safe connection to the SQLite database."""
    conn = sqlite3.connect(get_db_path(), timeout=10.0, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """Initializes SQLite schema if not already present."""
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS conversion_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_type TEXT NOT NULL,
                source_filename TEXT,
                output_filename TEXT NOT NULL,
                output_path TEXT NOT NULL,
                file_size INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT (datetime('now', 'localtime'))
            );
        """)
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_history_created 
            ON conversion_history(created_at DESC);
        """)
        conn.commit()


def record_history(
    task_type: str,
    source_filename: Optional[str],
    output_filename: str,
    output_path: str,
    file_size: int = 0
) -> int:
    """
    Records a completed file conversion or export into the history database.
    Calculates file size if not provided and file exists.
    """
    init_db()

    if file_size <= 0 and output_path and os.path.isfile(output_path):
        try:
            file_size = os.path.getsize(output_path)
        except Exception:
            file_size = 0

    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO conversion_history 
            (task_type, source_filename, output_filename, output_path, file_size)
            VALUES (?, ?, ?, ?, ?)
        """, (
            task_type,
            source_filename or "N/A",
            output_filename,
            output_path,
            file_size
        ))
        conn.commit()
        return cursor.lastrowid or 0


def get_recent_history(limit: int = 50) -> List[Dict[str, Any]]:
    """Retrieves most recent conversion history items sorted newest first."""
    init_db()
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, task_type, source_filename, output_filename, 
                   output_path, file_size, created_at
            FROM conversion_history
            ORDER BY id DESC
            LIMIT ?
        """, (limit,))
        rows = cursor.fetchall()
        
        results = []
        for r in rows:
            results.append({
                "id": r["id"],
                "task_type": r["task_type"],
                "source_filename": r["source_filename"],
                "output_filename": r["output_filename"],
                "file_size": r["file_size"],
                "download_url": f"/api/download/{r['output_filename']}",
                "created_at": r["created_at"],
            })
        return results


def delete_history_item(item_id: int) -> bool:
    """Deletes a specific history record by ID."""
    init_db()
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM conversion_history WHERE id = ?", (item_id,))
        conn.commit()
        return cursor.rowcount > 0


def clear_history() -> bool:
    """Clears all history records."""
    init_db()
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM conversion_history")
        conn.commit()
        return True
