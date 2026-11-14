# -*- coding: utf-8 -*-
"""
core.copilot.universal_rag — Universal Multi-Format SQLite FTS5 Semantic Chunking & RAG Engine.

Features:
1. Multi-Format Office Ingestion: Extracts & semantically chunks PDF, DOCX, XLSX, CSV, JSON, and Markdown.
2. SQLite FTS5 with Vietnamese Diacritic Folding: `unicode61 remove_diacritics 2` for sub-10ms retrieval.
3. Concurrency Guard (G1): Enforces `PRAGMA journal_mode = WAL;` and `PRAGMA busy_timeout = 5000;`.
4. Zero-Pip Dependency: Built on Python standard library + existing office core (`docx`, `openpyxl`, `fitz`).
"""

from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import sqlite3
import time
from dataclasses import asdict, dataclass
from typing import Any, Dict, List, Optional

DEFAULT_RAG_DB_PATH = os.path.abspath("./.universal_rag_fts.db")


@dataclass
class DocumentChunk:
    """Represents a semantic chunk extracted from an office or markdown document."""

    doc_id: str
    filename: str
    file_type: str
    section_title: str
    chunk_index: int
    content: str

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class UniversalRAGEngine:
    """
    Universal Office Document Knowledge Base powered by SQLite FTS5 and Semantic Chunking.
    """

    def __init__(self, db_path: Optional[str] = None) -> None:
        self.db_path = os.path.abspath(db_path or DEFAULT_RAG_DB_PATH)
        self._ensure_schema()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, timeout=10.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.execute("PRAGMA busy_timeout = 5000;")
        return conn

    def _ensure_schema(self) -> None:
        parent_dir = os.path.dirname(self.db_path)
        if parent_dir:
            os.makedirs(parent_dir, exist_ok=True)
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute(
                """
                CREATE VIRTUAL TABLE IF NOT EXISTS universal_rag_fts USING fts5(
                    doc_id,
                    filename,
                    file_type,
                    section_title,
                    chunk_index UNINDEXED,
                    content,
                    tokenize = 'unicode61 remove_diacritics 2'
                );
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS universal_rag_docs (
                    doc_id TEXT PRIMARY KEY,
                    filename TEXT,
                    file_type TEXT,
                    chunk_count INTEGER,
                    indexed_at REAL
                );
                """
            )
            conn.commit()
        finally:
            conn.close()

    def close(self) -> None:
        """Flushes WAL checkpoint and releases file handles cleanly on Windows."""
        if not os.path.exists(self.db_path):
            return
        try:
            conn = sqlite3.connect(self.db_path, timeout=5.0)
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            conn.close()
        except Exception:
            pass

    @staticmethod
    def _compute_doc_id(file_path: str) -> str:
        norm = os.path.abspath(file_path).lower()
        return "DOC-" + hashlib.sha1(norm.encode("utf-8")).hexdigest()[:12].upper()

    @staticmethod
    def split_text_Into_chunks(
        text: str,
        doc_id: str,
        filename: str,
        file_type: str,
        default_section: str = "General",
        max_chars: int = 900,
        overlap: int = 120,
    ) -> List[DocumentChunk]:
        """Splits text into overlapping semantic chunks while respecting section headers."""
        if not text or not text.strip():
            return []

        lines = text.splitlines()
        sections: List[tuple[str, List[str]]] = []
        current_sec = default_section
        current_buf: List[str] = []

        for line in lines:
            stripped = line.strip()
            if stripped.startswith("#"):
                if current_buf:
                    sections.append((current_sec, current_buf))
                    current_buf = []
                current_sec = stripped.lstrip("#").strip() or default_section
            elif stripped:
                current_buf.append(stripped)

        if current_buf:
            sections.append((current_sec, current_buf))

        chunks: List[DocumentChunk] = []
        chunk_idx = 0

        for sec_title, sec_lines in sections:
            joined = "\n".join(sec_lines).strip()
            if not joined:
                continue

            if len(joined) <= max_chars:
                chunks.append(
                    DocumentChunk(
                        doc_id=doc_id,
                        filename=filename,
                        file_type=file_type,
                        section_title=sec_title,
                        chunk_index=chunk_idx,
                        content=joined,
                    )
                )
                chunk_idx += 1
                continue

            start = 0
            length = len(joined)
            while start < length:
                end = min(start + max_chars, length)
                piece = joined[start:end].strip()
                if piece:
                    chunks.append(
                        DocumentChunk(
                            doc_id=doc_id,
                            filename=filename,
                            file_type=file_type,
                            section_title=sec_title,
                            chunk_index=chunk_idx,
                            content=piece,
                        )
                    )
                    chunk_idx += 1
                if end >= length:
                    break
                start = max(start + 1, end - overlap)

        return chunks

    @classmethod
    def extract_chunks_from_file(cls, file_path: str, max_chars: int = 900) -> List[DocumentChunk]:
        """Extracts structured semantic chunks from DOCX, XLSX, PDF, CSV, JSON, or Markdown files."""
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File not found for RAG indexing: {file_path}")

        doc_id = cls._compute_doc_id(file_path)
        filename = os.path.basename(file_path)
        ext = os.path.splitext(filename)[1].lower().lstrip(".")

        if ext in ("md", "txt", "yaml", "yml", "log"):
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                raw = f.read()
            return cls.split_text_Into_chunks(
                raw, doc_id=doc_id, filename=filename, file_type=ext, max_chars=max_chars
            )

        if ext == "json":
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                raw = f.read()
            return cls.split_text_Into_chunks(
                raw, doc_id=doc_id, filename=filename, file_type="json", default_section="JSON Data", max_chars=max_chars
            )

        if ext == "csv":
            rows_text: List[str] = []
            with open(file_path, "r", encoding="utf-8-sig", errors="ignore", newline="") as f:
                reader = csv.reader(f)
                for r in reader:
                    line = " | ".join(c.strip() for c in r if c and c.strip())
                    if line:
                        rows_text.append(line)
            return cls.split_text_Into_chunks(
                "\n".join(rows_text),
                doc_id=doc_id,
                filename=filename,
                file_type="csv",
                default_section="CSV Data",
                max_chars=max_chars,
            )

        if ext == "docx":
            from docx import Document

            doc = Document(file_path)
            md_lines: List[str] = []
            for p in doc.paragraphs:
                txt = (p.text or "").strip()
                if not txt:
                    continue
                style_name = (p.style.name if p.style else "").lower()
                if "heading" in style_name or "title" in style_name:
                    md_lines.append(f"# {txt}")
                else:
                    md_lines.append(txt)

            for t_idx, table in enumerate(doc.tables, start=1):
                md_lines.append(f"# Bảng {t_idx}")
                for row in table.rows:
                    cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                    if any(cells):
                        md_lines.append(" | ".join(cells))

            return cls.split_text_Into_chunks(
                "\n".join(md_lines),
                doc_id=doc_id,
                filename=filename,
                file_type="docx",
                default_section="Word Document",
                max_chars=max_chars,
            )

        if ext in ("xlsx", "xlsm"):
            import openpyxl

            wb = openpyxl.load_workbook(file_path, data_only=True)
            chunks: List[DocumentChunk] = []
            idx = 0
            try:
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    sheet_lines: List[str] = []
                    for r in range(1, min(ws.max_row + 1, 500)):
                        vals = [
                            str(ws.cell(r, c).value).strip()
                            for c in range(1, min(ws.max_column + 1, 30))
                            if ws.cell(r, c).value is not None and str(ws.cell(r, c).value).strip()
                        ]
                        if vals:
                            sheet_lines.append(" | ".join(vals))
                    sec_chunks = cls.split_text_Into_chunks(
                        "\n".join(sheet_lines),
                        doc_id=doc_id,
                        filename=filename,
                        file_type="xlsx",
                        default_section=f"Sheet: {sheet_name}",
                        max_chars=max_chars,
                    )
                    for sc in sec_chunks:
                        sc.chunk_index = idx
                        idx += 1
                        chunks.append(sc)
            finally:
                wb.close()
            return chunks

        if ext == "pdf":
            import fitz

            pdf_doc = fitz.open(file_path)
            chunks = []
            idx = 0
            try:
                for page_num in range(len(pdf_doc)):
                    page_text = pdf_doc[page_num].get_text("text") or ""
                    if not page_text.strip():
                        continue
                    page_chunks = cls.split_text_Into_chunks(
                        page_text,
                        doc_id=doc_id,
                        filename=filename,
                        file_type="pdf",
                        default_section=f"Trang {page_num + 1}",
                        max_chars=max_chars,
                    )
                    for pc in page_chunks:
                        pc.chunk_index = idx
                        idx += 1
                        chunks.append(pc)
            finally:
                pdf_doc.close()
            return chunks

        raise ValueError(f"Unsupported file type for Universal RAG indexing: .{ext}")

    def index_chunks(self, chunks: List[DocumentChunk]) -> int:
        """Persists a list of DocumentChunks into the SQLite FTS5 table with retry backoff."""
        if not chunks:
            return 0

        doc_id = chunks[0].doc_id
        filename = chunks[0].filename
        file_type = chunks[0].file_type

        for attempt in range(5):
            conn = self._get_connection()
            try:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM universal_rag_fts WHERE doc_id = ?;", (doc_id,))
                cursor.executemany(
                    """
                    INSERT INTO universal_rag_fts (doc_id, filename, file_type, section_title, chunk_index, content)
                    VALUES (?, ?, ?, ?, ?, ?);
                    """,
                    [
                        (c.doc_id, c.filename, c.file_type, c.section_title, c.chunk_index, c.content)
                        for c in chunks
                    ],
                )
                cursor.execute(
                    """
                    INSERT OR REPLACE INTO universal_rag_docs (doc_id, filename, file_type, chunk_count, indexed_at)
                    VALUES (?, ?, ?, ?, ?);
                    """,
                    (doc_id, filename, file_type, len(chunks), time.time()),
                )
                conn.commit()
                return len(chunks)
            except sqlite3.OperationalError:
                if attempt == 4:
                    raise
                time.sleep(0.05 * (2 ** attempt))
            finally:
                conn.close()
        return 0

    def index_file(self, file_path: str, max_chars: int = 900) -> Dict[str, Any]:
        """Extracts and indexes a single document into the FTS5 store."""
        t0 = time.perf_counter()
        chunks = self.extract_chunks_from_file(file_path, max_chars=max_chars)
        count = self.index_chunks(chunks)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0
        return {
            "success": True,
            "doc_id": chunks[0].doc_id if chunks else self._compute_doc_id(file_path),
            "filename": os.path.basename(file_path),
            "chunks_indexed": count,
            "elapsed_ms": round(elapsed_ms, 2),
        }

    def index_text(
        self,
        doc_id: str,
        filename: str,
        content: str,
        file_type: str = "md",
        section_title: str = "General",
    ) -> Dict[str, Any]:
        """Directly indexes raw text content into Universal RAG."""
        chunks = self.split_text_Into_chunks(
            content,
            doc_id=doc_id,
            filename=filename,
            file_type=file_type,
            default_section=section_title,
        )
        count = self.index_chunks(chunks)
        return {
            "success": True,
            "doc_id": doc_id,
            "filename": filename,
            "chunks_indexed": count,
        }

    @staticmethod
    def _build_fts_query(raw_query: str) -> str:
        tokens = re.findall(r"\w+", (raw_query or "").lower(), flags=re.UNICODE)
        stop_words = {"la", "va", "cua", "cho", "trong", "voi", "nhu", "the", "nao", "gi", "co"}
        filtered = [t for t in tokens if len(t) > 1 and t not in stop_words]
        if not filtered:
            filtered = tokens
        if not filtered:
            return ""
        return " OR ".join(f'"{t}"*' for t in filtered[:12])

    def search(
        self,
        query: str,
        top_k: int = 5,
        file_type: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Searches the FTS5 index and returns ranked semantic chunks."""
        fts_query = self._build_fts_query(query)
        if not fts_query:
            return []

        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            if file_type:
                cursor.execute(
                    """
                    SELECT
                        doc_id,
                        filename,
                        file_type,
                        section_title,
                        chunk_index,
                        content,
                        snippet(universal_rag_fts, 5, '[', ']', '...', 32) AS snippet,
                        bm25(universal_rag_fts) AS score
                    FROM universal_rag_fts
                    WHERE universal_rag_fts MATCH ? AND file_type = ?
                    ORDER BY score
                    LIMIT ?;
                    """,
                    (fts_query, file_type.lower().lstrip("."), int(top_k)),
                )
            else:
                cursor.execute(
                    """
                    SELECT
                        doc_id,
                        filename,
                        file_type,
                        section_title,
                        chunk_index,
                        content,
                        snippet(universal_rag_fts, 5, '[', ']', '...', 32) AS snippet,
                        bm25(universal_rag_fts) AS score
                    FROM universal_rag_fts
                    WHERE universal_rag_fts MATCH ?
                    ORDER BY score
                    LIMIT ?;
                    """,
                    (fts_query, int(top_k)),
                )

            rows = cursor.fetchall()
            results: List[Dict[str, Any]] = []
            for r in rows:
                results.append(
                    {
                        "doc_id": r["doc_id"],
                        "filename": r["filename"],
                        "file_type": r["file_type"],
                        "section_title": r["section_title"],
                        "chunk_index": r["chunk_index"],
                        "content": r["content"],
                        "snippet": r["snippet"],
                        "score": round(abs(float(r["score"] or 0.0)), 4),
                    }
                )
            return results
        except sqlite3.OperationalError:
            return []
        finally:
            conn.close()

    def get_stats(self) -> Dict[str, Any]:
        """Returns summary metrics of indexed documents and chunks."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) AS cnt, COALESCE(SUM(chunk_count), 0) AS total_chunks FROM universal_rag_docs;")
            row = cursor.fetchone()
            cursor.execute("SELECT doc_id, filename, file_type, chunk_count, indexed_at FROM universal_rag_docs ORDER BY indexed_at DESC LIMIT 20;")
            docs = [dict(r) for r in cursor.fetchall()]
            return {
                "total_documents": int(row["cnt"] if row else 0),
                "total_chunks": int(row["total_chunks"] if row else 0),
                "recent_documents": docs,
            }
        finally:
            conn.close()

