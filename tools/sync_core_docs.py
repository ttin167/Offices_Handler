#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tools/sync_core_docs.py — Core Technical & User Documentation Triad Synchronizer.

Implements Invariant ERR_DOC_SYNC_001 Section D:
Maintains the two foundational core documents in docs/:
  1. docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4 (Architecture Evolution Report)
  2. docs/HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG (Feature User Guide)
Enforces the Triad Invariant: MD source -> DOCX -> PDF (Lossless Vector / OpenXML).

Usage:
  python tools/sync_core_docs.py --build          # Recompile all core MD -> DOCX & PDF
  python tools/sync_core_docs.py --check          # Validate mtime & version synchronization
  python tools/sync_core_docs.py --doc <id>       # Build specific document ('arch' or 'guide')
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
import time
from typing import Any, Dict, List, Optional

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

CORE_DOCS: List[Dict[str, str]] = [
    {
        "id": "arch",
        "title": "Báo Cáo Tiến Hóa Kiến Trúc",
        "md": "docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md",
        "docx": "docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.docx",
        "pdf": "docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.pdf",
        "style": "docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.style.yaml",
    },
    {
        "id": "guide",
        "title": "Cẩm Nang Hướng Dẫn Sử Dụng Tính Năng",
        "md": "docs/HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.md",
        "docx": "docs/HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.docx",
        "pdf": "docs/HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.pdf",
        "style": "docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.style.yaml",
    },
]


def get_current_package_version() -> str:
    pkg_path = os.path.join(BASE_DIR, "package.json")
    try:
        with open(pkg_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("version", "4.3.0")
    except Exception:
        return "4.3.0"


def check_sync_status() -> int:
    print("=" * 72)
    print(" CORE DOCUMENTATION TRIAD AUDIT (ERR_DOC_SYNC_001 Section D)")
    print("=" * 72)
    pkg_ver = get_current_package_version()
    print(f"Target System Version: v{pkg_ver}\n")

    all_passed = True

    for doc in CORE_DOCS:
        md_abs = os.path.join(BASE_DIR, doc["md"])
        docx_abs = os.path.join(BASE_DIR, doc["docx"])
        pdf_abs = os.path.join(BASE_DIR, doc["pdf"])

        print(f"[*] Document: {doc['title']} ({doc['id']})")

        # 1. Check MD existence
        if not os.path.isfile(md_abs):
            print(f"    [FAIL] Source Markdown missing: {doc['md']}")
            all_passed = False
            continue

        md_mtime = os.path.getmtime(md_abs)
        md_time_str = datetime.datetime.fromtimestamp(md_mtime).strftime("%Y-%m-%d %H:%M:%S")
        print(f"    - MD   : {doc['md']} (Modified: {md_time_str})")

        # 2. Check DOCX existence and mtime
        if not os.path.isfile(docx_abs):
            print(f"    [FAIL] Exported DOCX missing: {doc['docx']}")
            all_passed = False
        else:
            docx_mtime = os.path.getmtime(docx_abs)
            docx_time_str = datetime.datetime.fromtimestamp(docx_mtime).strftime("%Y-%m-%d %H:%M:%S")
            if docx_mtime < md_mtime - 5.0:
                print(f"    [WARN] DOCX is older than MD by {md_mtime - docx_mtime:.1f}s ({doc['docx']})")
                all_passed = False
            else:
                print(f"    - DOCX : {doc['docx']} (Synced: {docx_time_str})")

        # 3. Check PDF existence and mtime
        if not os.path.isfile(pdf_abs):
            print(f"    [FAIL] Exported PDF missing: {doc['pdf']}")
            all_passed = False
        else:
            pdf_mtime = os.path.getmtime(pdf_abs)
            pdf_time_str = datetime.datetime.fromtimestamp(pdf_mtime).strftime("%Y-%m-%d %H:%M:%S")
            if pdf_mtime < md_mtime - 5.0:
                print(f"    [WARN] PDF is older than MD by {md_mtime - pdf_mtime:.1f}s ({doc['pdf']})")
                all_passed = False
            else:
                print(f"    - PDF  : {doc['pdf']} (Synced: {pdf_time_str})")

        print()

    if all_passed:
        print("[OK] All Core Documentation Triad assets are fully synchronized.")
        return 0
    else:
        print("[FAIL] Out-of-sync or missing documents detected. Run 'python tools/sync_core_docs.py --build' to repair.")
        return 1


def build_core_docs(doc_filter: Optional[str] = None) -> int:
    print("=" * 72)
    print(" CORE DOCUMENTATION TRIAD RECOMPILATION (MD -> DOCX -> PDF)")
    print("=" * 72)

    try:
        from core.converter.engine import convert_universal
    except (ImportError, ValueError):
        from converter_engine import convert_universal  # type: ignore

    targets = CORE_DOCS
    if doc_filter:
        targets = [d for d in CORE_DOCS if d["id"] == doc_filter or doc_filter in d["md"]]
        if not targets:
            print(f"[ERROR] Document filter '{doc_filter}' matched no core documents.", file=sys.stderr)
            return 1

    total_start = time.time()
    errors = 0

    for doc in targets:
        md_abs = os.path.join(BASE_DIR, doc["md"])
        docx_abs = os.path.join(BASE_DIR, doc["docx"])
        pdf_abs = os.path.join(BASE_DIR, doc["pdf"])
        style_abs = os.path.join(BASE_DIR, doc["style"]) if doc.get("style") else None

        print(f"\n[BUILDING] {doc['title']}")
        print(f"  Source : {doc['md']}")

        if not os.path.isfile(md_abs):
            print(f"  [ERROR] Source file not found: {md_abs}", file=sys.stderr)
            errors += 1
            continue

        # Step 1: MD -> DOCX
        t0 = time.time()
        ok_docx, out_docx, elap_docx = convert_universal(
            source_path=md_abs,
            target_path=docx_abs,
            style_path=style_abs if style_abs and os.path.isfile(style_abs) else None,
        )
        if not ok_docx or not os.path.isfile(docx_abs):
            print(f"  [ERROR] Failed to compile DOCX: {out_docx}", file=sys.stderr)
            errors += 1
            continue
        docx_size = os.path.getsize(docx_abs) / 1024.0
        print(f"  [OK] MD -> DOCX in {elap_docx:.2f}s ({docx_size:.1f} KB)")

        # Step 2: DOCX -> PDF
        t1 = time.time()
        ok_pdf, out_pdf, elap_pdf = convert_universal(
            source_path=docx_abs,
            target_path=pdf_abs,
        )
        if not ok_pdf or not os.path.isfile(pdf_abs):
            print(f"  [ERROR] Failed to export PDF: {out_pdf}", file=sys.stderr)
            errors += 1
            continue
        pdf_size = os.path.getsize(pdf_abs) / 1024.0
        print(f"  [OK] DOCX -> PDF in {elap_pdf:.2f}s ({pdf_size:.1f} KB)")

    total_elap = time.time() - total_start
    print("-" * 72)
    if errors == 0:
        print(f"[SUCCESS] Core Documentation Triad successfully compiled in {total_elap:.2f}s!")
        return 0
    else:
        print(f"[FAILURE] Completed with {errors} error(s) in {total_elap:.2f}s.", file=sys.stderr)
        return 2


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Core Technical & User Documentation Triad Synchronizer (ERR_DOC_SYNC_001 Section D)"
    )
    parser.add_argument("--build", action="store_true", help="Compile Markdown sources to DOCX and PDF")
    parser.add_argument("--check", action="store_true", help="Audit synchronization timestamps and existence")
    parser.add_argument("--doc", type=str, default=None, help="Target specific document id ('arch' or 'guide')")

    args = parser.parse_args()

    if args.check:
        return check_sync_status()

    # Default action is build
    return build_core_docs(doc_filter=args.doc)


if __name__ == "__main__":
    sys.exit(main())
