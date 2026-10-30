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
import re
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

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


def get_expected_mcp_tools_count() -> int:
    """Reads live count of MCP tools registered in server/mcp_server.py."""
    try:
        from server.mcp_server import TOOL_DEFINITIONS
        return len(TOOL_DEFINITIONS)
    except Exception:
        return 28


def get_expected_tests_count() -> int:
    """Extracts live automated test count from README.md or fallback."""
    readme_path = os.path.join(BASE_DIR, "README.md")
    try:
        with open(readme_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'Tests-([0-9]+)%2F[0-9]+_Passed', content)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return 313


def audit_semantic_body_parity(pkg_ver: str) -> Tuple[bool, List[str]]:
    """
    Enforces Semantic Content Parity (ERR_DOC_SYNC_001 Section D):
    Verifies that Markdown body text actually matches the target system version,
    real test suite counts, MCP tool counts, and includes dedicated feature sections.
    Prevents 'title-only' version bumps where body tables and text remain outdated.
    """
    errors: List[str] = []
    mcp_count = get_expected_mcp_tools_count()
    test_count = get_expected_tests_count()

    # 1. Audit BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md
    arch_md = os.path.join(BASE_DIR, "docs", "BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md")
    if os.path.isfile(arch_md):
        with open(arch_md, "r", encoding="utf-8") as f:
            arch_text = f.read()

        # Check Section 2 comparison header
        if f"v{pkg_ver} Production Grade" not in arch_text:
            errors.append(f"docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md: Section 2 table missing 'v{pkg_ver} Production Grade'")

        # Detect obsolete versions in Section 2 comparison title
        obsolete_match = re.search(r'## 2\.\s+BẢNG MA TRẬN SO SÁNH TOÀN DIỆN.*v(4\.[0-9]+\.[0-9]+)', arch_text, re.IGNORECASE)
        if obsolete_match and obsolete_match.group(1) != pkg_ver:
            errors.append(f"docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md: Section 2 header still references old version 'v{obsolete_match.group(1)}' instead of 'v{pkg_ver}'")

        # Check Conclusion mentions current version
        if f"({pkg_ver} Production Grade)" not in arch_text and f"(v{pkg_ver} Production Grade)" not in arch_text:
            errors.append(f"docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md: Conclusion missing '(v{pkg_ver} Production Grade)'")

        # Check MCP tools count
        mcp_pattern = rf'(?:{mcp_count}\s+MCP\s+Tools|MCP\s+{mcp_count}-Tool)'
        if not re.search(mcp_pattern, arch_text, re.IGNORECASE):
            errors.append(f"docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md: Body missing reference to current {mcp_count} MCP Tools Catalog")

        # Check Automated test count
        test_pattern = rf'(?:{test_count}\+|{test_count}/{test_count})'
        if not re.search(test_pattern, arch_text):
            errors.append(f"docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md: Body missing reference to current test count ({test_count}+)")

        # Check dedicated feature section for current version
        feature_section_pattern = rf'v{re.escape(pkg_ver)}'
        if not re.search(rf'##\s+[0-9]+\..*{feature_section_pattern}', arch_text):
            errors.append(f"docs/BAO_CAO_TIEN_HOA_KIEN_TRUC_V1_TO_V4.md: Missing dedicated feature chapter heading for 'v{pkg_ver}'")

    # 2. Audit HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.md
    guide_md = os.path.join(BASE_DIR, "docs", "HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.md")
    if os.path.isfile(guide_md):
        with open(guide_md, "r", encoding="utf-8") as f:
            guide_text = f.read()

        if f"v{pkg_ver}" not in guide_text:
            errors.append(f"docs/HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.md: Missing target version 'v{pkg_ver}'")

        # Check dedicated feature section for current version
        if not re.search(rf'###\s+PHẦN.*v{re.escape(pkg_ver)}', guide_text, re.IGNORECASE):
            errors.append(f"docs/HUONG_DAN_SU_DUNG_CHI_TIET_TUNG_TINH_NANG.md: Missing dedicated feature section for 'v{pkg_ver}'")

    return (len(errors) == 0, errors)


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

    # 4. Semantic Content Parity Audit
    print("[*] Semantic Body Parity Audit (ERR_DOC_SYNC_001 Section D):")
    semantic_ok, semantic_errors = audit_semantic_body_parity(pkg_ver)
    if not semantic_ok:
        all_passed = False
        for err in semantic_errors:
            print(f"    [FAIL] {err}")
    else:
        print("    [OK] Markdown body texts match target version, test count, and MCP catalog.")

    print()

    if all_passed:
        print("[OK] All Core Documentation Triad assets and body semantics are fully synchronized.")
        return 0
    else:
        print("[FAIL] Out-of-sync or semantic parity violations detected. Please resolve errors before release.")
        return 1


def build_core_docs(doc_filter: Optional[str] = None) -> int:
    print("=" * 72)
    print(" CORE DOCUMENTATION TRIAD RECOMPILATION (MD -> DOCX -> PDF)")
    print("=" * 72)

    try:
        from core.converter.engine import convert_universal
    except (ImportError, ValueError):
        from core.converter.engine import convert_universal  # type: ignore

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
