#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tools/check_ba_leak.py — Guard against committing/pushing local-only BA business files, JS modules, or routes.

Checks:
  1. Git index (tracked files) and staged changes for any blocked BA / Labor Law paths.
  2. Core tracked entrypoints (server/app.py, server/mcp_server.py, server/routes/__init__.py,
     src/main.js, index.html) for inline BA routes, static imports, or embedded HTML panels.
  3. Parity of BA exclusion patterns across .gitignore, .dockerignore, .antigravityignore,
     .geminiignore, and Dockerfile.

Usage:
  python tools/check_ba_leak.py            # Audit tracked files, code, and ignore parity
  python tools/check_ba_leak.py --staged   # Also suitable for .git/hooks/pre-commit
"""

from __future__ import annotations

import argparse
import fnmatch
import re
import subprocess
import sys
from pathlib import Path
from typing import Iterable, List, Tuple

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = Path(__file__).resolve().parent.parent

BLOCKED_PREFIXES: Tuple[str, ...] = (
    "core/ba_hub/",
    "src/templates/ba/",
    ".agents/skills/ba-workspace/",
    "labor_law_vault/",
    "ba_academy_vault/",
    "data/labor_law/",
)

BLOCKED_GLOBS: Tuple[str, ...] = (
    "cli/ba_commands.py",
    "cli/ba_print_commands.py",
    "server/routes/ba_*.py",
    "src/modules/ba_*.js",
    "src/modules/labor_law_vault.js",
    "src/modules/app_store.js",
    "specs/ba_*.json",
    "tests/test_ba_*.py",
    "tests/test_labor_law_*.py",
    "tests/test_regulatory_*.py",
    "tests/test_hr_governance.py",
    "*.labor_law_vault.zip",
    ".labor_law_fts.db*",
    ".ba_academy.db*",
)

REQUIRED_IGNORE_ENTRIES: Tuple[str, ...] = (
    "core/ba_hub/",
    "server/routes/ba_*.py",
    "src/modules/ba_*.js",
    "src/modules/labor_law_vault.js",
    "src/modules/app_store.js",
    "src/templates/ba/",
    "tests/test_hr_governance.py",
)


def _git_lines(args: List[str]) -> List[str]:
    proc = subprocess.run(
        ["git", *args],
        cwd=str(BASE_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if proc.returncode != 0:
        return []
    return [line.strip().replace("\\", "/") for line in proc.stdout.splitlines() if line.strip()]


def _is_blocked_path(rel_path: str) -> bool:
    norm = rel_path.lstrip("./").replace("\\", "/")
    if any(norm.startswith(prefix) for prefix in BLOCKED_PREFIXES):
        return True
    return any(fnmatch.fnmatch(norm, pat) for pat in BLOCKED_GLOBS)


def check_git_paths(staged_only: bool = False) -> List[str]:
    """Return error strings for any tracked or staged file matching blocked BA paths."""
    errors: List[str] = []
    tracked = _git_lines(["ls-files"])
    staged = _git_lines(["diff", "--cached", "--name-only", "--diff-filter=ACMRT"])

    candidates: Iterable[str] = staged if staged_only else sorted(set(tracked) | set(staged))
    for path in candidates:
        if _is_blocked_path(path):
            errors.append(f"[BLOCKED PATH IN GIT INDEX/STAGE] {path}")
    return errors


def check_inline_leaks() -> List[str]:
    """Verify core tracked files do not embed BA routes, tools, static imports, or HTML panels."""
    errors: List[str] = []

    app_py = BASE_DIR / "server" / "app.py"
    if app_py.is_file():
        txt = app_py.read_text(encoding="utf-8")
        if re.search(r'elif\s+path\s*==\s*["\']/(?:api/ba|api/law)/', txt):
            errors.append("[INLINE ROUTE LEAK] server/app.py still contains direct /api/ba/* or /api/law/* handlers.")
        if re.search(r'from\s+core\.ba_hub\s+import\s+\(', txt):
            errors.append("[INLINE IMPORT LEAK] server/app.py still contains bulk import from core.ba_hub.")

    mcp_py = BASE_DIR / "server" / "mcp_server.py"
    if mcp_py.is_file():
        txt = mcp_py.read_text(encoding="utf-8")
        for tool_name in ("ba_story_lint", "ba_urd_synthesize", "ba_fsd_synthesize", "law_knowledge_search", "hr_governance_blueprint"):
            if f'"{tool_name}"' in txt:
                errors.append(f"[INLINE MCP LEAK] server/mcp_server.py still defines '{tool_name}'.")

    routes_init = BASE_DIR / "server" / "routes" / "__init__.py"
    if routes_init.is_file():
        for idx, line in enumerate(routes_init.read_text(encoding="utf-8").splitlines(), start=1):
            stripped = line.strip()
            if stripped.startswith("from . import ba_") and not line.startswith("    "):
                errors.append(f"[UNGUARDED ROUTE IMPORT] server/routes/__init__.py:{idx}: {stripped}")

    main_js = BASE_DIR / "src" / "main.js"
    if main_js.is_file():
        txt = main_js.read_text(encoding="utf-8")
        if re.search(r"^\s*import\s+.*['\"]\./modules/(?:ba_|labor_law_vault|app_store)", txt, flags=re.MULTILINE):
            errors.append("[STATIC JS IMPORT LEAK] src/main.js statically imports a local-only BA/Law/Store JS module.")

    index_html = BASE_DIR / "index.html"
    if index_html.is_file():
        txt = index_html.read_text(encoding="utf-8")
        for panel_id in ("panel-requirements", "panel-ba-academy", "panel-law-vault", "panel-data-process", "panel-audit-recon"):
            if f'id="{panel_id}"' in txt:
                errors.append(f"[INLINE HTML PANEL LEAK] index.html still embeds #{panel_id}.")

    return errors


def check_ignore_parity() -> List[str]:
    """Verify .gitignore, .dockerignore, .antigravityignore, .geminiignore, and Dockerfile stay synced."""
    errors: List[str] = []
    ignore_files = (".gitignore", ".dockerignore", ".antigravityignore", ".geminiignore")
    for fname in ignore_files:
        fpath = BASE_DIR / fname
        if not fpath.is_file():
            errors.append(f"[MISSING IGNORE FILE] {fname}")
            continue
        txt = fpath.read_text(encoding="utf-8")
        for req in REQUIRED_IGNORE_ENTRIES:
            if req not in txt:
                errors.append(f"[IGNORE DESYNC] {fname} missing '{req}'")

    dockerfile = BASE_DIR / "Dockerfile"
    if dockerfile.is_file():
        d_txt = dockerfile.read_text(encoding="utf-8")
        for req in ("server/routes/ba_*.py", "src/modules/ba_*.js", "src/modules/labor_law_vault.js", "src/modules/app_store.js", "src/templates/ba", "tests/test_hr_governance.py"):
            if req not in d_txt:
                errors.append(f"[DOCKERFILE PRUNE DESYNC] Dockerfile missing '/app/{req}'")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify local-only BA files, JS modules, and routes are excluded from Git.")
    parser.add_argument("--staged", action="store_true", help="Check staged files only for git path rules (pre-commit mode).")
    args = parser.parse_args()

    failures: List[str] = []
    failures.extend(check_git_paths(staged_only=args.staged))
    failures.extend(check_inline_leaks())
    failures.extend(check_ignore_parity())

    if failures:
        print("[FAIL] BA Leak Guard detected issues:", file=sys.stderr)
        for item in failures:
            print(f"  - {item}", file=sys.stderr)
        return 1

    print("[OK] BA Leak Guard passed: 0 tracked/staged BA leaks, 0 inline leaks, ignore parity 100%.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

