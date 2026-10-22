"""
felo_adapter.py — Python Adapter for Felo-Slides Agent Skill.

Acts as a bridge between Antigravity / ai_tools_cli and Felo AI Cloud Service:
- Automatic discovery of FELO_API_KEY from environment or Windows User registry
- Execution of bundled Node.js script run_ppt_task.mjs
- Structured parsing of task results and URL logging
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from typing import Any, Dict, Optional


def get_felo_api_key() -> Optional[str]:
    """Retrieves FELO_API_KEY from environment or Windows registry."""
    key = os.environ.get("FELO_API_KEY")
    if key and key.strip():
        return key.strip()

    # On Windows, try reading User environment variable directly
    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as reg_key:
                val, _ = winreg.QueryValueEx(reg_key, "FELO_API_KEY")
                if val:
                    return val.strip()
        except Exception:
            pass

    return None


def run_felo_ppt_task(
    query: str,
    file_path: Optional[str] = None,
    theme_id: Optional[str] = None,
    task_id: Optional[str] = None,
    interval: int = 10,
    max_wait: int = 1800,
) -> Dict[str, Any]:
    """
    Executes run_ppt_task.mjs with Node.js and returns structured task output.
    """
    api_key = get_felo_api_key()
    if not api_key:
        raise ValueError(
            "FELO_API_KEY is not configured! Please set it in environment:\n"
            "  Windows PowerShell: $env:FELO_API_KEY='your-key'\n"
            "  Linux/macOS:        export FELO_API_KEY='your-key'"
        )

    script_dir = os.path.dirname(os.path.abspath(__file__))
    mjs_path = os.path.join(script_dir, "run_ppt_task.mjs")
    if not os.path.isfile(mjs_path):
        raise FileNotFoundError(f"Felo runner script not found: {mjs_path}")

    cmd = [
        "node",
        mjs_path,
        "--json",
        "--interval", str(interval),
        "--max-wait", str(max_wait),
    ]

    if task_id:
        cmd.extend(["--task-id", task_id])
    else:
        cmd.extend(["--query", query])

    if file_path:
        cmd.extend(["--file", os.path.abspath(file_path)])

    if theme_id:
        cmd.extend(["--theme", theme_id])

    env = os.environ.copy()
    env["FELO_API_KEY"] = api_key

    print(f"[INFO] Invoking Felo Cloud PPT generation with query: '{query}'...")
    proc = subprocess.run(cmd, capture_output=True, text=True, env=env)

    stdout = proc.stdout.strip()
    stderr = proc.stderr.strip()

    if proc.returncode != 0:
        raise RuntimeError(f"Felo task failed (code {proc.returncode}):\n{stderr or stdout}")

    # Parse JSON output
    try:
        # Sometimes there may be log lines before JSON; extract between first { and last }
        start = stdout.find("{")
        end = stdout.rfind("}")
        if start != -1 and end != -1:
            return json.loads(stdout[start:end+1])
        return {"raw_output": stdout}
    except Exception as exc:
        return {"raw_output": stdout, "parse_error": str(exc)}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python felo_adapter.py <query> [file_path]")
        sys.exit(1)

    q = sys.argv[1]
    f = sys.argv[2] if len(sys.argv) > 2 else None
    try:
        result = run_felo_ppt_task(query=q, file_path=f)
        print("\n=== FELO GENERATION RESULT ===")
        print(json.dumps(result, indent=2, ensure_ascii=False))
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
