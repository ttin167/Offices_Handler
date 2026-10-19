"""
tools/bump_version.py — Automated Version Synchronization & Invariant Audit Tool.

Invariant: ERR_VER_001 (Version Desync Invariant)
Ensures 100% dual-surface synchronization between Terminal/Backend/Manifests and Frontend/UI.

Usage:
  python tools/bump_version.py 4.2.0           # Bump and synchronize all files + rebuild frontend
  python tools/bump_version.py --check        # Audit all files for version consistency (exit 0 if clean, 1 if desynced)
  python tools/bump_version.py --get          # Print current project version
"""

from __future__ import annotations
import json
import os
import re
import subprocess
import sys
from typing import Dict, List, Tuple

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_current_package_version() -> str:
    pkg_path = os.path.join(BASE_DIR, "package.json")
    with open(pkg_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("version", "4.0.0")


def audit_version_consistency(target_version: str) -> Tuple[bool, List[str]]:
    """
    Checks all manifest, terminal, server, and frontend files for version synchronization.
    Returns (is_synced, list_of_errors).
    """
    errors: List[str] = []

    # 1. package.json
    pkg_path = os.path.join(BASE_DIR, "package.json")
    if os.path.isfile(pkg_path):
        with open(pkg_path, "r", encoding="utf-8") as f:
            v = json.load(f).get("version")
        if v != target_version:
            errors.append(f"package.json: expected '{target_version}' got '{v}'")

    # 2. package-lock.json
    lock_path = os.path.join(BASE_DIR, "package-lock.json")
    if os.path.isfile(lock_path):
        with open(lock_path, "r", encoding="utf-8") as f:
            lock_data = json.load(f)
            v_root = lock_data.get("version")
            v_pkg = lock_data.get("packages", {}).get("", {}).get("version")
        if v_root != target_version or v_pkg != target_version:
            errors.append(f"package-lock.json: expected '{target_version}' got root='{v_root}', pkg='{v_pkg}'")

    # 3. pyproject.toml
    pyproject_path = os.path.join(BASE_DIR, "pyproject.toml")
    if os.path.isfile(pyproject_path):
        with open(pyproject_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'version\s*=\s*["\']([^"\']+)["\']', content)
        if not m or m.group(1) != target_version:
            errors.append(f"pyproject.toml: expected '{target_version}' got '{m.group(1) if m else 'None'}'")

    # 4. dev_launcher.js
    launcher_path = os.path.join(BASE_DIR, "dev_launcher.js")
    if os.path.isfile(launcher_path):
        with open(launcher_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r"let\s+pkgVersion\s*=\s*['\"]([^'\"]+)['\"]", content)
        if not m or m.group(1) != target_version:
            errors.append(f"dev_launcher.js: expected '{target_version}' got '{m.group(1) if m else 'None'}'")

    # 5. server/app.py
    app_path = os.path.join(BASE_DIR, "server", "app.py")
    if os.path.isfile(app_path):
        with open(app_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'APP_VERSION\s*=\s*["\']([^"\']+)["\']', content)
        if not m or m.group(1) != target_version:
            errors.append(f"server/app.py: expected APP_VERSION = '{target_version}' got '{m.group(1) if m else 'None'}'")

    # 6. server/mcp_server.py
    mcp_path = os.path.join(BASE_DIR, "server", "mcp_server.py")
    if os.path.isfile(mcp_path):
        with open(mcp_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'SERVER_VERSION\s*=\s*["\']([^"\']+)["\']', content)
        if not m or m.group(1) != target_version:
            errors.append(f"server/mcp_server.py: expected SERVER_VERSION = '{target_version}' got '{m.group(1) if m else 'None'}'")

    # 7. index.html
    index_path = os.path.join(BASE_DIR, "index.html")
    if os.path.isfile(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            content = f.read()
        expected_sub = f"Office Studio v{target_version}"
        if expected_sub not in content:
            errors.append(f"index.html: missing expected '{expected_sub}' in brand subtitle")

    # 8. core/env_doctor.py
    doctor_path = os.path.join(BASE_DIR, "core", "env_doctor.py")
    if os.path.isfile(doctor_path):
        with open(doctor_path, "r", encoding="utf-8") as f:
            content = f.read()
        m = re.search(r'DOCTOR_VERSION\s*=\s*["\']([^"\']+)["\']', content)
        if not m or m.group(1) != target_version:
            errors.append(f"core/env_doctor.py: expected DOCTOR_VERSION = '{target_version}' got '{m.group(1) if m else 'None'}'")

    return (len(errors) == 0, errors)


def sync_version(new_version: str, rebuild_fe: bool = True) -> bool:
    """
    Atomically updates version across all manifest, backend, CLI, and frontend files.
    Optionally re-runs 'npm run build' to update production bundle in dist/.
    """
    print(f"\n🚀 [BUMP VERSION] Synchronizing all components to v{new_version}...\n")
    changes: List[str] = []

    # 1. package.json
    pkg_path = os.path.join(BASE_DIR, "package.json")
    if os.path.isfile(pkg_path):
        with open(pkg_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        data["version"] = new_version
        with open(pkg_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
            f.write("\n")
        changes.append(f"Updated package.json -> {new_version}")

    # 2. package-lock.json
    lock_path = os.path.join(BASE_DIR, "package-lock.json")
    if os.path.isfile(lock_path):
        with open(lock_path, "r", encoding="utf-8") as f:
            lock_data = json.load(f)
        lock_data["version"] = new_version
        if "packages" in lock_data and "" in lock_data["packages"]:
            lock_data["packages"][""]["version"] = new_version
        with open(lock_path, "w", encoding="utf-8") as f:
            json.dump(lock_data, f, indent=2)
            f.write("\n")
        changes.append(f"Updated package-lock.json -> {new_version}")

    # 3. pyproject.toml
    pyproject_path = os.path.join(BASE_DIR, "pyproject.toml")
    if os.path.isfile(pyproject_path):
        with open(pyproject_path, "r", encoding="utf-8") as f:
            txt = f.read()
        txt = re.sub(r'version\s*=\s*["\'][^"\']+["\']', f'version = "{new_version}"', txt, count=1)
        with open(pyproject_path, "w", encoding="utf-8") as f:
            f.write(txt)
        changes.append(f"Updated pyproject.toml -> {new_version}")

    # 4. dev_launcher.js
    launcher_path = os.path.join(BASE_DIR, "dev_launcher.js")
    if os.path.isfile(launcher_path):
        with open(launcher_path, "r", encoding="utf-8") as f:
            txt = f.read()
        txt = re.sub(r"let\s+pkgVersion\s*=\s*['\"][^'\"]+['\"];", f"let pkgVersion = '{new_version}';", txt, count=1)
        with open(launcher_path, "w", encoding="utf-8") as f:
            f.write(txt)
        changes.append(f"Updated dev_launcher.js fallback -> {new_version}")

    # 5. server/app.py
    app_path = os.path.join(BASE_DIR, "server", "app.py")
    if os.path.isfile(app_path):
        with open(app_path, "r", encoding="utf-8") as f:
            txt = f.read()
        if "APP_VERSION = " in txt:
            txt = re.sub(r'APP_VERSION\s*=\s*["\'][^"\']+["\']', f'APP_VERSION = "{new_version}"', txt, count=1)
        else:
            txt = f'APP_VERSION = "{new_version}"\n' + txt
        # Also replace in /api/health and /api/info
        txt = re.sub(r'"version":\s*"[^"]+"', f'"version": "{new_version}"', txt)
        with open(app_path, "w", encoding="utf-8") as f:
            f.write(txt)
        changes.append(f"Updated server/app.py (APP_VERSION, /api/health) -> {new_version}")

    # 6. server/mcp_server.py
    mcp_path = os.path.join(BASE_DIR, "server", "mcp_server.py")
    if os.path.isfile(mcp_path):
        with open(mcp_path, "r", encoding="utf-8") as f:
            txt = f.read()
        txt = re.sub(r'SERVER_VERSION\s*=\s*["\'][^"\']+["\']', f'SERVER_VERSION = "{new_version}"', txt, count=1)
        with open(mcp_path, "w", encoding="utf-8") as f:
            f.write(txt)
        changes.append(f"Updated server/mcp_server.py -> {new_version}")

    # 7. core/env_doctor.py
    doctor_path = os.path.join(BASE_DIR, "core", "env_doctor.py")
    if os.path.isfile(doctor_path):
        with open(doctor_path, "r", encoding="utf-8") as f:
            txt = f.read()
        if "DOCTOR_VERSION = " in txt:
            txt = re.sub(r'DOCTOR_VERSION\s*=\s*["\'][^"\']+["\']', f'DOCTOR_VERSION = "{new_version}"', txt, count=1)
        else:
            txt = f'DOCTOR_VERSION = "{new_version}"\n' + txt
        txt = re.sub(r'ANTIGRAVITY OFFICE ENVIRONMENT DOCTOR \(v[^\)]+\)', f'ANTIGRAVITY OFFICE ENVIRONMENT DOCTOR (v{new_version})', txt)
        with open(doctor_path, "w", encoding="utf-8") as f:
            f.write(txt)
        changes.append(f"Updated core/env_doctor.py -> {new_version}")

    # 8. index.html
    index_path = os.path.join(BASE_DIR, "index.html")
    if os.path.isfile(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            txt = f.read()
        txt = re.sub(r'<span class="brand-subtitle">Office Studio v[^<]+</span>', f'<span class="brand-subtitle">Office Studio v{new_version}</span>', txt)
        with open(index_path, "w", encoding="utf-8") as f:
            f.write(txt)
        changes.append(f"Updated index.html brand subtitle -> Office Studio v{new_version}")

    # 9. README.md badge
    readme_path = os.path.join(BASE_DIR, "README.md")
    if os.path.isfile(readme_path):
        with open(readme_path, "r", encoding="utf-8") as f:
            txt = f.read()
        txt = re.sub(r'Release-v[^_]+_Production', f'Release-v{new_version}_Production', txt)
        txt = re.sub(r'Release:\s*v[^\)\]]+', f'Release: v{new_version}', txt)
        with open(readme_path, "w", encoding="utf-8") as f:
            f.write(txt)
        changes.append(f"Updated README.md badge -> v{new_version}")

    for c in changes:
        print(f"  [OK] {c}")

    if rebuild_fe:
        print("\n🔨 [BUILD] Recompiling Vite production bundle (npm run build)...")
        shell_cmd = "npm.cmd run build" if sys.platform == "win32" else "npm run build"
        res = subprocess.run(shell_cmd, shell=True, cwd=BASE_DIR, capture_output=True, text=True)
        if res.returncode == 0:
            print("  [OK] Frontend production bundle rebuilt successfully in dist/.")
        else:
            print(f"  [WARN] Failed to rebuild bundle: {res.stderr or res.stdout}")

    # Self-audit check
    is_clean, errs = audit_version_consistency(new_version)
    if is_clean:
        print(f"\n✅ [SUCCESS] All components verified and 100% synchronized with v{new_version}!\n")
        return True
    else:
        print(f"\n❌ [ERR_VER_001] Incomplete synchronization detected:")
        for e in errs:
            print(f"  - {e}")
        return False


def main() -> int:
    args = sys.argv[1:]

    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    if args[0] == "--get":
        print(get_current_package_version())
        return 0

    if args[0] == "--check":
        target = args[1] if len(args) > 1 else get_current_package_version()
        is_clean, errs = audit_version_consistency(target)
        if is_clean:
            print(f"[OK] Invariant ERR_VER_001 PASSED: All components synchronized with v{target}.")
            return 0
        else:
            print(f"[FAIL] Invariant ERR_VER_001 VIOLATION: Version desynchronization detected for v{target}:")
            for e in errs:
                print(f"  - {e}")
            return 1

    # Bump mode
    new_ver = args[0].lstrip("v").strip()
    if not re.match(r"^\d+\.\d+\.\d+", new_ver):
        print(f"[ERROR] Invalid version format: '{new_ver}'. Expected format: X.Y.Z (e.g. 4.2.0)")
        return 1

    no_build = "--no-build" in args
    success = sync_version(new_ver, rebuild_fe=not no_build)
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
