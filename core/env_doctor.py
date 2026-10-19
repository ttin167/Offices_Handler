"""
core.env_doctor — Zero-Setup Environment Doctor & Binary Validator (v4.0).

Scans and validates external runtime dependencies across Windows, Linux, and macOS:
1. Java Runtime (java) & plantuml.jar
2. Chromium headless engine (Microsoft Edge / Google Chrome)
3. Document Export Engines (Microsoft Word COM, LibreOffice)
4. OCR Engine (Tesseract & Vietnamese language data)
5. Core Python Libraries (python-docx, python-pptx, openpyxl, pymupdf, Pillow)

Conforms to Clean Architecture: 100% self-contained in core/, pure dictionary output.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from typing import Any, Dict, List, Optional

_CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
_WORKSPACE_ROOT = os.path.dirname(_CURRENT_DIR)


class EnvironmentDoctor:
    """Diagnostic scanner for external binaries and runtimes."""

    def __init__(self, workspace_root: Optional[str] = None):
        self.workspace_root = workspace_root or _WORKSPACE_ROOT

    def check_java_plantuml(self) -> Dict[str, Any]:
        """Check Java JRE/JDK and bundled plantuml.jar."""
        java_cmd = shutil.which("java") or shutil.which("java.exe")
        has_java = bool(java_cmd)
        java_version = ""

        if has_java:
            try:
                proc = subprocess.run(
                    [java_cmd, "-version"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                    errors="replace"
                )
                output = proc.stderr or proc.stdout
                first_line = output.splitlines()[0] if output else ""
                java_version = first_line.strip()
            except Exception as exc:
                java_version = f"Error querying version: {exc}"

        # Locate plantuml.jar
        jar_candidates = [
            os.path.join(self.workspace_root, "tools", "plantuml.jar"),
            os.path.join(os.getcwd(), "tools", "plantuml.jar"),
            os.environ.get("PLANTUML_JAR", ""),
            shutil.which("plantuml.jar") or "",
        ]
        jar_path = next((c for c in jar_candidates if c and os.path.isfile(c)), None)
        jar_size_mb = round(os.path.getsize(jar_path) / (1024 * 1024), 2) if jar_path else 0.0

        # Check Graphviz dot binary (required for PlantUML Class, C4, UseCase, Activity diagrams)
        dot_cmd = shutil.which("dot") or shutil.which("dot.exe")
        has_dot = bool(dot_cmd)
        dot_version = ""
        if has_dot:
            try:
                proc = subprocess.run(
                    [dot_cmd, "-V"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                    errors="replace"
                )
                dot_version = (proc.stderr or proc.stdout).strip()
            except Exception:
                dot_version = "Detected"

        status = "healthy" if (has_java and jar_path) else "degraded"
        issues = []
        if not has_java:
            issues.append("Java Runtime (JRE >= 8) not found in PATH.")
        if not jar_path:
            issues.append("plantuml.jar not found in tools/plantuml.jar.")
        if not has_dot:
            issues.append("Graphviz 'dot' executable not detected. Non-sequence PlantUML diagrams require Graphviz.")

        return {
            "status": status,
            "has_java": has_java,
            "java_cmd": java_cmd,
            "java_version": java_version,
            "has_jar": bool(jar_path),
            "jar_path": jar_path,
            "jar_size_mb": jar_size_mb,
            "has_dot": has_dot,
            "dot_cmd": dot_cmd,
            "dot_version": dot_version,
            "issues": issues,
        }

    def check_chromium(self) -> Dict[str, Any]:
        """Check Microsoft Edge or Google Chrome for headless rendering."""
        candidates = []
        env_chromium = os.environ.get("CHROMIUM_PATH")
        if env_chromium:
            candidates.append(env_chromium)

        if sys.platform == "win32":
            candidates.extend([
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                shutil.which("msedge"),
                shutil.which("chrome"),
            ])
        elif sys.platform == "darwin":
            candidates.extend([
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
                shutil.which("google-chrome"),
                shutil.which("chromium"),
            ])
        else:
            candidates.extend([
                shutil.which("google-chrome"),
                shutil.which("google-chrome-stable"),
                shutil.which("chromium"),
                shutil.which("chromium-browser"),
                "/usr/bin/chromium",
                "/usr/bin/chromium-browser",
            ])

        found_path = next((c for c in candidates if c and os.path.isfile(c)), None)
        version_str = ""

        if found_path:
            try:
                proc = subprocess.run(
                    [found_path, "--version"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                    errors="replace"
                )
                version_str = (proc.stdout or proc.stderr).strip()
            except Exception:
                version_str = "Detected (version check failed)"

        status = "healthy" if found_path else "degraded"
        issues = []
        if not found_path:
            issues.append("Chromium runtime (Microsoft Edge or Google Chrome) not detected. Headless canvas & Draw.io rendering disabled.")

        return {
            "status": status,
            "found": bool(found_path),
            "path": found_path,
            "version": version_str,
            "issues": issues,
        }

    def check_chromium_runner(self) -> Dict[str, Any]:
        """Alias for check_chromium for naming consistency."""
        return self.check_chromium()

    def check_office_converters(self) -> Dict[str, Any]:
        """Check Microsoft Word COM and LibreOffice for vector PDF conversion."""
        word_com_available = False
        word_version = ""

        if sys.platform == "win32":
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as key:
                    curver, _ = winreg.QueryValueEx(key, "")
                    word_version = f"Microsoft Word (v{curver.split('.')[-1]}.0)"
                    word_com_available = True
            except Exception:
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                    import win32com.client as win32
                    w = win32.Dispatch("Word.Application")
                    word_version = f"Microsoft Word v{w.Version}"
                    w.Quit()
                    del w
                    pythoncom.CoUninitialize()
                    word_com_available = True
                except Exception as exc:
                    word_version = f"COM Unavailable ({exc})"

        soffice_candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.com",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            shutil.which("soffice"),
            shutil.which("libreoffice"),
        ]
        soffice_path = next((c for c in soffice_candidates if c and os.path.isfile(c)), None)

        status = "healthy" if (word_com_available or soffice_path) else "degraded"
        issues = []
        if not word_com_available and not soffice_path:
            issues.append("Neither Microsoft Word COM nor LibreOffice detected. PDF export will fall back to Edge Headless print.")

        return {
            "status": status,
            "word_com_available": word_com_available,
            "word_version": word_version,
            "libreoffice_available": bool(soffice_path),
            "libreoffice_path": soffice_path,
            "issues": issues,
        }

    def check_ocr_tesseract(self) -> Dict[str, Any]:
        """Check Tesseract OCR executable and Vietnamese language models."""
        candidates = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            shutil.which("tesseract"),
        ]
        tess_path = next((c for c in candidates if c and os.path.isfile(c)), None)
        languages: List[str] = []

        if tess_path:
            try:
                proc = subprocess.run(
                    [tess_path, "--list-langs"],
                    capture_output=True,
                    text=True,
                    timeout=5,
                    errors="replace"
                )
                output = proc.stdout or proc.stderr
                lines = [l.strip() for l in output.splitlines() if l.strip() and not l.startswith("List of")]
                languages = lines
            except Exception:
                pass

        has_vie = "vie" in languages
        status = "healthy" if (tess_path and has_vie) else "optional"
        issues = []
        if not tess_path:
            issues.append("Tesseract OCR is not installed (OCR features optional).")
        elif not has_vie:
            issues.append("Tesseract installed, but Vietnamese traineddata ('vie') not found in tessdata.")

        return {
            "status": status,
            "found": bool(tess_path),
            "path": tess_path,
            "languages": languages,
            "has_vietnamese": has_vie,
            "issues": issues,
        }

    def check_python_dependencies(self) -> Dict[str, Any]:
        """Verify core document processing Python libraries."""
        deps = {
            "docx": "python-docx",
            "pptx": "python-pptx",
            "openpyxl": "openpyxl",
            "fitz": "PyMuPDF",
            "PIL": "Pillow",
            "yaml": "PyYAML",
        }
        installed = {}
        missing = []

        for mod, pkg in deps.items():
            try:
                m = __import__(mod)
                ver = getattr(m, "__version__", "installed")
                installed[pkg] = ver
            except ImportError:
                missing.append(pkg)

        status = "healthy" if not missing else "critical"
        return {
            "status": status,
            "installed": installed,
            "missing": missing,
            "issues": [f"Missing required Python package: {p}" for p in missing],
        }

    def run_diagnostics(self) -> Dict[str, Any]:
        """Run complete environment diagnostic suite and aggregate results."""
        t0 = time.time()
        py_deps = self.check_python_dependencies()
        java_puml = self.check_java_plantuml()
        chromium = self.check_chromium()
        office = self.check_office_converters()
        ocr = self.check_ocr_tesseract()

        all_issues = []
        all_issues.extend(py_deps["issues"])
        all_issues.extend(java_puml["issues"])
        all_issues.extend(chromium["issues"])
        all_issues.extend(office["issues"])

        # Determine overall system health
        if py_deps["status"] == "critical":
            overall_status = "critical"
        elif any(c["status"] == "degraded" for c in (java_puml, chromium, office)):
            overall_status = "degraded"
        else:
            overall_status = "healthy"

        recommendations = []
        if not java_puml["has_java"]:
            recommendations.append("Install Java Runtime (e.g. Eclipse Temurin or OpenJDK 17+) to enable PlantUML diagram rendering.")
        if not chromium["found"]:
            recommendations.append("Ensure Microsoft Edge or Google Chrome is installed for Full HD Draw.io & Canvas rendering.")
        if not office["word_com_available"] and not office["libreoffice_available"]:
            recommendations.append("Install Microsoft Word or LibreOffice for 100% pixel-perfect vector PDF conversion.")
        if not ocr["found"]:
            recommendations.append("Optional: Install Tesseract OCR with Vietnamese language pack for scanned document processing.")

        return {
            "status": overall_status,
            "timestamp": time.time(),
            "elapsed_sec": round(time.time() - t0, 3),
            "platform": sys.platform,
            "python_version": sys.version.split()[0],
            "components": {
                "python_dependencies": py_deps,
                "java_plantuml": java_puml,
                "chromium_runner": chromium,
                "office_converters": office,
                "ocr_tesseract": ocr,
            },
            "issues_count": len(all_issues),
            "issues": all_issues,
            "recommendations": recommendations,
        }


def run_doctor() -> Dict[str, Any]:
    """Convenience functional API returning environment diagnostic dict."""
    return EnvironmentDoctor().run_diagnostics()


def format_doctor_report(report: Dict[str, Any]) -> str:
    """Format diagnostic report into human-readable terminal output string."""
    status_tag = "[HEALTHY]" if report["status"] == "healthy" else (
        "[DEGRADED]" if report["status"] == "degraded" else "[CRITICAL]"
    )
    lines = [
        "",
        "=" * 65,
        f" {status_tag} ANTIGRAVITY OFFICE ENVIRONMENT DOCTOR (v4.0)",
        "=" * 65,
        f"OS Platform    : {report['platform']} (Python {report['python_version']})",
        f"Overall Status : {report['status'].upper()} ({report['elapsed_sec']}s)",
        "-" * 65,
    ]

    comp = report["components"]

    # 1. Python Dependencies
    py_ok = comp["python_dependencies"]["status"] == "healthy"
    lines.append(f"[{'OK' if py_ok else 'FAIL'}] Python Libraries:")
    for pkg, ver in comp["python_dependencies"]["installed"].items():
        lines.append(f"     + {pkg:15s} : {ver}")
    for m in comp["python_dependencies"]["missing"]:
        lines.append(f"     x {m:15s} : MISSING")

    # 2. Java & PlantUML
    jp = comp["java_plantuml"]
    jp_ok = jp["status"] == "healthy"
    lines.append(f"[{'OK' if jp_ok else 'WARN'}] Java Runtime & PlantUML:")
    lines.append(f"     + Java          : {jp['java_version'] or 'Not found'}")
    lines.append(f"     + plantuml.jar  : {jp['jar_path'] or 'Not found'} ({jp['jar_size_mb']} MB)")

    # 3. Chromium Engine
    cr = comp["chromium_runner"]
    cr_ok = cr["status"] == "healthy"
    lines.append(f"[{'OK' if cr_ok else 'WARN'}] Chromium Canvas Runner:")
    lines.append(f"     + Browser       : {cr['version'] or 'Not found'}")
    lines.append(f"     + Path          : {cr['path'] or 'None'}")

    # 4. Office PDF Converters
    of = comp["office_converters"]
    of_ok = of["status"] == "healthy"
    lines.append(f"[{'OK' if of_ok else 'WARN'}] Office PDF Converters:")
    lines.append(f"     + Word COM      : {'Available' if of['word_com_available'] else 'Unavailable'} ({of['word_version']})")
    lines.append(f"     + LibreOffice   : {of['libreoffice_path'] or 'Not found'}")

    # 5. Tesseract OCR
    ocr = comp["ocr_tesseract"]
    lines.append(f"[{'OK' if ocr['found'] else 'INFO'}] Tesseract OCR (Optional):")
    lines.append(f"     + Executable    : {ocr['path'] or 'Not installed'}")
    lines.append(f"     + Vietnamese    : {'Supported (vie.traineddata)' if ocr['has_vietnamese'] else 'Missing'}")

    # Recommendations
    if report["recommendations"]:
        lines.append("-" * 65)
        lines.append("Recommendations for Maximum Fidelity:")
        for r in report["recommendations"]:
            lines.append(f"   -> {r}")

    lines.append("=" * 65)
    lines.append("")
    return "\n".join(lines)


def print_doctor_report(report: Dict[str, Any]) -> None:
    """Format diagnostic report into human-readable terminal output."""
    if sys.platform == "win32":
        try:
            if hasattr(sys.stdout, "reconfigure"):
                sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            if hasattr(sys.stderr, "reconfigure"):
                sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    print(format_doctor_report(report))


if __name__ == "__main__":
    rep = run_doctor()
    print_doctor_report(rep)

