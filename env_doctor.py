"""
env_doctor.py — Root Facade for Zero-Setup Environment Doctor.
"""

from __future__ import annotations

try:
    from core.env_doctor import EnvironmentDoctor, run_doctor, print_doctor_report, format_doctor_report
except (ImportError, ValueError):
    from .core.env_doctor import EnvironmentDoctor, run_doctor, print_doctor_report, format_doctor_report  # type: ignore

__all__ = ["EnvironmentDoctor", "run_doctor", "print_doctor_report", "format_doctor_report"]

if __name__ == "__main__":
    rep = run_doctor()
    print_doctor_report(rep)
