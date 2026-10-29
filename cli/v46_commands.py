# -*- coding: utf-8 -*-
"""
cli/v46_commands.py — v4.6.0 Enterprise Subcommands & CLI Extensions.

Provides CLI handlers and parser extensions for:
  1. pdf-tools extensions: `crypto-sign` (PAdES/PKCS#7) and `pdfa` (ISO 19005-2 PDF/A-2b)
  2. doc-security extensions: `--deep-cdr`, `--redact-pii`, `--pii-types`
  3. `xlsx-to-brief`: Convert Excel workbooks (.xlsx) into Word (.docx), Markdown (.md), or PDF (.pdf) reports
  4. `xlsx-consolidate`: Merge multiple departmental/branch Excel workbooks into a Master Workbook
  5. `run-pipeline`: Execute declarative multi-step workflows (single file or `--batch-dir` folder)
  6. `code-to-diagram`: AST Code-to-Diagram generator (Python / SQL / Prisma -> Class / ERD / Draw.io / SVG / PNG)
  7. `desktop-web`: Launch Native Desktop App Mode Web Studio with isolated `--user-data-dir`
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any, Callable


def cmd_xlsx_to_brief(args: argparse.Namespace) -> int:
    from core.xlsx.reporting import excel_to_brief

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Excel workbook not found: {src}", file=sys.stderr)
        return 1

    try:
        res = excel_to_brief(
            xlsx_path=src,
            output_path=args.output,
            target_format=args.format,
            title=args.title,
            include_kpi_summary=not getattr(args, "no_kpi", False),
            max_rows_per_sheet=int(getattr(args, "max_rows", 500)),
        )
        if getattr(args, "json", False):
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            print(
                f"[OK] Excel-to-Brief ({res['format'].upper()}) generated ({res['sheets_processed']} sheet(s)) -> {res['output_path']}"
            )
        return 0
    except Exception as exc:
        print(f"[ERROR] excel-to-brief failed: {exc}", file=sys.stderr)
        return 2


def cmd_xlsx_consolidate(args: argparse.Namespace) -> int:
    from core.xlsx.reporting import consolidate_workbooks

    try:
        res = consolidate_workbooks(
            input_paths=args.inputs,
            output_path=args.output,
            mode=args.mode,
            strict_schema=not getattr(args, "relaxed_schema", False),
        )
        if getattr(args, "json", False):
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            print(f"[OK] Consolidated {res['workbooks_merged']} workbook(s) ({res['mode']}) -> {res['output_path']}")
        return 0
    except Exception as exc:
        print(f"[ERROR] xlsx-consolidate failed: {exc}", file=sys.stderr)
        return 2


def cmd_run_pipeline(args: argparse.Namespace) -> int:
    from core.pipeline import run_bulk_directory_pipeline, run_pipeline

    try:
        if getattr(args, "batch_dir", None):
            out_dir = args.output or os.path.join(os.path.abspath(args.batch_dir), "pipeline_output")
            res = run_bulk_directory_pipeline(
                input_dir=args.batch_dir,
                output_dir=out_dir,
                pipeline_spec=args.spec,
                file_pattern=getattr(args, "pattern", "*.*"),
                max_workers=getattr(args, "workers", None),
            )
            if getattr(args, "json", False):
                print(json.dumps(res, ensure_ascii=False, indent=2))
            else:
                print(
                    f"[OK] Bulk Pipeline '{res['pipeline_name']}' finished: {res['succeeded']}/{res['total_files']} succeeded "
                    f"(workers={res['max_workers_used']}) -> {res['output_dir']}"
                )
            return 0 if res.get("success") else 1

        if not args.input:
            print("[ERROR] Either --input <file> or --batch-dir <folder> is required.", file=sys.stderr)
            return 1

        res = run_pipeline(
            input_path=args.input,
            pipeline_spec=args.spec,
            output_path=args.output,
        )
        if getattr(args, "json", False):
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            print(
                f"[OK] Pipeline '{res['pipeline_name']}' executed {res['steps_executed']} step(s) in {res['elapsed_sec']}s -> {res['output_file']}"
            )
        return 0
    except Exception as exc:
        print(f"[ERROR] run-pipeline failed: {exc}", file=sys.stderr)
        return 2


def cmd_code_to_diagram(args: argparse.Namespace) -> int:
    from core.diagrams.code_to_diagram import generate_diagram_from_code

    try:
        res = generate_diagram_from_code(
            source_path=args.source,
            diagram_type=args.type,
            output_spec_path=args.spec_out,
            output_png_path=args.output,
            output_drawio_path=args.drawio,
            output_svg_path=args.svg,
            inject_into=args.inject,
            title=args.title,
        )
        if getattr(args, "json", False):
            print(json.dumps(res, ensure_ascii=False, indent=2))
        else:
            print(f"[OK] Generated {res['diagram_type'].upper()} from source AST:")
            if res.get("drawio_path"):
                print(f"  + Draw.io : {res['drawio_path']}")
            if res.get("svg_path"):
                print(f"  + Vector  : {res['svg_path']}")
            if res.get("png_path"):
                print(f"  + PNG     : {res['png_path']}")
            if res.get("injected_into"):
                print(f"  + Injected: {res['injected_into']}")
        return 0
    except Exception as exc:
        print(f"[ERROR] code-to-diagram failed: {exc}", file=sys.stderr)
        return 2


def cmd_desktop_web(args: argparse.Namespace) -> int:
    from ui.webview_launcher import launch_native_web_studio

    res = launch_native_web_studio(
        port=int(getattr(args, "port", 0) or 0),
        dry_run=bool(getattr(args, "dry_run", False)),
    )
    if getattr(args, "dry_run", False) or getattr(args, "json", False):
        print(json.dumps(res, ensure_ascii=False, indent=2))
    return 0


def wrap_pdf_tools_handler(orig_handler: Callable[[argparse.Namespace], int]) -> Callable[[argparse.Namespace], int]:
    """Extends cmd_pdf_tools with `crypto-sign` and `pdfa` actions."""
    def _wrapped(args: argparse.Namespace) -> int:
        action = getattr(args, "action", "")
        if action in ("crypto-sign", "crypto_sign"):
            from core.converter.pdf_toolkit import crypto_sign_pdf

            if not args.input:
                print("[ERROR] 'crypto-sign' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_crypto_signed.pdf"
            res = crypto_sign_pdf(
                args.input,
                out,
                signer_name=getattr(args, "signer", None) or "Antigravity Authority",
                reason=getattr(args, "reason", None) or "Phê duyệt & Xác thực tài liệu (PAdES/PKCS#7)",
                location=getattr(args, "location", None) or "Việt Nam",
                pfx_path=getattr(args, "pfx", None),
                pfx_password=getattr(args, "pfx_pwd", None),
                stamp_image=getattr(args, "stamp", None) or getattr(args, "signature", None),
                page=getattr(args, "page", "last"),
            )
            print(f"[OK] Cryptographically signed PDF (SHA-256: {res['sha256_digest'][:16]}...): {res['output_path']}")
            return 0

        if action in ("pdfa", "to-pdfa", "pdfa-2b"):
            from core.converter.pdf_toolkit import convert_to_pdfa

            if not args.input:
                print("[ERROR] 'pdfa' requires an input file (-i file.pdf)", file=sys.stderr)
                return 1
            out = args.output or os.path.splitext(args.input)[0] + "_pdfa2b.pdf"
            res = convert_to_pdfa(
                args.input,
                out,
                part=int(getattr(args, "pdfa_part", 2) or 2),
                conformance=getattr(args, "pdfa_conf", "B") or "B",
            )
            print(f"[OK] Converted PDF to archival {res['pdfa_standard']}: {res['output_path']}")
            return 0

        return orig_handler(args)

    return _wrapped


def wrap_doc_security_handler(orig_handler: Callable[[argparse.Namespace], int]) -> Callable[[argparse.Namespace], int]:
    """Extends cmd_doc_security with `--deep-cdr`, `--redact-pii`, and `--pii-types`."""
    def _wrapped(args: argparse.Namespace) -> int:
        if getattr(args, "deep_cdr", False) or getattr(args, "redact_pii", False):
            from core.security.sanitizer import DocumentSanitizer

            sanitizer = DocumentSanitizer()
            out_p = getattr(args, "output", None)
            raw_pii = getattr(args, "pii_types", None)
            pii_list = [x.strip() for x in raw_pii.split(",") if x.strip()] if raw_pii else None
            san_res = sanitizer.sanitize(
                args.file,
                output_path=out_p,
                deep_cdr=bool(getattr(args, "deep_cdr", True)),
                redact_pii=bool(getattr(args, "redact_pii", False)),
                pii_types=pii_list,
            )
            print(f"[SECURITY DEEP CDR] Cleaned file saved to: {san_res.get('output_file')}")
            for act in san_res.get("actions_taken", []):
                print(f"  + {act}")
            # Prevent double-running basic sanitize in orig_handler
            args.sanitize = False
        return orig_handler(args)

    return _wrapped


def register_v46_subcommands(
    sub: Any,
    p_pdft: argparse.ArgumentParser,
    p_sec: argparse.ArgumentParser,
    orig_pdf_func: Callable[[argparse.Namespace], int],
    orig_sec_func: Callable[[argparse.Namespace], int],
) -> None:
    """Registers v4.6 subcommands and extends pdf-tools and doc-security parsers."""
    # 1. Extend pdf-tools choices & flags
    for act in p_pdft._actions:
        if act.dest == "action" and act.choices is not None:
            choices_list = list(act.choices)
            for extra in ("crypto-sign", "pdfa"):
                if extra not in choices_list:
                    choices_list.append(extra)
            act.choices = choices_list
    p_pdft.add_argument("--signer", default="Antigravity Authority", help="Signer name for crypto-sign")
    p_pdft.add_argument("--reason", default="Phê duyệt & Xác thực tài liệu (PAdES/PKCS#7)", help="Signing reason")
    p_pdft.add_argument("--location", default="Việt Nam", help="Signing location")
    p_pdft.add_argument("--pfx", default=None, help="Optional PKCS#12 (.pfx/.p12) certificate path")
    p_pdft.add_argument("--pfx-pwd", default=None, help="Optional PKCS#12 password")
    p_pdft.add_argument("--pdfa-part", type=int, default=2, choices=[1, 2, 3], help="PDF/A standard part (default: 2)")
    p_pdft.add_argument("--pdfa-conf", default="B", choices=["A", "B", "U"], help="PDF/A conformance level (default: B)")
    p_pdft.set_defaults(func=wrap_pdf_tools_handler(orig_pdf_func))

    # 2. Extend doc-security flags
    p_sec.add_argument("--deep-cdr", action="store_true", help="Enable Deep CDR (strip vbaProject.bin, OLE MZ binaries, PDF JS, and promote ContentType)")
    p_sec.add_argument("--redact-pii", action="store_true", help="Redact Vietnamese PII (CCCD, Phone, Contextual Bank Account, MST, Email)")
    p_sec.add_argument("--pii-types", default=None, help="Comma-separated PII types to redact (cccd,phone,bank_account,mst,email)")
    p_sec.set_defaults(func=wrap_doc_security_handler(orig_sec_func))

    # 3. Subcommand: xlsx-to-brief
    p_xtb = sub.add_parser("xlsx-to-brief", help="Convert Excel workbook (.xlsx) into Word (.docx), Markdown (.md), or PDF (.pdf) report")
    p_xtb.add_argument("file", help="Path to source .xlsx workbook")
    p_xtb.add_argument("-o", "--output", default=None, help="Output report path (.docx, .md, or .pdf)")
    p_xtb.add_argument("-f", "--format", default="docx", choices=["docx", "md", "pdf"], help="Target report format (default: docx)")
    p_xtb.add_argument("--title", default=None, help="Custom report title")
    p_xtb.add_argument("--no-kpi", action="store_true", help="Omit KPI summary tables")
    p_xtb.add_argument("--max-rows", type=int, default=500, help="Maximum rows per sheet (default: 500)")
    p_xtb.add_argument("--json", action="store_true", help="Output JSON summary")
    p_xtb.set_defaults(func=cmd_xlsx_to_brief)

    # 4. Subcommand: xlsx-consolidate
    p_xc = sub.add_parser("xlsx-consolidate", help="Consolidate multiple Excel workbooks into a Master Workbook")
    p_xc.add_argument("--inputs", nargs="+", required=True, help="Input .xlsx file paths to consolidate")
    p_xc.add_argument("-o", "--output", required=True, help="Output Master .xlsx path")
    p_xc.add_argument("--mode", default="append_rows", choices=["append_rows", "merge_sheets"], help="Consolidation mode")
    p_xc.add_argument("--relaxed-schema", action="store_true", help="Allow minor column header differences without failing")
    p_xc.add_argument("--json", action="store_true", help="Output JSON summary")
    p_xc.set_defaults(func=cmd_xlsx_consolidate)

    # 5. Subcommand: run-pipeline
    p_rp = sub.add_parser("run-pipeline", help="Execute a declarative multi-step workflow pipeline on a file or folder")
    p_rp.add_argument("-s", "--spec", required=True, help="Path to pipeline JSON/YAML spec file or raw JSON string")
    p_rp.add_argument("-i", "--input", default=None, help="Single input file path")
    p_rp.add_argument("--batch-dir", default=None, help="Input directory for bulk processing")
    p_rp.add_argument("-o", "--output", default=None, help="Output file path (single mode) or output directory (batch mode)")
    p_rp.add_argument("--pattern", default="*.*", help="File glob pattern for --batch-dir (default: *.*)")
    p_rp.add_argument("--workers", type=int, default=None, help="Max worker threads (default: MAX_JOB_WORKERS env or 2)")
    p_rp.add_argument("--json", action="store_true", help="Output JSON execution trace")
    p_rp.set_defaults(func=cmd_run_pipeline)

    # 6. Subcommand: code-to-diagram
    p_ctd = sub.add_parser("code-to-diagram", help="AST Code-to-Diagram generator (Python/SQL/Prisma -> Class/ERD/Draw.io/SVG/PNG)")
    p_ctd.add_argument("source", help="Path to source file (.py, .sql, .prisma) or package directory")
    p_ctd.add_argument("-t", "--type", default="auto", choices=["auto", "class", "erd", "c4"], help="Diagram type (default: auto)")
    p_ctd.add_argument("-o", "--output", default=None, help="Output PNG path")
    p_ctd.add_argument("--drawio", default=None, help="Output native .drawio path")
    p_ctd.add_argument("--svg", default=None, help="Output pure vector .svg path")
    p_ctd.add_argument("--spec-out", default=None, help="Output declarative JSON spec path")
    p_ctd.add_argument("--inject", default=None, help="Target .docx or .pptx file to auto-inject diagram into")
    p_ctd.add_argument("--title", default=None, help="Diagram title / caption for injection")
    p_ctd.add_argument("--json", action="store_true", help="Output JSON metadata")
    p_ctd.set_defaults(func=cmd_code_to_diagram)

    # 7. Subcommand: desktop-web
    p_dw = sub.add_parser("desktop-web", help="Launch Native Desktop App Mode Web Studio with isolated --user-data-dir")
    p_dw.add_argument("--port", type=int, default=0, help="Local server port (default: auto free port)")
    p_dw.add_argument("--dry-run", action="store_true", help="Verify server & print command without blocking on UI")
    p_dw.add_argument("--json", action="store_true", help="Output JSON launch info")
    p_dw.set_defaults(func=cmd_desktop_web)

