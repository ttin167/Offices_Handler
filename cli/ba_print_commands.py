"""
cli/ba_print_commands.py — BA Enterprise Workstation & Printing CLI Dispatcher.

Houses command handlers and argument parsers for:
1. BA Hub:
   - ba-story-lint, ba-story-export
   - ba-usecase-lint, ba-usecase-export
   - ba-rtm-matrix
   - ba-dbml-convert
   - ba-satsuma-lineage
   - ba-sql-suite
   - ba-scope-creep
   - ba-obsidian-export, ba-obsidian-compile
   - ba-openapi
   - ba-decision-table
   - ba-meeting-parser
   - wireframe-render
2. Enterprise Printing:
   - print-list
   - print-doc
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any, Dict, List


# ---------------------------------------------------------------------------
# Handlers: OpenAPI, Decision Table, Meeting Parser, Wireframe
# ---------------------------------------------------------------------------

def cmd_ba_openapi(args: argparse.Namespace) -> int:
    from core.ba_hub.data_dictionary import DataDictionaryScanner, FieldMappingEngine, scan_openapi_spec

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] OpenAPI spec not found: {src}", file=sys.stderr)
        return 1

    try:
        data = scan_openapi_spec(src)
        out_format = getattr(args, "format", "excel").lower()
        output_path = args.output

        if out_format == "json":
            if output_path:
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print(f"[OK] OpenAPI JSON written to: {output_path}")
            else:
                print(json.dumps(data, ensure_ascii=False, indent=2))
        elif out_format == "docx" or (output_path and output_path.endswith(".docx")):
            out_file = output_path or os.path.splitext(src)[0] + "_dict.docx"
            FieldMappingEngine.export_openapi_to_docx(data, out_file)
            print(f"[OK] OpenAPI Word Data Dictionary generated: {out_file}")
        else:
            out_file = output_path or os.path.splitext(src)[0] + "_dict.xlsx"
            FieldMappingEngine.export_openapi_to_excel(data, out_file)
            print(f"[OK] OpenAPI Excel Data Dictionary generated: {out_file}")
        return 0
    except Exception as exc:
        print(f"[ERROR] ba-openapi failed: {exc}", file=sys.stderr)
        return 2


def cmd_ba_decision_table(args: argparse.Namespace) -> int:
    from core.ba_hub.decision_table import DecisionTable

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Spec not found: {src}", file=sys.stderr)
        return 1

    try:
        with open(src, "r", encoding="utf-8") as f:
            raw_data = json.load(f)
        dt = DecisionTable.from_dict(raw_data)

        if getattr(args, "audit", False):
            audit_res = dt.audit_completeness_and_contradictions()
            print("=" * 60)
            print(f"  BABOK v3 DECISION TABLE AUDIT: {dt.table_id} ({dt.title})")
            print("=" * 60)
            print(f"  Score: {audit_res.score}/100 | Status: {'VALID' if audit_res.is_valid else 'NEEDS_REVISION'}")
            print(f"  Theoretical Combinations: {audit_res.total_theoretical_rules} | Rules Defined: {audit_res.actual_rules_count}")
            if audit_res.contradictions:
                print("  [!] Mâu Thuẫn Logic:")
                for c in audit_res.contradictions:
                    print(f"    - {c}")
            if audit_res.missing_combinations:
                print(f"  [?] Thiếu {len(audit_res.missing_combinations)} kịch bản kết hợp:")
                for m in audit_res.missing_combinations[:5]:
                    print(f"    - {m}")
            print("=" * 60)

        out_path = args.output
        if out_path:
            if out_path.endswith(".xlsx"):
                dt.export_to_excel(out_path)
            else:
                dt.export_to_docx(out_path)
            print(f"[OK] Decision Table exported to: {out_path}")
        return 0
    except Exception as exc:
        print(f"[ERROR] ba-decision-table failed: {exc}", file=sys.stderr)
        return 2


def cmd_ba_meeting_parser(args: argparse.Namespace) -> int:
    from core.ba_hub.meeting_parser import MeetingNotesParser

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Meeting notes file not found: {src}", file=sys.stderr)
        return 1

    try:
        with open(src, "r", encoding="utf-8") as f:
            raw_text = f.read()

        parsed = MeetingNotesParser.parse_text(raw_text)
        out_path = args.output
        out_fmt = getattr(args, "format", "docx").lower()

        if out_fmt == "json" or (out_path and out_path.endswith(".json")):
            target = out_path or os.path.splitext(src)[0] + "_analysis.json"
            with open(target, "w", encoding="utf-8") as f:
                json.dump(parsed.to_dict(), f, ensure_ascii=False, indent=2)
            print(f"[OK] Meeting analysis JSON written to: {target}")
        elif out_fmt == "md" or (out_path and out_path.endswith(".md")):
            target = out_path or os.path.splitext(src)[0] + "_analysis.md"
            MeetingNotesParser.export_to_markdown(parsed, target)
            print(f"[OK] Meeting Markdown (Obsidian) written to: {target}")
        else:
            target = out_path or os.path.splitext(src)[0] + "_analysis.docx"
            MeetingNotesParser.export_to_docx(parsed, target)
            print(f"[OK] Meeting Word Specification generated: {target}")
        return 0
    except Exception as exc:
        print(f"[ERROR] ba-meeting-parser failed: {exc}", file=sys.stderr)
        return 2


def cmd_wireframe_render(args: argparse.Namespace) -> int:
    from core.diagrams.salt_wireframe import SaltWireframeEngine

    src = args.spec
    output_path = args.output or "output/wireframe.png"
    scale = float(getattr(args, "scale", 2.0))

    try:
        if os.path.isfile(src):
            with open(src, "r", encoding="utf-8") as f:
                content = f.read().strip()
            try:
                spec_data = json.loads(content)
            except Exception:
                spec_data = content
        else:
            spec_data = src

        res = SaltWireframeEngine.render_salt_to_png(spec_data, output_path, scale=scale)
        print(f"[OK] Wireframe rendered successfully ({res['dimensions_px'][0]}x{res['dimensions_px'][1]} px) -> {res['png_path']}")
        return 0
    except Exception as exc:
        print(f"[ERROR] wireframe-render failed: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# Handlers: Enterprise Printing Engine
# ---------------------------------------------------------------------------

def cmd_print_list(args: argparse.Namespace) -> int:
    from core.printing import PrinterManager

    try:
        printers = PrinterManager.list_printers()
        if getattr(args, "json", False):
            print(json.dumps(printers, ensure_ascii=False, indent=2))
        else:
            print("=" * 70)
            print("  DANH SÁCH MÁY IN HỆ THỐNG (SYSTEM HARDWARE & VIRTUAL PRINTERS)")
            print("=" * 70)
            for idx, p in enumerate(printers, start=1):
                star = " [DEFAULT]" if p.get("is_default") else ""
                v_tag = " (Virtual)" if p.get("is_virtual") else " (Hardware)"
                print(f"  {idx}. {p['name']}{star}{v_tag}")
                if p.get("description"):
                    print(f"     Mô tả: {p['description']}")
            print("=" * 70)
        return 0
    except Exception as exc:
        print(f"[ERROR] print-list failed: {exc}", file=sys.stderr)
        return 2


def cmd_print_doc(args: argparse.Namespace) -> int:
    from core.printing import PrinterManager

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] File to print not found: {src}", file=sys.stderr)
        return 1

    try:
        opt_options = {
            "grayscale": getattr(args, "grayscale", False),
            "n_up": int(getattr(args, "n_up", 1)),
            "booklet": getattr(args, "booklet", False),
            "watermark": getattr(args, "watermark", None),
            "margin_guard_mm": float(getattr(args, "margin_guard", 5.0)),
        }
        res = PrinterManager.print_file(
            file_path=src,
            printer_name=getattr(args, "printer", None),
            copies=int(getattr(args, "copies", 1)),
            optimizer_options=opt_options
        )
        print(f"[OK] Print Job Dispatched: {res['job_id']} -> Printer: {res['printer']} ({res['status']})")
        return 0
    except Exception as exc:
        print(f"[ERROR] print-doc failed: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# Handlers: Core BA Hub (Stories, Use Cases, RTM, DBML, Satsuma, SQL, Vault)
# ---------------------------------------------------------------------------

def cmd_ba_story_lint(args: argparse.Namespace) -> int:
    from core.ba_hub.story_builder import StoryQualityLinter, UserStory

    story = UserStory(
        id=getattr(args, "story_id", "US-001"),
        title=getattr(args, "title", "Yêu cầu"),
        domain=getattr(args, "domain", "General"),
        role=getattr(args, "persona", "Người dùng"),
        goal=getattr(args, "action", "thao tác"),
        benefit=getattr(args, "value", "đáp ứng nhu cầu"),
    )
    res = StoryQualityLinter.lint_story(story)
    print(json.dumps(res, ensure_ascii=False, indent=2))
    return 0


def cmd_ba_story_export(args: argparse.Namespace) -> int:
    from core.ba_hub.story_builder import StoryBuilder

    src = os.path.abspath(args.file)
    with open(src, "r", encoding="utf-8") as f:
        data = json.load(f)
    stories = [StoryBuilder.parse_story_dict(d) for d in (data if isinstance(data, list) else [data])]
    out_path = args.output or "output/user_stories.docx"
    if out_path.endswith(".md"):
        StoryBuilder.export_to_markdown(stories, out_path)
    else:
        StoryBuilder.export_to_docx(stories, out_path)
    print(f"[OK] Exported {len(stories)} User Story(ies) -> {out_path}")
    return 0


def cmd_ba_usecase_export(args: argparse.Namespace) -> int:
    from core.ba_hub.usecase_builder import UseCaseBuilder

    src = os.path.abspath(args.file)
    with open(src, "r", encoding="utf-8") as f:
        data = json.load(f)
    specs = [UseCaseBuilder.parse_usecase_dict(d) for d in (data if isinstance(data, list) else [data])]
    out_path = args.output or "output/usecases.docx"
    if out_path.endswith(".md"):
        UseCaseBuilder.export_to_markdown(specs, out_path)
    else:
        UseCaseBuilder.export_to_docx(specs, out_path)
    print(f"[OK] Exported {len(specs)} Use Case(s) -> {out_path}")
    return 0


def cmd_ba_rtm_matrix(args: argparse.Namespace) -> int:
    from core.ba_hub.rtm_matrix import RTMEngine

    src = os.path.abspath(args.file)
    with open(src, "r", encoding="utf-8") as f:
        items = json.load(f)
    out_path = args.output or "output/RTM_Matrix.xlsx"
    RTMEngine.export_rtm_to_excel("Dự Án Nghiệp Vụ", items, out_path)
    print(f"[OK] RTM Matrix exported -> {out_path}")
    return 0


def cmd_ba_dbml_convert(args: argparse.Namespace) -> int:
    from core.ba_hub.dbml_satsuma import DBMLEngine

    src = os.path.abspath(args.file)
    out_path = args.output or "output/schema.dbml"
    with open(src, "r", encoding="utf-8") as f:
        content = f.read()

    if src.endswith(".dbml"):
        sql = DBMLEngine.dbml_to_sql(content)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(sql)
        print(f"[OK] DBML converted to SQL DDL -> {out_path}")
    else:
        dbml = DBMLEngine.sql_to_dbml(content)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(dbml)
        print(f"[OK] SQL DDL converted to DBML -> {out_path}")
    return 0


def cmd_ba_satsuma_lineage(args: argparse.Namespace) -> int:
    from core.ba_hub.dbml_satsuma import SatsumaMappingEngine

    src = os.path.abspath(args.file)
    with open(src, "r", encoding="utf-8") as f:
        code = f.read()
    rules = SatsumaMappingEngine.parse_satsuma_dsl(code)
    out_path = args.output or "output/data_lineage.png"
    puml = SatsumaMappingEngine.generate_lineage_plantuml(rules)
    from core.diagrams.bridge import render_diagram_code_to_png
    res = render_diagram_code_to_png(puml, "plantuml", output_png=out_path)
    print(f"[OK] Satsuma Lineage generated ({len(rules)} rules) -> {out_path}")
    return 0


def cmd_ba_sql_suite(args: argparse.Namespace) -> int:
    from core.ba_hub.sql_toolkit import SQLIntegrityTestSuite

    src = os.path.abspath(args.file)
    with open(src, "r", encoding="utf-8") as f:
        schema = f.read()
    suite = SQLIntegrityTestSuite.generate_from_schema(schema)
    out_path = args.output or "output/sql_tests.sql"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n\n".join([f"-- {t.name}: {t.description}\n{t.sql}" for t in suite]))
    print(f"[OK] Generated {len(suite)} SQL Integrity Tests -> {out_path}")
    return 0


def cmd_ba_scope_creep(args: argparse.Namespace) -> int:
    from core.ba_hub.scope_diff import ScopeCreepAuditor

    old_doc = os.path.abspath(args.v1)
    new_doc = os.path.abspath(args.v2)
    summary = ScopeCreepAuditor.compare_documents(old_doc, new_doc)
    out_path = args.output or "output/scope_creep_report.docx"
    ScopeCreepAuditor.export_report_docx(summary, out_path)
    print(f"[OK] Scope Creep Audit: {summary.total_changes} changes, {summary.total_estimated_mandays} mandays -> {out_path}")
    return 0


def cmd_ba_obsidian_export(args: argparse.Namespace) -> int:
    from core.ba_hub.obsidian_vault import ObsidianVaultExporter

    proj = getattr(args, "project", "Enterprise_BA_Project")
    out_path = args.output or f"output/{proj}_Obsidian_Vault.zip"
    exporter = ObsidianVaultExporter(project_name=proj)
    exporter.export_to_zip(out_path)
    print(f"[OK] Obsidian Vault (.zip) exported -> {out_path}")
    return 0


def cmd_ba_obsidian_compile(args: argparse.Namespace) -> int:
    from core.ba_hub.obsidian_vault import ObsidianVaultCompiler

    src = os.path.abspath(args.file)
    out_path = args.output or "output/BRD_From_Obsidian.docx"
    compiler = ObsidianVaultCompiler(src)
    compiler.compile_to_docx(out_path)
    print(f"[OK] Obsidian Vault compiled to Word BRD -> {out_path}")
    return 0


# ---------------------------------------------------------------------------
# Registration Facade
# ---------------------------------------------------------------------------

def register_ba_print_subcommands(sub: argparse._SubParsersAction) -> None:
    """Registers all BA Hub & Enterprise Printing subcommands into the CLI parser."""

    # 1. ba-openapi
    p_oapi = sub.add_parser("ba-openapi", help="Scan OpenAPI / Swagger contracts into structured Data Dictionaries")
    p_oapi.add_argument("file", help="Path to openapi.json, swagger.json, or yaml contract")
    p_oapi.add_argument("-o", "--output", default=None, help="Output file path (.xlsx, .docx, .json)")
    p_oapi.add_argument("-f", "--format", default="excel", choices=["excel", "docx", "json"], help="Output format")
    p_oapi.set_defaults(func=cmd_ba_openapi)

    # 2. ba-decision-table
    p_dt = sub.add_parser("ba-decision-table", help="BABOK v3 Decision Table Generator & Completeness/Contradiction Audit")
    p_dt.add_argument("file", help="Path to decision table spec .json")
    p_dt.add_argument("-o", "--output", default=None, help="Output file path (.docx or .xlsx)")
    p_dt.add_argument("--audit", action="store_true", help="Run completeness & contradiction audit")
    p_dt.set_defaults(func=cmd_ba_decision_table)

    # 3. ba-meeting-parser
    p_mp = sub.add_parser("ba-meeting-parser", help="Parse meeting notes into Stories, Rules, and Action Items")
    p_mp.add_argument("file", help="Path to raw meeting notes text/markdown file")
    p_mp.add_argument("-o", "--output", default=None, help="Output file path (.docx, .md, .json)")
    p_mp.add_argument("-f", "--format", default="docx", choices=["docx", "md", "json"], help="Output format")
    p_mp.set_defaults(func=cmd_ba_meeting_parser)

    # 4. wireframe-render
    p_wf = sub.add_parser("wireframe-render", help="Render PlantUML Salt UI Wireframe mockup to 300+ DPI PNG")
    p_wf.add_argument("spec", help="Path to wireframe JSON spec file or raw Salt code")
    p_wf.add_argument("-o", "--output", default=None, help="Output PNG path")
    p_wf.add_argument("--scale", type=float, default=2.0, help="Scale factor (default: 2.0)")
    p_wf.set_defaults(func=cmd_wireframe_render)

    # 5. print-list
    p_plist = sub.add_parser("print-list", help="Discover and list available system physical & virtual printers")
    p_plist.add_argument("--json", action="store_true", help="Output machine-readable JSON")
    p_plist.set_defaults(func=cmd_print_list)

    # 6. print-doc
    p_pdoc = sub.add_parser("print-doc", help="Silent document printing with ink-saver & booklet optimizer")
    p_pdoc.add_argument("file", help="Path to document (.pdf, .docx, .xlsx)")
    p_pdoc.add_argument("--printer", default=None, help="Target printer name (default: system default printer)")
    p_pdoc.add_argument("--copies", type=int, default=1, help="Number of copies (default: 1)")
    p_pdoc.add_argument("--grayscale", action="store_true", help="Convert pages to grayscale (toner saver)")
    p_pdoc.add_argument("--n-up", type=int, default=1, choices=[1, 2, 4], help="N-Up pages per sheet (1, 2, 4)")
    p_pdoc.add_argument("--booklet", action="store_true", help="Arrange pages in 2-sided booklet order")
    p_pdoc.add_argument("--watermark", default=None, help="Security watermark text across pages")
    p_pdoc.add_argument("--margin-guard", type=float, default=5.0, help="Safety margin in mm (ERR_PRINT_001)")
    p_pdoc.set_defaults(func=cmd_print_doc)

    # 7. ba-story-lint & export
    p_sl = sub.add_parser("ba-story-lint", help="Audit User Story quality against BABOK v3 & INVEST criteria")
    p_sl.add_argument("--story-id", default="US-001", help="Story ID")
    p_sl.add_argument("--title", default="Chức năng", help="Story title")
    p_sl.add_argument("--persona", default="Người dùng", help="Role / Persona")
    p_sl.add_argument("--action", default="thao tác", help="Action / Goal")
    p_sl.add_argument("--value", default="lợi ích", help="Business Value")
    p_sl.set_defaults(func=cmd_ba_story_lint)

    p_se = sub.add_parser("ba-story-export", help="Export User Stories to OpenXML Word or Markdown")
    p_se.add_argument("file", help="Path to story JSON")
    p_se.add_argument("-o", "--output", default=None, help="Output path (.docx or .md)")
    p_se.set_defaults(func=cmd_ba_story_export)

    # 8. ba-usecase-export
    p_ue = sub.add_parser("ba-usecase-export", help="Export Cockburn Use Cases to OpenXML Word")
    p_ue.add_argument("file", help="Path to use case JSON")
    p_ue.add_argument("-o", "--output", default=None, help="Output path (.docx or .md)")
    p_ue.set_defaults(func=cmd_ba_usecase_export)

    # 9. ba-rtm-matrix
    p_rtm = sub.add_parser("ba-rtm-matrix", help="Generate Requirements Traceability Matrix (RTM) Excel workbook")
    p_rtm.add_argument("file", help="Path to RTM spec JSON")
    p_rtm.add_argument("-o", "--output", default=None, help="Output .xlsx path")
    p_rtm.set_defaults(func=cmd_ba_rtm_matrix)

    # 10. ba-dbml-convert
    p_dbml = sub.add_parser("ba-dbml-convert", help="Convert between DBML and SQL DDL schemas")
    p_dbml.add_argument("file", help="Path to .dbml or .sql file")
    p_dbml.add_argument("-o", "--output", default=None, help="Output file path")
    p_dbml.set_defaults(func=cmd_ba_dbml_convert)

    # 11. ba-satsuma-lineage
    p_sat = sub.add_parser("ba-satsuma-lineage", help="Generate Data Lineage Graph from Satsuma DSL")
    p_sat.add_argument("file", help="Path to .satsuma file")
    p_sat.add_argument("-o", "--output", default=None, help="Output PNG path")
    p_sat.set_defaults(func=cmd_ba_satsuma_lineage)

    # 12. ba-sql-suite
    p_sqs = sub.add_parser("ba-sql-suite", help="Generate 7-Gate SQL Data Integrity Test Suite")
    p_sqs.add_argument("file", help="Path to SQL DDL schema file")
    p_sqs.add_argument("-o", "--output", default=None, help="Output SQL script path")
    p_sqs.set_defaults(func=cmd_ba_sql_suite)

    # 13. ba-scope-creep
    p_scp = sub.add_parser("ba-scope-creep", help="Audit Scope Creep & Semantic Deltas between two Word documents")
    p_scp.add_argument("v1", help="Path to baseline Word document v1")
    p_scp.add_argument("v2", help="Path to revised Word document v2")
    p_scp.add_argument("-o", "--output", default=None, help="Output report .docx path")
    p_scp.set_defaults(func=cmd_ba_scope_creep)

    # 14. ba-obsidian-export & compile
    p_ox = sub.add_parser("ba-obsidian-export", help="Export project requirements to Obsidian Vault (.zip)")
    p_ox.add_argument("-p", "--project", default="Enterprise_BA_Project", help="Project name")
    p_ox.add_argument("-o", "--output", default=None, help="Output .zip path")
    p_ox.set_defaults(func=cmd_ba_obsidian_export)

    p_oc = sub.add_parser("ba-obsidian-compile", help="Compile Obsidian Vault (.zip) into Word BRD (.docx)")
    p_oc.add_argument("file", help="Path to Obsidian Vault .zip file")
    p_oc.add_argument("-o", "--output", default=None, help="Output .docx path")
    p_oc.set_defaults(func=cmd_ba_obsidian_compile)

