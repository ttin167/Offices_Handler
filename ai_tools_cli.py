"""
ai_tools_cli.py — CLI toolkit for AI-driven document editing.

Provides structured read/write/diff operations on DOCX and XLSX files,
outputting JSON snapshots that AI (Antigravity IDE) can read and modify.

Subcommands:
  docx-read       — Read DOCX → JSON snapshot
  docx-write      — Write JSON snapshot → DOCX (with optional style template)
  docx-diff       — Compare two JSON snapshots → before/after diff report
  docx-inject     — Surgically inject paragraphs, runs, and images into DOCX
  mermaid-render  — Render Mermaid code to high-res PNG and inject into DOCX
  plantuml-render — Render PlantUML code to high-res PNG and inject into DOCX
  diagram-render  — Unified multi-engine dispatcher (canvas | mermaid | plantuml)
  spec-render     — Render SVG Canvas Precision Diagram to high-res PNG
  diagram-editor  — Interactive Canvas Editor for Precision Diagram JSON specs
  docx-to-pptx    — Synthesize 16:9 PowerPoint presentation deck from Word .docx
  pptx-build      — Compile 16:9 PowerPoint deck from structured JSON spec (or pptx-create)
  pptx-preview    — Export presentation slides to high-res PNG previews for QA
  pptx-from-md    — Compile Markdown slide deck ('---' delimiter) to .pptx
  xlsx-read       — Read XLSX → JSON snapshot (auto-hybrid, headless openpyxl fallback)
  xlsx-write      — Write JSON snapshot → XLSX
  xlsx-mutate     — Mutate Excel template in-place (row/col expansion, style cloning, chart relink)
  xlsx-validate   — Validate Excel sheet against reference template across 12 Quality Gates (Invariant E6)

  python ai_tools_cli.py pptx-build spec.json -o deck.pptx --theme corporate_blue
  python ai_tools_cli.py pptx-preview deck.pptx -o ./preview_slides --slides 1,5,11,15
  python ai_tools_cli.py docx-to-pptx document.docx -o presentation.pptx --theme thesis_blue
  python ai_tools_cli.py diagram-render --spec spec.json -o out.png
  python ai_tools_cli.py plantuml-render erd_spec.json -o erd.png
  python ai_tools_cli.py docx-read file.docx -o snap.json
  python ai_tools_cli.py docx-write snap.json --template file.docx -o out.docx
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            getattr(sys.stdout, "reconfigure")(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            getattr(sys.stderr, "reconfigure")(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# docx-read
# ---------------------------------------------------------------------------

def cmd_docx_read(args: argparse.Namespace) -> int:
    from .docx_reader import read_docx_to_json

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] File not found: {src}", file=sys.stderr)
        return 1

    t0 = time.time()
    try:
        result = read_docx_to_json(src, output_path=args.output)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    if args.output:
        print(f"[OK] Snapshot written to: {result}  ({elapsed:.2f}s)")
    else:
        print(result)
    return 0


# ---------------------------------------------------------------------------
# docx-write
# ---------------------------------------------------------------------------

def cmd_docx_write(args: argparse.Namespace) -> int:
    from .docx_writer import write_docx_from_json_file

    json_path = os.path.abspath(args.snapshot)
    if not os.path.isfile(json_path):
        print(f"[ERROR] Snapshot not found: {json_path}", file=sys.stderr)
        return 1

    template_path = os.path.abspath(args.template) if args.template else None
    if template_path and not os.path.isfile(template_path):
        print(f"[ERROR] Template DOCX not found: {template_path}", file=sys.stderr)
        return 1

    output_path = args.output or os.path.splitext(json_path)[0] + "_written.docx"
    output_path = os.path.abspath(output_path)

    t0 = time.time()
    try:
        result = write_docx_from_json_file(json_path, output_path, template_path)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] DOCX written to: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# docx-inject
# ---------------------------------------------------------------------------

def cmd_docx_inject(args: argparse.Namespace) -> int:
    from .docx_writer import inject_content_into_docx

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Injection spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    out_path = os.path.abspath(args.output) if args.output else src

    t0 = time.time()
    try:
        result = inject_content_into_docx(src, spec, out_path)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] Content injected into: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# docx-patch
# ---------------------------------------------------------------------------

def cmd_docx_patch(args: argparse.Namespace) -> int:
    try:
        from .docx_writer import patch_docx_template
    except (ImportError, ValueError):
        from docx_writer import patch_docx_template  # type: ignore

    template_path = os.path.abspath(args.template)
    if not os.path.isfile(template_path):
        print(f"[ERROR] Template DOCX not found: {template_path}", file=sys.stderr)
        return 1

    patch_path = os.path.abspath(args.patch)
    if not os.path.isfile(patch_path):
        print(f"[ERROR] Patch JSON spec not found: {patch_path}", file=sys.stderr)
        return 1

    with open(patch_path, "r", encoding="utf-8") as f:
        patch_spec = json.load(f)

    text_rep = patch_spec.get("text_replacements") or patch_spec.get("texts") or {}
    block_rep = patch_spec.get("block_replacements") or patch_spec.get("blocks") or {}

    out_path = os.path.abspath(args.output) if args.output else os.path.splitext(template_path)[0] + "_patched.docx"

    t0 = time.time()
    try:
        result = patch_docx_template(template_path, out_path, text_rep, block_rep)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] Patched DOCX written to: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# docx-page-numbers
# ---------------------------------------------------------------------------

def cmd_docx_page_numbers(args: argparse.Namespace) -> int:
    import docx
    try:
        from .docx_advanced_engine import inject_dynamic_page_numbers, safe_save_docx
    except (ImportError, ValueError):
        from docx_advanced_engine import inject_dynamic_page_numbers, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        inject_dynamic_page_numbers(doc, font_name=args.font, font_size_pt=args.size)
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Dynamic page numbers injected into: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-caption-fix
# ---------------------------------------------------------------------------

def cmd_docx_caption_fix(args: argparse.Namespace) -> int:
    import docx
    try:
        from .docx_advanced_engine import format_figure_captions, safe_save_docx
    except (ImportError, ValueError):
        from docx_advanced_engine import format_figure_captions, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        count = format_figure_captions(doc, font_name=args.font, font_size_pt=args.size)
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Standardized {count} captions in: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-sanitize-icons
# ---------------------------------------------------------------------------

def cmd_docx_sanitize_icons(args: argparse.Namespace) -> int:
    import docx
    try:
        from .docx_advanced_engine import sanitize_emojis_and_symbols, safe_save_docx
    except (ImportError, ValueError):
        from docx_advanced_engine import sanitize_emojis_and_symbols, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        count = sanitize_emojis_and_symbols(doc)
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Sanitized {count} elements (emojis/icons) in: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-translate
# ---------------------------------------------------------------------------

def cmd_docx_translate(args: argparse.Namespace) -> int:
    import docx
    try:
        from .docx_advanced_engine import translate_vn_to_en_protected, safe_save_docx
    except (ImportError, ValueError):
        from docx_advanced_engine import translate_vn_to_en_protected, safe_save_docx  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Source DOCX not found: {src}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else src
    try:
        doc = docx.Document(src)
        translated_count = 0
        for p in doc.paragraphs:
            old_t = p.text
            new_t = translate_vn_to_en_protected(old_t)
            if new_t != old_t:
                p.text = new_t
                translated_count += 1
        for tbl in doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        old_t = p.text
                        new_t = translate_vn_to_en_protected(old_t)
                        if new_t != old_t:
                            p.text = new_t
                            translated_count += 1
        saved = safe_save_docx(doc, out_path)
        print(f"[OK] Translated {translated_count} blocks in: {saved}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# mermaid-render
# ---------------------------------------------------------------------------

def cmd_mermaid_render(args: argparse.Namespace) -> int:
    try:
        from .mermaid_renderer import render_mermaid_to_png
    except (ImportError, ValueError):
        from mermaid_renderer import render_mermaid_to_png  # type: ignore

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Mermaid spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    # CLI flag overrides
    if args.output:
        spec["output_path"] = args.output
    if args.inject:
        spec["inject_into"] = args.inject
    if args.theme:
        spec["theme_preset"] = args.theme
    if args.scale:
        spec["scale"] = args.scale

    base_dir = os.path.dirname(spec_path)
    t0 = time.time()
    try:
        result = render_mermaid_to_png(spec, base_dir=base_dir)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    png_path = result["png_path"]
    w_cm = result["width_cm"]
    h_cm = result["height_cm"]
    w_px, h_px = result["dimensions_px"]
    f_size = result["file_size_bytes"]

    print(f"[OK] Diagram rendered: {png_path}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Doc Layout: {w_cm}cm x {h_cm}cm | Size: {f_size} bytes")

    if result.get("injected"):
        print(f"[OK] Injected into DOCX: {result.get('docx_path')}")

    return 0


# ---------------------------------------------------------------------------
# plantuml-render
# ---------------------------------------------------------------------------

def cmd_plantuml_render(args: argparse.Namespace) -> int:
    try:
        from .plantuml_renderer import render_plantuml_to_png
    except (ImportError, ValueError):
        from plantuml_renderer import render_plantuml_to_png  # type: ignore

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] PlantUML spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    if args.output:
        spec["output_path"] = args.output
    if args.inject:
        spec["inject_into"] = args.inject
    if args.dpi:
        spec["dpi"] = args.dpi

    base_dir = os.path.dirname(spec_path)
    t0 = time.time()
    try:
        result = render_plantuml_to_png(spec, base_dir=base_dir)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    png_path = result["png_path"]
    w_cm = result["width_cm"]
    h_cm = result["height_cm"]
    w_px, h_px = result["dimensions_px"]
    f_size = result["file_size_bytes"]

    print(f"[OK] PlantUML diagram rendered: {png_path}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Doc Layout: {w_cm}cm x {h_cm}cm | Size: {f_size} bytes")

    if result.get("injected"):
        print(f"[OK] Injected into DOCX: {result.get('docx_path')}")

    return 0


# ---------------------------------------------------------------------------
# diagram-render (Unified Multi-Engine Technical Diagram Dispatcher)
# ---------------------------------------------------------------------------

def cmd_diagram_render(args: argparse.Namespace) -> int:
    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Spec file not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    if args.output:
        spec["output_path"] = args.output
    if args.inject:
        spec["inject_into"] = args.inject

    engine = str(spec.get("engine", "mxgraph")).strip().lower()
    base_dir = os.path.dirname(spec_path)
    out_png = args.output or spec.get("output_path") or os.path.splitext(spec_path)[0] + ".png"
    out_png = os.path.abspath(out_png)
    spec["output_path"] = out_png

    t0 = time.time()
    try:
        if engine in ("mxgraph", "drawio"):
            try:
                from .mxgraph_engine import render_diagram
            except (ImportError, ValueError):
                from mxgraph_engine import render_diagram  # type: ignore
            if args.scale is not None:
                spec["scale"] = args.scale
            result = render_diagram(spec, base_dir=base_dir)
        elif engine == "canvas":
            try:
                from .spec_diagram_engine import PrecisionDiagram
            except (ImportError, ValueError):
                from spec_diagram_engine import PrecisionDiagram  # type: ignore
            scale = args.scale if args.scale is not None else spec.get("scale", 3)
            diag = PrecisionDiagram.from_spec(spec, base_dir=base_dir)
            result = diag.render_to_png(out_png, scale=scale)
        elif engine == "mermaid":
            try:
                from .mermaid_renderer import render_mermaid_to_png
            except (ImportError, ValueError):
                from mermaid_renderer import render_mermaid_to_png  # type: ignore
            result = render_mermaid_to_png(spec, base_dir=base_dir)
        elif engine == "plantuml":
            try:
                from .plantuml_renderer import render_plantuml_to_png
            except (ImportError, ValueError):
                from plantuml_renderer import render_plantuml_to_png  # type: ignore
            result = render_plantuml_to_png(spec, base_dir=base_dir)
        else:
            raise ValueError(f"Unknown diagram engine: '{engine}'. Supported: 'mxgraph', 'drawio', 'canvas', 'mermaid', 'plantuml'")
    except Exception as exc:
        print(f"[ERROR] Diagram rendering failed ({engine}): {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    w_px, h_px = result["dimensions_px"]
    f_size = result["file_size_bytes"]
    print(f"[OK] {engine.upper()} diagram rendered: {result['png_path']}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Size: {f_size} bytes")

    # Windows Trap 4: Sau khi render ra result["png_path"] thành công:
    # Auto-inject vào Word nếu spec có inject_into và chưa được inject
    docx_target = spec.get("inject_into")
    if docx_target and not result.get("injected"):
        if not os.path.isabs(docx_target) and base_dir:
            docx_target = os.path.abspath(os.path.join(base_dir, docx_target))
        if os.path.exists(docx_target):
            try:
                from .docx_writer import inject_diagram_into_docx
            except (ImportError, ValueError):
                from docx_writer import inject_diagram_into_docx  # type: ignore
            inject_diagram_into_docx(
                docx_path=docx_target,
                png_path=result["png_path"],
                heading=spec.get("target_heading"),
                placeholder=spec.get("placeholder"),
                caption=spec.get("caption_template") or spec.get("caption"),
                width_cm=spec.get("width_cm", 14.0),
                max_height_cm=spec.get("max_height_cm", 20.0),
            )
            print(f"[OK] Injected diagram into {docx_target}")
        else:
            print(f"[WARN] Target DOCX for injection not found: {docx_target}")
    elif result.get("injected"):
        print(f"[OK] Injected diagram into {docx_target}")

    return 0


# ---------------------------------------------------------------------------
# spec-render (Precision Technical & Architecture Diagram Engine)
# ---------------------------------------------------------------------------

def cmd_spec_render(args: argparse.Namespace) -> int:
    from .spec_diagram_engine import PrecisionDiagram

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Spec file not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec_dict = json.load(f)

    out_png = args.output or spec_dict.get("output_path") or os.path.splitext(spec_path)[0] + ".png"
    out_png = os.path.abspath(out_png)
    scale = args.scale if args.scale is not None else spec_dict.get("scale", 3)

    t0 = time.time()
    try:
        diag = PrecisionDiagram.from_spec(spec_dict)
        res = diag.render_to_png(out_png, scale=scale)
    except Exception as exc:
        print(f"[ERROR] Precision diagram rendering failed: {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    w_px, h_px = res["dimensions_px"]
    f_size = res["file_size_bytes"]
    ar = res["aspect_ratio"]

    print(f"[OK] Precision Diagram rendered: {res['png_path']}  ({elapsed:.2f}s)")
    print(f"     Resolution: {w_px}x{h_px}px | Aspect Ratio: {ar}:1 | Size: {f_size} bytes | Scale: {scale}x")
    return 0


def cmd_diagram_editor(args: argparse.Namespace) -> int:
    try:
        from .diagram_editor import DiagramEditorApp
    except (ImportError, ValueError):
        from diagram_editor import DiagramEditorApp  # type: ignore
    app = DiagramEditorApp(spec_path=args.spec)
    app.run()
    return 0

def _flatten_body(body: list[dict]) -> dict[int, dict]:
    """Return {element_index: element} mapping."""
    return {el.get("element_index", i): el for i, el in enumerate(body)}


def cmd_docx_diff(args: argparse.Namespace) -> int:
    before_path = os.path.abspath(args.before)
    after_path  = os.path.abspath(args.after)

    for p in (before_path, after_path):
        if not os.path.isfile(p):
            print(f"[ERROR] File not found: {p}", file=sys.stderr)
            return 1

    with open(before_path, "r", encoding="utf-8") as f:
        before_snap = json.load(f)
    with open(after_path, "r", encoding="utf-8") as f:
        after_snap = json.load(f)

    before_body = _flatten_body(before_snap.get("body", []))
    after_body  = _flatten_body(after_snap.get("body", []))

    all_indices = sorted(set(before_body.keys()) | set(after_body.keys()))

    changes = []
    for idx in all_indices:
        b = before_body.get(idx)
        a = after_body.get(idx)

        if b == a:
            continue  # No change

        entry: dict = {"element_index": idx}
        if b is None:
            entry["change_type"] = "added"
            entry["after"] = a
        elif a is None:
            entry["change_type"] = "removed"
            entry["before"] = b
        else:
            entry["change_type"] = "modified"
            entry["before"] = b
            entry["after"] = a

        changes.append(entry)

    diff_report = {
        "before_file": before_snap.get("source_file", before_path),
        "after_file":  after_snap.get("source_file", after_path),
        "total_before_elements": len(before_body),
        "total_after_elements":  len(after_body),
        "total_changes": len(changes),
        "changes": changes,
    }

    json_str = json.dumps(diff_report, ensure_ascii=False, indent=2)

    if args.output:
        out_path = os.path.abspath(args.output)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(json_str)
        print(f"[OK] Diff report written to: {out_path}  ({len(changes)} changes)")
    else:
        print(json_str)

    return 0


# ---------------------------------------------------------------------------
# xlsx-read
# ---------------------------------------------------------------------------

def cmd_xlsx_read(args: argparse.Namespace) -> int:
    from .xlsx_reader import read_xlsx_to_json

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] File not found: {src}", file=sys.stderr)
        return 1

    t0 = time.time()
    try:
        result = read_xlsx_to_json(
            src,
            output_path=args.output,
            sheet_name=args.sheet,
            cell_range=args.range,
            engine=getattr(args, "engine", "auto"),
        )
    except ImportError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 3
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    if args.output:
        print(f"[OK] Snapshot written to: {result}  ({elapsed:.2f}s)")
    else:
        print(result)
    return 0


# ---------------------------------------------------------------------------
# xlsx-write
# ---------------------------------------------------------------------------

def cmd_xlsx_write(args: argparse.Namespace) -> int:
    from .xlsx_writer import write_xlsx_from_json_file

    json_path = os.path.abspath(args.snapshot)
    if not os.path.isfile(json_path):
        print(f"[ERROR] Snapshot not found: {json_path}", file=sys.stderr)
        return 1

    template_path = os.path.abspath(args.template) if args.template else None
    if template_path and not os.path.isfile(template_path):
        print(f"[ERROR] Template XLSX not found: {template_path}", file=sys.stderr)
        return 1

    output_path = args.output or os.path.splitext(json_path)[0] + "_written.xlsx"
    output_path = os.path.abspath(output_path)

    t0 = time.time()
    try:
        result = write_xlsx_from_json_file(json_path, output_path, template_path)
    except ImportError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 3
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] XLSX written to: {result}  ({elapsed:.2f}s)")
    return 0


# ---------------------------------------------------------------------------
# xlsx-mutate
# ---------------------------------------------------------------------------

def cmd_xlsx_mutate(args: argparse.Namespace) -> int:
    try:
        from .xlsx_writer import mutate_template_excel
    except (ImportError, ValueError):
        from xlsx_writer import mutate_template_excel  # type: ignore

    template_path = os.path.abspath(args.template)
    if not os.path.isfile(template_path):
        print(f"[ERROR] Template XLSX not found: {template_path}", file=sys.stderr)
        return 1

    spec_path = os.path.abspath(args.spec)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Mutation JSON spec not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    out_path = os.path.abspath(args.output) if args.output else os.path.splitext(template_path)[0] + "_mutated.xlsx"

    t0 = time.time()
    try:
        result = mutate_template_excel(template_path, out_path, spec)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2

    elapsed = time.time() - t0
    print(f"[OK] Mutated XLSX written to: {result}  ({elapsed:.2f}s)")

    if getattr(args, "validate", False):
        try:
            from .xlsx_validator import validate_excel_sheet
        except (ImportError, ValueError):
            from xlsx_validator import validate_excel_sheet  # type: ignore

        sheets_spec = spec.get("sheets", {})
        for sheet_name in sheets_spec:
            try:
                res = validate_excel_sheet(
                    result,
                    target_sheet=sheet_name,
                    ref_sheet=getattr(args, "ref_sheet", None),
                    verbose=True,
                )
                if not res.passed:
                    print(f"[WARN] Invariant E6: {len(res.diffs)} format discrepancy(ies) on sheet '{sheet_name}'.", file=sys.stderr)
            except Exception as v_exc:
                print(f"[WARN] Validation skipped for sheet '{sheet_name}': {v_exc}")

    return 0


# ---------------------------------------------------------------------------
# xlsx-validate
# ---------------------------------------------------------------------------

def cmd_xlsx_validate(args: argparse.Namespace) -> int:
    try:
        from .xlsx_validator import validate_excel_sheet
    except (ImportError, ValueError):
        from xlsx_validator import validate_excel_sheet  # type: ignore

    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Workbook not found: {src}", file=sys.stderr)
        return 2

    t0 = time.time()
    try:
        res = validate_excel_sheet(
            src,
            target_sheet=args.target,
            ref_sheet=args.ref,
            verbose=True,
        )
        elapsed = time.time() - t0
        if res.passed:
            print(f"[OK] Sheet '{args.target}' passed all 12 Quality Gates! ({elapsed:.2f}s)")
            return 0
        else:
            print(f"[ERROR] Sheet '{args.target}' failed {res.gates_failed} Quality Gate(s) ({elapsed:.2f}s)", file=sys.stderr)
            return 1
    except Exception as exc:
        print(f"[ERROR] Validation failed: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-create
# ---------------------------------------------------------------------------

def cmd_pptx_create(args: argparse.Namespace) -> int:
    try:
        from .pptx_writer import write_pptx_from_spec
    except (ImportError, ValueError):
        from pptx_writer import write_pptx_from_spec  # type: ignore

    spec_raw = getattr(args, "spec_opt", None) or getattr(args, "spec", None)
    if not spec_raw:
        print("[ERROR] No slide spec JSON provided. Use --spec <file.json> or positional argument.", file=sys.stderr)
        return 1

    spec_path = os.path.abspath(spec_raw)
    if not os.path.isfile(spec_path):
        print(f"[ERROR] Slide spec JSON not found: {spec_path}", file=sys.stderr)
        return 1

    with open(spec_path, "r", encoding="utf-8") as f:
        spec = json.load(f)

    out_path = args.output or spec.get("output_path") or os.path.splitext(spec_path)[0] + ".pptx"
    out_path = os.path.abspath(out_path)

    tpl = args.template or spec.get("template")
    tpl_path = os.path.abspath(tpl) if tpl else None

    try:
        result = write_pptx_from_spec(
            spec=spec,
            output_path=out_path,
            theme_name=args.theme or spec.get("theme", "thesis_blue"),
            template_path=tpl_path,
        )
        print(f"[OK] PowerPoint presentation created: {result}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-from-md
# ---------------------------------------------------------------------------

def cmd_pptx_from_md(args: argparse.Namespace) -> int:
    try:
        from .pptx_engine import compile_markdown_file_to_pptx
    except (ImportError, ValueError):
        from pptx_engine import compile_markdown_file_to_pptx  # type: ignore

    md_path = os.path.abspath(args.file)
    if not os.path.isfile(md_path):
        print(f"[ERROR] Markdown file not found: {md_path}", file=sys.stderr)
        return 1

    out_path = args.output or os.path.splitext(md_path)[0] + ".pptx"
    out_path = os.path.abspath(out_path)

    try:
        result = compile_markdown_file_to_pptx(
            md_file_path=md_path,
            output_pptx_path=out_path,
            theme_name=args.theme,
            template_path=os.path.abspath(args.template) if args.template else None,
        )
        print(f"[OK] Compiled Markdown slide deck to: {result}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# docx-to-pptx
# ---------------------------------------------------------------------------

def cmd_docx_to_pptx(args: argparse.Namespace) -> int:
    try:
        from .docx_to_pptx import convert_docx_to_pptx
    except (ImportError, ValueError):
        from docx_to_pptx import convert_docx_to_pptx  # type: ignore

    docx_path = os.path.abspath(args.file)
    if not os.path.isfile(docx_path):
        print(f"[ERROR] Source DOCX not found: {docx_path}", file=sys.stderr)
        return 1

    try:
        pptx_res, spec_res = convert_docx_to_pptx(
            docx_path=docx_path,
            output_pptx=args.output,
            spec_out=args.spec_out,
            theme=args.theme,
            template_path=os.path.abspath(args.template) if args.template else None,
        )
        print(f"[OK] Successfully synthesized presentation: {pptx_res}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to convert DOCX to PPTX: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pdf-to-pptx
# ---------------------------------------------------------------------------

def cmd_pdf_to_pptx(args: argparse.Namespace) -> int:
    try:
        from .docx_to_pptx import convert_pdf_to_pptx
    except (ImportError, ValueError):
        from docx_to_pptx import convert_pdf_to_pptx  # type: ignore

    pdf_path = os.path.abspath(args.file)
    if not os.path.isfile(pdf_path):
        print(f"[ERROR] Source PDF not found: {pdf_path}", file=sys.stderr)
        return 1

    try:
        pptx_res, spec_res = convert_pdf_to_pptx(
            pdf_path=pdf_path,
            output_pptx=args.output,
            spec_out=args.spec_out,
            theme=args.theme,
            template_path=os.path.abspath(args.template) if args.template else None,
            page_range=getattr(args, "pages", None),
        )
        print(f"[OK] Successfully synthesized presentation from PDF: {pptx_res}")
        return 0
    except Exception as exc:
        print(f"[ERROR] Failed to convert PDF to PPTX: {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-insert-diagram
# ---------------------------------------------------------------------------

def cmd_pptx_insert_diagram(args: argparse.Namespace) -> int:
    try:
        from .pptx_engine import insert_diagram_into_pptx
    except (ImportError, ValueError):
        from pptx_engine import insert_diagram_into_pptx  # type: ignore

    pptx_path = os.path.abspath(args.file)
    img_path = os.path.abspath(args.image)
    if not os.path.isfile(pptx_path):
        print(f"[ERROR] Target PPTX not found: {pptx_path}", file=sys.stderr)
        return 1
    if not os.path.isfile(img_path):
        print(f"[ERROR] Diagram image not found: {img_path}", file=sys.stderr)
        return 1

    out_path = os.path.abspath(args.output) if args.output else pptx_path

    try:
        result = insert_diagram_into_pptx(
            pptx_path=pptx_path,
            image_path=img_path,
            slide_index=args.slide,
            title=args.title,
            caption=args.caption,
            output_path=out_path,
            theme_name=args.theme,
        )
        print(f"[OK] Diagram successfully inserted into: {result}")
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-inspect
# ---------------------------------------------------------------------------

def cmd_pptx_inspect(args: argparse.Namespace) -> int:
    try:
        from .pptx_reader import read_pptx_to_json, read_pptx
    except (ImportError, ValueError):
        from pptx_reader import read_pptx_to_json, read_pptx  # type: ignore

    pptx_path = os.path.abspath(args.file)
    if not os.path.isfile(pptx_path):
        print(f"[ERROR] PPTX not found: {pptx_path}", file=sys.stderr)
        return 1

    try:
        if args.output:
            out_json = read_pptx_to_json(pptx_path, args.output)
            print(f"[OK] Presentation AST written to: {out_json}")
        else:
            ast = read_pptx(pptx_path)
            print(json.dumps(ast, indent=2, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# pptx-preview
# ---------------------------------------------------------------------------

def cmd_pptx_preview(args: argparse.Namespace) -> int:
    """Exports PowerPoint presentation slides to high-res PNG previews."""
    import tempfile
    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        print(f"[ERROR] Presentation file not found: {src}", file=sys.stderr)
        return 1

    out_dir = os.path.abspath(args.output_dir or os.path.join(os.path.dirname(src), "preview_slides"))
    os.makedirs(out_dir, exist_ok=True)

    slides_arg = getattr(args, "slides", "all") or "all"
    slides_arg = slides_arg.strip()
    dpi = getattr(args, "dpi", 150) or 150
    t0 = time.time()

    exported_files = []

    # 1. Native Windows PowerPoint COM automation (Pixel-perfect 1920x1080)
    if sys.platform == "win32":
        try:
            import win32com.client
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            deck = powerpoint.Presentations.Open(src, WithWindow=False)
            total = deck.Slides.Count

            if slides_arg.lower() == "all":
                target_indices = list(range(1, total + 1))
            else:
                target_indices = [int(x.strip()) for x in slides_arg.split(",") if x.strip().isdigit()]

            for s_idx in target_indices:
                if 1 <= s_idx <= total:
                    out_png = os.path.join(out_dir, f"slide_{s_idx:02d}_preview.png")
                    deck.Slides(s_idx).Export(out_png, "PNG", 1920, 1080)
                    exported_files.append(out_png)

            deck.Close()
            powerpoint.Quit()
        except Exception as com_err:
            print(f"[WARN] PowerPoint COM export unavailable or encountered an error: {com_err}", file=sys.stderr)

    # 2. Headless LibreOffice / PyMuPDF fallback
    if not exported_files:
        try:
            import subprocess, pymupdf
            temp_pdf_dir = tempfile.mkdtemp()
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", src, "--outdir", temp_pdf_dir],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60
            )
            pdf_name = os.path.splitext(os.path.basename(src))[0] + ".pdf"
            temp_pdf = os.path.join(temp_pdf_dir, pdf_name)
            if os.path.isfile(temp_pdf):
                doc = pymupdf.open(temp_pdf)
                total = len(doc)
                if slides_arg.lower() == "all":
                    target_indices = list(range(1, total + 1))
                else:
                    target_indices = [int(x.strip()) for x in slides_arg.split(",") if x.strip().isdigit()]

                for s_idx in target_indices:
                    if 1 <= s_idx <= total:
                        page = doc[s_idx - 1]
                        pix = page.get_pixmap(dpi=dpi)
                        out_png = os.path.join(out_dir, f"slide_{s_idx:02d}_preview.png")
                        pix.save(out_png)
                        exported_files.append(out_png)
                doc.close()
        except Exception as fb_err:
            print(f"[WARN] Headless fallback failed: {fb_err}", file=sys.stderr)

    elapsed = time.time() - t0
    if exported_files:
        print(f"[OK] Exported {len(exported_files)} slide previews to: {out_dir}  ({elapsed:.2f}s)")
        for f in exported_files[:5]:
            print(f"     - {os.path.basename(f)}")
        if len(exported_files) > 5:
            print(f"     ... and {len(exported_files) - 5} more.")
        return 0
    else:
        print("[ERROR] Could not export previews. Ensure Microsoft PowerPoint (Windows) or LibreOffice is installed.", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# felo-generate
# ---------------------------------------------------------------------------

def cmd_felo_generate(args: argparse.Namespace) -> int:
    skill_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".agents", "skills", "felo-slides", "scripts")
    if not os.path.isdir(skill_dir):
        skill_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "skills", "felo-slides", "scripts")

    adapter_path = os.path.join(skill_dir, "felo_adapter.py")
    if not os.path.isfile(adapter_path):
        print(f"[ERROR] Felo adapter not found at: {adapter_path}", file=sys.stderr)
        return 1

    sys.path.insert(0, skill_dir)
    try:
        import felo_adapter  # type: ignore
        res = felo_adapter.run_felo_ppt_task(
            query=args.query,
            file_path=args.file,
            theme_id=args.theme,
            task_id=args.task_id,
        )
        print("\n=== FELO PRESENTATION GENERATION RESULT ===")
        print(json.dumps(res, indent=2, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2


# ---------------------------------------------------------------------------
# Argument parser
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ai_tools_cli",
        description=(
            "AI Document Editing Toolkit — Read/write DOCX and XLSX as JSON snapshots.\n"
            "Designed for Antigravity IDE to edit non-source-code files directly."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Read DOCX to JSON snapshot:
  python ai_tools_cli.py docx-read report.docx -o report_snap.json

  # Write modified snapshot back to DOCX (with original as style anchor):
  python ai_tools_cli.py docx-write report_snap.json --template report.docx -o report_v2.docx

  # Compare two snapshots (before AI edit vs after):
  python ai_tools_cli.py docx-diff before.json after.json -o diff.json

  # Read XLSX (specific sheet + range):
  python ai_tools_cli.py xlsx-read data.xlsx --sheet "Q1 Report" --range "A1:H50" -o snap.json

  # Write XLSX from snapshot (with original as template):
  python ai_tools_cli.py xlsx-write snap.json --template data.xlsx -o data_v2.xlsx
        """,
    )

    sub = parser.add_subparsers(dest="command", required=True)

    # --- docx-read ---
    p_dr = sub.add_parser("docx-read", help="Read DOCX → JSON snapshot")
    p_dr.add_argument("file", help="Path to source .docx file")
    p_dr.add_argument("-o", "--output", default=None,
                      help="Output JSON file path (prints to stdout if omitted)")
    p_dr.set_defaults(func=cmd_docx_read)

    # --- docx-write ---
    p_dw = sub.add_parser("docx-write", help="Write JSON snapshot → DOCX")
    p_dw.add_argument("snapshot", help="Path to .json snapshot file")
    p_dw.add_argument("--template", default=None,
                      help="Path to original .docx to use as style anchor (recommended)")
    p_dw.add_argument("-o", "--output", default=None,
                      help="Output .docx path (default: <snapshot>_written.docx)")
    p_dw.set_defaults(func=cmd_docx_write)

    # --- docx-inject ---
    p_di = sub.add_parser(
        "docx-inject",
        help="Surgically inject paragraphs, runs, and images into an existing DOCX without rewriting",
    )
    p_di.add_argument("file", help="Path to source/target .docx file")
    p_di.add_argument("spec", help="Path to injection spec .json file")
    p_di.add_argument(
        "-o", "--output", default=None,
        help="Output .docx path (defaults to in-place overwrite of source file)",
    )
    p_di.set_defaults(func=cmd_docx_inject)

    # --- docx-patch ---
    p_dp = sub.add_parser(
        "docx-patch",
        help="In-place template patching for DOCX (text placeholders and block insertion without wiping layout)",
    )
    p_dp.add_argument("template", help="Path to source .docx template file")
    p_dp.add_argument("patch", help="Path to patch spec .json file")
    p_dp.add_argument("-o", "--output", default=None, help="Output .docx path (default: <template>_patched.docx)")
    p_dp.set_defaults(func=cmd_docx_patch)

    # --- mermaid-render ---
    p_mr = sub.add_parser(
        "mermaid-render",
        help="Render Mermaid diagram from spec JSON to high-res PNG and optionally inject into DOCX",
    )
    p_mr.add_argument("spec", help="Path to mermaid diagram spec .json file")
    p_mr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_mr.add_argument("--inject", default=None,
                      help="Override target .docx to inject into")
    p_mr.add_argument("--theme", default=None,
                      help="Override theme preset ('clean_modern', 'enterprise', or json path)")
    p_mr.add_argument("--scale", type=int, default=None,
                      help="Override scale factor (default: 2)")
    p_mr.set_defaults(func=cmd_mermaid_render)

    # --- plantuml-render ---
    p_pr = sub.add_parser(
        "plantuml-render",
        help="Render PlantUML diagram from spec JSON to high-res PNG and optionally inject into DOCX",
    )
    p_pr.add_argument("spec", help="Path to plantuml diagram spec .json file")
    p_pr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_pr.add_argument("--inject", default=None,
                      help="Override target .docx to inject into")
    p_pr.add_argument("--dpi", type=int, default=300,
                      help="Override DPI (default: 300)")
    p_pr.set_defaults(func=cmd_plantuml_render)

    # --- diagram-render (Unified Dispatcher) ---
    p_dr = sub.add_parser(
        "diagram-render",
        help="Unified multi-engine diagram dispatcher (reads 'engine': 'canvas' | 'mermaid' | 'plantuml')",
    )
    p_dr.add_argument("spec", help="Path to unified diagram spec .json file")
    p_dr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_dr.add_argument("--inject", default=None,
                      help="Override target .docx to inject into")
    p_dr.add_argument("--scale", type=int, default=None,
                      help="Override scale factor for canvas engine (default: 3)")
    p_dr.set_defaults(func=cmd_diagram_render)

    # --- spec-render ---
    p_sr = sub.add_parser(
        "spec-render",
        help="Render Precision Technical & Architecture diagram from spec JSON to high-res PNG (Chromium 300+ DPI)",
    )
    p_sr.add_argument("spec", help="Path to precision diagram spec .json file")
    p_sr.add_argument("-o", "--output", default=None,
                      help="Override output PNG path")
    p_sr.add_argument("--scale", type=int, default=None,
                      help="Override scale factor (default: 3)")
    p_sr.set_defaults(func=cmd_spec_render)

    # --- diagram-editor ---
    p_de = sub.add_parser(
        "diagram-editor",
        help="Launch Interactive Canvas Editor for Precision Diagram JSON specs",
    )
    p_de.add_argument("spec", nargs="?", default=None, help="Path to precision diagram spec .json file (optional)")
    p_de.set_defaults(func=cmd_diagram_editor)


    # --- docx-page-numbers ---
    p_dpn = sub.add_parser("docx-page-numbers", help="Inject dynamic OpenXML page numbering into footers")
    p_dpn.add_argument("file", help="Path to .docx file")
    p_dpn.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dpn.add_argument("--font", default="Times New Roman", help="Font family (default: Times New Roman)")
    p_dpn.add_argument("--size", type=float, default=10.0, help="Font size in pt (default: 10.0)")
    p_dpn.set_defaults(func=cmd_docx_page_numbers)

    # --- docx-caption-fix ---
    p_dcf = sub.add_parser("docx-caption-fix", help="Normalize figure/table captions with En-dash and styling")
    p_dcf.add_argument("file", help="Path to .docx file")
    p_dcf.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dcf.add_argument("--font", default="Times New Roman", help="Font family (default: Times New Roman)")
    p_dcf.add_argument("--size", type=float, default=10.5, help="Font size in pt (default: 10.5)")
    p_dcf.set_defaults(func=cmd_docx_caption_fix)

    # --- docx-sanitize-icons ---
    p_dsi = sub.add_parser("docx-sanitize-icons", help="Remove emojis and special symbols from Word docx")
    p_dsi.add_argument("file", help="Path to .docx file")
    p_dsi.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dsi.set_defaults(func=cmd_docx_sanitize_icons)

    # --- docx-translate ---
    p_dt = sub.add_parser("docx-translate", help="Translate document while strictly protecting quoted terms")
    p_dt.add_argument("file", help="Path to .docx file")
    p_dt.add_argument("-o", "--output", default=None, help="Output .docx path (default: in-place safe save)")
    p_dt.add_argument("--sl", default="vi", help="Source language (default: vi)")
    p_dt.add_argument("--tl", default="en", help="Target language (default: en)")
    p_dt.set_defaults(func=cmd_docx_translate)

    # --- docx-diff ---
    p_dd = sub.add_parser("docx-diff", help="Compare two JSON snapshots → diff report")
    p_dd.add_argument("before", help="Path to 'before' .json snapshot")
    p_dd.add_argument("after",  help="Path to 'after'  .json snapshot")
    p_dd.add_argument("-o", "--output", default=None,
                      help="Output diff report .json (prints to stdout if omitted)")
    p_dd.set_defaults(func=cmd_docx_diff)

    # --- xlsx-read ---
    p_xr = sub.add_parser("xlsx-read", help="Read XLSX → JSON snapshot (auto-hybrid, headless openpyxl fallback)")
    p_xr.add_argument("file", help="Path to source .xlsx file")
    p_xr.add_argument("--sheet", default=None,
                      help="Sheet name to export (exports all sheets if omitted)")
    p_xr.add_argument("--range", default=None,
                      help="Cell range to limit export, e.g. 'A1:H20' (requires --sheet)")
    p_xr.add_argument("--engine", default="auto", choices=["auto", "openpyxl", "xlwings"],
                      help="Reading engine: auto (default), openpyxl (headless), or xlwings (requires Excel)")
    p_xr.add_argument("-o", "--output", default=None,
                      help="Output JSON file path (prints to stdout if omitted)")
    p_xr.set_defaults(func=cmd_xlsx_read)

    # --- xlsx-write ---
    p_xw = sub.add_parser("xlsx-write", help="Write JSON snapshot → XLSX")
    p_xw.add_argument("snapshot", help="Path to .json snapshot file")
    p_xw.add_argument("--template", default=None,
                      help="Path to original .xlsx to use as base (optional)")
    p_xw.add_argument("-o", "--output", default=None,
                      help="Output .xlsx path (default: <snapshot>_written.xlsx)")
    p_xw.set_defaults(func=cmd_xlsx_write)

    # --- xlsx-mutate ---
    p_xm = sub.add_parser(
        "xlsx-mutate",
        help="Mutate Excel template in-place with row/col expansion, style cloning, and formula shifting",
    )
    p_xm.add_argument("template", help="Path to base .xlsx template file")
    p_xm.add_argument("spec", help="Path to mutation spec .json file")
    p_xm.add_argument("-o", "--output", default=None, help="Output .xlsx path (default: <template>_mutated.xlsx)")
    p_xm.add_argument("--validate", action="store_true", help="Automatically run 12 Quality Gates diff after mutating (Invariant E6)")
    p_xm.add_argument("--ref-sheet", default=None, help="Reference sheet for validation (default: auto-discover Example/Template)")
    p_xm.set_defaults(func=cmd_xlsx_mutate)

    # --- xlsx-validate ---
    p_xv = sub.add_parser(
        "xlsx-validate",
        help="Validate an Excel sheet against in-template reference sheet across 12 Quality Gates (Invariant E6)",
    )
    p_xv.add_argument("file", help="Path to .xlsx workbook")
    p_xv.add_argument("--target", required=True, help="Target sheet name to validate")
    p_xv.add_argument("--ref", default=None, help="Reference sheet name (default: auto-discover Example/Template)")
    p_xv.set_defaults(func=cmd_xlsx_validate)

    # --- pptx-create / pptx-build ---
    for cmd_name in ["pptx-create", "pptx-build"]:
        p_pc = sub.add_parser(cmd_name, help="Generate 16:9 PowerPoint deck from structured JSON spec")
        p_pc.add_argument("spec", nargs="?", default=None, help="Path to slide spec .json file")
        p_pc.add_argument("-s", "--spec", dest="spec_opt", default=None, help="Path to slide spec .json file (flag option)")
        p_pc.add_argument("-o", "--output", default=None, help="Output .pptx path")
        p_pc.add_argument("--theme", default="thesis_blue",
                          choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                          help="Visual theme (default: thesis_blue)")
        p_pc.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
        p_pc.set_defaults(func=cmd_pptx_create)

    # --- pptx-from-md ---
    p_pm = sub.add_parser("pptx-from-md", help="Compile Markdown slide deck ('---' delimiter) to .pptx")
    p_pm.add_argument("file", help="Path to markdown slides .md file")
    p_pm.add_argument("-o", "--output", default=None, help="Output .pptx path")
    p_pm.add_argument("--theme", default="thesis_blue",
                      choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                      help="Visual theme (default: thesis_blue)")
    p_pm.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
    p_pm.set_defaults(func=cmd_pptx_from_md)

    # --- docx-to-pptx ---
    p_dtp = sub.add_parser("docx-to-pptx", help="Synthesize 16:9 PowerPoint presentation deck from Word .docx document")
    p_dtp.add_argument("file", help="Path to source Word .docx file")
    p_dtp.add_argument("-o", "--output", default=None, help="Output .pptx path")
    p_dtp.add_argument("--spec-out", default=None, help="Path to save intermediate slide_spec.json")
    p_dtp.add_argument("--theme", default="thesis_blue",
                       choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                       help="Visual theme (default: thesis_blue)")
    p_dtp.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
    p_dtp.set_defaults(func=cmd_docx_to_pptx)

    # --- pdf-to-pptx ---
    p_ptp = sub.add_parser("pdf-to-pptx", help="Synthesize 16:9 PowerPoint presentation deck directly from PDF document")
    p_ptp.add_argument("file", help="Path to source PDF file")
    p_ptp.add_argument("-o", "--output", default=None, help="Output .pptx path")
    p_ptp.add_argument("--spec-out", default=None, help="Path to save intermediate slide_spec.json")
    p_ptp.add_argument("--pages", default=None, help="Page range to extract, e.g. '1-10' or '1,3,5'")
    p_ptp.add_argument("--theme", default="thesis_blue",
                       choices=["thesis_blue", "corporate_blue", "modern_dark", "academic_light"],
                       help="Visual theme (default: thesis_blue)")
    p_ptp.add_argument("--template", default=None, help="Path to base .pptx / .potx template")
    p_ptp.set_defaults(func=cmd_pdf_to_pptx)

    # --- pptx-insert-diagram ---
    p_pid = sub.add_parser("pptx-insert-diagram", help="Insert technical diagram image into a PowerPoint slide")
    p_pid.add_argument("file", help="Path to .pptx file")
    p_pid.add_argument("image", help="Path to diagram image file (.png/.jpg)")
    p_pid.add_argument("--slide", type=int, default=None, help="Slide number (1-based); appends new slide if omitted")
    p_pid.add_argument("--title", default=None, help="Slide title")
    p_pid.add_argument("--caption", default=None, help="Figure caption underneath diagram")
    p_pid.add_argument("-o", "--output", default=None, help="Output .pptx path (default: in-place safe save)")
    p_pid.add_argument("--theme", default="corporate_blue", help="Visual theme for newly created slide")
    p_pid.set_defaults(func=cmd_pptx_insert_diagram)

    # --- pptx-inspect ---
    p_pi = sub.add_parser("pptx-inspect", help="Extract PowerPoint structure, slides, text and tables into JSON AST")
    p_pi.add_argument("file", help="Path to .pptx file")
    p_pi.add_argument("-o", "--output", default=None, help="Output JSON path (prints to stdout if omitted)")
    p_pi.set_defaults(func=cmd_pptx_inspect)

    # --- pptx-preview ---
    p_pp = sub.add_parser("pptx-preview", help="Export PowerPoint slides as high-res PNG previews for QA")
    p_pp.add_argument("file", help="Path to .pptx presentation file")
    p_pp.add_argument("-o", "--output-dir", default=None, help="Directory to save preview PNGs (default: <dir>/preview_slides)")
    p_pp.add_argument("--slides", default="all", help="Comma-separated slide indices, e.g. '1,5,11,15', or 'all' (default: all)")
    p_pp.add_argument("--dpi", type=int, default=150, help="Image resolution DPI (default: 150)")
    p_pp.set_defaults(func=cmd_pptx_preview)

    # --- felo-generate ---
    p_fg = sub.add_parser("felo-generate", help="Generate AI presentation via Felo Open Platform cloud service")
    p_fg.add_argument("query", help="Presentation prompt / topic")
    p_fg.add_argument("--file", default=None, help="Optional context file or image to upload")
    p_fg.add_argument("--theme", default=None, help="Felo PPT theme ID")
    p_fg.add_argument("--task-id", default=None, help="Resume polling an existing task")
    p_fg.set_defaults(func=cmd_felo_generate)

    return parser


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
