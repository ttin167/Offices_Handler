# -*- coding: utf-8 -*-
"""
core/diagrams/code_to_diagram.py — AST Code-to-Diagram Generator & Native Draw.io SVG Exporter (v4.6).

Capabilities:
  1. scan_python_ast_to_spec: Parses Python source files or package directories using stdlib `ast`
     to extract Classes, Inheritance, Typed Attributes, and Methods (+ public, # protected, - private)
     into PlantUML Class Diagram, Mermaid classDiagram, and Draw.io mxGraph specs.
  2. scan_sql_or_prisma_to_erd_spec: Parses SQL DDL ('CREATE TABLE', 'PRIMARY KEY', 'FOREIGN KEY')
     or Prisma schema models into Mermaid erDiagram and Native Draw.io <mxGraphModel> XML (MX_INV_01..05).
  3. export_mxgraph_to_svg: Pure-Python vector SVG exporter for <mxGraphModel> XML with collision-free
     white label masks (MX_INV_08).
  4. generate_diagram_from_code: End-to-end pipeline from source code -> JSON Spec / .drawio / .svg / .png
     with automatic injection into Word (.docx) or PowerPoint (.pptx).
"""

from __future__ import annotations

import ast
import html
import os
import re
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional, Set, Tuple

from .mxgraph import (
    ColumnInfo,
    EntityInfo,
    RelationshipInfo,
    build_mxgraph_model_xml,
    calculate_auto_layout,
    spec_to_mxgraph_xml,
)


def _parse_style_dict(style_str: str) -> Dict[str, str]:
    """Parses a Draw.io style string ('key=val;rounded=1;') into a dictionary."""
    res: Dict[str, str] = {}
    for token in (style_str or "").split(";"):
        token = token.strip()
        if not token:
            continue
        if "=" in token:
            k, v = token.split("=", 1)
            res[k.strip()] = v.strip()
        else:
            res[token] = "1"
    return res


def export_mxgraph_to_svg(
    xml_content: str,
    output_svg_path: str,
    padding: int = 32,
) -> Dict[str, Any]:
    """
    Exports a pure Draw.io <mxGraphModel> XML string directly to a standalone,
    publication-grade Vector SVG file (zero external browser required), conforming to
    MX_INV_08 (white background mask on edge labels).
    """
    dest = os.path.abspath(output_svg_path)
    os.makedirs(os.path.dirname(dest), exist_ok=True)

    root = ET.fromstring(xml_content)
    cells_by_id: Dict[str, ET.Element] = {}
    geom_by_id: Dict[str, Tuple[float, float, float, float]] = {}
    parent_by_id: Dict[str, str] = {}

    for cell in root.iter("mxCell"):
        cid = cell.attrib.get("id", "")
        if not cid:
            continue
        cells_by_id[cid] = cell
        parent_by_id[cid] = cell.attrib.get("parent", "")
        geom = cell.find("mxGeometry")
        if geom is not None:
            gx = float(geom.attrib.get("x", "0") or 0.0)
            gy = float(geom.attrib.get("y", "0") or 0.0)
            gw = float(geom.attrib.get("width", "120") or 120.0)
            gh = float(geom.attrib.get("height", "50") or 50.0)
            geom_by_id[cid] = (gx, gy, gw, gh)

    def _abs_coords(cid: str) -> Tuple[float, float, float, float]:
        gx, gy, gw, gh = geom_by_id.get(cid, (0.0, 0.0, 120.0, 50.0))
        pid = parent_by_id.get(cid, "")
        while pid and pid not in ("0", "1") and pid in geom_by_id:
            px, py, _, _ = geom_by_id[pid]
            gx += px
            gy += py
            pid = parent_by_id.get(pid, "")
        return gx, gy, gw, gh

    max_x, max_y = 600.0, 400.0
    for cid, cell in cells_by_id.items():
        if cell.attrib.get("vertex") == "1":
            ax, ay, aw, ah = _abs_coords(cid)
            max_x = max(max_x, ax + aw)
            max_y = max(max_y, ay + ah)

    svg_w = int(max_x + padding * 2)
    svg_h = int(max_y + padding * 2)

    svg_parts: List[str] = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{svg_w}" height="{svg_h}" viewBox="0 0 {svg_w} {svg_h}">',
        "  <defs>",
        '    <marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">',
        '      <path d="M 0 1 L 10 5 L 0 9 z" fill="#475569"/>',
        "    </marker>",
        "  </defs>",
        f'  <rect width="{svg_w}" height="{svg_h}" fill="#FFFFFF"/>',
    ]

    # 1. Render Vertex Cells
    for cid, cell in cells_by_id.items():
        if cell.attrib.get("vertex") != "1":
            continue
        style_map = _parse_style_dict(cell.attrib.get("style", ""))
        ax, ay, aw, ah = _abs_coords(cid)
        val_raw = html.unescape(cell.attrib.get("value", "") or "")
        val_clean = re.sub(r"<[^>]+>", "", val_raw).strip()

        fill_col = style_map.get("fillColor", "#F8FAFC")
        stroke_col = style_map.get("strokeColor", "#1E293B")
        stroke_w = style_map.get("strokeWidth", "1.5")
        font_col = style_map.get("fontColor", "#0F172A")
        font_sz = style_map.get("fontSize", "13")
        rx = "6" if style_map.get("rounded") == "1" else "0"

        shape_type = style_map.get("shape", "")
        if shape_type == "table":
            svg_parts.append(
                f'  <rect x="{ax:.1f}" y="{ay:.1f}" width="{aw:.1f}" height="{ah:.1f}" '
                f'fill="{fill_col}" stroke="{stroke_col}" stroke-width="{stroke_w}"/>'
            )
            svg_parts.append(
                f'  <rect x="{ax:.1f}" y="{ay:.1f}" width="{aw:.1f}" height="43" '
                f'fill="#F1F5F9" stroke="{stroke_col}" stroke-width="{stroke_w}"/>'
            )
            if val_clean:
                svg_parts.append(
                    f'  <text x="{ax + aw/2:.1f}" y="{ay + 26:.1f}" text-anchor="middle" '
                    f'font-family="Trebuchet MS, Arial, sans-serif" font-size="14" font-weight="bold" fill="{font_col}">'
                    f"{html.escape(val_clean)}</text>"
                )
        elif shape_type == "tableRow":
            continue
        else:
            svg_parts.append(
                f'  <rect x="{ax:.1f}" y="{ay:.1f}" width="{aw:.1f}" height="{ah:.1f}" rx="{rx}" '
                f'fill="{fill_col}" stroke="{stroke_col}" stroke-width="{stroke_w}"/>'
            )
            if val_clean:
                align = style_map.get("align", "center")
                if align == "left":
                    tx = ax + 8.0
                    anchor = "start"
                else:
                    tx = ax + aw / 2.0
                    anchor = "middle"
                ty = ay + ah / 2.0 + 4.5
                svg_parts.append(
                    f'  <text x="{tx:.1f}" y="{ty:.1f}" text-anchor="{anchor}" '
                    f'font-family="Trebuchet MS, Arial, sans-serif" font-size="{font_sz}" fill="{font_col}">'
                    f"{html.escape(val_clean)}</text>"
                )

    # 2. Render Orthogonal Edges & White-Masked Labels (MX_INV_08)
    for cid, cell in cells_by_id.items():
        if cell.attrib.get("edge") != "1":
            continue
        src_id = cell.attrib.get("source", "")
        tgt_id = cell.attrib.get("target", "")
        if src_id not in geom_by_id or tgt_id not in geom_by_id:
            continue

        sx, sy, sw, sh = _abs_coords(src_id)
        tx, ty, tw, th = _abs_coords(tgt_id)
        style_map = _parse_style_dict(cell.attrib.get("style", ""))

        exit_x = float(style_map.get("exitX", "1.0") or 1.0)
        exit_y = float(style_map.get("exitY", "0.5") or 0.5)
        entry_x = float(style_map.get("entryX", "0.0") or 0.0)
        entry_y = float(style_map.get("entryY", "0.5") or 0.5)

        p1x, p1y = sx + sw * exit_x, sy + sh * exit_y
        p2x, p2y = tx + tw * entry_x, ty + th * entry_y
        mid_x = (p1x + p2x) / 2.0

        path_d = f"M {p1x:.1f} {p1y:.1f} L {mid_x:.1f} {p1y:.1f} L {mid_x:.1f} {p2y:.1f} L {p2x:.1f} {p2y:.1f}"
        stroke_col = style_map.get("strokeColor", "#475569")
        svg_parts.append(
            f'  <path d="{path_d}" fill="none" stroke="{stroke_col}" stroke-width="1.6" marker-end="url(#arrow)"/>'
        )

        lbl_raw = html.unescape(cell.attrib.get("value", "") or "")
        lbl_clean = re.sub(r"<[^>]+>", "", lbl_raw).strip()
        if lbl_clean:
            lx = mid_x
            ly = (p1y + p2y) / 2.0
            box_w = max(44.0, len(lbl_clean) * 7.2 + 12.0)
            box_h = 20.0
            svg_parts.append(
                f'  <rect x="{lx - box_w/2:.1f}" y="{ly - box_h/2:.1f}" width="{box_w:.1f}" height="{box_h:.1f}" '
                f'rx="3" fill="#FFFFFF" stroke="#E2E8F0" stroke-width="0.8"/>'
            )
            svg_parts.append(
                f'  <text x="{lx:.1f}" y="{ly + 4:.1f}" text-anchor="middle" '
                f'font-family="Trebuchet MS, Arial, sans-serif" font-size="11" fill="#0F172A">'
                f"{html.escape(lbl_clean)}</text>"
            )

    svg_parts.append("</svg>")
    svg_str = "\n".join(svg_parts)
    with open(dest, "w", encoding="utf-8") as f:
        f.write(svg_str)

    return {
        "success": True,
        "svg_path": dest,
        "dimensions_px": (svg_w, svg_h),
        "file_size_bytes": os.path.getsize(dest),
    }


def _render_svg_or_mxgraph_fallback_png(
    xml_content: str,
    output_png_path: str,
) -> Dict[str, Any]:
    """
    Renders <mxGraphModel> XML to PNG using PyMuPDF SVG rasterizer (or Pillow fallback)
    when headless browser is unavailable.
    """
    dest = os.path.abspath(output_png_path)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    tmp_svg = f"{os.path.splitext(dest)[0]}__tmp_vec.svg"
    try:
        svg_info = export_mxgraph_to_svg(xml_content, tmp_svg)
        w, h = svg_info["dimensions_px"]
        try:
            import pymupdf as fitz

            svg_doc = fitz.open(tmp_svg)
            page = svg_doc[0]
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
            pix.save(dest)
            pw, ph = pix.width, pix.height
            svg_doc.close()
            return {"png_path": dest, "dimensions_px": (pw, ph), "file_size_bytes": os.path.getsize(dest)}
        except Exception:
            from PIL import Image, ImageDraw

            img = Image.new("RGB", (max(400, w), max(300, h)), (255, 255, 255))
            draw = ImageDraw.Draw(img)
            draw.rectangle([10, 10, max(400, w) - 10, max(300, h) - 10], outline=(30, 41, 59), width=2)
            img.save(dest)
            return {"png_path": dest, "dimensions_px": img.size, "file_size_bytes": os.path.getsize(dest)}
    finally:
        if os.path.isfile(tmp_svg):
            try:
                os.remove(tmp_svg)
            except OSError:
                pass


def _visibility_prefix(name: str) -> str:
    if name.startswith("__") and not name.endswith("__"):
        return "-"
    if name.startswith("_"):
        return "#"
    return "+"


def _unparse_annotation(node: Optional[ast.AST]) -> str:
    if node is None:
        return "Any"
    try:
        return ast.unparse(node)
    except Exception:
        return "Any"


def scan_python_ast_to_spec(source_path: str) -> Dict[str, Any]:
    """
    Scans a Python file or directory of Python files using `ast` to extract classes,
    attributes, methods, inheritance, and composition relationships.
    """
    py_files: List[str] = []
    if os.path.isfile(source_path):
        py_files.append(os.path.abspath(source_path))
    elif os.path.isdir(source_path):
        for root_dir, dirs, files in os.walk(source_path):
            dirs[:] = [d for d in dirs if not d.startswith(".") and d not in ("__pycache__", "node_modules", "venv", ".venv")]
            for fn in sorted(files):
                if fn.endswith(".py"):
                    py_files.append(os.path.join(root_dir, fn))
    else:
        # Raw python code string
        py_files = []

    classes_info: Dict[str, Dict[str, Any]] = {}

    def _process_tree(tree: ast.AST) -> None:
        for node in ast.walk(tree):
            if isinstance(node, ast.ClassDef):
                cname = node.name
                bases = [_unparse_annotation(b) for b in node.bases if _unparse_annotation(b) != "object"]
                attrs: Dict[str, str] = {}
                methods: List[Dict[str, str]] = []

                for item in node.body:
                    if isinstance(item, ast.AnnAssign) and isinstance(item.target, ast.Name):
                        attrs[item.target.id] = _unparse_annotation(item.annotation)
                    elif isinstance(item, ast.Assign):
                        for t in item.targets:
                            if isinstance(t, ast.Name):
                                attrs[t.id] = "Any"
                    elif isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
                        mname = item.name
                        ret_type = _unparse_annotation(item.returns) if item.returns else "None"
                        arg_names = [a.arg for a in item.args.args if a.arg not in ("self", "cls")]
                        methods.append({
                            "name": mname,
                            "args": ", ".join(arg_names[:4]),
                            "return_type": ret_type,
                            "visibility": _visibility_prefix(mname),
                        })
                        if mname == "__init__":
                            for sub in ast.walk(item):
                                if (
                                    isinstance(sub, ast.Assign)
                                    and len(sub.targets) == 1
                                    and isinstance(sub.targets[0], ast.Attribute)
                                    and isinstance(sub.targets[0].value, ast.Name)
                                    and sub.targets[0].value.id == "self"
                                ):
                                    attrs.setdefault(sub.targets[0].attr, "Any")
                                elif (
                                    isinstance(sub, ast.AnnAssign)
                                    and isinstance(sub.target, ast.Attribute)
                                    and isinstance(sub.target.value, ast.Name)
                                    and sub.target.value.id == "self"
                                ):
                                    attrs[sub.target.attr] = _unparse_annotation(sub.annotation)

                classes_info[cname] = {
                    "name": cname,
                    "bases": bases,
                    "attributes": attrs,
                    "methods": methods,
                }

    if py_files:
        for pf in py_files:
            try:
                with open(pf, "r", encoding="utf-8", errors="ignore") as f:
                    tree = ast.parse(f.read(), filename=pf)
                _process_tree(tree)
            except Exception:
                continue
    else:
        tree = ast.parse(source_path)
        _process_tree(tree)

    # Build relationships, PlantUML, Mermaid, and mxGraph spec
    puml_lines = ["@startuml", "skinparam classAttributeIconSize 0"]
    mmd_lines = ["classDiagram"]
    nodes: List[Dict[str, Any]] = []
    edges: List[Dict[str, Any]] = []
    entities: Dict[str, EntityInfo] = {}
    rels: List[RelationshipInfo] = []

    known_classes = set(classes_info.keys())

    for cname, cinfo in classes_info.items():
        puml_lines.append(f"class {cname} {{")
        mmd_lines.append(f"  class {cname} {{")
        cols: List[ColumnInfo] = []

        for aname, atype in cinfo["attributes"].items():
            vis = _visibility_prefix(aname)
            puml_lines.append(f"  {vis}{aname} : {atype}")
            mmd_lines.append(f"    {vis}{atype} {aname}")
            cols.append(ColumnInfo(name=aname, type=atype, key="ATTR"))

        for m in cinfo["methods"]:
            puml_lines.append(f"  {m['visibility']}{m['name']}({m['args']}) : {m['return_type']}")
            mmd_lines.append(f"    {m['visibility']}{m['name']}({m['args']}) {m['return_type']}")
            cols.append(ColumnInfo(name=f"{m['name']}()", type=m["return_type"], key="METH"))

        puml_lines.append("}")
        mmd_lines.append("  }")

        entities[cname] = EntityInfo(
            name=cname,
            columns=cols or [ColumnInfo(name="pass", type="void", key="")],
            width=290.0,
            height=43.0 * (max(1, len(cols)) + 1),
        )
        nodes.append({
            "id": cname,
            "label": cname,
            "type": "table",
            "columns": [
                {"name": c.name, "type": c.type, "is_pk": c.key == "ATTR", "is_fk": c.key == "METH"}
                for c in cols
            ],
        })

        # Inheritance edges
        for base in cinfo["bases"]:
            base_clean = base.split(".")[-1]
            if base_clean in known_classes:
                puml_lines.append(f"{base_clean} <|-- {cname}")
                mmd_lines.append(f"  {base_clean} <|-- {cname} : inherits")
                edges.append({"source": cname, "target": base_clean, "label": "inherits"})
                rels.append(RelationshipInfo(source=cname, target=base_clean, label="inherits"))

        # Composition / association edges from attribute types
        for aname, atype in cinfo["attributes"].items():
            for target_cls in known_classes:
                if target_cls != cname and re.search(rf"\b{re.escape(target_cls)}\b", atype):
                    puml_lines.append(f"{cname} *-- {target_cls} : {aname}")
                    mmd_lines.append(f"  {cname} *-- {target_cls} : {aname}")
                    edges.append({"source": cname, "target": target_cls, "label": aname})
                    rels.append(RelationshipInfo(source=cname, target=target_cls, label=aname))

    puml_lines.append("@enduml")
    if entities:
        calculate_auto_layout(entities, rels)
        xml_content = build_mxgraph_model_xml(entities, rels)
    else:
        xml_content = spec_to_mxgraph_xml({"nodes": [{"id": "Empty", "label": "No Classes Found"}], "edges": []})

    puml_str = "\n".join(puml_lines)
    mmd_str = "\n".join(mmd_lines)
    return {
        "engine": "plantuml",
        "diagram_type": "class_diagram",
        "classes_count": len(classes_info),
        "code": puml_str,
        "plantuml_code": puml_str,
        "mermaid_code": mmd_str,
        "xml_content": xml_content,
        "nodes": nodes,
        "edges": edges,
    }


CREATE_TABLE_RE = re.compile(
    r"CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?[`\"'\[]?([A-Za-z0-9_]+)[`\"'\]]?\s*\((.*?)\)\s*;",
    re.IGNORECASE | re.DOTALL,
)
PRISMA_MODEL_RE = re.compile(
    r"model\s+([A-Za-z0-9_]+)\s*\{(.*?)\}",
    re.DOTALL,
)


def scan_sql_or_prisma_to_erd_spec(source_path: str) -> Dict[str, Any]:
    """
    Parses SQL DDL ('CREATE TABLE ...') or Prisma schema ('model ...') into
    a standardized ERD spec, Mermaid erDiagram, and native Draw.io <mxGraphModel> XML.
    """
    if os.path.isfile(source_path):
        with open(source_path, "r", encoding="utf-8", errors="ignore") as f:
            raw_text = f.read()
    else:
        raw_text = source_path

    entities: Dict[str, EntityInfo] = {}
    rels: List[RelationshipInfo] = []
    nodes: List[Dict[str, Any]] = []
    edges: List[Dict[str, Any]] = []

    # 1. Check SQL CREATE TABLE blocks
    sql_matches = list(CREATE_TABLE_RE.finditer(raw_text))
    if sql_matches:
        for m in sql_matches:
            tname = m.group(1).strip()
            body = m.group(2)
            cols: List[ColumnInfo] = []
            pk_cols: Set[str] = set()
            fk_cols: Set[str] = set()

            for raw_line in body.split(","):
                line = raw_line.strip()
                if not line:
                    continue
                up = line.upper()

                # Table-level PRIMARY KEY (col)
                m_pk = re.search(r"PRIMARY\s+KEY\s*\(([^)]+)\)", line, re.IGNORECASE)
                if m_pk and up.startswith(("PRIMARY", "CONSTRAINT")):
                    for pk_c in m_pk.group(1).split(","):
                        pk_cols.add(pk_c.strip().strip("`'\"[]"))
                    continue

                # Table-level FOREIGN KEY (col) REFERENCES target(tcol)
                m_fk = re.search(
                    r"FOREIGN\s+KEY\s*\(([^)]+)\)\s*REFERENCES\s+[`\"'\[]?([A-Za-z0-9_]+)[`\"'\]]?",
                    line,
                    re.IGNORECASE,
                )
                if m_fk:
                    fk_c = m_fk.group(1).strip().strip("`'\"[]")
                    ref_table = m_fk.group(2).strip()
                    fk_cols.add(fk_c)
                    rels.append(RelationshipInfo(source=tname, target=ref_table, label=fk_c))
                    edges.append({"source": tname, "target": ref_table, "label": fk_c})
                    continue

                if up.startswith(("KEY ", "INDEX ", "UNIQUE ", "CHECK ", "CONSTRAINT ")):
                    continue

                tokens = line.split()
                if len(tokens) >= 2:
                    cname = tokens[0].strip("`'\"[]")
                    ctype = re.sub(r"\(.*?\)", "", tokens[1]).lower()
                    is_pk = "PRIMARY KEY" in up
                    is_fk = "REFERENCES" in up
                    if is_pk:
                        pk_cols.add(cname)
                    if is_fk:
                        fk_cols.add(cname)
                        m_inline_ref = re.search(r"REFERENCES\s+[`\"'\[]?([A-Za-z0-9_]+)[`\"'\]]?", line, re.IGNORECASE)
                        if m_inline_ref:
                            ref_t = m_inline_ref.group(1).strip()
                            rels.append(RelationshipInfo(source=tname, target=ref_t, label=cname))
                            edges.append({"source": tname, "target": ref_t, "label": cname})
                    cols.append(ColumnInfo(name=cname, type=ctype, key=""))

            for c in cols:
                if c.name in pk_cols:
                    c.key = "PK"
                elif c.name in fk_cols:
                    c.key = "FK"

            entities[tname] = EntityInfo(
                name=tname,
                columns=cols,
                width=278.0,
                height=43.0 * (len(cols) + 1),
            )
            nodes.append({
                "id": tname,
                "label": tname,
                "columns": [
                    {"name": c.name, "type": c.type, "is_pk": c.key == "PK", "is_fk": c.key == "FK"}
                    for c in cols
                ],
            })

    # 2. Check Prisma model blocks if no SQL tables matched
    if not entities:
        for m in PRISMA_MODEL_RE.finditer(raw_text):
            mname = m.group(1).strip()
            body = m.group(2)
            cols = []
            for raw_line in body.splitlines():
                line = raw_line.strip()
                if not line or line.startswith("//") or line.startswith("@@"):
                    continue
                parts = line.split()
                if len(parts) >= 2:
                    fname = parts[0]
                    ftype = parts[1].rstrip("?[]")
                    key_tag = "PK" if "@id" in line else ("FK" if "@relation" in line or fname.endswith("Id") else "")
                    cols.append(ColumnInfo(name=fname, type=ftype, key=key_tag))
            entities[mname] = EntityInfo(
                name=mname,
                columns=cols,
                width=278.0,
                height=43.0 * (len(cols) + 1),
            )
            nodes.append({
                "id": mname,
                "label": mname,
                "columns": [
                    {"name": c.name, "type": c.type, "is_pk": c.key == "PK", "is_fk": c.key == "FK"}
                    for c in cols
                ],
            })

    mmd_lines = ["erDiagram"]
    for tname, ent in entities.items():
        mmd_lines.append(f"  {tname} {{")
        for c in ent.columns:
            k_suffix = f" {c.key}" if c.key else ""
            mmd_lines.append(f"    {c.type} {c.name}{k_suffix}")
        mmd_lines.append("  }")
    for r in rels:
        mmd_lines.append(f'  {r.target} ||--o{{ {r.source} : "{r.label or "fk"}"')

    if entities:
        calculate_auto_layout(entities, rels)
        xml_content = build_mxgraph_model_xml(entities, rels)
    else:
        xml_content = spec_to_mxgraph_xml({"nodes": [{"id": "Empty", "label": "No Tables Found"}], "edges": []})

    mmd_str = "\n".join(mmd_lines)
    return {
        "engine": "mermaid",
        "diagram_type": "erd",
        "tables_count": len(entities),
        "code": mmd_str,
        "mermaid_code": mmd_str,
        "xml_content": xml_content,
        "nodes": nodes,
        "edges": edges,
    }


def generate_diagram_from_code(
    source_path: str,
    diagram_type: str = "auto",
    output_spec_path: Optional[str] = None,
    output_png_path: Optional[str] = None,
    output_drawio_path: Optional[str] = None,
    output_svg_path: Optional[str] = None,
    inject_into: Optional[str] = None,
    title: Optional[str] = None,
    output_path: Optional[str] = None,
    engine: Optional[str] = None,
    inject_heading: Optional[str] = None,
) -> Dict[str, Any]:
    """
    End-to-end AST Code-to-Diagram generator:
      - Scans Python (.py / folder) or SQL/Prisma (.sql / .prisma)
      - Exports JSON Spec, Native Draw.io (.drawio), Pure Vector SVG (.svg), and High-Res PNG (.png)
      - Supports automatic diagram injection into Word (.docx) OR PowerPoint (.pptx).
    """
    if output_path:
        low_out = output_path.lower()
        if low_out.endswith(".svg") and not output_svg_path:
            output_svg_path = output_path
        elif low_out.endswith((".drawio", ".xml")) and not output_drawio_path:
            output_drawio_path = output_path
        elif low_out.endswith(".json") and not output_spec_path:
            output_spec_path = output_path
        elif low_out.endswith(".png") and not output_png_path:
            output_png_path = output_path
    dtype = (diagram_type or "auto").strip().lower()
    if dtype == "auto":
        if os.path.isdir(source_path) or str(source_path).lower().endswith(".py") or "class " in str(source_path):
            dtype = "class"
        else:
            dtype = "erd"

    if dtype in ("class", "class_diagram", "python", "c4"):
        spec = scan_python_ast_to_spec(source_path)
    else:
        spec = scan_sql_or_prisma_to_erd_spec(source_path)

    xml_content = spec["xml_content"]

    if output_spec_path:
        import json

        abs_spec = os.path.abspath(output_spec_path)
        os.makedirs(os.path.dirname(abs_spec), exist_ok=True)
        with open(abs_spec, "w", encoding="utf-8") as f:
            json.dump(spec, f, ensure_ascii=False, indent=2)
        spec["spec_path"] = abs_spec

    if output_drawio_path:
        abs_drawio = os.path.abspath(output_drawio_path)
        os.makedirs(os.path.dirname(abs_drawio), exist_ok=True)
        with open(abs_drawio, "w", encoding="utf-8") as f:
            f.write(xml_content)
        spec["drawio_path"] = abs_drawio

    if output_svg_path:
        svg_res = export_mxgraph_to_svg(xml_content, output_svg_path)
        spec["svg_path"] = svg_res["svg_path"]

    if output_png_path or inject_into:
        target_png = os.path.abspath(output_png_path) if output_png_path else os.path.abspath("diagram_assets/auto_code_diagram.png")
        png_res = _render_svg_or_mxgraph_fallback_png(xml_content, target_png)
        spec["png_path"] = png_res["png_path"]
        spec["dimensions_px"] = png_res["dimensions_px"]

    if inject_into and spec.get("png_path"):
        abs_inject = os.path.abspath(inject_into)
        if os.path.isfile(abs_inject):
            if abs_inject.lower().endswith(".pptx"):
                from core.pptx.engine import insert_diagram_into_pptx

                insert_diagram_into_pptx(
                    pptx_path=abs_inject,
                    image_path=spec["png_path"],
                    title=title or f"Architecture Diagram ({spec['diagram_type'].upper()})",
                    caption=title or "Auto-generated from source code AST",
                )
                spec["injected_into"] = abs_inject
            elif abs_inject.lower().endswith(".docx"):
                from core.docx.writer import inject_diagram_into_docx

                inject_diagram_into_docx(
                    docx_path=abs_inject,
                    png_path=spec["png_path"],
                    heading=inject_heading,
                    caption=title or "Auto-generated from source code AST",
                )
                spec["injected_into"] = abs_inject

    spec["success"] = True
    return spec
