"""
bridge.py — Unified Multi-Engine Diagram Bridge (v4.0).

Centralized dispatch, rendering, dimension scaling, and caching for:
1. Mermaid (.mmd, ```mermaid)
2. PlantUML (.puml, ```plantuml)
3. Draw.io / mxGraph (.drawio, .xml, ```drawio)
4. Canvas Spec Engine (.json, ```canvas, ```spec, ```diagram)

Conforms to:
- MX_INV_01..06, MX_INV_08 (Draw.io Native XML & Collision-Free Masks)
- ERR_DOCX_005 (Printable Margin Overflow: W <= 14cm, H <= 18cm)
- PPTX_INV_02, PPTX_INV_05 (Aspect Ratio Lock & Diagram Centering)
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sys
from typing import Any, Dict, Optional, Tuple
from PIL import Image

# Ensure stdout/stderr handle UTF-8 on Windows
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

DIAGRAM_LANGUAGES = {
    "mermaid", "mmd",
    "plantuml", "puml",
    "drawio", "mxgraph", "xml",
    "canvas", "spec", "diagram-spec", "diagram", "json",
}

DIAGRAM_EXTENSIONS = {
    ".mmd", ".mermaid",
    ".puml", ".plantuml",
    ".drawio", ".xml",
    ".json",
}

IMAGE_EXTENSIONS = {
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tiff",
}


def calculate_aspect_dimensions(
    width_px: int,
    height_px: int,
    max_w_cm: float = 14.0,
    max_h_cm: float = 18.0,
) -> Tuple[float, float]:
    """
    Computes (width_cm, height_cm) fitting within max_w_cm and max_h_cm
    while strictly preserving the original aspect ratio (ERR_DOCX_005).
    """
    if width_px <= 0 or height_px <= 0:
        return max_w_cm, max_h_cm

    scaled_h = (height_px / width_px) * max_w_cm
    if scaled_h <= max_h_cm:
        return max_w_cm, round(scaled_h, 2)
    else:
        scaled_w = (width_px / height_px) * max_h_cm
        return round(min(scaled_w, max_w_cm), 2), max_h_cm


def _extract_diagram_title(code: str, lang: str) -> Optional[str]:
    """Extract optional title from diagram source code."""
    # Mermaid or PlantUML title: title "..." or title ...
    m = re.search(r'^\s*title\s+["\']?([^"\']+)["\']?', code, re.IGNORECASE | re.MULTILINE)
    if m:
        return m.group(1).strip()
    return None


def render_diagram_code_to_png(
    code: str,
    lang: str,
    output_png: Optional[str] = None,
    base_dir: Optional[str] = None,
    scale: int = 2,
) -> Optional[Dict[str, Any]]:
    """
    Renders diagram source code into a high-res PNG file.
    Returns metadata dict on success, None on failure.
    """
    clean_code = code.strip()
    if not clean_code:
        return None

    clean_lang = (lang or "").strip().lower()
    if clean_lang not in DIAGRAM_LANGUAGES:
        # Fallback check for XML mxGraphModel
        if "<mxGraphModel" in clean_code:
            clean_lang = "drawio"
        else:
            return None

    # Compute content hash for disk caching
    code_hash = hashlib.md5(clean_code.encode("utf-8")).hexdigest()[:12]

    # Destination PNG path
    if not output_png:
        target_dir = os.path.join(base_dir or os.getcwd(), "diagram_assets")
        os.makedirs(target_dir, exist_ok=True)
        output_png = os.path.join(target_dir, f"diag_{clean_lang}_{code_hash}.png")

    output_png = os.path.abspath(output_png)
    os.makedirs(os.path.dirname(output_png), exist_ok=True)

    # Return cached render if valid
    if os.path.isfile(output_png) and os.path.getsize(output_png) > 0:
        try:
            with Image.open(output_png) as im:
                w, h = im.size
            title = _extract_diagram_title(clean_code, clean_lang)
            return {
                "success": True,
                "png_path": output_png,
                "engine": clean_lang,
                "dimensions_px": (w, h),
                "title": title,
                "cached": True,
            }
        except Exception:
            pass

    # Dispatch to appropriate rendering engine
    try:
        # 1. Mermaid Engine
        if clean_lang in ("mermaid", "mmd"):
            try:
                from core.diagrams.legacy.mermaid import render_mermaid_to_png
            except (ImportError, ValueError):
                try:
                    from legacy_engines.mermaid_renderer import render_mermaid_to_png
                except (ImportError, ValueError):
                    from mermaid_renderer import render_mermaid_to_png

            res = render_mermaid_to_png({
                "mermaid_code": clean_code,
                "output_path": output_png,
                "scale": scale,
            }, base_dir=base_dir)
            if not res or not res.get("success") or not os.path.isfile(output_png):
                return None

        # 2. PlantUML Engine
        elif clean_lang in ("plantuml", "puml"):
            try:
                from core.diagrams.legacy.plantuml import render_plantuml_to_png
            except (ImportError, ValueError):
                try:
                    from legacy_engines.plantuml_renderer import render_plantuml_to_png
                except (ImportError, ValueError):
                    from plantuml_renderer import render_plantuml_to_png

            puml_code = clean_code
            if not puml_code.startswith("@startuml"):
                puml_code = f"@startuml\n{puml_code}\n@enduml"

            res = render_plantuml_to_png({
                "plantuml_code": puml_code,
                "output_path": output_png,
                "dpi": 300,
            }, base_dir=base_dir)
            if not res or not res.get("success") or not os.path.isfile(output_png):
                return None

        # 3. Draw.io / mxGraph Engine
        elif clean_lang in ("drawio", "mxgraph") or (clean_lang == "xml" and "<mxGraphModel" in clean_code):
            try:
                from core.diagrams.mxgraph import render_mxgraph_to_png
            except (ImportError, ValueError):
                from mxgraph_engine import render_mxgraph_to_png

            res = render_mxgraph_to_png(clean_code, output_png, scale=scale)
            actual_png = res.get("png_path") if isinstance(res, dict) else (res if isinstance(res, str) else output_png)
            if not actual_png or not os.path.isfile(actual_png):
                return None
            output_png = actual_png

        # 4. Canvas Spec / JSON Engine
        elif clean_lang in ("canvas", "spec", "diagram-spec", "diagram", "json"):
            try:
                spec_dict = json.loads(clean_code)
            except Exception:
                return None

            if not isinstance(spec_dict, dict):
                return None

            spec_dict["output_path"] = output_png
            spec_dict["scale"] = scale
            engine = str(spec_dict.get("engine", "")).strip().lower()

            if engine == "mermaid":
                return render_diagram_code_to_png(
                    spec_dict.get("mermaid_code") or spec_dict.get("code", ""),
                    "mermaid",
                    output_png=output_png,
                    base_dir=base_dir,
                    scale=scale,
                )
            elif engine == "plantuml":
                return render_diagram_code_to_png(
                    spec_dict.get("plantuml_code") or spec_dict.get("code", ""),
                    "plantuml",
                    output_png=output_png,
                    base_dir=base_dir,
                    scale=scale,
                )
            elif engine == "canvas" or ("nodes" in spec_dict and "edges" in spec_dict):
                try:
                    from core.diagrams.legacy.canvas import PrecisionDiagram
                except (ImportError, ValueError):
                    from legacy_engines.spec_diagram_engine import PrecisionDiagram

                diag = PrecisionDiagram.from_spec(spec_dict, base_dir=base_dir)
                res = diag.render_to_png(output_png, scale=scale)
                if not res or not res.get("success") or not os.path.isfile(output_png):
                    return None
            else:
                try:
                    from core.diagrams.mxgraph import render_diagram
                except (ImportError, ValueError):
                    from mxgraph_engine import render_diagram

                res = render_diagram(spec_dict, base_dir=base_dir)
                if not res or not os.path.isfile(output_png):
                    return None
        else:
            return None

        if os.path.isfile(output_png) and os.path.getsize(output_png) > 0:
            with Image.open(output_png) as im:
                w, h = im.size
            title = _extract_diagram_title(clean_code, clean_lang)
            return {
                "success": True,
                "png_path": output_png,
                "engine": clean_lang,
                "dimensions_px": (w, h),
                "title": title,
                "cached": False,
            }
    except Exception as exc:
        print(f"[WARN] Diagram bridge failed to render ({clean_lang}): {exc}", file=sys.stderr)
        return None

    return None


def render_diagram_file_to_png(
    file_path: str,
    output_png: Optional[str] = None,
    base_dir: Optional[str] = None,
    scale: int = 2,
) -> Optional[Dict[str, Any]]:
    """
    Renders or resolves a diagram file (.json, .puml, .mmd, .drawio, .xml, .png)
    to a high-resolution PNG image.
    """
    if not file_path:
        return None

    abs_path = file_path if os.path.isabs(file_path) else os.path.abspath(
        os.path.join(base_dir or os.getcwd(), file_path)
    )

    ext = os.path.splitext(abs_path)[1].lower()

    # Case 1: File is already an image
    if ext in IMAGE_EXTENSIONS:
        if os.path.isfile(abs_path) and os.path.getsize(abs_path) > 0:
            try:
                with Image.open(abs_path) as im:
                    w, h = im.size
                return {
                    "success": True,
                    "png_path": abs_path,
                    "engine": "image",
                    "dimensions_px": (w, h),
                    "title": os.path.splitext(os.path.basename(abs_path))[0],
                    "cached": True,
                }
            except Exception:
                pass

        # If PNG doesn't exist, check if sibling spec exists (e.g. flow.png -> flow.json or flow.drawio)
        base_stem = os.path.splitext(abs_path)[0]
        for candidate_ext in (".json", ".drawio", ".xml", ".puml", ".mmd"):
            candidate_file = base_stem + candidate_ext
            if os.path.isfile(candidate_file):
                return render_diagram_file_to_png(candidate_file, output_png=abs_path, base_dir=base_dir, scale=scale)
        return None

    # Case 2: File is a diagram source/spec file
    if not os.path.isfile(abs_path):
        return None

    with open(abs_path, "r", encoding="utf-8") as f:
        content = f.read()

    lang_map = {
        ".mmd": "mermaid",
        ".mermaid": "mermaid",
        ".puml": "plantuml",
        ".plantuml": "plantuml",
        ".drawio": "drawio",
        ".xml": "drawio",
        ".json": "json",
    }
    lang = lang_map.get(ext, "spec")

    if not output_png:
        out_dir = os.path.join(os.path.dirname(abs_path), "diagram_assets")
        os.makedirs(out_dir, exist_ok=True)
        output_png = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(abs_path))[0]}.png")

    return render_diagram_code_to_png(
        code=content,
        lang=lang,
        output_png=output_png,
        base_dir=os.path.dirname(abs_path),
        scale=scale,
    )
