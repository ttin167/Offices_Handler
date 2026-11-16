"""
core.docx.toolkit — Pure Office Word Mastery Toolkit.

Provides headless OpenXML automation for everyday administrative Word workflows:
1. polish_docx_document: Apply Vietnamese Decree 30/2020/NĐ-CP or Executive presets,
   VN typography cleanup, table pagination flags (cantSplit / tblHeader / trailing <w:p>),
   image printable-width clamping (ERR_DOCX_005), and dynamic Page X / Y numbering.
2. merge_docx_files: Concatenate multiple .docx documents cleanly while remapping
   image and hyperlink relationship IDs so embedded media renders without corruption.
3. split_docx_by_heading: Split a large Word document into per-chapter .docx files + .zip.
4. apply_docx_watermark_and_hf: Inject native diagonal VML watermark, headers/footers,
   and dynamic page numbering across all document sections.
"""

from __future__ import annotations

import copy
import io
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import docx
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

from .advanced_engine import (
    format_figure_captions,
    inject_dynamic_page_numbers,
    safe_save_docx,
    sanitize_emojis_and_symbols,
    sanitize_vietnamese_typography,
)
from .presets import apply_preset_to_document, build_administrative_header

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_IMAGE_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


def _safe_slug(text: str, fallback: str = "Section") -> str:
    cleaned = re.sub(r'[\\/*?:"<>|\x00-\x1f]+', "_", str(text or "").strip())
    cleaned = cleaned.strip(". ")
    return cleaned[:70] or fallback


def _enforce_docx_table_and_image_invariants(doc: docx.Document) -> Dict[str, int]:
    """
    Enforce ERR_DOCX_001 (last <w:p> in <w:tc>), ERR_DOCX_003 (cantSplit & tblHeader),
    and ERR_DOCX_005 (clamp inline images inside printable page width).
    """
    tables_fixed = 0
    images_clamped = 0

    for tbl in doc.tables:
        if not tbl.rows:
            continue
        tables_fixed += 1
        for r_idx, row in enumerate(tbl.rows):
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))
            if r_idx == 0 and trPr.find(qn("w:tblHeader")) is None:
                trPr.append(OxmlElement("w:tblHeader"))
            for cell in row.cells:
                # ERR_DOCX_001: <w:tc> must end with a <w:p>
                tc = cell._tc
                if len(tc) == 0 or tc[-1].tag != qn("w:p"):
                    tc.append(OxmlElement("w:p"))

    # ERR_DOCX_005: Printable body width guard
    sec = doc.sections[0]
    pw = int(sec.page_width or Cm(21.0))
    lm = int(sec.left_margin or Cm(3.0))
    rm = int(sec.right_margin or Cm(1.5))
    max_w_emu = max(int(Cm(10.0)), min(int(Cm(14.5)), pw - lm - rm))

    for shape in doc.inline_shapes:
        try:
            if shape.width and shape.width > max_w_emu and shape.width > 0:
                ratio = max_w_emu / float(shape.width)
                shape.width = max_w_emu
                if shape.height:
                    shape.height = int(shape.height * ratio)
                images_clamped += 1
        except Exception:
            continue

    return {"tables_fixed": tables_fixed, "images_clamped": images_clamped}


def _remap_element_relationships(
    element: Any, src_part: Any, dst_part: Any
) -> None:
    """Remap r:embed / r:id references when cloning XML paragraphs or tables across docs."""
    rid_map: Dict[str, str] = {}
    attr_keys = (qn("r:embed"), qn("r:id"), qn("r:link"))

    for node in element.iter():
        for attr in attr_keys:
            old_rid = node.attrib.get(attr)
            if not old_rid:
                continue
            if old_rid in rid_map:
                node.attrib[attr] = rid_map[old_rid]
                continue
            rel = src_part.rels.get(old_rid)
            if rel is None:
                continue
            try:
                if rel.is_external:
                    new_rid = dst_part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True
                    )
                elif rel.reltype == _IMAGE_RELTYPE and hasattr(rel, "target_part"):
                    _, new_rid = dst_part.get_or_add_image(
                        io.BytesIO(rel.target_part.blob)
                    )
                elif hasattr(rel, "target_part"):
                    new_rid = dst_part.relate_to(rel.target_part, rel.reltype)
                else:
                    continue
                rid_map[old_rid] = new_rid
                node.attrib[attr] = new_rid
            except Exception:
                continue


def _remap_num_ids_in_element(element: Any, num_id_map: Dict[str, str]) -> None:
    """Rewrite <w:numId w:val="..."/> references inside a cloned XML subtree."""
    if not num_id_map:
        return
    num_id_tag = qn("w:numId")
    val_attr = qn("w:val")
    for node in element.iter():
        if node.tag == num_id_tag:
            old_val = node.get(val_attr)
            if old_val and old_val in num_id_map:
                node.set(val_attr, num_id_map[old_val])


def _merge_numbering_definitions(
    sub_doc: docx.Document, master: docx.Document
) -> Dict[str, str]:
    """
    Merge numbering definitions (abstractNum + num) from `sub_doc` into `master`
    with collision-free IDs so bullet/numbered lists retain their definitions and restarts.
    Returns mapping `{old_numId_str: new_numId_str}`.
    """
    try:
        sub_num_part = getattr(sub_doc.part, "numbering_part", None)
    except Exception:
        return {}
    if sub_num_part is None:
        return {}

    sub_nums = sub_num_part.element.findall(qn("w:num"))
    if not sub_nums:
        return {}

    try:
        master_num_part = master.part.numbering_part
    except Exception:
        return {}
    if master_num_part is None:
        return {}

    m_el = master_num_part.element
    s_el = sub_num_part.element

    existing_abs_ids = {
        int(el.get(qn("w:abstractNumId")))
        for el in m_el.findall(qn("w:abstractNum"))
        if str(el.get(qn("w:abstractNumId") or "")).isdigit()
    }
    existing_num_ids = {
        int(el.get(qn("w:numId")))
        for el in m_el.findall(qn("w:num"))
        if str(el.get(qn("w:numId") or "")).isdigit()
    }
    next_abs_id = max(existing_abs_ids, default=0) + 1
    next_num_id = max(existing_num_ids, default=0) + 1

    # Index sub abstractNum definitions
    sub_abs_map: Dict[str, Any] = {}
    for el in s_el.findall(qn("w:abstractNum")):
        aid = el.get(qn("w:abstractNumId"))
        if aid is not None:
            sub_abs_map[aid] = el

    abs_id_map: Dict[str, str] = {}
    num_id_map: Dict[str, str] = {}

    first_num_el = m_el.find(qn("w:num"))
    for num_el in sub_nums:
        old_num_id = num_el.get(qn("w:numId"))
        if not old_num_id:
            continue
        abs_ref = num_el.find(qn("w:abstractNumId"))
        old_abs_id = abs_ref.get(qn("w:val")) if abs_ref is not None else None

        new_abs_id: Optional[str] = None
        if old_abs_id is not None and old_abs_id in sub_abs_map:
            if old_abs_id in abs_id_map:
                new_abs_id = abs_id_map[old_abs_id]
            else:
                new_abs_id = str(next_abs_id)
                next_abs_id += 1
                cloned_abs = copy.deepcopy(sub_abs_map[old_abs_id])
                cloned_abs.set(qn("w:abstractNumId"), new_abs_id)
                _remap_element_relationships(cloned_abs, sub_num_part, master_num_part)
                # ECMA-376 requires all <w:abstractNum> elements to appear before <w:num>
                if first_num_el is not None:
                    first_num_el.addprevious(cloned_abs)
                else:
                    m_el.append(cloned_abs)
                abs_id_map[old_abs_id] = new_abs_id

        new_num_id = str(next_num_id)
        next_num_id += 1
        cloned_num = copy.deepcopy(num_el)
        cloned_num.set(qn("w:numId"), new_num_id)
        cloned_abs_ref = cloned_num.find(qn("w:abstractNumId"))
        if cloned_abs_ref is not None and new_abs_id is not None:
            cloned_abs_ref.set(qn("w:val"), new_abs_id)
        m_el.append(cloned_num)
        num_id_map[old_num_id] = new_num_id

    return num_id_map


def _merge_missing_styles(
    sub_doc: docx.Document, master: docx.Document, num_id_map: Dict[str, str]
) -> int:
    """Copy non-colliding custom styles from `sub_doc` into `master`."""
    try:
        m_styles_el = master.styles.element
        s_styles_el = sub_doc.styles.element
    except Exception:
        return 0

    existing_ids = {
        el.get(qn("w:styleId"))
        for el in m_styles_el.findall(qn("w:style"))
        if el.get(qn("w:styleId"))
    }
    added = 0
    for st_el in s_styles_el.findall(qn("w:style")):
        sid = st_el.get(qn("w:styleId"))
        if not sid or sid in existing_ids:
            continue
        cloned_st = copy.deepcopy(st_el)
        _remap_num_ids_in_element(cloned_st, num_id_map)
        m_styles_el.append(cloned_st)
        existing_ids.add(sid)
        added += 1
    return added


def _hf_has_content(hf: Any) -> bool:
    """Return True if a section header/footer contains visible text, tables, or drawings."""
    for p in getattr(hf, "paragraphs", []):
        if _paragraph_has_content(p):
            return True
    if getattr(hf, "tables", []):
        return True
    return False


def _section_has_custom_hf(sec: Any) -> bool:
    for attr in ("header", "footer", "first_page_header", "first_page_footer"):
        hf = getattr(sec, attr, None)
        if hf is not None and _hf_has_content(hf):
            return True
    return False


def _copy_hf_content(src_hf: Any, dst_hf: Any, num_id_map: Dict[str, str]) -> None:
    """Copy header/footer XML body elements cleanly with remapped relationships and numIds."""
    dst_hf.is_linked_to_previous = False
    dst_el = dst_hf._element
    for existing in list(dst_el):
        dst_el.remove(existing)
    for child in list(src_hf._element):
        cloned = copy.deepcopy(child)
        _remap_element_relationships(cloned, src_hf.part, dst_hf.part)
        _remap_num_ids_in_element(cloned, num_id_map)
        dst_el.append(cloned)


def polish_docx_document(
    input_path: str,
    output_path: Optional[str] = None,
    preset: str = "nd30",
    add_page_numbers: bool = True,
    sanitize_typography: bool = True,
    fix_tables: bool = True,
    normalize_captions: bool = False,
    admin_header: Optional[Dict[str, str]] = None,
) -> Dict[str, Any]:
    """
    Apply full Vietnamese administrative / executive Word formatting & OpenXML cleanup:
    - Apply preset (`nd30`, `nd30_report`, `modern_report`) via `core.docx.presets`.
    - Optionally prepend Vietnamese Decree 30/2020/NĐ-CP administrative header block.
    - Clean quotes, dashes, emojis, enforce cantSplit/tblHeader/trailing <w:p>, clamp images.
    - Inject bottom-right / bottom-center dynamic `Trang X / Y` footer numbering.
    """
    src = Path(input_path)
    if not src.is_file():
        raise FileNotFoundError(f"DOCX file not found: {input_path}")

    dst = Path(output_path) if output_path else src.with_name(f"{src.stem}_polished.docx")
    dst.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(str(src))
    preset_res = apply_preset_to_document(doc, preset=preset)

    if admin_header and any(str(v).strip() for v in admin_header.values()):
        build_administrative_header(
            doc,
            agency_parent=str(admin_header.get("agency_parent") or "ỦY BAN NHÂN DÂN / TẬP ĐOÀN"),
            agency_name=str(admin_header.get("agency_name") or "VĂN PHÒNG / PHÒNG HÀNH CHÍNH"),
            document_no=str(admin_header.get("document_no") or "Số: .../CV-VP"),
            location_date=str(admin_header.get("location_date") or "Hà Nội, ngày ... tháng ... năm 2026"),
            insert_at_top=True,
        )

    typo_changes = 0
    emoji_changes = 0
    if sanitize_typography:
        typo_changes = sanitize_vietnamese_typography(doc)
        emoji_changes = sanitize_emojis_and_symbols(doc)

    captions_formatted = 0
    if normalize_captions:
        captions_formatted = format_figure_captions(doc)

    inv_stats = {"tables_fixed": 0, "images_clamped": 0}
    if fix_tables:
        inv_stats = _enforce_docx_table_and_image_invariants(doc)

    if add_page_numbers:
        inject_dynamic_page_numbers(doc)

    saved_path = safe_save_docx(doc, str(dst))
    return {
        "ok": True,
        "input_path": str(src),
        "output_path": saved_path,
        "preset_applied": preset_res,
        "typo_changes": typo_changes,
        "emoji_changes": emoji_changes,
        "captions_formatted": captions_formatted,
        "page_numbers_added": add_page_numbers,
        **inv_stats,
    }


def _extract_style_num_ids(doc_obj: docx.Document) -> Dict[str, str]:
    """Return mapping `{styleId: numId_str}` for styles that attach numbering."""
    res: Dict[str, str] = {}
    try:
        styles_el = doc_obj.styles.element
    except Exception:
        return res
    for st_el in styles_el.findall(qn("w:style")):
        sid = st_el.get(qn("w:styleId"))
        pPr = st_el.find(qn("w:pPr"))
        if not sid or pPr is None:
            continue
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            continue
        numId = numPr.find(qn("w:numId"))
        if numId is not None and numId.get(qn("w:val")):
            res[sid] = str(numId.get(qn("w:val")))
    return res


def merge_docx_files(
    input_paths: Sequence[str],
    output_path: str,
    add_page_breaks: bool = True,
    normalize_page_numbers: bool = True,
    preserve_sections: bool = True,
) -> Dict[str, Any]:
    """
    Merge multiple Word `.docx` files sequentially into one clean master document,
    remapping media/hyperlink relationships, merging `numbering.xml` (`w:numId` /
    `w:abstractNumId`) and missing `styles.xml` definitions, and preserving distinct
    sub-document headers/footers across sections.
    """
    valid_paths = [Path(p) for p in input_paths if p and Path(p).is_file()]
    if not valid_paths:
        raise ValueError("At least one valid DOCX file is required for merging.")

    dst = Path(output_path)
    dst.parent.mkdir(parents=True, exist_ok=True)

    master = Document(str(valid_paths[0]))
    master_body = master.element.body

    def _insert_before_sect(el: Any) -> None:
        sect_pr = master_body.find(qn("w:sectPr"))
        if sect_pr is not None:
            sect_pr.addprevious(el)
        else:
            master_body.append(el)

    merged_count = 1
    numbering_remapped = 0
    styles_merged = 0
    sections_preserved = 1

    for sub_path in valid_paths[1:]:
        sub_doc = Document(str(sub_path))
        num_id_map = _merge_numbering_definitions(sub_doc, master)
        numbering_remapped += len(num_id_map)
        styles_merged += _merge_missing_styles(sub_doc, master, num_id_map)
        sub_style_nums = _extract_style_num_ids(sub_doc)

        sub_sec = sub_doc.sections[0] if sub_doc.sections else None
        has_custom_hf = bool(sub_sec is not None and _section_has_custom_hf(sub_sec))

        if preserve_sections and has_custom_hf and sub_sec is not None:
            start_type = WD_SECTION.NEW_PAGE if add_page_breaks else WD_SECTION.CONTINUOUS
            new_sec = master.add_section(start_type)
            sections_preserved += 1
            for dim_attr in (
                "page_width",
                "page_height",
                "left_margin",
                "right_margin",
                "top_margin",
                "bottom_margin",
            ):
                val = getattr(sub_sec, dim_attr, None)
                if val is not None:
                    try:
                        setattr(new_sec, dim_attr, val)
                    except Exception:
                        pass
            new_sec.different_first_page_header_footer = bool(
                sub_sec.different_first_page_header_footer
            )
            _copy_hf_content(sub_sec.header, new_sec.header, num_id_map)
            _copy_hf_content(sub_sec.footer, new_sec.footer, num_id_map)
            if sub_sec.different_first_page_header_footer:
                _copy_hf_content(sub_sec.first_page_header, new_sec.first_page_header, num_id_map)
                _copy_hf_content(sub_sec.first_page_footer, new_sec.first_page_footer, num_id_map)
        elif add_page_breaks:
            p_break = OxmlElement("w:p")
            r_el = OxmlElement("w:r")
            br_el = OxmlElement("w:br")
            br_el.set(qn("w:type"), "page")
            r_el.append(br_el)
            p_break.append(r_el)
            _insert_before_sect(p_break)

        for child in list(sub_doc.element.body):
            if child.tag == qn("w:sectPr"):
                continue
            cloned = copy.deepcopy(child)
            _remap_element_relationships(cloned, sub_doc.part, master.part)
            _remap_num_ids_in_element(cloned, num_id_map)

            # Isolate style-bound list numbering (e.g. ListNumber / ListBullet) onto remapped numId
            if num_id_map and sub_style_nums:
                for p_el in cloned.iter(qn("w:p")):
                    pPr = p_el.find(qn("w:pPr"))
                    if pPr is None:
                        continue
                    if pPr.find(qn("w:numPr")) is not None:
                        continue
                    pStyle = pPr.find(qn("w:pStyle"))
                    if pStyle is None:
                        continue
                    sid = pStyle.get(qn("w:val"))
                    if sid and sid in sub_style_nums:
                        mapped = num_id_map.get(sub_style_nums[sid])
                        if mapped:
                            numPr = OxmlElement("w:numPr")
                            numId_el = OxmlElement("w:numId")
                            numId_el.set(qn("w:val"), mapped)
                            numPr.append(numId_el)
                            pPr.append(numPr)

            _insert_before_sect(cloned)
        merged_count += 1

    _enforce_docx_table_and_image_invariants(master)
    if normalize_page_numbers:
        inject_dynamic_page_numbers(master)

    saved_path = safe_save_docx(master, str(dst))
    return {
        "ok": True,
        "merged_files": merged_count,
        "output_path": saved_path,
        "numbering_remapped": numbering_remapped,
        "styles_merged": styles_merged,
        "sections_preserved": sections_preserved,
        "sources": [str(p) for p in valid_paths],
    }


def split_docx_by_heading(
    input_path: str,
    output_path: Optional[str] = None,
    heading_level: int = 1,
    as_zip: bool = True,
) -> Dict[str, Any]:
    """
    Split a master `.docx` file into separate `.docx` chapter files whenever a
    paragraph with `Heading {heading_level}` appears, bundled into a ZIP archive.
    Cleans up temporary staging directories cleanly when `as_zip=True`.
    """
    src = Path(input_path)
    if not src.is_file():
        raise FileNotFoundError(f"DOCX file not found: {input_path}")

    doc = Document(str(src))
    target_style_names = {
        f"heading {heading_level}",
        f"heading{heading_level}",
    }

    sections_data: List[Dict[str, Any]] = []
    current_title = f"{src.stem}_Intro"
    current_elements: List[Any] = []

    for child in doc.element.body:
        if child.tag == qn("w:sectPr"):
            continue
        is_split_heading = False
        heading_text = ""
        if child.tag == qn("w:p"):
            p_obj = docx.text.paragraph.Paragraph(child, doc)
            st_name = (p_obj.style.name or "").strip().lower() if p_obj.style else ""
            txt = p_obj.text.strip()
            if st_name in target_style_names and txt:
                is_split_heading = True
                heading_text = txt

        if is_split_heading:
            if current_elements:
                sections_data.append({"title": current_title, "elements": current_elements})
            current_title = heading_text
            current_elements = [child]
        else:
            current_elements.append(child)

    if current_elements:
        sections_data.append({"title": current_title, "elements": current_elements})

    if as_zip:
        out_zip = Path(output_path) if output_path else src.with_name(f"{src.stem}_chapters.zip")
        out_zip.parent.mkdir(parents=True, exist_ok=True)
        work_dir = Path(tempfile.mkdtemp(prefix="antigravity_docx_split_"))
    else:
        work_dir = Path(output_path) if output_path else src.with_name(f"{src.stem}_chapters")
        work_dir.mkdir(parents=True, exist_ok=True)
        out_zip = work_dir

    # Use original doc as template to retain exact styles, margins, fonts
    base_bytes = src.read_bytes()
    generated: List[Dict[str, Any]] = []
    used_names: set[str] = set()

    try:
        for idx, sec in enumerate(sections_data, start=1):
            slug = _safe_slug(sec["title"], fallback=f"Section_{idx}")
            fname = f"{idx:02d}_{slug}.docx"
            n = 2
            while fname.lower() in used_names:
                fname = f"{idx:02d}_{slug}_{n}.docx"
                n += 1
            used_names.add(fname.lower())

            child_doc = Document(io.BytesIO(base_bytes))
            body = child_doc.element.body
            sectPr = body.find(qn("w:sectPr"))
            for el in list(body):
                if el.tag != qn("w:sectPr"):
                    body.remove(el)

            for orig_el in sec["elements"]:
                cloned = copy.deepcopy(orig_el)
                _remap_element_relationships(cloned, doc.part, child_doc.part)
                if sectPr is not None:
                    sectPr.addprevious(cloned)
                else:
                    body.append(cloned)

            out_file = work_dir / fname
            child_doc.save(str(out_file))
            generated.append(
                {
                    "title": sec["title"],
                    "filename": fname,
                    "path": str(out_zip) if as_zip else str(out_file),
                    "_local_file": str(out_file),
                }
            )

        if as_zip:
            with zipfile.ZipFile(str(out_zip), "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for item in generated:
                    zf.write(item["_local_file"], arcname=item["filename"])

        for item in generated:
            item.pop("_local_file", None)

        return {
            "ok": True,
            "input_path": str(src),
            "output_path": str(out_zip),
            "heading_level": heading_level,
            "total_chapters": len(generated),
            "chapters": generated,
        }
    finally:
        if as_zip and work_dir.exists():
            shutil.rmtree(work_dir, ignore_errors=True)


def _build_watermark_run_xml(text: str) -> Any:
    """Create native diagonal VML WordArt watermark <w:r> element."""
    safe_txt = (
        str(text or "BẢN NHÁP")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
    xml_str = (
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:o="urn:schemas-microsoft-com:office:office">'
        "<w:rPr><w:noProof/></w:rPr>"
        "<w:pict>"
        '<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;width:460pt;height:140pt;'
        'rotation:315;z-index:-251657216;mso-position-horizontal:center;'
        'mso-position-horizontal-relative:margin;mso-position-vertical:center;'
        'mso-position-vertical-relative:margin" '
        'fillcolor="#CBD5E1" stroked="f">'
        '<v:fill opacity=".35"/>'
        f'<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="{safe_txt}"/>'
        "</v:shape>"
        "</w:pict>"
        "</w:r>"
    )
    return parse_xml(xml_str)


def _paragraph_has_content(p: Any) -> bool:
    """Return True if a paragraph contains visible text, drawings, shapes, or field codes."""
    if (p.text or "").strip():
        return True
    return bool(
        p._p.xpath(
            ".//*[local-name()='drawing' or local-name()='pict' or local-name()='fldChar' or local-name()='instrText']"
        )
    )


def apply_docx_watermark_and_hf(
    input_path: str,
    output_path: Optional[str] = None,
    watermark_text: Optional[str] = None,
    header_text: Optional[str] = None,
    footer_text: Optional[str] = None,
    add_page_numbers: bool = True,
) -> Dict[str, Any]:
    """
    Inject clean header text, footer text, dynamic Page X/Y numbering, and a
    native diagonal VML Word watermark across every section of a Word document
    without overwriting existing header/footer logos or fields.
    """
    src = Path(input_path)
    if not src.is_file():
        raise FileNotFoundError(f"DOCX file not found: {input_path}")

    dst = Path(output_path) if output_path else src.with_name(f"{src.stem}_branded.docx")
    dst.parent.mkdir(parents=True, exist_ok=True)

    doc = Document(str(src))
    for sec in doc.sections:
        sec.different_first_page_header_footer = False
        hdr = sec.header
        if header_text and header_text.strip():
            first_hp = hdr.paragraphs[0] if hdr.paragraphs else None
            if first_hp is not None and not _paragraph_has_content(first_hp):
                hp = first_hp
            else:
                hp = hdr.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            hr = hp.add_run(header_text.strip())
            hr.font.name = "Times New Roman"
            hr.font.size = Pt(9.5)
            hr.font.italic = True
            hr.font.color.rgb = RGBColor(100, 116, 139)

        if watermark_text and watermark_text.strip():
            wp = hdr.add_paragraph()
            wp._p.append(_build_watermark_run_xml(watermark_text.strip()))

        ftr = sec.footer
        if footer_text and footer_text.strip():
            first_fp = ftr.paragraphs[0] if ftr.paragraphs else None
            if first_fp is not None and not _paragraph_has_content(first_fp):
                fp = first_fp
            else:
                fp = ftr.add_paragraph()
            fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            fr = fp.add_run(footer_text.strip())
            fr.font.name = "Times New Roman"
            fr.font.size = Pt(9.5)
            fr.font.color.rgb = RGBColor(100, 116, 139)

    if add_page_numbers:
        inject_dynamic_page_numbers(doc)

    saved_path = safe_save_docx(doc, str(dst))
    return {
        "ok": True,
        "input_path": str(src),
        "output_path": saved_path,
        "watermark_text": watermark_text or "",
        "header_text": header_text or "",
        "footer_text": footer_text or "",
        "page_numbers_added": add_page_numbers,
    }
