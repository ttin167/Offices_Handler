"""
core/yearbook/handoff_packager.py — Professional Pre-press Designer Handoff Packager.

Assembles and exports a 5-in-1 production handoff package (.zip) for commercial print shops:
1. 01_Digital_Yearbook_Spread_Preview.pdf (Spread View for designer & client)
2. 02_Print_Ready_Pages_300DPI.pdf (Single pages + 3mm Bleed + Crop Marks)
3. 02_Original_Photos_Sorted/ (Original uncompressed photos systematically renamed by page and slot)
4. 03_Designer_Handoff_Brief.md (Markdown specification document)
5. 03_Design_Brief_And_Captions.docx (Word table brief with OpenXML table invariants cantSplit, tblHeader, vAlign)
6. 04_Yearbook_Project_Spec.json (Declarative project spec for 100% round-trip fidelity)
"""

from __future__ import annotations

import json
import os
import shutil
import tempfile
import zipfile
from typing import Any, Dict, List, Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .layout_engine import PRESETS
from .pdf_press_exporter import YearbookPressEngine, generate_mock_photo_bytes


def _apply_table_invariants(table: Any) -> None:
    """
    Applies OpenXML table invariants (cantSplit, tblHeader, vAlign="center")
    to guarantee compliance with ERR_DOCX_001, ERR_DOCX_002, and ERR_DOCX_003.
    """
    # 1. Table header flag on first row
    header_tr = table.rows[0]._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    header_tr.append(tbl_header)

    # 2. cantSplit and vertical centering on all rows
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        trPr.append(cant_split)

        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            v_align = OxmlElement("w:vAlign")
            v_align.set(qn("w:val"), "center")
            tcPr.append(v_align)

            # ERR_DOCX_001: Ensure cell ends with at least one paragraph
            if len(cell.paragraphs) == 0:
                cell.add_paragraph()


class YearbookHandoffPackager:
    """Assembles and zips the designer handoff bundle."""

    @classmethod
    def generate_docx_brief(
        cls,
        yearbook_spec: Dict[str, Any],
        output_docx_path: str,
    ) -> str:
        """
        Generates 03_Design_Brief_And_Captions.docx with OpenXML table invariants.
        """
        doc = Document()

        # Set page margins (1 inch)
        for sec in doc.sections:
            sec.top_margin = Inches(0.8)
            sec.bottom_margin = Inches(0.8)
            sec.left_margin = Inches(0.8)
            sec.right_margin = Inches(0.8)

        preset_key = yearbook_spec.get("preset", "photobook_square_25x25")
        preset = PRESETS.get(preset_key, PRESETS["photobook_square_25x25"])
        title = yearbook_spec.get("title", "KỶ YẾU TỐT NGHIỆP 2026")
        author = yearbook_spec.get("author", "Biện Trung Tín")
        degree = yearbook_spec.get("degree", "Cử Nhân Kỹ Thuật Phần Mềm")
        date_str = yearbook_spec.get("date", "Tháng 11 / 2026")
        bleed_mm = yearbook_spec.get("bleed_mm", 3.0)
        pages = yearbook_spec.get("pages") or yearbook_spec.get("spreads") or []

        # Document Header
        h1 = doc.add_heading(f"BẢN ĐẶC TẢ THIẾT KẾ KỶ YẾU TỐT NGHIỆP", level=1)
        h1.runs[0].font.color.rgb = RGBColor(15, 23, 42)
        h1.runs[0].font.size = Pt(16)

        p_meta = doc.add_paragraph()
        r = p_meta.add_run(f"Dự Án: {title}\n")
        r.bold = True
        p_meta.add_run(f"Chủ Sở Hữu: {author} • {degree}\n")
        p_meta.add_run(f"Khổ Thành Phẩm (Trim Size): {preset['name']} ({preset['w_mm']} x {preset['h_mm']} mm)\n")
        p_meta.add_run(f"Lề Xén Kỹ Thuật (Bleed): {bleed_mm} mm mỗi cạnh | Vùng An Toàn (Safe Zone): 8.0 mm\n")
        p_meta.add_run(f"Tổng Số Trang Đôi: {len(pages)} spreads | Ngày Bàn Giao: {date_str}\n")
        p_meta.paragraph_format.space_after = Pt(14)

        # Instructions for the Designer
        p_note = doc.add_paragraph()
        r_note = p_note.add_run("LƯU Ý DÀNH CHO THỢ THIẾT KẾ & NHÀ IN:\n")
        r_note.bold = True
        r_note.font.color.rgb = RGBColor(30, 58, 138)
        p_note.add_run("1. Toàn bộ ảnh gốc độ phân giải cao đã được đổi tên tương ứng trong thư mục `02_Original_Photos_Sorted/`.\n")
        p_note.add_run("2. Bảng bên dưới quy định rõ ảnh nào thuộc trang nào, lời tựa (caption) và ghi chú retouch đặc thù.\n")
        p_note.add_run("3. Vui lòng tham khảo file `01_Digital_Yearbook_Spread_Preview.pdf` để nắm bắt tinh thần bố cục mong muốn.")
        p_note.paragraph_format.space_after = Pt(16)

        # Specifications Table
        table = doc.add_table(rows=1, cols=5)
        table.style = "Table Grid"

        hdr_cells = table.rows[0].cells
        headers = ["Trang (Spread)", "Mẫu Layout", "Ảnh Gốc Đã Chuẩn Hóa", "Lời Tựa (Caption)", "Ghi Chú Cho Designer"]
        for idx, text in enumerate(headers):
            cell = hdr_cells[idx]
            cell.text = text
            # Format header text
            p = cell.paragraphs[0]
            p.runs[0].bold = True
            p.runs[0].font.size = Pt(9.5)
            p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            # Navy header fill
            tcPr = cell._tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "0F172A")
            tcPr.append(shd)

        for p_idx, p_spec in enumerate(pages, 1):
            row_cells = table.add_row().cells
            layout_type = p_spec.get("layout", p_spec.get("type", "cover_hero"))
            caption = p_spec.get("caption", "—")
            designer_note = p_spec.get("designer_note", "Retouch tone màu tự nhiên, làm mịn da vừa phải.")
            slots = p_spec.get("slots", [])

            photo_names = []
            for s_idx, slot in enumerate(slots, 1):
                clean_name = f"Page_{p_idx:02d}_Slot_{s_idx:02d}.jpg"
                photo_names.append(clean_name)
            sorted_photos_str = "\n".join(photo_names) if photo_names else "—"

            row_cells[0].text = f"Spread {p_idx:02d}\n(Trang {p_idx*2-1}-{p_idx*2})"
            row_cells[1].text = layout_type
            row_cells[2].text = sorted_photos_str
            row_cells[3].text = caption
            row_cells[4].text = designer_note

            # Set small font size
            for c in row_cells:
                for p in c.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8.5)

        # Apply OpenXML table invariants
        _apply_table_invariants(table)

        os.makedirs(os.path.dirname(os.path.abspath(output_docx_path)), exist_ok=True)
        doc.save(output_docx_path)
        return output_docx_path

    @classmethod
    def export_handoff_bundle(
        cls,
        yearbook_spec: Dict[str, Any],
        output_zip_path: str,
    ) -> str:
        """
        Assembles and outputs the complete 5-in-1 ZIP handoff package:
        - Streams files directly to disk to maintain RAM < 512MB (ERR_RAM_001).
        - Generates 01_Spread_Preview.pdf, 02_Print_Ready.pdf, Sorted Photos, Word brief, and JSON spec.
        """
        preset_key = yearbook_spec.get("preset", "photobook_square_25x25")
        preset = PRESETS.get(preset_key, PRESETS["photobook_square_25x25"])
        title = yearbook_spec.get("title", "KỶ YẾU TỐT NGHIỆP 2026")
        author = yearbook_spec.get("author", "Biện Trung Tín")
        degree = yearbook_spec.get("degree", "Cử Nhân Kỹ Thuật Phần Mềm")
        date_str = yearbook_spec.get("date", "Tháng 11 / 2026")
        bleed_mm = yearbook_spec.get("bleed_mm", 3.0)
        pages = yearbook_spec.get("pages") or yearbook_spec.get("spreads") or []

        temp_dir = tempfile.mkdtemp(prefix="yearbook_handoff_")
        try:
            # 1. Render Spread Preview PDF
            spread_pdf_path = os.path.join(temp_dir, "01_Digital_Yearbook_Spread_Preview.pdf")
            YearbookPressEngine.export_spread_preview_pdf(yearbook_spec, spread_pdf_path)

            # 2. Render Print Ready PDF (Bleed 3mm + Crop Marks)
            print_pdf_path = os.path.join(temp_dir, "02_Print_Ready_Pages_300DPI.pdf")
            YearbookPressEngine.export_print_ready_pdf(yearbook_spec, print_pdf_path)

            # 3. Generate Word Brief DOCX
            docx_brief_path = os.path.join(temp_dir, "03_Design_Brief_And_Captions.docx")
            cls.generate_docx_brief(yearbook_spec, docx_brief_path)

            # 4. Generate Markdown Brief
            md_lines = [
                f"# BẢN ĐẶC TẢ THIẾT KẾ KỶ YẾU (DESIGNER HANDOFF BRIEF)",
                f"- **Dự án**: {title}",
                f"- **Chủ sở hữu**: {author} • {degree}",
                f"- **Khổ thành phẩm (Trim Size)**: {preset['name']} ({preset['w_mm']}x{preset['h_mm']} mm)",
                f"- **Lề xén (Bleed)**: {bleed_mm} mm mỗi cạnh | **Vùng an toàn (Safe Zone)**: 8.0 mm",
                f"- **Ngày bàn giao**: {date_str}\n",
                "## CHI TIẾT CÁC TRANG & YÊU CẦU THIẾT KẾ:\n",
            ]

            for p_idx, p_spec in enumerate(pages, 1):
                layout_type = p_spec.get("layout", p_spec.get("type", "cover_hero"))
                caption = p_spec.get("caption", "—")
                designer_note = p_spec.get("designer_note", "Giữ tone màu tự nhiên, làm mịn da vừa phải.")
                slots = p_spec.get("slots", [])

                md_lines.append(f"### Spread {p_idx:02d} (Trang {p_idx*2-1}-{p_idx*2}) — Layout: `{layout_type}`")
                md_lines.append(f"- **Lời tựa in trên trang (Caption)**: \"{caption}\"")
                md_lines.append(f"- **Ghi chú cho Designer**: {designer_note}")
                md_lines.append("- **Danh sách ảnh trong khung:**")
                for s_idx, slot in enumerate(slots, 1):
                    clean_name = f"Page_{p_idx:02d}_Slot_{s_idx:02d}.jpg"
                    md_lines.append(f"  + Khung {s_idx:02d}: `02_Original_Photos_Sorted/{clean_name}` ({slot.get('caption', 'Ảnh')})")
                md_lines.append("")

            md_brief_path = os.path.join(temp_dir, "03_Designer_Handoff_Brief.md")
            with open(md_brief_path, "w", encoding="utf-8") as f:
                f.write("\n".join(md_lines))

            # 5. Project Spec JSON
            spec_json_path = os.path.join(temp_dir, "04_Yearbook_Project_Spec.json")
            with open(spec_json_path, "w", encoding="utf-8") as f:
                json.dump(yearbook_spec, f, ensure_ascii=False, indent=2)

            # 6. Stream all components into final ZIP archive (ERR_RAM_001)
            os.makedirs(os.path.dirname(os.path.abspath(output_zip_path)), exist_ok=True)
            with zipfile.ZipFile(output_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                # Add generated documents
                zf.write(spread_pdf_path, "01_Digital_Yearbook_Spread_Preview.pdf")
                zf.write(print_pdf_path, "02_Print_Ready_Pages_300DPI.pdf")
                zf.write(docx_brief_path, "03_Design_Brief_And_Captions.docx")
                zf.write(md_brief_path, "03_Designer_Handoff_Brief.md")
                zf.write(spec_json_path, "04_Yearbook_Project_Spec.json")

                # Add sorted original photos
                for p_idx, p_spec in enumerate(pages, 1):
                    slots = p_spec.get("slots", [])
                    for s_idx, slot in enumerate(slots, 1):
                        photo_path = slot.get("photo_path", "")
                        ext = os.path.splitext(photo_path)[1].lower() if photo_path else ".jpg"
                        if not ext:
                            ext = ".jpg"
                        clean_img_name = f"Page_{p_idx:02d}_Slot_{s_idx:02d}{ext}"
                        arc_target = f"02_Original_Photos_Sorted/{clean_img_name}"

                        if photo_path and os.path.isfile(photo_path):
                            # Stream directly from disk (zero RAM load)
                            zf.write(photo_path, arc_target)
                        else:
                            # Write mock photo placeholder
                            mock_bytes = generate_mock_photo_bytes(800, 800, slot.get("mock_theme", "gown"), clean_img_name)
                            zf.writestr(arc_target, mock_bytes)

            return output_zip_path
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
