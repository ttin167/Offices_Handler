# -*- coding: utf-8 -*-
"""
tests/test_docx_mastery.py — Unit Tests for Pillar 2: Vietnamese Words Converter,
Mail Merge Formatting Specs, Decree 30/2020 Polish, Merge, Split, and Watermark/HF.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches
from PIL import Image

from core.docx.design_system import format_token_value, number_to_vietnamese_words
from core.docx.mail_merge import execute_mail_merge, inspect_template_tokens
from core.docx.toolkit import (
    apply_docx_watermark_and_hf,
    merge_docx_files,
    polish_docx_document,
    split_docx_by_heading,
)


def test_number_to_vietnamese_words_and_format_token_specs() -> None:
    assert number_to_vietnamese_words(0) == "Không đồng"
    assert number_to_vietnamese_words(15_000_000) == "Mười lăm triệu đồng"
    assert number_to_vietnamese_words(1_005_000) == "Một triệu không trăm lẻ năm nghìn đồng"
    assert number_to_vietnamese_words(1_000_000_000) == "Một tỷ đồng"
    assert number_to_vietnamese_words(-500_000) == "Âm năm trăm nghìn đồng"
    assert number_to_vietnamese_words(2_050_000_100) == "Hai tỷ không trăm năm mươi triệu một trăm đồng"

    assert format_token_value(15_000_000, "vn_words") == "Mười lăm triệu đồng"
    assert format_token_value(15_000_000, "words_plain") == "Mười lăm triệu"
    assert format_token_value("2026-11-14", "date_vn") == "14 tháng 11 năm 2026"
    assert format_token_value("nguyễn văn an", "title") == "Nguyễn Văn An"


def test_mail_merge_with_vn_words_and_header_footer_tables(tmp_path: Path) -> None:
    tpl_path = tmp_path / "hop_dong_tpl.docx"
    out_dir = tmp_path / "merged_out"

    doc = Document()
    doc.add_paragraph("Kính gửi ông/bà {{HoTen:title}}, số tiền {{SoTien:vnd}} ({{SoTien:vn_words}}).")
    doc.add_paragraph("Ngày ký hợp đồng: {{NgayKy:date_vn}}.")
    hdr_tbl = doc.sections[0].header.add_table(rows=1, cols=1, width=Inches(5))
    hdr_tbl.cell(0, 0).text = "Mã HĐ: {{MaHD:upper}}"
    ftr_tbl = doc.sections[0].footer.add_table(rows=1, cols=1, width=Inches(5))
    ftr_tbl.cell(0, 0).text = "Người nhận: {{HoTen:lower}}"
    doc.save(str(tpl_path))

    inspected = inspect_template_tokens(str(tpl_path))
    assert inspected["template_file"] == "hop_dong_tpl.docx"
    assert set(["HoTen", "SoTien", "NgayKy", "MaHD"]).issubset(set(inspected["variables"]))
    assert len(inspected.get("formatted_tokens", [])) >= 4

    merge_res = execute_mail_merge(
        template_path=str(tpl_path),
        data_source=[
            {
                "HoTen": "trần thị mai",
                "SoTien": 15000000,
                "NgayKy": "2026-11-14",
                "MaHD": "hd-2026-09",
            }
        ],
        output_dir=str(out_dir),
        filename_pattern="HopDong_{{MaHD}}.docx",
    )
    assert merge_res["success"] is True
    assert merge_res["total_records"] == 1
    generated = list(out_dir.glob("*.docx"))
    assert len(generated) == 1

    merged_doc = Document(str(generated[0]))
    body_text = "\n".join(p.text for p in merged_doc.paragraphs)
    assert "Trần Thị Mai" in body_text
    assert "Mười lăm triệu đồng" in body_text
    assert "14 tháng 11 năm 2026" in body_text
    hdr_text = merged_doc.sections[0].header.tables[0].cell(0, 0).text
    ftr_text = merged_doc.sections[0].footer.tables[0].cell(0, 0).text
    assert "HD-2026-09" in hdr_text
    assert "trần thị mai" in ftr_text


def test_polish_merge_split_and_watermark_hf(tmp_path: Path) -> None:
    img_path = tmp_path / "wide_banner.png"
    Image.new("RGB", (800, 200), color=(30, 80, 160)).save(str(img_path))

    doc1_path = tmp_path / "ch1.docx"
    d1 = Document()
    d1.add_heading("Chương 1: Quy định chung", level=1)
    d1.add_paragraph("Nội dung chuẩn hóa thể thức Nghị định 30/2020/NĐ-CP .")
    tbl = d1.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = "STT"
    tbl.cell(0, 1).text = "Hạng mục"
    tbl.cell(1, 0).text = "1"
    tbl.cell(1, 1).text = "Kiểm tra bảng"
    d1.add_picture(str(img_path), width=Inches(8.5))  # Deliberately oversized (> 14cm)
    d1.save(str(doc1_path))

    polished_path = tmp_path / "ch1_polished.docx"
    res_pol = polish_docx_document(
        input_path=str(doc1_path),
        output_path=str(polished_path),
        preset="nd30",
        add_page_numbers=True,
        sanitize_typography=True,
        fix_tables=True,
        admin_header={"agency_parent": "UBND THÀNH PHỐ", "agency_name": "VĂN PHÒNG"},
    )
    assert res_pol["ok"] is True
    assert polished_path.is_file()
    assert res_pol["tables_fixed"] >= 1
    assert res_pol["images_clamped"] >= 1

    doc2_path = tmp_path / "ch2.docx"
    d2 = Document()
    d2.add_heading("Chương 2: Triển khai thực hiện", level=1)
    d2.add_paragraph("Tài liệu đính kèm hình ảnh minh họa.")
    d2.add_picture(str(img_path), width=Inches(3.0))
    d2.save(str(doc2_path))

    merged_path = tmp_path / "master_merged.docx"
    res_merge = merge_docx_files([str(doc1_path), str(doc2_path)], output_path=str(merged_path))
    assert res_merge["ok"] is True
    assert res_merge["merged_files"] == 2
    assert merged_path.is_file()

    m_doc = Document(str(merged_path))
    assert len(m_doc.inline_shapes) >= 2

    split_zip = tmp_path / "chapters.zip"
    res_split = split_docx_by_heading(str(merged_path), output_path=str(split_zip), heading_level=1, as_zip=True)
    assert res_split["ok"] is True
    assert res_split["total_chapters"] >= 2
    with zipfile.ZipFile(str(split_zip), "r") as zf:
        assert len([n for n in zf.namelist() if n.endswith(".docx")]) >= 2

    branded_path = tmp_path / "branded.docx"
    res_wm = apply_docx_watermark_and_hf(
        input_path=str(doc2_path),
        output_path=str(branded_path),
        watermark_text="TÀI LIỆU NỘI BỘ",
        header_text="CÔNG TY ANTIGRAVITY",
        footer_text="Lưu hành nội bộ",
        add_page_numbers=True,
    )
    assert res_wm["ok"] is True
    assert branded_path.is_file()
    with zipfile.ZipFile(str(branded_path), "r") as zf:
        header_xml = zf.read("word/header1.xml").decode("utf-8", errors="ignore")
        assert "PowerPlusWaterMarkObject" in header_xml
        assert "CÔNG TY ANTIGRAVITY" in header_xml
