"""
tests/test_docx_design_system.py — Comprehensive Unit & Integration Tests
for Centralized Enterprise Document Design System (core.docx.design_system).

Verifies:
1. Multi-Theme Registry (modern_tech, executive_classic, nghidinh30).
2. Executive Cover Page Builder & Metadata Card.
3. Native Dynamic Table of Contents (Word F9 TOC Field).
4. Dual Header/Footer with 'Trang X / Y' Native Fields.
5. High-Impact Callout Cards (NOTE, TIP, WARNING, IMPORTANT).
6. Fenced Code Block Cards with Consolas Typography.
7. Modern Clean Studio Table Engine & Invariants (cantSplit, tblHeader, tcMar, tcBorders).
8. Dynamic Template Token Formatters (Currency, Dates, Number formatting).
9. Integration with decoupled_to_docx and DocxDiagnosticEngine verification (0 issues).
"""

from __future__ import annotations

import os
import tempfile
import docx
from docx import Document
from docx.oxml.ns import qn
import pytest

from core.docx.design_system import (
    DocxTheme,
    get_theme,
    THEME_MODERN_TECH,
    THEME_EXECUTIVE_CLASSIC,
    THEME_NGHIDINH30,
    build_cover_page,
    inject_native_toc,
    setup_header_footer,
    build_callout_box,
    build_code_card,
    format_table_modern_clean,
    format_token_value,
    detect_column_alignment,
)
from core.docx.writer import _replace_text_in_paragraph
from core.converter.markdown import decoupled_to_docx
from tools.unified_qa_diagnostic import DocxDiagnosticEngine


class TestDocxThemes:
    def test_theme_registry_retrieval(self):
        th_default = get_theme()
        assert th_default.name == "modern_tech"
        assert th_default.font_primary == "Segoe UI"
        assert th_default.color_primary == "0F172A"

        th_exec = get_theme("executive_classic")
        assert th_exec.name == "executive_classic"
        assert th_exec.font_primary == "Times New Roman"
        assert th_exec.color_primary == "0F294A"

        th_nd30 = get_theme("nghidinh30")
        assert th_nd30.name == "nghidinh30"
        assert th_nd30.font_primary == "Times New Roman"

        # Fallback for unknown theme name
        th_fallback = get_theme("unknown_custom_theme")
        assert th_fallback.name == "modern_tech"

    def test_theme_hex_to_rgb(self):
        th = THEME_MODERN_TECH
        rgb = th.hex_to_rgb("2563EB")
        assert rgb[0] == 0x25
        assert rgb[1] == 0x63
        assert rgb[2] == 0xEB

        rgb_default = th.hex_to_rgb("invalid")
        assert rgb_default[0] == 15
        assert rgb_default[1] == 23
        assert rgb_default[2] == 42


class TestExecutiveCoverPage:
    def test_build_cover_page_structure(self):
        doc = Document()
        build_cover_page(
            doc,
            title="Hệ Thống Thanh Toán Số",
            subtitle="Đặc Tả Kỹ Thuật & Kiến Trúc Microservices",
            author="Ban Chuyển Đổi Số",
            org="TẬP ĐOÀN TÀI CHÍNH CÔNG NGHỆ",
            version_str="v2.5.0",
            theme=THEME_MODERN_TECH,
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "cover_page.docx")
            doc.save(out_path)
            assert os.path.exists(out_path)

            doc_read = Document(out_path)
            full_text = "\n".join([p.text for p in doc_read.paragraphs])
            assert "HỆ THỐNG THANH TOÁN SỐ" in full_text
            assert "Đặc Tả Kỹ Thuật & Kiến Trúc Microservices" in full_text
            assert "TẬP ĐOÀN TÀI CHÍNH CÔNG NGHỆ" in full_text

            # Check metadata card table
            assert len(doc_read.tables) >= 1
            card_table = doc_read.tables[0]
            card_text = card_table.rows[0].cells[0].text
            assert "Ban Chuyển Đổi Số" in card_text
            assert "v2.5.0" in card_text

            # Check OpenXML invariant: cantSplit on card row
            trPr = card_table.rows[0]._tr.get_or_add_trPr()
            assert trPr.find(qn("w:cantSplit")) is not None


class TestNativeTOC:
    def test_inject_native_toc(self):
        doc = Document()
        doc.add_heading("Tài Liệu Mẫu", level=1)
        inject_native_toc(doc, title="MỤC LỤC TỔNG QUAN", max_level=3)

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "native_toc.docx")
            doc.save(out_path)

            doc_read = Document(out_path)
            full_text = "\n".join([p.text for p in doc_read.paragraphs])
            assert "MỤC LỤC TỔNG QUAN" in full_text

            # Check that fldSimple with TOC exists in the XML
            xml_str = doc_read._element.xml
            assert 'w:instr="TOC \\o &quot;1-3&quot; \\h \\z \\u"' in xml_str


class TestDualHeaderFooter:
    def test_setup_header_footer(self):
        doc = Document()
        doc.add_paragraph("Nội dung trang 1")
        setup_header_footer(doc, header_title="BÁO CÁO TỔNG KẾT DỰ ÁN", confidentiality="NỘI BỘ")

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "header_footer.docx")
            doc.save(out_path)

            doc_read = Document(out_path)
            sec = doc_read.sections[0]

            # Header text check
            header_text = "\n".join([p.text for p in sec.header.paragraphs])
            assert "BÁO CÁO TỔNG KẾT DỰ ÁN" in header_text

            # Footer XML check for PAGE and NUMPAGES
            footer_xml = sec.footer._element.xml
            assert 'w:instr="PAGE"' in footer_xml
            assert 'w:instr="NUMPAGES"' in footer_xml
            assert "Trang " in footer_xml
            assert "NỘI BỘ" in footer_xml


class TestCalloutBoxes:
    @pytest.mark.parametrize("callout_type", ["NOTE", "TIP", "WARNING", "IMPORTANT", "CAUTION"])
    def test_build_callout_box_types(self, callout_type):
        doc = Document()
        build_callout_box(
            doc,
            callout_type=callout_type,
            content=["Dòng 1: Chi tiết lưu ý quan trọng.", "Dòng 2: Hướng dẫn xử lý."],
            theme=THEME_MODERN_TECH,
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, f"callout_{callout_type}.docx")
            doc.save(out_path)

            doc_read = Document(out_path)
            assert len(doc_read.tables) == 1
            tbl = doc_read.tables[0]
            cell_text = tbl.rows[0].cells[0].text
            assert "Dòng 1: Chi tiết lưu ý quan trọng." in cell_text
            assert "Dòng 2: Hướng dẫn xử lý." in cell_text

            # Verify table OpenXML invariants
            trPr = tbl.rows[0]._tr.get_or_add_trPr()
            assert trPr.find(qn("w:cantSplit")) is not None
            tcPr = tbl.rows[0].cells[0]._tc.get_or_add_tcPr()
            assert tcPr.find(qn("w:tcBorders")) is not None
            assert tcPr.find(qn("w:shd")) is not None

    def test_callout_inline_markdown_parsing(self):
        doc = Document()
        lines = [
            "**Tài liệu tra cứu chuẩn hóa cho Ban Lãnh Đạo và Chuyên Viên**",
            "**Cập nhật chính sách**: Luật BHXH 2024 có hiệu lực từ 01/07/2025.",
            "**Đơn vị ban hành**: Bộ Phận Phân Tích & Quản Trị Rủi Ro Nghiệp Vụ.",
            "**Căn cứ pháp lý**: `Điều 20, 24-27` Bộ luật Lao động 2019.",
        ]
        build_callout_box(
            doc,
            callout_type="INFO",
            content=lines,
            title="GHI CHÚ / TRÍCH DẪN",
            theme=THEME_MODERN_TECH,
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "callout_markdown.docx")
            doc.save(out_path)

            doc_read = Document(out_path)
            tbl = doc_read.tables[0]
            cell = tbl.rows[0].cells[0]

            # 1. Ensure NO raw '**' artifacts remain in the text
            for p in cell.paragraphs:
                assert "**" not in p.text, f"Unparsed markdown '**' found: {p.text}"

            # 2. Ensure bold runs are present for headings and labels
            bold_runs = [r for p in cell.paragraphs for r in p.runs if r.bold]
            assert len(bold_runs) >= 4, "Expected at least 4 bold runs for markdown headers/labels"

            # 3. Ensure inline code was converted to Consolas
            code_runs = [r for p in cell.paragraphs for r in p.runs if r.font.name == "Consolas"]
            assert len(code_runs) >= 1
            assert "Điều 20, 24-27" in code_runs[0].text


class TestCodeBlockCards:
    def test_build_code_card(self):
        doc = Document()
        code = "def process_data(items):\n    return [x * 2 for x in items]"
        build_code_card(doc, code=code, lang="python", theme=THEME_MODERN_TECH)

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "code_card.docx")
            doc.save(out_path)

            doc_read = Document(out_path)
            assert len(doc_read.tables) == 1
            tbl = doc_read.tables[0]
            cell_text = tbl.rows[0].cells[0].text
            assert "PYTHON" in cell_text
            assert "def process_data(items):" in cell_text


class TestModernCleanTables:
    def test_format_table_modern_clean(self):
        doc = Document()
        table = doc.add_table(rows=3, cols=3)
        headers = ["Mã SV", "Họ và Tên", "Điểm TB"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h

        data = [
            ["SV001", "Nguyễn Văn A", "8.5"],
            ["SV002", "Trần Thị B", "9.2"],
        ]
        for r_i, row_data in enumerate(data, start=1):
            for c_i, val in enumerate(row_data):
                table.rows[r_i].cells[c_i].text = val

        format_table_modern_clean(table, theme=THEME_MODERN_TECH)

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "clean_table.docx")
            doc.save(out_path)

            doc_read = Document(out_path)
            tbl = doc_read.tables[0]

            # Invariant checks: tblHeader on row 0, cantSplit on all rows
            assert tbl.rows[0]._tr.get_or_add_trPr().find(qn("w:tblHeader")) is not None
            for row in tbl.rows:
                assert row._tr.get_or_add_trPr().find(qn("w:cantSplit")) is not None

            # Cell padding & borders checks
            for row in tbl.rows:
                for cell in row.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    assert tcPr.find(qn("w:tcMar")) is not None
                    assert tcPr.find(qn("w:tcBorders")) is not None

    def test_detect_column_alignment(self):
        numeric_col = ["100", "2,500.50", "30%", "45.000 ₫"]
        assert detect_column_alignment(numeric_col) == docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT

        id_col = ["US-001", "UC-02", "ERR-404"]
        assert detect_column_alignment(id_col) == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

        text_col = ["Người dùng đăng nhập thành công vào ứng dụng", "Gửi email thông báo"]
        assert detect_column_alignment(text_col) == docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

    def test_markdown_table_header_contrast_white(self):
        md_table_doc = """# TỶ LỆ ĐÓNG BẢO HIỂM BẮT BUỘC

| Quỹ Thành Phần | Doanh Nghiệp (ER) | Người Lao Động (EE) | Tổng Cộng | Ghi Chú |
| :--- | :---: | :---: | :---: | :--- |
| BHXH (Hưu trí & Tử tuất) | 14.0% | 8.0% | 22.0% | Bắt buộc |
| BHYT | 3.0% | 1.5% | 4.5% | Khám chữa bệnh |
"""
        with tempfile.TemporaryDirectory() as tmpdir:
            md_path = os.path.join(tmpdir, "table_contrast.md")
            docx_path = os.path.join(tmpdir, "table_contrast.docx")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_table_doc)

            ok, out_docx = decoupled_to_docx(md_path, docx_path=docx_path)
            assert ok
            doc_read = Document(out_docx)
            assert len(doc_read.tables) >= 1

            tbl = doc_read.tables[0]
            header_row = tbl.rows[0]

            # Verify that all header cell text is Pure White (255, 255, 255) and bold
            for cell in header_row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        assert r.bold is True, f"Header run '{r.text}' should be bold"
                        assert r.font.color.rgb == docx.shared.RGBColor(255, 255, 255), (
                            f"Header run '{r.text}' should have pure white font color, got {r.font.color.rgb}"
                        )


class TestTokenFormatters:
    def test_format_token_value_currency(self):
        assert format_token_value("15000000", "VND") == "15.000.000 ₫"
        assert format_token_value("1250.75", "USD") == "$1,250.75"

    def test_format_token_value_date(self):
        assert format_token_value("2026-11-06", "DATE") == "06/11/2026"
        assert format_token_value("2026-05-15", "DATE_VN") == "15 tháng 05 năm 2026"

    def test_format_token_value_num(self):
        assert format_token_value("1234567.89", "NUM") == "1,234,567.89"
        assert format_token_value("1234567.89", "INT") == "1,234,568"

    def test_paragraph_replace_with_formatters(self):
        doc = Document()
        p = doc.add_paragraph("Lương: {{SALARY:VND}} | Ngày bắt đầu: {{START_DATE:DATE}}")
        replacements = {
            "SALARY": "25000000",
            "START_DATE": "2026-11-06",
        }
        replaced = _replace_text_in_paragraph(p, replacements)
        assert replaced >= 1
        assert "25.000.000 ₫" in p.text
        assert "06/11/2026" in p.text


class TestMarkdownPipelineIntegration:
    def test_decoupled_to_docx_with_design_system_and_diagnose(self):
        md_content = """# BÁO CÁO NGHIỆP VỤ CAO CẤP

> [!NOTE]
> Đây là hộp thông báo lưu ý theo phong cách GitHub Alert hiện đại.

> [!WARNING]
> Cảnh báo về vấn đề bảo mật hệ thống thanh toán.

## 1. Dữ Liệu Nhân Sự

| Mã NV | Họ và Tên | Chức Vụ | Mức Lương |
| :---: | :--- | :--- | ---: |
| NV01 | Nguyễn Văn An | Kỹ sư trưởng | 45.000.000 ₫ |
| NV02 | Trần Thị Bình | Trưởng phòng BA | 42.000.000 ₫ |
| NV03 | Lê Hoàng Cường | Chuyên viên Dev | 30.000.000 ₫ |

## 2. Mã Nguồn Minh Họa

```python
def authenticate(token: str) -> bool:
    return len(token) > 32
```
"""
        with tempfile.TemporaryDirectory() as tmpdir:
            md_file = os.path.join(tmpdir, "report.md")
            docx_file = os.path.join(tmpdir, "report.docx")
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(md_content)

            ok, out_docx = decoupled_to_docx(md_file, docx_path=docx_file)
            assert ok
            assert os.path.exists(out_docx)

            # Diagnose OpenXML quality via DocxDiagnosticEngine
            issues, summary = DocxDiagnosticEngine.diagnose(out_docx)
            assert len(issues) == 0, f"Diagnostic issues found: {issues}"
            assert summary["paragraphs_count"] > 0
            assert summary["tables_count"] > 0
