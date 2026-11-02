"""
core.ba_hub.brd_generator — Enterprise BRD & SRS Document Synthesizer for Business Analysts.

Assembles complete, C-level presentation-grade Word (.docx) specifications according to:
- IEEE 830 / ISO 29148 Requirements Engineering Standards.
- BABOK Guide v3 Best Practices.
- OpenXML Strict Invariants (cantSplit, tblHeader, vAlign).
"""

from __future__ import annotations

import datetime
import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor


@dataclass
class BRDMetadata:
    title: str = "TÀI LIỆU ĐẶC TẢ YÊU CẦU NGHIỆP VỤ (BRD)"
    project_name: str = "Hệ Thống Số Hóa Quy Trình & Thanh Toán Doanh Nghiệp"
    version: str = "1.0"
    author: str = "Lead Business Analyst"
    organization: str = "Khối Chuyển Đổi Số & Công Nghệ Thông Tin"
    confidentiality: str = "Nội bộ / Confidential"
    approver: str = "Head of Product / Project Director"
    created_date: str = field(default_factory=lambda: datetime.date.today().strftime("%d/%m/%Y"))


class BRDDocumentGenerator:
    """Assembles and generates publication-grade Word DOCX BRD/SRS documents."""

    @classmethod
    def generate_brd(
        cls,
        meta: Any,
        stories: Optional[List[Any]] = None,
        usecases: Optional[List[Any]] = None,
        data_fields: Optional[List[Dict[str, Any]]] = None,
        scope_in: Optional[List[str]] = None,
        scope_out: Optional[List[str]] = None,
        output_path: str = "dist/BRD_Specification.docx",
    ) -> str:
        if isinstance(usecases, str):
            output_path = usecases
            usecases = None

        if isinstance(meta, dict):
            meta = BRDMetadata(
                title=meta.get("title", "TÀI LIỆU YÊU CẦU NGHIỆP VỤ (BRD)"),
                project_name=meta.get("project_name", "Dự Án Nghiệp Vụ"),
                version=meta.get("version", "1.0"),
                author=meta.get("author", "Business Analyst"),
                organization=meta.get("organization", "Doanh Nghiệp"),
                created_date=meta.get("created_date", meta.get("date_created", "2026")),
                confidentiality=meta.get("confidentiality", "Nội bộ / Confidential"),
            )

        doc = Document()

        # Page setup A4
        for sec in doc.sections:
            sec.top_margin = Inches(1.0)
            sec.bottom_margin = Inches(1.0)
            sec.left_margin = Inches(1.0)
            sec.right_margin = Inches(0.8)

        # ----------------------------------------------------------------------
        # 1. ENTERPRISE COVER PAGE
        # ----------------------------------------------------------------------
        p_org = doc.add_paragraph()
        r_org = p_org.add_run(meta.organization.upper())
        r_org.font.size = Pt(11)
        r_org.bold = True
        r_org.font.color.rgb = RGBColor(100, 116, 139)

        doc.add_paragraph().paragraph_format.space_after = Pt(72)

        p_t = doc.add_paragraph()
        r_t = p_t.add_run(meta.title)
        r_t.font.size = Pt(24)
        r_t.bold = True
        r_t.font.color.rgb = RGBColor(15, 23, 42)

        p_p = doc.add_paragraph()
        r_p = p_p.add_run(f"Dự Án: {meta.project_name}")
        r_p.font.size = Pt(15)
        r_p.bold = True
        r_p.font.color.rgb = RGBColor(2, 132, 199)

        p_conf = doc.add_paragraph()
        r_conf = p_conf.add_run(f"Cấp độ bảo mật: {meta.confidentiality}")
        r_conf.italic = True
        r_conf.font.size = Pt(10)
        r_conf.font.color.rgb = RGBColor(148, 163, 184)

        doc.add_paragraph().paragraph_format.space_after = Pt(120)

        # Meta Table on Cover
        t_meta = doc.add_table(rows=4, cols=2)
        t_meta.alignment = WD_TABLE_ALIGNMENT.LEFT
        meta_items = [
            ("Tác giả lập tài liệu:", meta.author),
            ("Người phê duyệt:", meta.approver),
            ("Phiên bản tài liệu:", f"v{meta.version} (Official Release)"),
            ("Ngày hoàn tất:", meta.created_date),
        ]
        for idx, (k, v) in enumerate(meta_items):
            row = t_meta.rows[idx]
            c0, c1 = row.cells[0], row.cells[1]
            c0.width, c1.width = Inches(2.2), Inches(4.0)
            c0.paragraphs[0].add_run(k).bold = True
            c1.paragraphs[0].add_run(v)
            for cell in (c0, c1):
                cell.paragraphs[0].runs[0].font.size = Pt(9.5)

        # Page break after cover
        doc.add_page_break()

        # ----------------------------------------------------------------------
        # 2. DOCUMENT CONTROL & REVISION HISTORY
        # ----------------------------------------------------------------------
        h1 = doc.add_heading("QUẢN TRỊ TÀI LIỆU & LỊCH SỬ PHIÊN BẢN (DOCUMENT CONTROL)", level=1)
        h1.runs[0].font.size = Pt(13)
        h1.runs[0].font.color.rgb = RGBColor(15, 23, 42)

        t_rev = doc.add_table(rows=1, cols=5)
        t_rev.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_rev = t_rev.rows[0]
        hdr_rev._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        hdr_rev._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        rev_cols = ["Phiên bản", "Ngày cập nhật", "Tác giả", "Mô tả thay đổi nghiệp vụ", "Trạng thái"]
        rev_widths = [Inches(0.9), Inches(1.2), Inches(1.4), Inches(2.3), Inches(1.0)]

        for idx, text in enumerate(rev_cols):
            c = hdr_rev.cells[idx]
            c.width = rev_widths[idx]
            c_pr = c._tc.get_or_add_tcPr()
            c_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1E293B"/>'))
            c_pr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text)
            r.bold = True
            r.font.size = Pt(9)
            r.font.color.rgb = RGBColor(255, 255, 255)

        rev_rows_data = [
            ("v1.0", meta.created_date, meta.author, "Khởi tạo tài liệu đặc tả yêu cầu nghiệp vụ ban đầu (Initial Draft).", "Đã duyệt"),
            ("v1.1", meta.created_date, meta.author, "Bổ sung tiêu chí nghiệm thu Gherkin và bảng từ điển dữ liệu tích hợp.", "Hiện tại"),
        ]
        for row_data in rev_rows_data:
            r_obj = t_rev.add_row()
            r_obj._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            for c_idx, val in enumerate(row_data):
                cell = r_obj.cells[c_idx]
                cell.width = rev_widths[c_idx]
                cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CBD5E1"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="CBD5E1"/><w:right w:val="single" w:sz="4" w:color="CBD5E1"/></w:tcBorders>'))
                p = cell.paragraphs[0]
                if c_idx in (0, 1, 4):
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(val).font.size = Pt(9)

        doc.add_paragraph().paragraph_format.space_after = Pt(16)

        # ----------------------------------------------------------------------
        # 3. EXECUTIVE SUMMARY & SCOPE
        # ----------------------------------------------------------------------
        doc.add_heading("1. TỔNG QUAN DỰ ÁN & MỤC TIÊU NGHIỆP VỤ", level=1)
        p_exec = doc.add_paragraph()
        p_exec.add_run(
            f"Dự án '{meta.project_name}' được triển khai nhằm mục đích tự động hóa, tối ưu hóa quy trình vận hành và nâng cao trải nghiệm người dùng cuối. "
            "Tài liệu này cung cấp nền tảng pháp lý và kỹ thuật duy nhất (Single Source of Truth) phục vụ đội ngũ phát triển, kiểm thử chất lượng (QA) và các bên liên quan nghiệm thu."
        )

        doc.add_heading("2. PHẠM VI DỰ ÁN (PROJECT SCOPE)", level=1)
        p_in_hdr = doc.add_paragraph()
        p_in_hdr.add_run("2.1. Nằm trong phạm vi triển khai (In-Scope):").bold = True
        default_in = scope_in or [
            "Chuẩn hóa quy trình giao dịch trực tuyến và tích hợp cổng thanh toán.",
            "Xây dựng API đồng bộ dữ liệu thời gian thực giữa hệ thống lõi và hệ thống vệ tinh.",
            "Bộ quy tắc đối soát số liệu tự động và phát hiện chênh lệch.",
            "Cung cấp báo cáo phân tích quản trị cho ban giám đốc."
        ]
        for s in default_in:
            doc.add_paragraph(s, style='List Bullet')

        p_out_hdr = doc.add_paragraph()
        p_out_hdr.add_run("2.2. Nằm ngoài phạm vi giai đoạn này (Out-of-Scope):").bold = True
        default_out = scope_out or [
            "Tích hợp cổng thanh toán quốc tế bằng tiền điện tử / Crypto.",
            "Nâng cấp phần cứng máy chủ vật lý on-premise.",
            "Hỗ trợ giao dịch ngoại hối phức tạp phái sinh."
        ]
        for s in default_out:
            doc.add_paragraph(s, style='List Bullet')

        # ----------------------------------------------------------------------
        # 4. USER STORIES & ACCEPTANCE CRITERIA
        # ----------------------------------------------------------------------
        doc.add_heading("3. YÊU CẦU CHỨC NĂNG (FUNCTIONAL REQUIREMENTS - USER STORIES)", level=1)
        if stories:
            for raw_st in stories:
                st = raw_st.to_dict() if hasattr(raw_st, "to_dict") else raw_st
                p_st_title = doc.add_paragraph()
                r_st_id = p_st_title.add_run(f"[{st.get('id', 'US-001')}] {st.get('title', '')}")
                r_st_id.bold = True
                r_st_id.font.size = Pt(11)

                p_st_body = doc.add_paragraph()
                p_st_body.paragraph_format.left_indent = Inches(0.25)
                p_st_body.add_run("As a ").bold = True
                p_st_body.add_run(f"{st.get('role', 'User')}, ")
                p_st_body.add_run("I want ").bold = True
                p_st_body.add_run(f"{st.get('goal', '')}, ")
                p_st_body.add_run("So that ").bold = True
                p_st_body.add_run(f"{st.get('benefit', '')}.\n")

                scenarios = st.get("scenarios", [])
                if scenarios:
                    p_st_body.add_run("Tiêu chí nghiệm thu (Acceptance Criteria):\n").bold = True
                    for sc in scenarios:
                        p_st_body.add_run(f"  • [{sc.get('type', 'Scenario')}]: {sc.get('title', '')}\n")
                        p_st_body.add_run(f"    - Given: {sc.get('given', '')}\n")
                        p_st_body.add_run(f"    - When: {sc.get('when', '')}\n")
                        p_st_body.add_run(f"    - Then: {sc.get('then', '')}\n")
        else:
            doc.add_paragraph("Chi tiết các User Story được đính kèm tại Phụ lục phân rã tính năng.")

        # ----------------------------------------------------------------------
        # 5. DATA DICTIONARY SECTION
        # ----------------------------------------------------------------------
        doc.add_heading("4. TỪ ĐIỂN DỮ LIỆU & RÀNG BUỘC (DATA DICTIONARY)", level=1)
        if data_fields:
            t_dd = doc.add_table(rows=1, cols=4)
            t_dd.alignment = WD_TABLE_ALIGNMENT.CENTER
            hdr_dd = t_dd.rows[0]
            hdr_dd._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

            dd_cols = ["Tên Trường (Field)", "Kiểu Dữ Liệu", "Ràng Buộc (Null/PK)", "Mô Tả Nghiệp Vụ"]
            dd_w = [Inches(1.8), Inches(1.3), Inches(1.3), Inches(2.4)]
            for idx, text in enumerate(dd_cols):
                c = hdr_dd.cells[idx]
                c.width = dd_w[idx]
                c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0F172A"/>'))
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(text)
                r.bold = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(255, 255, 255)

            for raw_fld in data_fields:
                fld = raw_fld.to_dict() if hasattr(raw_fld, "to_dict") else raw_fld
                r_obj = t_dd.add_row()
                r_obj._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                f_name = fld.get("field", "")
                f_type = fld.get("type", "VARCHAR")
                f_null = fld.get("nullable", "YES")
                f_pk = "PK" if fld.get("pk") == "YES" else ("FK" if fld.get("fk") != "-" else "-")
                f_desc = fld.get("description", "")

                for c_idx, val in enumerate([f_name, f_type, f"{f_null} | {f_pk}", f_desc]):
                    cell = r_obj.cells[c_idx]
                    cell.width = dd_w[c_idx]
                    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CBD5E1"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="CBD5E1"/><w:right w:val="single" w:sz="4" w:color="CBD5E1"/></w:tcBorders>'))
                    cell.paragraphs[0].add_run(val).font.size = Pt(8.5)

        # ----------------------------------------------------------------------
        # 5. REQUIREMENTS TRACEABILITY MATRIX (RTM)
        # ----------------------------------------------------------------------
        doc.add_heading("5. MA TRẬN TRUY XUẤT NGUỒN GỐC (REQUIREMENTS TRACEABILITY MATRIX - RTM)", level=1)
        if stories:
            from .rtm_matrix import RTMEngine
            rtm_items = RTMEngine.build_rtm(stories, usecases)
            if rtm_items:
                t_rtm = doc.add_table(rows=1, cols=5)
                t_rtm.alignment = WD_TABLE_ALIGNMENT.CENTER
                hdr_rtm = t_rtm.rows[0]
                hdr_rtm._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

                rtm_cols = ["Mã Yêu Cầu (BR)", "User Story ID", "Use Case ID", "Tiêu Chí Nghiệm Thu (AC)", "Trạng Thái"]
                rtm_w = [Inches(1.5), Inches(1.3), Inches(1.3), Inches(2.2), Inches(1.1)]
                for idx, text in enumerate(rtm_cols):
                    c = hdr_rtm.cells[idx]
                    c.width = rtm_w[idx]
                    c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0F172A"/>'))
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(text)
                    r.bold = True
                    r.font.size = Pt(8.5)
                    r.font.color.rgb = RGBColor(255, 255, 255)

                for it in rtm_items[:25]:
                    r_obj = t_rtm.add_row()
                    r_obj._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                    for c_idx, val in enumerate([it.req_id, it.story_id, it.usecase_id, it.scenario_title[:45], it.status]):
                        cell = r_obj.cells[c_idx]
                        cell.width = rtm_w[c_idx]
                        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:color="CBD5E1"/><w:bottom w:val="single" w:sz="4" w:color="CBD5E1"/><w:left w:val="single" w:sz="4" w:color="CBD5E1"/><w:right w:val="single" w:sz="4" w:color="CBD5E1"/></w:tcBorders>'))
                        cell.paragraphs[0].add_run(str(val)).font.size = Pt(8)
        else:
            doc.add_paragraph("Ma trận RTM chi tiết sẽ được tự động liên kết sau khi xác nhận danh sách User Stories.")

        # ----------------------------------------------------------------------
        # 6. NON-FUNCTIONAL REQUIREMENTS (NFR)
        # ----------------------------------------------------------------------
        doc.add_heading("6. YÊU CẦU PHI CHỨC NĂNG (NON-FUNCTIONAL REQUIREMENTS - NFR)", level=1)
        nfr_items = [
            ("NFR-PERF-01 (Hiệu năng):", "Thời gian phản hồi API trung bình đạt dưới 300ms với lưu lượng 1,000 requests/giây (TPS)."),
            ("NFR-SEC-01 (Bảo mật):", "Toàn bộ dữ liệu định danh khách hàng (CCCD, SĐT, STK) bắt buộc mã hóa AES-256 at-rest và TLS 1.3 in-transit."),
            ("NFR-AVAIL-01 (Tính sẵn sàng):", "Hệ thống cam kết độ sẵn sàng vận hành đạt mức SLA 99.9% (High Availability)."),
            ("NFR-AUDIT-01 (Kiểm toán):", "Ghi log đầy đủ mọi thao tác chỉnh sửa số liệu và lưu vết kiểm toán (Audit Trail) tối thiểu 5 năm theo quy định pháp lý.")
        ]
        for code, desc in nfr_items:
            p = doc.add_paragraph()
            p.add_run(f"• {code} ").bold = True
            p.add_run(desc)

        # ----------------------------------------------------------------------
        # 7. SIGN-OFF APPROVALS
        # ----------------------------------------------------------------------
        doc.add_heading("7. KÝ DUYỆT & XÁC NHẬN BÀN GIAO (SIGN-OFF & APPROVALS)", level=1)
        p_sign = doc.add_paragraph()
        p_sign.add_run("Bằng chữ ký dưới đây, các bên liên quan xác nhận nội dung tài liệu yêu cầu nghiệp vụ này đã được rà soát đầy đủ và thống nhất triển khai:")

        t_sign = doc.add_table(rows=2, cols=3)
        t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
        sign_roles = ["Đại Diện Khối Nghiệp Vụ (Business Owner)", "Quản Trị Dự Án / PO (Product Owner)", "Trưởng Nhóm Phân Tích (Lead BA)"]

        for idx, role in enumerate(sign_roles):
            c_top = t_sign.rows[0].cells[idx]
            c_bot = t_sign.rows[1].cells[idx]
            c_top.width = Inches(2.2)
            c_bot.width = Inches(2.2)

            p_r = c_top.paragraphs[0]
            p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p_r.add_run(role)
            r.bold = True
            r.font.size = Pt(9.5)

            c_bot.paragraphs[0].add_run("\n\n\n\n___________________________\nNgày: ...../...../2026").font.size = Pt(9)
            c_bot.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path
