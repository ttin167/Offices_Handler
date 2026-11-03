"""
tests/test_ba_meeting_and_wireframe.py — Test Suite for Meeting Notes Parser & PlantUML Salt Wireframe Engine.
"""

import os
import tempfile
import pytest

from core.ba_hub.meeting_parser import MeetingNotesParser
from core.diagrams.salt_wireframe import SaltWireframeEngine
from core.diagrams.bridge import render_diagram_code_to_png


SAMPLE_MEETING_NOTES = """# Họp Đặc Tả Hệ Thống Ví Điện Tử Doanh Nghiệp
Ngày: 15/10/2026
Thành phần tham gia: Nguyễn Văn An (Product Owner), Lê Thị Bình (Lead BA), Trần Hữu Cường (Tech Lead)

Quyết định: Thống nhất giới hạn chuyển tiền tối đa 100.000.000 VNĐ / ngày đối với tài khoản cá nhân.
Quy định: Bắt buộc xác thực sinh trắc học khuôn mặt đối với mọi giao dịch trên 10.000.000 VNĐ.

Khách hàng cá nhân đã liên kết ngân hàng cần có thể nạp tiền vào ví điện tử để thanh toán hóa đơn điện nước nhanh chóng.

Action item: Hoàn thiện tài liệu kiến trúc kỹ thuật (Cường, 20/10/2026)
Todo: Soạn thảo ma trận RTM và tiêu chí UAT (Bình, 18/10/2026)

Cần làm rõ: Đối soát với cổng thanh toán đối tác diễn ra vào 23h hàng ngày hay real-time?
"""


def test_meeting_notes_parser():
    parsed = MeetingNotesParser.parse_text(SAMPLE_MEETING_NOTES)
    assert "Ví Điện Tử Doanh Nghiệp" in parsed.project_title
    assert "15/10/2026" in parsed.meeting_date
    assert len(parsed.attendees) >= 2
    assert len(parsed.decisions) >= 1
    assert len(parsed.business_rules) >= 1
    assert len(parsed.user_stories) >= 1
    assert len(parsed.action_items) >= 2

    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, "meeting.docx")
        md_path = os.path.join(tmp, "meeting.md")

        MeetingNotesParser.export_to_docx(parsed, docx_path)
        assert os.path.isfile(docx_path)
        assert os.path.getsize(docx_path) > 1000

        MeetingNotesParser.export_to_markdown(parsed, md_path)
        assert os.path.isfile(md_path)
        assert os.path.getsize(md_path) > 100


def test_salt_wireframe_engine():
    spec = {
        "title": "Cửa Sổ Xác Thực Giao Dịch",
        "elements": [
            {"type": "input", "label": "Mã OTP SMS:"},
            {"type": "checkbox", "label": "Ghi nhớ thiết bị này", "checked": True},
            {"type": "button_group", "buttons": ["Xác Nhận", "Gửi Lại OTP", "Hủy"]}
        ]
    }
    with tempfile.TemporaryDirectory() as tmp:
        out_png = os.path.join(tmp, "wireframe.png")
        res = SaltWireframeEngine.render_salt_to_png(spec, out_png)
        assert res["success"] is True
        assert os.path.isfile(out_png)
        assert res["width"] > 50
        assert res["height"] > 50

        # Also test via bridge
        bridge_out = os.path.join(tmp, "bridge_wireframe.png")
        b_res = render_diagram_code_to_png(spec, "salt", output_png=bridge_out)
        assert b_res is not None
        assert b_res["success"] is True
        assert os.path.isfile(bridge_out)

