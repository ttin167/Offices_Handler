"""
core.diagrams.salt_wireframe — PlantUML Salt UI Wireframe Engine (Rapid Low-Fi Wireframe Studio).

Features:
1. Declarative UI Mockup & Wireframe generation from JSON Spec or pure PlantUML Salt text.
2. Built-in Catalog of 6 Enterprise BA Wireframe Templates:
   - login_auth: Màn hình Đăng Nhập & Xác Thực 2FA (OTP)
   - crud_table: Bảng Quản Trị Dữ Liệu, Bộ Lọc & Phân Trang
   - wizard_form: Biểu Mẫu Nhập Liệu Đa Bước (Multi-Step Wizard)
   - modal_confirm: Hộp Thoại Xác Nhận & Cảnh Báo Rủi Ro (Modal Dialog)
   - mobile_screen: Màn Hình Ứng Dụng Di Động (Mobile App View)
   - dashboard_kpi: Bảng Điều Khiển Quản Trị & Chỉ Số KPI (Executive Dashboard)
3. Dual-Output Rendering:
   - Pure Vector SVG (-tsvg) with transparent background & monochrome skinparams for instant Web Studio preview.
   - 300+ DPI crisp PNG (-tpng) with mandatory -charset UTF-8 for Windows Vietnamese typography fidelity.
4. Direct 1-Click Word (.docx) OpenXML Injection enforcing Printable Margin Guard (ERR_DOCX_005 <= 14.0 cm),
   Last Paragraph Rule (ERR_DOCX_001), keepNext caption binding, and deterministic temp file cleanup (ERR_PIPE_001).
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union
from PIL import Image


WIREFRAME_TEMPLATES: Dict[str, Dict[str, Any]] = {
    "login_auth": {
        "id": "login_auth",
        "title": "Màn Hình Đăng Nhập & Xác Thực 2FA",
        "category": "Authentication",
        "description": "Khung đăng nhập hệ thống doanh nghiệp kèm xác thực hai lớp OTP SMS và ghi nhớ thiết bị.",
        "json_spec": {
            "title": "CỔNG ĐĂNG NHẬP HỆ THỐNG DOANH NGHIỆP (SSO & 2FA)",
            "frame": "+",
            "elements": [
                {"type": "separator"},
                {"type": "input", "label": "Tài khoản / Email:", "value": "ba.lead@enterprise.vn"},
                {"type": "password", "label": "Mật khẩu:", "value": "************"},
                {"type": "droplist", "label": "Vai trò truy cập:", "options": ["Business Analyst", "Product Owner", "System Admin"]},
                {"type": "input", "label": "Mã xác thực OTP (6 số):", "value": "849201"},
                {"type": "checkbox", "label": "Ghi nhớ thiết bị tin cậy trong 30 ngày", "checked": True},
                {"type": "separator"},
                {"type": "button_group", "buttons": ["Đăng Nhập Hệ Thống", "Gửi Lại OTP", "Quên Mật Khẩu?"]}
            ]
        },
        "salt_code": """@startsalt
skinparam BackgroundColor transparent
skinparam monochrome true
{+
  <b>CỔNG ĐĂNG NHẬP HỆ THỐNG DOANH NGHIỆP (SSO & 2FA)
  --
  Tài khoản / Email:      | "ba.lead@enterprise.vn   "
  Mật khẩu:               | "************            "
  Vai trò truy cập:       | ^Business Analyst|Product Owner|System Admin^
  Mã xác thực OTP (6 số): | "849201                  "
  [X] Ghi nhớ thiết bị tin cậy trong 30 ngày
  --
  { [ Đăng Nhập Hệ Thống ] | [ Gửi Lại OTP ] | [ Quên Mật Khẩu? ] }
}
@endsalt"""
    },
    "crud_table": {
        "id": "crud_table",
        "title": "Bảng Quản Trị & Bộ Lọc Dữ Liệu (CRUD Grid)",
        "category": "Data Grid",
        "description": "Màn hình danh sách bản ghi có thanh tìm kiếm, bộ lọc trạng thái, bảng dữ liệu và phân trang.",
        "json_spec": {
            "title": "QUẢN LÝ DANH MỤC YÊU CẦU NGHIỆP VỤ (REQUIREMENTS REGISTRY)",
            "frame": "+",
            "menus": ["Hệ Thống", "Quản Lý Yêu Cầu", "Ma Trận RTM", "Báo Cáo UAT", "Trợ Giúp"],
            "elements": [
                {"type": "input", "label": "Từ khóa tìm kiếm:", "value": "Chuyển tiền 24/7"},
                {"type": "droplist", "label": "Trạng thái:", "options": ["Tất cả trạng thái", "Đã phê duyệt (Approved)", "Đang thẩm định (In Review)", "Bản nháp (Draft)"]},
                {"type": "button_group", "buttons": ["Tìm Kiếm", "+ Thêm Yêu Cầu Mới", "Xuất Excel (.xlsx)"]},
                {"type": "separator"},
                {
                    "type": "table",
                    "headers": ["Mã REQ", "Tên Chức Năng Nghiệp Vụ", "Độ Ưu Tiên", "Trạng Thái", "Thao Tác"],
                    "rows": [
                        ["REQ-001", "Chuyển tiền nhanh NAPAS 24/7", "Cao (High)", "Approved", "[Sửa] | [Xóa]"],
                        ["REQ-002", "Định danh khách hàng điện tử eKYC", "Cao (High)", "Approved", "[Sửa] | [Xóa]"],
                        ["REQ-003", "Đối soát giao dịch tự động cuối ngày", "Trung bình", "In Review", "[Sửa] | [Xóa]"],
                        ["REQ-004", "Cảnh báo giao dịch đáng ngờ AML", "Thấp (Low)", "Draft", "[Sửa] | [Xóa]"]
                    ]
                },
                {"type": "separator"},
                {"type": "button_group", "buttons": ["<< Trang đầu", "< Trước", "Trang 1 / 12", "Tiếp >", "Trang cuối >>"]}
            ]
        },
        "salt_code": """@startsalt
skinparam BackgroundColor transparent
skinparam monochrome true
{+
  <b>QUẢN LÝ DANH MỤC YÊU CẦU NGHIỆP VỤ (REQUIREMENTS REGISTRY)
  {* Hệ Thống | Quản Lý Yêu Cầu | Ma Trận RTM | Báo Cáo UAT | Trợ Giúp }
  --
  {
    Từ khóa: | "Chuyển tiền 24/7        " | Trạng thái: | ^Tất cả|Approved|In Review|Draft^ | [ Tìm Kiếm ] | [ + Thêm Mới ]
  }
  --
  {#
    <b>Mã REQ | <b>Tên Chức Năng Nghiệp Vụ | <b>Độ Ưu Tiên | <b>Trạng Thái | <b>Thao Tác
    REQ-001 | Chuyển tiền nhanh NAPAS 24/7 | Cao (High) | Approved | [Sửa] [Xóa]
    REQ-002 | Định danh khách hàng điện tử eKYC | Cao (High) | Approved | [Sửa] [Xóa]
    REQ-003 | Đối soát giao dịch tự động cuối ngày | Trung bình | In Review | [Sửa] [Xóa]
    REQ-004 | Cảnh báo giao dịch đáng ngờ AML | Thấp (Low) | Draft | [Sửa] [Xóa]
  }
  --
  { Hiển thị 1-4 / 48 bản ghi | [ << ] | [ < Trước ] | Trang 1 / 12 | [ Tiếp > ] | [ >> ] }
}
@endsalt"""
    },
    "wizard_form": {
        "id": "wizard_form",
        "title": "Biểu Mẫu Nhập Liệu Đa Bước (Multi-Step Wizard)",
        "category": "Multi-Step Form",
        "description": "Quy trình khởi tạo hồ sơ gồm nhiều bước tuần tự (Step 1 -> Step 2 -> Step 3) kèm điều hướng.",
        "json_spec": {
            "title": "QUY TRÌNH MỞ TÀI KHOẢN DOANH NGHIỆP (BƯỚC 2/3: THÔNG TIN PHÁP LÝ)",
            "frame": "+",
            "tabs": ["Bước 1: Thông Tin Chung", "Bước 2: Hồ Sơ Pháp Lý (Đang chọn)", "Bước 3: Xác Nhận & Ký Số"],
            "elements": [
                {"type": "separator"},
                {"type": "input", "label": "Tên doanh nghiệp (*):", "value": "CÔNG TY CỔ PHẦN CÔNG NGHỆ SỐ ANTIGRAVITY"},
                {"type": "input", "label": "Mã số thuế / ĐKKD (*):", "value": "0109876543"},
                {"type": "input", "label": "Người đại diện pháp luật (*):", "value": "Nguyễn Văn An"},
                {"type": "droplist", "label": "Loại hình doanh nghiệp:", "options": ["Công ty Cổ phần", "Công ty TNHH 1 Thành viên", "Công ty TNHH 2 TV trở lên"]},
                {"type": "radio", "label": "Phương thức ký hợp đồng: Chữ ký số USB Token / HSM", "selected": True},
                {"type": "radio", "label": "Phương thức ký hợp đồng: Ký trực tiếp tại quầy giao dịch", "selected": False},
                {"type": "checkbox", "label": "Tôi cam kết thông tin khai báo hoàn toàn chính xác theo quy định pháp luật", "checked": True},
                {"type": "separator"},
                {"type": "button_group", "buttons": ["< Quay Lại Bước 1", "Lưu Bản Nháp", "Tiếp Tục Bước 3 >"]}
            ]
        },
        "salt_code": """@startsalt
skinparam BackgroundColor transparent
skinparam monochrome true
{+
  <b>QUY TRÌNH MỞ TÀI KHOẢN DOANH NGHIỆP (BƯỚC 2/3: THÔNG TIN PHÁP LÝ)
  {/ Bước 1: Thông Tin Chung | <b>Bước 2: Hồ Sơ Pháp Lý | Bước 3: Xác Nhận & Ký Số }
  --
  Tên doanh nghiệp (*):         | "CÔNG TY CỔ PHẦN CÔNG NGHỆ SỐ ANTIGRAVITY"
  Mã số thuế / ĐKKD (*):        | "0109876543                              "
  Người đại diện pháp luật (*): | "Nguyễn Văn An                           "
  Loại hình doanh nghiệp:       | ^Công ty Cổ phần|Công ty TNHH 1 TV|Công ty TNHH 2 TV^
  --
  (X) Phương thức ký hợp đồng: Chữ ký số USB Token / HSM
  ( ) Phương thức ký hợp đồng: Ký trực tiếp tại quầy giao dịch
  [X] Tôi cam kết thông tin khai báo hoàn toàn chính xác theo quy định pháp luật
  --
  { [ < Quay Lại Bước 1 ] | [ Lưu Bản Nháp ] | [ Tiếp Tục Bước 3 > ] }
}
@endsalt"""
    },
    "modal_confirm": {
        "id": "modal_confirm",
        "title": "Hộp Thoại Xác Nhận & Cảnh Báo Rủi Ro (Modal Dialog)",
        "category": "Dialog / Alert",
        "description": "Popup xác nhận thao tác quan trọng (phê duyệt, xóa dữ liệu, chuyển tiền vượt hạn mức).",
        "json_spec": {
            "title": "CẢNH BÁO XÁC NHẬN PHÊ DUYỆT LỆNH CHI LƯƠNG HÀNG LOẠT",
            "frame": "+",
            "elements": [
                {"type": "separator"},
                {"type": "text", "label": "Bạn đang thực hiện phê duyệt lệnh chi lương kỳ Tháng 11/2026:"},
                {
                    "type": "table",
                    "headers": ["Chỉ Tiêu Kiểm Soát", "Giá Trị Xác Nhận"],
                    "rows": [
                        ["Tổng số nhân sự thụ hưởng", "245 Nhân viên"],
                        ["Tổng số tiền thực lĩnh (Net Pay)", "6.850.000.000 VNĐ"],
                        ["Tài khoản trích nợ", "190088889999 - Ngân hàng TMCP Ngoại Thương"]
                    ]
                },
                {"type": "input", "label": "Lý do / Ghi chú phê duyệt:", "value": "Đồng ý chi lương đợt 1 tháng 11/2026"},
                {"type": "checkbox", "label": "Đã đối soát khớp 100% với bảng tổng hợp thuế TNCN và BHXH", "checked": True},
                {"type": "separator"},
                {"type": "button_group", "buttons": ["Xác Nhận & Ký Duyệt", "Hủy Bỏ Thao Tác"]}
            ]
        },
        "salt_code": """@startsalt
skinparam BackgroundColor transparent
skinparam monochrome true
{+
  <b>CẢNH BÁO XÁC NHẬN PHÊ DUYỆT LỆNH CHI LƯƠNG HÀNG LOẠT
  --
  Bạn đang thực hiện phê duyệt lệnh chi lương kỳ Tháng 11/2026:
  {#
    <b>Chỉ Tiêu Kiểm Soát | <b>Giá Trị Xác Nhận
    Tổng số nhân sự thụ hưởng | 245 Nhân viên
    Tổng số tiền thực lĩnh (Net Pay) | 6.850.000.000 VNĐ
    Tài khoản trích nợ | 190088889999 - Ngân hàng TMCP Ngoại Thương
  }
  Ghi chú phê duyệt: | "Đồng ý chi lương đợt 1 tháng 11/2026"
  [X] Đã đối soát khớp 100% với bảng tổng hợp thuế TNCN và BHXH
  --
  { [ Xác Nhận & Ký Duyệt ] | [ Hủy Bỏ Thao Tác ] }
}
@endsalt"""
    },
    "mobile_screen": {
        "id": "mobile_screen",
        "title": "Màn Hình Ứng Dụng Di Động (Mobile App View)",
        "category": "Mobile UI",
        "description": "Bố cục màn hình dọc trên điện thoại thông minh với thanh trạng thái, thẻ thông tin và Bottom Tab Bar.",
        "json_spec": {
            "title": "9:41 AM          5G [100%] | VÍ ĐIỆN TỬ DOANH NGHIỆP",
            "frame": "+",
            "elements": [
                {"type": "separator"},
                {"type": "text", "label": "Xin chào, NGUYỄN VĂN AN (Hạng Kim Cương)"},
                {
                    "type": "table",
                    "headers": ["Số Dư Khả Dụng", "Điểm Thưởng Tích Lũy"],
                    "rows": [
                        ["125.450.000 VNĐ", "4.850 Điểm"]
                    ]
                },
                {"type": "button_group", "buttons": ["Chuyển Tiền", "Nạp Tiền", "Quét QR", "Hóa Đơn"]},
                {"type": "separator"},
                {"type": "text", "label": "GIAO DỊCH GẦN ĐÂY:"},
                {
                    "type": "table",
                    "headers": ["Thời Gian", "Nội Dung Giao Dịch", "Số Tiền"],
                    "rows": [
                        ["10:15 05/11", "Thanh toán hóa đơn điện EVN", "-1.450.000đ"],
                        ["08:30 04/11", "Nhận tiền chuyển khoản 24/7", "+25.000.000đ"],
                        ["19:20 02/11", "Mua vé máy bay công tác", "-3.200.000đ"]
                    ]
                },
                {"type": "separator"},
                {"type": "button_group", "buttons": ["[Trang Chủ]", "[Lịch Sử]", "[Thông Báo]", "[Cá Nhân]"]}
            ]
        },
        "salt_code": """@startsalt
skinparam BackgroundColor transparent
skinparam monochrome true
{+
  { 9:41 AM | . | . | 5G [100%] }
  ==
  <b>VÍ ĐIỆN TỬ DOANH NGHIỆP - TRANG CHỦ
  Xin chào, <b>NGUYỄN VĂN AN</b> (Hạng Kim Cương)
  {#
    <b>Số Dư Khả Dụng | <b>Điểm Thưởng Tích Lũy
    125.450.000 VNĐ | 4.850 Điểm
  }
  { [ Chuyển Tiền ] | [ Nạp Tiền ] | [ Quét QR ] | [ Hóa Đơn ] }
  --
  <b>GIAO DỊCH GẦN ĐÂY
  {#
    <b>Thời Gian | <b>Nội Dung Giao Dịch | <b>Số Tiền
    10:15 05/11 | Thanh toán hóa đơn điện EVN | -1.450.000đ
    08:30 04/11 | Nhận tiền chuyển khoản 24/7 | +25.000.000đ
    19:20 02/11 | Mua vé máy bay công tác | -3.200.000đ
  }
  ==
  {/ <b>[Trang Chủ] | [Lịch Sử] | [Thông Báo] | [Cá Nhân] }
}
@endsalt"""
    },
    "dashboard_kpi": {
        "id": "dashboard_kpi",
        "title": "Bảng Điều Khiển Quản Trị & Chỉ Số KPI (Executive Dashboard)",
        "category": "Analytics",
        "description": "Màn hình tổng quan quản trị hiển thị thẻ KPI trọng yếu, bộ lọc kỳ báo cáo và bảng theo dõi tiến độ.",
        "json_spec": {
            "title": "TRUNG TÂM ĐIỀU HÀNH & GIÁM SÁT CHỈ SỐ KPI DỰ ÁN (EXECUTIVE DASHBOARD)",
            "frame": "+",
            "menus": ["Tổng Quan KPI", "Tiến Độ Dự Án", "Chất Lượng UAT", "Ngân Sách", "Cấu Hình"],
            "elements": [
                {"type": "droplist", "label": "Kỳ báo cáo:", "options": ["Quý 4 / 2026", "Quý 3 / 2026", "Cả năm 2026"]},
                {"type": "separator"},
                {
                    "type": "table",
                    "headers": ["Tổng Yêu Cầu (Stories)", "Độ Phủ RTM Coverage", "Tỷ Lệ Pass UAT", "Rủi Ro Tồn Đọng"],
                    "rows": [
                        ["148 Stories (+12%)", "98.5% (Đạt chuẩn)", "94.2% (132/140)", "02 Cảnh báo nhẹ"]
                    ]
                },
                {"type": "separator"},
                {"type": "text", "label": "BẢNG THEO DÕI TIẾN ĐỘ CÁC PHÂN HỆ TRỌNG ĐIỂM:"},
                {
                    "type": "table",
                    "headers": ["Phân Hệ (Module)", "Lead BA Phụ Trách", "Tiến Độ Đặc Tả", "Trạng Thái UAT", "Hành Động"],
                    "rows": [
                        ["Core Banking & Thanh Toán", "Lê Thị Bình", "100% Hoàn tất", "Đang UAT (96%)", "[Xem Chi Tiết]"],
                        ["Quản Trị Nhân Sự & Lương", "Trần Hữu Cường", "100% Hoàn tất", "Đã Nghiệm Thu", "[Xem Chi Tiết]"],
                        ["Cổng Dịch Vụ Khách Hàng", "Phạm Minh Tuấn", "85% Đang viết", "Chờ Kiểm Thử", "[Xem Chi Tiết]"]
                    ]
                },
                {"type": "separator"},
                {"type": "button_group", "buttons": ["Làm Mới Dữ Liệu", "Xuất Báo Cáo Word BRD", "Tải Bảng Tổng Hợp Excel"]}
            ]
        },
        "salt_code": """@startsalt
skinparam BackgroundColor transparent
skinparam monochrome true
{+
  <b>TRUNG TÂM ĐIỀU HÀNH & GIÁM SÁT CHỈ SỐ KPI DỰ ÁN (EXECUTIVE DASHBOARD)
  {* Tổng Quan KPI | Tiến Độ Dự Án | Chất Lượng UAT | Ngân Sách | Cấu Hình }
  --
  { Kỳ báo cáo: | ^Quý 4 / 2026|Quý 3 / 2026|Cả năm 2026^ | Đơn vị: | ^Toàn Khối Công Nghệ^ | [ Lọc Dữ Liệu ] }
  --
  {#
    <b>Tổng Yêu Cầu (Stories) | <b>Độ Phủ RTM Coverage | <b>Tỷ Lệ Pass UAT | <b>Rủi Ro Tồn Đọng
    148 Stories (+12%) | 98.5% (Đạt chuẩn) | 94.2% (132/140) | 02 Cảnh báo nhẹ
  }
  --
  <b>BẢNG THEO DÕI TIẾN ĐỘ CÁC PHÂN HỆ TRỌNG ĐIỂM
  {#
    <b>Phân Hệ (Module) | <b>Lead BA Phụ Trách | <b>Tiến Độ Đặc Tả | <b>Trạng Thái UAT | <b>Hành Động
    Core Banking & Thanh Toán | Lê Thị Bình | 100% Hoàn tất | Đang UAT (96%) | [Xem Chi Tiết]
    Quản Trị Nhân Sự & Lương | Trần Hữu Cường | 100% Hoàn tất | Đã Nghiệm Thu | [Xem Chi Tiết]
    Cổng Dịch Vụ Khách Hàng | Phạm Minh Tuấn | 85% Đang viết | Chờ Kiểm Thử | [Xem Chi Tiết]
  }
  --
  { [ Làm Mới Dữ Liệu ] | [ Xuất Báo Cáo Word BRD ] | [ Tải Bảng Tổng Hợp Excel ] }
}
@endsalt"""
    },
}


class SaltWireframeEngine:
    """Renders PlantUML Salt UI Wireframes from declarative JSON specs, built-in templates, or raw Salt text."""

    DEFAULT_SKINPARAMS = (
        "skinparam BackgroundColor transparent\n"
        "skinparam monochrome true"
    )

    @classmethod
    def get_templates(cls) -> Dict[str, Dict[str, Any]]:
        """Returns the catalog of 6 built-in enterprise BA wireframe templates."""
        return WIREFRAME_TEMPLATES

    @classmethod
    def get_template(cls, template_id: str) -> Optional[Dict[str, Any]]:
        """Retrieves a specific wireframe template by ID."""
        return WIREFRAME_TEMPLATES.get(template_id.strip().lower())

    @classmethod
    def find_plantuml_jar(cls) -> Optional[str]:
        candidates = [
            os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "tools", "plantuml.jar")),
            os.path.abspath(os.path.join(os.path.dirname(__file__), "legacy", "plantuml.jar")),
            os.path.abspath(os.path.join(os.path.dirname(__file__), "plantuml.jar")),
        ]
        for p in candidates:
            if os.path.isfile(p) and os.path.getsize(p) > 1000:
                return p
        return None

    @classmethod
    def _ensure_clean_skinparams(cls, salt_block: str) -> str:
        """Ensures transparent background and monochrome skinparams are injected inside @startsalt block."""
        lines = salt_block.strip().splitlines()
        if not lines:
            return f"@startsalt\n{cls.DEFAULT_SKINPARAMS}\n{{\n  [Wireframe]\n}}\n@endsalt"

        has_bg = any("backgroundcolor" in l.lower() for l in lines)
        has_mono = any("monochrome" in l.lower() for l in lines)

        if lines[0].strip().startswith("@startsalt"):
            inject_lines = []
            if not has_bg:
                inject_lines.append("skinparam BackgroundColor transparent")
            if not has_mono:
                inject_lines.append("skinparam monochrome true")
            if inject_lines:
                lines = [lines[0]] + inject_lines + lines[1:]
            return "\n".join(lines)

        return salt_block

    @classmethod
    def spec_to_salt_code(cls, spec: Union[str, Dict[str, Any]]) -> str:
        """Converts raw text, template ID, or structured JSON specification into PlantUML Salt code."""
        if isinstance(spec, str):
            clean = spec.strip()
            # Check if string is a template ID
            if clean.lower() in WIREFRAME_TEMPLATES:
                return WIREFRAME_TEMPLATES[clean.lower()]["salt_code"]

            # Check if string is JSON
            if clean.startswith("{") and clean.endswith("}") and ('"elements"' in clean or '"title"' in clean or '"salt"' in clean):
                try:
                    parsed_dict = json.loads(clean)
                    if isinstance(parsed_dict, dict):
                        return cls.spec_to_salt_code(parsed_dict)
                except Exception:
                    pass

            if clean.startswith("@startsalt") and clean.endswith("@endsalt"):
                return cls._ensure_clean_skinparams(clean)
            if clean.startswith("{") and clean.endswith("}"):
                return cls._ensure_clean_skinparams(f"@startsalt\n{clean}\n@endsalt")
            return cls._ensure_clean_skinparams(f"@startsalt\n{{\n{clean}\n}}\n@endsalt")

        if not isinstance(spec, dict):
            return cls._ensure_clean_skinparams("@startsalt\n{\n[Wireframe Spec Error]\n}\n@endsalt")

        # Check if spec references a template_id directly without custom elements
        if "template" in spec and spec["template"] in WIREFRAME_TEMPLATES and "elements" not in spec and "salt" not in spec and "code" not in spec:
            return WIREFRAME_TEMPLATES[spec["template"]]["salt_code"]

        # If spec has direct "code" or "salt" field
        if "salt" in spec and spec["salt"]:
            return cls.spec_to_salt_code(spec["salt"])
        if "code" in spec and spec["code"]:
            return cls.spec_to_salt_code(spec["code"])
        if "json_spec" in spec and isinstance(spec["json_spec"], dict):
            return cls.spec_to_salt_code(spec["json_spec"])

        # Build declarative Salt syntax
        title = spec.get("title", "")
        frame_type = spec.get("frame", "+")  # '+' for window with border, '#' for table, ' ' for plain
        elements = spec.get("elements", [])
        menus = spec.get("menus", [])
        tabs = spec.get("tabs", [])

        lines = [f"{{{frame_type}"]

        if title:
            lines.append(f"  <b>{title}")

        if menus:
            menu_str = " | ".join(menus)
            lines.append(f"  {{* {menu_str} }}")

        if tabs:
            tab_str = " | ".join(tabs)
            lines.append(f"  {{/ {tab_str} }}")

        # Process UI elements
        for elem in elements:
            if isinstance(elem, str):
                lines.append(f"  {elem}")
                continue

            if not isinstance(elem, dict):
                continue

            e_type = elem.get("type", "text").lower()
            label = elem.get("label", "")
            val = elem.get("value") or elem.get("default") or ""

            if e_type in ("input", "textbox", "text_field"):
                placeholder = val or " " * 15
                if label:
                    lines.append(f'  {{ {label} | "{placeholder}" }}')
                else:
                    lines.append(f'  "{placeholder}"')

            elif e_type in ("password", "pwd"):
                placeholder = val or "****"
                if label:
                    lines.append(f'  {{ {label} | "{placeholder}" }}')
                else:
                    lines.append(f'  "{placeholder}"')

            elif e_type in ("droplist", "combobox", "select"):
                options = elem.get("options", ["-- Chọn --"])
                opt_str = "|".join(options)
                if label:
                    lines.append(f"  {{ {label} | ^{opt_str}^ }}")
                else:
                    lines.append(f"  ^{opt_str}^")

            elif e_type in ("checkbox", "check"):
                checked = "[X]" if elem.get("checked", False) else "[ ]"
                lines.append(f"  {checked} {label}")

            elif e_type in ("radio", "radio_button"):
                selected = "(X)" if elem.get("selected", False) else "( )"
                lines.append(f"  {selected} {label}")

            elif e_type in ("button", "btn"):
                lines.append(f"  [ {label} ]")

            elif e_type == "button_group":
                buttons = elem.get("buttons", [label] if label else ["OK", "Cancel"])
                b_str = " | ".join([f"[ {b} ]" for b in buttons])
                lines.append(f"  {{ {b_str} }}")

            elif e_type in ("table", "grid"):
                headers = elem.get("headers", [])
                rows = elem.get("rows", [])
                table_lines = ["  {#"]
                if headers:
                    table_lines.append("    <b>" + " | <b>".join([str(h) for h in headers]))
                for row in rows:
                    if isinstance(row, list):
                        table_lines.append("    " + " | ".join([str(c) for c in row]))
                table_lines.append("  }")
                lines.extend(table_lines)

            elif e_type == "tree":
                nodes = elem.get("nodes", [])
                lines.append("  {T")
                for n in nodes:
                    lines.append(f"    {n}")
                lines.append("  }")

            elif e_type == "separator":
                lines.append("  --")

            elif e_type == "double_separator":
                lines.append("  ==")

            else:
                lines.append(f"  {label}")

        lines.append("}")
        salt_body = "\n".join(lines)
        return cls._ensure_clean_skinparams(f"@startsalt\n{salt_body}\n@endsalt")

    @staticmethod
    def _extract_svg_dimensions(svg_text: str) -> Tuple[int, int]:
        """Extracts pixel width and height from PlantUML SVG markup."""
        w_match = re.search(r'width=["\'](\d+(?:\.\d+)?)(?:px)?["\']', svg_text)
        h_match = re.search(r'height=["\'](\d+(?:\.\d+)?)(?:px)?["\']', svg_text)
        if w_match and h_match:
            return int(float(w_match.group(1))), int(float(h_match.group(1)))

        vb_match = re.search(r'viewBox=["\']\s*[\d\.\-]+\s+[\d\.\-]+\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*["\']', svg_text)
        if vb_match:
            return int(float(vb_match.group(1))), int(float(vb_match.group(2)))

        return 480, 320

    @classmethod
    def render_salt_to_svg(
        cls,
        spec_or_code: Union[str, Dict[str, Any]],
        output_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Renders Salt UI Wireframe to pure vector SVG string (and optional .svg file) with UTF-8 encoding."""
        jar_path = cls.find_plantuml_jar()
        if not jar_path:
            raise FileNotFoundError("Không tìm thấy plantuml.jar trong thư mục tools/.")

        salt_code = cls.spec_to_salt_code(spec_or_code)

        cmd = [
            "java",
            "-Djava.awt.headless=true",
            "-DPLANTUML_LIMIT_SIZE=16384",
            "-jar",
            jar_path,
            "-charset",
            "UTF-8",
            "-tsvg",
            "-pipe",
        ]

        proc = subprocess.run(
            cmd,
            input=salt_code,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            timeout=30,
        )
        if proc.returncode != 0 or "<svg" not in proc.stdout:
            raise RuntimeError(f"Lỗi kết xuất PlantUML Salt SVG: {proc.stderr or proc.stdout}")

        svg_content = proc.stdout.strip()
        idx = svg_content.find("<svg")
        if idx > 0:
            svg_content = svg_content[idx:]

        w, h = cls._extract_svg_dimensions(svg_content)

        saved_path: Optional[str] = None
        if output_path:
            saved_path = os.path.abspath(output_path)
            os.makedirs(os.path.dirname(saved_path), exist_ok=True)
            with open(saved_path, "w", encoding="utf-8") as f:
                f.write(svg_content)

        return {
            "success": True,
            "svg": svg_content,
            "svg_path": saved_path,
            "dimensions_px": (w, h),
            "width": w,
            "height": h,
            "salt_code": salt_code,
        }

    @classmethod
    def render_salt_to_png(
        cls,
        spec_or_code: Union[str, Dict[str, Any]],
        output_path: str,
        scale: float = 2.0,
        dpi: int = 300,
    ) -> Dict[str, Any]:
        """Renders Salt UI Wireframe to a high-resolution PNG file with mandatory -charset UTF-8."""
        jar_path = cls.find_plantuml_jar()
        if not jar_path:
            raise FileNotFoundError("Không tìm thấy plantuml.jar trong thư mục tools/.")

        salt_code = cls.spec_to_salt_code(spec_or_code)

        with tempfile.TemporaryDirectory() as tmp_dir:
            puml_file = os.path.join(tmp_dir, "wireframe.puml")
            with open(puml_file, "w", encoding="utf-8") as f:
                f.write(salt_code)

            cmd = [
                "java",
                "-Djava.awt.headless=true",
                "-DPLANTUML_LIMIT_SIZE=16384",
                "-jar",
                jar_path,
                "-charset",
                "UTF-8",
                "-tpng",
                puml_file,
                "-o",
                tmp_dir,
            ]

            proc = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                timeout=30,
            )
            if proc.returncode != 0:
                raise RuntimeError(f"Lỗi render PlantUML Salt: {proc.stderr or proc.stdout}")

            expected_png = os.path.join(tmp_dir, "wireframe.png")
            if not os.path.isfile(expected_png):
                raise FileNotFoundError(f"PlantUML không sinh ra tệp PNG mong muốn: {proc.stderr}")

            abs_out = os.path.abspath(output_path)
            os.makedirs(os.path.dirname(abs_out), exist_ok=True)
            shutil.copy2(expected_png, abs_out)

        with Image.open(abs_out) as im:
            w, h = im.size

        return {
            "success": True,
            "png_path": abs_out,
            "dimensions_px": (w, h),
            "width": w,
            "height": h,
            "salt_code": salt_code,
        }

    @classmethod
    def inject_wireframe_into_docx(
        cls,
        docx_path: str,
        spec_or_code: Union[str, Dict[str, Any]],
        output_path: Optional[str] = None,
        heading: Optional[str] = None,
        caption: Optional[str] = "Hình: Phác thảo giao diện màn hình (UI Wireframe Mockup)",
        max_width_cm: float = 14.0,
    ) -> Dict[str, Any]:
        """
        Renders a wireframe to 300 DPI PNG and injects it into a Word (.docx) document.
        Enforces Printable Margin Guard (ERR_DOCX_005 <= 14.0 cm), Last Paragraph Rule (ERR_DOCX_001),
        and deterministic temporary file cleanup (ERR_PIPE_001).
        """
        from docx import Document
        from docx.shared import Cm, Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        abs_docx = os.path.abspath(docx_path)
        dest_docx = os.path.abspath(output_path) if output_path else abs_docx

        fd, temp_png = tempfile.mkstemp(prefix="step_wireframe_temp_", suffix=".png")
        os.close(fd)

        try:
            render_res = cls.render_salt_to_png(spec_or_code, temp_png, scale=2.0, dpi=300)
            img_w_px, img_h_px = render_res["dimensions_px"]
            aspect = max(img_w_px, 1) / max(img_h_px, 1)

            if os.path.isfile(abs_docx):
                doc = Document(abs_docx)
            else:
                doc = Document()

            # Calculate printable page width (ERR_DOCX_005)
            current_sec = doc.sections[-1]
            pw_cm = current_sec.page_width.cm if current_sec.page_width else 21.0
            lm_cm = current_sec.left_margin.cm if current_sec.left_margin else 2.54
            rm_cm = current_sec.right_margin.cm if current_sec.right_margin else 2.54
            usable_width_cm = max(pw_cm - lm_cm - rm_cm, 10.0)

            # Safe width cap <= 14.0 cm and <= usable_width_cm
            safe_limit_cm = min(float(max_width_cm), usable_width_cm, 14.0)
            # Natural width estimation at 144 effective screen DPI, capped at safe_limit_cm
            natural_w_cm = (img_w_px / 110.0) * 2.54
            final_w_cm = min(max(natural_w_cm, 8.5), safe_limit_cm)
            final_h_cm = final_w_cm / aspect

            # Locate target heading if provided
            anchor_para = None
            if heading:
                target_idx = None
                for i, p in enumerate(doc.paragraphs):
                    if heading.strip().lower() in p.text.strip().lower():
                        target_idx = i
                        break
                if target_idx is not None and (target_idx + 1) < len(doc.paragraphs):
                    anchor_para = doc.paragraphs[target_idx + 1]
                elif target_idx is None:
                    h_para = doc.add_heading(heading, level=2)
                    pPr = h_para._element.get_or_add_pPr()
                    if pPr.find(qn("w:keepNext")) is None:
                        pPr.append(OxmlElement("w:keepNext"))

            if anchor_para is not None:
                img_para = anchor_para.insert_paragraph_before()
            else:
                img_para = doc.add_paragraph()

            img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            img_para.paragraph_format.space_before = Pt(6.0)
            img_para.paragraph_format.space_after = Pt(4.0)

            if caption:
                pPr = img_para._element.get_or_add_pPr()
                if pPr.find(qn("w:keepNext")) is None:
                    pPr.append(OxmlElement("w:keepNext"))

            run = img_para.add_run()
            run.add_picture(temp_png, width=Cm(final_w_cm), height=Cm(final_h_cm))

            if caption:
                if anchor_para is not None:
                    cap_para = anchor_para.insert_paragraph_before()
                else:
                    cap_para = doc.add_paragraph()
                cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap_para.paragraph_format.space_before = Pt(4.0)
                cap_para.paragraph_format.space_after = Pt(12.0)
                c_run = cap_para.add_run(caption)
                c_run.italic = True
                c_run.font.name = "Times New Roman"
                c_run.font.size = Pt(10.0)
                c_run.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

            # Enforce ERR_DOCX_001: ensure every table cell ends with a <w:p>
            for tbl in doc.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        tc = cell._tc
                        if len(tc.xpath("./w:p")) == 0:
                            tc.append(OxmlElement("w:p"))

            os.makedirs(os.path.dirname(dest_docx), exist_ok=True)
            doc.save(dest_docx)

            return {
                "success": True,
                "docx_path": dest_docx,
                "width_cm": round(final_w_cm, 2),
                "height_cm": round(final_h_cm, 2),
                "usable_margin_cm": round(usable_width_cm, 2),
                "dimensions_px": (img_w_px, img_h_px),
                "heading": heading,
                "caption": caption,
            }
        finally:
            if os.path.exists(temp_png):
                try:
                    os.remove(temp_png)
                except OSError:
                    pass
