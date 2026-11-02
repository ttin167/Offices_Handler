"""
core.ba_hub.templates — Standard Industry Templates & Blueprints for Business Analysts.

Provides production-grade, BABOK & IEEE-compliant templates for:
1. User Story & Gherkin Acceptance Criteria (Happy Path, Negative, Boundary).
2. Use Case Specifications (E-commerce, Banking, Logistics).
3. Business Requirements Document (BRD / SRS) structural blueprints.
4. Data Dictionary & Source-to-Target Mapping matrix.
5. UAT Test Case & Acceptance Checklist formats.
"""

from __future__ import annotations
from typing import Any, Dict, List

# ------------------------------------------------------------------------------
# 1. INDUSTRY USER STORY & GHERKIN TEMPLATES
# ------------------------------------------------------------------------------
STORY_TEMPLATES: Dict[str, Dict[str, Any]] = {
    "banking_fund_transfer": {
        "title": "Chuyển tiền nhanh liên ngân hàng 24/7 (NAPAS 247)",
        "domain": "Ngân hàng số (Digital Banking)",
        "role": "Khách hàng cá nhân đã định danh (eKYC)",
        "goal": "thực hiện chuyển tiền nhanh 24/7 sang ngân hàng khác bằng số tài khoản hoặc số thẻ",
        "benefit": "người thụ hưởng nhận được tiền ngay lập tức mà không cần đến quầy giao dịch",
        "business_rules": [
            "Hạn mức tối đa 500,000,000 VNĐ / giao dịch đối với xác thực Smart OTP.",
            "Số dư tài khoản trích nợ phải lớn hơn hoặc bằng (Số tiền chuyển + Phí giao dịch).",
            "Giao dịch trên 10,000,000 VNĐ bắt buộc xác thực sinh trắc học khuôn mặt theo QĐ 2345/QĐ-NHNN."
        ],
        "scenarios": [
            {
                "title": "Chuyển tiền thành công trong hạn mức thông thường",
                "type": "Happy Path",
                "given": "Khách hàng đã đăng nhập ứng dụng Mobile Banking và có số dư khả dụng 5,000,000 VNĐ",
                "when": "Khách hàng nhập số tài khoản đích hợp lệ, số tiền 2,000,000 VNĐ và xác thực Smart OTP thành công",
                "then": "Hệ thống trích nợ 2,000,000 VNĐ từ tài khoản nguồn, cộng tiền tài khoản đích qua NAPAS",
                "and": "Hiển thị màn hình biên lai giao dịch thành công kèm mã tham chiếu FT-XXXXXXXX và gửi thông báo biến động số dư qua OTT/SMS"
            },
            {
                "title": "Từ chối giao dịch do số dư khả dụng không đủ",
                "type": "Negative Path",
                "given": "Khách hàng có số dư khả dụng 1,000,000 VNĐ và phí giao dịch là 2,200 VNĐ",
                "when": "Khách hàng yêu cầu chuyển số tiền 1,000,000 VNĐ",
                "then": "Hệ thống hiển thị thông báo lỗi 'Số dư khả dụng không đủ để thanh toán số tiền và phí giao dịch'",
                "and": "Không gửi lệnh sang cổng thanh toán NAPAS và không trừ tiền tài khoản nguồn"
            },
            {
                "title": "Bắt buộc xác thực sinh trắc học với giao dịch lớn",
                "type": "Regulatory Boundary",
                "given": "Khách hàng thực hiện lệnh chuyển tiền 15,000,000 VNĐ (vượt ngưỡng 10 triệu đồng)",
                "when": "Khách hàng nhấn nút 'Tiếp tục' tại màn hình xác nhận thông tin",
                "then": "Hệ thống kích hoạt camera quét khuôn mặt đối chiếu với dữ liệu CCCD gắn chip lưu tại hệ thống",
                "and": "Chỉ kích hoạt bước nhập mã PIN Smart OTP sau khi xác thực khuôn mặt đạt độ khớp >= 90%"
            }
        ]
    },
    "ecommerce_checkout": {
        "title": "Thanh toán đơn hàng và Áp dụng Mã giảm giá (Coupon Checkout)",
        "domain": "Thương mại điện tử (E-Commerce)",
        "role": "Người mua hàng trực tuyến (Registered Shopper)",
        "goal": "chọn phương thức thanh toán và áp dụng mã voucher giảm giá tại bước thanh toán",
        "benefit": "tiết kiệm chi phí mua sắm và hoàn tất đơn hàng nhanh chóng",
        "business_rules": [
            "Mỗi đơn hàng chỉ được áp dụng tối đa 1 mã Giảm giá sàn và 1 mã Miễn phí vận chuyển.",
            "Tồn kho của sản phẩm phải được khóa tạm thời (Reservation Lock) trong vòng 15 phút.",
            "Nếu quá 15 phút chưa hoàn tất thanh toán qua cổng trực tuyến, đơn hàng tự động hủy và hoàn trả tồn kho."
        ],
        "scenarios": [
            {
                "title": "Áp dụng voucher hợp lệ và thanh toán COD",
                "type": "Happy Path",
                "given": "Giỏ hàng có tổng giá trị 600,000 VNĐ và voucher 'SALE50K' có hiệu lực cho đơn từ 500,000 VNĐ",
                "when": "Người mua áp dụng mã 'SALE50K' và chọn phương thức 'Thanh toán khi nhận hàng (COD)'",
                "then": "Tổng tiền thanh toán được trừ 50,000 VNĐ còn 550,000 VNĐ",
                "and": "Đơn hàng được khởi tạo ở trạng thái 'Chờ đóng gói' và gửi email xác nhận cho người mua"
            },
            {
                "title": "Voucher hết lượt sử dụng hoặc không đạt giá trị tối thiểu",
                "type": "Negative Path",
                "given": "Giỏ hàng có tổng giá trị 300,000 VNĐ",
                "when": "Người mua nhập mã 'VIP100K' (yêu cầu đơn từ 500,000 VNĐ)",
                "then": "Hệ thống hiển thị cảnh báo đỏ 'Đơn hàng chưa đạt giá trị tối thiểu 500,000 VNĐ để áp dụng mã này'",
                "and": "Tổng tiền đơn hàng giữ nguyên không giảm giá"
            }
        ]
    },
    "wms_inventory_inbound": {
        "title": "Nhập kho sản phẩm theo Đơn đặt hàng mua (PO Inbound Receiving)",
        "domain": "Quản lý chuỗi cung ứng & Kho bãi (WMS)",
        "role": "Nhân viên kiểm hàng nhập kho (Inbound Inspector)",
        "goal": "quét mã vạch barcode/RFID của các pallet hàng nhập theo mã đơn PO",
        "benefit": "đối soát số lượng thực tế với đơn đặt hàng và cập nhật tồn kho chính xác theo thời gian thực",
        "business_rules": [
            "Hàng hóa không có mã PO hợp lệ ở trạng thái 'Approved' sẽ bị từ chối nhập.",
            "Số lượng nhập thừa vượt quá 5% dung sai (Over-receiving tolerance) phải có phê duyệt của Trưởng kho.",
            "Hàng cận hạn sử dụng (< 6 tháng) bắt buộc chuyển vào khu vực cách ly (Quarantine Area)."
        ],
        "scenarios": [
            {
                "title": "Nhập kho đủ số lượng và tự động cập nhật vị trí lưu kho (Putaway)",
                "type": "Happy Path",
                "given": "Đơn hàng PO-2026-088 yêu cầu nhập 100 thùng sữa, trạng thái PO là 'Approved'",
                "when": "Nhân viên quét đủ 100 thùng đạt tiêu chuẩn chất lượng và xác nhận 'Hoàn tất nhập hàng'",
                "then": "Hệ thống WMS tăng tồn kho khả dụng thêm 100 thùng tại khu vực Inbound",
                "and": "Gợi ý danh sách vị trí kệ trống (Bin Location) tối ưu để xếp hàng lên kệ"
            }
        ]
    }
}

# ------------------------------------------------------------------------------
# 2. USE CASE SPECIFICATION TEMPLATES
# ------------------------------------------------------------------------------
USECASE_TEMPLATES: Dict[str, Dict[str, Any]] = {
    "uc_fund_transfer": {
        "id": "UC-PAY-001",
        "name": "Chuyển Tiền Liên Ngân Hàng Nhanh 24/7",
        "module": "Phân hệ Thanh toán & Giao dịch",
        "primary_actor": "Khách hàng cá nhân (Mobile App User)",
        "secondary_actors": ["Hệ thống Core Banking (T24/Finacle)", "Cổng chuyển mạch tài chính NAPAS", "Hệ thống SMS/OTT Gateway"],
        "preconditions": [
            "Khách hàng đã đăng nhập thành công vào ứng dụng ngân hàng số.",
            "Tài khoản thanh toán nguồn ở trạng thái hoạt động (Active) và không bị phong tỏa.",
            "Phương thức xác thực Smart OTP đã được kích hoạt trên thiết bị."
        ],
        "postconditions": [
            "Tài khoản nguồn bị trích trừ đúng số tiền chuyển và phí giao dịch tương ứng.",
            "Tài khoản thụ hưởng tại ngân hàng đích nhận được tiền ghi có ngay lập tức.",
            "Bản ghi giao dịch được lưu trữ đầy đủ trong lịch sử giao dịch Core Banking và phục vụ đối soát."
        ],
        "trigger": "Khách hàng chọn chức năng 'Chuyển tiền nhanh 24/7' trên menu trang chủ ứng dụng.",
        "main_flow": [
            {"step": "1.0", "actor": "Khách hàng", "action": "Chọn tài khoản nguồn trích nợ từ danh sách tài khoản thanh toán khả dụng."},
            {"step": "2.0", "actor": "Khách hàng", "action": "Chọn ngân hàng thụ hưởng và nhập số tài khoản hoặc số thẻ người nhận."},
            {"step": "3.0", "actor": "Hệ thống", "action": "Gửi yêu cầu truy vấn thông tin tài khoản sang NAPAS và tự động hiển thị tên chủ tài khoản thụ hưởng (Inquiry Name)."},
            {"step": "4.0", "actor": "Khách hàng", "action": "Nhập số tiền cần chuyển và nội dung thanh toán, sau đó nhấn 'Tiếp tục'."},
            {"step": "5.0", "actor": "Hệ thống", "action": "Kiểm tra hạn mức giao dịch, tính toán biểu phí dịch vụ và hiển thị màn hình xác nhận thông tin tổng thể."},
            {"step": "6.0", "actor": "Khách hàng", "action": "Kiểm tra lại toàn bộ thông tin thụ hưởng và xác nhận giao dịch qua Smart OTP (kèm Sinh trắc học nếu > 10 triệu đồng)."},
            {"step": "7.0", "actor": "Hệ thống", "action": "Gửi lệnh trích nợ sang Core Banking, chuyển mạch giao dịch qua NAPAS và thông báo kết quả thành công kèm biên lai."}
        ],
        "alternative_flows": [
            {
                "id": "Alt-1",
                "condition": "Tại bước 2.0, khách hàng chọn người nhận từ danh bạ thụ hưởng đã lưu",
                "steps": [
                    "2.1: Hệ thống tự động điền ngân hàng đích, số tài khoản và tên người nhận.",
                    "2.2: Luồng nghiệp vụ tiếp tục từ bước 4.0 của luồng chính."
                ]
            }
        ],
        "exception_flows": [
            {
                "id": "Exc-1",
                "condition": "Tại bước 3.0, ngân hàng thụ hưởng hoặc NAPAS bảo trì hoặc số tài khoản không tồn tại",
                "steps": [
                    "3.E1: Hệ thống hiển thị thông báo lỗi 'Tài khoản thụ hưởng không tồn tại hoặc ngân hàng đích đang bảo trì kết nối'.",
                    "3.E2: Hệ thống cho phép khách hàng sửa lại thông tin hoặc quay về trang chủ."
                ]
            },
            {
                "id": "Exc-2",
                "condition": "Tại bước 5.0, số dư khả dụng không đủ để trích nợ số tiền và phí",
                "steps": [
                    "5.E1: Hệ thống tô đỏ trường số tiền và hiển thị cảnh báo 'Số dư không đủ'.",
                    "5.E2: Nút 'Tiếp tục' bị vô hiệu hóa cho đến khi khách hàng điều chỉnh lại số tiền hợp lệ."
                ]
            }
        ],
        "business_rules": [
            "BR-01: Tên người thụ hưởng hiển thị ở bước 3.0 phải là chữ in hoa không dấu chuẩn định dạng liên ngân hàng.",
            "BR-02: Nội dung chuyển tiền tự động loại bỏ các ký tự đặc biệt nguy hiểm (ngăn ngừa SQL/XSS Injection).",
            "BR-03: Giao dịch thất bại tại cổng NAPAS phải kích hoạt cơ chế hoàn tiền tự động (Auto-Reversal) trong vòng tối đa 30 giây."
        ]
    }
}

# ------------------------------------------------------------------------------
# 3. DATA DICTIONARY BLUEPRINTS
# ------------------------------------------------------------------------------
DATA_DICT_TEMPLATES: Dict[str, List[Dict[str, Any]]] = {
    "banking_customer_account": [
        {"field": "customer_id", "name": "Mã Khách Hàng (CIF)", "type": "VARCHAR(20)", "nullable": "NO", "pk": "YES", "fk": "-", "description": "Mã định danh duy nhất của khách hàng trên toàn hệ thống ngân hàng (CIF).", "rules": "Định dạng CIF-XXXXXXXX, sinh tự động từ Core Banking"},
        {"field": "id_card_no", "name": "Số CCCD / Hộ Chiếu", "type": "VARCHAR(20)", "nullable": "NO", "pk": "NO", "fk": "-", "description": "Số giấy tờ tùy thân hợp lệ đã xác thực eKYC.", "rules": "12 chữ số hợp lệ theo chuẩn Bộ Công An hoặc số hộ chiếu"},
        {"field": "account_number", "name": "Số Tài Khoản Thanh Toán", "type": "VARCHAR(25)", "nullable": "NO", "pk": "NO", "fk": "-", "description": "Số tài khoản thanh toán VND phục vụ nhận và chuyển tiền.", "rules": "Chuỗi ký tự số, độ dài 10-16 ký tự, kiểm tra thuật toán Modulo 11"},
        {"field": "available_balance", "name": "Số Dư Khả Dụng", "type": "DECIMAL(18,2)", "nullable": "NO", "pk": "NO", "fk": "-", "description": "Số dư thực tế có thể sử dụng giao dịch (không bao gồm số dư phong tỏa).", "rules": "Giá trị >= 0, định dạng tiền tệ VNĐ"},
        {"field": "currency_code", "name": "Mã Tiền Tệ", "type": "CHAR(3)", "nullable": "NO", "pk": "NO", "fk": "-", "description": "Đơn vị tiền tệ theo chuẩn ISO 4217.", "rules": "Mặc định 'VND', hỗ trợ 'USD', 'EUR'"},
        {"field": "account_status", "name": "Trạng Thái Tài Khoản", "type": "VARCHAR(15)", "nullable": "NO", "pk": "NO", "fk": "-", "description": "Trạng thái vận hành của tài khoản.", "rules": "Giá trị trong tập: 'ACTIVE', 'BLOCKED', 'DORMANT', 'CLOSED'"}
    ]
}

# ------------------------------------------------------------------------------
# 4. BRD DOCUMENT STRUCTURE OUTLINE
# ------------------------------------------------------------------------------
BRD_DEFAULT_OUTLINE: Dict[str, Any] = {
    "doc_title": "TÀI LIỆU YÊU CẦU NGHIỆP VỤ (BRD - BUSINESS REQUIREMENTS DOCUMENT)",
    "system_name": "Hệ Thống Ngân Hàng Số & Thanh Toán Đa Kênh",
    "version": "1.0",
    "confidentiality": "Nội bộ / Confidential",
    "organization": "Khối Công Nghệ & Khối Vận Hành Doanh Nghiệp",
    "sections": [
        {"id": "sec_1", "title": "1. Tổng Quan Dự Án & Mục Tiêu Nghiệp Vụ (Executive Summary)", "desc": "Bối cảnh phát sinh dự án, vấn đề cần giải quyết, và các chỉ số đo lường hiệu quả (KPI/ROI)."},
        {"id": "sec_2", "title": "2. Phạm Vi Dự Án (Project Scope)", "desc": "Chi tiết các chức năng nằm trong phạm vi triển khai (In-Scope) và các phần hoãn lại (Out-of-Scope)."},
        {"id": "sec_3", "title": "3. Ma Trận Các Bên Liên Quan (Stakeholders & User Personas)", "desc": "Danh sách các nhóm người dùng nội bộ, khách hàng cuối và quyền hạn tương ứng."},
        {"id": "sec_4", "title": "4. Kiến Trúc Quy Trình Nghiệp Vụ (Business Process Architecture)", "desc": "Sơ đồ luồng quy trình tổng thể (BPMN / Swimlane) và tương tác hệ thống."},
        {"id": "sec_5", "title": "5. Yêu Cầu Chức Năng (Functional Requirements - User Stories & Use Cases)", "desc": "Danh mục chi tiết Use Case, User Story kèm tiêu chí nghiệm thu Gherkin."},
        {"id": "sec_6", "title": "6. Yêu Cầu Dữ Liệu & Từ Điển Dữ Liệu (Data Requirements)", "desc": "Cấu trúc thực thể ERD, danh mục trường dữ liệu và quy tắc ràng buộc nghiệp vụ."},
        {"id": "sec_7", "title": "7. Ma Trận Truy Xuất Nguồn Gốc (RTM - Requirements Traceability Matrix)", "desc": "Bảng ánh xạ từ Mục tiêu nghiệp vụ -> Mã yêu cầu -> Kịch bản kiểm thử UAT tương ứng."},
        {"id": "sec_8", "title": "8. Yêu Cầu Phi Chức Năng (Non-Functional Requirements - NFR)", "desc": "Quy định về hiệu năng TPS, thời gian phản hồi, bảo mật PCI-DSS, sao lưu và khôi phục sự cố."},
        {"id": "sec_9", "title": "9. Kế Hoạch Nghiệm Thu & Ký Duyệt (Acceptance Criteria & Approvals)", "desc": "Tiêu chí nghiệm thu tổng thể và chữ ký của Business Owner, Product Owner và Lead BA."}
    ]
}

