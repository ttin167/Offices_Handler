"""
build_qa_docx.py — Compiles 30 Mock Defense Q&As into an Academic Executive DOCX Document.
Tailored for Capstone Project:
"Design of an Autonomous Vacuum Cleaning Robot for Indoor Applications"
Student: Nguyen Vu Nguyen (MSSV: 2252948)
Supervisor: Dr. Pham Phuong Tung
Faculty of Mechanical Engineering - Mechatronics Department, HCMUT.
"""

from __future__ import annotations

import json
import os
import re
import sys
from typing import Any, Dict, List

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml

# ---------------------------------------------------------------------------
# COLOR PALETTE CONSTANTS
# ---------------------------------------------------------------------------
COLOR_PRIMARY_NAVY = RGBColor(0x0F, 0x29, 0x4A)   # #0F294A
COLOR_ACCENT_BLUE   = RGBColor(0x00, 0x51, 0xE2)   # #0051E2
COLOR_TEAL          = RGBColor(0x0D, 0x94, 0x88)   # #0D9488
COLOR_TEXT_DARK     = RGBColor(0x0F, 0x17, 0x2A)   # #0F172A
COLOR_TEXT_MUTED    = RGBColor(0x47, 0x55, 0x69)   # #475569
COLOR_TEXT_LIGHT    = RGBColor(0x64, 0x74, 0x8B)   # #64748B
COLOR_WHITE         = RGBColor(0xFF, 0xFF, 0xFF)   # #FFFFFF

HEX_PRIMARY_NAVY    = "0F294A"
HEX_ACCENT_BLUE     = "0051E2"
HEX_TEAL            = "0D9488"
HEX_BG_LIGHT        = "F8FAFC"
HEX_BG_CALLOUT      = "F8FAFC"
HEX_BG_AIM          = "F1F5F9"
HEX_BG_KEY          = "F0FDF4"
HEX_BORDER_LIGHT    = "E2E8F0"
HEX_BORDER_SLATE    = "CBD5E1"
HEX_BORDER_KEY      = "0D9488"


# ---------------------------------------------------------------------------
# OPENXML TABLE & CELL HELPERS
# ---------------------------------------------------------------------------
def set_cell_shading(tc_elem, hex_color: str) -> None:
    """Set cell background fill color."""
    tcPr = tc_elem.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color.lstrip("#").upper())
    tcPr.append(shd)


def set_cell_margins(tc_elem, top: int = 120, bottom: int = 120, left: int = 160, right: int = 160) -> None:
    """Set cell internal padding in dxa (1 pt = 20 dxa)."""
    tcPr = tc_elem.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(existing)
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_custom_cell_borders(tc_elem, top=None, bottom=None, left=None, right=None) -> None:
    """Set individual borders for a table cell."""
    tcPr = tc_elem.get_or_add_tcPr()
    for existing in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(existing)
    tcBorders = OxmlElement("w:tcBorders")
    
    borders = {"top": top, "bottom": bottom, "left": left, "right": right}
    for side, cfg in borders.items():
        b_elem = OxmlElement(f"w:{side}")
        if cfg:
            b_elem.set(qn("w:val"), cfg.get("val", "single"))
            b_elem.set(qn("w:sz"), str(cfg.get("sz", 4)))
            b_elem.set(qn("w:space"), "0")
            b_elem.set(qn("w:color"), cfg.get("color", "CBD5E1").lstrip("#"))
        else:
            b_elem.set(qn("w:val"), "none")
        tcBorders.append(b_elem)
    tcPr.append(tcBorders)


def set_row_flags(tr_elem, is_header: bool = False, cant_split: bool = True) -> None:
    """Enforce cantSplit and tblHeader flags for OpenXML tables."""
    trPr = tr_elem.get_or_add_trPr()
    if cant_split and trPr.find(qn("w:cantSplit")) is None:
        trPr.append(OxmlElement("w:cantSplit"))
    if is_header and trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def ensure_cell_has_paragraph(tc_elem) -> None:
    """OpenXML invariant ERR_DOCX_001: Every cell must terminate with at least one paragraph."""
    paragraphs = tc_elem.findall(qn("w:p"))
    if not paragraphs:
        p = OxmlElement("w:p")
        tc_elem.append(p)


def add_page_number_to_paragraph(p):
    """Inserts a native dynamic page number field into a Word paragraph."""
    run = p.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def clean_math_and_plain_text(text: str) -> str:
    """Ensures 100% clean readable text without LaTeX residues."""
    if not text:
        return ""
    replacements = [
        (r"\$\\rightarrow\$", "→"),
        (r"\\rightarrow", "→"),
        (r"\$\\approx\$", "≈"),
        (r"\\approx", "≈"),
        (r"\$\\ge\$", "≥"),
        (r"\\ge", "≥"),
        (r"\$\\le\$", "≤"),
        (r"\\le", "≤"),
        (r"\$\\pm\$", "±"),
        (r"\\pm", "±"),
        (r"\$\\degree\$", "°"),
        (r"\\degree", "°"),
        (r"\$\\cdot\$", "·"),
        (r"\\cdot", "·"),
        (r"\$\\Omega\$", "Ω"),
        (r"\\Omega", "Ω"),
        (r"\$", ""),  # remove loose dollar signs
    ]
    res = text
    for pat, rep in replacements:
        res = re.sub(pat, rep, res)
    return res


def add_markdown_formatted_text(p, text: str, default_font_size: Pt = Pt(10), default_color: RGBColor = COLOR_TEXT_DARK):
    """
    Parses basic bold formatting (**text**) and renders runs cleanly.
    """
    cleaned = clean_math_and_plain_text(text)
    # Split by **bold**
    tokens = re.split(r"(\*\*.*?\*\*)", cleaned)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**") and len(token) >= 4:
            bold_text = token[2:-2]
            run = p.add_run(bold_text)
            run.bold = True
            run.font.name = "Arial"
            run.font.size = default_font_size
            run.font.color.rgb = default_color
        else:
            run = p.add_run(token)
            run.font.name = "Arial"
            run.font.size = default_font_size
            run.font.color.rgb = default_color


# ---------------------------------------------------------------------------
# METADATA ENRICHMENT FOR 30 QUESTIONS
# ---------------------------------------------------------------------------
THEMATIC_CATEGORIES = {
    1: {
        "title": "CHUYÊN ĐỀ 1: THIẾT KẾ CƠ KHÍ, ĐỘNG LỰC HỌC & CƠ CẤU CHẤP HÀNH",
        "desc": "Bao gồm tính toán mô-men dẫn động, cấu hình bánh vi phân, bộ truyền đai chổi cuộn, cản va cơ học và hộp chứa bụi.",
        "badge": "Cơ khí & Động lực học",
        "badge_color": HEX_PRIMARY_NAVY,
        "questions": [1, 2, 21, 23, 24, 25]
    },
    2: {
        "title": "CHUYÊN ĐỀ 2: HỆ THỐNG ĐIỆN, NĂNG LƯỢNG PIN & AN TOÀN VẬN HÀNH",
        "desc": "Bao gồm phân bổ công suất, hệ thống pin LiFePO4, chống kẹt bảo vệ mô-tơ, chống nhiễu EMI và kiến trúc nguồn.",
        "badge": "Điện & Năng lượng",
        "badge_color": "C2410C",  # Orange-700
        "questions": [3, 4, 15, 16, 17, 19, 22, 26]
    },
    3: {
        "title": "CHUYÊN ĐỀ 3: KIẾN TRÚC NHÚNG, ĐIỀU KHIỂN TỰ ĐỘNG & THUẬT TOÁN ĐIỀU HƯỚNG",
        "desc": "Bao gồm Dual-MCU SPI, nhận dạng hàm nấc, điều khiển PI, Anti-windup, FSM, lọc trôi góc IMU và di chuyển xoắn ốc.",
        "badge": "Điều khiển & Nhúng",
        "badge_color": HEX_ACCENT_BLUE,
        "questions": [5, 6, 7, 8, 11, 12, 13, 14, 28, 29]
    },
    4: {
        "title": "CHUYÊN ĐỀ 4: THỰC NGHIỆM ĐO ĐẠC ĐỘ PHỦ, THU GOM RÁC & ĐÁNH GIÁ ĐỀ TÀI",
        "desc": "Bao gồm phương pháp đo độ phủ sàn IEC 62929, camera overhead, hiệu suất thu gom rác, phân tích sai số và chi phí.",
        "badge": "Thực nghiệm & Đóng góp",
        "badge_color": HEX_TEAL,
        "questions": [9, 10, 18, 20, 27, 30]
    }
}

QUESTION_EXTRA_METADATA = {
    1: {
        "ref": "Thuyết minh Chương 3 | Slide 8",
        "takeaways": [
            "Cấu hình 3 điểm tiếp xúc đảm bảo 2 bánh chủ động luôn bám chặt mặt sàn phẳng, không bao giờ bị kênh.",
            "Khả năng xoay tròn tại chỗ 360 độ (bán kính quay R = 0), tối ưu di chuyển trong gầm bàn ghế hẹp.",
            "Triệt tiêu lực cản của bánh caster khi đổi hướng bằng bộ điều khiển PID bù hướng IMU thời gian thực."
        ]
    },
    2: {
        "ref": "Thuyết minh Chương 3 | Slide 8",
        "takeaways": [
            "Khối lượng m = 4 kg, bán kính bánh r = 33 mm (0.033 m), gia tốc thiết kế a_max = 0.3 m/s².",
            "Mô-men yêu cầu danh định: 0.04 N·m (kết hợp lực quán tính 1.2 N và lực cản lăn 0.6 N).",
            "Hệ số an toàn k = 2.0 → chọn động cơ giảm tốc DC có mô-men định mức 0.08 – 0.1 N·m."
        ]
    },
    3: {
        "ref": "Thuyết minh Chương 4 | Slide 10 & 23",
        "takeaways": [
            "Phân chia độc lập: 4.7 W cho mạch logic/điều khiển và 93.44 W cho cơ cấu chấp hành (bánh xe, chổi cuộn, quạt hút).",
            "Hạ áp qua các module Buck DC-DC cách ly, ngăn chặn hiện tượng sụt áp vi điều khiển khi tải động lực khởi động.",
            "Cầu chì chính 10A bảo vệ ngắn mạch toàn hệ thống đặt ngay sau cực dương khối pin."
        ]
    },
    4: {
        "ref": "Thuyết minh Chương 4 | Slide 10",
        "takeaways": [
            "STM32F407 (Master, 168 MHz) chịu trách nhiệm FSM cấp cao, lập kế hoạch điều hướng và đọc cảm biến môi trường.",
            "STM32F103 (Slave, 72 MHz) chuyên trách vòng lặp kín PID điều khiển vận tốc 2 động cơ và đọc Encoder tần số cao.",
            "Giao tiếp qua bus SPI phần cứng 1 – 2 MHz có kiểm tra Checksum, đảm bảo tính thời gian thực chính xác."
        ]
    },
    5: {
        "ref": "Thuyết minh Chương 5 | Slide 14",
        "takeaways": [
            "Thực nghiệm đáp ứng hàm nấc (Step-Response): Cấp điện áp nấc 12V và đo đáp ứng tốc độ qua Encoder chu kỳ 10 ms.",
            "Mô hình hóa động cơ thành hàm truyền bậc 1: G(s) = K / (tau·s + 1), xác định chính xác hệ số tĩnh K và hằng số thời gian tau.",
            "Cơ sở khoa học vững chắc để tính toán bộ điều khiển PI giải tích thay vì dò sai thủ công."
        ]
    },
    6: {
        "ref": "Thuyết minh Chương 5 | Slide 14",
        "takeaways": [
            "Bộ điều khiển PI: Khâu P đảm bảo thời gian đáp ứng nhanh, khâu I triệt tiêu hoàn toàn sai số xác lập khi có tải cản lăn.",
            "Loại bỏ khâu D (Đạo hàm): Tín hiệu xung từ Encoder sau khi vi phân rất nhạy cảm với nhiễu (noise amplifier), làm PWM rung giật.",
            "Hệ thống đạt tính ổn định cao và đáp ứng trơn tru trên mọi loại bề mặt sàn."
        ]
    },
    7: {
        "ref": "Thuyết minh Chương 5 | Slide 14",
        "takeaways": [
            "Nguyên nhân: Khi xe khởi động hoặc gặp chướng ngại, sai số lớn làm thành phần tích phân tích lũy vượt quá 100% duty cycle (12V).",
            "Giải pháp Clamping (Khóa tích phân): Đóng băng khâu tích phân ngay khi tín hiệu điều khiển chạm ngưỡng bão hòa [0, 12V].",
            "Kết quả: Triệt tiêu hiện tượng vọt lố (overshoot) và rút ngắn thời gian ổn định tốc độ."
        ]
    },
    8: {
        "ref": "Thuyết minh Chương 6 | Slide 17",
        "takeaways": [
            "Cơ chế phát hiện kẹt: Bánh xe quay (Encoder có xung) nhưng gia tốc IMU = 0, hoặc Bumper bị kẹt quá 1.5 giây.",
            "Quy trình tự giải thoát 3 bước: (1) Lùi xe 10 cm → (2) Xoay ngẫu nhiên 90° – 135° → (3) Tiến lên thử lại.",
            "Giới hạn 3 chu kỳ lặp: Nếu sau 3 lần vẫn kẹt, robot chuyển trạng thái an toàn dừng khẩn cấp và phát còi báo."
        ]
    },
    9: {
        "ref": "Thuyết minh Chương 7 | Slide 4 & 19",
        "takeaways": [
            "Hệ thống đo đạc khách quan: Camera góc rộng gắn trần quay toàn bộ sàn thử nghiệm chuẩn.",
            "Xử lý ảnh chia lưới Occupancy Grid (kích thước ô 5x5 cm): Đếm số ô được đầu hút quét qua ít nhất một lần.",
            "Độ phủ đạt 96.30% trong phòng trống và 72.99% trong bài test phức tạp, tuân thủ tiêu chuẩn IEC 62929."
        ]
    },
    10: {
        "ref": "Thuyết minh Chương 7 | Slide 20",
        "takeaways": [
            "Hiệu suất thực nghiệm: Đạt 90.0% với rác nhẹ (giấy vụn) và 75.33% với bụi nặng/cát tiêu chuẩn.",
            "Phân tích nguyên nhân: Hạt cát có trọng lượng riêng lớn, cần áp suất tĩnh chân không cao hơn để hút ngược lên khoang chứa.",
            "Hướng cải tiến: Nâng cấp quạt hút BLDC áp suất cao và điều chỉnh góc gạt cao su chổi cuộn quét sát sàn."
        ]
    },
    11: {
        "ref": "Thuyết minh Chương 4 | Slide 9 & 17",
        "takeaways": [
            "3 cụm cảm biến phản xạ hồng ngoại mép sàn đo mức tín hiệu ánh sáng phản hồi từ bề mặt nền.",
            "Xử lý sàn đen/thảm tối: Hiệu chuẩn ngưỡng ADC phần mềm thực tế kết hợp bộ lọc thời gian Debounce (3 chu kỳ liên tiếp).",
            "Phanh khẩn cấp ngay khi phát hiện khoảng cách tăng đột biến, bảo vệ robot không bị rơi cầu thang."
        ]
    },
    12: {
        "ref": "Thuyết minh Chương 5 & 6 | Slide 16",
        "takeaways": [
            "Nguyên nhân: Con quay MEMS Gyro tích lũy sai số trôi điểm không (bias drift) theo thời gian tích phân góc.",
            "Quy trình Zero-Rate Calibration: Tự động đo và trừ độ lệch tĩnh trong 2 giây đầu khi robot đứng yên lúc khởi động.",
            "Kết hợp bộ lọc bổ sung (Complementary Filter) giúp góc hướng Yaw ổn định chính xác trong toàn bộ chu trình."
        ]
    },
    13: {
        "ref": "Thuyết minh Chương 6 | Slide 16 & 17",
        "takeaways": [
            "Khác biệt hoàn toàn với xe đồ chơi: Thuật toán phản xạ có chủ đích (Smart Reactive Navigation).",
            "Kết hợp 2 chiến thuật cốt lõi: Men theo tường (Wall-Following) để dọn chân tường và Xoắn ốc mở rộng (Spiral) dọn khoảng trống.",
            "Máy trạng thái FSM chuyển đổi mượt mà dựa trên dữ liệu 4 cảm biến khoảng cách và thanh cản va."
        ]
    },
    14: {
        "ref": "Thuyết minh Chương 6 | Slide 16",
        "takeaways": [
            "Quỹ đạo xoắn ốc Archimedes: Giữ tốc độ bánh ngoài cố định v_max, tăng dần vận tốc bánh trong theo thời gian để nới bán kính.",
            "Bán kính cong tăng đều từ tâm ra ngoài, đảm bảo vệt hút quét sạch sẽ không bỏ sót tâm phòng.",
            "Thoát chế độ xoắn ốc ngay khi phát hiện chướng ngại vật hoặc khi đạt bán kính giới hạn an toàn R_max = 0.8 m."
        ]
    },
    15: {
        "ref": "Thuyết minh Chương 4 | Slide 10 & 23",
        "takeaways": [
            "Điện trở Shunt đo dòng tải thực tế của mô-tơ chổi cuộn chính qua mạch khuếch đại dòng.",
            "Bảo vệ 2 tầng: Tầng phần mềm ngắt xung PWM khi dòng vượt ngưỡng 1.5A quá 200 ms; Tầng phần cứng bảo vệ chống ngắn mạch.",
            "Cơ chế tự đảo chiều ngắn để nhả dị vật (tóc, sợi vải) trước khi thử quay lại."
        ]
    },
    16: {
        "ref": "Thuyết minh Chương 4 | Slide 10",
        "takeaways": [
            "Bộ pin LiFePO4 4S dung lượng 3000 mAh (12.8V danh định, năng lượng lưu trữ 38.4 Wh).",
            "Công suất tiêu thụ trung bình thực tế khi dọn dẹp chỉ khoảng 25 – 30 W (tải không kẹt liên tục).",
            "Thời gian hoạt động liên tục đạt 75 phút, hoàn toàn đáp ứng chu trình làm sạch căn hộ gia đình 40 – 60 m²."
        ]
    },
    17: {
        "ref": "Thuyết minh Chương 4 | Slide 10",
        "takeaways": [
            "Tụ gốm 100 nF mắc song song dập hồ quang tia lửa điện trực tiếp tại 2 cực chổi than động cơ.",
            "Thiết kế mạch in 2 lớp phân tách độc lập Power Ground (dòng lớn động cơ) và Signal Ground (vi điều khiển).",
            "Điểm nối đất hình sao (Star Grounding) duy nhất tại cực âm pin loại bỏ hoàn toàn hiện tượng sụt mass gây reset vi điều khiển."
        ]
    },
    18: {
        "ref": "Thuyết minh Chương 7 | Slide 18",
        "takeaways": [
            "Thử nghiệm chạy thẳng 1 mét: Dán thước milimet chuẩn trên sàn, quay video từ trần đo độ lệch ngang (sai số dưới 5 mm = 0.5%).",
            "Thử nghiệm xoay góc 90° và 180°: Kiểm chứng qua bàn độ chia và góc quay tích phân của IMU.",
            "Minh chứng cho chất lượng đồng đều của bộ điều khiển PI cân bằng tốc độ hai bánh xe."
        ]
    },
    19: {
        "ref": "Thuyết minh Chương 4 | Slide 10",
        "takeaways": [
            "Robot hoạt động hoàn toàn tự hành (Autonomous Offline): ESP32-C3 chỉ đóng vai trò cầu nối truyền telemetry qua UART.",
            "Mất kết nối Bluetooth/Wi-Fi không ảnh hưởng đến thuật toán hút bụi và an toàn di chuyển của STM32 Master.",
            "Bộ đệm vòng tròn (Circular Buffer) lưu trữ log cục bộ và tự động đồng bộ lại khi kết nối phục hồi."
        ]
    },
    20: {
        "ref": "Thuyết minh Chương 7 | Slide 21",
        "takeaways": [
            "Tổng chi phí linh kiện và gia công mẫu thử (BOM Prototype): khoảng 2.5 – 2.8 triệu VNĐ.",
            "Tối ưu chi phí nhờ tự in 3D kết cấu khung vỏ, thiết kế mạch PCB riêng và sử dụng dòng chip nhúng công nghiệp phổ thông.",
            "Tiềm năng thương mại hóa lớn ở phân khúc robot hút bụi bình dân, dễ sửa chữa và thay thế linh kiện."
        ]
    },
    21: {
        "ref": "Thuyết minh Chương 3 | Slide 8",
        "takeaways": [
            "Bộ truyền đai răng (Timing Belt): Giảm chấn và bảo vệ mô-tơ khi chổi kẹt cứng, chống mẻ bánh răng.",
            "Bố trí mô-tơ lệch sang một bên giúp tối ưu không gian họng hút trung tâm và hạ thấp trọng tâm xe.",
            "Vận hành êm ái, không cần tra dầu mỡ bôi trơn (tránh bám bụi bẩn gây kẹt cơ khí)."
        ]
    },
    22: {
        "ref": "Thuyết minh Chương 4 | Slide 9",
        "takeaways": [
            "Cảm biến siêu âm có búp sóng rộng, dễ bị tán xạ ở góc hẹp và chân bàn tròn, tốc độ phản hồi chậm.",
            "Cảm biến khoảng cách hồng ngoại kích thước gọn gàng, tần số lấy mẫu cao (trên 100 Hz), phát hiện vật cản tầm gần rất nhạy.",
            "Chi phí thấp và dễ bố trí xung quanh chu vi thân tròn của robot."
        ]
    },
    23: {
        "ref": "Thuyết minh Chương 3 & 4 | Slide 8",
        "takeaways": [
            "Thanh cản va cơ học (Bumper) là chốt an toàn vật lý cuối cùng (Fail-safe) khi cảm biến quang học bị lóa hoặc gặp kính trong suốt.",
            "Sử dụng 2 công tắc hành trình (Microswitch) độ nhạy cao phân biệt va chạm cánh trái / cánh phải.",
            "Lò xo đàn hồi tạo hành trình tự do 5 – 8 mm hấp thụ xung lực va đập cơ học."
        ]
    },
    24: {
        "ref": "Thuyết minh Chương 5 | Slide 13",
        "takeaways": [
            "Mô hình động học vi phân: v_R = v + (omega·b)/2 và v_L = v - (omega·b)/2 với khoảng cách 2 bánh b = 232 mm (0.232 m).",
            "Quy đổi sang vận tốc góc bánh xe bằng cách chia cho bán kính bánh r = 33 mm.",
            "Cơ sở toán học cốt lõi để Slave MCU điều khiển bám vận tốc độc lập cho từng động cơ."
        ]
    },
    25: {
        "ref": "Thuyết minh Chương 3 | Slide 8",
        "takeaways": [
            "Dung tích hộp bụi 0.2 lít đáp ứng tiêu chuẩn dọn dẹp phòng gia đình 20 – 30 m² trong 1 chu trình.",
            "Màng lọc 2 tầng: Lưới lọc thô giữ rác lớn và màng HEPA H11 lọc 95% bụi mịn kích thước 0.3 micromet.",
            "Van silicone một chiều ở cửa nạp chống rơi ngược bụi rác khi tắt quạt hút."
        ]
    },
    26: {
        "ref": "Thuyết minh Chương 4 | Slide 10",
        "takeaways": [
            "Tuổi thọ vượt trội: Pin LiFePO4 đạt trên 2000 chu kỳ sạc-xả (gấp 4 lần so với 300 – 500 chu kỳ của Li-ion thông thường).",
            "Độ an toàn nhiệt tuyệt đối: Chống cháy nổ ngay cả khi quá sạc, nhiệt độ đánh thủng cấu trúc trên 270°C.",
            "Đường đặc tính phóng điện rất phẳng, giữ điện áp ổn định trong hầu hết thời gian làm việc."
        ]
    },
    27: {
        "ref": "Thuyết minh Chương 7 | Slide 19",
        "takeaways": [
            "Độ phủ giảm nhẹ 3.90% (từ 96.30% xuống 92.40%) khi đưa vào 2 vật cản kích thước 20x20 cm.",
            "Bản chất: Robot né tránh tạo vùng bóng râm (Shadow Zone) xung quanh vật cản, phản ánh đúng thực tế di chuyển.",
            "Khẳng định độ tin cậy và tính trung thực của phương pháp đo đạc thực nghiệm theo tiêu chuẩn IEC 62929."
        ]
    },
    28: {
        "ref": "Thuyết minh Chương 6 | Slide 15",
        "takeaways": [
            "Máy trạng thái FSM 5 trạng thái: Init (Khởi động), Clean (Dọn dẹp), Wall-Follow (Men tường), Avoid/Turn (Tránh vật cản), Stuck Recovery (Giải thoát).",
            "Chống treo Deadlock bằng cơ chế Timeout (Software Watchdog) cho từng trạng thái.",
            "Tự động kích hoạt chu trình thoát hiểm nếu robot bị kẹt trong một vòng lặp logic quá thời gian quy định."
        ]
    },
    29: {
        "ref": "Thuyết minh Chương 5 & 6 | Slide 16",
        "takeaways": [
            "Hai bánh xe quay ngược chiều nhau với cùng tốc độ để robot xoay tròn quanh tâm mà không bị tịnh tiến.",
            "Thuật toán Turn Angle giám sát liên tục góc xoay từ IMU, ngắt động cơ khi góc lệch nhỏ hơn ngưỡng 0.5°.",
            "Thời gian thực hiện xoay 90° chỉ mất khoảng 1.2 giây với độ chính xác cao."
        ]
    },
    30: {
        "ref": "Thuyết minh Toàn văn | Slide 21",
        "takeaways": [
            "Tự chủ 100% toàn bộ chuỗi công nghệ Cơ – Điện – Lập trình nhúng từ mô hình SolidWorks đến mạch in PCB và Firmware C/C++.",
            "Làm chủ thuật toán điều khiển kín PI động cơ và máy trạng thái FSM điều hướng phản xạ thông minh.",
            "Xây dựng bài toán thực nghiệm đo đạc độ phủ sàn và hiệu suất thu gom rác định lượng theo tiêu chuẩn quốc tế IEC 62929."
        ]
    }
}


def build_defense_qa_document(output_path: str, questions_data: List[Dict[str, Any]]):
    """Compiles all 30 questions into a publication-grade Word document."""
    doc = Document()
    
    # -----------------------------------------------------------------------
    # PAGE SETUP: Standard A4 with professional academic margins
    # -----------------------------------------------------------------------
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(1.8)
    
    # Configure Header & Footer
    header = section.header
    header_para = header.paragraphs[0]
    header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = header_para.add_run("TRƯỜNG ĐH BÁCH KHOA - ĐHQG TP.HCM | KHOA CƠ KHÍ — BỘ MÔN CƠ ĐIỆN TỬ")
    hrun.font.name = "Arial"
    hrun.font.size = Pt(8)
    hrun.font.color.rgb = COLOR_TEXT_LIGHT
    
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    frun1 = footer_para.add_run("ĐATN: Robot Hút Bụi Tự Hành | SV: Nguyễn Vũ Nguyên (2252948)           Trang ")
    frun1.font.name = "Arial"
    frun1.font.size = Pt(8.5)
    frun1.font.color.rgb = COLOR_TEXT_LIGHT
    add_page_number_to_paragraph(footer_para)

    # -----------------------------------------------------------------------
    # 1. ACADEMIC TITLE & METADATA HEADER
    # -----------------------------------------------------------------------
    p_inst = doc.add_paragraph()
    p_inst.paragraph_format.space_before = Pt(0)
    p_inst.paragraph_format.space_after = Pt(2)
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_inst.add_run("ĐẠI HỌC QUỐC GIA THÀNH PHỐ HỒ CHÍ MINH\nTRƯỜNG ĐẠI HỌC BÁCH KHOA\nKHOA CƠ KHÍ — BỘ MÔN CƠ ĐIỆN TỬ")
    r_inst.bold = True
    r_inst.font.name = "Arial"
    r_inst.font.size = Pt(10)
    r_inst.font.color.rgb = COLOR_TEXT_MUTED

    # Separator rule
    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_before = Pt(4)
    p_line.paragraph_format.space_after = Pt(12)
    r_line = p_line.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    r_line.font.name = "Arial"
    r_line.font.size = Pt(9)
    r_line.font.color.rgb = COLOR_ACCENT_BLUE

    # Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run("TÀI LIỆU ÔN TẬP & BỘ 30 CÂU HỎI PHẢN BIỆN MẪU\nBẢO VỆ ĐỒ ÁN TỐT NGHIỆP CƠ ĐIỆN TỬ")
    r_title.bold = True
    r_title.font.name = "Arial"
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = COLOR_PRIMARY_NAVY

    # Subtitle (Topic)
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(4)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run('Đề tài: "THIẾT KẾ ROBOT HÚT BỤI TỰ HÀNH CHO ỨNG DỤNG TRONG NHÀ"\n(Design of an Autonomous Vacuum Cleaning Robot for Indoor Applications)')
    r_sub.italic = True
    r_sub.bold = True
    r_sub.font.name = "Arial"
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = COLOR_TEXT_DARK

    # Metadata Summary Box (Table)
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    
    meta_data = [
        ("Sinh viên thực hiện:", "Nguyễn Vũ Nguyên — MSSV: 2252948"),
        ("Cán bộ hướng dẫn:", "TS. Phạm Phương Tùng"),
        ("Chuyên ngành đào tạo:", "Kỹ thuật Cơ điện tử (Khoa Cơ khí, ĐHBK TP.HCM)"),
        ("Tiêu chuẩn văn bản:", "100% Plain Text, Chuẩn hóa số liệu khớp Thuyết minh & Slide V6")
    ]
    
    for row_idx, (label, val) in enumerate(meta_data):
        row = meta_table.rows[row_idx]
        set_row_flags(row._tr, cant_split=True)
        
        c0 = row.cells[0]
        c0.width = Cm(5.2)
        set_cell_shading(c0._tc, HEX_BG_LIGHT)
        set_cell_margins(c0._tc, top=80, bottom=80, left=140, right=100)
        set_custom_cell_borders(c0._tc, 
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                left={"sz": 12, "color": HEX_PRIMARY_NAVY},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.name = "Arial"
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_TEXT_MUTED
        ensure_cell_has_paragraph(c0._tc)

        c1 = row.cells[1]
        c1.width = Cm(11.3)
        set_cell_shading(c1._tc, HEX_BG_LIGHT)
        set_cell_margins(c1._tc, top=80, bottom=80, left=120, right=140)
        set_custom_cell_borders(c1._tc,
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                left={"sz": 4, "color": HEX_BORDER_LIGHT},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(val)
        r1.bold = (row_idx < 2)
        r1.font.name = "Arial"
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_TEXT_DARK
        ensure_cell_has_paragraph(c1._tc)

    # -----------------------------------------------------------------------
    # 2. EXECUTIVE DEFENSE STRATEGY GUIDE
    # -----------------------------------------------------------------------
    p_strat_h = doc.add_paragraph()
    p_strat_h.paragraph_format.space_before = Pt(18)
    p_strat_h.paragraph_format.space_after = Pt(6)
    p_strat_h.paragraph_format.keep_with_next = True
    r_strat_h = p_strat_h.add_run("CẨM NANG TÁC PHONG & CHIẾN THUẬT PHẢN BIỆN TRƯỚC HỘI ĐỒNG")
    r_strat_h.bold = True
    r_strat_h.font.name = "Arial"
    r_strat_h.font.size = Pt(12)
    r_strat_h.font.color.rgb = COLOR_PRIMARY_NAVY

    strat_table = doc.add_table(rows=1, cols=1)
    strat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    strat_cell = strat_table.cell(0, 0)
    strat_cell.width = Cm(16.5)
    set_row_flags(strat_table.rows[0]._tr, cant_split=True)
    set_cell_shading(strat_cell._tc, HEX_BG_AIM)
    set_cell_margins(strat_cell._tc, top=140, bottom=140, left=180, right=160)
    set_custom_cell_borders(strat_cell._tc,
                            left={"sz": 24, "color": HEX_ACCENT_BLUE},
                            top={"sz": 4, "color": HEX_BORDER_LIGHT},
                            bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                            right={"sz": 4, "color": HEX_BORDER_LIGHT})
    
    ps = strat_cell.paragraphs[0]
    ps.paragraph_format.space_before = Pt(0)
    ps.paragraph_format.space_after = Pt(4)
    rs_head = ps.add_run("Quy tắc vàng 3 bước trả lời mọi câu hỏi kỹ thuật:")
    rs_head.bold = True
    rs_head.font.name = "Arial"
    rs_head.font.size = Pt(10)
    rs_head.font.color.rgb = COLOR_ACCENT_BLUE

    points = [
        ("Bước 1 — Khẳng định trực tiếp & gãy gọn:", " Dạ thưa Thầy Cô, [nêu thẳng phương án hoặc kết quả lựa chọn], ví dụ: Em chọn cấu hình 3 bánh vi phân... (Không nói vòng vo, không ngập ngừng)."),
        ("Bước 2 — Bản chất kỹ thuật & lý do lựa chọn:", " Phân tích cơ sở khoa học, nguyên lý vật lý hoặc sự đánh đổi (Trade-off): Vì sao chọn giải pháp này mà không chọn giải pháp khác."),
        ("Bước 3 — Số liệu thực chứng từ đồ án:", " Đưa ra con số cụ thể đo đạc được trong Thuyết minh/Slide (ví dụ: mô-men 0.04 N·m, độ phủ 96.30%, sai số thẳng 0.5%, thời gian chạy 75 phút). Số liệu thực tế là vũ khí mạnh nhất chứng minh sinh viên tự tay làm thật.")
    ]
    for b_title, b_content in points:
        p_pt = strat_cell.add_paragraph()
        p_pt.paragraph_format.space_before = Pt(2)
        p_pt.paragraph_format.space_after = Pt(2)
        r_pt1 = p_pt.add_run("• " + b_title)
        r_pt1.bold = True
        r_pt1.font.name = "Arial"
        r_pt1.font.size = Pt(9.5)
        r_pt1.font.color.rgb = COLOR_TEXT_DARK
        r_pt2 = p_pt.add_run(b_content)
        r_pt2.font.name = "Arial"
        r_pt2.font.size = Pt(9.5)
        r_pt2.font.color.rgb = COLOR_TEXT_MUTED
    ensure_cell_has_paragraph(strat_cell._tc)

    # -----------------------------------------------------------------------
    # 3. EXECUTIVE INDEX MATRIX TABLE (BẢNG TRA CỨU NHANH 30 CÂU)
    # -----------------------------------------------------------------------
    doc.add_page_break()
    p_idx_h = doc.add_paragraph()
    p_idx_h.paragraph_format.space_before = Pt(4)
    p_idx_h.paragraph_format.space_after = Pt(2)
    p_idx_h.paragraph_format.keep_with_next = True
    r_idx_h = p_idx_h.add_run("BẢNG TRA CỨU NHANH DANH MỤC 30 CÂU HỎI PHẢN BIỆN THEO CHUYÊN ĐỀ")
    r_idx_h.bold = True
    r_idx_h.font.name = "Arial"
    r_idx_h.font.size = Pt(11.5)
    r_idx_h.font.color.rgb = COLOR_PRIMARY_NAVY

    p_idx_sub = doc.add_paragraph()
    p_idx_sub.paragraph_format.space_before = Pt(0)
    p_idx_sub.paragraph_format.space_after = Pt(4)
    r_idx_sub = p_idx_sub.add_run("Toàn bộ 30 câu hỏi được phân loại theo 4 trục kiến trúc kỹ thuật Cơ – Điện – Điều khiển – Thực nghiệm:")
    r_idx_sub.italic = True
    r_idx_sub.font.name = "Arial"
    r_idx_sub.font.size = Pt(8.5)
    r_idx_sub.font.color.rgb = COLOR_TEXT_MUTED

    # Table: STT, Chủ đề, Câu hỏi tóm tắt, Chuyên đề, Tham chiếu
    idx_table = doc.add_table(rows=len(questions_data) + 1, cols=4)
    idx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    idx_table.autofit = False

    col_widths = [Cm(1.1), Cm(4.8), Cm(7.6), Cm(3.0)]

    # Header Row
    hdr_row = idx_table.rows[0]
    set_row_flags(hdr_row._tr, is_header=True, cant_split=True)
    headers = ["STT", "Chủ Đề Câu Hỏi", "Trọng Tâm Đánh Giá", "Tham Chiếu"]
    for i, h_text in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.width = col_widths[i]
        set_cell_shading(cell._tc, HEX_PRIMARY_NAVY)
        set_cell_margins(cell._tc, top=50, bottom=50, left=50, right=50)
        set_custom_cell_borders(cell._tc,
                                top={"sz": 4, "color": HEX_PRIMARY_NAVY},
                                bottom={"sz": 10, "color": HEX_ACCENT_BLUE},
                                left={"sz": 4, "color": "334155"},
                                right={"sz": 4, "color": "334155"})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h_text)
        r.bold = True
        r.font.name = "Arial"
        r.font.size = Pt(8)
        r.font.color.rgb = COLOR_WHITE
        ensure_cell_has_paragraph(cell._tc)

    # Question rows
    for idx, q in enumerate(questions_data):
        row = idx_table.rows[idx + 1]
        set_row_flags(row._tr, is_header=False, cant_split=True)
        q_num = q["num"]
        meta = QUESTION_EXTRA_METADATA.get(q_num, {"ref": "Thuyết minh & Slide", "takeaways": []})
        
        # Determine category badge
        cat_badge = ""
        for c_id, c_info in THEMATIC_CATEGORIES.items():
            if q_num in c_info["questions"]:
                cat_badge = f"CĐ {c_id}"
                break
        
        bg_color = HEX_BG_LIGHT if (idx % 2 == 1) else "FFFFFF"

        # C0: STT
        c0 = row.cells[0]
        c0.width = col_widths[0]
        set_cell_shading(c0._tc, bg_color)
        set_cell_margins(c0._tc, top=25, bottom=25, left=30, right=30)
        set_custom_cell_borders(c0._tc,
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                left={"sz": 4, "color": HEX_BORDER_LIGHT},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(f"{q_num:02d}")
        r0.bold = True
        r0.font.name = "Arial"
        r0.font.size = Pt(7.5)
        r0.font.color.rgb = COLOR_PRIMARY_NAVY
        ensure_cell_has_paragraph(c0._tc)

        # C1: Chủ đề & Chuyên đề
        c1 = row.cells[1]
        c1.width = col_widths[1]
        set_cell_shading(c1._tc, bg_color)
        set_cell_margins(c1._tc, top=25, bottom=25, left=50, right=50)
        set_custom_cell_borders(c1._tc,
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                left={"sz": 4, "color": HEX_BORDER_LIGHT},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1_cat = p1.add_run(f"[{cat_badge}] ")
        r1_cat.bold = True
        r1_cat.font.name = "Arial"
        r1_cat.font.size = Pt(7.5)
        r1_cat.font.color.rgb = COLOR_ACCENT_BLUE
        
        r1 = p1.add_run(q["topic"])
        r1.bold = True
        r1.font.name = "Arial"
        r1.font.size = Pt(7.5)
        r1.font.color.rgb = COLOR_TEXT_DARK
        ensure_cell_has_paragraph(c1._tc)

        # C2: Trọng tâm / Ý đồ
        c2 = row.cells[2]
        c2.width = col_widths[2]
        set_cell_shading(c2._tc, bg_color)
        set_cell_margins(c2._tc, top=25, bottom=25, left=50, right=50)
        set_custom_cell_borders(c2._tc,
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                left={"sz": 4, "color": HEX_BORDER_LIGHT},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        aim_snippet = q["aim"].replace("\n", " ")
        if len(aim_snippet) > 85:
            aim_snippet = aim_snippet[:82] + "..."
        r2 = p2.add_run(aim_snippet)
        r2.font.name = "Arial"
        r2.font.size = Pt(7.5)
        r2.font.color.rgb = COLOR_TEXT_MUTED
        ensure_cell_has_paragraph(c2._tc)

        # C3: Tham chiếu
        c3 = row.cells[3]
        c3.width = col_widths[3]
        set_cell_shading(c3._tc, bg_color)
        set_cell_margins(c3._tc, top=25, bottom=25, left=40, right=40)
        set_custom_cell_borders(c3._tc,
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                left={"sz": 4, "color": HEX_BORDER_LIGHT},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        p3 = c3.paragraphs[0]
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(0)
        r3 = p3.add_run(meta["ref"])
        r3.italic = True
        r3.font.name = "Arial"
        r3.font.size = Pt(7.5)
        r3.font.color.rgb = COLOR_TEXT_LIGHT
        ensure_cell_has_paragraph(c3._tc)

    # -----------------------------------------------------------------------
    # 4. DETAILED Q&A SECTIONS (30 QUESTIONS IN HIGH-FIDELITY CALLOUT BLOCKS)
    # -----------------------------------------------------------------------
    doc.add_page_break()

    # We will iterate questions sequentially (1 -> 30)
    # And insert Thematic Section Dividers when transitioning into a new thematic group
    current_category_id = None

    for q in questions_data:
        q_num = q["num"]
        meta = QUESTION_EXTRA_METADATA.get(q_num, {"ref": "Thuyết minh & Slide", "takeaways": []})
        
        # Check which category this question belongs to
        q_cat_id = None
        for c_id, c_info in THEMATIC_CATEGORIES.items():
            if q_num in c_info["questions"]:
                q_cat_id = c_id
                break

        # Check if we should insert Section Banner
        if q_cat_id != current_category_id and q_cat_id is not None:
            current_category_id = q_cat_id
            c_info = THEMATIC_CATEGORIES[q_cat_id]
            
            p_sec = doc.add_paragraph()
            p_sec.paragraph_format.space_before = Pt(20)
            p_sec.paragraph_format.space_after = Pt(4)
            p_sec.paragraph_format.keep_with_next = True
            
            r_sec = p_sec.add_run(c_info["title"])
            r_sec.bold = True
            r_sec.font.name = "Arial"
            r_sec.font.size = Pt(13)
            r_sec.font.color.rgb = COLOR_PRIMARY_NAVY
            
            p_sec_desc = doc.add_paragraph()
            p_sec_desc.paragraph_format.space_before = Pt(0)
            p_sec_desc.paragraph_format.space_after = Pt(12)
            p_sec_desc.paragraph_format.keep_with_next = True
            r_sdesc = p_sec_desc.add_run(c_info["desc"])
            r_sdesc.italic = True
            r_sdesc.font.name = "Arial"
            r_sdesc.font.size = Pt(9.5)
            r_sdesc.font.color.rgb = COLOR_TEXT_MUTED

        # --- QUESTION HEADING ---
        p_qh = doc.add_paragraph()
        p_qh.paragraph_format.space_before = Pt(16)
        p_qh.paragraph_format.space_after = Pt(3)
        p_qh.paragraph_format.keep_with_next = True
        
        # Category tag
        badge_text = THEMATIC_CATEGORIES[q_cat_id]["badge"] if q_cat_id else "Chuyên đề"
        r_tag = p_qh.add_run(f"[{badge_text}]  ")
        r_tag.bold = True
        r_tag.font.name = "Arial"
        r_tag.font.size = Pt(9)
        r_tag.font.color.rgb = COLOR_ACCENT_BLUE

        # Question title
        r_qtitle = p_qh.add_run(f"CÂU {q_num:02d}: {q['topic'].upper()}")
        r_qtitle.bold = True
        r_qtitle.font.name = "Arial"
        r_qtitle.font.size = Pt(12)
        r_qtitle.font.color.rgb = COLOR_PRIMARY_NAVY

        # Reference subtext
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.space_before = Pt(0)
        p_ref.paragraph_format.space_after = Pt(6)
        p_ref.paragraph_format.keep_with_next = True
        r_ref = p_ref.add_run(f"📍 Tham chiếu tài liệu: {meta['ref']}")
        r_ref.italic = True
        r_ref.font.name = "Arial"
        r_ref.font.size = Pt(8.5)
        r_ref.font.color.rgb = COLOR_TEXT_LIGHT

        # --- QUESTION TEXT BOX ---
        q_box = doc.add_table(rows=1, cols=1)
        q_box.alignment = WD_TABLE_ALIGNMENT.CENTER
        q_cell = q_box.cell(0, 0)
        q_cell.width = Cm(16.5)
        set_row_flags(q_box.rows[0]._tr, cant_split=True)
        set_cell_shading(q_cell._tc, HEX_BG_AIM)
        set_cell_margins(q_cell._tc, top=100, bottom=100, left=160, right=140)
        set_custom_cell_borders(q_cell._tc,
                                left={"sz": 20, "color": HEX_PRIMARY_NAVY},
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        
        pq = q_cell.paragraphs[0]
        pq.paragraph_format.space_before = Pt(0)
        pq.paragraph_format.space_after = Pt(0)
        rq_lbl = pq.add_run("Hội đồng hỏi: ")
        rq_lbl.bold = True
        rq_lbl.font.name = "Arial"
        rq_lbl.font.size = Pt(10)
        rq_lbl.font.color.rgb = COLOR_PRIMARY_NAVY
        
        clean_q_text = clean_math_and_plain_text(q["question"]).replace("\n", " ").strip()
        rq_txt = pq.add_run(f'"{clean_q_text}"')
        rq_txt.bold = False
        rq_txt.font.name = "Arial"
        rq_txt.font.size = Pt(10)
        rq_txt.font.color.rgb = COLOR_TEXT_DARK
        ensure_cell_has_paragraph(q_cell._tc)

        # --- AIM / Ý ĐỒ CỦA HỘI ĐỒNG ---
        p_aim = doc.add_paragraph()
        p_aim.paragraph_format.space_before = Pt(6)
        p_aim.paragraph_format.space_after = Pt(8)
        p_aim.paragraph_format.keep_with_next = True
        r_aim_lbl = p_aim.add_run("🎯 Ý đồ đánh giá của Thầy Cô: ")
        r_aim_lbl.bold = True
        r_aim_lbl.font.name = "Arial"
        r_aim_lbl.font.size = Pt(9)
        r_aim_lbl.font.color.rgb = COLOR_TEXT_MUTED
        
        clean_aim_text = clean_math_and_plain_text(q["aim"]).replace("\n", " ").strip()
        r_aim_txt = p_aim.add_run(clean_aim_text)
        r_aim_txt.italic = True
        r_aim_txt.font.name = "Arial"
        r_aim_txt.font.size = Pt(9)
        r_aim_txt.font.color.rgb = COLOR_TEXT_MUTED

        # --- SAMPLE ANSWER CALLOUT BOX ---
        ans_table = doc.add_table(rows=1, cols=1)
        ans_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        ans_cell = ans_table.cell(0, 0)
        ans_cell.width = Cm(16.5)
        set_row_flags(ans_table.rows[0]._tr, cant_split=True)
        set_cell_shading(ans_cell._tc, HEX_BG_CALLOUT)
        set_cell_margins(ans_cell._tc, top=140, bottom=140, left=180, right=160)
        set_custom_cell_borders(ans_cell._tc,
                                left={"sz": 32, "color": HEX_ACCENT_BLUE},
                                top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                right={"sz": 4, "color": HEX_BORDER_LIGHT})
        
        # Header inside answer callout
        pa_hdr = ans_cell.paragraphs[0]
        pa_hdr.paragraph_format.space_before = Pt(0)
        pa_hdr.paragraph_format.space_after = Pt(6)
        ra_hlbl = pa_hdr.add_run("💬 Câu trả lời mẫu của Sinh viên:")
        ra_hlbl.bold = True
        ra_hlbl.font.name = "Arial"
        ra_hlbl.font.size = Pt(10)
        ra_hlbl.font.color.rgb = COLOR_ACCENT_BLUE

        # Answer body paragraphs
        raw_ans = q["answer"].strip()
        ans_paragraphs = [p.strip() for p in raw_ans.split("\n\n") if p.strip()]
        
        for pa_idx, para_str in enumerate(ans_paragraphs):
            p_ans = ans_cell.add_paragraph()
            p_ans.paragraph_format.space_before = Pt(2)
            p_ans.paragraph_format.space_after = Pt(4)
            p_ans.paragraph_format.line_spacing = 1.2
            add_markdown_formatted_text(p_ans, para_str, default_font_size=Pt(9.5), default_color=COLOR_TEXT_DARK)

        ensure_cell_has_paragraph(ans_cell._tc)

        # --- KEY TAKEAWAYS (TỪ KHÓA GHI ĐIỂM) ---
        if meta.get("takeaways"):
            p_take_hdr = doc.add_paragraph()
            p_take_hdr.paragraph_format.space_before = Pt(6)
            p_take_hdr.paragraph_format.space_after = Pt(2)
            p_take_hdr.paragraph_format.keep_with_next = True
            r_tk_lbl = p_take_hdr.add_run("⚡ Từ khóa ghi điểm & Lưu ý thực chiến:")
            r_tk_lbl.bold = True
            r_tk_lbl.font.name = "Arial"
            r_tk_lbl.font.size = Pt(9)
            r_tk_lbl.font.color.rgb = COLOR_TEAL

            take_table = doc.add_table(rows=1, cols=1)
            take_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            take_cell = take_table.cell(0, 0)
            take_cell.width = Cm(16.5)
            set_row_flags(take_table.rows[0]._tr, cant_split=True)
            set_cell_shading(take_cell._tc, HEX_BG_KEY)
            set_cell_margins(take_cell._tc, top=80, bottom=80, left=140, right=140)
            set_custom_cell_borders(take_cell._tc,
                                    left={"sz": 20, "color": HEX_TEAL},
                                    top={"sz": 4, "color": HEX_BORDER_LIGHT},
                                    bottom={"sz": 4, "color": HEX_BORDER_LIGHT},
                                    right={"sz": 4, "color": HEX_BORDER_LIGHT})
            
            for t_idx, takeaway in enumerate(meta["takeaways"]):
                pt = take_cell.paragraphs[0] if t_idx == 0 else take_cell.add_paragraph()
                pt.paragraph_format.space_before = Pt(1)
                pt.paragraph_format.space_after = Pt(1)
                rt_bullet = pt.add_run("✔ ")
                rt_bullet.bold = True
                rt_bullet.font.name = "Arial"
                rt_bullet.font.size = Pt(8.5)
                rt_bullet.font.color.rgb = COLOR_TEAL
                
                rt_txt = pt.add_run(takeaway)
                rt_txt.font.name = "Arial"
                rt_txt.font.size = Pt(8.5)
                rt_txt.font.color.rgb = COLOR_TEXT_DARK
            ensure_cell_has_paragraph(take_cell._tc)

        # Spacing divider between questions
        p_div = doc.add_paragraph()
        p_div.paragraph_format.space_before = Pt(4)
        p_div.paragraph_format.space_after = Pt(10)
        r_div = p_div.add_run("──────────────────────────────────────────")
        r_div.font.name = "Arial"
        r_div.font.size = Pt(7)
        r_div.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)

    # -----------------------------------------------------------------------
    # 5. CONCLUSION & FINAL ADVICE
    # -----------------------------------------------------------------------
    doc.add_page_break()
    p_concl_h = doc.add_paragraph()
    p_concl_h.paragraph_format.space_before = Pt(12)
    p_concl_h.paragraph_format.space_after = Pt(6)
    r_ch = p_concl_h.add_run("LỜI KẾT & TÂM THẾ BẢO VỆ ĐỒ ÁN TỐT NGHIỆP THÀNH CÔNG")
    r_ch.bold = True
    r_ch.font.name = "Arial"
    r_ch.font.size = Pt(13)
    r_ch.font.color.rgb = COLOR_PRIMARY_NAVY

    final_tips = [
        ("Nắm chắc bức tranh tổng thể Cơ - Điện - Lập trình:", " Đồ án tốt nghiệp ngành Cơ điện tử đòi hỏi sự kết nối chặt chẽ giữa 3 phân hệ. Khi trả lời về phần cứng cơ khí, hãy luôn liên hệ sang phần điều khiển nhúng (như cách bánh caster ảnh hưởng đến bù hướng IMU)."),
        ("Tự tin vào kết quả đo đạc thực tế:", " Mọi con số trong Thuyết minh (96.30% độ phủ, 75 phút thời gian chạy, 0.5% sai số thẳng, 0.04 N·m mô-men) đều là minh chứng thuyết phục nhất. Thầy Cô đánh giá cao sự trung thực khoa học hơn là những con số lý thuyết hoàn hảo không tì vết."),
        ("Bình tĩnh lắng nghe hết câu hỏi:", " Không ngắt lời Thầy Cô. Dành 2 - 3 giây định hình câu trả lời theo công thức 3 bước. Nếu câu hỏi vượt ngoài phạm vi đề tài, hãy thẳng thắn ghi nhận đó là hướng phát triển tương lai rất có giá trị."),
        ("Chúc bạn bảo vệ Đồ án Tốt nghiệp đạt kết quả xuất sắc nhất và hoàn thành trọn vẹn bậc học Kỹ sư Cơ Điện Tử ĐHBK TP.HCM!")
    ]

    for f_title, f_desc in [
        (final_tips[0][0], final_tips[0][1]),
        (final_tips[1][0], final_tips[1][1]),
        (final_tips[2][0], final_tips[2][1]),
        ("", final_tips[3][0])
    ]:
        pf = doc.add_paragraph()
        pf.paragraph_format.space_before = Pt(4)
        pf.paragraph_format.space_after = Pt(4)
        if f_title:
            rf1 = pf.add_run("⭐ " + f_title)
            rf1.bold = True
            rf1.font.name = "Arial"
            rf1.font.size = Pt(10)
            rf1.font.color.rgb = COLOR_PRIMARY_NAVY
            rf2 = pf.add_run(f_desc)
            rf2.font.name = "Arial"
            rf2.font.size = Pt(10)
            rf2.font.color.rgb = COLOR_TEXT_DARK
        else:
            rf = pf.add_run(f_desc)
            rf.bold = True
            rf.italic = True
            rf.font.name = "Arial"
            rf.font.size = Pt(10.5)
            rf.font.color.rgb = COLOR_ACCENT_BLUE

    # Ensure parent directory exists and save
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    print(f"Successfully generated high-fidelity Word document at:\n{output_path}")


if __name__ == "__main__":
    json_path = "scratch/all_30_questions.json"
    if not os.path.isfile(json_path):
        print(f"Error: {json_path} does not exist.")
        sys.exit(1)
        
    with open(json_path, "r", encoding="utf-8") as f:
        questions_data = json.load(f)
        
    out_docx = os.path.abspath("docs/Bo_30_Cau_Hoi_Phan_Bien_Va_Tra_Loi_Mau_DATN.docx")
    build_defense_qa_document(out_docx, questions_data)
