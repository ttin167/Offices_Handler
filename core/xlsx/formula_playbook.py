# -*- coding: utf-8 -*-
"""
core.xlsx.formula_playbook — Enterprise Excel Formula Playbook & Cheat Sheet System.

Provides a comprehensive curated catalog of 32+ essential Excel formulas structured across:
1. Tên hàm (Function Name & Category)
2. Công thức mẫu (Syntax & Real-World Practical Examples)
3. Dùng trong trường hợp nào (When to Use / Business Scenarios)
4. Tại sao lại chọn hàm đó (Why Choose This Function / Technical Rationale & Alternatives Comparison)

Includes automated export to dual-sheet Excel workbook (.xlsx):
- Sheet 1: 'Formula_Playbook' (Enterprise Reference Guide)
- Sheet 2: 'Live_Interactive_Formulas' (Executable KPI Formulas & Business Demo)
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Tuple

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    openpyxl = None  # type: ignore

try:
    from core.xlsx.writer import apply_worksheet_print_setup
except ImportError:
    apply_worksheet_print_setup = None  # type: ignore


# ---------------------------------------------------------------------------
# Visual Tokens & Styles
# ---------------------------------------------------------------------------

NAVY_TITLE_FILL = PatternFill(start_color="0F294A", end_color="0F294A", fill_type="solid") if openpyxl else None
NAVY_HEADER_FILL = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") if openpyxl else None
SUBHEADER_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid") if openpyxl else None
ALT_ROW_FILL = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") if openpyxl else None
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") if openpyxl else None
HIGHLIGHT_KPI_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") if openpyxl else None

TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFF") if openpyxl else None
SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="475569") if openpyxl else None
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF") if openpyxl else None
DATA_FONT = Font(name="Calibri", size=10) if openpyxl else None
DATA_BOLD_FONT = Font(name="Calibri", size=10, bold=True) if openpyxl else None
CODE_FONT = Font(name="Consolas", size=9.5, color="0F172A") if openpyxl else None
KPI_VAL_FONT = Font(name="Calibri", size=11, bold=True, color="0F294A") if openpyxl else None

THIN_BORDER = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
) if openpyxl else None

HEADER_BORDER = Border(
    left=Side(style="thin", color="3B82F6"),
    right=Side(style="thin", color="3B82F6"),
    top=Side(style="medium", color="1D4ED8"),
    bottom=Side(style="medium", color="1D4ED8"),
) if openpyxl else None


# ---------------------------------------------------------------------------
# Comprehensive Master Catalog (32 Curated Formulas)
# ---------------------------------------------------------------------------

FORMULA_CATALOG: List[Dict[str, Any]] = [
    # -----------------------------------------------------------------------
    # 1. TRA CỨU & THAM CHIẾU (LOOKUP & REFERENCE)
    # -----------------------------------------------------------------------
    {
        "id": "xlookup",
        "name": "XLOOKUP",
        "category": "lookup",
        "category_vi": "Tra Cứu & Tham Chiếu",
        "template": "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
        "example": '=XLOOKUP("NV001", A2:A100, C2:C100, "Không tìm thấy", 0)',
        "when_to_use": "Dùng khi cần tìm kiếm một giá trị trong bảng dữ liệu và lấy dữ liệu tương ứng ở một cột hoặc hàng khác. Thích hợp cho tra cứu mã nhân viên, mã sản phẩm, tỷ giá hối đoái, tra cứu từ phải sang trái hoặc tra cứu 2 chiều.",
        "why_choose": "Ưu việt vượt trội hơn VLOOKUP ở 4 điểm: 1) Tra cứu được từ phải qua trái (VLOOKUP chỉ tìm cột bên phải); 2) Không bị vỡ công thức khi chèn thêm hay xóa bớt cột; 3) Tích hợp sẵn giá trị mặc định [if_not_found] không cần lồng IFERROR; 4) Mặc định tìm kiếm chính xác (match_mode=0) an toàn tuyệt đối.",
    },
    {
        "id": "index_match",
        "name": "INDEX / MATCH",
        "category": "lookup",
        "category_vi": "Tra Cứu & Tham Chiếu",
        "template": "=INDEX(return_range, MATCH(lookup_val, lookup_range, 0))",
        "example": '=INDEX(C2:C50000, MATCH("SP_A", A2:A50000, 0))',
        "when_to_use": "Dùng khi làm việc với các bảng dữ liệu khổng lồ (hàng trăm ngàn dòng) hoặc cần tương thích với mọi phiên bản Excel cũ (Excel 2010/2013/2016/2019 không có XLOOKUP).",
        "why_choose": "Tốc độ xử lý (benchmark hiệu năng) nhanh hơn VLOOKUP 30-50% trên tập dữ liệu lớn vì MATCH chỉ quét 1 cột duy nhất thay vì load toàn bộ mảng bảng vào bộ nhớ. Cho phép tra cứu linh hoạt 2 chiều (Giao điểm Hàng x Cột) với cú pháp INDEX(vùng, MATCH_hàng, MATCH_cột).",
    },
    {
        "id": "filter",
        "name": "FILTER",
        "category": "lookup",
        "category_vi": "Tra Cứu & Tham Chiếu",
        "template": "=FILTER(array, include, [if_empty])",
        "example": '=FILTER(A2:E100, (C2:C100="Hà Nội") * (E2:E100>50000000), "Không có dữ liệu")',
        "when_to_use": "Dùng khi cần lọc và bóc tách một danh sách thỏa mãn một hoặc nhiều điều kiện đồng thời sang một vùng báo cáo riêng mà không cần dùng công cụ Data Filter thủ công hay viết code VBA.",
        "why_choose": "Khác với XLOOKUP chỉ trả về 1 dòng đầu tiên, FILTER là hàm mảng động (Dynamic Array Spill) tự động trả về toàn bộ các dòng thỏa mãn điều kiện. Hỗ trợ toán tử nhân (*) cho điều kiện VÀ (AND) và toán tử cộng (+) cho điều kiện HOẶC (OR).",
    },
    {
        "id": "chooserows_choosecols",
        "name": "CHOOSECOLS / CHOOSEROWS",
        "category": "lookup",
        "category_vi": "Tra Cứu & Tham Chiếu",
        "template": "=CHOOSECOLS(array, col_num1, [col_num2], ...)",
        "example": "=CHOOSECOLS(FILTER(A2:Z100, C2:C100=\"Miền Nam\"), 1, 3, 5)",
        "when_to_use": "Dùng khi trích xuất báo cáo từ bảng dữ liệu gốc có quá nhiều cột không cần thiết, chỉ muốn giữ lại các cột cốt lõi (ví dụ cột 1: Mã, cột 3: Tên, cột 5: Doanh số).",
        "why_choose": "Loại bỏ hoàn toàn việc phải ẩn cột hoặc copy-paste thủ công. Kết hợp hoàn hảo với FILTER, SORT để tạo bảng điều khiển (Dashboard) tự động mà không làm biến đổi cấu trúc bảng nguồn.",
    },
    {
        "id": "offset",
        "name": "OFFSET",
        "category": "lookup",
        "category_vi": "Tra Cứu & Tham Chiếu",
        "template": "=OFFSET(reference, rows, cols, [height], [width])",
        "example": "=OFFSET(B1, 1, 0, COUNTA(B:B)-1, 1)",
        "when_to_use": "Dùng để tạo Dải ô động (Dynamic Named Range) cho đồ thị (Charts) hoặc danh sách thả xuống (Data Validation List) tự động co giãn khi người dùng nhập thêm dữ liệu mới.",
        "why_choose": "Cho phép dịch chuyển tọa độ linh hoạt theo cả trục ngang và trục dọc. Lưu ý: Do là hàm Volatile (tính lại mỗi khi có bất kỳ thay đổi nào trong bảng tính), nên dùng có kiểm soát hoặc thay bằng INDEX nếu file quá nặng.",
    },
    {
        "id": "indirect",
        "name": "INDIRECT",
        "category": "lookup",
        "category_vi": "Tra Cứu & Tham Chiếu",
        "template": '=INDIRECT(ref_text, [a1])',
        "example": '=INDIRECT("\'" & A2 & "\'!B5")',
        "when_to_use": "Dùng khi cần tham chiếu động đến tên một trang tính (Sheet) dựa trên giá trị của một ô (ví dụ: ô A2 ghi tên tháng 'Thang_05', công thức tự động kéo số liệu từ ô B5 của Sheet 'Thang_05').",
        "why_choose": "Giải quyết bài toán tổng hợp đa sheet mà không cần gõ cứng công thức ='Sheet1'!B5, ='Sheet2'!B5 cho từng dòng. Rất hữu ích khi xây dựng mô hình tài chính hợp nhất nhiều chi nhánh.",
    },
    {
        "id": "vlookup_legacy",
        "name": "VLOOKUP (Legacy Mode)",
        "category": "lookup",
        "category_vi": "Tra Cứu & Tham Chiếu",
        "template": "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
        "example": '=VLOOKUP("SP01", A2:D100, 3, FALSE)',
        "when_to_use": "Dùng khi bắt buộc phải duy trì tính tương thích với các hệ thống cổ hoặc tệp Excel xuất cho đối tác đang dùng phiên bản Excel 2003-2007 cũ kỹ.",
        "why_choose": "Hàm phổ biến nhất lịch sử Excel, ai cũng hiểu. Tuy nhiên ngày nay NÊN THAY THẾ bằng XLOOKUP hoặc INDEX/MATCH vì VLOOKUP dễ bị lỗi khi chèn thêm cột và không tra cứu ngược từ phải qua trái được.",
    },

    # -----------------------------------------------------------------------
    # 2. TOÁN HỌC & THỐNG KÊ TỔNG HỢP (AGGREGATION & MATH)
    # -----------------------------------------------------------------------
    {
        "id": "sumifs",
        "name": "SUMIFS",
        "category": "aggregation",
        "category_vi": "Toán Học & Thống Kê Tổng Hợp",
        "template": "=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
        "example": '=SUMIFS(E2:E100, B2:B100, "Kinh doanh", D2:D100, ">=2026-01-01")',
        "when_to_use": "Dùng khi cần tính tổng một cột số liệu dựa trên từ 2 điều kiện trở lên (ví dụ: Tổng doanh thu của phòng Kinh doanh trong năm 2026 tại thị trường Miền Bắc).",
        "why_choose": "Tốc độ xử lý cực nhanh, cú pháp trực quan hơn SUMPRODUCT. Luôn ưu tiên dùng SUMIFS ngay cả khi chỉ có 1 điều kiện (thay cho SUMIF) để thống nhất vị trí sum_range ở đầu, dễ mở rộng thêm tiêu chí sau này.",
    },
    {
        "id": "countifs",
        "name": "COUNTIFS",
        "category": "aggregation",
        "category_vi": "Toán Học & Thống Kê Tổng Hợp",
        "template": "=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
        "example": '=COUNTIFS(C2:C100, "Nam", E2:E100, ">=50000000")',
        "when_to_use": "Dùng khi cần đếm số lượng giao dịch, số nhân sự, hoặc số hợp đồng thỏa mãn đồng thời nhiều tiêu chí lọc.",
        "why_choose": "Là công cụ đếm đa điều kiện chuẩn mực nhất của Excel. Xử lý tốt các toán tử so sánh chuỗi ('>=', '<>', '*', '?') mà không cần viết công thức mảng phức tạp.",
    },
    {
        "id": "averageifs",
        "name": "AVERAGEIFS",
        "category": "aggregation",
        "category_vi": "Toán Học & Thống Kê Tổng Hợp",
        "template": "=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)",
        "example": '=AVERAGEIFS(F2:F100, B2:B100, "CNTT", D2:D100, "Senior")',
        "when_to_use": "Dùng để tính mức lương bình quân, giá bán trung bình, hoặc điểm số trung bình của một nhóm đối tượng cụ thể thỏa mãn nhiều tiêu chuẩn.",
        "why_choose": "Tự động loại trừ các ô rỗng và chỉ tính trung bình các dòng thỏa mãn tiêu chí, tránh sai lệch do chia cho 0 hoặc tính cả ô trống.",
    },
    {
        "id": "aggregate",
        "name": "AGGREGATE",
        "category": "aggregation",
        "category_vi": "Toán Học & Thống Kê Tổng Hợp",
        "template": "=AGGREGATE(function_num, options, array, [k])",
        "example": "=AGGREGATE(9, 6, C2:C100)",
        "when_to_use": "Dùng khi tính tổng (SUM) hoặc trung bình trong bảng dữ liệu có chứa các ô bị lỗi (#N/A, #VALUE!, #DIV/0!) hoặc các dòng bị ẩn do Filter.",
        "why_choose": "Hàm SUM thông thường sẽ báo lỗi ngay nếu cột có 1 ô lỗi bất kỳ. AGGREGATE với tùy chọn option 6 cho phép tự động bỏ qua toàn bộ ô lỗi; option 5 bỏ qua dòng ẩn. Một giải pháp cứu cánh cho dữ liệu kế toán chưa sạch.",
    },
    {
        "id": "sumproduct",
        "name": "SUMPRODUCT",
        "category": "aggregation",
        "category_vi": "Toán Học & Thống Kê Tổng Hợp",
        "template": "=SUMPRODUCT(array1, [array2], ...)",
        "example": "=SUMPRODUCT(C2:C50, D2:D50)",
        "when_to_use": "Dùng khi cần tính tổng thành tiền = Số lượng x Đơn giá cho toàn bộ danh sách mà không cần phải tạo thêm cột phụ 'Thành tiền'.",
        "why_choose": "Tính toán ma trận trực tiếp trong bộ nhớ, tiết kiệm tài nguyên bảng tính. Ngoài ra có thể dùng làm bộ lọc nâng cao cho các điều kiện so sánh độ dài chuỗi hoặc phân tích logic không hỗ trợ trong SUMIFS.",
    },
    {
        "id": "subtotal",
        "name": "SUBTOTAL",
        "category": "aggregation",
        "category_vi": "Toán Học & Thống Kê Tổng Hợp",
        "template": "=SUBTOTAL(function_num, ref1, ...)",
        "example": "=SUBTOTAL(109, D2:D100)",
        "when_to_use": "Dùng cho dòng Tổng cộng ở chân các bảng biểu có sử dụng chức năng AutoFilter hoặc Data Outline.",
        "why_choose": "Số hiệu 109 giúp hàm tự động tính lại tổng chỉ cho các dòng đang hiển thị sau khi lọc, không bao gồm các dòng bị ẩn. Ngăn chặn hiện tượng tính trùng 2 lần nếu trong dải ô có chứa các dòng subtotal con.",
    },

    # -----------------------------------------------------------------------
    # 3. XỬ LÝ CHUỖI & VĂN BẢN (TEXT & MANIPULATION)
    # -----------------------------------------------------------------------
    {
        "id": "textjoin",
        "name": "TEXTJOIN",
        "category": "text",
        "category_vi": "Xử Lý Chuỗi & Văn Bản",
        "template": "=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)",
        "example": '=TEXTJOIN(", ", TRUE, B2:B20)',
        "when_to_use": "Dùng khi cần gộp nhiều ô, nhiều dòng thành một chuỗi duy nhất có dấu phẩy hoặc dấu gạch nối ngăn cách (ví dụ: tạo chuỗi danh sách email, danh sách mã hợp đồng).",
        "why_choose": "Vượt trội hơn dấu '&' và hàm CONCATENATE: Cho phép truyền trực tiếp một dải ô (Range B2:B20) thay vì phải gõ B2 & ', ' & B3 & ', ' thủ công; cờ ignore_empty=TRUE tự động bỏ qua các ô trống không tạo ra dấu phẩy thừa.",
    },
    {
        "id": "textsplit",
        "name": "TEXTSPLIT",
        "category": "text",
        "category_vi": "Xử Lý Chuỗi & Văn Bản",
        "template": "=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])",
        "example": '=TEXTSPLIT(A2, "-")',
        "when_to_use": "Dùng khi cần phân tách chuỗi 'Họ - Đệm - Tên' hoặc mã giao dịch 'VCB-2026-05' thành các cột riêng biệt.",
        "why_choose": "Hoạt động hoàn toàn bằng công thức mảng động tự tràn kết quả sang các cột bên cạnh, thay thế hoàn toàn tính năng Text-to-Columns tĩnh của Excel. Bất cứ khi nào ô nguồn thay đổi, kết quả phân tách tự động cập nhật.",
    },
    {
        "id": "textbefore_after",
        "name": "TEXTBEFORE / TEXTAFTER",
        "category": "text",
        "category_vi": "Xử Lý Chuỗi & Văn Bản",
        "template": "=TEXTBEFORE(text, delimiter, [instance_num]) / =TEXTAFTER(...)",
        "example": '=TEXTAFTER(A2, "@")',
        "when_to_use": "Dùng khi cần trích xuất tên miền từ địa chỉ email ('gmail.com' từ 'user@gmail.com') hoặc lấy mã tỉnh từ số căn cước/biển số xe.",
        "why_choose": "Cực kỳ ngắn gọn và dễ hiểu! Loại bỏ hoàn toàn sự kết hợp ác mộng của =MID(A2, FIND('@', A2)+1, LEN(A2)) trong các phiên bản Excel cũ.",
    },
    {
        "id": "trim_clean",
        "name": "TRIM & CLEAN",
        "category": "text",
        "category_vi": "Xử Lý Chuỗi & Văn Bản",
        "template": "=TRIM(CLEAN(text))",
        "example": "=TRIM(CLEAN(A2))",
        "when_to_use": "Dùng để làm sạch dữ liệu thô xuất ra từ các phần mềm kế toán, ERP hoặc web scraper thường bị dính khoảng trắng đầu cuối hoặc ký tự xuống dòng ẩn.",
        "why_choose": "CLEAN loại bỏ 32 ký tự không in được đầu tiên trong bảng mã ASCII (mã 0 đến 31); TRIM loại bỏ toàn bộ khoảng trắng ở đầu, cuối và chuẩn hóa khoảng trắng kép giữa các từ về 1 khoảng trắng đơn.",
    },
    {
        "id": "exact",
        "name": "EXACT",
        "category": "text",
        "category_vi": "Xử Lý Chuỗi & Văn Bản",
        "template": "=EXACT(text1, text2)",
        "example": '=EXACT(A2, "Admin")',
        "when_to_use": "Dùng khi kiểm tra mật khẩu, mã xác thực hoặc mã phân loại yêu cầu độ chính xác tuyệt đối có phân biệt chữ HOA / chữ thường.",
        "why_choose": "Toán tử so sánh thông thường trong Excel (=) không phân biệt hoa thường ('abc' = 'ABC' trả về TRUE). Hàm EXACT bắt buộc phải khớp 100% từng ký tự hoa/thường mới trả về TRUE.",
    },

    # -----------------------------------------------------------------------
    # 4. LOGIC & KIỂM SOÁT LỖI (LOGIC & ERROR HANDLING)
    # -----------------------------------------------------------------------
    {
        "id": "ifs",
        "name": "IFS",
        "category": "logic",
        "category_vi": "Logic & Kiểm Soát Lỗi",
        "template": "=IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ..., [TRUE, default_value])",
        "example": '=IFS(A2>=85, "Xuất sắc", A2>=70, "Khá", A2>=50, "Trung bình", TRUE, "Yếu")',
        "when_to_use": "Dùng khi xếp loại học lực, phân loại thang điểm KPI, tính hoa hồng bán hàng nhiều bậc.",
        "why_choose": "Triệt tiêu hoàn toàn cấu trúc IF lồng nhau (Nested IF) sâu 5-7 tầng dễ nhầm lẫn dấu đóng ngoặc. Cú pháp phẳng, dễ đọc, dễ kiểm tra logic; tham số cuối cùng TRUE đóng vai trò là giá trị mặc định (ELSE).",
    },
    {
        "id": "switch",
        "name": "SWITCH",
        "category": "logic",
        "category_vi": "Logic & Kiểm Soát Lỗi",
        "template": "=SWITCH(expression, val1, result1, [val2, result2], ..., [default])",
        "example": '=SWITCH(WEEKDAY(TODAY()), 1, "Chủ Nhật", 2, "Thứ Hai", 7, "Thứ Bảy", "Ngày thường")',
        "when_to_use": "Dùng khi đối chiếu một giá trị duy nhất với một danh sách các kết quả cố định (ví dụ: mã trạng thái 1, 2, 3 tương ứng với 'Chờ duyệt', 'Đã duyệt', 'Từ chối').",
        "why_choose": "Không cần phải lặp lại tên biểu thức so sánh nhiều lần như IFS (IFS(A2=1, ..., A2=2, ...)). SWITCH chỉ đánh giá ô nguồn 1 lần duy nhất, tăng tốc độ tính toán.",
    },
    {
        "id": "iferror",
        "name": "IFERROR",
        "category": "logic",
        "category_vi": "Logic & Kiểm Soát Lỗi",
        "template": "=IFERROR(value, value_if_error)",
        "example": '=IFERROR(VLOOKUP(A2, B:C, 2, 0), "Chưa đăng ký")',
        "when_to_use": "Dùng để bọc bên ngoài các công thức có khả năng sinh lỗi như phép chia cho 0 hoặc hàm tìm kiếm không thấy giá trị.",
        "why_choose": "Biến bảng tính trở nên chuyên nghiệp, loại bỏ các ký hiệu lỗi xấu xí (#DIV/0!, #N/A, #VALUE!) làm ảnh hưởng đến thẩm mỹ báo cáo gửi sếp hoặc đối tác.",
    },
    {
        "id": "isnumber_match",
        "name": "ISNUMBER / MATCH",
        "category": "logic",
        "category_vi": "Logic & Kiểm Soát Lỗi",
        "template": "=ISNUMBER(MATCH(lookup_val, lookup_range, 0))",
        "example": "=ISNUMBER(MATCH(A2, DanhSachCam, 0))",
        "when_to_use": "Dùng trong định dạng có điều kiện (Conditional Formatting) để tự động tô màu các dòng có mã nằm trong danh sách cấm / danh sách VIP.",
        "why_choose": "Trả về đúng giá trị Boolean chuẩn (TRUE hoặc FALSE) thay vì trả về số thứ tự như MATCH hoặc lỗi #N/A, rất an toàn để kích hoạt quy tắc tô màu hoặc Data Validation.",
    },

    # -----------------------------------------------------------------------
    # 5. MẢNG ĐỘNG & HIỆN ĐẠI 365 (DYNAMIC ARRAYS & MODERN)
    # -----------------------------------------------------------------------
    {
        "id": "let",
        "name": "LET",
        "category": "dynamic",
        "category_vi": "Mảng Động & Hiện Đại (Excel 365)",
        "template": "=LET(name1, name_value1, [name2, name_value2], ..., calculation)",
        "example": '=LET(doanh_so, SUM(B2:B50), ty_le, 0.05, doanh_so * ty_le)',
        "when_to_use": "Dùng khi xây dựng các công thức tài chính hoặc kỹ thuật phức tạp có các biểu thức con lặp lại nhiều lần trong cùng một ô.",
        "why_choose": "Bước nhảy vọt về hiệu năng: Excel chỉ tính toán biến số 1 lần duy nhất và lưu tạm vào bộ nhớ thay vì tính lại 3-4 lần. Giúp công thức ngắn gọn, dễ bảo trì như viết code lập trình.",
    },
    {
        "id": "unique",
        "name": "UNIQUE",
        "category": "dynamic",
        "category_vi": "Mảng Động & Hiện Đại (Excel 365)",
        "template": "=UNIQUE(array, [by_col], [exactly_once])",
        "example": "=UNIQUE(C2:C500)",
        "when_to_use": "Dùng khi cần trích xuất danh mục các sản phẩm hoặc danh sách phòng ban không trùng lặp từ một bảng dữ liệu nhập liệu hàng nghìn dòng.",
        "why_choose": "Loại bỏ hoàn toàn công cụ Remove Duplicates thủ công. Khi bảng nguồn thêm bản ghi mới, danh sách UNIQUE tự động cập nhật ngay lập tức.",
    },
    {
        "id": "sort_sortby",
        "name": "SORT / SORTBY",
        "category": "dynamic",
        "category_vi": "Mảng Động & Hiện Đại (Excel 365)",
        "template": "=SORT(array, [sort_index], [sort_order])",
        "example": "=SORT(FILTER(A2:D100, D2:D100>0), 4, -1)",
        "when_to_use": "Dùng khi cần tự động xếp hạng Top 10 khách hàng có doanh thu cao nhất theo thứ tự từ lớn đến bé.",
        "why_choose": "Sắp xếp tự động bằng công thức động, không làm thay đổi trật tự bảng dữ liệu gốc. Kết hợp mượt mà với UNIQUE và FILTER.",
    },
    {
        "id": "sequence",
        "name": "SEQUENCE",
        "category": "dynamic",
        "category_vi": "Mảng Động & Hiện Đại (Excel 365)",
        "template": "=SEQUENCE(rows, [columns], [start], [step])",
        "example": "=SEQUENCE(COUNTA(B2:B100), 1, 1, 1)",
        "when_to_use": "Dùng để tự động tạo cột Số Thứ Tự (STT: 1, 2, 3...) tự co giãn theo số dòng thực tế của bảng dữ liệu.",
        "why_choose": "Không cần phải kéo chuột 'Fill Handle' thủ công. Dữ liệu có 50 dòng thì STT tự sinh 50 dòng, dữ liệu giảm xuống 20 dòng thì STT tự co lại, không bao giờ bị thừa số.",
    },
    {
        "id": "lambda",
        "name": "LAMBDA",
        "category": "dynamic",
        "category_vi": "Mảng Động & Hiện Đại (Excel 365)",
        "template": "=LAMBDA([parameter1, parameter2, ...], calculation)",
        "example": '=LAMBDA(gia, vat, gia * (1 + vat))',
        "when_to_use": "Dùng khi muốn tự tạo một hàm riêng theo nghiệp vụ đặc thù của công ty (ví dụ hàm TinhThueTNCN) và đặt tên trong Name Manager để toàn công ty dùng chung.",
        "why_choose": "Tạo hàm tùy biến mà KHÔNG CẦN viết mã VBA hay macro, giúp tệp Excel lưu dưới dạng .xlsx thông thường mà không bị phần mềm diệt virus hoặc chính sách bảo mật chặn macro (.xlsm).",
    },

    # -----------------------------------------------------------------------
    # 6. THỜI GIAN & TÀI CHÍNH (DATE & FINANCIAL)
    # -----------------------------------------------------------------------
    {
        "id": "eomonth",
        "name": "EOMONTH",
        "category": "financial",
        "category_vi": "Thời Gian & Tài Chính",
        "template": "=EOMONTH(start_date, months)",
        "example": "=EOMONTH(TODAY(), 0)",
        "when_to_use": "Dùng để xác định ngày cuối cùng của tháng hiện tại, tháng trước (months=-1) hoặc tháng sau (months=1), dùng trong tính khấu hao, tính lãi vay và chốt sổ kế toán.",
        "why_choose": "Tự động xử lý chính xác số ngày của từng tháng (28, 29, 30 hay 31 ngày) kể cả năm nhuận, tránh hoàn toàn sai sót khi người dùng gán cứng ngày 30 hay 31.",
    },
    {
        "id": "edate",
        "name": "EDATE",
        "category": "financial",
        "category_vi": "Thời Gian & Tài Chính",
        "template": "=EDATE(start_date, months)",
        "example": "=EDATE(A2, 12)",
        "when_to_use": "Dùng để tính ngày đáo hạn hợp đồng, thời hạn bảo hành sản phẩm (ví dụ cộng thêm 6 tháng, 12 tháng, 24 tháng).",
        "why_choose": "Giữ nguyên ngày trong tháng và tự động tịnh tiến tháng/năm, xử lý an toàn các mốc ngày cuối tháng mà không bị lệch sang tháng kế tiếp.",
    },
    {
        "id": "networkdays_intl",
        "name": "NETWORKDAYS.INTL",
        "category": "financial",
        "category_vi": "Thời Gian & Tài Chính",
        "template": "=NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])",
        "example": '=NETWORKDAYS.INTL(A2, B2, 1, NgayLe)',
        "when_to_use": "Dùng để tính số ngày làm việc thực tế giữa 2 mốc thời gian để tính lương, tính thời gian hoàn thành dự án (SLA).",
        "why_choose": "Linh hoạt tuyệt đối: Cho phép tùy chọn chế độ nghỉ cuối tuần (chỉ nghỉ Chủ Nhật hoặc nghỉ cả Thứ 7 & CN) và tự động trừ các ngày nghỉ lễ quốc gia trong danh sách [holidays].",
    },
    {
        "id": "pmt",
        "name": "PMT",
        "category": "financial",
        "category_vi": "Thời Gian & Tài Chính",
        "template": "=PMT(rate, nper, pv, [fv], [type])",
        "example": "=PMT(8%/12, 24, -500000000)",
        "when_to_use": "Dùng trong ngân hàng và quản lý tài chính cá nhân để tính số tiền trả góp gốc + lãi cố định đều đặn hàng tháng cho khoản vay mua nhà/xe.",
        "why_choose": "Chuẩn mực ngân hàng quốc tế, tự động tính toán cân đối dòng tiền chiết khấu theo lãi suất kỳ hạn.",
    },
    {
        "id": "xirr",
        "name": "XIRR",
        "category": "financial",
        "category_vi": "Thời Gian & Tài Chính",
        "template": "=XIRR(values, dates, [guess])",
        "example": "=XIRR(B2:B10, A2:A10)",
        "when_to_use": "Dùng trong phân tích đầu tư dự án để tính Tỷ suất hoàn vốn nội bộ (Internal Rate of Return) cho các khoản đầu tư và thu hồi vốn diễn ra ở các mốc thời gian không đều nhau.",
        "why_choose": "Vượt trội hơn hàm IRR truyền thống (IRR giả định các khoản tiền phát sinh ở các khoảng thời gian đều đặn hàng năm). XIRR tính toán chuẩn xác theo từng mốc ngày thực tế phát sinh dòng tiền.",
    },
]


# ---------------------------------------------------------------------------
# API Functions
# ---------------------------------------------------------------------------

def get_formula_catalog(category: Optional[str] = None, search: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Returns filtered formula list based on category and search query.
    """
    items = FORMULA_CATALOG
    if category and category.lower() != "all":
        items = [i for i in items if i["category"].lower() == category.lower()]

    if search:
        q = search.strip().lower()
        filtered = []
        for i in items:
            searchable_text = f"{i['name']} {i['category_vi']} {i['template']} {i['when_to_use']} {i['why_choose']}".lower()
            if q in searchable_text:
                filtered.append(i)
        items = filtered

    return items


def generate_formula_playbook_xlsx(output_path: str) -> Tuple[bool, str]:
    """
    Generates a publication-grade, dual-sheet Excel (.xlsx) workbook:
    - Sheet 1: 'Formula_Playbook' (Master formula cheat sheet & technical catalog)
    - Sheet 2: 'Live_Interactive_Formulas' (Real business demo with active calculating formulas)
    """
    if openpyxl is None:
        return False, "Thư viện 'openpyxl' chưa được cài đặt."

    dest = os.path.abspath(output_path)
    out_dir = os.path.dirname(dest)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # =======================================================================
    # SHEET 1: FORMULA PLAYBOOK (CATALOG GUIDE)
    # =======================================================================
    ws1 = wb.create_sheet(title="Formula_Playbook")

    # Banner Title
    ws1.merge_cells("A1:G1")
    t_cell = ws1.cell(row=1, column=1)
    t_cell.value = "CẨM NANG & THƯ VIỆN CÔNG THỨC EXCEL CHUYÊN NGHIỆP (EXCEL FORMULA PLAYBOOK)"
    t_cell.fill = NAVY_TITLE_FILL
    t_cell.font = TITLE_FONT
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 36

    # Subtitle
    ws1.merge_cells("A2:G2")
    sub_cell = ws1.cell(row=2, column=1)
    sub_cell.value = "Tổng hợp 32 công thức thực chiến chuẩn Enterprise — Cập nhật phiên bản Excel 365 & Classic | Antigravity Office Studio"
    sub_cell.fill = SUBHEADER_FILL
    sub_cell.font = SUBTITLE_FONT
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 22

    # Headers
    headers = [
        ("STT", 6, "center"),
        ("Nhóm Nghiệp Vụ", 24, "center"),
        ("Tên Hàm", 18, "center"),
        ("Cú Pháp Công Thức Mẫu", 42, "left"),
        ("Ví Dụ Thực Hành Cụ Thể", 36, "left"),
        ("Dùng Trong Trường Hợp Nào?", 44, "left"),
        ("Tại Sao Lại Chọn Hàm Này? (Lợi Thế Kỹ Thuật)", 56, "left"),
    ]

    ws1.row_dimensions[3].height = 28
    for c_idx, (h_title, _, h_align) in enumerate(headers, start=1):
        cell = ws1.cell(row=3, column=c_idx)
        cell.value = h_title
        cell.fill = NAVY_HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=True)
        cell.border = HEADER_BORDER

    # Populate Data Rows
    current_row = 4
    for idx, item in enumerate(FORMULA_CATALOG, start=1):
        ws1.row_dimensions[current_row].height = 54
        fill = ALT_ROW_FILL if idx % 2 == 0 else WHITE_FILL

        # Col A: STT
        c_stt = ws1.cell(row=current_row, column=1, value=idx)
        c_stt.alignment = Alignment(horizontal="center", vertical="center")

        # Col B: Nhóm
        c_cat = ws1.cell(row=current_row, column=2, value=item["category_vi"])
        c_cat.alignment = Alignment(horizontal="center", vertical="center")

        # Col C: Tên Hàm
        c_name = ws1.cell(row=current_row, column=3, value=item["name"])
        c_name.alignment = Alignment(horizontal="center", vertical="center")
        c_name.font = DATA_BOLD_FONT

        # Col D: Cú Pháp
        c_tpl = ws1.cell(row=current_row, column=4, value=item["template"])
        c_tpl.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_tpl.font = CODE_FONT

        # Col E: Ví Dụ
        c_ex = ws1.cell(row=current_row, column=5, value=item["example"])
        c_ex.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_ex.font = CODE_FONT

        # Col F: Trường Hợp Dùng
        c_when = ws1.cell(row=current_row, column=6, value=item["when_to_use"])
        c_when.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Col G: Lý Do Chọn
        c_why = ws1.cell(row=current_row, column=7, value=item["why_choose"])
        c_why.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for col_k in range(1, 8):
            cell = ws1.cell(row=current_row, column=col_k)
            cell.fill = fill
            cell.border = THIN_BORDER
            if col_k not in (3, 4, 5):
                cell.font = DATA_FONT

        current_row += 1

    # Freeze panes at A4
    ws1.freeze_panes = "A4"

    # Set Column Widths
    for c_idx, (_, col_w, _) in enumerate(headers, start=1):
        col_letter = get_column_letter(c_idx)
        ws1.column_dimensions[col_letter].width = col_w

    # Print setup
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
    if apply_worksheet_print_setup:
        apply_worksheet_print_setup(ws1, orientation="landscape")
    else:
        ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE

    # =======================================================================
    # SHEET 2: LIVE INTERACTIVE FORMULAS (DEMO WITH LIVE CALCULATIONS)
    # =======================================================================
    ws2 = wb.create_sheet(title="Live_Interactive_Formulas")

    # Sheet 2 Title
    ws2.merge_cells("A1:G1")
    s2_title = ws2.cell(row=1, column=1)
    s2_title.value = "MINH HỌA CÔNG THỨC EXCEL THỰC THI SỐNG ĐỘNG (LIVE EXECUTABLE DEMO)"
    s2_title.fill = NAVY_TITLE_FILL
    s2_title.font = TITLE_FONT
    s2_title.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 32

    # Section 1: Raw Data Table
    ws2.merge_cells("A3:F3")
    sec1_hdr = ws2.cell(row=3, column=1, value="1. BẢNG DỮ LIỆU BÁN HÀNG DOANH NGHIỆP (MẪU RAW DATA)")
    sec1_hdr.font = Font(name="Calibri", size=11, bold=True, color="0F294A")

    raw_headers = ["Mã NV", "Họ Và Tên", "Phòng Ban", "Khu Vực", "Doanh Số (VNĐ)", "Xếp Loại"]
    ws2.row_dimensions[4].height = 24
    for c_i, h in enumerate(raw_headers, start=1):
        cell = ws2.cell(row=4, column=c_i, value=h)
        cell.fill = NAVY_HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = HEADER_BORDER

    mock_employees = [
        ("NV001", "Nguyễn Văn An", "Kinh doanh", "Miền Bắc", 125000000),
        ("NV002", "Trần Thị Bình", "Kinh doanh", "Miền Nam", 85000000),
        ("NV003", "Lê Hoàng Cường", "Kỹ thuật", "Miền Bắc", 45000000),
        ("NV004", "Phạm Minh Đức", "Kinh doanh", "Miền Trung", 110000000),
        ("NV005", "Hoàng Thu Giang", "Kế toán", "Miền Bắc", 30000000),
        ("NV006", "Vũ Hải Nam", "Kinh doanh", "Miền Bắc", 95000000),
        ("NV007", "Đỗ Bích Ngọc", "Marketing", "Miền Nam", 60000000),
        ("NV008", "Ngô Quốc Phong", "Kinh doanh", "Miền Nam", 140000000),
    ]

    for r_offset, (emp_id, emp_name, emp_dept, emp_region, emp_sales) in enumerate(mock_employees, start=5):
        ws2.row_dimensions[r_offset].height = 20
        fill = ALT_ROW_FILL if r_offset % 2 == 0 else WHITE_FILL

        c_id = ws2.cell(row=r_offset, column=1, value=emp_id)
        c_id.number_format = "@"
        c_id.alignment = Alignment(horizontal="center", vertical="center")

        c_nm = ws2.cell(row=r_offset, column=2, value=emp_name)
        c_nm.alignment = Alignment(horizontal="left", vertical="center")

        c_dp = ws2.cell(row=r_offset, column=3, value=emp_dept)
        c_dp.alignment = Alignment(horizontal="center", vertical="center")

        c_rg = ws2.cell(row=r_offset, column=4, value=emp_region)
        c_rg.alignment = Alignment(horizontal="center", vertical="center")

        c_sl = ws2.cell(row=r_offset, column=5, value=emp_sales)
        c_sl.number_format = "#,##0"
        c_sl.alignment = Alignment(horizontal="right", vertical="center")

        # Dynamic formula in Column F: IFS formula!
        c_rk = ws2.cell(row=r_offset, column=6)
        c_rk.value = f'=IFS(E{r_offset}>=100000000, "Xuất Sắc", E{r_offset}>=50000000, "Đạt KPI", TRUE, "Cần Cải Thiện")'
        c_rk.alignment = Alignment(horizontal="center", vertical="center")
        c_rk.font = DATA_BOLD_FONT

        for c_k in range(1, 7):
            cell = ws2.cell(row=r_offset, column=c_k)
            cell.fill = fill
            cell.border = THIN_BORDER
            if c_k != 6:
                cell.font = DATA_FONT

    # Section 2: Live KPI Calculations & Practical Formula Demos
    sec2_start = 15
    ws2.merge_cells(f"A{sec2_start}:G{sec2_start}")
    sec2_hdr = ws2.cell(row=sec2_start, column=1, value="2. BẢNG TỔNG HỢP KPI & TRA CỨU TỰ ĐỘNG BẰNG CÔNG THỨC LIVE")
    sec2_hdr.font = Font(name="Calibri", size=11, bold=True, color="0F294A")

    kpi_headers = ["Chỉ Tiêu / Nghiệp Vụ", "Công Thức Excel Thực Thi", "Giá Trị Tính Toán Trực Tiếp", "Ý Nghĩa Nghiệp Vụ & Vì Sao Chọn"]
    ws2.row_dimensions[sec2_start + 1].height = 24
    ws2.cell(row=sec2_start + 1, column=1, value=kpi_headers[0]).fill = NAVY_HEADER_FILL
    ws2.cell(row=sec2_start + 1, column=2, value=kpi_headers[1]).fill = NAVY_HEADER_FILL
    ws2.cell(row=sec2_start + 1, column=3, value=kpi_headers[2]).fill = NAVY_HEADER_FILL
    ws2.cell(row=sec2_start + 1, column=4, value=kpi_headers[3]).fill = NAVY_HEADER_FILL
    for c_i in range(1, 5):
        cell = ws2.cell(row=sec2_start + 1, column=c_i)
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = HEADER_BORDER

    live_kpis = [
        (
            "Tổng Doanh Số Phòng Kinh Doanh",
            '=SUMIFS(E5:E12, C5:C12, "Kinh doanh")',
            "#,##0",
            "SUMIFS tính tổng doanh thu chỉ riêng phòng Kinh doanh, bỏ qua phòng Kế toán/Kỹ thuật.",
        ),
        (
            "Tổng Doanh Số Kinh Doanh Miền Bắc",
            '=SUMIFS(E5:E12, C5:C12, "Kinh doanh", D5:D12, "Miền Bắc")',
            "#,##0",
            "SUMIFS lọc 2 điều kiện đồng thời (Phòng ban + Khu vực) không cần cột phụ.",
        ),
        (
            "Số Lượng Nhân Viên Đạt Mốc Xuất Sắc (>=100tr)",
            '=COUNTIFS(E5:E12, ">=100000000")',
            "0",
            "COUNTIFS đếm nhanh các nhân sự vượt chỉ tiêu để tính thưởng quý.",
        ),
        (
            "Doanh Số Trung Bình Miền Nam",
            '=AVERAGEIFS(E5:E12, D5:D12, "Miền Nam")',
            "#,##0",
            "AVERAGEIFS tự động tính trung bình cộng không bị ảnh hưởng bởi ô trống.",
        ),
        (
            "Tra Cứu Họ Tên Của Mã Nhân Viên 'NV004'",
            '=XLOOKUP("NV004", A5:A12, B5:B12, "Không tìm thấy")',
            "@",
            "XLOOKUP tra cứu chính xác, tự động fallback an toàn nếu mã không tồn tại.",
        ),
        (
            "Tra Cứu Doanh Số Của Mã Nhân Viên 'NV008'",
            '=XLOOKUP("NV008", A5:A12, E5:E12, 0)',
            "#,##0",
            "XLOOKUP lấy số liệu doanh số từ phải qua trái hoặc ngang dọc bất kỳ.",
        ),
        (
            "Chuỗi Danh Sách Họ Tên Nhân Sự",
            '=TEXTJOIN("; ", TRUE, B5:B12)',
            "@",
            "TEXTJOIN gộp toàn bộ họ tên thành 1 chuỗi ngăn cách bởi dấu chấm phẩy.",
        ),
    ]

    for k_idx, (kpi_title, kpi_formula, kpi_fmt, kpi_desc) in enumerate(live_kpis, start=sec2_start + 2):
        ws2.row_dimensions[k_idx].height = 26
        fill = HIGHLIGHT_KPI_FILL if k_idx % 2 == 0 else WHITE_FILL

        # Col 1: Title
        c1 = ws2.cell(row=k_idx, column=1, value=kpi_title)
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c1.font = DATA_BOLD_FONT

        # Col 2: Formula string representation
        c2 = ws2.cell(row=k_idx, column=2, value=kpi_formula)
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c2.font = CODE_FONT

        # Col 3: Live Executing Formula!
        c3 = ws2.cell(row=k_idx, column=3)
        c3.value = kpi_formula
        c3.number_format = kpi_fmt
        c3.alignment = Alignment(horizontal="right" if kpi_fmt != "@" else "left", vertical="center")
        c3.font = KPI_VAL_FONT

        # Col 4: Description
        c4 = ws2.cell(row=k_idx, column=4, value=kpi_desc)
        c4.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c4.font = DATA_FONT

        for c_k in range(1, 5):
            cell = ws2.cell(row=k_idx, column=c_k)
            cell.fill = fill
            cell.border = THIN_BORDER

    # Set column widths for Sheet 2
    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 36
    ws2.column_dimensions["C"].width = 28
    ws2.column_dimensions["D"].width = 46
    ws2.column_dimensions["E"].width = 22
    ws2.column_dimensions["F"].width = 20
    ws2.column_dimensions["G"].width = 16

    ws2.freeze_panes = "A5"
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
    if apply_worksheet_print_setup:
        apply_worksheet_print_setup(ws2, orientation="landscape")
    else:
        ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE

    wb.save(dest)
    return True, dest
