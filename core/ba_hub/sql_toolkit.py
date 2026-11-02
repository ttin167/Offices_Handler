"""
core.ba_hub.sql_toolkit — 7 Automated Core SQL Integrity Tests & Data Audit Suite.

Inspired by prachitparikh/ba-toolkit:
1. Seven Core SQL Data Integrity Tests:
   - Test 1: Orphan Records Check (Foreign Key referential integrity).
   - Test 2: Duplicate Primary/Composite Key Check.
   - Test 3: Critical Business Nulls Check.
   - Test 4: Business Range & Boundary Violations (Negative values, >100%).
   - Test 5: Referential Status Conflict (e.g. Order Cancelled vs Item Shipped).
   - Test 6: Chronological / Timestamp Anomalies (e.g. delivery_date < order_date).
   - Test 7: Value Reconciliation Variance Sum (Header Total vs Sum of Line Items).
2. Multi-Dialect Generation: PostgreSQL, MSSQL (T-SQL), Oracle (PL/SQL), MySQL.
3. Natural Language to Business Query Template Assistant.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class SQLIntegrityTestItem:
    test_id: str
    test_name: str
    description: str
    severity: str  # CRITICAL, MAJOR, MINOR
    sql_query: str
    dialect: str = "postgres"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "id": self.test_id,
            "name": self.test_name,
            "description": self.description,
            "severity": self.severity,
            "sql": self.sql_query,
            "dialect": self.dialect,
        }


class SQLIntegrityTestSuite:
    """Generates the 7 Core Automated SQL Data Integrity Tests for BA audits."""

    @classmethod
    def generate_suite(
        cls,
        target_table: str = "orders",
        parent_table: Optional[str] = "customers",
        primary_key: str = "order_id",
        foreign_key: Optional[str] = "customer_id",
        child_table: Optional[str] = "order_items",
        amount_column: str = "total_amount",
        date_column: str = "order_date",
        status_column: str = "status",
        dialect: str = "postgres",
    ) -> List[SQLIntegrityTestItem]:
        tests: List[SQLIntegrityTestItem] = []
        d = dialect.lower()

        # ----------------------------------------------------------------------
        # Test 1: Orphan Records Check
        # ----------------------------------------------------------------------
        if parent_table and foreign_key:
            sql_1 = (
                f"-- 1. KIỂM TRA BẢN GHI MỒ CÔI (ORPHAN RECORDS)\n"
                f"-- Dò tìm các bản ghi trong '{target_table}' có khóa ngoại '{foreign_key}' không tồn tại trong '{parent_table}'\n"
                f"SELECT t.{primary_key}, t.{foreign_key}\n"
                f"FROM {target_table} t\n"
                f"LEFT JOIN {parent_table} p ON t.{foreign_key} = p.{foreign_key}\n"
                f"WHERE t.{foreign_key} IS NOT NULL AND p.{foreign_key} IS NULL;\n"
            )
            tests.append(
                SQLIntegrityTestItem(
                    test_id="SQL-INT-01",
                    test_name="Orphan Records Check (Khóa Ngoại Mồ Côi)",
                    description=f"Dò tìm các dòng trong '{target_table}' tham chiếu đến khóa không tồn tại trong '{parent_table}'.",
                    severity="CRITICAL",
                    sql_query=sql_1,
                    dialect=d,
                )
            )

        # ----------------------------------------------------------------------
        # Test 2: Duplicate Key Check
        # ----------------------------------------------------------------------
        sql_2 = (
            f"-- 2. KIỂM TRA TRÙNG LẶP KHÓA (DUPLICATE PRIMARY / COMPOSITE KEY)\n"
            f"-- Dò tìm các bản ghi bị nhân bản khóa '{primary_key}' trong bảng '{target_table}'\n"
            f"SELECT {primary_key}, COUNT(*) AS duplicate_count\n"
            f"FROM {target_table}\n"
            f"GROUP BY {primary_key}\n"
            f"HAVING COUNT(*) > 1;\n"
        )
        tests.append(
            SQLIntegrityTestItem(
                test_id="SQL-INT-02",
                test_name="Duplicate Keys Check (Trùng Khóa Chính)",
                description=f"Phát hiện bản ghi bị trùng lặp khóa chính '{primary_key}' do lỗi ETL/migration.",
                severity="CRITICAL",
                sql_query=sql_2,
                dialect=d,
            )
        )

        # ----------------------------------------------------------------------
        # Test 3: Critical Business Nulls Check
        # ----------------------------------------------------------------------
        trim_fn = "LTRIM(RTRIM" if d in ("mssql", "sqlserver") else "TRIM"
        sql_3 = (
            f"-- 3. KIỂM TRA Ô TRỐNG BẤT HỢP LỆ (CRITICAL BUSINESS NULLS)\n"
            f"-- Dò tìm các trường bắt buộc nghiệp vụ bị NULL hoặc khoảng trắng trong '{target_table}'\n"
            f"SELECT {primary_key}, {status_column}, {amount_column}\n"
            f"FROM {target_table}\n"
            f"WHERE {status_column} IS NULL OR {trim_fn}({status_column}) = ''\n"
            f"   OR {amount_column} IS NULL;\n"
        )
        tests.append(
            SQLIntegrityTestItem(
                test_id="SQL-INT-03",
                test_name="Critical Nulls Check (Ô Trống Nghiệp Vụ)",
                description=f"Quét các bản ghi thiếu thông tin bắt buộc trong '{target_table}'.",
                severity="MAJOR",
                sql_query=sql_3,
                dialect=d,
            )
        )

        # ----------------------------------------------------------------------
        # Test 4: Business Range & Boundary Violations
        # ----------------------------------------------------------------------
        sql_4 = (
            f"-- 4. KIỂM TRA VI PHẠM NGƯỠNG NGHIỆP VỤ (RANGE & BOUNDARY VIOLATIONS)\n"
            f"-- Phát hiện số tiền <= 0 hoặc vượt ngưỡng bất thường trong '{target_table}'\n"
            f"SELECT {primary_key}, {amount_column}\n"
            f"FROM {target_table}\n"
            f"WHERE {amount_column} <= 0\n"
            f"   OR {amount_column} > 1000000000; -- Hạn mức cảnh báo 1 tỷ VNĐ\n"
        )
        tests.append(
            SQLIntegrityTestItem(
                test_id="SQL-INT-04",
                test_name="Boundary Violations (Vi Phạm Ngưỡng Giá Trị)",
                description=f"Kiểm tra các giá trị âm vô lý hoặc vượt hạn mức tối đa cho phép trên '{amount_column}'.",
                severity="MAJOR",
                sql_query=sql_4,
                dialect=d,
            )
        )

        # ----------------------------------------------------------------------
        # Test 5: Referential Status Conflict
        # ----------------------------------------------------------------------
        if child_table:
            sql_5 = (
                f"-- 5. XUNG ĐỘT TRẠNG THÁI LOGIC (REFERENTIAL STATUS CONFLICT)\n"
                f"-- Đơn hàng cha đã HỦY ('CANCELLED') nhưng dòng chi tiết con '{child_table}' vẫn 'SHIPPED'\n"
                f"SELECT t.{primary_key}, t.{status_column} AS parent_status, c.item_status AS child_status\n"
                f"FROM {target_table} t\n"
                f"JOIN {child_table} c ON t.{primary_key} = c.{primary_key}\n"
                f"WHERE t.{status_column} IN ('CANCELLED', 'VOID', 'REJECTED')\n"
                f"  AND c.item_status IN ('SHIPPED', 'DELIVERED', 'COMPLETED');\n"
            )
            tests.append(
                SQLIntegrityTestItem(
                    test_id="SQL-INT-05",
                    test_name="Referential Status Conflict (Mâu Thuẫn Trạng Thái)",
                    description=f"Phát hiện trạng thái bất nhất giữa '{target_table}' và bảng chi tiết con '{child_table}'.",
                    severity="CRITICAL",
                    sql_query=sql_5,
                    dialect=d,
                )
            )

        # ----------------------------------------------------------------------
        # Test 6: Chronological / Timestamp Anomalies
        # ----------------------------------------------------------------------
        sql_6 = (
            f"-- 6. NGHỊCH LÝ THỜI GIAN GIAO DỊCH (CHRONOLOGICAL ANOMALIES)\n"
            f"-- Phát hiện ngày tạo lớn hơn ngày giao dịch hoặc thời gian hoàn tất trước khi tạo\n"
            f"SELECT {primary_key}, {date_column}, created_at\n"
            f"FROM {target_table}\n"
            f"WHERE {date_column} > CURRENT_DATE + 1 -- Giao dịch tương lai bất hợp lệ\n"
            f"   OR (updated_at IS NOT NULL AND updated_at < created_at);\n"
        )
        tests.append(
            SQLIntegrityTestItem(
                test_id="SQL-INT-06",
                test_name="Chronological Anomalies (Nghịch Lý Thời Gian)",
                description=f"Kiểm tra sai lệch mốc thời gian logic trong '{target_table}'.",
                severity="MAJOR",
                sql_query=sql_6,
                dialect=d,
            )
        )

        # ----------------------------------------------------------------------
        # Test 7: Value Reconciliation Variance Sum
        # ----------------------------------------------------------------------
        if child_table:
            sql_7 = (
                f"-- 7. ĐỐI SOÁT TỔNG TIỀN VÀ CHI TIẾT (VALUE RECONCILIATION VARIANCE SUM)\n"
                f"-- So khớp tổng tiền '{target_table}.{amount_column}' với tổng các dòng con '{child_table}.line_amount'\n"
                f"SELECT t.{primary_key}, t.{amount_column} AS header_total,\n"
                f"       SUM(c.line_amount) AS items_sum,\n"
                f"       (t.{amount_column} - SUM(c.line_amount)) AS variance_diff\n"
                f"FROM {target_table} t\n"
                f"JOIN {child_table} c ON t.{primary_key} = c.{primary_key}\n"
                f"GROUP BY t.{primary_key}, t.{amount_column}\n"
                f"HAVING ABS(t.{amount_column} - SUM(c.line_amount)) > 0.01;\n"
            )
            tests.append(
                SQLIntegrityTestItem(
                    test_id="SQL-INT-07",
                    test_name="Value Reconciliation Sum (Lệch Tổng Tiền Chi Tiết)",
                    description=f"Đối soát tổng tiền trên header '{target_table}' và tổng các dòng chi tiết '{child_table}'.",
                    severity="CRITICAL",
                    sql_query=sql_7,
                    dialect=d,
                )
            )

        return tests
