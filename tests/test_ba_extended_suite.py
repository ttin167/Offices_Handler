"""
tests.test_ba_extended_suite — Comprehensive Tests for Top 4 Open-Source BA Technologies.

Verifies:
1. RTM Engine (Traceability Linking & Coverage Gap Audit).
2. DBML Bidirectional Engine (DBML <-> SQL DDL <-> Mermaid ERD).
3. Satsuma Data Mapping DSL & Type Incompatibility Guard.
4. Scope Creep & Change Impact Analysis Engine.
5. 7 Core Automated SQL Data Integrity Tests Suite.
"""

import os
import tempfile
import pytest
from core.ba_hub.templates import STORY_TEMPLATES, USECASE_TEMPLATES
from core.ba_hub.story_builder import UserStory
from core.ba_hub.usecase_builder import UseCaseSpec
from core.ba_hub.rtm_matrix import RTMItem, RTMCoverage, RTMEngine
from core.ba_hub.dbml_satsuma import (
    DBMLField,
    DBMLTable,
    DBMLEngine,
    SatsumaMappingRule,
    SatsumaMappingEngine,
)
from core.ba_hub.scope_diff import ScopeDeltaItem, ScopeChangeSummary, ScopeCreepAuditor
from core.ba_hub.sql_toolkit import SQLIntegrityTestItem, SQLIntegrityTestSuite


# ------------------------------------------------------------------------------
# 1. RTM ENGINE TESTS
# ------------------------------------------------------------------------------
class TestRTMEngine:
    def test_build_rtm_and_coverage(self):
        stories = [
            UserStory.from_dict({
                "id": "US-001",
                **STORY_TEMPLATES["banking_fund_transfer"]
            })
        ]
        usecases = [UseCaseSpec.from_dict(USECASE_TEMPLATES["uc_fund_transfer"])]
        testcases = [
            {"id": "TC-001", "title": "Chuyển tiền thành công trong hạn mức thông thường", "status": "PASS"},
            {"id": "TC-002", "title": "Từ chối giao dịch do số dư khả dụng không đủ", "status": "FAIL"},
        ]

        rtm_items = RTMEngine.build_rtm(stories, usecases, testcases)
        assert len(rtm_items) >= 3

        cov = RTMEngine.calculate_coverage(rtm_items)
        assert cov.total_requirements == 1
        assert cov.coverage_pct == 100.0
        assert cov.total_testcases >= 2
        assert cov.passed_testcases >= 1
        assert cov.failed_testcases >= 1

    def test_export_rtm_to_excel(self):
        items = [
            RTMItem(
                req_id="BR-PAY-01",
                req_title="Chuyển tiền nhanh 24/7",
                story_id="US-001",
                usecase_id="UC-001",
                scenario_id="AC-01",
                scenario_title="Thành công",
                testcase_id="TC-001",
                status="PASS",
            ),
            RTMItem(
                req_id="BR-PAY-02",
                req_title="Sinh trắc học",
                story_id="US-001",
                usecase_id="UC-001",
                scenario_id="AC-02",
                scenario_title="Bắt buộc quét mặt",
                testcase_id="TC-002",
                status="FAIL",
                defect_id="DEF-001",
            ),
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "rtm_test.xlsx")
            RTMEngine.export_rtm_to_excel(items, out_xlsx, project_name="Ngân Hàng Số")
            assert os.path.exists(out_xlsx)
            assert os.path.getsize(out_xlsx) > 2000


# ------------------------------------------------------------------------------
# 2. DBML BIDIRECTIONAL ENGINE TESTS
# ------------------------------------------------------------------------------
class TestDBMLEngine:
    SAMPLE_DBML = """
    Table users {
      id integer [pk]
      username varchar(50) [not null, unique]
      created_at timestamp
    }
    Table orders {
      order_id integer [pk]
      user_id integer [ref: > users.id]
      total_amount decimal(12,2)
    }
    """

    def test_parse_dbml_and_to_sql(self):
        tables = DBMLEngine.parse_dbml(self.SAMPLE_DBML)
        assert len(tables) == 2
        tbl_names = [t.name for t in tables]
        assert "users" in tbl_names
        assert "orders" in tbl_names

        sql = DBMLEngine.dbml_to_sql(self.SAMPLE_DBML, dialect="postgres")
        assert "CREATE TABLE users" in sql
        assert "CREATE TABLE orders" in sql
        assert "PRIMARY KEY" in sql
        assert "REFERENCES users(id)" in sql

    def test_sql_to_dbml(self):
        sql = """
        CREATE TABLE products (
            prod_id VARCHAR(20) PRIMARY KEY,
            prod_name VARCHAR(100) NOT NULL,
            price DECIMAL(10,2)
        );
        """
        dbml = DBMLEngine.sql_to_dbml(sql)
        assert "Table products" in dbml
        assert "prod_id" in dbml
        assert "pk" in dbml

    def test_dbml_to_mermaid_erd(self):
        mermaid = DBMLEngine.dbml_to_mermaid_erd(self.SAMPLE_DBML)
        assert "erDiagram" in mermaid
        assert "users" in mermaid
        assert "orders" in mermaid


# ------------------------------------------------------------------------------
# 3. SATSUMA DATA MAPPING DSL TESTS
# ------------------------------------------------------------------------------
class TestSatsumaMappingEngine:
    def test_parse_satsuma_dsl_and_lineage(self):
        dsl = """
        # Data Mapping Spec
        Source_CRM.cust_id -> Target_Core.cif_number [Direct: Khóa chính CIF]
        Source_CRM.id_card -> Target_Core.national_id [Mask_PII: Che 6 số giữa]
        Source_CRM.first_name + " " + Source_CRM.last_name -> Target_Core.full_name [Concat: Ghép họ tên]
        """
        rules = SatsumaMappingEngine.parse_satsuma_dsl(dsl)
        assert len(rules) == 3
        assert rules[0].rule_type == "Direct"
        assert rules[1].rule_type == "Mask_PII"
        assert rules[2].rule_type == "Concat"

        lineage = SatsumaMappingEngine.generate_data_lineage_mermaid(rules)
        assert "flowchart LR" in lineage
        assert "Source_CRM.cust_id" in lineage
        assert "Target_Core.cif_number" in lineage

    def test_type_incompatibility_guard(self):
        rule = SatsumaMappingRule(
            source_expr="Source.notes",
            source_type="VARCHAR(200)",
            target_field="Target.short_desc",
            target_type="VARCHAR(50)",
            rule_type="Direct",
        )
        SatsumaMappingEngine._audit_type_compatibility(rule)
        assert len(rule.warnings) > 0
        assert any("TRUNCATION RISK" in w for w in rule.warnings)
        assert rule.is_valid is False


# ------------------------------------------------------------------------------
# 4. SCOPE CREEP AUDITOR TESTS
# ------------------------------------------------------------------------------
class TestScopeCreepAuditor:
    def test_scope_diff_analysis_and_export(self):
        import docx

        with tempfile.TemporaryDirectory() as tmpdir:
            base_file = os.path.join(tmpdir, "srs_v1.docx")
            rev_file = os.path.join(tmpdir, "srs_v2.docx")

            doc_v1 = docx.Document()
            doc_v1.add_heading("1. Đăng nhập", level=1)
            doc_v1.add_paragraph("Hệ thống cho phép đăng nhập bằng số điện thoại và mật khẩu thông thường.")
            doc_v1.save(base_file)

            doc_v2 = docx.Document()
            doc_v2.add_heading("1. Đăng nhập", level=1)
            doc_v2.add_paragraph("Hệ thống cho phép đăng nhập bằng số điện thoại và mật khẩu thông thường kèm mã OTP xác thực.")
            doc_v2.add_heading("2. Thanh toán hóa đơn", level=1)
            doc_v2.add_paragraph("Hệ thống tích hợp cổng thanh toán trực tuyến NAPAS và VNPay để thanh toán tự động.")
            doc_v2.save(rev_file)

            summary = ScopeCreepAuditor.analyze_documents(base_file, rev_file)
            assert summary.total_changes >= 1
            assert summary.new_requirements_count >= 1
            assert summary.total_estimated_mandays > 0.0

            out_report = os.path.join(tmpdir, "scope_report.docx")
            ScopeCreepAuditor.export_scope_report_docx(summary, out_report, project_name="Ngân Hàng")
            assert os.path.exists(out_report)
            assert os.path.getsize(out_report) > 2000


# ------------------------------------------------------------------------------
# 5. 7 CORE SQL INTEGRITY TESTS
# ------------------------------------------------------------------------------
class TestSQLIntegritySuite:
    def test_generate_7_sql_tests(self):
        tests = SQLIntegrityTestSuite.generate_suite(
            target_table="orders",
            parent_table="customers",
            primary_key="order_id",
            foreign_key="customer_id",
            child_table="order_items",
            amount_column="total_amount",
            dialect="postgres",
        )
        assert len(tests) == 7
        t_ids = [t.test_id for t in tests]
        assert "SQL-INT-01" in t_ids
        assert "SQL-INT-02" in t_ids
        assert "SQL-INT-03" in t_ids
        assert "SQL-INT-04" in t_ids
        assert "SQL-INT-05" in t_ids
        assert "SQL-INT-06" in t_ids
        assert "SQL-INT-07" in t_ids

        # Check SQL syntax content
        orphan_test = next(t for t in tests if t.test_id == "SQL-INT-01")
        assert "LEFT JOIN customers" in orphan_test.sql_query
        assert "IS NULL" in orphan_test.sql_query

        dup_test = next(t for t in tests if t.test_id == "SQL-INT-02")
        assert "HAVING COUNT(*) > 1" in dup_test.sql_query

        variance_test = next(t for t in tests if t.test_id == "SQL-INT-07")
        assert "SUM(c.line_amount)" in variance_test.sql_query
