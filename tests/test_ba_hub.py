"""
tests.test_ba_hub — Comprehensive Unit & Integration Tests for BA Enterprise Workstation.

Verifies:
1. User Story & Gherkin AC Builder + Quality Linter (BABOK Scorecard).
2. OpenXML Word and Markdown Exporters for User Stories.
3. Use Case Specification Builder, Linter & Word Exporter.
4. Data Dictionary SQL DDL & Excel Scanners.
5. Two-Way Data Reconciliation Engine (Composite Keys, Numeric Tolerance, Variance Report).
6. UAT Testcase Matrix Generator (In-cell validations & OpenXML styling).
7. Full Enterprise BRD / SRS Document Synthesizer.
"""

import os
import tempfile
import pytest
from core.ba_hub.templates import (
    STORY_TEMPLATES,
    USECASE_TEMPLATES,
    DATA_DICT_TEMPLATES,
    BRD_DEFAULT_OUTLINE,
)
from core.ba_hub.story_builder import (
    UserStory,
    GherkinScenario,
    StoryQualityLinter,
    StoryBuilder,
)
from core.ba_hub.usecase_builder import (
    UseCaseSpec,
    UseCaseQualityLinter,
    UseCaseBuilder,
)
from core.ba_hub.data_dictionary import (
    DataField,
    DataDictionaryScanner,
    FieldMappingEngine,
)
from core.ba_hub.reconciliation import (
    DataReconciliationEngine,
    ReconciliationSummary,
)
from core.ba_hub.testcase_generator import (
    UATScenario,
    UATTestcaseGenerator,
)
from core.ba_hub.brd_generator import BRDDocumentGenerator


# ------------------------------------------------------------------------------
# 1. USER STORY & GHERKIN LINTER TESTS
# ------------------------------------------------------------------------------
class TestUserStoryAndLinter:
    def test_templates_load_cleanly(self):
        assert "banking_fund_transfer" in STORY_TEMPLATES
        tpl = STORY_TEMPLATES["banking_fund_transfer"]
        assert tpl["role"] != ""
        assert len(tpl["scenarios"]) >= 3

    def test_story_quality_linter_perfect_score(self):
        tpl = STORY_TEMPLATES["banking_fund_transfer"]
        story = UserStory.from_dict({
            "id": "US-TEST-001",
            "title": tpl["title"],
            "domain": tpl["domain"],
            "role": tpl["role"],
            "goal": tpl["goal"],
            "benefit": tpl["benefit"],
            "business_rules": tpl["business_rules"],
            "scenarios": tpl["scenarios"],
        })
        audit = StoryQualityLinter.lint_story(story)
        assert audit.score >= 85
        assert "A+" in audit.grade or "Xuất sắc" in audit.grade
        assert audit.total_scenarios >= 3

    def test_story_quality_linter_penalties(self):
        # Flawed story: vague persona "Người dùng", no business rules, no scenarios
        flawed = UserStory(
            id="US-FLAW-001",
            title="Đăng nhập",
            role="Người dùng",
            goal="đăng nhập vào hệ thống",
            benefit="",
            business_rules=[],
            scenarios=[],
        )
        audit = StoryQualityLinter.lint_story(flawed)
        assert audit.score < 50
        severities = [f.severity for f in audit.findings]
        assert "CRITICAL" in severities
        messages = [f.message for f in audit.findings]
        assert any("người dùng" in m.lower() for m in messages)

    def test_story_word_and_markdown_export(self):
        tpl = STORY_TEMPLATES["ecommerce_checkout"]
        story = UserStory.from_dict({
            "id": "US-ECOMM-001",
            "title": tpl["title"],
            "domain": tpl["domain"],
            "role": tpl["role"],
            "goal": tpl["goal"],
            "benefit": tpl["benefit"],
            "business_rules": tpl["business_rules"],
            "scenarios": tpl["scenarios"],
        })

        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, "story.docx")
            md_path = os.path.join(tmpdir, "story.md")

            StoryBuilder.export_to_docx([story], docx_path, doc_title="Kiểm Thử Export Word")
            assert os.path.exists(docx_path)
            assert os.path.getsize(docx_path) > 1000

            StoryBuilder.export_to_markdown([story], md_path)
            assert os.path.exists(md_path)
            with open(md_path, "r", encoding="utf-8") as f:
                md_content = f.read()
            assert "As a" in md_content
            assert "Scenario:" in md_content


# ------------------------------------------------------------------------------
# 2. USE CASE SPECIFICATION TESTS
# ------------------------------------------------------------------------------
class TestUseCaseBuilder:
    def test_usecase_linter(self):
        uc_data = USECASE_TEMPLATES["uc_fund_transfer"]
        uc = UseCaseSpec.from_dict(uc_data)
        audit = UseCaseQualityLinter.lint_usecase(uc)
        assert audit.score >= 85
        assert "A+" in audit.grade or "Xuất sắc" in audit.grade

    def test_usecase_word_export(self):
        uc_data = USECASE_TEMPLATES["uc_fund_transfer"]
        uc = UseCaseSpec.from_dict(uc_data)
        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = os.path.join(tmpdir, "usecase.docx")
            UseCaseBuilder.export_to_docx(uc, out_docx)
            assert os.path.exists(out_docx)
            assert os.path.getsize(out_docx) > 1000


# ------------------------------------------------------------------------------
# 3. DATA DICTIONARY & FIELD MAPPING TESTS
# ------------------------------------------------------------------------------
class TestDataDictionary:
    def test_scan_sql_ddl(self):
        ddl = """
        CREATE TABLE accounts (
            account_id VARCHAR(36) PRIMARY KEY,
            account_no VARCHAR(20) NOT NULL,
            customer_id VARCHAR(20) REFERENCES customers(id),
            balance DECIMAL(15,2) DEFAULT 0.00,
            is_active BOOLEAN NOT NULL,
            created_at TIMESTAMP
        );
        """
        tables = DataDictionaryScanner.scan_sql_ddl(ddl)
        table_map = {t["table_name"]: t["fields"] for t in tables}
        assert "accounts" in table_map
        fields = table_map["accounts"]
        assert len(fields) == 6

        # Check primary key
        pk_field = next(f for f in fields if f.get("field") == "account_id")
        assert pk_field.get("pk") == "YES"
        assert pk_field.get("nullable") == "NO"

        # Check foreign key
        fk_field = next(f for f in fields if f.get("field") == "customer_id")
        assert fk_field.get("fk") != "-"
        assert "customers(id)" in fk_field.get("fk")

    def test_export_data_dictionary_to_excel(self):
        fields = [
            DataField("cif_no", "Mã CIF", "VARCHAR(20)", "NO", True, False, "-", "Mã định danh duy nhất", "Sinh từ Core"),
            DataField("balance", "Số Dư", "DECIMAL(18,2)", "NO", False, False, "-", "Số dư khả dụng", ">= 0"),
        ]
        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "dict.xlsx")
            FieldMappingEngine.export_data_dictionary_to_excel("CUSTOMER_ACCOUNTS", fields, out_xlsx)
            assert os.path.exists(out_xlsx)
            assert os.path.getsize(out_xlsx) > 1000


# ------------------------------------------------------------------------------
# 4. TWO-WAY DATA RECONCILIATION TESTS
# ------------------------------------------------------------------------------
class TestDataReconciliation:
    def test_reconcile_in_memory_records(self):
        source_a = [
            {"trans_id": "FT001", "amount": 1000000, "status": "SUCCESS"},
            {"trans_id": "FT002", "amount": 2500000, "status": "SUCCESS"},
            {"trans_id": "FT003", "amount": 500000, "status": "FAILED"}, # Missing in B
            {"trans_id": "FT004", "amount": 750000, "status": "SUCCESS"}, # Variance in B
        ]
        source_b = [
            {"trans_id": "FT001", "amount": 1000000, "status": "SUCCESS"},
            {"trans_id": "FT002", "amount": 2500000, "status": "SUCCESS"},
            {"trans_id": "FT004", "amount": 745000, "status": "SUCCESS"}, # Variance 5000
            {"trans_id": "FT005", "amount": 120000, "status": "SUCCESS"}, # Missing in A
        ]

        summary = DataReconciliationEngine.reconcile(
            source_a=source_a,
            source_b=source_b,
            key_columns=["trans_id"],
            compare_columns=["amount"],
            numeric_tolerance=0.0,
            source_a_name="Hệ Thống Nội Bộ",
            source_b_name="Sao Kê NAPAS",
        )

        assert summary.total_a == 4
        assert summary.total_b == 4
        assert summary.matched_exact == 2  # FT001, FT002
        assert summary.matched_discrepant == 1  # FT004
        assert summary.missing_in_b == 1  # FT003
        assert summary.missing_in_a == 1  # FT005
        assert summary.match_rate_pct == 50.0

        # Export report to Excel
        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "recon_report.xlsx")
            DataReconciliationEngine.export_report_to_excel(summary, out_xlsx)
            assert os.path.exists(out_xlsx)
            assert os.path.getsize(out_xlsx) > 2000


# ------------------------------------------------------------------------------
# 5. UAT TESTCASE GENERATOR TESTS
# ------------------------------------------------------------------------------
class TestUATGenerator:
    def test_generate_and_export_uat(self):
        stories = [STORY_TEMPLATES["banking_fund_transfer"]]
        testcases = UATTestcaseGenerator.generate_from_stories(stories)
        assert len(testcases) >= 3

        with tempfile.TemporaryDirectory() as tmpdir:
            out_xlsx = os.path.join(tmpdir, "uat.xlsx")
            UATTestcaseGenerator.export_uat_checklist_to_excel("Mobile Banking", testcases, out_xlsx)
            assert os.path.exists(out_xlsx)
            assert os.path.getsize(out_xlsx) > 1000


# ------------------------------------------------------------------------------
# 6. ENTERPRISE BRD DOCUMENT SYNTHESIZER TESTS
# ------------------------------------------------------------------------------
class TestBRDGenerator:
    def test_generate_complete_brd(self):
        stories = [
            UserStory.from_dict({
                "id": "US-001",
                **STORY_TEMPLATES["banking_fund_transfer"]
            })
        ]
        meta = {
            "title": "TÀI LIỆU YÊU CẦU NGHIỆP VỤ (BRD)",
            "project_name": "Ngân Hàng Số 2026",
            "version": "1.0",
            "author": "Senior BA",
            "organization": "Ngân Hàng TMCP Công Nghệ",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = os.path.join(tmpdir, "brd_complete.docx")
            BRDDocumentGenerator.generate_brd(meta, stories, out_docx)
            assert os.path.exists(out_docx)
            assert os.path.getsize(out_docx) > 5000
