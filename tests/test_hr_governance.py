"""
tests/test_hr_governance.py — Automated Unit Tests for Enterprise HR Governance,
6 Data Scopes, Segregation of Duties (SoD) & DBML Schema.
"""

import json
import unittest

from core.ba_hub.hr_governance import (
    HRDataScope,
    DATA_SCOPE_METADATA,
    DataScopeEvaluator,
    SoDRiskLevel,
    SoDRuleDefinition,
    SOD_CORE_RULES,
    SoDConflictAuditor,
    HRCoreDBSchema,
)
from core.ba_hub.templates import STORY_TEMPLATES
from server.mcp_server import execute_hr_governance_blueprint


class TestHRDataScopes(unittest.TestCase):
    """Test suite for the 6 Canonical HR Data Scopes & SQL Predicate Generator."""

    def test_six_data_scopes_exist(self):
        expected_scopes = {
            "GLOBAL",
            "LEGAL_ENTITY",
            "BUSINESS_UNIT",
            "DEPARTMENT_TREE",
            "DIRECT_TEAM",
            "SELF_ONLY"
        }
        actual_scopes = {s.value for s in HRDataScope}
        self.assertEqual(expected_scopes, actual_scopes)

    def test_scope_metadata_levels(self):
        for scope in HRDataScope:
            meta = DataScopeEvaluator.get_scope_metadata(scope)
            self.assertIn("level", meta)
            self.assertIn("name_vi", meta)
            self.assertIn("om_boundary", meta)
            self.assertIn("filter_description", meta)
            self.assertTrue(1 <= meta["level"] <= 6)

    def test_build_scope_sql_filter_global(self):
        clause, params = DataScopeEvaluator.build_scope_sql_filter(HRDataScope.GLOBAL, {})
        self.assertEqual(clause, "1=1")
        self.assertEqual(params, {})

    def test_build_scope_sql_filter_legal_entity(self):
        clause, params = DataScopeEvaluator.build_scope_sql_filter(
            HRDataScope.LEGAL_ENTITY,
            {"legal_entity_id": "LE-001"}
        )
        self.assertEqual(clause, "emp.legal_entity_id = :scope_legal_entity_id")
        self.assertEqual(params, {"scope_legal_entity_id": "LE-001"})

    def test_build_scope_sql_filter_business_unit(self):
        clause, params = DataScopeEvaluator.build_scope_sql_filter(
            HRDataScope.BUSINESS_UNIT,
            {"business_unit_id": "BU-TECH"}
        )
        self.assertEqual(clause, "dept.business_unit_id = :scope_business_unit_id")
        self.assertEqual(params, {"scope_business_unit_id": "BU-TECH"})

    def test_build_scope_sql_filter_department_tree_with_path(self):
        clause, params = DataScopeEvaluator.build_scope_sql_filter(
            HRDataScope.DEPARTMENT_TREE,
            {"department_path": "/1/4/12/"}
        )
        self.assertIn("dept.path = :scope_dept_exact", clause)
        self.assertIn("dept.path LIKE :scope_dept_pattern", clause)
        self.assertEqual(params["scope_dept_exact"], "/1/4/12/")
        self.assertEqual(params["scope_dept_pattern"], "/1/4/12/%")

    def test_build_scope_sql_filter_department_tree_with_child_ids(self):
        clause, params = DataScopeEvaluator.build_scope_sql_filter(
            HRDataScope.DEPARTMENT_TREE,
            {"child_department_ids": ["d1", "d2", "d3"]}
        )
        self.assertIn("emp.department_id IN", clause)
        self.assertEqual(len(params), 3)

    def test_build_scope_sql_filter_direct_team(self):
        clause, params = DataScopeEvaluator.build_scope_sql_filter(
            HRDataScope.DIRECT_TEAM,
            {"department_id": "dept-99", "current_user_emp_id": "EMP-007"}
        )
        self.assertIn("emp.department_id = :scope_direct_dept_id", clause)
        self.assertIn("emp.direct_manager_id = :scope_manager_emp_id", clause)
        self.assertEqual(params["scope_direct_dept_id"], "dept-99")
        self.assertEqual(params["scope_manager_emp_id"], "EMP-007")

    def test_build_scope_sql_filter_self_only(self):
        clause, params = DataScopeEvaluator.build_scope_sql_filter(
            HRDataScope.SELF_ONLY,
            {"current_user_emp_id": "EMP-ME"}
        )
        self.assertEqual(clause, "emp.id = :scope_self_emp_id")
        self.assertEqual(params, {"scope_self_emp_id": "EMP-ME"})

    def test_build_scope_sql_missing_context_raises(self):
        with self.assertRaises(ValueError):
            DataScopeEvaluator.build_scope_sql_filter(HRDataScope.LEGAL_ENTITY, {})

        with self.assertRaises(ValueError):
            DataScopeEvaluator.build_scope_sql_filter(HRDataScope.SELF_ONLY, {})

    def test_format_scope_matrix_markdown(self):
        md = DataScopeEvaluator.format_scope_matrix_markdown()
        self.assertIn("| Cấp | Mã Scope | Tên Phạm Vi Dữ Liệu", md)
        self.assertIn("`GLOBAL`", md)
        self.assertIn("`SELF_ONLY`", md)


class TestSoDRuleEngine(unittest.TestCase):
    """Test suite for Segregation of Duties (SoD) Rules & Conflict Auditor."""

    def test_sod_core_rules_loaded(self):
        self.assertTrue(len(SOD_CORE_RULES) >= 5)
        rule_ids = {r.rule_id for r in SOD_CORE_RULES}
        for expected in ["SOD-01", "SOD-02", "SOD-03", "SOD-04", "SOD-05"]:
            self.assertIn(expected, rule_ids)

    def test_audit_maker_checker_conflict(self):
        # User has both contract create and contract approve -> SOD-01
        assigned_perms = ["hrm.contract.create", "hrm.contract.approve", "hrm.employee.view"]
        res = SoDConflictAuditor.audit_permission_set(assigned_perms, strict=True)
        self.assertFalse(res["compliant"])
        self.assertEqual(res["verdict"], "HARD_BLOCK_REJECTED")
        self.assertGreaterEqual(res["critical_conflicts"], 1)
        conflict_rule_ids = [v["rule_id"] for v in res["violations"]]
        self.assertIn("SOD-01", conflict_rule_ids)

    def test_audit_admin_business_conflict(self):
        # User has both system role assign and employee create -> SOD-02
        assigned_perms = ["admin.role.assign", "hrm.employee.create"]
        res = SoDConflictAuditor.audit_permission_set(assigned_perms, strict=True)
        self.assertFalse(res["compliant"])
        self.assertEqual(res["verdict"], "HARD_BLOCK_REJECTED")
        conflict_rule_ids = [v["rule_id"] for v in res["violations"]]
        self.assertIn("SOD-02", conflict_rule_ids)

    def test_audit_clean_permissions_passed(self):
        # Clean operator: Only creates drafts, cannot approve or assign roles
        assigned_perms = ["hrm.contract.create", "hrm.employee.create", "hrm.employee.view"]
        res = SoDConflictAuditor.audit_permission_set(assigned_perms, strict=True)
        self.assertTrue(res["compliant"])
        self.assertEqual(res["score"], 100)
        self.assertEqual(res["verdict"], "PASSED")
        self.assertEqual(len(res["violations"]), 0)

    def test_offboarding_4_gatekeepers_blocked(self):
        # Only 2 gates cleared: manager and IT, missing facility and finance
        status = {
            "gate_manager_work": True,
            "gate_it_assets": True,
            "gate_facility_admin": False,
            "gate_finance_debt": False
        }
        res = SoDConflictAuditor.audit_offboarding_readiness(status)
        self.assertFalse(res["can_terminate_contract"])
        self.assertEqual(res["verdict"], "BLOCKED_PENDING_HANDOVERS")
        self.assertEqual(res["cleared_count"], 2)
        self.assertEqual(len(res["missing_gates"]), 2)

    def test_offboarding_4_gatekeepers_ready(self):
        # All 4 gates cleared
        status = {
            "gate_manager_work": True,
            "gate_it_assets": True,
            "gate_facility_admin": True,
            "gate_finance_debt": True
        }
        res = SoDConflictAuditor.audit_offboarding_readiness(status)
        self.assertTrue(res["can_terminate_contract"])
        self.assertEqual(res["verdict"], "READY_FOR_HR_TERMINATION")
        self.assertEqual(res["cleared_count"], 4)
        self.assertEqual(len(res["missing_gates"]), 0)


class TestHRSchemaAndTemplates(unittest.TestCase):
    """Test suite for DBML/DDL Schemas and Templates integration."""

    def test_dbml_schema_content(self):
        dbml = HRCoreDBSchema.get_dbml_schema()
        self.assertIn("Table om_legal_entities", dbml)
        self.assertIn("Table om_companies", dbml)
        self.assertIn("Table om_business_units", dbml)
        self.assertIn("Table om_departments", dbml)
        self.assertIn("Table om_cost_centers", dbml)
        self.assertIn("Table hrm_employees", dbml)
        self.assertIn("Table hrm_contracts", dbml)
        self.assertIn("Table hrm_job_history", dbml)
        self.assertIn("Table ofb_termination_requests", dbml)
        self.assertIn("Table ofb_handover_checklists", dbml)

    def test_sql_ddl_schema_content(self):
        sql = HRCoreDBSchema.get_sql_ddl()
        self.assertIn("CREATE TABLE IF NOT EXISTS om_departments", sql)
        self.assertIn("CREATE TABLE IF NOT EXISTS hrm_employees", sql)
        self.assertIn("CREATE INDEX IF NOT EXISTS idx_om_dept_path", sql)

    def test_hr_templates_registered(self):
        self.assertIn("om_organization_tree", STORY_TEMPLATES)
        self.assertIn("hrm_employee_movement", STORY_TEMPLATES)
        self.assertIn("ofb_termination_workflow", STORY_TEMPLATES)
        self.assertIn("admin_data_scope_rbac", STORY_TEMPLATES)

        om_story = STORY_TEMPLATES["om_organization_tree"]
        self.assertEqual(om_story["domain"], "Quản trị Cơ cấu Tổ chức (Organizational Management - OM)")
        self.assertTrue(len(om_story["business_rules"]) >= 3)
        self.assertTrue(len(om_story["scenarios"]) >= 2)


class TestHRGovernanceMCPTool(unittest.TestCase):
    """Test suite for mcp_server hr_governance_blueprint tool integration."""

    def test_mcp_data_scopes_markdown(self):
        res = execute_hr_governance_blueprint({"action": "data_scopes", "format": "markdown"})
        self.assertIn("markdown_table", res)
        self.assertIn("GLOBAL", res["markdown_table"])

    def test_mcp_data_scopes_evaluate(self):
        res = execute_hr_governance_blueprint({
            "action": "data_scopes",
            "scope": "DEPARTMENT_TREE",
            "context": {"department_path": "/1/2/"}
        })
        self.assertEqual(res["scope"], "DEPARTMENT_TREE")
        self.assertIn("dept.path", res["sql_filter"])

    def test_mcp_sod_audit(self):
        res = execute_hr_governance_blueprint({
            "action": "sod_audit",
            "permissions": ["hrm.contract.create", "hrm.contract.approve"],
            "strict": True
        })
        self.assertFalse(res["compliant"])
        self.assertEqual(res["verdict"], "HARD_BLOCK_REJECTED")

    def test_mcp_ofb_workflow(self):
        res = execute_hr_governance_blueprint({
            "action": "ofb_workflow",
            "checklist": {
                "gate_manager_work": True,
                "gate_it_assets": True,
                "gate_facility_admin": True,
                "gate_finance_debt": True
            }
        })
        self.assertTrue(res["can_terminate_contract"])

    def test_mcp_export_schema(self):
        res_dbml = execute_hr_governance_blueprint({"action": "export_schema", "format": "dbml"})
        self.assertIn("om_departments", res_dbml["schema"])

        res_sql = execute_hr_governance_blueprint({"action": "export_schema", "format": "sql"})
        self.assertIn("CREATE TABLE", res_sql["schema"])


if __name__ == "__main__":
    unittest.main()

