"""
tests/test_ba_openapi_and_decision_table.py — Test Suite for OpenAPI Scanner & BABOK Decision Table.
"""

import json
import os
import tempfile
import pytest

from core.ba_hub.data_dictionary import (
    DataDictionaryScanner,
    FieldMappingEngine,
    scan_openapi_spec,
)
from core.ba_hub.decision_table import (
    DecisionTable,
    ConditionRule,
    ActionRule,
    DecisionRuleEntry,
)


SAMPLE_OPENAPI_SPEC = {
    "openapi": "3.0.1",
    "info": {
        "title": "Digital Banking Payment Service",
        "version": "2.4.0",
        "description": "API Gateway for NAPAS 247 Instant Transfer and QR Code Payments."
    },
    "paths": {
        "/api/v1/payments/napas247": {
            "post": {
                "summary": "Khởi tạo giao dịch chuyển tiền nhanh NAPAS 247",
                "operationId": "initiateNapasTransfer",
                "parameters": [
                    {
                        "name": "X-Client-Id",
                        "in": "header",
                        "required": True,
                        "schema": {"type": "string", "maxLength": 50}
                    }
                ],
                "responses": {
                    "200": {"description": "Giao dịch thành công"},
                    "400": {"description": "Tham số không hợp lệ"}
                }
            }
        }
    },
    "components": {
        "schemas": {
            "PaymentOrder": {
                "type": "object",
                "required": ["order_id", "sender_account", "amount"],
                "properties": {
                    "order_id": {"type": "string", "format": "uuid", "description": "Mã đơn hàng duy nhất"},
                    "sender_account": {"type": "string", "maxLength": 20, "description": "Số tài khoản nguồn"},
                    "beneficiary_account": {"type": "string", "maxLength": 20, "description": "Số tài khoản thụ hưởng"},
                    "amount": {"type": "number", "minimum": 1000, "maximum": 500000000, "description": "Số tiền giao dịch VND"},
                    "fee": {"type": "number", "default": 0},
                    "memo": {"type": "string", "maxLength": 100}
                }
            }
        }
    }
}


def test_scan_openapi_spec():
    data = scan_openapi_spec(SAMPLE_OPENAPI_SPEC)
    assert data["title"] == "Digital Banking Payment Service"
    assert data["version"] == "2.4.0"
    assert data["endpoints_count"] == 1
    assert data["schemas_count"] == 1

    tbl = data["tables"][0]
    assert tbl["table_name"] == "PaymentOrder"
    fields = {f["field"]: f for f in tbl["fields"]}
    assert "order_id" in fields
    assert fields["order_id"]["nullable"] == "NO"
    assert "amount" in fields
    assert "min: 1000" in fields["amount"]["rules"]


def test_export_openapi_excel_and_docx():
    data = scan_openapi_spec(SAMPLE_OPENAPI_SPEC)
    with tempfile.TemporaryDirectory() as tmp:
        xlsx_path = os.path.join(tmp, "openapi_dict.xlsx")
        docx_path = os.path.join(tmp, "openapi_dict.docx")

        FieldMappingEngine.export_openapi_to_excel(data, xlsx_path)
        assert os.path.isfile(xlsx_path)
        assert os.path.getsize(xlsx_path) > 1000

        FieldMappingEngine.export_openapi_to_docx(data, docx_path)
        assert os.path.isfile(docx_path)
        assert os.path.getsize(docx_path) > 1000


def test_decision_table_completeness_and_contradiction():
    dt = DecisionTable(
        table_id="DT-CREDIT-01",
        title="Duyệt Hạn Mức Thẻ Tín Dụng",
        domain="Ngân Hàng Số",
        conditions=[
            ConditionRule(name="Thu Nhập >= 15tr", possible_values=["Y", "N"]),
            ConditionRule(name="Lịch Sử CIC Tốt", possible_values=["Y", "N"]),
        ],
        actions=[
            ActionRule(name="Cấp Thẻ Chuẩn"),
            ActionRule(name="Từ Chối"),
        ],
        rules=[
            DecisionRuleEntry("R01", {"Thu Nhập >= 15tr": "Y", "Lịch Sử CIC Tốt": "Y"}, {"Cấp Thẻ Chuẩn": "X", "Từ Chối": "-"}),
            DecisionRuleEntry("R02", {"Thu Nhập >= 15tr": "N", "Lịch Sử CIC Tốt": "N"}, {"Cấp Thẻ Chuẩn": "-", "Từ Chối": "X"}),
        ]
    )

    audit = dt.audit_completeness_and_contradictions()
    assert audit.total_theoretical_rules == 4
    assert audit.actual_rules_count == 2
    assert len(audit.missing_combinations) == 2
    assert len(audit.contradictions) == 0

    with tempfile.TemporaryDirectory() as tmp:
        xlsx_path = os.path.join(tmp, "decision_table.xlsx")
        docx_path = os.path.join(tmp, "decision_table.docx")

        dt.export_to_excel(xlsx_path)
        assert os.path.isfile(xlsx_path)
        assert os.path.getsize(xlsx_path) > 1000

        dt.export_to_docx(docx_path)
        assert os.path.isfile(docx_path)
        assert os.path.getsize(docx_path) > 1000

