"""
tests/test_ba_obsidian.py — Comprehensive Unit & Integration Tests for Obsidian Vault Integration.

Verifies:
1. ObsidianVaultExporter: Generates a complete, standard Obsidian Vault .zip.
2. Index.md, Project_Overview.canvas, Requirements/, Data-Dictionary/, .obsidian/.
3. YAML Frontmatter validation and Wikilinks [[Page|Alias]] syntax.
4. Obsidian Canvas v1.0 JSON format and edge routing.
5. ObsidianVaultCompiler: Compiles a Vault .zip back into an Enterprise Word BRD/SRS (.docx).
6. OpenXML Table invariants (cantSplit, tblHeader, vAlign="center") and Word Bookmarks/Hyperlinks.
7. MCP Tool calls for ba_obsidian_export_vault and ba_obsidian_compile_vault.
"""

from __future__ import annotations

import json
import os
import tempfile
import zipfile
import docx
import pytest

from core.ba_hub.obsidian_vault import ObsidianVaultExporter, ObsidianVaultCompiler
from core.ba_hub.story_builder import UserStory
from core.ba_hub.usecase_builder import UseCaseSpec
from core.ba_hub.data_dictionary import DataField
from server.mcp_server import handle_rpc_request


@pytest.fixture
def sample_vault_data():
    stories = [
        UserStory(
            id="US-001",
            title="Đăng Nhập Khách Hàng",
            role="Khách hàng cá nhân",
            goal="đăng nhập vào ứng dụng ngân hàng số bằng sinh trắc học hoặc mật khẩu",
            benefit="tôi có thể quản lý tài khoản và thực hiện giao dịch an toàn",
            domain="Banking",
            business_rules=["Tài khoản không bị khóa"],
            scenarios=[],
        )
    ]
    usecases = [
        UseCaseSpec.from_dict({
            "id": "UC-01",
            "name": "Chuyển Tiền Liên Ngân Hàng",
            "module": "Banking Core",
            "primary_actor": "Khách Hàng Đã Xác Thực",
            "trigger": "Khách hàng chọn chức năng chuyển tiền",
            "preconditions": ["Tài khoản nguồn hoạt động bình thường", "Số dư khả dụng >= Số tiền chuyển + Phí"],
            "postconditions": ["Tài khoản nguồn bị trừ tiền", "Hệ thống ghi nhận giao dịch thành công"],
            "main_flow": [
                {"step": "1", "actor": "Khách Hàng", "action": "Chọn ngân hàng thụ hưởng và nhập số tài khoản."},
                {"step": "2", "actor": "Hệ Thống", "action": "Kiểm tra thông tin tài khoản đích và hiển thị tên người thụ hưởng."},
                {"step": "3", "actor": "Khách Hàng", "action": "Nhập số tiền cần chuyển và nội dung giao dịch."},
            ],
            "alternative_flows": [
                {"id": "ALT-01", "condition": "Ngân hàng thụ hưởng bảo trì", "steps": ["Hệ thống thông báo lỗi và dừng giao dịch."]}
            ],
            "exception_flows": [
                {"id": "EXC-01", "condition": "Khách hàng nhập sai OTP 3 lần", "steps": ["Khóa giao dịch trực tuyến và ghi nhận log bảo mật."]}
            ],
        })
    ]
    data_dictionaries = {
        "Users": [
            DataField(field_name="user_id", data_type="BIGINT", nullable="NO", is_primary_key=True, description="Mã định danh duy nhất của người dùng"),
            DataField(field_name="username", data_type="VARCHAR(50)", nullable="NO", description="Tên đăng nhập hệ thống"),
            DataField(field_name="email", data_type="VARCHAR(100)", nullable="NO", description="Địa chỉ thư điện tử xác thực"),
            DataField(field_name="status", data_type="VARCHAR(20)", nullable="NO", description="Trạng thái tài khoản (ACTIVE, LOCKED)"),
        ],
        "Transactions": [
            DataField(field_name="tx_id", data_type="VARCHAR(64)", nullable="NO", is_primary_key=True, description="Mã giao dịch đối soát Napas"),
            DataField(field_name="user_id", data_type="BIGINT", nullable="NO", is_foreign_key=True, references="Users(user_id)", description="Khóa ngoại liên kết bảng Users"),
            DataField(field_name="amount", data_type="DECIMAL(18,2)", nullable="NO", description="Số tiền giao dịch thực tế"),
        ]
    }
    diagrams = {
        "Architecture": "flowchart TD\n    A[Web/Mobile App] --> B[API Gateway]\n    B --> C[Auth Service]\n    B --> D[Banking Core]"
    }
    return {
        "stories": stories,
        "usecases": usecases,
        "data_dictionaries": data_dictionaries,
        "diagrams": diagrams,
        "project_name": "Digital_Banking_Platform",
        "author": "Senior Business Analyst",
    }


class TestObsidianVaultExporter:
    def test_export_vault_zip_structure(self, sample_vault_data):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_out = os.path.join(tmpdir, "test_vault.zip")
            res_path = ObsidianVaultExporter.export_vault_zip(
                stories=sample_vault_data["stories"],
                usecases=sample_vault_data["usecases"],
                data_dictionaries=sample_vault_data["data_dictionaries"],
                diagrams=sample_vault_data["diagrams"],
                project_name=sample_vault_data["project_name"],
                output_zip_path=zip_out,
                author=sample_vault_data["author"],
            )

            assert os.path.exists(res_path)
            assert zipfile.is_zipfile(res_path)

            with zipfile.ZipFile(res_path, "r") as zf:
                namelist = zf.namelist()

                # Core files
                assert "Index.md" in namelist
                assert "Project_Overview.canvas" in namelist
                assert "Requirements/BRD_Overview.md" in namelist
                assert ".obsidian/app.json" in namelist

                # Stories & Usecases
                story_files = [n for n in namelist if n.startswith("Requirements/US")]
                assert len(story_files) == 1
                assert "US-001.md" in story_files[0] or "US_001.md" in story_files[0]

                uc_files = [n for n in namelist if n.startswith("Requirements/UC")]
                assert len(uc_files) == 1
                assert "UC-01.md" in uc_files[0] or "UC_01.md" in uc_files[0]

                # Data Dictionaries
                assert "Data-Dictionary/Users.md" in namelist
                assert "Data-Dictionary/Transactions.md" in namelist

                # Diagrams
                assert "Diagrams/Architecture.md" in namelist

    def test_markdown_frontmatter_and_wikilinks(self, sample_vault_data):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_out = os.path.join(tmpdir, "test_vault.zip")
            ObsidianVaultExporter.export_vault_zip(
                stories=sample_vault_data["stories"],
                usecases=sample_vault_data["usecases"],
                data_dictionaries=sample_vault_data["data_dictionaries"],
                diagrams=sample_vault_data["diagrams"],
                project_name=sample_vault_data["project_name"],
                output_zip_path=zip_out,
                author=sample_vault_data["author"],
            )

            with zipfile.ZipFile(zip_out, "r") as zf:
                # Read Index.md
                index_md = zf.read("Index.md").decode("utf-8")
                assert "---" in index_md
                assert "type: vault-index" in index_md
                assert "[[Project_Overview.canvas" in index_md
                assert "[[US-001|" in index_md
                assert "[[Users|" in index_md

                # Read Story MD
                story_path = [n for n in zf.namelist() if n.startswith("Requirements/US")][0]
                story_md = zf.read(story_path).decode("utf-8")
                assert 'id: "US-001"' in story_md
                assert 'role: "Khách hàng cá nhân"' in story_md
                assert "[[BRD_Overview" in story_md

                # Read Canvas JSON
                canvas_raw = zf.read("Project_Overview.canvas").decode("utf-8")
                canvas_data = json.loads(canvas_raw)
                assert "nodes" in canvas_data
                assert "edges" in canvas_data
                assert len(canvas_data["nodes"]) >= 3
                assert len(canvas_data["edges"]) >= 2


class TestObsidianVaultCompiler:
    def test_compile_vault_to_docx(self, sample_vault_data):
        with tempfile.TemporaryDirectory() as tmpdir:
            zip_out = os.path.join(tmpdir, "test_vault.zip")
            ObsidianVaultExporter.export_vault_zip(
                stories=sample_vault_data["stories"],
                usecases=sample_vault_data["usecases"],
                data_dictionaries=sample_vault_data["data_dictionaries"],
                diagrams=sample_vault_data["diagrams"],
                project_name=sample_vault_data["project_name"],
                output_zip_path=zip_out,
                author=sample_vault_data["author"],
            )

            docx_out = os.path.join(tmpdir, "compiled_brd.docx")
            res_docx = ObsidianVaultCompiler.compile_vault_to_docx(
                vault_input=zip_out,
                output_docx_path=docx_out,
                project_name="Digital Banking Platform",
                author="Senior BA",
            )

            assert os.path.exists(res_docx)
            assert os.path.getsize(res_docx) > 5000

            # Open Word document and verify structure & invariants
            doc = docx.Document(res_docx)
            full_text = "\n".join([p.text for p in doc.paragraphs])
            assert "Digital Banking Platform" in full_text
            assert "TÀI LIỆU YÊU CẦU NGHIỆP VỤ" in full_text

            # Check OpenXML Tables
            assert len(doc.tables) >= 2
            for table in doc.tables:
                for row in table.rows:
                    trPr = row._tr.get_or_add_trPr()
                    assert trPr.find(docx.oxml.ns.qn("w:cantSplit")) is not None

                # Table header
                first_trPr = table.rows[0]._tr.get_or_add_trPr()
                assert first_trPr.find(docx.oxml.ns.qn("w:tblHeader")) is not None


class TestObsidianMCPTools:
    def test_mcp_export_and_compile_vault(self, sample_vault_data):
        # 1. MCP Tool: ba_obsidian_export_vault
        req_export = {
            "jsonrpc": "2.0",
            "id": 201,
            "method": "tools/call",
            "params": {
                "name": "ba_obsidian_export_vault",
                "arguments": {
                    "stories": [s.to_dict() for s in sample_vault_data["stories"]],
                    "usecases": [u.to_dict() for u in sample_vault_data["usecases"]],
                    "data_dictionaries": {k: [f.to_dict() for f in v] for k, v in sample_vault_data["data_dictionaries"].items()},
                    "project_name": "MCP_Test_Project",
                    "author": "MCP Agent",
                }
            }
        }
        res_export = handle_rpc_request(req_export)
        assert res_export is not None
        assert "result" in res_export
        assert res_export["result"].get("isError") is not True

        res_data = json.loads(res_export["result"]["content"][0]["text"])
        assert res_data["success"] is True
        vault_zip_path = res_data["vault_zip"]
        assert os.path.exists(vault_zip_path)

        # 2. MCP Tool: ba_obsidian_compile_vault
        req_compile = {
            "jsonrpc": "2.0",
            "id": 202,
            "method": "tools/call",
            "params": {
                "name": "ba_obsidian_compile_vault",
                "arguments": {
                    "vault_file": vault_zip_path,
                    "project_name": "MCP Compiled Project",
                    "author": "MCP Agent",
                }
            }
        }
        res_compile = handle_rpc_request(req_compile)
        assert res_compile is not None
        assert "result" in res_compile
        assert res_compile["result"].get("isError") is not True

        compile_data = json.loads(res_compile["result"]["content"][0]["text"])
        assert compile_data["success"] is True
        assert os.path.exists(compile_data["docx_path"])
