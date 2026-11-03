---
name: ba-workspace
description: Enterprise Business Analysis Workstation operations guide, mental model, and decision tree for BABOK v3 / IEEE 830 BRD/SRS synthesis, User Story & Gherkin AC linter, Cockburn Use Cases, RTM traceability, DBML ERD, Satsuma data lineage, SQL integrity verification, Scope Creep audit, Obsidian Vault two-way compilation, OpenAPI contract scanner, Decision Tables, Meeting Notes parser, and PlantUML Salt wireframes.
---

# Antigravity BA Enterprise Workstation Skill (`ba-workspace`)

Use this skill whenever you need to elicit, structure, validate, trace, audit, or publish software requirements, data architectures, and business logic into publication-grade documents (Word `.docx`, Excel `.xlsx`, Markdown `.md`, Obsidian Vault `.zip`, and PlantUML Salt UI Wireframes).

---

## 🧠 Core Philosophy & Mental Model

### 1. The BABOK v3 & INVEST Specification Standard
- **INVEST Principle for User Stories**: Every story must be Independent, Negotiable, Valuable, Estimable, Small, and Testable.
- **Three-Part User Story Structure**:
  - `As a <Persona>`: Must be a specific business role (e.g., `Khách hàng cá nhân đã eKYC`, NOT vague `Người dùng`).
  - `I want <Action/Feature>`: Clear business capability.
  - `So that <Business Value>`: Explicit ROI, operational saving, or user outcome.
- **Three-Tier Gherkin Acceptance Criteria (AC)**:
  - `Happy Path (Scenario 1)`: Standard expected success flow.
  - `Negative Path (Scenario 2)`: Validation errors, timeouts, or rejection handling.
  - `Boundary Conditions (Scenario 3)`: Numeric limits, concurrency thresholds, or maximum quotas.

### 2. Complete Requirements Traceability Matrix (RTM)
- Strict bidirectional mapping: **Business Need &rarr; Epic &rarr; User Story &rarr; Use Case &rarr; API Endpoint / DB Table &rarr; UAT Test Case &rarr; Defect / Sign-off**.
- Zero orphan requirements: Every functional requirement must map to at least one test case and at least one system component.

### 3. Local-First Knowledge Mesh (Obsidian Vault Integration)
- Every requirement artifact is compiled into a local-first Obsidian Vault with clean two-way `[[Wikilinks]]`.
- Bidirectional Compilation:
  - Export: BA App &rarr; `.zip` Obsidian Vault with Graph View relationships.
  - Ingest: `.zip` Obsidian Vault &rarr; Analyzes Wikilinks graph &rarr; Compiles into OpenXML Word `.docx` with native `<w:bookmarkStart>` and `<w:hyperlink>`.

---

## 🌲 Decision Tree & Use Cases

```
                            [Business Analysis Task]
                                       │
        ┌──────────────────────────────┼──────────────────────────────┐
        ▼                              ▼                              ▼
 [Requirements & Stories]     [Data & Architecture]         [Audit & Delivery]
        │                              │                              │
 ┌──────┴──────┐                ┌──────┴──────┐                ┌──────┴──────┐
 ▼             ▼                ▼             ▼                ▼             ▼
User Stories  Use Cases      DBML / DDL    OpenAPI Scan     RTM Matrix    Obsidian Vault
(INVEST Linter)(Cockburn)    (ERD 2-Way)   (Data Dict)      (Traceability)(Word Compiler)
```

### Case 1: Linter & Thẩm Định Chất Lượng User Story
- **Scenario**: Đánh giá tính hoàn thiện của User Story theo chuẩn BABOK v3 / IEEE 830 trước khi đưa vào Sprint Backlog.
- **CLI Command**:
  ```bash
  python ai_tools_cli.py ba-story-lint --story-id "US-PAY-001" --title "Chuyển tiền nhanh 24/7" \
    --persona "Khách hàng cá nhân đã xác thực eKYC" \
    --action "chuyển tiền liên ngân hàng qua số thẻ hoặc số tài khoản" \
    --value "người thụ hưởng nhận tiền ngay lập tức 24/7 không cần đến quầy" \
    --ac "Kịch bản 1: Số dư đủ và hạn mức hợp lệ..." \
    --rules "Hạn mức tối đa 500 triệu/lần"
  ```
- **Success Signal**: JSON trả về `score >= 80`, `grade: "EXCELLENT"|"GOOD"`, không có lỗi critical.

### Case 2: Bóc Tách API Contract (OpenAPI / Swagger) Sang Data Dictionary
- **Scenario**: Ingest tệp OpenAPI/Swagger (`openapi.json` hoặc `swagger.yaml`) thành bảng từ điển dữ liệu chuẩn OpenXML Word/Excel.
- **CLI Command**:
  ```bash
  python ai_tools_cli.py ba-openapi swagger.json -o data_dict.xlsx --format excel
  ```

### Case 3: Kiểm Toán Bảng Quyết Định (Decision Table) BABOK v3
- **Scenario**: Xác thực ma trận logic điều kiện - hành động (Condition / Action Matrix), kiểm tra 100% không bị mâu thuẫn (Contradiction) và không bị bỏ sót kịch bản (Completeness).
- **CLI Command**:
  ```bash
  python ai_tools_cli.py ba-decision-table spec.json -o decision_matrix.docx --audit
  ```

### Case 4: Chuyển Biên Bản Họp (Meeting Notes) Sang User Stories & Rules
- **Scenario**: Phân tích ghi chép thô sau buổi họp nghiệp vụ với Stakeholders thành các User Stories 3 thành phần, Quy tắc nghiệp vụ và Action Items.
- **CLI Command**:
  ```bash
  python ai_tools_cli.py ba-meeting-parser raw_meeting_notes.md -o meeting_analysis.docx
  ```

### Case 5: Vẽ Phác Thảo Giao Diện (PlantUML Salt UI Wireframe)
- **Scenario**: Tạo mockup wireframe màn hình đăng nhập, dashboard hoặc bảng dữ liệu từ văn bản phác thảo thuần túy.
- **CLI Command**:
  ```bash
  python ai_tools_cli.py wireframe-render wireframe_spec.json -o wireframe.png --scale 2
  ```

### Case 6: Hai Chiều Obsidian Vault ↔ OpenXML Word BRD
- **Scenario**: Đóng gói toàn bộ tài liệu dự án thành Obsidian Vault `.zip` có đồ thị Graph View, hoặc biên dịch tệp Vault `.zip` thành tài liệu Word BRD 9 phần hoàn chỉnh.
- **CLI Commands**:
  ```bash
  # Xuất Obsidian Vault:
  python ai_tools_cli.py ba-obsidian-export --project "Digital_Banking" -o Digital_Banking_Vault.zip

  # Biên dịch ngược sang Word BRD:
  python ai_tools_cli.py ba-obsidian-compile Digital_Banking_Vault.zip -o BRD_Digital_Banking.docx
  ```

---

## 📋 OpenXML Table & Format Invariants for BA Artifacts

1. **Table Invariants**:
   - Mọi bảng đặc tả (Data Dictionary, RTM, Decision Table, Use Case Steps) bắt buộc phải có `<w:cantSplit/>`, `<w:tblHeader/>`, và `<w:vAlign w:val="center"/>`.
2. **Identifier Columns**:
   - Cột Story ID, Field Code, API Path trong Excel bắt buộc thiết lập `number_format = '@'` và kiểu chuỗi (`str`).
3. **Traceability Links**:
   - Các mã liên kết chéo (`US-xxx`, `UC-xxx`, `TC-xxx`) khi xuất Word phải được neo bằng `<w:bookmarkStart>` và hyperlink `<w:hyperlink w:anchor="...">`.

