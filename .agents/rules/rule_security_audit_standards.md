# Rule: Codebase & Document Security Audit Standards (Cloudflare Invariants)

> **Scope**: Mandatory security principles, adversarial candidate validation, and document safety invariants (`ERR_SEC_001`–`005`) for all code reviews, document processing, and agent audits.

---

## 1. Adversarial Candidate Verification Invariants

When an AI Agent reviews code or audits a target repository:

| # | Invariant | Rule Description |
|---|---|---|
| **S1** | **No Unverified Confirmations** | An agent MUST NOT mark any vulnerability as `confirmed` without an independent adversarial verifier explicitly attempting to disprove it. |
| **S2** | **Source Trace Completeness** | Every `confirmed` finding must include the exact source taint trace from untrusted input entry point to sink. Vague suggestions without file and line numbers are forbidden. |
| **S3** | **Strict Three-Verdict Output** | Findings must strictly partition into `confirmed` (proven with bounded reproduction), `needs_validation` (missing runtime context, no severity), and `rejected` (disproven with rationale). |
| **S4** | **Zero-Dependency Validation** | The parent agent must run local schema validators (`validate-findings.cjs`, `validate-coverage-ledger.cjs`) to verify ledger integrity without relying on LLM memory. |
| **S5** | **Read-Only Sandbox Isolation** | Source code inspection is strictly read-only. Dynamic reproductions must run inside bounded, isolated scratch directories with no external network access. |

---

## 2. Document & Office Automation Security Invariants (ERR_SEC)

When parsing, converting, or generating office documents (`.docx`, `.xlsx`, `.pptx`, `.pdf`, `.zip`):

| Invariant Code | Name | Mandatory Rule & Verification |
|---|---|---|
| **`ERR_SEC_001`** | **Macro & Malicious Code Execution** | OpenXML documents must be scanned for macro streams (`vbaProject.bin`, `word/vbaData.xml`, XLM macros). Uninspected macro code must be flagged or stripped. |
| **`ERR_SEC_002`** | **Spreadsheet Formula Injection / DDE** | Any cell text starting with `=`, `+`, `-`, or `@` that invokes dangerous functions (`cmd|`, `powershell|`, `mshta|`, `WEBSERVICE`, `HYPERLINK` to external protocols) must be neutralized by prepending `'` upon sanitization. |
| **`ERR_SEC_003`** | **Dangerous OLE Package Embeddings** | Embedded objects under `embeddings/oleObject*.bin` must be scanned for PE headers (`MZ`) or executable script extensions (`.exe`, `.bat`, `.vbs`, `.ps1`). |
| **`ERR_SEC_004`** | **Hidden PII & Sensitive Metadata Leakage** | Documents prepared for public delivery must strip author usernames (`dc:creator`), comments, revision history markup, and restore or unhide `veryHidden` sheets. |
| **`ERR_SEC_005`** | **Safe XML & Zip Slip Defense** | Archive extraction must verify `os.path.commonpath([dest, target]) == dest`. XML parsers must disable DTD external entities (`<!ENTITY`, `SYSTEM`) to prevent XXE. |

---

## 3. Safe Sanitization Code Pattern

```python
import re

def sanitize_formula_injection(val: str) -> str:
    """Neutralize DDE / Formula Injection without destroying user data."""
    if not isinstance(val, str):
        return val
    # Dangerous trigger prefixes
    if re.match(r"^[=+\-@]", val):
        # Look for shell/network triggers
        if re.search(r"(cmd|powershell|mshta|cscript|wscript|calc|webservice|iex)", val, re.IGNORECASE):
            return f"'{val}"
    return val
```
