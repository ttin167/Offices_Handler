---
name: security-audit
description: Security guidance, vulnerability review, and adversarial candidate verification for codebases, APIs, office documents, parsers, and toolchains based on Cloudflare's security-audit harness.
---

# Security Audit Skill (Cloudflare-Seeded Harness)

Find vulnerabilities that violate a real trust boundary, then give owners source evidence, safe reproduction, priority, and the smallest effective fix. This is a defensive, source-first workflow. A candidate without a concrete affected principal, resource, or security outcome is not a confirmed finding.

## Operating Modes

This skill operates in two distinct modes:

- **Guidance Mode (Default)**: For security questions, focused code reviews, triage, or investigation of specific findings. Use only the relevant parts of this skill. Do not automatically run all six phases or create audit ledgers.
- **Full Audit Mode**: Use the complete workflow when the user explicitly asks to audit or pen-test a codebase, asks for a full/comprehensive security review, or requests formal audit reports. Run all six phases and maintain state ledgers.

## Platform Terminology

This skill is agent-neutral:
- **Parent Agent**: Coordinates the audit run and owns shared state (`coverage-ledger.json`, `findings.json`).
- **Hunter Agent**: Investigates specific units/subsystems and records candidate findings.
- **Verifier Agent (Adversarial Verifier)**: A fresh, independent agent whose explicit mandate is to **try and disprove (bác bỏ)** each candidate finding. Only candidates that survive adversarial verification become `confirmed`.

## The 6-Phase Audit Pipeline

```
┌──────────────────────────────────────────────────────────────────────────────────┐
│                          CLOUDFLARE 6-PHASE AUDIT HARNESS                        │
│                                                                                  │
│  Phase 1: Reconnaissance ──► Phase 2: Coverage-Led Hunting ──► Phase 3: Validate │
│  • Architecture mapping      • Ledger units (files/classes)   • Fresh Verifier   │
│  • Trust boundaries          • Target-neutral attack classes  • Disprove attempt │
│  • Entry points & parsers    • coverage-ledger.json           • Evidence bounds  │
│                                                                        │         │
│  Phase 6: Reporting ◄─────── Phase 5: Verification ◄────────── Phase 4: Schema   │
│  • REPORT.md                 • Zero-dependency script checks   • findings.json   │
│  • FINDINGS-DETAIL.md        • No model memory reliance       • 3 verdicts       │
│  • NEEDS-VALIDATION.md       • Final source verification      • report-schema    │
└──────────────────────────────────────────────────────────────────────────────────┘
```

1. **Phase 1: Reconnaissance** — Map architecture, trust boundaries, entry points, input parsers, and prior evidence into `architecture.md` and initial `coverage-ledger.json`. (See [RECONNAISSANCE.md](RECONNAISSANCE.md)).
2. **Phase 2: Coverage-Led Hunting** — Assign isolated hunters to ledger units, systematically reviewing source code against relevant attack classes. (See [HUNTING.md](HUNTING.md) and [ATTACK-CLASSES.md](ATTACK-CLASSES.md)).
3. **Phase 3: Candidate Validation (Adversarial Review)** — Every unique candidate is submitted to a fresh, independent verifier sub-agent tasked with actively attempting to disprove the vulnerability. (See [VALIDATION-AND-REPORTING.md](VALIDATION-AND-REPORTING.md)).
4. **Phase 4: Structured Schema Outputs** — Record findings into `findings.json` strictly conforming to `report-schema.json`.
5. **Phase 5: Independent Record Verification** — Execute local zero-dependency validators (`validate-findings.cjs` and `validate-coverage-ledger.cjs`) to ensure data integrity without relying on LLM memory.
6. **Phase 6: Target-Neutral Reporting** — Generate human-readable reports: `REPORT.md`, `FINDINGS-DETAIL.md`, and `NEEDS-VALIDATION.md`.

## The Three Strict Verdicts

Every evaluated item must resolve into exactly one of three states:
- **`confirmed`**: A concrete vulnerability with a complete source trace, affected trust boundary, reproducible input, and bounded impact.
- **`needs_validation`**: A suspicious pattern where a decisive fact is missing runtime context (e.g., external server environment, missing sandbox capability). Has an exact unresolved fact and NO severity score.
- **`rejected`**: A candidate disproven by the adversarial verifier (e.g., input is sanitized earlier in the pipeline, type system strictly prevents bad data, or defense-in-depth neutralizes effect).

## Office & Document Parsing Attack Classes

For document automation engines (Word, Excel, PowerPoint, PDF, ZIP), consult [OFFICE-AND-DOCUMENT-PARSERS.md](OFFICE-AND-DOCUMENT-PARSERS.md) covering:
- **`ERR_SEC_001`**: Macro & Malicious Code Execution (`vbaProject.bin`, XLM macros).
- **`ERR_SEC_002`**: Spreadsheet Formula & DDE Injection (`=cmd|`, `=powershell|`, `=WEBSERVICE`).
- **`ERR_SEC_003`**: Dangerous OLE Package & Binary Embeddings (`embeddings/oleObject*.bin`).
- **`ERR_SEC_004`**: Hidden PII & Sensitive Metadata Leakage (Revisions, comments, `veryHidden` sheets).
- **`ERR_SEC_005`**: Safe XML Parsing & Zip Slip Traversal (`../` directory traversal in archives).
