# Phase 3 & Phase 6: Validation, Schema & Reporting

This document defines the core adversarial protocol that distinguishes Cloudflare's harness from standard prompt-based audits.

## Phase 3: Adversarial Candidate Validation

Every candidate finding generated in Phase 2 must be delegated to a **fresh, independent Verifier Sub-Agent**.

### Mandate of the Verifier Sub-Agent
The verifier's explicit goal is **adversarial**:
- **Attempt to DISPROVE (bác bỏ) the candidate finding**.
- Question every premise:
  - Is the input sanitized by a prior layer, router, or framework middleware?
  - Does the language runtime or type system prevent this attack condition?
  - Is the sink actually reachable by an untrusted actor across the trust boundary?
  - Is the affected data synthetic, internal-only, or non-sensitive?
  - Does defense-in-depth prevent exploitation in realistic deployments?

### Outcomes of Verification

1. **`confirmed`**:
   - The verifier was **unable to disprove** the candidate.
   - The verifier produced a concrete, deterministic reproduction path or proof trace.
   - The finding includes: Title, Severity (Critical, High, Medium, Low), Affected Trust Boundary, Source Trace, Proof-of-Concept or reproduction steps, and the Smallest Effective Fix.
2. **`rejected`**:
   - The verifier successfully disproved the vulnerability.
   - The finding records the exact reason for rejection (e.g. `Sanitized by resolve_file_path() at line 42`).
   - Recorded in `findings.json` under `rejected` to prevent duplicate re-investigation in future runs.
3. **`needs_validation`**:
   - The validity depends on an environmental fact that cannot be verified statically (e.g. external network configuration, specific production OS permissions).
   - Must specify the **exact unresolved fact** and a safe validation procedure.
   - Must NOT assign a severity score.

---

## Phase 6: Target-Neutral Reporting

After validation and running zero-dependency schema checks (`validate-findings.cjs`), compile final reports:

1. **`REPORT.md`**: Executive summary, trust boundary analysis, statistics by severity, and list of confirmed findings.
2. **`FINDINGS-DETAIL.md`**: In-depth analysis of each confirmed vulnerability with full code snippets, trace evidence, and diffs for smallest effective fix.
3. **`NEEDS-VALIDATION.md`**: List of suspicious patterns requiring runtime environment verification.
