# Phase 2: Coverage-Led Hunting

Coverage-Led Hunting enforces systematic, non-overlapping code review. Instead of random guessing or relying on prompt intuition, hunters evaluate each assigned unit from `coverage-ledger.json` against relevant attack classes.

## Rules of Hunting

1. **Unit Assignment**:
   - Each hunter is assigned a specific set of files/units from `coverage-ledger.json`.
   - The hunter must inspect the unit thoroughly and record either candidate findings or a clean check with checked attack classes.
2. **Selective Domain Application**:
   - Only apply attack classes relevant to the unit's technology stack:
     - For Web Server (`server/app.py`): Focus on Path Traversal, SSRF, DoS, Header Injection, LFI.
     - For Document Parsers (`core/docx/`, `core/xlsx/`, `core/pdf/`): Focus on Zip Slip, XXE, Entity Expansion, Formula Injection, Deserialization, Memory Exhaustion.
     - For Frontend (`src/`): Focus on DOM-based XSS, Content Injection, Prototype Pollution.
3. **Candidate Recording (Do Not Confirm Prematurely)**:
   - When a suspicious flow is detected, create a **Candidate Record**:
     - `candidate_id`: Stable slug (e.g. `cand-app-path-traversal-01`).
     - `file`: Path to source file.
     - `line_range`: Exact line numbers.
     - `attack_class`: Name of attack class.
     - `source_trace`: Taint flow from input entry to sink.
     - `claimed_impact`: Potential violation of trust boundary.
   - **Crucial Rule**: Hunters NEVER mark a finding as `confirmed`. Only an independent verifier can do that.
4. **Updating the Ledger**:
   - Once a unit has been reviewed, update its status in `coverage-ledger.json` to `covered`.
   - Record the timestamp, checked attack classes, and any generated candidate IDs.
