# -*- coding: utf-8 -*-
"""
core/pipeline/engine.py — Declarative Multi-Step Document Workflow Pipeline & Bulk Runner (v4.6).

Enforces Guard #4 (ERR_PIPE_001):
  1. Memory-Safe Concurrency: Defaults to max_workers = int(os.environ.get("MAX_JOB_WORKERS", 2))
     to prevent Out-Of-Memory (OOM) crashes on 512MB RAM cloud containers (Render.com).
  2. Deterministic Intermediate File Cleanup: Every intermediate file (step_1_temp.docx,
     step_2_temp.pdf, ...) is tracked and strictly removed inside a try...finally block
     immediately after each file finishes or fails the pipeline.
"""

from __future__ import annotations

import concurrent.futures
import fnmatch
import json
import os
import shutil
import time
from typing import Any, Dict, List, Optional, Union

import yaml


def get_default_max_workers() -> int:
    """Returns memory-safe worker count from MAX_JOB_WORKERS env var (default: 2)."""
    try:
        val = int(os.environ.get("MAX_JOB_WORKERS", "2"))
        return max(1, val)
    except (ValueError, TypeError):
        return 2


def _is_zombie_or_lock_file(filename: str) -> bool:
    """ERR_CONV_004: Ignore Office lock files (~$*, .~lock.*) and temp artifacts."""
    base = os.path.basename(filename)
    return (
        base.startswith("~$")
        or base.startswith(".~lock.")
        or "_temp" in base
        or base.endswith(".tmp")
    )


class WorkflowPipeline:
    """Executes a declarative sequence of document processing steps on a source file."""

    def __init__(self, spec: Union[Dict[str, Any], str]):
        if isinstance(spec, str):
            if os.path.isfile(spec):
                with open(spec, "r", encoding="utf-8") as f:
                    if spec.lower().endswith((".yaml", ".yml")):
                        self.spec = yaml.safe_load(f) or {}
                    else:
                        self.spec = json.load(f)
            else:
                try:
                    self.spec = json.loads(spec)
                except Exception:
                    self.spec = yaml.safe_load(spec) or {}
        elif isinstance(spec, dict):
            self.spec = dict(spec)
        else:
            raise ValueError("Pipeline spec must be a dict, JSON/YAML string, or file path.")

        self.name: str = str(self.spec.get("name", "Antigravity Document Pipeline"))
        self.steps: List[Dict[str, Any]] = list(self.spec.get("steps", []))
        if not self.steps:
            raise ValueError("Pipeline spec must contain a non-empty 'steps' list.")

    def _determine_step_ext(self, current_ext: str, action: str, params: Dict[str, Any]) -> str:
        """Determines the output file extension produced by a given pipeline step."""
        act = action.strip().lower().replace("-", "_")
        if act == "convert":
            t_fmt = str(params.get("target_format") or params.get("target_ext") or ".pdf").strip().lower()
            return t_fmt if t_fmt.startswith(".") else f".{t_fmt}"
        if act == "xlsx_to_brief":
            t_fmt = str(params.get("target_format") or "docx").strip().lower().lstrip(".")
            return f".{t_fmt}"
        if act in ("pdf_watermark", "pdf_giap_lai", "pdf_sign", "pdf_crypto_sign", "pdf_flatten", "pdf_to_pdfa"):
            return ".pdf"
        return current_ext

    def _execute_single_step(
        self,
        action: str,
        input_file: str,
        output_file: str,
        params: Dict[str, Any],
    ) -> Dict[str, Any]:
        """Dispatches a single step action to its corresponding core engine."""
        act = action.strip().lower().replace("-", "_")

        if act == "convert":
            from core.converter.engine import convert_universal

            target_fmt = params.get("target_format") or params.get("target_ext")
            ok, out_p, elapsed = convert_universal(
                source_path=input_file,
                target_path=output_file,
                target_ext=target_fmt,
                page_range=params.get("page_range"),
                style_path=params.get("style_path"),
            )
            if not ok:
                raise RuntimeError(f"Step 'convert' failed: {out_p}")
            return {"output_file": out_p, "elapsed": elapsed}

        if act == "docx_sanitize_vietnamese":
            from core.docx.vietnamese_sanitizer import sanitize_docx_vietnamese

            shutil.copy2(input_file, output_file)
            res = sanitize_docx_vietnamese(output_file, output_path=output_file)
            return {"output_file": output_file, "details": res}

        if act == "docx_apply_preset":
            from core.docx.nd30_preset import apply_nd30_preset

            shutil.copy2(input_file, output_file)
            res = apply_nd30_preset(
                output_file,
                output_path=output_file,
                org_name=params.get("org_name", "CƠ QUAN CHỦ QUẢN"),
                dept_name=params.get("dept_name", "ĐƠN VỊ BAN HÀNH"),
                doc_number=params.get("doc_number", "01/BC-ĐV"),
                location_date=params.get("location_date"),
            )
            return {"output_file": output_file, "details": res}

        if act == "docx_replace":
            from core.docx.writer import patch_docx_template

            replacements = params.get("replacements") or params.get("data") or {}
            res = patch_docx_template(input_file, output_file, replacements=replacements)
            return {"output_file": output_file, "details": res}

        if act in ("doc_security_sanitize", "security_sanitize", "deep_cdr", "sanitize"):
            from core.security.sanitizer import sanitize_document

            res = sanitize_document(
                input_file,
                output_path=output_file,
                defuse_formulas=bool(params.get("defuse_formulas", True)),
                strip_metadata=bool(params.get("strip_metadata", True)),
                unhide_sheets=bool(params.get("unhide_sheets", True)),
                deep_cdr=bool(params.get("deep_cdr", True)),
                redact_pii=bool(params.get("redact_pii", False)),
                pii_types=params.get("pii_types"),
            )
            if not res.get("success"):
                raise RuntimeError(res.get("error", "Sanitization failed"))
            return {"output_file": res["output_file"], "metrics": res.get("metrics")}

        if act in ("doc_redact_pii", "redact_pii"):
            from core.security.sanitizer import sanitize_document

            res = sanitize_document(
                input_file,
                output_path=output_file,
                defuse_formulas=False,
                strip_metadata=bool(params.get("strip_metadata", False)),
                unhide_sheets=False,
                deep_cdr=False,
                redact_pii=True,
                pii_types=params.get("pii_types"),
            )
            return {"output_file": res["output_file"], "metrics": res.get("metrics")}

        if act == "pdf_watermark":
            from core.converter.pdf_toolkit import stamp_watermark_pdf

            out_p = stamp_watermark_pdf(
                input_file,
                output_file,
                watermark_text=str(params.get("text", "CONFIDENTIAL")),
                opacity=float(params.get("opacity", 0.18)),
            )
            return {"output_file": out_p}

        if act == "pdf_giap_lai":
            from core.converter.pdf_toolkit import stamp_giap_lai_pdf

            stamp_path = params.get("stamp_path") or params.get("seal_path")
            if not stamp_path:
                raise ValueError("pdf_giap_lai requires 'stamp_path' parameter.")
            res = stamp_giap_lai_pdf(
                input_file,
                output_file,
                seal_image_path=stamp_path,
                position=str(params.get("position", "right")),
            )
            return {"output_file": output_file, "details": res}

        if act == "pdf_sign":
            from core.converter.pdf_toolkit import sign_pdf_with_stamp

            signatures = params.get("signatures") or []
            res = sign_pdf_with_stamp(input_file, output_file, signatures=signatures)
            return {"output_file": output_file, "details": res}

        if act == "pdf_crypto_sign":
            from core.converter.pdf_toolkit import crypto_sign_pdf

            res = crypto_sign_pdf(
                input_file,
                output_file,
                signer_name=str(params.get("signer_name", "Antigravity Authority")),
                reason=str(params.get("reason", "Phê duyệt tài liệu tự động")),
                location=str(params.get("location", "Việt Nam")),
                pfx_path=params.get("pfx_path"),
                pfx_password=params.get("pfx_password"),
                stamp_image=params.get("stamp_image"),
                page=params.get("page", "last"),
            )
            return {"output_file": output_file, "details": res}

        if act == "pdf_flatten":
            from core.converter.pdf_toolkit import flatten_and_sanitize_pdf

            res = flatten_and_sanitize_pdf(
                input_file,
                output_file,
                flatten_vector=bool(params.get("flatten_vector", True)),
                strict_raster=bool(params.get("strict_raster", False)),
                sanitize_metadata=bool(params.get("sanitize_metadata", True)),
            )
            return {"output_file": output_file, "details": res}

        if act == "pdf_to_pdfa":
            from core.converter.pdf_toolkit import convert_to_pdfa

            res = convert_to_pdfa(
                input_file,
                output_file,
                part=int(params.get("part", 2)),
                conformance=str(params.get("conformance", "B")),
            )
            return {"output_file": output_file, "details": res}

        if act == "xlsx_to_brief":
            from core.xlsx.reporting import excel_to_brief

            res = excel_to_brief(
                input_file,
                output_path=output_file,
                target_format=str(params.get("target_format", "docx")),
                title=params.get("title"),
            )
            return {"output_file": res["output_path"], "details": res}

        if act in ("qa_audit", "audit_qa", "security_audit", "audit"):
            from core.security.document_auditor import audit_document_security

            shutil.copy2(input_file, output_file)
            audit_res = audit_document_security(output_file)
            return {"output_file": output_file, "audit": audit_res}

        raise ValueError(f"Unsupported pipeline step action: '{action}'")

    def run(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        work_dir: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Executes all configured steps sequentially on input_path.
        Enforces Guard #4 (ERR_PIPE_001): Deterministically cleans up all intermediate
        step_*_temp.* files inside a try...finally block.
        """
        t0 = time.time()
        src = os.path.abspath(input_path)
        if not os.path.isfile(src):
            raise FileNotFoundError(f"Pipeline input file not found: {src}")

        base_dir = os.path.abspath(work_dir) if work_dir else os.path.dirname(src)
        os.makedirs(base_dir, exist_ok=True)

        # Compute final extension after all steps
        curr_ext = os.path.splitext(src)[1].lower()
        for st in self.steps:
            curr_ext = self._determine_step_ext(curr_ext, st.get("action", ""), st.get("params", {}))

        if not output_path:
            stem = os.path.splitext(os.path.basename(src))[0]
            output_path = os.path.join(base_dir, f"{stem}_pipeline_out{curr_ext}")
        final_dest = os.path.abspath(output_path)
        os.makedirs(os.path.dirname(final_dest), exist_ok=True)

        intermediate_files: List[str] = []
        step_traces: List[Dict[str, Any]] = []
        current_input = src
        pid_tag = f"{os.getpid()}_{int(time.time() * 1000) % 100000}"

        try:
            total_steps = len(self.steps)
            for idx, step_cfg in enumerate(self.steps, start=1):
                action = str(step_cfg.get("action", "")).strip()
                params = dict(step_cfg.get("params") or {})
                for k, v in step_cfg.items():
                    if k not in ("action", "params") and k not in params:
                        params[k] = v
                if not action:
                    raise ValueError(f"Step #{idx} is missing 'action'.")

                in_ext = os.path.splitext(current_input)[1].lower()
                out_ext = self._determine_step_ext(in_ext, action, params)

                if idx < total_steps:
                    step_out = os.path.join(base_dir, f"step_{idx}_temp_{pid_tag}{out_ext}")
                    intermediate_files.append(step_out)
                else:
                    step_out = final_dest

                st_t0 = time.time()
                res = self._execute_single_step(action, current_input, step_out, params)
                actual_out = os.path.abspath(res.get("output_file", step_out))
                if idx < total_steps and actual_out != step_out and actual_out != src:
                    intermediate_files.append(actual_out)

                current_input = actual_out
                step_traces.append({
                    "step": idx,
                    "action": action,
                    "output_file": actual_out,
                    "elapsed_sec": round(time.time() - st_t0, 3),
                })

            if os.path.abspath(current_input) != final_dest and os.path.isfile(current_input):
                shutil.copy2(current_input, final_dest)

            return {
                "success": True,
                "pipeline_name": self.name,
                "input_file": src,
                "output_file": final_dest,
                "output_path": final_dest,
                "steps_executed": len(step_traces),
                "step_traces": step_traces,
                "elapsed_sec": round(time.time() - t0, 3),
            }
        finally:
            # Guard #4 (ERR_PIPE_001): Guaranteed cleanup of all intermediate step_*_temp files
            for tmp_path in intermediate_files:
                if os.path.isfile(tmp_path) and os.path.abspath(tmp_path) != final_dest and os.path.abspath(tmp_path) != src:
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass


def run_pipeline(
    input_path: str,
    pipeline_spec: Union[Dict[str, Any], str],
    output_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Convenience wrapper to run a WorkflowPipeline on a single input file."""
    pipeline = WorkflowPipeline(pipeline_spec)
    return pipeline.run(input_path, output_path=output_path)


def run_bulk_directory_pipeline(
    input_dir: str,
    output_dir: str,
    pipeline_spec: Optional[Union[Dict[str, Any], str]] = None,
    file_pattern: str = "*.*",
    max_workers: Optional[int] = None,
    spec: Optional[Union[Dict[str, Any], str]] = None,
) -> Dict[str, Any]:
    """
    Executes a WorkflowPipeline across all matching files in input_dir.
    Enforces Guard #4 (ERR_PIPE_001): Defaults max_workers to int(os.environ.get('MAX_JOB_WORKERS', 2)).
    """
    t0 = time.time()
    resolved_spec = pipeline_spec if pipeline_spec is not None else spec
    if resolved_spec is None:
        raise ValueError("pipeline_spec or spec is required")

    src_dir = os.path.abspath(input_dir)
    if not os.path.isdir(src_dir):
        raise FileNotFoundError(f"Bulk input directory not found: {src_dir}")

    dest_dir = os.path.abspath(output_dir)
    os.makedirs(dest_dir, exist_ok=True)

    workers = max_workers if (max_workers is not None and max_workers >= 1) else get_default_max_workers()
    pipeline = WorkflowPipeline(resolved_spec)

    candidate_files: List[str] = []
    for entry in sorted(os.listdir(src_dir)):
        full_p = os.path.join(src_dir, entry)
        if not os.path.isfile(full_p):
            continue
        if _is_zombie_or_lock_file(entry):
            continue
        if fnmatch.fnmatch(entry.lower(), file_pattern.lower()):
            candidate_files.append(full_p)

    results: List[Dict[str, Any]] = []
    succeeded = 0
    failed = 0

    def _process_one(f_path: str) -> Dict[str, Any]:
        stem, orig_ext = os.path.splitext(os.path.basename(f_path))
        final_ext = orig_ext.lower()
        for st in pipeline.steps:
            final_ext = pipeline._determine_step_ext(final_ext, st.get("action", ""), st.get("params", {}))
        out_f = os.path.join(dest_dir, f"{stem}{final_ext}")
        return pipeline.run(f_path, output_path=out_f, work_dir=dest_dir)

    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        future_map = {executor.submit(_process_one, fp): fp for fp in candidate_files}
        for fut in concurrent.futures.as_completed(future_map):
            src_f = future_map[fut]
            try:
                res = fut.result()
                results.append(res)
                succeeded += 1
            except Exception as exc:
                failed += 1
                results.append({
                    "success": False,
                    "input_file": src_f,
                    "error": str(exc),
                })

    return {
        "success": failed == 0,
        "pipeline_name": pipeline.name,
        "input_dir": src_dir,
        "output_dir": dest_dir,
        "max_workers": workers,
        "max_workers_used": workers,
        "total_files": len(candidate_files),
        "succeeded": succeeded,
        "failed": failed,
        "results": results,
        "elapsed_sec": round(time.time() - t0, 3),
    }
