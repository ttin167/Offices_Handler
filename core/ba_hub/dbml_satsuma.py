"""
core.ba_hub.dbml_satsuma — DBML Bidirectional Engine & Satsuma Data Mapping DSL.

Inspired by @software-mansion/dbml & EqualExperts/satsuma-lang:
1. DBML Bidirectional Conversion:
   DBML <-> SQL DDL <-> Data Dictionary <-> Mermaid ERD.
2. Satsuma-style Data Mapping DSL:
   Source.Field -> Target.Field [Transformation_Rule]
3. Type Incompatibility & String Truncation Guard:
   Catches size overflows, illegal type casting, and timestamp precision loss.
4. Automated Data Lineage Graph Generation (Mermaid).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple


@dataclass
class DBMLField:
    name: str
    data_type: str
    is_pk: bool = False
    is_unique: bool = False
    not_null: bool = False
    default_val: Optional[str] = None
    note: str = ""
    ref_target: Optional[str] = None  # e.g. "users.id"


@dataclass
class DBMLTable:
    name: str
    fields: List[DBMLField] = field(default_factory=list)
    note: str = ""


@dataclass
class SatsumaMappingRule:
    source_expr: str
    source_type: str
    target_field: str
    target_type: str
    rule_type: str  # Direct, Concat, Lookup, Mask_PII, Cast, Custom
    logic_desc: str = ""
    is_valid: bool = True
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "source": self.source_expr,
            "source_type": self.source_type,
            "target": self.target_field,
            "target_type": self.target_type,
            "rule": self.rule_type,
            "logic": self.logic_desc,
            "is_valid": self.is_valid,
            "warnings": self.warnings,
        }


class DBMLEngine:
    """Parses and converts Database Markup Language (DBML) to/from SQL and Mermaid."""

    @classmethod
    def parse_dbml(cls, dbml_code: str) -> List[DBMLTable]:
        tables: List[DBMLTable] = []
        table_pattern = re.compile(r"Table\s+([A-Za-z0-9_]+)\s*\{([^}]+)\}", re.IGNORECASE)
        field_line_pattern = re.compile(r"^\s*([A-Za-z0-9_]+)\s+([A-Za-z0-9_(),\s]+)(?:\[(.*?)\])?", re.MULTILINE)

        for match in table_pattern.finditer(dbml_code):
            tbl_name = match.group(1).strip()
            body = match.group(2)
            fields: List[DBMLField] = []

            for line in body.strip().splitlines():
                line = line.strip()
                if not line or line.startswith("//"):
                    continue

                f_match = field_line_pattern.match(line)
                if f_match:
                    f_name = f_match.group(1)
                    f_type = f_match.group(2).strip()
                    attrs_str = f_match.group(3) or ""

                    is_pk = "pk" in attrs_str.lower()
                    is_unique = "unique" in attrs_str.lower()
                    not_null = "not null" in attrs_str.lower()

                    # Extract default
                    default_m = re.search(r"default:\s*([^,\]]+)", attrs_str, re.IGNORECASE)
                    default_val = default_m.group(1).strip() if default_m else None

                    # Extract note
                    note_m = re.search(r"note:\s*'([^']+)'", attrs_str, re.IGNORECASE)
                    note = note_m.group(1).strip() if note_m else ""

                    # Extract ref
                    ref_m = re.search(r"ref:\s*>\s*([A-Za-z0-9_.]+)", attrs_str, re.IGNORECASE)
                    ref_target = ref_m.group(1).strip() if ref_m else None

                    fields.append(
                        DBMLField(
                            name=f_name,
                            data_type=f_type,
                            is_pk=is_pk,
                            is_unique=is_unique,
                            not_null=not_null,
                            default_val=default_val,
                            note=note,
                            ref_target=ref_target,
                        )
                    )

            tables.append(DBMLTable(name=tbl_name, fields=fields))
        return tables

    @classmethod
    def dbml_to_sql(cls, dbml_code: str, dialect: str = "postgres") -> str:
        tables = cls.parse_dbml(dbml_code)
        if not tables:
            return "-- No tables parsed from DBML."

        sql_statements = []
        for tbl in tables:
            cols_def = []
            for f in tbl.fields:
                col_sql = f"    {f.name} {f.data_type.upper()}"
                if f.is_pk:
                    col_sql += " PRIMARY KEY"
                if f.not_null:
                    col_sql += " NOT NULL"
                if f.is_unique:
                    col_sql += " UNIQUE"
                if f.default_val:
                    col_sql += f" DEFAULT {f.default_val}"
                if f.ref_target:
                    # e.g. ref: > customers.id
                    target_tbl, target_col = f.ref_target.split(".") if "." in f.ref_target else (f.ref_target, "id")
                    col_sql += f" REFERENCES {target_tbl}({target_col})"
                cols_def.append(col_sql)

            stmt = f"CREATE TABLE {tbl.name} (\n" + ",\n".join(cols_def) + "\n);"
            sql_statements.append(stmt)

        return "\n\n".join(sql_statements)

    @classmethod
    def sql_to_dbml(cls, sql_ddl: str) -> str:
        # Regex for CREATE TABLE
        tbl_regex = re.compile(r"CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?([A-Za-z0-9_\"`]+)\s*\((.*?)\);", re.DOTALL | re.IGNORECASE)
        col_regex = re.compile(r"^\s*([A-Za-z0-9_\"`]+)\s+([A-Za-z0-9_(),]+)(.*)$", re.IGNORECASE)

        dbml_out = []
        for match in tbl_regex.finditer(sql_ddl):
            raw_tbl = match.group(1).replace('"', '').replace('`', '').strip()
            body = match.group(2)
            lines = [l.strip() for l in body.split(",") if l.strip()]

            table_lines = [f"Table {raw_tbl} {{"]
            for col_line in lines:
                if col_line.upper().startswith(("PRIMARY KEY", "FOREIGN KEY", "CONSTRAINT", "INDEX")):
                    continue
                cm = col_regex.match(col_line)
                if cm:
                    c_name = cm.group(1).replace('"', '').replace('`', '').strip()
                    c_type = cm.group(2).strip().lower()
                    rest = cm.group(3).upper()

                    attrs = []
                    if "PRIMARY KEY" in rest:
                        attrs.append("pk")
                    if "NOT NULL" in rest:
                        attrs.append("not null")
                    if "UNIQUE" in rest:
                        attrs.append("unique")

                    # Check references
                    ref_match = re.search(r"REFERENCES\s+([A-Za-z0-9_]+)\s*\(([A-Za-z0-9_]+)\)", rest, re.IGNORECASE)
                    if ref_match:
                        ref_t = ref_match.group(1)
                        ref_c = ref_match.group(2)
                        attrs.append(f"ref: > {ref_t}.{ref_c}")

                    attr_str = f" [{', '.join(attrs)}]" if attrs else ""
                    table_lines.append(f"  {c_name} {c_type}{attr_str}")

            table_lines.append("}")
            dbml_out.append("\n".join(table_lines))

        return "\n\n".join(dbml_out) if dbml_out else "// Không nhận dạng được bảng SQL hợp lệ."

    @classmethod
    def data_dict_to_dbml(cls, table_name: str, fields: List[Any]) -> str:
        lines = [f"Table {table_name} {{"]
        for raw_f in fields:
            f = raw_f if isinstance(raw_f, dict) else raw_f.to_dict()
            col_name = f.get("field", "col")
            col_type = f.get("type", "varchar").lower()
            attrs = []
            if f.get("pk") == "YES" or f.get("is_primary_key"):
                attrs.append("pk")
            if f.get("nullable") == "NO":
                attrs.append("not null")
            fk = f.get("fk", "-")
            if fk and fk != "-":
                attrs.append(f"ref: > {fk}")
            desc = f.get("description", "")
            if desc:
                attrs.append(f"note: '{desc[:50]}'")

            attr_str = f" [{', '.join(attrs)}]" if attrs else ""
            lines.append(f"  {col_name} {col_type}{attr_str}")
        lines.append("}")
        return "\n".join(lines)

    @classmethod
    def dbml_to_mermaid_erd(cls, dbml_code: str) -> str:
        tables = cls.parse_dbml(dbml_code)
        lines = ["erDiagram"]
        refs = []

        for tbl in tables:
            lines.append(f"  {tbl.name} {{")
            for f in tbl.fields:
                clean_type = re.sub(r"[(),0-9\s]", "_", f.data_type).strip("_") or "string"
                pk_marker = " PK" if f.is_pk else ""
                fk_marker = " FK" if f.ref_target else ""
                comment = f' "{f.note}"' if f.note else ""
                lines.append(f"    {clean_type} {f.name}{pk_marker}{fk_marker}{comment}")
                if f.ref_target and "." in f.ref_target:
                    target_tbl = f.ref_target.split(".")[0]
                    refs.append(f"  {tbl.name} }}|--|| {target_tbl} : references")
            lines.append("  }")

        for r in refs:
            lines.append(r)

        return "\n".join(lines)


class SatsumaMappingEngine:
    """Parses Satsuma-style Data Mapping DSL and performs type & truncation audits."""

    @classmethod
    def parse_satsuma_dsl(cls, dsl_text: str) -> List[SatsumaMappingRule]:
        rules: List[SatsumaMappingRule] = []
        # Pattern: Source.Col -> Target.Col [Rule_Type: Note]
        pattern = re.compile(
            r"^\s*([A-Za-z0-9_.]+(?:\s*\+\s*[A-Za-z0-9_.\"\s]+)*)\s*->\s*([A-Za-z0-9_.]+)\s*(?:\[([^:\]]+)(?::\s*([^\]]+))?\])?",
            re.MULTILINE,
        )

        for line in dsl_text.splitlines():
            line = line.strip()
            if not line or line.startswith("#") or line.startswith("//"):
                continue

            m = pattern.match(line)
            if m:
                src_expr = m.group(1).strip()
                target_field = m.group(2).strip()
                rule_type = (m.group(3) or "Direct").strip()
                logic = (m.group(4) or "").strip()

                # Infer default types
                src_type = "VARCHAR(100)"
                target_type = "VARCHAR(100)"

                rule = SatsumaMappingRule(
                    source_expr=src_expr,
                    source_type=src_type,
                    target_field=target_field,
                    target_type=target_type,
                    rule_type=rule_type,
                    logic_desc=logic,
                )
                cls._audit_type_compatibility(rule)
                rules.append(rule)

        return rules

    @classmethod
    def _audit_type_compatibility(cls, rule: SatsumaMappingRule) -> None:
        warnings = []
        src_t = rule.source_type.upper()
        tgt_t = rule.target_type.upper()

        # Check String Truncation: e.g. VARCHAR(100) -> VARCHAR(30)
        src_len_m = re.search(r"VARCHAR\((\d+)\)", src_t)
        tgt_len_m = re.search(r"VARCHAR\((\d+)\)", tgt_t)

        if src_len_m and tgt_len_m:
            src_len = int(src_len_m.group(1))
            tgt_len = int(tgt_len_m.group(1))
            if src_len > tgt_len:
                warnings.append(
                    f"[TRUNCATION RISK] Độ dài nguồn ({src_len}) lớn hơn đích ({tgt_len}). Nguy cơ bị cắt ngắn dữ liệu."
                )

        # Check String to Number without Cast
        if "VARCHAR" in src_t and any(num in tgt_t for num in ["INT", "DECIMAL", "NUMERIC", "FLOAT"]):
            if "CAST" not in rule.rule_type.upper():
                warnings.append(
                    f"[CAST RISK] Chuyển từ chuỗi ({src_t}) sang số ({tgt_t}) cần có validation kiểm tra ký tự số."
                )

        # Check Timestamp to Date
        if ("TIMESTAMP" in src_t or "DATETIME" in src_t) and tgt_t == "DATE":
            warnings.append("[PRECISION LOSS] Chuyển Timestamp sang Date sẽ làm mất thông tin giờ:phút:giây.")

        rule.warnings = warnings
        rule.is_valid = len([w for w in warnings if "TRUNCATION" in w]) == 0

    @classmethod
    def generate_data_lineage_mermaid(cls, rules: List[SatsumaMappingRule]) -> str:
        lines = ["flowchart LR", "  classDef src fill:#0F172A,stroke:#38BDF8,stroke-width:2px,color:#F8FAFC;", "  classDef tgt fill:#1E293B,stroke:#4ADE80,stroke-width:2px,color:#F8FAFC;", "  classDef rule fill:#1E293B,stroke:#FBBF24,stroke-width:1px,color:#FBBF24,stroke-dasharray: 4 4;"]

        for idx, r in enumerate(rules, start=1):
            src_node = f"S_{idx}[\"{r.source_expr}\"]"
            tgt_node = f"T_{idx}[\"{r.target_field}\"]"
            rule_node = f"R_{idx}{{{{\"[{r.rule_type}] {r.logic_desc[:20]}\"}}}}"

            lines.append(f"  {src_node} --> {rule_node}")
            lines.append(f"  {rule_node} --> {tgt_node}")
            lines.append(f"  class S_{idx} src;")
            lines.append(f"  class T_{idx} tgt;")
            lines.append(f"  class R_{idx} rule;")

        return "\n".join(lines)
