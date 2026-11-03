"""
core.ba_hub.data_dictionary — Enterprise Data Dictionary & Data Mapping Engine for BAs.

Features:
1. Automated Excel Schema Scanner: Infers types (VARCHAR, INT, DECIMAL, DATE, BOOL), nullability, and PK candidates.
2. SQL DDL Parser: Parses CREATE TABLE definitions into structured data dictionaries.
3. Source-to-Target Data Mapping Matrix: Helps BAs specify integration transformations (e.g. SAP to WMS).
4. Professional OpenXML Excel (.xlsx) & Word (.docx) export with frozen headers and auto-adjusted widths.
"""

from __future__ import annotations

import datetime
import json
import os
from pathlib import Path
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import yaml
except ImportError:
    yaml = None

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
except ImportError:
    Document = None


@dataclass
class DataField:
    field_name: str
    display_name: str = ""
    data_type: str = "VARCHAR"
    nullable: str = "YES"
    is_primary_key: bool = False
    is_foreign_key: bool = False
    references: str = "-"
    sample_values: List[str] = field(default_factory=list)
    description: str = ""
    business_rules: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "field": self.field_name,
            "name": self.display_name or self.field_name,
            "type": self.data_type,
            "nullable": self.nullable,
            "pk": "YES" if self.is_primary_key else "NO",
            "fk": self.references if self.is_foreign_key else "-",
            "description": self.description,
            "rules": self.business_rules,
            "samples": ", ".join(self.sample_values[:3])
        }


class DataDictionaryScanner:
    """Scans and extracts schema metadata from Excel sheets or SQL DDL statements."""

    @staticmethod
    def _infer_type(values: List[Any]) -> str:
        non_null = [v for v in values if v is not None and str(v).strip() != ""]
        if not non_null:
            return "VARCHAR(255)"

        is_int = True
        is_float = True
        is_date = True
        is_bool = True
        max_len = 0

        for val in non_null:
            s_val = str(val).strip()
            max_len = max(max_len, len(s_val))

            if isinstance(val, bool):
                is_int = is_float = is_date = False
                continue
            else:
                is_bool = False

            if isinstance(val, (datetime.date, datetime.datetime)):
                is_int = is_float = False
                continue
            else:
                # check string date
                if not re.match(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}', s_val) and not re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{4}', s_val):
                    is_date = False

            if isinstance(val, int) and not isinstance(val, bool):
                pass
            else:
                try:
                    int(s_val)
                except ValueError:
                    is_int = False

            if isinstance(val, (int, float)):
                pass
            else:
                try:
                    float(s_val.replace(",", "."))
                except ValueError:
                    is_float = False

        if is_bool:
            return "BOOLEAN"
        if is_int:
            return "BIGINT" if max_len > 9 else "INTEGER"
        if is_float:
            return "DECIMAL(18,2)"
        if is_date:
            return "DATE"
        return f"VARCHAR({max(32, min(500, ((max_len // 10) + 1) * 20))})"

    @classmethod
    def scan_excel(cls, file_path: str, sheet_name: Optional[str] = None) -> Tuple[str, List[DataField]]:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        table_name = ws.title

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return table_name, []

        # Find header row (first row with multiple string values)
        header_idx = 0
        for idx, row in enumerate(rows[:5]):
            str_cells = [c for c in row if c is not None and str(c).strip()]
            if len(str_cells) >= 2:
                header_idx = idx
                break

        header_row = rows[header_idx]
        data_rows = rows[header_idx + 1: header_idx + 201]  # Sample up to 200 rows

        fields: List[DataField] = []
        for col_idx, col_name in enumerate(header_row):
            if col_name is None or str(col_name).strip() == "":
                continue

            raw_col_str = str(col_name).strip()
            col_values = [r[col_idx] for r in data_rows if col_idx < len(r)]

            inferred_type = cls._infer_type(col_values)
            non_null_count = sum(1 for v in col_values if v is not None and str(v).strip() != "")
            has_null = non_null_count < len(col_values)
            unique_count = len(set(col_values))

            # PK heuristic: no nulls, high row count, unique values, name has 'id' or 'code' or 'no'
            is_pk = (not has_null and unique_count == len(col_values) and len(col_values) >= 3 and
                     any(k in raw_col_str.lower() for k in ["id", "code", "mã", "no", "cif"]))

            sample_strings = [str(v) for v in col_values if v is not None and str(v).strip() != ""][:3]

            field_obj = DataField(
                field_name=re.sub(r'[\s\-]+', '_', raw_col_str.lower()),
                display_name=raw_col_str,
                data_type=inferred_type,
                nullable="YES" if has_null else "NO",
                is_primary_key=is_pk,
                sample_values=sample_strings,
                description=f"Dữ liệu trích xuất từ cột '{raw_col_str}'.",
                business_rules="Giá trị duy nhất bắt buộc" if is_pk else ""
            )
            fields.append(field_obj)

        return table_name, fields

    @classmethod
    def scan_sql_ddl(cls, ddl_text: str) -> List[Dict[str, Any]]:
        """Parses standard SQL DDL CREATE TABLE statements into table schemas."""
        tables: List[Dict[str, Any]] = []

        # Find CREATE TABLE blocks
        pattern = re.compile(r'CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?([`"\[]?\w+[`"\]]?)\s*\((.*?)\);', re.DOTALL | re.IGNORECASE)
        for match in pattern.finditer(ddl_text):
            table_name = re.sub(r'[`"\[\]]', '', match.group(1))
            body = match.group(2)

            columns: List[DataField] = []
            lines = [l.strip() for l in body.split('\n') if l.strip()]

            # Also catch comma-separated if on single lines
            raw_clauses: List[str] = []
            curr_clause = []
            paren_depth = 0
            for char in body:
                if char == '(': paren_depth += 1
                elif char == ')': paren_depth -= 1
                elif char == ',' and paren_depth == 0:
                    raw_clauses.append("".join(curr_clause).strip())
                    curr_clause = []
                    continue
                curr_clause.append(char)
            if curr_clause:
                raw_clauses.append("".join(curr_clause).strip())

            table_pks: List[str] = []
            for clause in raw_clauses:
                # Table level primary key
                pk_match = re.match(r'PRIMARY\s+KEY\s*\((.*?)\)', clause, re.IGNORECASE)
                if pk_match:
                    pks = [re.sub(r'[`"\[\]\s]', '', p) for p in pk_match.group(1).split(',')]
                    table_pks.extend(pks)
                    continue

                if re.match(r'CONSTRAINT|FOREIGN\s+KEY|KEY|INDEX|UNIQUE', clause, re.IGNORECASE):
                    continue

                parts = clause.split()
                if len(parts) >= 2:
                    col_name = re.sub(r'[`"\[\]]', '', parts[0])
                    col_type = parts[1].upper()
                    # Collect parameters like VARCHAR(50) or DECIMAL(10,2)
                    if len(parts) > 2 and parts[2].startswith('('):
                        col_type += parts[2].upper()

                    is_pk = "PRIMARY KEY" in clause.upper() or col_name in table_pks
                    is_not_null = "NOT NULL" in clause.upper()
                    fk_ref = "-"
                    ref_match = re.search(r'REFERENCES\s+([`"\[]?\w+[`"\]]?)\s*\((.*?)\)', clause, re.IGNORECASE)
                    if ref_match:
                        fk_ref = f"{ref_match.group(1)}({ref_match.group(2)})"

                    df = DataField(
                        field_name=col_name,
                        display_name=col_name.replace("_", " ").title(),
                        data_type=col_type,
                        nullable="NO" if (is_not_null or is_pk) else "YES",
                        is_primary_key=is_pk,
                        is_foreign_key=bool(ref_match),
                        references=fk_ref,
                        description=f"Trường dữ liệu {col_name} trong bảng {table_name}."
                    )
                    columns.append(df)

            tables.append({
                "table_name": table_name,
                "fields": [c.to_dict() for c in columns]
            })

        return tables

    @classmethod
    def scan_openapi(cls, spec_input: Union[str, dict, Path]) -> Dict[str, Any]:
        """
        Parses an OpenAPI 3.x or Swagger 2.0 specification (dict, JSON/YAML string, or file path).
        Extracts endpoints, schemas, parameters, and translates them into DataField dictionaries.
        """
        raw_dict: Dict[str, Any] = {}
        if isinstance(spec_input, dict):
            raw_dict = spec_input
        elif isinstance(spec_input, (str, Path)):
            spec_str = str(spec_input).strip()
            if os.path.isfile(spec_str):
                with open(spec_str, "r", encoding="utf-8") as f:
                    content = f.read()
            else:
                content = spec_str

            try:
                raw_dict = json.loads(content)
            except Exception:
                if yaml is not None:
                    try:
                        raw_dict = yaml.safe_load(content) or {}
                    except Exception as ye:
                        raise ValueError(f"Không thể phân tích định dạng OpenAPI (JSON/YAML): {ye}")
                else:
                    raise ValueError("Tệp không phải JSON hợp lệ và thư viện PyYAML chưa được cài đặt.")

        info = raw_dict.get("info", {})
        title = info.get("title", "API Contract Specification")
        version = str(info.get("version", "1.0.0"))
        description = info.get("description", "")
        spec_version = raw_dict.get("openapi") or raw_dict.get("swagger") or "3.0.0"

        components = raw_dict.get("components", {})
        schemas_raw = components.get("schemas", {}) if "components" in raw_dict else raw_dict.get("definitions", {})

        tables: List[Dict[str, Any]] = []
        all_fields: List[DataField] = []

        def _map_openapi_type(prop: Dict[str, Any]) -> str:
            t = prop.get("type", "string")
            fmt = prop.get("format", "")
            if t == "integer":
                return "BIGINT" if fmt == "int64" else "INT"
            elif t == "number":
                return "FLOAT" if fmt == "float" else "DECIMAL(18,4)"
            elif t == "boolean":
                return "BOOLEAN"
            elif t == "string":
                if fmt in ("date", "date-time"):
                    return "DATETIME"
                elif fmt == "binary":
                    return "BLOB/BYTEA"
                elif fmt == "uuid":
                    return "UUID"
                max_l = prop.get("maxLength")
                return f"VARCHAR({max_l})" if max_l else "VARCHAR(255)"
            elif t == "array":
                items = prop.get("items", {})
                inner_type = items.get("type", "object") if isinstance(items, dict) else "OBJECT"
                return f"ARRAY<{str(inner_type).upper()}>"
            elif t == "object":
                return "JSON/OBJECT"
            return str(t).upper()

        for schema_name, schema_def in schemas_raw.items():
            if not isinstance(schema_def, dict):
                continue
            properties = schema_def.get("properties", {})
            required_props = set(schema_def.get("required", []))

            schema_fields: List[DataField] = []
            for prop_name, prop_def in properties.items():
                if not isinstance(prop_def, dict):
                    continue
                d_type = _map_openapi_type(prop_def)
                is_req = prop_name in required_props
                nullable_val = "NO" if is_req else "YES"
                if "nullable" in prop_def:
                    nullable_val = "YES" if prop_def["nullable"] else "NO"

                rules_list = []
                if "minimum" in prop_def: rules_list.append(f"min: {prop_def['minimum']}")
                if "maximum" in prop_def: rules_list.append(f"max: {prop_def['maximum']}")
                if "minLength" in prop_def: rules_list.append(f"minLen: {prop_def['minLength']}")
                if "maxLength" in prop_def: rules_list.append(f"maxLen: {prop_def['maxLength']}")
                if "pattern" in prop_def: rules_list.append(f"regex: {prop_def['pattern']}")
                if "enum" in prop_def: rules_list.append(f"enum: {prop_def['enum']}")

                samples = []
                if "example" in prop_def:
                    samples.append(str(prop_def["example"]))
                elif "default" in prop_def:
                    samples.append(str(prop_def["default"]))

                is_pk = prop_name.lower() in ("id", f"{schema_name.lower()}_id", f"{schema_name.lower()}id")

                df = DataField(
                    field_name=prop_name,
                    display_name=prop_def.get("title") or prop_name.replace("_", " ").title(),
                    data_type=d_type,
                    nullable=nullable_val,
                    is_primary_key=is_pk,
                    description=prop_def.get("description") or f"Thuộc tính {prop_name} trong đối tượng {schema_name}.",
                    business_rules="; ".join(rules_list),
                    sample_values=samples
                )
                schema_fields.append(df)
                all_fields.append(df)

            tables.append({
                "table_name": schema_name,
                "fields": [f.to_dict() for f in schema_fields],
                "raw_fields": schema_fields
            })

        endpoints: List[Dict[str, Any]] = []
        paths = raw_dict.get("paths", {})
        for path_str, path_item in paths.items():
            if not isinstance(path_item, dict):
                continue
            for method in ("get", "post", "put", "patch", "delete", "options", "head"):
                op = path_item.get(method)
                if not isinstance(op, dict):
                    continue
                op_id = op.get("operationId") or f"{method.upper()} {path_str}"
                summary = op.get("summary") or op.get("description") or ""
                tags = op.get("tags", [])

                param_fields = []
                for p in op.get("parameters", []):
                    if not isinstance(p, dict):
                        continue
                    p_name = p.get("name", "")
                    p_in = p.get("in", "query")
                    p_req = p.get("required", False)
                    p_schema = p.get("schema", {})
                    p_type = _map_openapi_type(p_schema) if isinstance(p_schema, dict) and p_schema else "STRING"
                    param_fields.append({
                        "name": p_name,
                        "in": p_in,
                        "required": p_req,
                        "type": p_type,
                        "description": p.get("description", "")
                    })

                resps = []
                for status_code, resp_def in op.get("responses", {}).items():
                    resps.append({
                        "status": str(status_code),
                        "description": resp_def.get("description", "") if isinstance(resp_def, dict) else str(resp_def)
                    })

                endpoints.append({
                    "method": method.upper(),
                    "path": path_str,
                    "operation_id": op_id,
                    "summary": summary,
                    "tags": tags,
                    "parameters": param_fields,
                    "responses": resps
                })

        return {
            "title": title,
            "version": version,
            "spec_version": spec_version,
            "description": description,
            "endpoints_count": len(endpoints),
            "schemas_count": len(tables),
            "tables": tables,
            "endpoints": endpoints,
            "all_fields_count": len(all_fields)
        }


class FieldMappingEngine:
    """Generates Source-to-Target Data Mapping Matrices for System Integrations."""

    @staticmethod
    def export_data_dictionary_to_excel(table_name: str, fields: List[DataField], output_path: str) -> str:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = re.sub(r'[\\/*?:\[\]]', '_', table_name)[:30]
        ws.views.sheetView[0].showGridLines = True

        # Header Title
        ws.merge_cells("A1:H1")
        title_cell = ws["A1"]
        title_cell.value = f"BẢNG TỪ ĐIỂN DỮ LIỆU (DATA DICTIONARY) — BẢNG: {table_name.upper()}"
        title_cell.font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Subtitle
        ws.merge_cells("A2:H2")
        sub_cell = ws["A2"]
        sub_cell.value = f"Tài liệu phân tích cấu trúc dữ liệu và quy tắc nghiệp vụ | Xuất bản ngày: {datetime.date.today().strftime('%d/%m/%Y')}"
        sub_cell.font = Font(name="Segoe UI", size=9.5, italic=True, color="64748B")
        sub_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Column Headers
        headers = ["STT", "Tên Trường (Field Name)", "Tên Hiển Thị (Label)", "Kiểu Dữ Liệu", "Bắt Buộc (Null?)", "Khóa Chính (PK)", "Khóa Ngoại (FK)", "Mô Tả & Quy Tắc Nghiệp Vụ"]
        ws.row_dimensions[4].height = 28

        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        header_border = Border(
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="medium", color="0284C7"),
            left=Side(style="thin", color="CBD5E1"),
            right=Side(style="thin", color="CBD5E1"),
        )

        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = header_border

        # Data Rows
        data_font = Font(name="Segoe UI", size=9.5)
        code_font = Font(name="Consolas", size=9.5, color="0369A1")
        cell_border = Border(
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0"),
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
        )
        zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        start_row = 5
        for idx, fld in enumerate(fields, start=1):
            curr_row = start_row + idx - 1
            ws.row_dimensions[curr_row].height = 24

            ws.cell(row=curr_row, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
            c_fld = ws.cell(row=curr_row, column=2, value=fld.field_name)
            c_fld.font = code_font
            c_fld.alignment = Alignment(horizontal="left", vertical="center")

            ws.cell(row=curr_row, column=3, value=fld.display_name or fld.field_name).alignment = Alignment(horizontal="left", vertical="center")

            c_type = ws.cell(row=curr_row, column=4, value=fld.data_type)
            c_type.font = Font(name="Consolas", size=9, bold=True, color="475569")
            c_type.alignment = Alignment(horizontal="center", vertical="center")

            c_null = ws.cell(row=curr_row, column=5, value="Bắt buộc" if fld.nullable == "NO" else "Tùy chọn")
            c_null.alignment = Alignment(horizontal="center", vertical="center")
            if fld.nullable == "NO":
                c_null.font = Font(name="Segoe UI", size=9.5, bold=True, color="BE123C")

            c_pk = ws.cell(row=curr_row, column=6, value="PK" if fld.is_primary_key else "-")
            c_pk.alignment = Alignment(horizontal="center", vertical="center")
            if fld.is_primary_key:
                c_pk.font = Font(name="Segoe UI", size=10, bold=True, color="D97706")

            ws.cell(row=curr_row, column=7, value=fld.references if fld.is_foreign_key else "-").alignment = Alignment(horizontal="center", vertical="center")

            desc_text = fld.description
            if fld.business_rules:
                desc_text += f" [Ràng buộc: {fld.business_rules}]"
            ws.cell(row=curr_row, column=8, value=desc_text).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            for col in range(1, 9):
                c = ws.cell(row=curr_row, column=col)
                c.border = cell_border
                if not c.font or c.font.name != "Consolas":
                    c.font = data_font
                if idx % 2 == 0:
                    c.fill = zebra_fill

        # Auto-fit column widths
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 24
        ws.column_dimensions["C"].width = 24
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 12
        ws.column_dimensions["G"].width = 16
        ws.column_dimensions["H"].width = 45

        # Freeze panes below header
        ws.freeze_panes = "A5"

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        return output_path

    @staticmethod
    def export_data_mapping_to_excel(source_system: str, target_system: str, mapping_rows: List[Dict[str, Any]], output_path: str) -> str:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Source_To_Target_Mapping"
        ws.views.sheetView[0].showGridLines = True

        # Header Title
        ws.merge_cells("A1:G1")
        title_cell = ws["A1"]
        title_cell.value = f"MA TRẬN ÁNH XẠ TRƯỜNG DỮ LIỆU (DATA MAPPING): {source_system.upper()} ➔ {target_system.upper()}"
        title_cell.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        headers = ["STT", f"Nguồn: {source_system} (Source)", "Kiểu Nguồn", "Quy Tắc Chuyển Đổi (Transformation Rule)", f"Đích: {target_system} (Target)", "Kiểu Đích", "Ghi Chú Nghiệp Vụ"]
        ws.row_dimensions[3].height = 28

        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for idx, row_data in enumerate(mapping_rows, start=1):
            r = 3 + idx
            ws.row_dimensions[r].height = 24
            ws.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r, column=2, value=row_data.get("src_field", "")).font = Font(name="Consolas", size=9.5, color="0369A1")
            ws.cell(row=r, column=3, value=row_data.get("src_type", "VARCHAR")).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r, column=4, value=row_data.get("transform", "Direct Copy (1:1)")).alignment = Alignment(horizontal="left", vertical="center")
            ws.cell(row=r, column=5, value=row_data.get("target_field", "")).font = Font(name="Consolas", size=9.5, bold=True, color="0D9488")
            ws.cell(row=r, column=6, value=row_data.get("target_type", "VARCHAR")).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r, column=7, value=row_data.get("notes", "")).alignment = Alignment(horizontal="left", vertical="center")

            for col in range(1, 8):
                ws.cell(row=r, column=col).border = Border(
                    top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"),
                    left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0")
                )

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 32
        ws.column_dimensions["E"].width = 25
        ws.column_dimensions["F"].width = 14
        ws.column_dimensions["G"].width = 30

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        return output_path

    @staticmethod
    def export_openapi_to_excel(api_data: Dict[str, Any], output_path: str) -> str:
        wb = openpyxl.Workbook()
        
        # Sheet 1: Endpoints
        ws1 = wb.active
        ws1.title = "API_Endpoints"
        ws1.views.sheetView[0].showGridLines = True

        title = api_data.get("title", "API Contract")
        version = api_data.get("version", "1.0.0")

        ws1.merge_cells("A1:E1")
        c1 = ws1["A1"]
        c1.value = f"DANH MỤC API ENDPOINTS — {title.upper()} (v{version})"
        c1.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
        c1.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        c1.alignment = Alignment(horizontal="center", vertical="center")
        ws1.row_dimensions[1].height = 36

        headers1 = ["STT", "Phương Thức", "Đường Dẫn Endpoint", "Tóm Tắt Chức Năng", "Tham Số / Yêu Cầu"]
        ws1.row_dimensions[3].height = 26
        header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        for col_idx, h in enumerate(headers1, start=1):
            cell = ws1.cell(row=3, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        method_colors = {
            "GET": ("0284C7", "F0F9FF"),
            "POST": ("16A34A", "F0FDF4"),
            "PUT": ("D97706", "FFFBEB"),
            "PATCH": ("CA8A04", "FEFCE8"),
            "DELETE": ("DC2626", "FEF2F2"),
        }

        endpoints = api_data.get("endpoints", [])
        for idx, ep in enumerate(endpoints, start=1):
            r = 3 + idx
            ws1.row_dimensions[r].height = 24
            ws1.cell(row=r, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")

            m = ep.get("method", "GET")
            m_cell = ws1.cell(row=r, column=2, value=m)
            fg_color, bg_color = method_colors.get(m, ("475569", "F8FAFC"))
            m_cell.font = Font(name="Segoe UI", size=9.5, bold=True, color=fg_color)
            m_cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            m_cell.alignment = Alignment(horizontal="center", vertical="center")

            p_cell = ws1.cell(row=r, column=3, value=ep.get("path", ""))
            p_cell.font = Font(name="Consolas", size=9.5, bold=True, color="0F172A")
            p_cell.alignment = Alignment(horizontal="left", vertical="center")

            ws1.cell(row=r, column=4, value=ep.get("summary", "")).alignment = Alignment(horizontal="left", vertical="center")

            # Format parameters
            params = ep.get("parameters", [])
            p_desc = ", ".join([f"{p.get('name')}:{p.get('type')}" for p in params[:4]]) if params else "-"
            ws1.cell(row=r, column=5, value=p_desc).alignment = Alignment(horizontal="left", vertical="center")

            for c in range(1, 6):
                ws1.cell(row=r, column=c).border = Border(
                    top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"),
                    left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0")
                )

        ws1.column_dimensions["A"].width = 6
        ws1.column_dimensions["B"].width = 14
        ws1.column_dimensions["C"].width = 38
        ws1.column_dimensions["D"].width = 36
        ws1.column_dimensions["E"].width = 30

        # Sheet 2: Data Dictionary (All Schemas)
        ws2 = wb.create_sheet(title="Data_Schemas")
        ws2.views.sheetView[0].showGridLines = True

        ws2.merge_cells("A1:H1")
        c2 = ws2["A1"]
        c2.value = f"TỪ ĐIỂN DỮ LIỆU ĐỐI TƯỢNG (DATA SCHEMAS DICTIONARY) — {title.upper()}"
        c2.font = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
        c2.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        ws2.row_dimensions[1].height = 36

        headers2 = ["STT", "Đối Tượng / Schema", "Tên Thuộc Tính", "Kiểu Dữ Liệu", "Bắt Buộc?", "Ràng Buộc / Rules", "Mô Tả", "Ví Dụ (Sample)"]
        ws2.row_dimensions[3].height = 26
        for col_idx, h in enumerate(headers2, start=1):
            cell = ws2.cell(row=3, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        curr_row = 4
        item_stt = 1
        for tbl in api_data.get("tables", []):
            s_name = tbl.get("table_name", "")
            for fld in tbl.get("fields", []):
                ws2.row_dimensions[curr_row].height = 22
                ws2.cell(row=curr_row, column=1, value=item_stt).alignment = Alignment(horizontal="center", vertical="center")
                ws2.cell(row=curr_row, column=2, value=s_name).font = Font(name="Segoe UI", size=9.5, bold=True, color="1E3A8A")
                ws2.cell(row=curr_row, column=3, value=fld.get("field", "")).font = Font(name="Consolas", size=9.5, color="0284C7")
                ws2.cell(row=curr_row, column=4, value=fld.get("type", "VARCHAR")).alignment = Alignment(horizontal="center", vertical="center")

                null_cell = ws2.cell(row=curr_row, column=5, value=fld.get("nullable", "YES"))
                null_cell.alignment = Alignment(horizontal="center", vertical="center")
                if fld.get("nullable") == "NO":
                    null_cell.font = Font(name="Segoe UI", size=9.5, bold=True, color="DC2626")

                ws2.cell(row=curr_row, column=6, value=fld.get("rules", "-")).alignment = Alignment(horizontal="left", vertical="center")
                ws2.cell(row=curr_row, column=7, value=fld.get("description", "")).alignment = Alignment(horizontal="left", vertical="center")
                ws2.cell(row=curr_row, column=8, value=fld.get("samples", "-")).alignment = Alignment(horizontal="left", vertical="center")

                for c in range(1, 9):
                    ws2.cell(row=curr_row, column=c).border = Border(
                        top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"),
                        left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0")
                    )

                curr_row += 1
                item_stt += 1

        ws2.column_dimensions["A"].width = 6
        ws2.column_dimensions["B"].width = 22
        ws2.column_dimensions["C"].width = 25
        ws2.column_dimensions["D"].width = 16
        ws2.column_dimensions["E"].width = 12
        ws2.column_dimensions["F"].width = 24
        ws2.column_dimensions["G"].width = 34
        ws2.column_dimensions["H"].width = 20

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        return output_path

    @staticmethod
    def export_openapi_to_docx(api_data: Dict[str, Any], output_path: str) -> str:
        if Document is None:
            raise RuntimeError("Thư viện python-docx chưa được cài đặt.")

        doc = Document()
        for sec in doc.sections:
            sec.top_margin = Inches(0.8)
            sec.bottom_margin = Inches(0.8)
            sec.left_margin = Inches(0.9)
            sec.right_margin = Inches(0.9)

        # Title
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title = api_data.get("title", "API Contract Specification")
        version = api_data.get("version", "1.0.0")
        r_title = p_title.add_run(f"ĐẶC TẢ GIAO DIỆN LẬP TRÌNH ỨNG DỤNG (API SPECIFICATION)\n{title.upper()}")
        r_title.bold = True
        r_title.font.size = Pt(16)
        r_title.font.color.rgb = RGBColor(15, 23, 42)

        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sub = p_sub.add_run(f"Phiên bản API: v{version} | OpenAPI: {api_data.get('spec_version', '3.0')} | Ngày xuất: {datetime.date.today().strftime('%d/%m/%Y')}")
        r_sub.italic = True
        r_sub.font.size = Pt(10)
        r_sub.font.color.rgb = RGBColor(100, 116, 139)

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # 1. Endpoints Table
        doc.add_heading("1. Danh Mục Các Điểm Cuối (API Endpoints)", level=1)
        endpoints = api_data.get("endpoints", [])
        if endpoints:
            tbl_ep = doc.add_table(rows=len(endpoints) + 1, cols=4)
            tbl_ep.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl_ep.autofit = False

            headers = ["STT", "Phương Thức", "Đường Dẫn Endpoint", "Mô Tả Chức Năng"]
            col_widths = [Inches(0.6), Inches(1.3), Inches(2.8), Inches(2.0)]

            # Header row
            hdr_cells = tbl_ep.rows[0].cells
            tbl_ep.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            tbl_ep.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

            for i, h_text in enumerate(headers):
                cell = hdr_cells[i]
                cell.width = col_widths[i]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(h_text)
                r.bold = True
                r.font.size = Pt(9.5)
                r.font.color.rgb = RGBColor(255, 255, 255)
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0F172A"/>'))
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

            for idx, ep in enumerate(endpoints, start=1):
                row = tbl_ep.rows[idx]
                row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                
                cells = row.cells
                for ci, w in enumerate(col_widths):
                    cells[ci].width = w
                    tcPr = cells[ci]._tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                    if idx % 2 == 0:
                        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))

                cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cells[0].paragraphs[0].add_run(str(idx)).font.size = Pt(9)

                m = ep.get("method", "GET")
                p_m = cells[1].paragraphs[0]
                p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_m = p_m.add_run(m)
                r_m.bold = True
                r_m.font.size = Pt(9)
                if m == "GET": r_m.font.color.rgb = RGBColor(2, 132, 199)
                elif m == "POST": r_m.font.color.rgb = RGBColor(22, 163, 74)
                elif m == "PUT": r_m.font.color.rgb = RGBColor(217, 119, 6)
                elif m == "DELETE": r_m.font.color.rgb = RGBColor(220, 38, 38)

                p_path = cells[2].paragraphs[0]
                r_path = p_path.add_run(ep.get("path", ""))
                r_path.font.name = "Consolas"
                r_path.font.size = Pt(8.5)

                p_desc = cells[3].paragraphs[0]
                r_desc = p_desc.add_run(ep.get("summary", "-"))
                r_desc.font.size = Pt(9)

        doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # 2. Schemas Table
        doc.add_heading("2. Từ Điển Dữ Liệu Đối Tượng (Data Schemas)", level=1)
        for tbl in api_data.get("tables", []):
            s_name = tbl.get("table_name", "Schema")
            fields = tbl.get("fields", [])

            h2 = doc.add_heading(f"Đối tượng (Schema): {s_name}", level=2)
            h2.paragraph_format.space_before = Pt(8)
            h2.paragraph_format.space_after = Pt(4)

            if not fields:
                doc.add_paragraph("Không có thuộc tính khai báo.")
                continue

            tbl_schema = doc.add_table(rows=len(fields) + 1, cols=6)
            tbl_schema.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl_schema.autofit = False

            s_headers = ["STT", "Thuộc Tính (Field)", "Kiểu", "Null?", "Rules / Ràng Buộc", "Mô Tả Nghiệp Vụ"]
            s_widths = [Inches(0.5), Inches(1.8), Inches(1.1), Inches(0.7), Inches(1.4), Inches(1.5)]

            # Header
            tbl_schema.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            tbl_schema.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            for ci, h_text in enumerate(s_headers):
                cell = tbl_schema.rows[0].cells[ci]
                cell.width = s_widths[ci]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(h_text)
                r.bold = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(255, 255, 255)
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1E293B"/>'))
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))

            for f_idx, fld in enumerate(fields, start=1):
                row = tbl_schema.rows[f_idx]
                row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                cells = row.cells
                for ci, w in enumerate(s_widths):
                    cells[ci].width = w
                    tcPr = cells[ci]._tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                    if f_idx % 2 == 0:
                        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))

                cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cells[0].paragraphs[0].add_run(str(f_idx)).font.size = Pt(8.5)

                cells[1].paragraphs[0].add_run(fld.get("field", "")).font.name = "Consolas"
                cells[1].paragraphs[0].runs[0].font.size = Pt(8.5)
                cells[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(2, 132, 199)

                cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cells[2].paragraphs[0].add_run(fld.get("type", "VARCHAR")).font.size = Pt(8.5)

                cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_null = cells[3].paragraphs[0].add_run(fld.get("nullable", "YES"))
                r_null.font.size = Pt(8.5)
                if fld.get("nullable") == "NO":
                    r_null.bold = True
                    r_null.font.color.rgb = RGBColor(220, 38, 38)

                cells[4].paragraphs[0].add_run(fld.get("rules", "-")).font.size = Pt(8.5)
                cells[5].paragraphs[0].add_run(fld.get("description", "")).font.size = Pt(8.5)

            doc.add_paragraph().paragraph_format.space_after = Pt(8)

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        doc.save(output_path)
        return output_path


def scan_openapi_spec(spec_input: Union[str, dict, Path]) -> Dict[str, Any]:
    """Standalone helper function to scan OpenAPI / Swagger specifications."""
    return DataDictionaryScanner.scan_openapi(spec_input)


