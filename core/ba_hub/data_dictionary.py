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
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


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

