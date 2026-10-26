"""
server/mcp_server.py — Standard Model Context Protocol (MCP) STDIO Server (v4.0).

Enables AI Agents (Antigravity IDE, Claude Desktop, Cursor, Roo-Code) to invoke
the Antigravity Office Engine via standard JSON-RPC 2.0 over STDIO.

Registered Tools:
  1. doc_convert        — High-fidelity bidirectional document conversion.
  2. doc_inspect        — Parse DOCX / PPTX / XLSX into structured JSON AST.
  3. doc_insert_diagram — Surgically insert diagram into Word DOCX or PowerPoint PPTX.
  4. diagram_render     — Multi-engine diagram renderer (Mermaid, PlantUML, Draw.io, Canvas).
  5. pptx_synthesize    — Synthesize 16:9 presentation deck from DOCX, PDF, or Markdown.
  6. env_doctor         — Zero-Setup Environment Doctor diagnostic suite.
"""

from __future__ import annotations

import json
import os
import sys
import traceback
from typing import Any, Dict, List, Optional

# Ensure standard UTF-8 handling on Windows
if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Core imports with dual-fallback
try:
    from core.converter.engine import convert_universal
    from core.docx.reader import read_docx, read_docx_to_json
    from core.pptx.reader import read_pptx
    from core.xlsx.reader import read_xlsx
    from core.docx.writer import inject_diagram_into_docx
    from core.pptx.engine import insert_diagram_into_pptx, parse_markdown_to_slides_spec
    from core.pptx.writer import write_pptx_from_spec
    from core.pptx.docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx
    from core.diagrams.bridge import render_diagram_file_to_png, render_diagram_code_to_png
    from core.env_doctor import run_doctor
    from core.xlsx.writer import UniversalExcelEngine
    from core.docx.writer import patch_docx_template
    from core.docx.mail_merge import inspect_template_tokens
except (ImportError, ValueError):
    from converter_engine import convert_universal  # type: ignore
    from docx_reader import read_docx, read_docx_to_json  # type: ignore
    from pptx_reader import read_pptx  # type: ignore
    from xlsx_reader import read_xlsx  # type: ignore
    from docx_writer import inject_diagram_into_docx, patch_docx_template  # type: ignore
    from pptx_engine import insert_diagram_into_pptx, parse_markdown_to_slides_spec  # type: ignore
    from pptx_writer import write_pptx_from_spec  # type: ignore
    from docx_to_pptx import convert_docx_to_pptx, convert_pdf_to_pptx  # type: ignore
    from diagram_bridge import render_diagram_file_to_png, render_diagram_code_to_png  # type: ignore
    from env_doctor import run_doctor  # type: ignore
    from xlsx_writer import UniversalExcelEngine  # type: ignore
    from docx_mail_merge import inspect_template_tokens  # type: ignore


MCP_PROTOCOL_VERSION = "2024-11-05"
SERVER_NAME = "antigravity-doc-handler"
SERVER_VERSION = "4.5.0"

TOOL_DEFINITIONS = [
    {
        "name": "doc_convert",
        "description": "Universal document conversion between DOCX, PDF, Markdown, PPTX, and Excel (.xlsx) with OpenXML table repair and semantic table extraction.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "source_path": {
                    "type": "string",
                    "description": "Absolute path to source document (.docx, .pdf, .md, .pptx)"
                },
                "target_ext": {
                    "type": "string",
                    "description": "Target extension (.docx, .pdf, .md, .pptx, .xlsx, .html)",
                    "default": ".docx"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path. If omitted, writes next to source file."
                },
                "page_range": {
                    "type": "string",
                    "description": "Optional page range for PDF sources (e.g. '1-5', '3')"
                }
            },
            "required": ["source_path"]
        }
    },
    {
        "name": "doc_inspect",
        "description": "Extract structured JSON AST (headings, paragraphs, tables, slide geometries, cell values) from DOCX, PPTX, or XLSX.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to .docx, .pptx, or .xlsx document"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional worksheet name when inspecting XLSX files"
                }
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "doc_insert_diagram",
        "description": "Surgically inject a high-resolution diagram image into Word DOCX or PowerPoint PPTX with caption and keepNext chaining.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_path": {
                    "type": "string",
                    "description": "Absolute path to target .docx or .pptx document"
                },
                "image_path": {
                    "type": "string",
                    "description": "Absolute path to diagram PNG or JPG image"
                },
                "heading": {
                    "type": "string",
                    "description": "Target section heading in DOCX or slide title in PPTX"
                },
                "caption": {
                    "type": "string",
                    "description": "Caption text under the diagram (e.g. 'Figure 1: Architecture Overview')"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path (defaults to safe in-place update)"
                },
                "slide_index": {
                    "type": "integer",
                    "description": "1-based slide number for PPTX target (appends new slide if omitted)"
                }
            },
            "required": ["doc_path", "image_path"]
        }
    },
    {
        "name": "diagram_render",
        "description": "Render technical diagram spec file or source code (Mermaid, PlantUML, Draw.io XML, Canvas Spec JSON) to high-res PNG (300 DPI, A4-clamped).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to diagram source file (.json, .puml, .mmd, .drawio, .xml)"
                },
                "code": {
                    "type": "string",
                    "description": "Optional raw diagram source code string if file_path is omitted"
                },
                "engine": {
                    "type": "string",
                    "description": "Engine name for raw code: 'mermaid', 'plantuml', 'drawio', 'canvas'",
                    "default": "canvas"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path for rendered PNG"
                },
                "scale": {
                    "type": "number",
                    "description": "Resolution scale multiplier (default: 2)",
                    "default": 2
                }
            }
        }
    },
    {
        "name": "pptx_synthesize",
        "description": "Synthesize a 16:9 widescreen presentation deck from Word (.docx), PDF, or Markdown with Executive Deck Sizing.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "source_path": {
                    "type": "string",
                    "description": "Absolute path to source document (.docx, .pdf, or .md)"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path for the synthesized .pptx deck"
                },
                "theme": {
                    "type": "string",
                    "description": "Visual theme: 'thesis_blue', 'corporate_blue', 'modern_dark', 'academic_light'",
                    "default": "thesis_blue"
                },
                "template": {
                    "type": "string",
                    "description": "Optional path to base .pptx / .potx template for layout inheritance"
                },
                "deck_scope": {
                    "type": "string",
                    "description": "Presentation scope: 'executive' (15-25 slides for >= 50-100 page docs), 'standard' (30-45 slides), 'comprehensive' (full AST), or 'auto'",
                    "enum": ["auto", "executive", "standard", "comprehensive"],
                    "default": "auto"
                }
            },
            "required": ["source_path"]
        }
    },
    {
        "name": "env_doctor",
        "description": "Run Zero-Setup Environment Doctor diagnostics to verify Java/PlantUML, Chromium/Edge, Word COM, LibreOffice, Tesseract, and Python libraries.",
        "inputSchema": {
            "type": "object",
            "properties": {}
        }
    },
    {
        "name": "docx_apply_preset",
        "description": "Apply standard Vietnamese Administrative styling (Decree 30/2020/ND-CP), Academic Thesis, or Corporate Report preset to DOCX.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to target .docx file"
                },
                "preset": {
                    "type": "string",
                    "description": "Preset name: 'nd30', 'academic_thesis', 'corporate_report'",
                    "default": "nd30"
                },
                "add_header": {
                    "type": "boolean",
                    "description": "Whether to insert 2-column administrative header at the top",
                    "default": False
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path"
                }
            },
            "required": ["file_path"]
        }
    },
    {
        "name": "pdf_quick_tools",
        "description": "Comprehensive PDF Toolkit: merge, split, rotate, remove blank pages, watermark, compress, permanent redact sensitive info (CCCD/phone/email/salary), protect with AES-256 password & permissions, normalize to standard A4 with Bates Numbering, apply Giáp Lai stamp across pages, and extract tables to Excel (.xlsx).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "description": "Action: 'merge', 'split', 'rotate', 'clean', 'watermark', 'compress', 'redact', 'protect', 'normalize_a4', 'giap_lai', 'extract_tables'",
                    "enum": ["merge", "split", "rotate", "clean", "watermark", "compress", "redact", "protect", "normalize_a4", "giap_lai", "extract_tables"]
                },
                "input_path": {
                    "type": "string",
                    "description": "Path to primary input PDF"
                },
                "inputs": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of input PDF paths (for 'merge')"
                },
                "output_path": {
                    "type": "string",
                    "description": "Output PDF file path or output directory (for split) or Excel file path"
                },
                "ranges": {
                    "type": "string",
                    "description": "Page ranges for 'split' (e.g. '1-3, 4-5')"
                },
                "watermark_text": {
                    "type": "string",
                    "description": "Watermark text string (for 'watermark')"
                },
                "angle": {
                    "type": "integer",
                    "description": "Rotation angle (90, 180, 270)"
                },
                "patterns": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Sensitive pattern categories for 'redact' (e.g. ['cccd', 'phone', 'email', 'bank_card', 'salary'])"
                },
                "custom_terms": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Custom keywords or phrases to redact"
                },
                "user_password": {
                    "type": "string",
                    "description": "User password to open PDF (for 'protect')"
                },
                "owner_password": {
                    "type": "string",
                    "description": "Owner password for full permissions (for 'protect')"
                },
                "stamp_image_path": {
                    "type": "string",
                    "description": "Path to round/oval stamp image file (for 'giap_lai')"
                },
                "bates_prefix": {
                    "type": "string",
                    "description": "Prefix for Bates numbering (e.g. 'HD-2026')"
                }
            },
            "required": ["action"]
        }
    },
    {
        "name": "doc_mail_merge",
        "description": "Batch Mail Merge Studio: inject structured Excel/JSON rows into Word template with leading zero preservation (ERR_XLSX_007), conditions {{IF:...}} and ZIP bundling.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "template_path": {
                    "type": "string",
                    "description": "Path to Word .docx template with {{Token}} placeholders"
                },
                "data_source": {
                    "description": "Path to .xlsx / .json file or direct list of dictionaries"
                },
                "output_dir": {
                    "type": "string",
                    "description": "Output directory for merged .docx documents",
                    "default": "./mail_merge_output"
                },
                "filename_pattern": {
                    "type": "string",
                    "description": "Pattern for naming output files, e.g. 'HDLD_{{MaNV}}_{{HoTen}}.docx'"
                },
                "zip_output": {
                    "type": "string",
                    "description": "Optional ZIP archive path to bundle all generated documents"
                },
                "inspect_only": {
                    "type": "boolean",
                    "description": "Inspect and return all {{Token}} placeholders without generating documents",
                    "default": False
                }
            },
            "required": ["template_path"]
        }
    },
    {
        "name": "doc_redline_diff",
        "description": "Compare two Word documents and generate a legal redline Track Changes (<w:ins>, <w:del>) document with comments (ERR_DOCX_008).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_a_path": {
                    "type": "string",
                    "description": "Path to original/baseline .docx document"
                },
                "doc_b_path": {
                    "type": "string",
                    "description": "Path to revised/modified .docx document"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path for redline .docx (defaults to <doc_b>_redline.docx)"
                },
                "author": {
                    "type": "string",
                    "description": "Author name for revisions and comments",
                    "default": "AI Reviewer"
                }
            },
            "required": ["doc_a_path", "doc_b_path"]
        }
    },
    {
        "name": "doc_ocr_scan",
        "description": "Vietnamese OCR Pipeline for scanned PDFs and images: deskews, neutralizes red stamps (ERR_CONV_005), and recovers text/tables into Markdown or DOCX.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {
                    "type": "string",
                    "description": "Path to scanned PDF or image file (.png, .jpg, .jpeg, .tiff)"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path (.md or .docx)"
                },
                "lang": {
                    "type": "string",
                    "description": "OCR languages (default: 'vie+eng')",
                    "default": "vie+eng"
                },
                "filter_red_stamp": {
                    "type": "boolean",
                    "description": "Whether to filter and neutralize red circular seal stamps (ERR_CONV_005)",
                    "default": True
                }
            },
            "required": ["input_path"]
        }
    },
    {
        "name": "xlsx_set_cell",
        "description": "Set value, formula, or formatting for an Excel cell in-place preserving DrawingML logos (ERR_XLSX_006) and safe merged cells (ERR_XLSX_004).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to target .xlsx document"
                },
                "coordinate": {
                    "type": "string",
                    "description": "Cell coordinate (e.g. 'F12', 'B5')"
                },
                "value": {
                    "description": "New value to assign to cell"
                },
                "formula": {
                    "type": "string",
                    "description": "Optional formula string starting with '=', e.g. '=SUM(B2:B10)'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional worksheet name (defaults to active sheet)"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path (defaults to in-place update)"
                }
            },
            "required": ["file_path", "coordinate"]
        }
    },
    {
        "name": "xlsx_set_range",
        "description": "Batch update multiple cells in Excel workbook in-place, e.g. {'B5': '1,500,000', 'C5': 'Approved'}.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to target .xlsx document"
                },
                "updates": {
                    "type": "object",
                    "description": "Dictionary mapping cell coordinates to values/formula objects"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional worksheet name"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path"
                }
            },
            "required": ["file_path", "updates"]
        }
    },
    {
        "name": "xlsx_insert_rows",
        "description": "Insert N rows at specified index, cloning prototype formatting (ERR_XLSX_001) and shifting downstream formulas via Regex AST (ERR_XLSX_002).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to target .xlsx document"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Worksheet name"
                },
                "start_row": {
                    "type": "integer",
                    "description": "1-based row index where new rows are inserted"
                },
                "count": {
                    "type": "integer",
                    "description": "Number of rows to insert",
                    "default": 1
                },
                "prototype_row": {
                    "type": "integer",
                    "description": "Optional row index to clone formatting from"
                },
                "shift_formulas": {
                    "type": "boolean",
                    "description": "Whether to shift downstream formulas",
                    "default": True
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path"
                }
            },
            "required": ["file_path", "sheet_name", "start_row"]
        }
    },
    {
        "name": "xlsx_find_anchor",
        "description": "Semantic Anchor Discovery (ERR_XLSX_003): Scans worksheet to discover row and column coordinates of a keyword.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to target .xlsx document"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Worksheet name"
                },
                "keyword": {
                    "type": "string",
                    "description": "Keyword string to search for"
                },
                "max_row": {
                    "type": "integer",
                    "description": "Maximum rows to scan (default: 150)",
                    "default": 150
                }
            },
            "required": ["file_path", "sheet_name", "keyword"]
        }
    },
    {
        "name": "docx_inspect_placeholders",
        "description": "Inspect all {{TOKEN}} variables, {{IF:...}} conditions, and {{BLOCK:...}} anchors in a Word .docx template.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "template_path": {
                    "type": "string",
                    "description": "Path to Word .docx template file"
                }
            },
            "required": ["template_path"]
        }
    },
    {
        "name": "docx_replace_text",
        "description": "Replace text/placeholders in Word DOCX body, tables, and headers/footers in-place preserving run formatting (ERR_DOCX_002).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_path": {
                    "type": "string",
                    "description": "Path to target Word .docx document"
                },
                "replacements": {
                    "type": "object",
                    "description": "Dict of {'KEY': 'Val'} or {'{{KEY}}': 'Val'} or {'old_text': 'new_text'}"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path"
                }
            },
            "required": ["doc_path", "replacements"]
        }
    },
    {
        "name": "docx_insert_paragraph",
        "description": "Insert a paragraph with explicit typography (style, bold, italic, color) after a heading or at document end.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_path": {
                    "type": "string",
                    "description": "Path to target Word .docx document"
                },
                "text": {
                    "type": "string",
                    "description": "Paragraph text content"
                },
                "after_heading": {
                    "type": "string",
                    "description": "Optional heading text to insert after"
                },
                "style": {
                    "type": "string",
                    "description": "Style name ('Normal', 'Heading 1', 'List Bullet')",
                    "default": "Normal"
                },
                "bold": {
                    "type": "boolean",
                    "description": "Bold text",
                    "default": False
                },
                "italic": {
                    "type": "boolean",
                    "description": "Italic text",
                    "default": False
                },
                "font_size": {
                    "type": "number",
                    "description": "Font size in pt"
                },
                "color": {
                    "type": "string",
                    "description": "Hex color string (e.g. '003057')"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path"
                }
            },
            "required": ["doc_path", "text"]
        }
    },
    {
        "name": "docx_insert_table",
        "description": "Insert an OpenXML-compliant table (<w:cantSplit/>, <w:tblHeader/>, centered) into Word document.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_path": {
                    "type": "string",
                    "description": "Path to target Word .docx document"
                },
                "headers": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Table header column titles"
                },
                "rows": {
                    "type": "array",
                    "items": {
                        "type": "array",
                        "items": {"type": "string"}
                    },
                    "description": "2D table rows data"
                },
                "after_heading": {
                    "type": "string",
                    "description": "Optional heading text to insert after"
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path"
                }
            },
            "required": ["doc_path", "headers", "rows"]
        }
    },
    {
        "name": "xlsx_formula_helper",
        "description": "Look up Excel formulas, syntax templates, real-world use cases, and technical architectural rationales across 32+ curated formulas.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "category": {
                    "type": "string",
                    "description": "Optional category filter: 'lookup', 'aggregation', 'text', 'logic', 'dynamic', 'financial', or 'all'",
                    "default": "all"
                },
                "search": {
                    "type": "string",
                    "description": "Optional search query by function name, keywords, or scenario"
                }
            }
        }
    },
    {
        "name": "yearbook_auto_layout",
        "description": "Smart auto-layout and aspect-ratio clustering for graduation yearbooks and photobooks into 9 editorial spread templates.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "photos": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "List of photo file paths or directory path"
                },
                "title": {
                    "type": "string",
                    "description": "Album title",
                    "default": "KỶ YẾU TỐT NGHIỆP 2026"
                },
                "subtitle": {
                    "type": "string",
                    "description": "Album subtitle"
                },
                "author": {
                    "type": "string",
                    "description": "Graduate / Owner name",
                    "default": "Biện Trung Tín"
                },
                "degree": {
                    "type": "string",
                    "description": "Academic degree / major"
                },
                "date": {
                    "type": "string",
                    "description": "Graduation date string"
                },
                "preset": {
                    "type": "string",
                    "description": "Preset key: 'photobook_square_25x25', 'a4_landscape', 'a4_portrait', 'magazine_b5'",
                    "default": "photobook_square_25x25"
                },
                "bleed_mm": {
                    "type": "number",
                    "description": "Pre-press bleed margin in mm (default: 3.0)",
                    "default": 3.0
                }
            }
        }
    },
    {
        "name": "yearbook_export",
        "description": "Build industrial print-ready PDF (300 DPI, 3mm Bleed, Crop Marks) or full 5-in-1 Designer Handoff Bundle (.zip).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "spec": {
                    "type": "object",
                    "description": "Yearbook Project Spec JSON object"
                },
                "output_path": {
                    "type": "string",
                    "description": "Destination file path (.zip or .pdf)"
                },
                "export_type": {
                    "type": "string",
                    "description": "Export type: 'handoff_zip', 'print_pdf', or 'spread_pdf'",
                    "default": "handoff_zip"
                }
            },
            "required": ["spec", "output_path"]
        }
    },
    {
        "name": "doc_security_audit",
        "description": "Audit & sanitize document security threats, malicious formulas/DDE, OLE binaries, Zip Slip, XXE, and sensitive PII leaks (Cloudflare Harness ERR_SEC_001–005).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Path to document to audit (.docx, .xlsx, .pptx, .pdf, .csv, .drawio, .xml, .zip)"
                },
                "sanitize": {
                    "type": "boolean",
                    "description": "Whether to sanitize/defuse detected threats",
                    "default": False
                },
                "output_path": {
                    "type": "string",
                    "description": "Optional destination path for sanitized file"
                },
                "defuse_formulas": {
                    "type": "boolean",
                    "description": "Defuse formula/DDE injection by prepending apostrophe",
                    "default": True
                },
                "strip_metadata": {
                    "type": "boolean",
                    "description": "Strip author and sensitive metadata from OpenXML core properties",
                    "default": True
                },
                "unhide_sheets": {
                    "type": "boolean",
                    "description": "Unhide veryHidden sheets in Excel workbooks",
                    "default": True
                }
            },
            "required": ["file_path"]
        }
    }
]


def log_debug(message: str) -> None:
    """Logs debug message to stderr so stdout remains clean JSON-RPC."""
    print(f"[MCP-SERVER] {message}", file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# Tool Implementations
# ---------------------------------------------------------------------------

def execute_doc_convert(args: Dict[str, Any]) -> Dict[str, Any]:
    src = os.path.abspath(args["source_path"])
    target_ext = args.get("target_ext", ".docx")
    if not target_ext.startswith("."):
        target_ext = "." + target_ext

    out_path = args.get("output_path")
    if not out_path:
        base, _ = os.path.splitext(src)
        out_path = f"{base}_converted{target_ext}"
    else:
        out_path = os.path.abspath(out_path)

    page_range = args.get("page_range")
    ok, res, elap = convert_universal(src, out_path, page_range=page_range)
    if not ok:
        raise RuntimeError(f"Conversion failed: {res}")

    return {
        "success": True,
        "output_path": out_path,
        "elapsed_sec": round(elap, 3),
        "source_format": os.path.splitext(src)[1],
        "target_format": target_ext,
    }


def execute_doc_inspect(args: Dict[str, Any]) -> Dict[str, Any]:
    file_path = os.path.abspath(args["file_path"])
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        return read_docx(file_path)
    elif ext == ".pptx":
        return read_pptx(file_path)
    elif ext in (".xlsx", ".xlsm"):
        sheet_name = args.get("sheet_name")
        return read_xlsx(file_path, sheet_name=sheet_name)
    else:
        raise ValueError(f"Unsupported document format for inspect: {ext}")


def execute_doc_insert_diagram(args: Dict[str, Any]) -> Dict[str, Any]:
    doc_path = os.path.abspath(args["doc_path"])
    image_path = os.path.abspath(args["image_path"])
    heading = args.get("heading")
    caption = args.get("caption")
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else doc_path

    if not os.path.isfile(doc_path):
        raise FileNotFoundError(f"Document not found: {doc_path}")
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"Diagram image not found: {image_path}")

    ext = os.path.splitext(doc_path)[1].lower()
    if ext == ".docx":
        saved = inject_diagram_into_docx(
            docx_path=doc_path,
            png_path=image_path,
            heading=heading,
            caption=caption,
            output_path=out_path,
        )
        return {"success": True, "output_path": saved, "document_type": "docx"}
    elif ext == ".pptx":
        slide_idx = args.get("slide_index")
        saved = insert_diagram_into_pptx(
            pptx_path=doc_path,
            image_path=image_path,
            slide_index=slide_idx,
            title=heading,
            caption=caption,
            output_path=out_path,
        )
        return {"success": True, "output_path": saved, "document_type": "pptx"}
    else:
        raise ValueError(f"Unsupported document format for diagram insertion: {ext}")


def execute_diagram_render(args: Dict[str, Any]) -> Dict[str, Any]:
    file_path = args.get("file_path")
    code = args.get("code")
    out_path = args.get("output_path")
    scale = int(args.get("scale", 2))

    if file_path:
        res = render_diagram_file_to_png(file_path, output_png=out_path, scale=scale)
        if not res or not res.get("success"):
            raise RuntimeError(f"Failed to render diagram file: {file_path}")
        return res
    elif code:
        engine = args.get("engine", "canvas")
        res = render_diagram_code_to_png(code, lang=engine, output_png=out_path, scale=scale)
        if not res or not res.get("success"):
            raise RuntimeError(f"Failed to render diagram code via {engine}")
        return res
    else:
        raise ValueError("Either 'file_path' or 'code' must be provided.")


def execute_pptx_synthesize(args: Dict[str, Any]) -> Dict[str, Any]:
    src = os.path.abspath(args["source_path"])
    if not os.path.isfile(src):
        raise FileNotFoundError(f"Source file not found: {src}")

    theme = args.get("theme", "thesis_blue")
    template = args.get("template")
    deck_scope = args.get("deck_scope", args.get("scope", "auto"))

    out_path = args.get("output_path")
    if not out_path:
        base, _ = os.path.splitext(src)
        out_path = f"{base}_presentation.pptx"
    else:
        out_path = os.path.abspath(out_path)

    ext = os.path.splitext(src)[1].lower()
    if ext == ".docx":
        final_pptx, _ = convert_docx_to_pptx(
            src, out_path, theme=theme, template_path=template, deck_scope=deck_scope
        )
        res_path = final_pptx
    elif ext == ".pdf":
        final_pptx, _ = convert_pdf_to_pptx(
            src, out_path, theme=theme, template_path=template, deck_scope=deck_scope
        )
        res_path = final_pptx
    elif ext in (".md", ".markdown"):
        with open(src, "r", encoding="utf-8") as f:
            md_text = f.read()
        slides_spec = parse_markdown_to_slides_spec(md_text, default_theme=theme)
        res_path = write_pptx_from_spec(slides_spec, out_path, template_path=template)
    else:
        raise ValueError(f"Unsupported source format for presentation synthesis: {ext}")


    return {
        "success": True,
        "output_path": res_path,
        "theme": theme,
        "source_format": ext,
    }


def execute_env_doctor(args: Dict[str, Any]) -> Dict[str, Any]:
    return run_doctor()


def execute_docx_apply_preset(args: Dict[str, Any]) -> Dict[str, Any]:
    import docx
    try:
        from core.docx.presets import apply_preset_to_document, build_administrative_header
        from core.docx.advanced_engine import safe_save_docx
    except (ImportError, ValueError):
        from docx_presets import apply_preset_to_document, build_administrative_header  # type: ignore
        from docx_advanced_engine import safe_save_docx  # type: ignore

    file_path = os.path.abspath(args["file_path"])
    preset = args.get("preset", "nd30")
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else file_path

    doc = docx.Document(file_path)
    res = apply_preset_to_document(doc, preset=preset)
    if args.get("add_header", False):
        build_administrative_header(
            doc,
            agency_parent=args.get("agency_parent", "BỘ THÔNG TIN VÀ TRUYỀN THÔNG"),
            agency_name=args.get("agency_name", "CỤC CÔNG NGHỆ THÔNG TIN"),
            document_no=args.get("doc_no", "Số: 01/QĐ-CNTT"),
            location_date=args.get("date", "Hà Nội, ngày 18 tháng 10 năm 2026"),
            insert_at_top=True,
        )
    saved = safe_save_docx(doc, out_path)
    return {
        "success": True,
        "output_path": saved,
        "preset": preset,
        "details": res,
    }


def execute_pdf_quick_tools(args: Dict[str, Any]) -> Dict[str, Any]:
    try:
        from core.converter.pdf_toolkit import (
            merge_pdfs,
            split_pdf,
            rotate_and_clean_pdf,
            stamp_watermark_pdf,
            compress_pdf,
            redact_sensitive_pdf,
            protect_pdf,
            normalize_and_number_pdf,
            stamp_giap_lai_pdf,
            extract_pdf_tables_to_xlsx,
        )
    except (ImportError, ValueError):
        from pdf_toolkit import (  # type: ignore
            merge_pdfs,
            split_pdf,
            rotate_and_clean_pdf,
            stamp_watermark_pdf,
            compress_pdf,
            redact_sensitive_pdf,
            protect_pdf,
            normalize_and_number_pdf,
            stamp_giap_lai_pdf,
            extract_pdf_tables_to_xlsx,
        )

    action = args["action"]
    inp = args.get("input_path")
    out = args.get("output_path")

    if action == "merge":
        inputs = args.get("inputs") or ([inp] if inp else [])
        if not inputs or len(inputs) < 2:
            raise ValueError("'merge' requires at least 2 input paths")
        res = merge_pdfs(inputs, out or "merged.pdf")
        return {"success": True, "output_path": res, "merged_files": inputs}

    elif action == "split":
        if not inp:
            raise ValueError("'split' requires 'input_path'")
        files = split_pdf(inp, out or "./split_output", page_ranges=args.get("ranges"))
        return {"success": True, "generated_files": files, "count": len(files)}

    elif action in ("rotate", "clean"):
        if not inp:
            raise ValueError(f"'{action}' requires 'input_path'")
        return rotate_and_clean_pdf(
            inp,
            out or inp,
            rotation=int(args.get("angle", 0) if action == "rotate" else 0),
            remove_blank=bool(args.get("remove_blank", False) or action == "clean"),
        )

    elif action == "watermark":
        if not inp:
            raise ValueError("'watermark' requires 'input_path'")
        text = args.get("watermark_text", "BẢN QUYỀN")
        res = stamp_watermark_pdf(inp, out or inp, watermark_text=text)
        return {"success": True, "output_path": res, "watermark": text}

    elif action == "compress":
        if not inp:
            raise ValueError("'compress' requires 'input_path'")
        return compress_pdf(inp, out or inp)

    elif action == "redact":
        if not inp:
            raise ValueError("'redact' requires 'input_path'")
        return redact_sensitive_pdf(
            inp,
            out or os.path.splitext(inp)[0] + "_redacted.pdf",
            patterns=args.get("patterns"),
            custom_terms=args.get("custom_terms"),
        )

    elif action == "protect":
        if not inp:
            raise ValueError("'protect' requires 'input_path'")
        return protect_pdf(
            inp,
            out or os.path.splitext(inp)[0] + "_protected.pdf",
            user_password=args.get("user_password"),
            owner_password=args.get("owner_password"),
            allow_print=bool(args.get("allow_print", True)),
            allow_copy=bool(args.get("allow_copy", False)),
            allow_modify=bool(args.get("allow_modify", False)),
        )

    elif action in ("normalize_a4", "normalize-a4"):
        if not inp:
            raise ValueError("'normalize_a4' requires 'input_path'")
        return normalize_and_number_pdf(
            inp,
            out or os.path.splitext(inp)[0] + "_a4.pdf",
            bates_prefix=args.get("bates_prefix"),
        )

    elif action in ("giap_lai", "giap-lai"):
        if not inp:
            raise ValueError("'giap_lai' requires 'input_path'")
        stamp = args.get("stamp_image_path")
        if not stamp:
            raise ValueError("'giap_lai' requires 'stamp_image_path'")
        return stamp_giap_lai_pdf(
            inp,
            out or os.path.splitext(inp)[0] + "_giap_lai.pdf",
            stamp,
            side=args.get("side", "right"),
        )

    elif action in ("extract_tables", "extract-tables"):
        if not inp:
            raise ValueError("'extract_tables' requires 'input_path'")
        return extract_pdf_tables_to_xlsx(
            inp,
            out or os.path.splitext(inp)[0] + "_tables.xlsx",
        )

    else:
        raise ValueError(f"Unknown PDF action: {action}")


def execute_doc_mail_merge(args: Dict[str, Any]) -> Dict[str, Any]:
    try:
        from core.docx.mail_merge import execute_mail_merge, inspect_template_tokens
    except (ImportError, ValueError):
        from docx_mail_merge import execute_mail_merge, inspect_template_tokens  # type: ignore

    template_path = args.get("template_path")
    if not template_path:
        raise ValueError("template_path is required for doc_mail_merge")
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Template DOCX file not found: {template_path}")

    if args.get("inspect_only"):
        tokens = inspect_template_tokens(template_path)
        return {"success": True, "template": template_path, "tokens": tokens}

    data_source = args.get("data_source")
    if not data_source:
        raise ValueError("data_source (.xlsx, .json or list) is required for mail merge")

    out_dir = args.get("output_dir", "./mail_merge_output")
    pattern = args.get("filename_pattern")
    zip_out = args.get("zip_output")

    return execute_mail_merge(
        template_path=template_path,
        data_source=data_source,
        output_dir=out_dir,
        filename_pattern=pattern,
        zip_output=zip_out,
    )


def execute_doc_redline_diff(args: Dict[str, Any]) -> Dict[str, Any]:
    try:
        from core.docx.diff_engine import compare_docx_documents
    except (ImportError, ValueError):
        from docx_diff_engine import compare_docx_documents  # type: ignore

    doc_a = args.get("doc_a_path")
    doc_b = args.get("doc_b_path")
    if not doc_a or not doc_b:
        raise ValueError("doc_a_path and doc_b_path are required for doc_redline_diff")

    out = args.get("output_path") or os.path.splitext(doc_b)[0] + "_redline.docx"
    author = args.get("author", "AI Reviewer")

    return compare_docx_documents(doc_a, doc_b, out, author=author)


def execute_doc_ocr_scan(args: Dict[str, Any]) -> Dict[str, Any]:
    try:
        from core.converter.ocr_pipeline import VietnameseOcrPipeline
    except (ImportError, ValueError):
        from ocr_pipeline import VietnameseOcrPipeline  # type: ignore

    inp = args.get("input_path")
    if not inp:
        raise ValueError("input_path is required for doc_ocr_scan")

    out = args.get("output_path")
    lang = args.get("lang", "vie+eng")
    filter_stamp = bool(args.get("filter_red_stamp", True))

    pipeline = VietnameseOcrPipeline(lang=lang, filter_red_stamp=filter_stamp)
    return pipeline.process_document(inp, output_path=out)


def execute_xlsx_set_cell(args: Dict[str, Any]) -> Dict[str, Any]:
    file_path = os.path.abspath(args["file_path"])
    coord_param = args.get("coordinate") or args.get("cell")
    if not coord_param:
        raise KeyError("'coordinate' or 'cell' must be provided")
    coordinate = str(coord_param).strip().upper()
    value = args.get("value")
    formula = args.get("formula")
    sheet_name = args.get("sheet_name")
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else file_path

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    engine = UniversalExcelEngine(file_path)
    sheet_obj = engine.get_sheet(sheet_name) if sheet_name else engine.wb.active
    ws = sheet_obj if sheet_obj is not None else engine.wb.worksheets[0]
    engine.safe_set_cell(ws, coordinate, value=value, formula=formula)
    engine.save(out_path)
    sheet_title = getattr(ws, "title", "Sheet")
    return {"success": True, "output_path": out_path, "coordinate": coordinate, "sheet": sheet_title}


def execute_xlsx_set_range(args: Dict[str, Any]) -> Dict[str, Any]:
    file_path = os.path.abspath(args["file_path"])
    updates = args["updates"]
    sheet_name = args.get("sheet_name")
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else file_path

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")
    if not isinstance(updates, dict):
        raise ValueError("'updates' must be a dictionary of cell coordinates to values/specs")

    engine = UniversalExcelEngine(file_path)
    sheet_obj = engine.get_sheet(sheet_name) if sheet_name else engine.wb.active
    ws = sheet_obj if sheet_obj is not None else engine.wb.worksheets[0]
    for coord, val in updates.items():
        coord_clean = str(coord).strip().upper()
        if isinstance(val, dict):
            engine.safe_set_cell(ws, coord_clean, value=val.get("value"), formula=val.get("formula"))
        else:
            engine.safe_set_cell(ws, coord_clean, value=val)
    engine.save(out_path)
    sheet_title = getattr(ws, "title", "Sheet")
    return {"success": True, "output_path": out_path, "updated_cells": len(updates), "sheet": sheet_title}


def execute_xlsx_insert_rows(args: Dict[str, Any]) -> Dict[str, Any]:
    file_path = os.path.abspath(args["file_path"])
    sheet_name = args["sheet_name"]
    start_row = int(args["start_row"])
    count = int(args.get("count", 1))
    proto_row = args.get("prototype_row")
    if proto_row is not None:
        proto_row = int(proto_row)
    shift_formulas = bool(args.get("shift_formulas", True))
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else file_path

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    engine = UniversalExcelEngine(file_path)
    created_rows = engine.expand_table_rows(
        sheet_name=sheet_name,
        start_row=start_row,
        count=count,
        prototype_row=proto_row,
        shift_formulas=shift_formulas,
    )
    engine.save(out_path)
    return {"success": True, "output_path": out_path, "created_rows": created_rows, "count": count, "sheet": sheet_name}


def execute_xlsx_find_anchor(args: Dict[str, Any]) -> Dict[str, Any]:
    from openpyxl.utils import get_column_letter

    file_path = os.path.abspath(args["file_path"])
    sheet_name = args["sheet_name"]
    keyword = str(args["keyword"])
    max_row = int(args.get("max_row", 150))

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    engine = UniversalExcelEngine(file_path)
    coord = engine.find_anchor(sheet_name, keyword, max_row=max_row)
    if coord:
        r, c = coord
        return {
            "success": True,
            "found": True,
            "row": r,
            "col": c,
            "coordinate": f"{get_column_letter(c)}{r}",
            "sheet": sheet_name,
        }
    return {"success": True, "found": False, "keyword": keyword, "sheet": sheet_name}


def execute_docx_inspect_placeholders(args: Dict[str, Any]) -> Dict[str, Any]:
    template_path = os.path.abspath(args["template_path"])
    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Word template not found: {template_path}")
    return inspect_template_tokens(template_path)


def execute_docx_replace_text(args: Dict[str, Any]) -> Dict[str, Any]:
    doc_path = os.path.abspath(args["doc_path"])
    replacements = args["replacements"]
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else doc_path

    if not os.path.isfile(doc_path):
        raise FileNotFoundError(f"Word document not found: {doc_path}")
    if not isinstance(replacements, dict):
        raise ValueError("'replacements' must be a dictionary")

    # Normalize token keys: support both "KEY" and "{{KEY}}"
    normalized_rep: Dict[str, str] = {}
    for k, v in replacements.items():
        k_str = str(k).strip()
        v_str = str(v)
        normalized_rep[k_str] = v_str
        if not k_str.startswith("{{") and not k_str.endswith("}}"):
            normalized_rep[f"{{{{{k_str}}}}}"] = v_str

    res_path = patch_docx_template(doc_path, out_path, text_replacements=normalized_rep)
    return {"success": True, "output_path": res_path, "replacements_count": len(replacements)}


def execute_docx_insert_paragraph(args: Dict[str, Any]) -> Dict[str, Any]:
    from docx import Document
    from docx.shared import Pt, RGBColor

    doc_path = os.path.abspath(args["doc_path"])
    text = str(args["text"])
    after_heading = args.get("after_heading")
    style = args.get("style", "Normal")
    bold = bool(args.get("bold", False))
    italic = bool(args.get("italic", False))
    font_size = args.get("font_size")
    color = args.get("color")
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else doc_path

    if not os.path.isfile(doc_path):
        raise FileNotFoundError(f"Word document not found: {doc_path}")

    doc = Document(doc_path)
    target_idx = None
    if after_heading:
        h_clean = after_heading.strip().lower()
        for idx, p in enumerate(doc.paragraphs):
            if p.text and h_clean in p.text.strip().lower():
                target_idx = idx
                break

    if target_idx is not None and target_idx < len(doc.paragraphs) - 1:
        ref_p = doc.paragraphs[target_idx]
        new_p = doc.add_paragraph()
        ref_p._p.addnext(new_p._p)
    else:
        new_p = doc.add_paragraph()

    run = new_p.add_run(text)
    if style in doc.styles:
        new_p.style = doc.styles[style]
    run.font.bold = bold
    run.font.italic = italic
    if font_size:
        run.font.size = Pt(float(font_size))
    if color:
        c_clean = str(color).lstrip("#").strip()
        if len(c_clean) == 6:
            run.font.color.rgb = RGBColor.from_string(c_clean)

    doc.save(out_path)
    return {"success": True, "output_path": out_path, "text": text}


def execute_docx_insert_table(args: Dict[str, Any]) -> Dict[str, Any]:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import parse_xml

    doc_path = os.path.abspath(args["doc_path"])
    headers = args["headers"]
    rows = args["rows"]
    after_heading = args.get("after_heading")
    out_path = os.path.abspath(args["output_path"]) if args.get("output_path") else doc_path

    if not os.path.isfile(doc_path):
        raise FileNotFoundError(f"Word document not found: {doc_path}")
    if not isinstance(headers, list) or not isinstance(rows, list):
        raise ValueError("'headers' and 'rows' must be lists")

    doc = Document(doc_path)
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for r_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        if r_idx == 0:
            trPr.append(parse_xml(r'<w:tblHeader xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
            for c_idx, h in enumerate(headers):
                cell = row.cells[c_idx]
                cell.paragraphs[0].text = str(h)
                if cell.paragraphs[0].runs:
                    cell.paragraphs[0].runs[0].font.bold = True
        else:
            data_row = rows[r_idx - 1]
            for c_idx, val in enumerate(data_row):
                if c_idx < len(row.cells):
                    row.cells[c_idx].paragraphs[0].text = str(val)

    if after_heading:
        h_clean = after_heading.strip().lower()
        for p in doc.paragraphs:
            if p.text and h_clean in p.text.strip().lower():
                p._p.addnext(table._tbl)
                break

    doc.save(out_path)
    return {"success": True, "output_path": out_path, "total_rows": len(rows) + 1, "cols": len(headers)}


def execute_xlsx_formula_helper(args: Dict[str, Any]) -> Dict[str, Any]:
    cat = args.get("category", "all")
    search = args.get("search")
    try:
        from core.xlsx.formula_playbook import get_formula_catalog
    except (ImportError, ValueError):
        from formula_playbook import get_formula_catalog  # type: ignore

    formulas = get_formula_catalog(category=cat, search=search)
    return {
        "success": True,
        "category": cat,
        "search": search,
        "total_count": len(formulas),
        "formulas": formulas,
    }


def execute_yearbook_auto_layout(args: Dict[str, Any]) -> Dict[str, Any]:
    from core.yearbook import YearbookLayoutEngine, get_default_mock_spec

    photos = args.get("photos", [])
    title = args.get("title", "KỶ YẾU TỐT NGHIỆP 2026")
    subtitle = args.get("subtitle", "Hành Trình Rực Rỡ & Những Năm Tháng Thanh Xuân")
    author = args.get("author", "Biện Trung Tín")
    degree = args.get("degree", "Cử Nhân Kỹ Thuật Phần Mềm")
    date_str = args.get("date", "Tháng 11 / 2026")
    preset = args.get("preset", "photobook_square_25x25")
    bleed_mm = float(args.get("bleed_mm", 3.0))

    if not photos:
        spec = get_default_mock_spec(
            title=title, subtitle=subtitle, author=author, degree=degree,
            date_str=date_str, preset_key=preset, bleed_mm=bleed_mm
        )
    else:
        spec = YearbookLayoutEngine.auto_layout(
            photo_paths=photos,
            title=title,
            subtitle=subtitle,
            author=author,
            degree=degree,
            date_str=date_str,
            preset_key=preset,
            bleed_mm=bleed_mm,
        )

    return {
        "success": True,
        "total_spreads": len(spec.get("pages", [])),
        "spec": spec,
    }


def execute_yearbook_export(args: Dict[str, Any]) -> Dict[str, Any]:
    from core.yearbook import YearbookPressEngine, YearbookHandoffPackager

    spec = args.get("spec")
    if not spec:
        raise ValueError("spec is required")
    out_path = args.get("output_path")
    if not out_path:
        raise ValueError("output_path is required")

    export_type = args.get("export_type", "handoff_zip").lower()
    if export_type == "print_pdf":
        actual_out = YearbookPressEngine.export_print_ready_pdf(spec, out_path)
    elif export_type == "spread_pdf":
        actual_out = YearbookPressEngine.export_spread_preview_pdf(spec, out_path)
    else:
        actual_out = YearbookHandoffPackager.export_handoff_bundle(spec, out_path)

    return {
        "success": True,
        "export_type": export_type,
        "output_path": actual_out,
        "file_size": os.path.getsize(actual_out) if os.path.isfile(actual_out) else 0,
    }


def execute_doc_security_audit(args: Dict[str, Any]) -> Dict[str, Any]:
    file_path = args.get("file_path")
    if not file_path:
        raise ValueError("file_path is required")
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"File not found: {abs_path}")

    from core.security.document_auditor import DocumentSecurityAuditor
    auditor = DocumentSecurityAuditor()
    report = auditor.audit_file(abs_path)

    result = {
        "success": True,
        "target": report.target,
        "is_clean": report.is_clean,
        "critical_count": report.critical_count,
        "warning_count": report.warning_count,
        "info_count": report.info_count,
        "report": report.to_dict(),
    }

    if args.get("sanitize", False):
        from core.security.sanitizer import DocumentSanitizer
        sanitizer = DocumentSanitizer()
        san_res = sanitizer.sanitize(
            abs_path,
            output_path=args.get("output_path"),
            defuse_formulas=args.get("defuse_formulas", True),
            strip_metadata=args.get("strip_metadata", True),
            unhide_sheets=args.get("unhide_sheets", True),
        )
        result["sanitization"] = san_res

    return result


TOOL_DISPATCHER = {
    "doc_convert": execute_doc_convert,
    "doc_inspect": execute_doc_inspect,
    "doc_insert_diagram": execute_doc_insert_diagram,
    "diagram_render": execute_diagram_render,
    "pptx_synthesize": execute_pptx_synthesize,
    "env_doctor": execute_env_doctor,
    "docx_apply_preset": execute_docx_apply_preset,
    "pdf_quick_tools": execute_pdf_quick_tools,
    "doc_mail_merge": execute_doc_mail_merge,
    "doc_redline_diff": execute_doc_redline_diff,
    "doc_ocr_scan": execute_doc_ocr_scan,
    # Atomic Office Tools
    "xlsx_set_cell": execute_xlsx_set_cell,
    "xlsx_set_range": execute_xlsx_set_range,
    "xlsx_insert_rows": execute_xlsx_insert_rows,
    "xlsx_find_anchor": execute_xlsx_find_anchor,
    "docx_inspect_placeholders": execute_docx_inspect_placeholders,
    "docx_replace_text": execute_docx_replace_text,
    "docx_insert_paragraph": execute_docx_insert_paragraph,
    "docx_insert_table": execute_docx_insert_table,
    "xlsx_formula_helper": execute_xlsx_formula_helper,
    # Yearbook Studio Tools
    "yearbook_auto_layout": execute_yearbook_auto_layout,
    "yearbook_export": execute_yearbook_export,
    # Cloudflare Security Audit Tool
    "doc_security_audit": execute_doc_security_audit,
}


# ---------------------------------------------------------------------------
# JSON-RPC Message Dispatcher
# ---------------------------------------------------------------------------

def handle_rpc_request(request: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Process single JSON-RPC 2.0 request dictionary and return response dict."""
    req_id = request.get("id")
    method = request.get("method")
    params = request.get("params", {})

    # Notification without id
    if req_id is None and method == "notifications/initialized":
        log_debug("Client initialized.")
        return None

    if method == "initialize":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "protocolVersion": MCP_PROTOCOL_VERSION,
                "capabilities": {
                    "tools": {}
                },
                "serverInfo": {
                    "name": SERVER_NAME,
                    "version": SERVER_VERSION
                }
            }
        }

    elif method == "ping":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {}
        }

    elif method == "tools/list":
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "result": {
                "tools": TOOL_DEFINITIONS
            }
        }

    elif method == "tools/call":
        tool_name = params.get("name")
        arguments = params.get("arguments", {})

        handler = TOOL_DISPATCHER.get(tool_name)
        if not handler:
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "error": {
                    "code": -32601,
                    "message": f"Unknown tool: '{tool_name}'"
                }
            }

        try:
            result_data = handler(arguments)
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [
                        {
                            "type": "text",
                            "text": json.dumps(result_data, indent=2, ensure_ascii=False)
                        }
                    ],
                    "isError": False
                }
            }
        except Exception as exc:
            err_msg = str(exc)
            log_debug(f"Error executing {tool_name}: {err_msg}")
            return {
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "content": [
                        {
                            "type": "text",
                            "text": json.dumps({
                                "error": err_msg,
                                "traceback": traceback.format_exc()
                            }, indent=2, ensure_ascii=False)
                        }
                    ],
                    "isError": True
                }
            }

    else:
        return {
            "jsonrpc": "2.0",
            "id": req_id,
            "error": {
                "code": -32601,
                "message": f"Method not found: '{method}'"
            }
        }


def run_stdio_server() -> None:
    """Reads NDJSON RPC requests from stdin and writes responses to stdout."""
    log_debug(f"Starting {SERVER_NAME} v{SERVER_VERSION} (STDIO transport)...")
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
        except Exception as exc:
            err_res = {
                "jsonrpc": "2.0",
                "id": None,
                "error": {
                    "code": -32700,
                    "message": f"Parse error: {str(exc)}"
                }
            }
            sys.stdout.write(json.dumps(err_res) + "\n")
            sys.stdout.flush()
            continue

        response = handle_rpc_request(request)
        if response is not None:
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()


def main() -> None:
    """Entry point for console scripts."""
    run_stdio_server()


if __name__ == "__main__":
    main()

