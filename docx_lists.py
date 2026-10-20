"""
docx_lists.py — High-Fidelity OpenXML List, Numbering & Bullet Engine for Word (.docx).

Provides native OpenXML ECMA-376 list formatting:
1. Native Dynamic Numbering:
   - Proper <w:numPr> with <w:numId> and <w:ilvl>.
   - Dynamic allocation of <w:numId> ensuring clean restarts at 1.0 for each new list.
   - Smart continuation: Sub-bullets/interleaved notes do not break or reset parent numbering.
   - Pressing Enter in Microsoft Word dynamically increments the next number.
2. 3-Tier Multi-Level Nesting:
   - Numbered: Level 0 (1, 2, 3) -> Level 1 (a, b, c) -> Level 2 (i, ii, iii).
   - Bullets: Level 0 (• Solid Disc) -> Level 1 (○ Hollow Circle) -> Level 2 (■ Solid Square).
3. Checkbox Task Lists:
   - - [ ] renders as ☐ (ballot box).
   - - [x] renders as ☑ (checked ballot box).
4. Invariant Protection:
   - Strict hanging indentation (<w:ind w:left="..." w:hanging="360"/>).
   - Zero auto-numbering on plain text unless explicitly formatted as a list.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple, Union

import docx
from docx.document import Document as DocumentType
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

NS_W = nsdecls("w")

# Base abstractNum IDs reserved for this engine
ABSTRACT_NUM_DECIMAL_ID = 2001
ABSTRACT_NUM_BULLET_ID = 2002

# Standard indentation unit in twips/dxa (1 inch = 1440 twips; 0.25 in = 360 twips; 0.5 in = 720 twips)
INDENT_HANGING_TWIPS = 360
INDENT_LEVEL_STEP_TWIPS = 720


@dataclass
class ParsedListItem:
    """Represents a structured parsed list item from text or AST."""
    item_type: str  # "ordered", "bullet", "checkbox"
    level: int  # 0, 1, 2
    text: str  # Clean content without list prefix
    is_checked: bool = False  # For checkbox items
    raw_number: Optional[int] = None  # Original number if ordered
    prefix: str = ""  # The original prefix matched


class ListContextTracker:
    """
    Tracks list nesting context, active numId instances, and handles
    smart continuation and restart triggers across a stream of lines.
    """

    def __init__(self, doc: DocumentType):
        self.doc = doc
        self.active_ordered_num_id: Optional[int] = None
        self.active_bullet_num_id: Optional[int] = None
        self.next_num_id_counter = 3000
        self.last_item_type: Optional[str] = None
        self.last_level: int = 0
        self.ordered_level_stack: List[int] = []

        # Ensure base abstractNum templates are registered in numbering.xml
        ensure_numbering_definitions(self.doc)

    def allocate_ordered_num_id(self) -> int:
        """Creates a brand new <w:num> element with startOverride=1 for clean restart across levels 0, 1, 2."""
        num_id = self.next_num_id_counter
        self.next_num_id_counter += 1

        n_el = self.doc.part.numbering_part.element
        new_num = parse_xml(
            f'<w:num {NS_W} w:numId="{num_id}">'
            f'  <w:abstractNumId w:val="{ABSTRACT_NUM_DECIMAL_ID}"/>'
            f'  <w:lvlOverride w:ilvl="0">'
            f'    <w:startOverride w:val="1"/>'
            f'  </w:lvlOverride>'
            f'  <w:lvlOverride w:ilvl="1">'
            f'    <w:startOverride w:val="1"/>'
            f'  </w:lvlOverride>'
            f'  <w:lvlOverride w:ilvl="2">'
            f'    <w:startOverride w:val="1"/>'
            f'  </w:lvlOverride>'
            f'</w:num>'
        )
        n_el.append(new_num)
        return num_id

    def get_bullet_num_id(self) -> int:
        """Retrieves or registers the shared bullet <w:num> instance."""
        if self.active_bullet_num_id is not None:
            return self.active_bullet_num_id

        num_id = self.next_num_id_counter
        self.next_num_id_counter += 1

        n_el = self.doc.part.numbering_part.element
        new_num = parse_xml(
            f'<w:num {NS_W} w:numId="{num_id}">'
            f'  <w:abstractNumId w:val="{ABSTRACT_NUM_BULLET_ID}"/>'
            f'</w:num>'
        )
        n_el.append(new_num)
        self.active_bullet_num_id = num_id
        return num_id

    def reset_ordered_list(self) -> None:
        """Explicitly resets the active ordered list context so the next item starts at 1."""
        self.active_ordered_num_id = None
        self.ordered_level_stack.clear()

    def reset_all_lists(self) -> None:
        """Resets all list context (called on headings, tables, or non-list breaks)."""
        self.reset_ordered_list()
        self.last_item_type = None
        self.last_level = 0

    def parse_markdown_line(self, line: str) -> Optional[ParsedListItem]:
        """
        Parses a markdown line to determine if it is an ordered, bullet, or checkbox list item.
        Respects leading spaces to determine indentation level (0, 1, 2).
        Returns None if the line is not a list item.
        """
        if not line or not line.strip():
            return None

        # Calculate indentation level based on leading whitespace
        l_stripped = line.lstrip(" \t")
        indent_len = len(line) - len(l_stripped)

        # 0-1 spaces -> Level 0; 2-3 spaces -> Level 1; >= 4 spaces -> Level 2
        if indent_len < 2:
            level = 0
        elif indent_len < 4:
            level = 1
        else:
            level = min(2, indent_len // 2)

        # 1. Checkbox Task List: - [ ] or - [x] or * [ ]
        m_check = re.match(r"^[-*+]\s+\[([ xX])\]\s+(.*)$", l_stripped)
        if m_check:
            is_checked = m_check.group(1).lower() == "x"
            clean_text = m_check.group(2).strip()
            return ParsedListItem(
                item_type="checkbox",
                level=level,
                text=clean_text,
                is_checked=is_checked,
                prefix=m_check.group(0)[:m_check.group(0).index(clean_text)]
            )

        # 2. Bullet List: - item, * item, + item
        m_bullet = re.match(r"^[-*+]\s+(.*)$", l_stripped)
        if m_bullet:
            clean_text = m_bullet.group(1).strip()
            return ParsedListItem(
                item_type="bullet",
                level=level,
                text=clean_text,
                is_checked=False,
                prefix=m_bullet.group(0)[:m_bullet.group(0).index(clean_text)]
            )

        # 3. Ordered List: 1. item, 2) item
        m_ordered = re.match(r"^(\d+)[\.\)]\s+(.*)$", l_stripped)
        if m_ordered:
            raw_num = int(m_ordered.group(1))
            clean_text = m_ordered.group(2).strip()
            return ParsedListItem(
                item_type="ordered",
                level=level,
                text=clean_text,
                raw_number=raw_num,
                prefix=m_ordered.group(0)[:m_ordered.group(0).index(clean_text)]
            )

        return None

    def get_num_id_for_item(self, item: ParsedListItem) -> Optional[int]:
        """
        Determines the appropriate numId for an item, applying restart rules
        or smart continuation if sub-items were interleaved.
        """
        if item.item_type == "ordered":
            if self.active_ordered_num_id is None:
                # Brand new list -> allocate fresh numId restarting at 1 (for all levels)
                self.active_ordered_num_id = self.allocate_ordered_num_id()
            elif item.raw_number == 1 and self.last_item_type != "ordered":
                # Explicit user '1.' after an interruption (bullet, note, heading) -> restart
                self.active_ordered_num_id = self.allocate_ordered_num_id()
            self.last_item_type = "ordered"
            self.last_level = item.level
            return self.active_ordered_num_id

        elif item.item_type == "bullet":
            self.last_item_type = "bullet"
            self.last_level = item.level
            return self.get_bullet_num_id()

        return None


def ensure_numbering_definitions(doc: DocumentType) -> None:
    """
    Registers standard ECMA-376 abstractNum definitions in numbering.xml:
    - ABSTRACT_NUM_DECIMAL_ID (2001): 3 levels (1., a., i.)
    - ABSTRACT_NUM_BULLET_ID (2002): 3 levels (•, ○, ■)
    Idempotent: skips if already defined.
    """
    n_part = getattr(doc.part, "numbering_part", None)
    if not n_part:
        return

    n_el = n_part.element
    existing_abs = {
        int(el.get(qn("w:abstractNumId")))
        for el in n_el.findall(qn("w:abstractNum"))
        if el.get(qn("w:abstractNumId")) is not None
    }

    # 1. Decimal Multi-Level Abstract Definition
    if ABSTRACT_NUM_DECIMAL_ID not in existing_abs:
        xml_dec = (
            f'<w:abstractNum {NS_W} w:abstractNumId="{ABSTRACT_NUM_DECIMAL_ID}">'
            f'  <w:multiLevelType w:val="multilevel"/>'
            f'  <w:lvl w:ilvl="0">'
            f'    <w:start w:val="1"/>'
            f'    <w:numFmt w:val="decimal"/>'
            f'    <w:lvlText w:val="%1."/>'
            f'    <w:lvlJc w:val="left"/>'
            f'    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
            f'  </w:lvl>'
            f'  <w:lvl w:ilvl="1">'
            f'    <w:start w:val="1"/>'
            f'    <w:numFmt w:val="lowerLetter"/>'
            f'    <w:lvlText w:val="%2."/>'
            f'    <w:lvlJc w:val="left"/>'
            f'    <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
            f'  </w:lvl>'
            f'  <w:lvl w:ilvl="2">'
            f'    <w:start w:val="1"/>'
            f'    <w:numFmt w:val="lowerRoman"/>'
            f'    <w:lvlText w:val="%3."/>'
            f'    <w:lvlJc w:val="left"/>'
            f'    <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
            f'  </w:lvl>'
            f'</w:abstractNum>'
        )
        n_el.insert(0, parse_xml(xml_dec))

    # 2. Bullet Multi-Level Abstract Definition
    if ABSTRACT_NUM_BULLET_ID not in existing_abs:
        xml_bullet = (
            f'<w:abstractNum {NS_W} w:abstractNumId="{ABSTRACT_NUM_BULLET_ID}">'
            f'  <w:multiLevelType w:val="multilevel"/>'
            f'  <w:lvl w:ilvl="0">'
            f'    <w:numFmt w:val="bullet"/>'
            f'    <w:lvlText w:val="&#x2022;"/>'
            f'    <w:lvlJc w:val="left"/>'
            f'    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
            f'    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'
            f'  </w:lvl>'
            f'  <w:lvl w:ilvl="1">'
            f'    <w:numFmt w:val="bullet"/>'
            f'    <w:lvlText w:val="&#x25CB;"/>'
            f'    <w:lvlJc w:val="left"/>'
            f'    <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>'
            f'    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'
            f'  </w:lvl>'
            f'  <w:lvl w:ilvl="2">'
            f'    <w:numFmt w:val="bullet"/>'
            f'    <w:lvlText w:val="&#x25A0;"/>'
            f'    <w:lvlJc w:val="left"/>'
            f'    <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>'
            f'    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'
            f'  </w:lvl>'
            f'</w:abstractNum>'
        )
        n_el.insert(0, parse_xml(xml_bullet))


def apply_list_formatting(
    paragraph: docx.text.paragraph.Paragraph,
    item_type: str,
    level: int = 0,
    num_id: Optional[int] = None,
    is_checked: bool = False
) -> None:
    """
    Applies OpenXML <w:numPr> and proportional indentation to a paragraph.
    Handles checkbox rendering using clean ballot box glyphs.
    """
    level = max(0, min(2, level))
    left_twips = (level + 1) * INDENT_LEVEL_STEP_TWIPS
    hanging_twips = INDENT_HANGING_TWIPS

    pPr = paragraph._p.get_or_add_pPr()

    # Clear existing numPr if any
    for existing in pPr.findall(qn("w:numPr")):
        pPr.remove(existing)

    if item_type in ("ordered", "bullet") and num_id is not None:
        numPr = parse_xml(
            f'<w:numPr {NS_W}>'
            f'  <w:ilvl w:val="{level}"/>'
            f'  <w:numId w:val="{num_id}"/>'
            f'</w:numPr>'
        )
        pPr.append(numPr)
    elif item_type == "checkbox":
        # Checkboxes use explicit indented symbol
        box_glyph = "☑  " if is_checked else "☐  "
        r_box = paragraph.insert_paragraph_before() if False else None
        # Insert check box symbol at the very beginning of the paragraph
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        if paragraph.runs:
            # Prepend to first run or insert as first run
            first_run = paragraph.runs[0]
            first_run.text = box_glyph + first_run.text
            first_run.font.bold = True
        else:
            r = paragraph.add_run(box_glyph)
            r.bold = True

    # Indentation (<w:ind w:left="X" w:hanging="Y"/>)
    for existing in pPr.findall(qn("w:ind")):
        pPr.remove(existing)
    ind = parse_xml(f'<w:ind {NS_W} w:left="{left_twips}" w:hanging="{hanging_twips}"/>')
    pPr.append(ind)

    # Paragraph Spacing Invariants
    paragraph.paragraph_format.space_before = Pt(1.5)
    paragraph.paragraph_format.space_after = Pt(1.5)
    paragraph.paragraph_format.line_spacing = 1.15
