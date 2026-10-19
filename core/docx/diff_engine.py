"""
core/docx/diff_engine.py — Enterprise Legal Redline Diff Engine & AI Track Changes.

Compares two Word (.docx) documents and produces a publication-grade redline diff:
  - Generates valid OpenXML <w:ins> and <w:del> tags with <w:delText> (ERR_DOCX_008).
  - Preserves unique w:id, w:author, and ISO-8601 UTC w:date attributes.
  - Injects AI Reviewer side margin comments into word/comments.xml.
  - Enforces The Last Paragraph Rule (<w:tc> ends with <w:p>) and cantSplit in tables.
"""

from __future__ import annotations

import copy
import datetime
import difflib
import io
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    import docx
    from docx import Document
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.shared import Pt, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)


class TrackChangeIdGenerator:
    """Generates sequential, positive unique w:id numbers for OpenXML track changes."""

    def __init__(self, start_id: int = 1) -> None:
        self._current_id = max(1, start_id)

    def next_id(self) -> int:
        cur = self._current_id
        self._current_id += 1
        return cur


def tokenize_text(text: str) -> List[str]:
    """
    Split text into words, whitespace, and punctuation tokens so that
    ''.join(tokenize_text(text)) == text exactly.
    """
    if not text:
        return []
    return re.findall(r"\s+|[^\w\s]+|\w+", text, re.UNICODE)


def diff_token_sequences(
    tokens_a: List[str], tokens_b: List[str]
) -> List[Tuple[str, str]]:
    """
    Compute diff opcodes between two token sequences using SequenceMatcher.
    Returns list of (opcode, text_chunk):
      'equal', 'insert', 'delete'
    """
    matcher = difflib.SequenceMatcher(None, tokens_a, tokens_b)
    diff_chunks: List[Tuple[str, str]] = []

    for tag, alo, ahi, blo, bhi in matcher.get_opcodes():
        if tag == "equal":
            diff_chunks.append(("equal", "".join(tokens_a[alo:ahi])))
        elif tag == "delete":
            diff_chunks.append(("delete", "".join(tokens_a[alo:ahi])))
        elif tag == "insert":
            diff_chunks.append(("insert", "".join(tokens_b[blo:bhi])))
        elif tag == "replace":
            # Treat replace as delete original followed by insert revised
            diff_chunks.append(("delete", "".join(tokens_a[alo:ahi])))
            diff_chunks.append(("insert", "".join(tokens_b[blo:bhi])))

    return diff_chunks


def create_ins_element(
    text: str,
    change_id: int,
    author: str,
    date_str: str,
    font_name: Optional[str] = None,
    font_size_pt: Optional[float] = None,
) -> Any:
    """
    Build an OpenXML <w:ins> element containing inserted text (ERR_DOCX_008 compliant).
    """
    ins_xml = f'<w:ins {nsdecls("w")} w:id="{change_id}" w:author="{author}" w:date="{date_str}">'
    ins_xml += '<w:r>'
    if font_name or font_size_pt:
        ins_xml += '<w:rPr>'
        if font_name:
            ins_xml += f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}"/>'
        if font_size_pt:
            half_pts = int(round(font_size_pt * 2))
            ins_xml += f'<w:sz w:val="{half_pts}"/>'
        ins_xml += '</w:rPr>'
    
    # Escape XML special characters
    safe_text = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    ins_xml += f'<w:t xml:space="preserve">{safe_text}</w:t>'
    ins_xml += '</w:r></w:ins>'
    return parse_xml(ins_xml)


def create_del_element(
    text: str,
    change_id: int,
    author: str,
    date_str: str,
    font_name: Optional[str] = None,
    font_size_pt: Optional[float] = None,
) -> Any:
    """
    Build an OpenXML <w:del> element containing deleted text using <w:delText> (ERR_DOCX_008 compliant).
    """
    del_xml = f'<w:del {nsdecls("w")} w:id="{change_id}" w:author="{author}" w:date="{date_str}">'
    del_xml += '<w:r>'
    if font_name or font_size_pt:
        del_xml += '<w:rPr>'
        if font_name:
            del_xml += f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}"/>'
        if font_size_pt:
            half_pts = int(round(font_size_pt * 2))
            del_xml += f'<w:sz w:val="{half_pts}"/>'
        del_xml += '</w:rPr>'

    safe_text = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    del_xml += f'<w:delText xml:space="preserve">{safe_text}</w:delText>'
    del_xml += '</w:r></w:del>'
    return parse_xml(del_xml)


def create_equal_run(
    text: str,
    font_name: Optional[str] = None,
    font_size_pt: Optional[float] = None,
) -> Any:
    """Build a normal un-tracked run element."""
    r_xml = f'<w:r {nsdecls("w")}>'
    if font_name or font_size_pt:
        r_xml += '<w:rPr>'
        if font_name:
            r_xml += f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}"/>'
        if font_size_pt:
            half_pts = int(round(font_size_pt * 2))
            r_xml += f'<w:sz w:val="{half_pts}"/>'
        r_xml += '</w:rPr>'

    safe_text = (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    r_xml += f'<w:t xml:space="preserve">{safe_text}</w:t></w:r>'
    return parse_xml(r_xml)


def create_comment_marker_start(comment_id: int) -> Any:
    return parse_xml(f'<w:commentRangeStart {nsdecls("w")} w:id="{comment_id}"/>')


def create_comment_marker_end(comment_id: int) -> Any:
    return parse_xml(f'<w:commentRangeEnd {nsdecls("w")} w:id="{comment_id}"/>')


def create_comment_reference(comment_id: int) -> Any:
    return parse_xml(
        f'<w:r {nsdecls("w")}><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{comment_id}"/></w:r>'
    )


class RedlineDiffEngine:
    """
    Enterprise Legal Redline Diff Engine.
    Executes bidirectional structural diff between two DOCX documents.
    """

    def __init__(
        self,
        author: str = "AI Reviewer",
        date_str: Optional[str] = None,
        start_id: int = 1,
    ) -> None:
        self.author = author
        self.date_str = date_str or datetime.datetime.now(
            datetime.timezone.utc
        ).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.id_gen = TrackChangeIdGenerator(start_id=start_id)
        self.comments: List[Dict[str, Any]] = []

    def add_comment(
        self,
        text: str,
        author: Optional[str] = None,
        date_str: Optional[str] = None,
    ) -> int:
        """Register a comment to be packaged into word/comments.xml."""
        cid = len(self.comments) + 1
        self.comments.append(
            {
                "id": cid,
                "text": text,
                "author": author or self.author,
                "date": date_str or self.date_str,
                "initials": (author or self.author)[:2].upper(),
            }
        )
        return cid

    def diff_paragraph_pair(
        self,
        target_p: Any,
        text_a: str,
        text_b: str,
        comment_text: Optional[str] = None,
    ) -> Dict[str, int]:
        """
        Diff text_a and text_b, populating target_p._p with <w:ins>, <w:del>, and normal runs.
        """
        tokens_a = tokenize_text(text_a)
        tokens_b = tokenize_text(text_b)
        chunks = diff_token_sequences(tokens_a, tokens_b)

        stats = {"insertions": 0, "deletions": 0, "equals": 0}

        comment_id = None
        if comment_text:
            comment_id = self.add_comment(comment_text)
            target_p._p.append(create_comment_marker_start(comment_id))

        for tag, chunk_text in chunks:
            if not chunk_text:
                continue

            if tag == "equal":
                target_p._p.append(create_equal_run(chunk_text))
                stats["equals"] += 1
            elif tag == "insert":
                cid = self.id_gen.next_id()
                target_p._p.append(
                    create_ins_element(
                        chunk_text, cid, self.author, self.date_str
                    )
                )
                stats["insertions"] += 1
            elif tag == "delete":
                cid = self.id_gen.next_id()
                target_p._p.append(
                    create_del_element(
                        chunk_text, cid, self.author, self.date_str
                    )
                )
                stats["deletions"] += 1

        if comment_id is not None:
            target_p._p.append(create_comment_marker_end(comment_id))
            target_p._p.append(create_comment_reference(comment_id))

        return stats

    def compare_documents(
        self,
        doc_a_path: str,
        doc_b_path: str,
        output_path: str,
        user_comments: Optional[List[Dict[str, Any]]] = None,
    ) -> Dict[str, Any]:
        """
        Compare two Word documents and write a track-changed redline document to output_path.
        """
        if not HAS_DOCX:
            raise RuntimeError("python-docx is required for RedlineDiffEngine")

        if not os.path.exists(doc_a_path):
            raise FileNotFoundError(f"Original document not found: {doc_a_path}")
        if not os.path.exists(doc_b_path):
            raise FileNotFoundError(f"Revised document not found: {doc_b_path}")

        doc_a = Document(doc_a_path)
        doc_b = Document(doc_b_path)

        out_doc = Document()
        if len(out_doc.paragraphs) > 0:
            p_elem = out_doc.paragraphs[0]._p
            p_parent = p_elem.getparent()
            if p_parent is not None:
                p_parent.remove(p_elem)

        total_stats = {
            "insertions_count": 0,
            "deletions_count": 0,
            "equals_count": 0,
            "comments_count": 0,
            "paragraphs_compared": 0,
            "tables_compared": 0,
            "output_path": output_path,
        }

        # 1. Compare Paragraphs
        paras_a = [p.text for p in doc_a.paragraphs]
        paras_b = [p.text for p in doc_b.paragraphs]

        p_matcher = difflib.SequenceMatcher(None, paras_a, paras_b)
        comment_idx = 0

        for tag, alo, ahi, blo, bhi in p_matcher.get_opcodes():
            if tag == "equal":
                for idx in range(alo, ahi):
                    new_p = out_doc.add_paragraph()
                    self.diff_paragraph_pair(new_p, paras_a[idx], paras_a[idx])
                    total_stats["equals_count"] += 1
                    total_stats["paragraphs_compared"] += 1

            elif tag == "delete":
                for idx in range(alo, ahi):
                    new_p = out_doc.add_paragraph()
                    cid = self.id_gen.next_id()
                    new_p._p.append(
                        create_del_element(
                            paras_a[idx], cid, self.author, self.date_str
                        )
                    )
                    total_stats["deletions_count"] += 1
                    total_stats["paragraphs_compared"] += 1

            elif tag == "insert":
                for idx in range(blo, bhi):
                    new_p = out_doc.add_paragraph()
                    cid = self.id_gen.next_id()
                    new_p._p.append(
                        create_ins_element(
                            paras_b[idx], cid, self.author, self.date_str
                        )
                    )
                    total_stats["insertions_count"] += 1
                    total_stats["paragraphs_compared"] += 1

            elif tag == "replace":
                len_a = ahi - alo
                len_b = bhi - blo
                common_len = min(len_a, len_b)

                for k in range(common_len):
                    new_p = out_doc.add_paragraph()
                    cmt = None
                    if user_comments and comment_idx < len(user_comments):
                        cmt = user_comments[comment_idx].get("text")
                        comment_idx += 1
                    p_stats = self.diff_paragraph_pair(
                        new_p, paras_a[alo + k], paras_b[blo + k], comment_text=cmt
                    )
                    total_stats["insertions_count"] += p_stats["insertions"]
                    total_stats["deletions_count"] += p_stats["deletions"]
                    total_stats["equals_count"] += p_stats["equals"]
                    total_stats["paragraphs_compared"] += 1

                if len_a > common_len:
                    for k in range(common_len, len_a):
                        new_p = out_doc.add_paragraph()
                        cid = self.id_gen.next_id()
                        new_p._p.append(
                            create_del_element(
                                paras_a[alo + k], cid, self.author, self.date_str
                            )
                        )
                        total_stats["deletions_count"] += 1
                        total_stats["paragraphs_compared"] += 1
                elif len_b > common_len:
                    for k in range(common_len, len_b):
                        new_p = out_doc.add_paragraph()
                        cid = self.id_gen.next_id()
                        new_p._p.append(
                            create_ins_element(
                                paras_b[blo + k], cid, self.author, self.date_str
                            )
                        )
                        total_stats["insertions_count"] += 1
                        total_stats["paragraphs_compared"] += 1

        # 2. Compare Tables
        tables_b = doc_b.tables
        tables_a = doc_a.tables

        for t_idx, tbl_b in enumerate(tables_b):
            tbl_a = tables_a[t_idx] if t_idx < len(tables_a) else None
            num_rows = len(tbl_b.rows)
            num_cols = len(tbl_b.columns) if num_rows > 0 else 0

            if num_rows == 0 or num_cols == 0:
                continue

            new_tbl = out_doc.add_table(rows=num_rows, cols=num_cols)
            new_tbl.style = "Table Grid"

            for r_idx, row_b in enumerate(tbl_b.rows):
                trPr = new_tbl.rows[r_idx]._tr.get_or_add_trPr()
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                if r_idx == 0:
                    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

                row_a = None
                if tbl_a and r_idx < len(tbl_a.rows):
                    row_a = tbl_a.rows[r_idx]

                for c_idx, cell_b in enumerate(row_b.cells):
                    cell_a = (
                        row_a.cells[c_idx]
                        if (row_a and c_idx < len(row_a.cells))
                        else None
                    )
                    text_a = cell_a.text if cell_a else ""
                    text_b = cell_b.text

                    target_cell = new_tbl.rows[r_idx].cells[c_idx]
                    p = target_cell.paragraphs[0]
                    p_stats = self.diff_paragraph_pair(p, text_a, text_b)
                    total_stats["insertions_count"] += p_stats["insertions"]
                    total_stats["deletions_count"] += p_stats["deletions"]

                    # Enforce The Last Paragraph Rule (ERR_DOCX_001)
                    if len(target_cell._tc.findall(qn("w:p"))) == 0:
                        target_cell._tc.append(
                            parse_xml(f'<w:p {nsdecls("w")}/>')
                        )

            total_stats["tables_compared"] += 1

        # Save base document
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        out_doc.save(output_path)

        # 3. Inject comments into word/comments.xml if any comments were created
        if self.comments:
            _inject_comments_to_package(output_path, self.comments)
            total_stats["comments_count"] = len(self.comments)

        return total_stats


def _inject_comments_to_package(
    docx_path: str, comments: List[Dict[str, Any]]
) -> None:
    """
    Directly injects word/comments.xml, updates [Content_Types].xml and
    word/_rels/document.xml.rels in the DOCX zip package.
    """
    ET.register_namespace("w", W_NS)
    ET.register_namespace("r", R_NS)

    root = ET.Element(
        f"{{{W_NS}}}comments",
        attrib={f"{{{R_NS}}}id": "rIdComments"},
    )

    for c in comments:
        c_elem = ET.SubElement(
            root,
            f"{{{W_NS}}}comment",
            attrib={
                f"{{{W_NS}}}id": str(c["id"]),
                f"{{{W_NS}}}author": str(c["author"]),
                f"{{{W_NS}}}date": str(c["date"]),
                f"{{{W_NS}}}initials": str(c.get("initials", "AI")),
            },
        )
        p = ET.SubElement(c_elem, f"{{{W_NS}}}p")
        pPr = ET.SubElement(p, f"{{{W_NS}}}pPr")
        pStyle = ET.SubElement(
            pPr, f"{{{W_NS}}}pStyle", attrib={f"{{{W_NS}}}val": "CommentText"}
        )
        r_ref = ET.SubElement(p, f"{{{W_NS}}}r")
        rPr_ref = ET.SubElement(r_ref, f"{{{W_NS}}}rPr")
        rStyle_ref = ET.SubElement(
            rPr_ref,
            f"{{{W_NS}}}rStyle",
            attrib={f"{{{W_NS}}}val": "CommentReference"},
        )
        ET.SubElement(r_ref, f"{{{W_NS}}}annotationRef")

        r_text = ET.SubElement(p, f"{{{W_NS}}}r")
        t = ET.SubElement(r_text, f"{{{W_NS}}}t")
        t.text = c["text"]

    comments_xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    temp_zip = docx_path + ".tmp.zip"
    with zipfile.ZipFile(docx_path, "r") as zin:
        with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)

                if item.filename == "[Content_Types].xml":
                    ct_root = ET.fromstring(content)
                    has_override = False
                    for child in ct_root:
                        if child.attrib.get("PartName") == "/word/comments.xml":
                            has_override = True
                            break
                    if not has_override:
                        ET.SubElement(
                            ct_root,
                            "Override",
                            attrib={
                                "PartName": "/word/comments.xml",
                                "ContentType": COMMENTS_CONTENT_TYPE,
                            },
                        )
                    content = ET.tostring(
                        ct_root, encoding="utf-8", xml_declaration=True
                    )

                elif item.filename == "word/_rels/document.xml.rels":
                    rels_root = ET.fromstring(content)
                    has_comments_rel = False
                    for child in rels_root:
                        if (
                            child.attrib.get("Type") == COMMENTS_REL_TYPE
                            or child.attrib.get("Target") == "comments.xml"
                        ):
                            has_comments_rel = True
                            break
                    if not has_comments_rel:
                        max_rid = 10
                        for child in rels_root:
                            rid_str = child.attrib.get("Id", "")
                            if rid_str.startswith("rId"):
                                try:
                                    num = int(rid_str[3:])
                                    if num > max_rid:
                                        max_rid = num
                                except ValueError:
                                    pass
                        new_rid = f"rId{max_rid + 1}"
                        ET.SubElement(
                            rels_root,
                            "Relationship",
                            attrib={
                                "Id": new_rid,
                                "Type": COMMENTS_REL_TYPE,
                                "Target": "comments.xml",
                            },
                        )
                    content = ET.tostring(
                        rels_root, encoding="utf-8", xml_declaration=True
                    )

                zout.writestr(item, content)

            zout.writestr("word/comments.xml", comments_xml_bytes)

    os.replace(temp_zip, docx_path)


def compare_docx_documents(
    doc_a_path: str,
    doc_b_path: str,
    output_docx_path: str,
    author: str = "AI Reviewer",
    date_str: Optional[str] = None,
    comments: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Convenience function for Redline Diff comparison.
    """
    engine = RedlineDiffEngine(author=author, date_str=date_str)
    return engine.compare_documents(
        doc_a_path=doc_a_path,
        doc_b_path=doc_b_path,
        output_path=output_docx_path,
        user_comments=comments,
    )
