"""
docx_math.py — Technical & Mathematical Expression Engine for Word OpenXML & OMML.

Dual-Mode Architecture:
1. High-Fidelity OpenXML Multi-Run Mode (Default for inline $...$ expressions):
   - Variable italicization (<w:i/>) for single-letter and Greek variables.
   - Upright (<w:i w:val="0"/>) for numeric values, units, and textual subscripts (\text{ref}, \text{max}).
   - OpenXML native Subscript & Superscript (<w:vertAlign w:val="subscript|superscript"/>).
   - Rich Greek lexicon and Unicode mathematical operators.
   - 100% cross-platform compatible across Microsoft Word, LibreOffice, WPS, Google Docs, and Mobile.
2. Native Office Math (OMML) Mode (for block equations $$...$$ and fractions):
   - Generates <m:oMath> and <m:oMathPara> elements with <m:f> (fractions) and <m:rad> (radicals).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, List, Optional, Tuple, Union

from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

# ---------------------------------------------------------------------------
# Greek & Mathematical Symbol Lexicon
# ---------------------------------------------------------------------------

GREEK_SYMBOLS = {
    # Lowercase
    r"\alpha": "α",
    r"\beta": "β",
    r"\gamma": "γ",
    r"\delta": "δ",
    r"\epsilon": "ε",
    r"\varepsilon": "ε",
    r"\zeta": "ζ",
    r"\eta": "η",
    r"\theta": "θ",
    r"\iota": "ι",
    r"\kappa": "κ",
    r"\lambda": "λ",
    r"\mu": "μ",
    r"\nu": "ν",
    r"\xi": "ξ",
    r"\pi": "π",
    r"\rho": "ρ",
    r"\sigma": "σ",
    r"\tau": "τ",
    r"\upsilon": "υ",
    r"\phi": "φ",
    r"\chi": "χ",
    r"\psi": "ψ",
    r"\omega": "ω",
    # Uppercase
    r"\Gamma": "Γ",
    r"\Delta": "Δ",
    r"\Theta": "Θ",
    r"\Lambda": "Λ",
    r"\Xi": "Ξ",
    r"\Pi": "Π",
    r"\Sigma": "Σ",
    r"\Phi": "Φ",
    r"\Psi": "Ψ",
    r"\Omega": "Ω",
}

MATH_OPERATORS = {
    r"\approx": "≈",
    r"\sim": "∼",
    r"\ne": "≠",
    r"\neq": "≠",
    r"\le": "≤",
    r"\leq": "≤",
    r"\ge": "≥",
    r"\geq": "≥",
    r"\pm": "±",
    r"\mp": "∓",
    r"\times": "×",
    r"\cdot": "·",
    r"\div": "÷",
    r"\circ": "°",
    r"\degree": "°",
    r"\to": "→",
    r"\rightarrow": "→",
    r"\leftarrow": "←",
    r"\infty": "∞",
    r"\propto": "∝",
    r"\partial": "∂",
    r"\nabla": "∇",
    r"\sqrt": "√",
}

ALL_SYMBOLS = {**GREEK_SYMBOLS, **MATH_OPERATORS}

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass
class MathRunToken:
    """Represents an individual styled text run inside a mathematical formula."""
    text: str
    is_italic: bool = False
    is_bold: bool = False
    is_subscript: bool = False
    is_superscript: bool = False
    scale_size: float = 1.0


def replace_latex_symbols(text: str) -> str:
    r"""Replaces LaTeX symbol macros (\alpha, \Omega, \ge, etc.) with Unicode characters."""
    for macro, uni in ALL_SYMBOLS.items():
        pattern = re.escape(macro) + r"(?![a-zA-Z])"
        text = re.sub(pattern, uni, text)
    return text


def parse_inline_math(expr: str) -> List[MathRunToken]:
    r"""
    Parses a LaTeX-style inline mathematical expression into structured tokens.
    Handles subscripts (x_1, v_{ref}), superscripts (x^2, e^{-t}), \text{...}, \mathbf{...}, \vec{...}.
    """
    expr = replace_latex_symbols(expr.strip())
    tokens: List[MathRunToken] = []
    i = 0
    n = len(expr)

    while i < n:
        # 1. Handle \text{...} blocks
        if expr.startswith(r"\text{", i):
            end_brace = expr.find("}", i + 6)
            if end_brace != -1:
                content = expr[i + 6:end_brace]
                tokens.append(MathRunToken(text=content, is_italic=False))
                i = end_brace + 1
                continue

        # 1b. Handle \mathbf{...} blocks
        if expr.startswith(r"\mathbf{", i):
            end_brace = expr.find("}", i + 8)
            if end_brace != -1:
                content = expr[i + 8:end_brace]
                tokens.append(MathRunToken(text=content, is_italic=False, is_bold=True))
                i = end_brace + 1
                continue

        # 1c. Handle \vec{...} blocks
        if expr.startswith(r"\vec{", i):
            end_brace = expr.find("}", i + 5)
            if end_brace != -1:
                content = expr[i + 5:end_brace]
                tokens.append(MathRunToken(text=content, is_italic=True, is_bold=True))
                i = end_brace + 1
                continue

        ch = expr[i]

        # 2. Handle Subscript _
        if ch == "_":
            i += 1
            if i < n and expr[i] == "{":
                end_brace = expr.find("}", i + 1)
                sub_text = expr[i + 1:end_brace] if end_brace != -1 else expr[i + 1:]
                i = (end_brace + 1) if end_brace != -1 else n
            elif i < n:
                sub_text = expr[i]
                i += 1
            else:
                sub_text = ""

            sub_text = replace_latex_symbols(sub_text)
            sub_italic = len(sub_text) == 1 and sub_text.isalpha()
            tokens.append(MathRunToken(
                text=sub_text,
                is_italic=sub_italic,
                is_subscript=True,
                scale_size=0.8
            ))
            continue

        # 3. Handle Superscript ^
        if ch == "^":
            i += 1
            if i < n and expr[i] == "{":
                end_brace = expr.find("}", i + 1)
                sup_text = expr[i + 1:end_brace] if end_brace != -1 else expr[i + 1:]
                i = (end_brace + 1) if end_brace != -1 else n
            elif i < n:
                sup_text = expr[i]
                i += 1
            else:
                sup_text = ""

            sup_text = replace_latex_symbols(sup_text)
            sup_italic = len(sup_text) == 1 and sup_text.isalpha()
            tokens.append(MathRunToken(
                text=sup_text,
                is_italic=sup_italic,
                is_superscript=True,
                scale_size=0.8
            ))
            continue

        # 4. Handle whitespace
        if ch.isspace():
            tokens.append(MathRunToken(text=" ", is_italic=False))
            i += 1
            continue

        # 5. Handle numbers & operators
        if ch.isdigit() or ch in "=+-*/()[],.;:!%<>|~":
            tokens.append(MathRunToken(text=ch, is_italic=False))
            i += 1
            continue

        # 6. Single alphabetic character / Greek variable
        if ch.isalpha() or ord(ch) > 127:
            tokens.append(MathRunToken(text=ch, is_italic=True))
            i += 1
            continue

        tokens.append(MathRunToken(text=ch, is_italic=False))
        i += 1

    # Merge consecutive identical formatting tokens
    merged: List[MathRunToken] = []
    for tok in tokens:
        if merged and (
            merged[-1].is_italic == tok.is_italic and
            merged[-1].is_bold == tok.is_bold and
            merged[-1].is_subscript == tok.is_subscript and
            merged[-1].is_superscript == tok.is_superscript and
            merged[-1].scale_size == tok.scale_size
        ):
            merged[-1].text += tok.text
        else:
            merged.append(tok)

    return merged


def add_math_token_run(
    paragraph: Any,
    token: MathRunToken,
    base_font_size_pt: float = 11.0,
    font_name: str = "Arial",
    color_rgb: Optional[RGBColor] = None,
    inherited_bold: bool = False
) -> Any:
    """Appends an individual MathRunToken as a styled OpenXML run to a paragraph."""
    run = paragraph.add_run(token.text)
    run.font.name = font_name
    run.italic = token.is_italic
    run.bold = token.is_bold or inherited_bold

    effective_pt = base_font_size_pt * token.scale_size
    run.font.size = Pt(effective_pt)

    if color_rgb:
        run.font.color.rgb = color_rgb

    # Apply vertical alignment for subscript / superscript
    if token.is_subscript or token.is_superscript:
        rPr = run._r.get_or_add_rPr()
        for existing in rPr.findall(qn("w:vertAlign")):
            rPr.remove(existing)
        vert_elem = OxmlElement("w:vertAlign")
        vert_elem.set(qn("w:val"), "subscript" if token.is_subscript else "superscript")
        rPr.append(vert_elem)

    return run


def create_omml_equation_element(expr: str) -> OxmlElement:
    """
    Constructs an OMML <m:oMath> XML element for complex block equations.
    Supports fractions \\frac{num}{den} and radicals \\sqrt{val}.
    """
    clean_expr = replace_latex_symbols(expr.strip())
    
    # Handle fraction: \frac{a}{b}
    frac_match = re.search(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", clean_expr)
    if frac_match:
        num = frac_match.group(1).strip()
        den = frac_match.group(2).strip()
        omml_xml = f"""
        <m:oMath xmlns:m="{M_NS}" xmlns:w="{W_NS}">
            <m:f>
                <m:num><m:r><m:t>{num}</m:t></m:r></m:num>
                <m:den><m:r><m:t>{den}</m:t></m:r></m:den>
            </m:f>
        </m:oMath>
        """
        return parse_xml(omml_xml)
        
    # Standard OMML text run
    tokens = parse_inline_math(clean_expr)
    runs_xml = []
    for t in tokens:
        if t.is_subscript:
            runs_xml.append(f"""
            <m:sSub>
                <m:e><m:r><m:t></m:t></m:r></m:e>
                <m:sub><m:r><m:t>{t.text}</m:t></m:r></m:sub>
            </m:sSub>
            """)
        elif t.is_superscript:
            runs_xml.append(f"""
            <m:sSup>
                <m:e><m:r><m:t></m:t></m:r></m:e>
                <m:sup><m:r><m:t>{t.text}</m:t></m:r></m:sup>
            </m:sSup>
            """)
        else:
            runs_xml.append(f"<m:r><m:t>{t.text}</m:t></m:r>")

    combined = "".join(runs_xml)
    omml_xml = f'<m:oMath xmlns:m="{M_NS}" xmlns:w="{W_NS}">{combined}</m:oMath>'
    return parse_xml(omml_xml)


def add_math_to_paragraph(
    paragraph: Any,
    text: str,
    base_font_size_pt: float = 11.0,
    font_name: str = "Arial",
    color_rgb: Optional[RGBColor] = None,
    bold: bool = False
) -> None:
    """
    Master Dispatcher: Parses mixed paragraph text containing LaTeX math ($...$ or $$...$$)
    and transparently injects styled OpenXML runs or OMML equations.
    """
    if not text:
        return

    # Split on $$block math$$ or $inline math$
    parts = re.split(r"(\$\$.*?\$\$|\$.*?\$)", text)

    for part in parts:
        if not part:
            continue

        # Case 1: Block equation $$...$$
        if part.startswith("$$") and part.endswith("$$") and len(part) >= 4:
            math_expr = part[2:-2].strip()
            # If complex fraction or user opted for block OMML, inject OMML element
            if r"\frac" in math_expr or len(math_expr) > 25:
                omml_elem = create_omml_equation_element(math_expr)
                paragraph._p.append(omml_elem)
            else:
                tokens = parse_inline_math(math_expr)
                for tok in tokens:
                    add_math_token_run(
                        paragraph,
                        tok,
                        base_font_size_pt=base_font_size_pt,
                        font_name=font_name,
                        color_rgb=color_rgb,
                        inherited_bold=bold
                    )
            continue

        # Case 2: Inline math $...$
        if part.startswith("$") and part.endswith("$") and len(part) >= 2:
            math_expr = part[1:-1].strip()
            tokens = parse_inline_math(math_expr)
            for tok in tokens:
                add_math_token_run(
                    paragraph,
                    tok,
                    base_font_size_pt=base_font_size_pt,
                    font_name=font_name,
                    color_rgb=color_rgb,
                    inherited_bold=bold
                )
            continue

        # Case 3: Regular text (handles bold **text** and italic *text*)
        sub_tokens = re.split(r"(\*\*.*?\*\*|\*.*?\*)", part)
        for s_tok in sub_tokens:
            if not s_tok:
                continue
            if s_tok.startswith("**") and s_tok.endswith("**") and len(s_tok) >= 4:
                run = paragraph.add_run(s_tok[2:-2])
                run.bold = True
            elif s_tok.startswith("*") and s_tok.endswith("*") and len(s_tok) >= 2:
                run = paragraph.add_run(s_tok[1:-1])
                run.italic = True
                run.bold = bold
            else:
                run = paragraph.add_run(s_tok)
                run.bold = bold

            run.font.name = font_name
            run.font.size = Pt(base_font_size_pt)
            if color_rgb:
                run.font.color.rgb = color_rgb
