"""
pptx_math.py — Technical & Mathematical Expression Engine for PowerPoint DrawingML.

Conforms to Enterprise Invariants:
- PPTX_INV_09: Run-Level Formatting Supremacy (all styles in a:rPr, zero defRPr locking)
- PPTX_INV_10: Default Pure Black Typography (#000000)
- PPTX_INV_12: Native DrawingML Math Subscript/Superscript (no raw a14:m schema hazards)
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, List, Optional, Tuple
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

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


@dataclass
class MathRunToken:
    """Represents an individual text run inside a mathematical formula."""
    text: str
    is_italic: bool = False
    is_bold: bool = False
    is_subscript: bool = False
    is_superscript: bool = False
    scale_size: float = 1.0  # Fraction of base font size (e.g. 0.75 for scripts)


def replace_latex_symbols(text: str) -> str:
    r"""Replaces LaTeX symbol macros (\alpha, \Omega, \ge, etc.) with Unicode characters."""
    for macro, uni in ALL_SYMBOLS.items():
        # Match macro when followed by non-alphanumeric or end of string
        pattern = re.escape(macro) + r"(?![a-zA-Z])"
        text = re.sub(pattern, uni, text)
    return text


def parse_inline_math(expr: str) -> List[MathRunToken]:
    """
    Parses a LaTeX-style inline mathematical expression into structured tokens.
    Handles:
    - Subscripts: x_1, v_{ref}, \theta_{cmd}
    - Superscripts: x^2, e^{-t/\tau}
    - Text blocks: \text{m/s}, \text{rad}
    - Variables vs numbers/operators
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
            # Acronyms (ref, max, min, avg) are upright; single variables (i, j, k) are italic
            sub_italic = len(sub_text) == 1 and sub_text.isalpha()
            tokens.append(MathRunToken(
                text=sub_text,
                is_italic=sub_italic,
                is_subscript=True,
                scale_size=0.75
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
                scale_size=0.75
            ))
            continue

        # 4. Handle whitespace
        if ch.isspace():
            tokens.append(MathRunToken(text=" ", is_italic=False))
            i += 1
            continue

        # 5. Handle numbers & punctuation / operators
        if ch.isdigit() or ch in "=+-*/()[],.;:!%<>|~":
            tokens.append(MathRunToken(text=ch, is_italic=False))
            i += 1
            continue

        # 6. Single alphabetic character / Greek variable
        if ch.isalpha() or ord(ch) > 127:
            tokens.append(MathRunToken(text=ch, is_italic=True))
            i += 1
            continue

        # Fallback character
        tokens.append(MathRunToken(text=ch, is_italic=False))
        i += 1

    # Merge consecutive identical tokens (e.g. adjacent numbers or spaces)
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


def create_drawingml_run(
    text: str,
    base_sz: int = 1200,
    scale_size: float = 1.0,
    is_italic: bool = False,
    is_bold: bool = False,
    is_subscript: bool = False,
    is_superscript: bool = False,
    color: str = "000000",
    typeface: str = "Arial",
) -> Any:
    """
    Creates an OpenXML <a:r> DrawingML element with fully specified <a:rPr>.
    Adheres strictly to PPTX_INV_09 (Run-Level Formatting) & PPTX_INV_12 (Native DrawingML Subscript).
    """
    final_sz = int(round(base_sz * scale_size))
    b_val = "1" if is_bold else "0"
    i_val = "1" if is_italic else "0"

    baseline_attr = ""
    if is_subscript:
        baseline_attr = ' baseline="-25000"'
    elif is_superscript:
        baseline_attr = ' baseline="30000"'

    run_xml = (
        f'<a:r {nsdecls("a")}>'
        f'<a:rPr sz="{final_sz}" b="{b_val}" i="{i_val}"{baseline_attr} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{typeface}"/>'
        f'</a:rPr>'
        f'<a:t>{_escape_xml(text)}</a:t>'
        f'</a:r>'
    )
    return parse_xml(run_xml)


def _escape_xml(text: str) -> str:
    """Safely escapes text content for DrawingML <a:t> elements."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def add_math_text_to_paragraph(
    paragraph: Any,
    text: str,
    base_sz: int = 1200,
    is_bold: bool = False,
    color: str = "000000",
    typeface: str = "Arial",
) -> None:
    """
    Appends mixed text containing inline LaTeX math ($...$) to an OpenXML paragraph.
    Normal text segments are added as standard black runs.
    Math segments are parsed and injected as native DrawingML runs (italic variables, subscripts, superscripts).
    """
    p_elem = paragraph._p

    # Split text into alternating normal and math segments
    # Pattern matches $...$ but excludes escaped \$
    parts = re.split(r'(?<!\\)\$(.*?)(?<!\\)\$', text)

    for idx, part in enumerate(parts):
        if not part:
            continue

        is_math = (idx % 2 == 1)

        if not is_math:
            # Normal text run
            cleaned_text = part.replace(r"\$", "$")
            run = create_drawingml_run(
                text=cleaned_text,
                base_sz=base_sz,
                scale_size=1.0,
                is_italic=False,
                is_bold=is_bold,
                is_subscript=False,
                is_superscript=False,
                color=color,
                typeface=typeface,
            )
            p_elem.append(run)
        else:
            # Inline math formula
            math_tokens = parse_inline_math(part)
            for tok in math_tokens:
                run = create_drawingml_run(
                    text=tok.text,
                    base_sz=base_sz,
                    scale_size=tok.scale_size,
                    is_italic=tok.is_italic,
                    is_bold=(is_bold or tok.is_bold),
                    is_subscript=tok.is_subscript,
                    is_superscript=tok.is_superscript,
                    color=color,
                    typeface=typeface,
                )
                p_elem.append(run)
