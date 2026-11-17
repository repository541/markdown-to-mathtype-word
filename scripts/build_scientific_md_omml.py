"""Build a Times New Roman 11-pt OMML intermediate from scientific Markdown.

The supported profile preserves research-paper prose, headings, bullets, and
basic HTML tables. Every supported TeX expression becomes native Office Math
(OMML), which is then consumed by the separate MathType 7 conversion step.
"""

from __future__ import annotations

import html
import json
import re
import sys
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree


FONT = "Times New Roman"
FONT_SIZE_PT = 11
BODY_LINE_SPACING = 1.5
MML2OMML_PATH = Path(r"D:\Program Files\Microsoft Office\Office16\MML2OMML.XSL")
DISPLAY_OPEN = r"\["
DISPLAY_CLOSE = r"\]"
INLINE_MATH = re.compile(r"\\\((.*?)\\\)", re.DOTALL)
TAG = re.compile(r"\\tag\{([^{}]*)\}")
HEADING = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
LIST_ITEM = re.compile(r"^[-*+]\s+(.+?)\s*$")
TABLE_BLOCK = re.compile(r"(?s)<table>(.*?)</table>")
ROW_BLOCK = re.compile(r"(?s)<tr>(.*?)</tr>")
CELL_BLOCK = re.compile(r"(?s)<t([hd])(?:\s+[^>]*)?>(.*?)</t\1>")
TABLE_CAPTION = re.compile(r"^Table\s+(S\d+(?:[a-z])?)\.", re.IGNORECASE)

# The final SI uses 11-pt text, so these wide scientific tables are given
# landscape pages and explicit Word grid widths rather than being squeezed into
# portrait equal-width columns.  The values are inches and stay inside a 10.2-
# inch landscape text area (Letter, 0.4-in side margins).
WIDE_TABLE_WIDTHS: dict[str, list[float]] = {
    "S3": [0.72] * 13,
    "S4": [1.25, 1.10, 1.10, 1.10, 1.10, 1.10, 1.10, 1.10, 1.10],
    "S5": [0.65, 1.17, 1.17, 1.17, 1.17, 1.17, 1.17, 1.17, 1.17],
    "S7": [0.65, 1.20, 1.10, 1.15, 0.95, 0.95, 1.20, 1.30, 0.90],
    "S10": [0.45, 0.70, 0.85, 0.85, 0.75, 0.85, 0.85, 0.75, 1.20, 0.80, 1.05],
    "S11": [1.45, 1.55, 0.78, 0.78, 0.78, 0.78, 0.78, 0.78, 0.78, 0.78, 0.78],
    "S13e": [0.80, 0.55, 1.25, 1.20, 1.30, 1.15, 1.20, 2.60],
    "S13f": [0.35, 0.75, 1.15, 1.55, 1.25, 1.35, 3.55],
    "S13g": [1.00, 1.45, 0.77, 0.77, 0.77, 0.77, 0.77, 0.77, 0.75, 0.80, 0.95],
    "S14": [1.35, 1.60, 1.50, 1.30, 1.55, 1.65, 1.05],
}


def set_run_font(run, *, bold: bool | None = None, italic: bool | None = None,
                 superscript: bool = False, subscript: bool = False) -> None:
    run.font.name = FONT
    run.font.size = Pt(FONT_SIZE_PT)
    run._element.rPr.rFonts.set(qn("w:ascii"), FONT)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run._element.rPr.rFonts.set(qn("w:cs"), FONT)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    run.font.superscript = superscript
    run.font.subscript = subscript


def configure_style(style, *, bold: bool = False, italic: bool = False,
                    before: float | None = None, after: float | None = None) -> None:
    style.font.name = FONT
    style.font.size = Pt(FONT_SIZE_PT)
    style.font.bold = bold
    style.font.italic = italic
    rpr = style._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)
    pf = style.paragraph_format
    pf.line_spacing = BODY_LINE_SPACING
    if before is not None:
        pf.space_before = Pt(before)
    if after is not None:
        pf.space_after = Pt(after)


def configure_section(section, *, landscape: bool) -> None:
    """Set a compact, publication-friendly Letter geometry for one section."""
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    section.page_width = Inches(11 if landscape else 8.5)
    section.page_height = Inches(8.5 if landscape else 11)
    if landscape:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.4)
        section.right_margin = Inches(0.4)
    else:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)


def configure_document(doc: Document) -> None:
    configure_section(doc.sections[0], landscape=False)

    configure_style(doc.styles["Normal"], after=6)
    for style_name in ("Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3",
                       "Heading 4", "Caption"):
        style = doc.styles[style_name]
        configure_style(
            style,
            bold=style_name not in ("Subtitle",),
            before=10 if style_name.startswith("Heading") else 0,
            after=5,
        )
        style.font.color.rgb = RGBColor(0, 0, 0)

    for style_name in ("List Bullet", "List Bullet 2", "List Bullet 3", "List Paragraph"):
        if style_name in doc.styles:
            configure_style(doc.styles[style_name], after=2)

    # Add a document-level math-font declaration.  It controls the OMML
    # intermediate; MathType will apply its own verified preferences later.
    settings = doc.settings.element
    for child in list(settings):
        if child.tag == qn("m:mathPr"):
            settings.remove(child)
    math_pr = OxmlElement("m:mathPr")
    math_font = OxmlElement("m:mathFont")
    math_font.set(qn("m:val"), FONT)
    math_pr.append(math_font)
    settings.append(math_pr)


def source_math_expressions(lines: list[str]) -> list[str]:
    expressions: list[str] = []
    index = 0
    while index < len(lines):
        if lines[index].strip() == DISPLAY_OPEN:
            index += 1
            body: list[str] = []
            while index < len(lines) and lines[index].strip() != DISPLAY_CLOSE:
                body.append(lines[index])
                index += 1
            expressions.append(TAG.sub("", "\n".join(body)).strip())
            index += 1
            continue
        expressions.extend(match.group(1).strip() for match in INLINE_MATH.finditer(lines[index]))
        index += 1
    return expressions


MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def mt(tag: str) -> str:
    return f"{{{MATH_NS}}}{tag}"


def wt(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "eta": "η", "theta": "θ", "lambda": "λ", "mu": "μ",
    "rho": "ρ", "sigma": "σ", "phi": "φ", "psi": "ψ", "Delta": "Δ",
    "Phi": "Φ", "Gamma": "Γ", "Sigma": "Σ", "top": "⊤", "times": "×",
    "cdot": "·", "leq": "≤", "geq": "≥", "sim": "∼", "in": "∈",
    "cup": "∪", "setminus": "∖", "subseteq": "⊆", "ldots": "…", "ell": "ℓ",
}


class TexOmmlParser:
    """Small TeX parser for the explicitly enumerated notation in this SI."""

    def __init__(self, tex: str, *, warnings: set[str] | None = None) -> None:
        self.tex = tex
        self.pos = 0
        self.warnings = warnings if warnings is not None else set()

    @staticmethod
    def _run(text: str, style: dict | None = None):
        run = OxmlElement("m:r")
        style = style or {}
        if style:
            rpr = OxmlElement("m:rPr")
            if style.get("script"):
                scr = OxmlElement("m:scr")
                scr.set(qn("m:val"), style["script"])
                rpr.append(scr)
            if style.get("kind"):
                sty = OxmlElement("m:sty")
                sty.set(qn("m:val"), style["kind"])
                rpr.append(sty)
            if style.get("plain"):
                rpr.append(OxmlElement("m:nor"))
            run.append(rpr)
        node = OxmlElement("m:t")
        node.text = text
        run.append(node)
        return run

    @staticmethod
    def _container(tag: str, children: list):
        node = OxmlElement(f"m:{tag}")
        for child in children:
            node.append(child)
        return node

    @staticmethod
    def _slot(tag: str, children: list):
        node = OxmlElement(f"m:{tag}")
        for child in children:
            node.append(child)
        return node

    def _skip_space(self) -> None:
        while self.pos < len(self.tex) and self.tex[self.pos].isspace():
            self.pos += 1

    def _group_text(self) -> str:
        self._skip_space()
        if self.pos >= len(self.tex) or self.tex[self.pos] != "{":
            return ""
        start = self.pos + 1
        depth = 1
        self.pos += 1
        while self.pos < len(self.tex) and depth:
            if self.tex[self.pos] == "{":
                depth += 1
            elif self.tex[self.pos] == "}":
                depth -= 1
                if depth == 0:
                    result = self.tex[start:self.pos]
                    self.pos += 1
                    return result
            self.pos += 1
        return self.tex[start:]

    def _command(self) -> str:
        if self.pos >= len(self.tex) or self.tex[self.pos] != "\\":
            return ""
        self.pos += 1
        if self.pos >= len(self.tex):
            return "\\"
        if self.tex[self.pos].isalpha():
            start = self.pos
            while self.pos < len(self.tex) and self.tex[self.pos].isalpha():
                self.pos += 1
            return self.tex[start:self.pos]
        value = self.tex[self.pos]
        self.pos += 1
        return value

    def _group(self, style: dict | None = None) -> list:
        self._skip_space()
        if self.pos < len(self.tex) and self.tex[self.pos] == "{":
            contents = self._group_text()
            return TexOmmlParser(contents, warnings=self.warnings).parse(style=style)
        return self._atom(style)

    def _delimiter(self) -> str:
        self._skip_space()
        if self.pos >= len(self.tex):
            return ""
        if self.tex[self.pos] == "\\":
            command = self._command()
            return {"{": "{", "}": "}", "|": "|"}.get(command, SYMBOLS.get(command, command))
        value = self.tex[self.pos]
        self.pos += 1
        return "" if value == "." else value

    def _fraction(self, style: dict | None):
        numerator = self._group(style)
        denominator = self._group(style)
        node = OxmlElement("m:f")
        node.append(self._slot("num", numerator))
        node.append(self._slot("den", denominator))
        return node

    def _radical(self, style: dict | None):
        content = self._group(style)
        node = OxmlElement("m:rad")
        props = OxmlElement("m:radPr")
        props.append(OxmlElement("m:degHide"))
        node.append(props)
        node.append(OxmlElement("m:deg"))
        node.append(self._slot("e", content))
        return node

    def _accent(self, kind: str, style: dict | None):
        content = self._group(style)
        if kind == "bar":
            node = OxmlElement("m:bar")
            props = OxmlElement("m:barPr")
            pos = OxmlElement("m:pos")
            pos.set(qn("m:val"), "top")
            props.append(pos)
            node.append(props)
            node.append(self._slot("e", content))
            return node
        node = OxmlElement("m:acc")
        props = OxmlElement("m:accPr")
        char = OxmlElement("m:chr")
        char.set(qn("m:val"), "ˆ")
        props.append(char)
        node.append(props)
        node.append(self._slot("e", content))
        return node

    def _nary(self, symbol: str):
        node = OxmlElement("m:nary")
        props = OxmlElement("m:naryPr")
        char = OxmlElement("m:chr")
        char.set(qn("m:val"), symbol)
        props.append(char)
        lim = OxmlElement("m:limLoc")
        lim.set(qn("m:val"), "undOvr")
        props.append(lim)
        node.append(props)
        node.append(OxmlElement("m:sub"))
        node.append(OxmlElement("m:sup"))
        node.append(OxmlElement("m:e"))
        return node

    def _aligned(self, style: dict | None):
        # The supplied SI contains no nested aligned environment.  Splitting
        # rows here preserves the equation array semantics in Office Math.
        end_token = r"\end{aligned}"
        end = self.tex.find(end_token, self.pos)
        if end < 0:
            self.warnings.add("unclosed aligned environment")
            raw = self.tex[self.pos:]
            self.pos = len(self.tex)
        else:
            raw = self.tex[self.pos:end]
            self.pos = end + len(end_token)
        rows = re.split(r"\\\\", raw)
        array = OxmlElement("m:eqArr")
        for row in rows:
            expression = row.replace("&", "").strip()
            array.append(self._slot("e", TexOmmlParser(expression, warnings=self.warnings).parse(style=style)))
        return array

    def _atom(self, style: dict | None = None) -> list:
        self._skip_space()
        if self.pos >= len(self.tex) or self.tex[self.pos] == "}":
            return []
        char = self.tex[self.pos]
        if char == "{":
            return self._group(style)
        if char == "&":
            self.pos += 1
            return []
        if char != "\\":
            self.pos += 1
            return [self._run(char, style)]

        command = self._command()
        if command == "frac":
            return [self._fraction(style)]
        if command == "sqrt":
            return [self._radical(style)]
        if command in {"bar", "widehat"}:
            return [self._accent(command, style)]
        if command == "sum":
            return [self._nary("∑")]
        if command == "left" or command == "right":
            delimiter = self._delimiter()
            return [self._run(delimiter, style)] if delimiter else []
        if command == "begin":
            environment = self._group_text()
            if environment == "aligned":
                return [self._aligned(style)]
            self.warnings.add(f"unsupported environment: {environment}")
            return []
        if command in {"mathrm", "text", "mbox"}:
            return self._group({**(style or {}), "plain": True, "kind": "p"})
        if command == "mathbf" or command == "boldsymbol":
            return self._group({**(style or {}), "kind": "b"})
        if command == "mathcal":
            return self._group({**(style or {}), "script": "script"})
        if command == "mathbb":
            return self._group({**(style or {}), "script": "double-struck"})
        if command in {"log", "max", "min", "exp"}:
            return [self._run(command, {**(style or {}), "plain": True, "kind": "p"})]
        if command in {",", "!", ";", ":", "quad", "qquad", " "}:
            return [self._run(" ", {**(style or {}), "plain": True})] if command != "!" else []
        if command == "\\":
            return []
        if command in {"{", "}"}:
            return [self._run(command, style)]
        if command in SYMBOLS:
            return [self._run(SYMBOLS[command], style)]
        self.warnings.add(f"unsupported TeX command: \\{command}")
        return [self._run(command, style)]

    def _attach_scripts(self, base: list, sub: list | None, sup: list | None) -> list:
        if not sub and not sup:
            return base
        if len(base) == 1 and base[0].tag == mt("nary"):
            nary = base[0]
            for child in list(nary):
                if child.tag == mt("sub"):
                    for value in sub or []:
                        child.append(value)
                if child.tag == mt("sup"):
                    for value in sup or []:
                        child.append(value)
            return base
        if sub and sup:
            node = OxmlElement("m:sSubSup")
            node.append(self._slot("e", base))
            node.append(self._slot("sub", sub))
            node.append(self._slot("sup", sup))
            return [node]
        if sub:
            node = OxmlElement("m:sSub")
            node.append(self._slot("e", base))
            node.append(self._slot("sub", sub))
            return [node]
        node = OxmlElement("m:sSup")
        node.append(self._slot("e", base))
        node.append(self._slot("sup", sup or []))
        return [node]

    def parse(self, style: dict | None = None) -> list:
        result: list = []
        while self.pos < len(self.tex):
            self._skip_space()
            if self.pos >= len(self.tex) or self.tex[self.pos] == "}":
                break
            if self.tex[self.pos] in "_^":
                # The source intentionally writes units as ordinary text plus
                # a short TeX exponent, e.g. kg\(^{-1}\). Preserve that
                # exponent as Office Math with an empty base so it remains a
                # convertible equation object rather than text formatting.
                marker = self.tex[self.pos]
                self.pos += 1
                value = self._group(style)
                result.extend(self._attach_scripts([], value if marker == "_" else None,
                                                   value if marker == "^" else None))
                continue
            base = self._atom(style)
            self._skip_space()
            sub = sup = None
            while self.pos < len(self.tex) and self.tex[self.pos] in "_^":
                marker = self.tex[self.pos]
                self.pos += 1
                value = self._group(style)
                if marker == "_":
                    sub = value
                else:
                    sup = value
                self._skip_space()
            result.extend(self._attach_scripts(base, sub, sup))
        return result


class OmmlConverter:
    def __init__(self) -> None:
        self._omml_by_tex: dict[str, str] = {}
        self.warnings: set[str] = set()

    def omml(self, tex: str):
        tex = tex.strip()
        if tex not in self._omml_by_tex:
            root = OxmlElement("m:oMath")
            for child in TexOmmlParser(tex, warnings=self.warnings).parse():
                root.append(child)
            self._omml_by_tex[tex] = etree.tostring(root, encoding="unicode")
        return parse_xml(self._omml_by_tex[tex])


def append_omml(paragraph, tex: str, converter: OmmlConverter) -> None:
    paragraph._p.append(converter.omml(tex))


class MathMlOmmlConverter:
    """Standards-based TeX -> MathML -> OMML conversion for final output."""

    def __init__(self, mathml_by_tex: dict[str, str]) -> None:
        if not MML2OMML_PATH.exists():
            raise FileNotFoundError(f"Microsoft MML2OMML stylesheet was not found: {MML2OMML_PATH}")
        self._transform = etree.XSLT(etree.parse(str(MML2OMML_PATH)))
        self._mathml_by_tex = mathml_by_tex
        self._omml_by_tex: dict[str, str] = {}
        self.warnings: set[str] = set()

    def omml(self, tex: str):
        tex = tex.strip()
        if tex not in self._omml_by_tex:
            if tex not in self._mathml_by_tex:
                raise KeyError(f"No MathML mapping was supplied for TeX: {tex}")
            root = self._transform(etree.fromstring(self._mathml_by_tex[tex].encode("utf-8"))).getroot()
            self._omml_by_tex[tex] = etree.tostring(root, encoding="unicode")
        return parse_xml(self._omml_by_tex[tex])


def append_text_run(paragraph, text: str, *, bold: bool = False, italic: bool = False,
                    superscript: bool = False, subscript: bool = False) -> None:
    if not text:
        return
    run = paragraph.add_run(html.unescape(text))
    set_run_font(
        run,
        bold=bold,
        italic=italic,
        superscript=superscript,
        subscript=subscript,
    )


def add_inline_content(paragraph, text: str, converter: OmmlConverter, *, bold: bool = False,
                       italic: bool = False) -> None:
    """Render source prose with nested emphasis, sup/sub tags, and TeX spans."""
    position = 0
    plain: list[str] = []

    def flush() -> None:
        if plain:
            append_text_run(paragraph, "".join(plain), bold=bold, italic=italic)
            plain.clear()

    while position < len(text):
        if text.startswith(r"\(", position):
            end = text.find(r"\)", position + 2)
            if end >= 0:
                flush()
                append_omml(paragraph, text[position + 2:end].strip(), converter)
                position = end + 2
                continue
        if text.startswith("**", position):
            end = text.find("**", position + 2)
            if end >= 0:
                flush()
                add_inline_content(paragraph, text[position + 2:end], converter, bold=True, italic=italic)
                position = end + 2
                continue
        if text.startswith("<sup>", position):
            end = text.find("</sup>", position + 5)
            if end >= 0:
                flush()
                append_text_run(paragraph, text[position + 5:end], bold=bold, italic=italic, superscript=True)
                position = end + 6
                continue
        if text.startswith("<sub>", position):
            end = text.find("</sub>", position + 5)
            if end >= 0:
                flush()
                append_text_run(paragraph, text[position + 5:end], bold=bold, italic=italic, subscript=True)
                position = end + 6
                continue
        if text.startswith("<br>", position) or text.startswith("<br/>", position):
            flush()
            paragraph.add_run().add_break()
            position += 4 if text.startswith("<br>", position) else 5
            continue
        if text[position] == "*" and not text.startswith("**", position):
            end = text.find("*", position + 1)
            if end >= 0:
                flush()
                add_inline_content(paragraph, text[position + 1:end], converter, bold=bold, italic=True)
                position = end + 1
                continue
        plain.append(text[position])
        position += 1
    flush()


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    tr_pr.append(hdr)


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def set_fixed_table_layout(table, widths: list[float]) -> None:
    """Apply explicit Word grid and cell widths in twips."""
    if len(widths) != len(table.columns):
        raise ValueError(f"Expected {len(table.columns)} table widths, got {len(widths)}")
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    total_twips = sum(round(width * 1440) for width in widths)
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.insert(0, tbl_w)
    tbl_w.set(qn("w:w"), str(total_twips))
    tbl_w.set(qn("w:type"), "dxa")

    grid_cols = table._tbl.tblGrid.findall(qn("w:gridCol"))
    for index, width in enumerate(widths):
        twips = round(width * 1440)
        table.columns[index].width = Inches(width)
        if index < len(grid_cols):
            grid_cols[index].set(qn("w:w"), str(twips))
        for cell in table.columns[index].cells:
            cell.width = Inches(width)
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(twips))
            tc_w.set(qn("w:type"), "dxa")


def table_rows(block: str) -> list[list[tuple[str, bool]]]:
    result: list[list[tuple[str, bool]]] = []
    for row_match in ROW_BLOCK.finditer(block):
        cells: list[tuple[str, bool]] = []
        for cell_match in CELL_BLOCK.finditer(row_match.group(1)):
            is_header = cell_match.group(1).lower() == "h"
            raw = re.sub(r"\s+", " ", cell_match.group(2)).strip()
            cells.append((html.unescape(raw), is_header))
        if cells:
            result.append(cells)
    return result


def add_html_table(doc: Document, block: str, converter: OmmlConverter,
                   table_id: str | None = None) -> None:
    rows = table_rows(block)
    if not rows:
        return
    column_count = max(len(row) for row in rows)
    table = doc.add_table(rows=len(rows), cols=column_count)
    table.style = "Table Grid"
    if table_id in WIDE_TABLE_WIDTHS:
        set_fixed_table_layout(table, WIDE_TABLE_WIDTHS[table_id])
    else:
        table.autofit = True
    for row_index, row_values in enumerate(rows):
        row = table.rows[row_index]
        prevent_row_split(row)
        if row_index == 0:
            set_repeat_table_header(row)
        for column_index in range(column_count):
            cell = row.cells[column_index]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_index == 0 else WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            if column_index < len(row_values):
                content, is_header = row_values[column_index]
                add_inline_content(paragraph, content, converter, bold=is_header)
            if row_index == 0:
                shade_cell(cell, "D9E2F3")
    doc.add_paragraph()


def paragraph_for(doc: Document, text: str, converter: OmmlConverter, *, style: str | None = None,
                  alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, before: float = 0,
                  after: float = 6, keep_with_next: bool = False):
    paragraph = doc.add_paragraph(style=style)
    paragraph.alignment = alignment
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = BODY_LINE_SPACING
    pf.keep_with_next = keep_with_next
    add_inline_content(paragraph, text, converter)
    return paragraph


def add_display_equation(doc: Document, tex: str, converter: OmmlConverter) -> None:
    tag_match = TAG.search(tex)
    tag = tag_match.group(1).strip() if tag_match else ""
    equation = TAG.sub("", tex).strip()
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(3)
    paragraph.paragraph_format.space_after = Pt(3)
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.keep_together = True
    section = doc.sections[-1]
    usable_width = section.page_width.inches - section.left_margin.inches - section.right_margin.inches
    tabs = paragraph.paragraph_format.tab_stops
    tabs.add_tab_stop(Inches(usable_width / 2), WD_TAB_ALIGNMENT.CENTER)
    tabs.add_tab_stop(Inches(usable_width), WD_TAB_ALIGNMENT.RIGHT)
    paragraph.add_run().add_tab()
    append_omml(paragraph, equation, converter)
    if tag:
        paragraph.add_run().add_tab()
        append_text_run(paragraph, f"({tag})")


def add_bullet(doc: Document, text: str, converter: OmmlConverter) -> None:
    paragraph = doc.add_paragraph(style="List Bullet")
    paragraph.paragraph_format.space_after = Pt(2)
    paragraph.paragraph_format.line_spacing = BODY_LINE_SPACING
    add_inline_content(paragraph, text, converter)


def build_document(source: Path, output: Path, supplied_mathml: Path | None = None) -> None:
    lines = source.read_text(encoding="utf-8-sig").splitlines()
    converter = (
        MathMlOmmlConverter(json.loads(supplied_mathml.read_text(encoding="utf-8")))
        if supplied_mathml is not None
        else OmmlConverter()
    )
    doc = Document()
    configure_document(doc)
    doc.core_properties.title = source.stem or "Scientific Markdown"
    doc.core_properties.subject = "Scientific Markdown converted to an OMML intermediate"
    doc.core_properties.author = ""

    index = 0
    title_count = 0
    current_layout = "portrait"
    pending_table_id: str | None = None
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()
        if not stripped:
            index += 1
            continue
        if stripped == DISPLAY_OPEN:
            index += 1
            equation_lines: list[str] = []
            while index < len(lines) and lines[index].strip() != DISPLAY_CLOSE:
                equation_lines.append(lines[index])
                index += 1
            if index >= len(lines):
                raise ValueError("Unclosed display-math block in source")
            add_display_equation(doc, "\n".join(equation_lines), converter)
            index += 1
            continue
        if stripped == "<table>":
            index += 1
            table_lines: list[str] = []
            while index < len(lines) and lines[index].strip() != "</table>":
                table_lines.append(lines[index])
                index += 1
            if index >= len(lines):
                raise ValueError("Unclosed HTML table in source")
            add_html_table(doc, "\n".join(table_lines), converter, pending_table_id)
            pending_table_id = None
            index += 1
            continue
        heading = HEADING.match(stripped)
        if heading:
            level = len(heading.group(1))
            label = heading.group(2)
            table_caption = TABLE_CAPTION.match(label)
            table_id = table_caption.group(1) if table_caption else None

            # Every wide table begins on a fresh landscape page. Notes and
            # explanatory text remain in the same section until the next
            # heading, where portrait geometry is restored as needed.
            if table_id in WIDE_TABLE_WIDTHS:
                section = doc.add_section(WD_SECTION.NEW_PAGE)
                configure_section(section, landscape=True)
                current_layout = "landscape"
            elif current_layout == "landscape":
                section = doc.add_section(WD_SECTION.NEW_PAGE)
                configure_section(section, landscape=False)
                current_layout = "portrait"

            if level == 1:
                title_count += 1
                style = "Title" if title_count == 1 else "Heading 1"
                paragraph_for(
                    doc,
                    label,
                    converter,
                    style=style,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER if title_count <= 2 else WD_ALIGN_PARAGRAPH.LEFT,
                    after=8,
                    keep_with_next=True,
                )
            elif table_id:
                paragraph_for(
                    doc,
                    label,
                    converter,
                    style="Caption",
                    alignment=WD_ALIGN_PARAGRAPH.LEFT,
                    before=0,
                    after=4,
                    keep_with_next=True,
                )
                pending_table_id = table_id
            elif level == 2:
                paragraph_for(doc, label, converter, style="Heading 1", alignment=WD_ALIGN_PARAGRAPH.LEFT,
                              before=10, after=5, keep_with_next=True)
            elif level == 3:
                paragraph_for(doc, label, converter, style="Heading 2", alignment=WD_ALIGN_PARAGRAPH.LEFT,
                              before=7, after=4, keep_with_next=True)
            else:
                paragraph_for(doc, label, converter, style="Heading 3", alignment=WD_ALIGN_PARAGRAPH.LEFT,
                              before=5, after=3, keep_with_next=True)
            index += 1
            continue
        bullet = LIST_ITEM.match(stripped)
        if bullet:
            add_bullet(doc, bullet.group(1), converter)
            index += 1
            continue
        paragraph_lines = [stripped]
        index += 1
        while index < len(lines):
            probe = lines[index].strip()
            if not probe or probe in (DISPLAY_OPEN, "<table>") or HEADING.match(probe) or LIST_ITEM.match(probe):
                break
            paragraph_lines.append(probe)
            index += 1
        text = " ".join(paragraph_lines)
        # Figure and table notes are left aligned; regular source prose is justified.
        left = text.startswith("**Fig.") or text.startswith("**Note.") or text.startswith("[")
        paragraph_for(
            doc,
            text,
            converter,
            alignment=WD_ALIGN_PARAGRAPH.LEFT if left else WD_ALIGN_PARAGRAPH.JUSTIFY,
            after=4 if left else 6,
        )

    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)
    if converter.warnings:
        print("OMML parser warnings: " + "; ".join(sorted(converter.warnings)))


def main() -> None:
    if len(sys.argv) == 4 and sys.argv[1] == "--extract-math":
        source = Path(sys.argv[2])
        destination = Path(sys.argv[3])
        destination.write_text(
            json.dumps(list(dict.fromkeys(source_math_expressions(source.read_text(encoding="utf-8-sig").splitlines()))), ensure_ascii=False),
            encoding="utf-8",
        )
        print(f"Extracted math to {destination}")
        return
    if len(sys.argv) not in (3, 4):
        raise SystemExit("Usage: build_si_omml_intermediate.py <source-markdown> <output.docx> [mathml.json]")
    source = Path(sys.argv[1])
    output = Path(sys.argv[2])
    if not source.is_file():
        raise FileNotFoundError(f"Source Markdown was not found: {source}")
    supplied_mathml = Path(sys.argv[3]) if len(sys.argv) == 4 else None
    build_document(source, output, supplied_mathml)
    print(f"Created {output}")


if __name__ == "__main__":
    main()
