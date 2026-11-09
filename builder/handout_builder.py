#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Course handout builder: JSON spec + Markdown bodies -> styled DOCX.

Usage:
    python handout_builder.py build --spec handout.json --out handout.docx [--force]
    python handout_builder.py check --spec handout.json

The spec schema is documented in skills/course-handout.md. All output is a
single JSON object on stdout; human diagnostics go to stderr.
"""

import argparse
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BODY_EA = "SimSun"
HEAD_EA = "Microsoft YaHei"
MONO = "Consolas"

INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x59, 0x59, 0x59)
NAVY = RGBColor(0x14, 0x36, 0x5C)
BLUE = RGBColor(0x1F, 0x4E, 0x79)
HEAD3 = RGBColor(0x2E, 0x54, 0x7A)
CODE_INK = RGBColor(0x1F, 0x23, 0x28)
AMBER = RGBColor(0x9A, 0x5B, 0x00)
RED = RGBColor(0x9C, 0x2F, 0x2F)
GREEN = RGBColor(0x1E, 0x5B, 0x33)

TEXT_WIDTH_CM = 16.2

PALETTES = {
    "goal": ("EAF1FB", "1F4E79", BLUE),
    "summary": ("EAF5EC", "2E7D32", GREEN),
    "key": ("FFF8E7", "E3A33B", AMBER),
    "tip": ("FFF8E7", "E3A33B", AMBER),
    "error": ("FDECEA", "C0392B", RED),
    "quote": ("F4F6F8", "9AA5B1", MUTED),
}

_PBDR_SUCCESSORS = (
    "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
)
_SHD_SUCCESSORS = _PBDR_SUCCESSORS[1:]


# ---------------------------------------------------------------------------
# Low-level DOCX helpers
# ---------------------------------------------------------------------------

def _twips(width):
    return int(round(width.cm / 2.54 * 1440))


def set_run_font(run, size, bold=False, color=INK, latin=BODY_EA, ea=BODY_EA,
                 italic=False, spacing=None):
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), latin)
    r_fonts.set(qn("w:hAnsi"), latin)
    r_fonts.set(qn("w:eastAsia"), ea)
    if spacing is not None:
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(spacing))
        r_pr.append(sp)


def style_east_asia(style, ea, latin=None):
    r_pr = style.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:eastAsia"), ea)
    if latin is not None:
        r_fonts.set(qn("w:ascii"), latin)
        r_fonts.set(qn("w:hAnsi"), latin)


def add_bottom_border(paragraph, color="14365C", sz=8, space=4):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.insert_element_before(p_bdr, *_PBDR_SUCCESSORS)


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.insert_element_before(
        shd, "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
        "w:vAlign", "w:hideMark",
    )


def set_cell_borders(cell, color="D8DEE4", sz=4,
                     edges=("top", "left", "bottom", "right")):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.insert_element_before(
            tc_borders, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
            "w:tcFitText", "w:vAlign", "w:hideMark",
        )
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        if edge in edges:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "nil")
        tc_borders.append(el)


def set_table_width(table, width):
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.insert_element_before(
            tbl_w, "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders",
            "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        )
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(_twips(width)))


def set_table_borders(table, color="C9D3DF", sz=4):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tbl_pr.insert_element_before(
        borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription",
    )


def set_table_cell_margins(table, top=45, left=100, bottom=45, right=100):
    tbl_pr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for tag, val in (("top", top), ("left", left), ("bottom", bottom),
                     ("right", right)):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tbl_pr.insert_element_before(
        mar, "w:tblLook", "w:tblCaption", "w:tblDescription",
    )


def set_row_flags(row, header=False):
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))
    if header:
        hdr = OxmlElement("w:tblHeader")
        hdr.set(qn("w:val"), "true")
        tr_pr.append(hdr)


def add_bookmark(paragraph, name, bid):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    idx = 1 if paragraph._p.pPr is not None else 0
    paragraph._p.insert(idx, start)
    paragraph._p.append(end)


def add_internal_link(paragraph, text, anchor, size, bold=False, color=BLUE):
    run = paragraph.add_run(text)
    set_run_font(run, size, bold, color)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    paragraph._p.append(hyperlink)
    hyperlink.append(run._r)


# ---------------------------------------------------------------------------
# Inline markup: **bold** and `code`
# ---------------------------------------------------------------------------

INLINE_RE = re.compile(r"(\*\*.+?\*\*|`[^`\n]+`)")


def add_inline(paragraph, text, size=10.5, bold=False, color=INK):
    code_size = max(size - 1.2, 7.5)
    for part in INLINE_RE.split(text or ""):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            inner = part[2:-2]
            if "`" in inner:
                add_inline(paragraph, inner, size=size, bold=True, color=color)
            else:
                run = paragraph.add_run(inner)
                set_run_font(run, size, True, color)
        elif part.startswith("`") and part.endswith("`") and len(part) > 2:
            run = paragraph.add_run(part[1:-1])
            set_run_font(run, code_size, bold, CODE_INK, latin=MONO, ea=BODY_EA)
        else:
            run = paragraph.add_run(part)
            set_run_font(run, size, bold, color)


# ---------------------------------------------------------------------------
# Block renderers
# ---------------------------------------------------------------------------

def body(doc, text, size=10.5, after=4, before=0, line=1.45, color=INK,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY, left=0.0, right=0.0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.alignment = align
    if left:
        pf.left_indent = Cm(left)
    if right:
        pf.right_indent = Cm(right)
    add_inline(p, text, size=size, color=color)
    return p


def bullet(doc, text, size=10.5, after=2.5):
    p = doc.add_paragraph(style="List Bullet")
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(after)
    pf.line_spacing = 1.4
    add_inline(p, text, size=size)
    return p


def numbered(doc, index, text, size=10.5):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.45
    pf.left_indent = Cm(0.75)
    pf.first_line_indent = Cm(-0.75)
    run = p.add_run(f"{index}. ")
    set_run_font(run, size, True, BLUE)
    add_inline(p, text, size=size)
    return p


def heading(doc, text, level):
    p = doc.add_paragraph(style=f"Heading {level}")
    run = p.add_run(text)
    if level == 1:
        set_run_font(run, 20, True, NAVY, latin=HEAD_EA, ea=HEAD_EA)
        add_bottom_border(p, "B7C9DE", 8, 6)
    elif level == 2:
        set_run_font(run, 14.5, True, BLUE, latin=HEAD_EA, ea=HEAD_EA)
    else:
        set_run_font(run, 12, True, HEAD3, latin=HEAD_EA, ea=HEAD_EA)
    return p


def spacer(doc, points=3):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(points)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("")
    run.font.size = Pt(1)
    return p


def _chunk_lines(lines, limit=26):
    chunks = []
    current = []
    for line in lines:
        current.append(line)
        if len(current) >= limit:
            cut = None
            for i in range(len(current) - 1, 9, -1):
                if current[i].strip() == "":
                    cut = i
                    break
            if cut is not None:
                chunks.append(current[:cut])
                current = current[cut:]
            else:
                chunks.append(current)
                current = []
    if current:
        chunks.append(current)
    return chunks


def _code_table(doc, lines):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_width(table, Cm(TEXT_WIDTH_CM))
    cell = table.cell(0, 0)
    cell.width = Cm(TEXT_WIDTH_CM)
    shade_cell(cell, "F6F8FA")
    set_cell_borders(cell, "D5DCE3", 4)
    set_table_cell_margins(table, top=85, left=115, bottom=85, right=115)
    set_row_flags(table.rows[0])

    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(11.6)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("\n".join(lines))
    set_run_font(run, 8.6, False, CODE_INK, latin=MONO, ea=MONO)
    for t in run._r.findall(qn("w:t")):
        t.set(qn("xml:space"), "preserve")


def code_block(doc, code):
    lines = code.strip("\n").split("\n")
    chunks = _chunk_lines(lines)
    for idx, chunk in enumerate(chunks):
        _code_table(doc, chunk)
        if idx != len(chunks) - 1:
            spacer(doc, 2)
    spacer(doc, 4)


def data_table(doc, headers, rows, widths=None):
    ncols = max(len(headers), 1)
    table = doc.add_table(rows=1, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_width(table, Cm(TEXT_WIDTH_CM))
    set_table_borders(table, "C9D3DF", 4)
    set_table_cell_margins(table, top=26, left=95, bottom=26, right=95)

    if widths is None:
        widths = [100.0 / ncols] * ncols
    col_widths = [Cm(TEXT_WIDTH_CM * w / 100.0) for w in widths]
    for i, width in enumerate(col_widths):
        table.columns[i].width = width

    header_row = table.rows[0]
    set_row_flags(header_row, header=True)
    for i, text in enumerate(headers):
        cell = header_row.cells[i]
        cell.width = col_widths[i]
        shade_cell(cell, "DCE6F2")
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        pf = p.paragraph_format
        pf.space_before = Pt(0.8)
        pf.space_after = Pt(0.8)
        pf.line_spacing = 1.05
        add_inline(p, text, size=9.3, bold=True, color=NAVY)

    for idx, row in enumerate(rows):
        cells = table.add_row().cells
        set_row_flags(table.rows[-1])
        for i in range(ncols):
            cell = cells[i]
            cell.width = col_widths[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            if idx % 2 == 1:
                shade_cell(cell, "F6F8FB")
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.space_before = Pt(0.8)
            pf.space_after = Pt(0.8)
            pf.line_spacing = 1.05
            if len(rows) <= 8 and idx < len(rows) - 1:
                pf.keep_with_next = True
            add_inline(p, row[i] if i < len(row) else "", size=9.2)
    if len(rows) <= 8:
        for cell in header_row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_with_next = True
    spacer(doc, 3)


def callout(doc, title, lines, kind="tip"):
    fill, bar, title_color = PALETTES.get(kind, PALETTES["tip"])
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_width(table, Cm(TEXT_WIDTH_CM))
    cell = table.cell(0, 0)
    cell.width = Cm(TEXT_WIDTH_CM)
    shade_cell(cell, fill)
    set_cell_borders(cell, color=bar, sz=24, edges=("left",))
    set_table_cell_margins(table, top=110, left=190, bottom=110, right=160)
    set_row_flags(table.rows[0])

    p = cell.paragraphs[0]
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.45
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(title + "　")
    set_run_font(run, 9.5, True, title_color, latin=HEAD_EA, ea=HEAD_EA)
    add_inline(p, lines[0] if lines else "", size=9.8)

    for line in lines[1:]:
        p2 = cell.add_paragraph()
        pf2 = p2.paragraph_format
        pf2.space_before = Pt(2)
        pf2.space_after = Pt(0)
        pf2.line_spacing = 1.45
        p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_inline(p2, line, size=9.8)
    spacer(doc, 4)


def figure(doc, image_path, caption=""):
    if not os.path.isfile(image_path):
        body(doc, f"（图片缺失：{image_path}）", size=9, color=RED)
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    width_cm = 14.5
    try:
        from PIL import Image
        with Image.open(image_path) as im:
            w, h = im.size
        if w and h:
            height_cm = width_cm * h / w
            if height_cm > 17.5:
                width_cm = 17.5 * w / h
    except Exception:  # noqa: BLE001 - size probing is best effort
        pass
    try:
        run.add_picture(image_path, width=Cm(width_cm))
    except Exception as exc:  # noqa: BLE001 - report and continue
        body(doc, f"（图片无法插入：{image_path}，{exc}）", size=9, color=RED)
        return
    if caption:
        c = doc.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraph_format.space_before = Pt(0)
        c.paragraph_format.space_after = Pt(6)
        r = c.add_run(caption)
        set_run_font(r, 9, False, MUTED)


# ---------------------------------------------------------------------------
# Markdown-ish body parser
# ---------------------------------------------------------------------------

NOTE_LABELS = ("提示", "注意", "易错", "结论", "要点", "思考", "说明", "警告")


def _is_special(line):
    s = line.strip()
    if not s:
        return True
    if s.startswith("```") or s.startswith("|") or s.startswith(">"):
        return True
    if s == "---":
        return True
    if re.match(r"^#{1,4}\s+", s):
        return True
    if re.match(r"^!\[.*?\]\(.*?\)\s*$", s):
        return True
    if re.match(r"^[-*]\s+", s) or re.match(r"^\d+[.)]\s+", s):
        return True
    return False


def _join_lines(lines):
    out = lines[0]
    for nxt in lines[1:]:
        if out and nxt:
            a, b = out[-1], nxt[0]
            if a.isascii() and b.isascii() and a.isalnum() and b.isalnum():
                out += " " + nxt
            else:
                out += nxt
        else:
            out += nxt
    return out


def _split_row(line):
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def parse_markdown(text, base_dir):
    blocks = []
    if not text:
        return blocks
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
            continue

        if line.lstrip().startswith("```"):
            i += 1
            code = []
            while i < len(lines) and not lines[i].lstrip().startswith("```"):
                code.append(lines[i])
                i += 1
            i += 1
            blocks.append({"type": "code", "text": "\n".join(code)})
            continue

        if line.strip() == "---":
            i += 1
            continue

        m = re.match(r"^(#{1,4})\s+(.*)$", line.strip())
        if m:
            level = min(len(m.group(1)), 3)
            blocks.append({"type": f"h{level}", "text": m.group(2).strip()})
            i += 1
            continue

        m = re.match(r"^!\[(.*?)\]\((.*?)\)\s*$", line.strip())
        if m:
            path = m.group(2)
            if not os.path.isabs(path):
                path = os.path.join(base_dir, path)
            blocks.append({"type": "figure", "path": path, "caption": m.group(1)})
            i += 1
            continue

        if line.lstrip().startswith("|"):
            raw = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                raw.append(lines[i])
                i += 1
            if len(raw) >= 2:
                headers = _split_row(raw[0])
                body_rows = []
                for r in raw[1:]:
                    cells = _split_row(r)
                    if cells and all(re.match(r"^:?-{2,}:?$", c.replace(" ", "")) for c in cells):
                        continue
                    body_rows.append(cells)
                blocks.append({"type": "table", "headers": headers, "rows": body_rows})
            continue

        if line.startswith(">"):
            quoted = []
            while i < len(lines) and lines[i].startswith(">"):
                quoted.append(re.sub(r"^>\s?", "", lines[i]))
                i += 1
            text_q = _join_lines([q.strip() for q in quoted if q.strip()])
            label = None
            m = re.match(r"^(%s)[：:]\s*(.*)$" % "|".join(NOTE_LABELS), text_q)
            if m:
                label = m.group(1)
                text_q = m.group(2)
            blocks.append({"type": "note", "label": label, "text": text_q})
            continue

        if re.match(r"^[-*]\s+", line):
            items = []
            while i < len(lines) and re.match(r"^[-*]\s+", lines[i]):
                items.append(re.sub(r"^[-*]\s+", "", lines[i]).strip())
                i += 1
            blocks.append({"type": "bullets", "items": items})
            continue

        if re.match(r"^\d+[.)]\s+", line):
            items = []
            while i < len(lines) and re.match(r"^\d+[.)]\s+", lines[i]):
                items.append(re.sub(r"^\d+[.)]\s+", "", lines[i]).strip())
                i += 1
            blocks.append({"type": "numbered", "items": items})
            continue

        para = [line.strip()]
        i += 1
        while i < len(lines) and lines[i].strip() and not _is_special(lines[i]):
            para.append(lines[i].strip())
            i += 1
        blocks.append({"type": "p", "text": _join_lines(para)})
    return blocks


def render_blocks(doc, blocks, base_dir):
    for block in blocks:
        kind = block["type"]
        if kind == "p":
            body(doc, block.get("text", ""))
        elif kind == "bullets":
            for item in block.get("items", []):
                bullet(doc, item)
            spacer(doc, 3)
        elif kind == "numbered":
            for idx, item in enumerate(block.get("items", []), 1):
                numbered(doc, idx, item)
            spacer(doc, 3)
        elif kind == "code":
            code_block(doc, block.get("text", ""))
        elif kind == "table":
            data_table(doc, block.get("headers", []), block.get("rows", []))
        elif kind == "note":
            label = block.get("label") or "提示"
            callout(doc, label, [block.get("text", "")],
                    "error" if label in ("易错", "警告") else
                    "goal" if label in ("结论", "要点") else "tip")
        elif kind == "h2":
            heading(doc, block.get("text", ""), 2)
        elif kind == "h3":
            heading(doc, block.get("text", ""), 3)
        elif kind == "figure":
            figure(doc, block.get("path", ""), block.get("caption", ""))


# ---------------------------------------------------------------------------
# Spec helpers, validation, stats
# ---------------------------------------------------------------------------

def _iter_chapters(spec):
    for mi, module in enumerate(spec.get("modules") or []):
        for ci, chapter in enumerate(module.get("chapters") or []):
            yield mi, ci, module, chapter


def _all_strings(node):
    if isinstance(node, str):
        yield node
    elif isinstance(node, list):
        for item in node:
            yield from _all_strings(item)
    elif isinstance(node, dict):
        for value in node.values():
            yield from _all_strings(value)


def collect_stats(spec):
    modules = len(spec.get("modules") or [])
    chapters = 0
    sections = 0
    questions = 0
    code_lines = 0
    blocks = 0

    for module in spec.get("modules") or []:
        for chapter in module.get("chapters") or []:
            chapters += 1
            sections += len(chapter.get("sections") or [])
            ex = chapter.get("exercises") or {}
            if isinstance(ex, dict):
                for items in ex.values():
                    questions += len(items or [])
            for section in chapter.get("sections") or []:
                for block in parse_markdown(section.get("body", ""), "."):
                    blocks += 1
                    if block["type"] == "code":
                        code_lines += block["text"].count("\n") + 1
    chars = sum(len(s) for s in _all_strings(spec))
    est_pages = int(round(chars / 850 + code_lines / 46 + chapters * 0.55 + 2))
    return {
        "modules": modules,
        "chapters": chapters,
        "sections": sections,
        "blocks": blocks,
        "questions": questions,
        "chars": chars,
        "code_lines": code_lines,
        "est_pages": est_pages,
    }


def check_spec(spec):
    issues = []
    stats = collect_stats(spec)
    meta = spec.get("meta") or {}

    def err(path, message):
        issues.append({"level": "error", "path": path, "message": message})

    def warn(path, message):
        issues.append({"level": "warning", "path": path, "message": message})

    if not meta.get("title"):
        err("meta.title", "缺少讲义标题")
    if not spec.get("modules"):
        err("modules", "至少需要一个模块")
        return issues, stats

    seen = {}
    for mi, module in enumerate(spec.get("modules") or []):
        chapters = module.get("chapters") or []
        if not chapters:
            err(f"modules[{mi}]", "模块下没有章节")
        if len(chapters) == 1 and len(spec.get("modules") or []) > 1:
            warn(f"modules[{mi}]", "模块只有一章，考虑合并或补充章节")
        for ci, chapter in enumerate(chapters):
            path = f"modules[{mi}].chapters[{ci}]"
            title = (chapter.get("title") or "").strip()
            if not title:
                err(path, "章节缺少标题")
            else:
                if title in seen:
                    warn(path, f"章节标题重复：{title}")
                seen[title] = path
            sections = chapter.get("sections") or []
            if not sections:
                err(path, "章节没有小节")
            elif len(sections) < 2:
                warn(path, "章节只有一节，建议拆成 2–6 节")
            objectives = chapter.get("objectives") or []
            if len(objectives) < 2:
                warn(path, "学习目标少于 2 条")
            summary = chapter.get("summary") or []
            if len(summary) < 2:
                warn(path, "本章小结少于 2 条")
            ex = chapter.get("exercises") or {}
            chapter_questions = sum(len(v or []) for v in ex.values()) if isinstance(ex, dict) else 0
            if chapter_questions < 3:
                warn(path, "习题少于 3 道")
            answers = chapter.get("answers") or []
            if chapter_questions and len(answers) < chapter_questions:
                warn(path, f"参考答案 {len(answers)} 条，少于题目数 {chapter_questions}")
            chapter_chars = 0
            for section in sections:
                chapter_chars += len(section.get("body") or "")
                if len((section.get("body") or "").strip()) < 200:
                    warn(f"{path}.sections", "小节正文偏短（少于 200 字）")
            if chapter_chars < 600:
                warn(path, "本章正文偏少（少于 600 字）")

    if stats["chars"] < 3000:
        warn("meta", "全文内容偏少，预计不足 5 页")
    return issues, stats


# ---------------------------------------------------------------------------
# Document builder
# ---------------------------------------------------------------------------

class HandoutBuilder:
    def __init__(self, spec, out_path):
        self.spec = spec
        self.out_path = out_path
        self.spec_dir = os.path.dirname(os.path.abspath(spec.get("__path__", ".")))
        self.doc = Document()
        self.bid = 1
        self.chapter_no = 0
        self.module_no = 0

    # -- setup ---------------------------------------------------------
    def setup(self):
        sec = self.doc.sections[0]
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.2)
        sec.bottom_margin = Cm(2.1)
        sec.left_margin = Cm(2.4)
        sec.right_margin = Cm(2.4)
        sec.header_distance = Cm(1.1)
        sec.footer_distance = Cm(1.1)
        sec.different_first_page_header_footer = True

        normal = self.doc.styles["Normal"]
        normal.font.name = BODY_EA
        normal.font.size = Pt(10.5)
        normal.font.color.rgb = INK
        style_east_asia(normal, BODY_EA, BODY_EA)
        normal.paragraph_format.line_spacing = 1.45
        normal.paragraph_format.space_after = Pt(4)

        h1 = self.doc.styles["Heading 1"]
        h1.font.name = HEAD_EA
        h1.font.size = Pt(20)
        h1.font.bold = True
        h1.font.italic = False
        h1.font.color.rgb = NAVY
        style_east_asia(h1, HEAD_EA, HEAD_EA)
        h1.paragraph_format.page_break_before = True
        h1.paragraph_format.space_before = Pt(0)
        h1.paragraph_format.space_after = Pt(16)
        h1.paragraph_format.keep_with_next = True
        h1.paragraph_format.line_spacing = 1.15

        h2 = self.doc.styles["Heading 2"]
        h2.font.name = HEAD_EA
        h2.font.size = Pt(14.5)
        h2.font.bold = True
        h2.font.italic = False
        h2.font.color.rgb = BLUE
        style_east_asia(h2, HEAD_EA, HEAD_EA)
        h2.paragraph_format.space_before = Pt(15)
        h2.paragraph_format.space_after = Pt(6)
        h2.paragraph_format.keep_with_next = True
        h2.paragraph_format.line_spacing = 1.2

        h3 = self.doc.styles["Heading 3"]
        h3.font.name = HEAD_EA
        h3.font.size = Pt(12)
        h3.font.bold = True
        h3.font.italic = False
        h3.font.color.rgb = HEAD3
        style_east_asia(h3, HEAD_EA, HEAD_EA)
        h3.paragraph_format.space_before = Pt(10)
        h3.paragraph_format.space_after = Pt(4)
        h3.paragraph_format.keep_with_next = True

        lb = self.doc.styles["List Bullet"]
        lb.font.name = BODY_EA
        lb.font.size = Pt(10.5)
        style_east_asia(lb, BODY_EA, BODY_EA)

        self.build_footer(sec)

    def build_footer(self, sec):
        p = sec.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)

        def txt(value):
            run = p.add_run(value)
            set_run_font(run, 9, False, MUTED, latin=HEAD_EA, ea=HEAD_EA)

        title = (self.spec.get("meta") or {}).get("title") or "课程讲义"
        txt(f"{title}　·　第 ")
        self._add_field(p, "PAGE")
        txt(" 页 / 共 ")
        self._add_field(p, "NUMPAGES")
        txt(" 页")

    @staticmethod
    def _add_field(paragraph, instr, placeholder="1"):
        run = paragraph.add_run()
        set_run_font(run, 9, False, MUTED, latin=HEAD_EA, ea=HEAD_EA)
        r = run._r
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instr_el = OxmlElement("w:instrText")
        instr_el.set(qn("xml:space"), "preserve")
        instr_el.text = f" {instr} "
        sep = OxmlElement("w:fldChar")
        sep.set(qn("w:fldCharType"), "separate")
        text_el = OxmlElement("w:t")
        text_el.text = placeholder
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        for el in (begin, instr_el, sep, text_el, end):
            r.append(el)

    # -- cover / front matter ------------------------------------------
    def cover(self):
        meta = self.spec.get("meta") or {}
        title = meta.get("title") or "课程讲义"
        subtitle = meta.get("subtitle") or ""
        series = meta.get("series") or "课程讲义"
        author = meta.get("author") or ""
        date = meta.get("date") or ""
        audience = meta.get("audience") or ""
        stats = collect_stats(self.spec)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(100)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(series)
        set_run_font(run, 10.5, True, MUTED, latin=HEAD_EA, ea=HEAD_EA, spacing=40)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(26)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(title)
        set_run_font(run, 34 if len(title) <= 12 else 28, True, NAVY,
                     latin=HEAD_EA, ea=HEAD_EA)

        if subtitle:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(subtitle)
            set_run_font(run, 16, False, MUTED, latin=HEAD_EA, ea=HEAD_EA)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(30)
        p.paragraph_format.space_after = Pt(0)
        add_bottom_border(p, "14365C", 12, 1)

        if audience:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(24)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run("适用对象：" + audience)
            set_run_font(run, 10.5, False, MUTED, latin=HEAD_EA, ea=HEAD_EA)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(0)
        extras = f" · 含 {stats['questions']} 道习题" if stats["questions"] else ""
        run = p.add_run(
            f"共 {stats['modules']} 个模块 · {stats['chapters']} 章 · "
            f"{stats['sections']} 节{extras}"
        )
        set_run_font(run, 9.5, False, MUTED, latin=HEAD_EA, ea=HEAD_EA)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(215)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(author or "课程讲义")
        set_run_font(run, 10.5, False, MUTED, latin=HEAD_EA, ea=HEAD_EA)

        if date:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(date)
            set_run_font(run, 9, False, MUTED, latin=HEAD_EA, ea=HEAD_EA)

    def preface(self):
        meta = self.spec.get("meta") or {}
        h = heading(self.doc, "前言", 1)
        add_bookmark(h, "front-preface", self._next_bid())

        custom = self.spec.get("preface")
        if custom:
            render_blocks(self.doc, parse_markdown(custom, self.spec_dir), self.spec_dir)
        else:
            title = meta.get("title") or "这份讲义"
            audience = meta.get("audience") or "学习者"
            body(self.doc,
                 f"这份讲义围绕**{title}**展开，面向{audience}。"
                 "它把内容拆成若干模块与章节：每一章先给出学习目标与先修知识，"
                 "再用“直觉引入 → 严格定义 → 例题 → 常见错误 → 小结”的顺序展开，"
                 "最后用分层习题检验掌握程度。")
            body(self.doc,
                 "章节的组织参考了两类教材的长处：一类是中文教材的严谨结构与完整知识链，"
                 "另一类是英文教材的直观解释、图示与“为什么”。"
                 "阅读时不必平均用力：先抓住每章的目标和小结，再回头精读推导与例题。")

        heading(self.doc, "怎么用这份讲义", 2)
        bullet(self.doc, "先看每章开头的**学习目标**，带着问题读正文，读完对照目标自检。")
        bullet(self.doc, "正文中的**例题**先遮住解答自己做一遍，再看解析。")
        bullet(self.doc, "每章末尾的**本章小结**用来串联知识点，**习题**用来暴露盲区。")
        bullet(self.doc, "遇到标有“易错”的方框，停下来想想自己是否也这样理解。")

        heading(self.doc, "符号约定", 2)
        bullet(self.doc, "**粗体**表示重点结论或关键词；`等宽字体`表示代码、公式中的变量名或文件名。")
        bullet(self.doc, "浅蓝方框是学习目标，浅绿方框是本章小结，浅黄方框是提示，浅红方框是易错点。")
        bullet(self.doc, "灰底代码块可直接复制运行；表格用于对比与速查。")

    def toc(self):
        p = self.doc.add_paragraph()
        p.paragraph_format.page_break_before = True
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run("目　录")
        set_run_font(run, 26, True, NAVY, latin=HEAD_EA, ea=HEAD_EA, spacing=60)
        add_bottom_border(p, "B7C9DE", 8, 8)
        spacer(self.doc, 10)

        ch_no = 0
        for mi, module in enumerate(self.spec.get("modules") or [], 1):
            mp = self.doc.add_paragraph()
            mp.paragraph_format.space_before = Pt(10 if mi > 1 else 2)
            mp.paragraph_format.space_after = Pt(3)
            run = mp.add_run(module.get("title") or f"模块 {mi}")
            set_run_font(run, 11.5, True, NAVY, latin=HEAD_EA, ea=HEAD_EA)
            for chapter in module.get("chapters") or []:
                ch_no += 1
                cp = self.doc.add_paragraph()
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after = Pt(5)
                cp.paragraph_format.line_spacing = 1.0
                cp.paragraph_format.left_indent = Cm(0.7)
                add_internal_link(cp, self._chapter_title(chapter, ch_no),
                                  f"ch-{ch_no}", 10.5)
                for si, section in enumerate(chapter.get("sections") or [], 1):
                    sp = self.doc.add_paragraph()
                    sp.paragraph_format.space_before = Pt(0)
                    sp.paragraph_format.space_after = Pt(3)
                    sp.paragraph_format.line_spacing = 1.0
                    sp.paragraph_format.left_indent = Cm(1.45)
                    add_internal_link(
                        sp,
                        self._section_title(ch_no, si, section.get("title")),
                        f"sec-{ch_no}-{si}", 9.3, color=MUTED,
                    )

    # -- structure -----------------------------------------------------
    def _next_bid(self):
        bid = self.bid
        self.bid += 1
        return bid

    @staticmethod
    def _chapter_title(chapter, number):
        title = (chapter.get("title") or "").strip()
        if re.match(r"^第\s*[0-9一二三四五六七八九十百]+\s*章", title):
            return title
        return f"第 {number} 章　{title}"

    @staticmethod
    def _section_title(chapter_no, index, title):
        title = (title or "").strip()
        if re.match(r"^\d+(\.\d+)*[\s　]", title):
            return title
        return f"{chapter_no}.{index}　{title}"

    def module_page(self, index, module):
        p = self.doc.add_paragraph()
        p.paragraph_format.page_break_before = True
        p.paragraph_format.space_before = Pt(150)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(f"模块 {index}")
        set_run_font(run, 11, True, MUTED, latin=HEAD_EA, ea=HEAD_EA, spacing=60)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(module.get("title") or "")
        set_run_font(run, 26, True, NAVY, latin=HEAD_EA, ea=HEAD_EA)

        p = self.doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(16)
        add_bottom_border(p, "B7C9DE", 8, 2)

        if module.get("description"):
            body(self.doc, module["description"], size=11, after=14, line=1.6)

        chapters = module.get("chapters") or []
        for chapter in chapters:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.left_indent = Cm(0.4)
            run = p.add_run("·  " + (chapter.get("title") or ""))
            set_run_font(run, 10.5, False, MUTED)

    def chapter(self, chapter, number):
        title = self._chapter_title(chapter, number)
        h = heading(self.doc, title, 1)
        add_bookmark(h, f"ch-{number}", self._next_bid())

        objectives = chapter.get("objectives") or []
        if objectives:
            callout(self.doc, "学习目标",
                    ["• " + item for item in objectives], "goal")

        prerequisites = chapter.get("prerequisites") or []
        if prerequisites:
            p = self.doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run("先修知识：")
            set_run_font(run, 9.5, True, MUTED, latin=HEAD_EA, ea=HEAD_EA)
            run = p.add_run("　".join(prerequisites))
            set_run_font(run, 9.5, False, MUTED)

        if chapter.get("lead"):
            render_blocks(self.doc, parse_markdown(chapter["lead"], self.spec_dir),
                          self.spec_dir)

        for si, section in enumerate(chapter.get("sections") or [], 1):
            sh = heading(self.doc, self._section_title(number, si,
                                                      section.get("title")), 2)
            add_bookmark(sh, f"sec-{number}-{si}", self._next_bid())
            render_blocks(self.doc,
                          parse_markdown(section.get("body", ""), self.spec_dir),
                          self.spec_dir)

        summary = chapter.get("summary") or []
        if summary:
            callout(self.doc, "本章小结",
                    ["• " + item for item in summary], "summary")

        key_points = chapter.get("key_points") or []
        if key_points:
            callout(self.doc, "核心要点",
                    ["• " + item for item in key_points], "key")

        self.exercises(chapter, number)

    def exercises(self, chapter, number):
        ex = chapter.get("exercises") or {}
        if not isinstance(ex, dict):
            return
        levels = [("basic", "基础题"), ("intermediate", "提高题"),
                  ("challenge", "挑战题")]
        present = [(key, label) for key, label in levels if ex.get(key)]
        if not present:
            return

        heading(self.doc, "习题", 2)
        counter = 0
        for key, label in present:
            if len(present) > 1:
                heading(self.doc, label, 3)
            for item in ex.get(key) or []:
                counter += 1
                hint = ""
                if isinstance(item, dict):
                    text = item.get("q") or ""
                    hint = item.get("hint") or ""
                else:
                    text = str(item)
                p = self.doc.add_paragraph()
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(4)
                pf.line_spacing = 1.45
                pf.left_indent = Cm(1.0)
                pf.first_line_indent = Cm(-1.0)
                run = p.add_run(f"{number}-{counter}　")
                set_run_font(run, 10, True, BLUE)
                add_inline(p, text, size=10.5)
                if hint:
                    run = p.add_run(f"（提示：{hint}）")
                    set_run_font(run, 8.8, False, MUTED)

        answers = chapter.get("answers") or []
        if answers and (self.spec.get("meta") or {}).get("include_answers", True):
            heading(self.doc, "参考答案", 2)
            for i in range(counter):
                p = self.doc.add_paragraph()
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(3)
                pf.line_spacing = 1.4
                run = p.add_run(f"{number}-{i + 1}　")
                set_run_font(run, 9.5, True, MUTED)
                answer = answers[i] if i < len(answers) else "略"
                add_inline(p, answer, size=9.5, color=MUTED)

    def appendix(self):
        appendix = self.spec.get("appendix") or {}
        glossary = appendix.get("glossary") or []
        references = appendix.get("references") or []
        notes = appendix.get("notes")
        if not (glossary or references or notes):
            return

        h = heading(self.doc, "附录", 1)
        add_bookmark(h, "appendix", self._next_bid())

        if glossary:
            heading(self.doc, "术语表", 2)
            data_table(self.doc, ["术语", "含义"],
                       [[term, meaning] for term, meaning in glossary],
                       [26, 74])

        if references:
            heading(self.doc, "配套视频与延伸阅读", 2)
            body(self.doc,
                 "下列资料用于配合本讲义学习。它们只作为结构参照与延伸阅读，"
                 "具体内容请以原课程为准。", size=9.8, color=MUTED)
            for i, ref in enumerate(references, 1):
                p = self.doc.add_paragraph()
                pf = p.paragraph_format
                pf.space_before = Pt(0)
                pf.space_after = Pt(5)
                pf.line_spacing = 1.35
                run = p.add_run(f"{i}. ")
                set_run_font(run, 9.5, True, BLUE)
                run = p.add_run(ref.get("title") or "")
                set_run_font(run, 9.5, True, INK)
                if ref.get("note"):
                    run = p.add_run("　" + ref["note"])
                    set_run_font(run, 9, False, MUTED)
                if ref.get("url"):
                    run = p.add_run("\n" + ref["url"])
                    set_run_font(run, 8.5, False, BLUE, latin=MONO, ea=BODY_EA)

        if notes:
            heading(self.doc, "说明", 2)
            render_blocks(self.doc, parse_markdown(notes, self.spec_dir),
                          self.spec_dir)

    # -- run -----------------------------------------------------------
    def build(self):
        self.setup()
        self.cover()
        self.preface()
        if (self.spec.get("meta") or {}).get("include_toc", True):
            self.toc()

        modules = self.spec.get("modules") or []
        number = 0
        for mi, module in enumerate(modules, 1):
            if len(modules) > 1:
                self.module_page(mi, module)
            for chapter in module.get("chapters") or []:
                number += 1
                self.chapter(chapter, number)

        self.appendix()

        self.doc.core_properties.title = (self.spec.get("meta") or {}).get("title") or "课程讲义"
        self.doc.core_properties.author = (self.spec.get("meta") or {}).get("author") or ""
        self.doc.core_properties.subject = (self.spec.get("meta") or {}).get("subtitle") or ""
        self.doc.save(self.out_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(description="Course handout builder")
    sub = parser.add_subparsers(dest="command", required=True)

    build_p = sub.add_parser("build")
    build_p.add_argument("--spec", required=True)
    build_p.add_argument("--out", required=True)
    build_p.add_argument("--force", action="store_true")

    check_p = sub.add_parser("check")
    check_p.add_argument("--spec", required=True)

    args = parser.parse_args(argv)

    try:
        with open(args.spec, "r", encoding="utf-8") as fh:
            spec = json.load(fh)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": f"无法读取规格文件：{exc}"},
                         ensure_ascii=True))
        return 2

    spec["__path__"] = os.path.abspath(args.spec)
    issues, stats = check_spec(spec)
    errors = [i for i in issues if i["level"] == "error"]

    if args.command == "check":
        print(json.dumps({"ok": not errors, "issues": issues, "stats": stats},
                         ensure_ascii=True))
        return 0 if not errors else 2

    if errors and not args.force:
        print(json.dumps({"ok": False, "error": "规格检查未通过",
                          "issues": issues, "stats": stats},
                         ensure_ascii=True))
        return 2

    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    try:
        HandoutBuilder(spec, args.out).build()
    except Exception as exc:  # noqa: BLE001
        import traceback
        traceback.print_exc(file=sys.stderr)
        print(json.dumps({"ok": False, "error": f"生成失败：{exc}"},
                         ensure_ascii=True))
        return 3

    print(json.dumps({
        "ok": True,
        "docx": os.path.abspath(args.out),
        "stats": stats,
        "issues": issues,
    }, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    sys.exit(main())
