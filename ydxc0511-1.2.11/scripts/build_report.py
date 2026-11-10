#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按茅台数科公司版式生成会议材料或制度文本 docx。

用法：
    python build_report.py content.json -o 报告.docx

content.json 结构见 references/出Word格式整理.md。脚本只负责正文和附件版式，
不生成红头、正式文号、签发人、印章或版记，也不代替内容和权限审核。
"""

import argparse
import datetime as _dt
import json
import os
import re
import sys

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor, Twips
from docx.enum.style import WD_STYLE_TYPE


# ---------- 公司版式常量（单位：厘米/磅） ----------
PAGE_W, PAGE_H = 21.0, 29.7
MARGIN_TOP, MARGIN_BOTTOM = 3.70, 3.50
MARGIN_LEFT, MARGIN_RIGHT = 2.80, 2.60
HEADER_DIST, FOOTER_DIST = 1.50, 1.75

FONT_TITLE = "方正小标宋简体"
FONT_BODY = "仿宋_GB2312"
FONT_H1 = "黑体"
FONT_H2 = "楷体_GB2312"
FONT_TABLE = "宋体"
FONT_PAGENO = "宋体"

SIZE_TITLE = 22.0
SIZE_BODY = 16.0
SIZE_TABLE = 12.0
SIZE_PAGENO = 14.0

LINE_TITLE = 31.0
LINE_BODY = 28.95  # build() 按 profile 切换：meeting=28.95，regulation=31
LINE_REGULATION = 31.0
LINE_TABLE = 14.0

TABLE_WIDTH_CM = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT  # 15.6 cm版心
TABLE_WIDTH_TWIPS = int(round(TABLE_WIDTH_CM / 2.54 * 1440))
DEFAULT_COL_WIDTHS_CM = [1.40, 3.00, 3.20, 8.00]


# ---------- 基础工具 ----------
def set_run_font(run, name, size, bold=None):
    run.font.name = name
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(attr), name)
    rFonts.set(qn("w:cs"), "Times New Roman")


def set_style_font(style, name, size, bold=None):
    style.font.name = name
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        rFonts.set(qn(attr), name)
    rFonts.set(qn("w:cs"), "Times New Roman")
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]
    # 去掉模板自带的主题字体/字号/颜色/加粗/字距，统一为黑色常规
    for tag in ("w:b", "w:bCs", "w:color", "w:spacing"):
        el = rPr.find(qn(tag))
        if el is not None:
            rPr.remove(el)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    insert_ordered(rPr, color, RPR_ORDER)
    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        insert_ordered(rPr, szCs, RPR_ORDER)
    szCs.set(qn("w:val"), str(int(size * 2)))
    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        insert_ordered(rPr, lang, RPR_ORDER)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")


def _indent_el(pPr):
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        insert_ordered(pPr, ind, PPR_ORDER)
    return ind


def set_first_line_chars(p, chars, font_pt=SIZE_BODY):
    """首行缩进按“字符”计；同时写 twips 兜底，便于其他软件识别。"""
    ind = _indent_el(p._p.get_or_add_pPr())
    ind.set(qn("w:firstLine"), str(int(chars * font_pt * 20)))
    ind.set(qn("w:firstLineChars"), str(int(chars * 100)))
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:leftChars"), "0")
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:rightChars"), "0")
    for attr in ("w:hanging", "w:hangingChars"):
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]


def set_hanging_chars(p, left_chars=5, hanging_chars=3, font_pt=SIZE_BODY):
    ind = _indent_el(p._p.get_or_add_pPr())
    ind.set(qn("w:left"), str(int(left_chars * font_pt * 20)))
    ind.set(qn("w:leftChars"), str(int(left_chars * 100)))
    ind.set(qn("w:hanging"), str(int(hanging_chars * font_pt * 20)))
    ind.set(qn("w:hangingChars"), str(int(hanging_chars * 100)))
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:rightChars"), "0")
    for attr in ("w:firstLine", "w:firstLineChars"):
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]


def set_right_indent_chars(p, chars, font_pt=SIZE_BODY):
    ind = _indent_el(p._p.get_or_add_pPr())
    ind.set(qn("w:right"), str(int(chars * font_pt * 20)))
    ind.set(qn("w:rightChars"), str(int(chars * 100)))
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:leftChars"), "0")


def set_left_indent_chars(p, chars, font_pt=SIZE_BODY):
    ind = _indent_el(p._p.get_or_add_pPr())
    ind.set(qn("w:left"), str(int(chars * font_pt * 20)))
    ind.set(qn("w:leftChars"), str(int(chars * 100)))
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:rightChars"), "0")


def apply_para(p, line_pt=LINE_BODY, first_chars=2, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
               keep_next=False, keep_together=False):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(line_pt)
    pf.alignment = align
    pf.widow_control = False
    pf.keep_with_next = keep_next
    pf.keep_together = keep_together
    set_first_line_chars(p, first_chars)



def add_rich_text(p, text, font=FONT_BODY, size=SIZE_BODY, bold=None):
    """支持 **加粗** 行内标记；其余按普通文本处理。"""
    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            r = p.add_run(part[2:-2])
            set_run_font(r, font, size, True)
        else:
            r = p.add_run(part)
            set_run_font(r, font, size, bold)


# ---------- 样式 ----------
def build_styles(doc):
    normal = doc.styles["Normal"]
    set_style_font(normal, FONT_BODY, SIZE_BODY)
    npf = normal.paragraph_format
    npf.space_before = Pt(0)
    npf.space_after = Pt(0)
    npf.line_spacing = Pt(LINE_BODY)
    npf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    npf.widow_control = False

    body = doc.styles["Body Text"]
    body.base_style = normal
    set_style_font(body, FONT_BODY, SIZE_BODY)
    bpf = body.paragraph_format
    bpf.space_before = Pt(0)
    bpf.space_after = Pt(0)
    bpf.line_spacing = Pt(LINE_BODY)
    bpf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    bpf.widow_control = False

    title = doc.styles["Title"]
    set_style_font(title, FONT_TITLE, SIZE_TITLE)
    tpPr = title.element.get_or_add_pPr()
    pBdr = tpPr.find(qn("w:pBdr"))
    if pBdr is not None:
        tpPr.remove(pBdr)
    tpf = title.paragraph_format
    tpf.space_before = Pt(0)
    tpf.space_after = Pt(0)
    tpf.line_spacing = Pt(LINE_TITLE)
    tpf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tpf.keep_together = True

    for name, font, level in (("Heading 1", FONT_H1, 0), ("Heading 2", FONT_H2, 1)):
        st = doc.styles[name]
        set_style_font(st, font, SIZE_BODY)
        pf = st.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(LINE_BODY)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.keep_with_next = True
        pf.keep_together = True
        pf.widow_control = False
        pPr = st.element.get_or_add_pPr()
        old = pPr.find(qn("w:outlineLvl"))
        if old is not None:
            pPr.remove(old)
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), str(level))
        pPr.append(outline)

    # 链接字符样式同样会被部分渲染器套用，一并改为黑色常规
    for char_name, char_font, char_size in (
        ("Title Char", FONT_TITLE, SIZE_TITLE),
        ("Heading 1 Char", FONT_H1, SIZE_BODY),
        ("Heading 2 Char", FONT_H2, SIZE_BODY),
    ):
        try:
            set_style_font(doc.styles[char_name], char_font, char_size)
        except KeyError:
            pass
    return doc


def ensure_appendix_style(doc):
    try:
        st = doc.styles["附录标题"]
    except KeyError:
        st = doc.styles.add_style("附录标题", WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
    set_style_font(st, FONT_H1, SIZE_BODY)
    pf = st.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(LINE_BODY)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.keep_with_next = True
    return st


# ---------- 页面与页脚 ----------
def setup_page(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(PAGE_W)
    sec.page_height = Cm(PAGE_H)
    sec.top_margin = Cm(MARGIN_TOP)
    sec.bottom_margin = Cm(MARGIN_BOTTOM)
    sec.left_margin = Cm(MARGIN_LEFT)
    sec.right_margin = Cm(MARGIN_RIGHT)
    sec.header_distance = Cm(HEADER_DIST)
    sec.footer_distance = Cm(FOOTER_DIST)
    sec.different_first_page_header_footer = False
    # 文档网格随配置行距设置，保持公司版心22行口径。
    sectPr = sec._sectPr
    old = sectPr.find(qn("w:docGrid"))
    if old is not None:
        sectPr.remove(old)
    grid = OxmlElement("w:docGrid")
    grid.set(qn("w:type"), "lines")
    grid.set(qn("w:linePitch"), str(int(round(LINE_BODY * 20))))
    grid.set(qn("w:charSpace"), "0")
    sectPr.append(grid)
    return sec


def clear_paragraph(p):
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for child in list(p._p):
        if child.tag not in (qn("w:pPr"),):
            p._p.remove(child)


def setup_header_footer(doc, profile="meeting", sensitivity=""):
    settings = doc.settings.element
    old_even = settings.find(qn("w:evenAndOddHeaders"))
    if old_even is not None:
        settings.remove(old_even)
    even = OxmlElement("w:evenAndOddHeaders")
    even.set(qn("w:val"), "true")
    settings.append(even)

    sec = doc.sections[0]
    sec.different_first_page_header_footer = bool(sensitivity)
    for header in (sec.header, sec.even_page_header):
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        clear_paragraph(hp)
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(0)
        hp.paragraph_format.line_spacing = 1.0

    if sensitivity:
        header = sec.first_page_header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        clear_paragraph(hp)
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.paragraph_format.space_before = Pt(0)
        hp.paragraph_format.space_after = Pt(0)
        hp.paragraph_format.line_spacing = 1.0
        set_first_line_chars(hp, 0)
        run = hp.add_run(sensitivity)
        set_run_font(run, FONT_H1, SIZE_BODY)

    for footer, alignment in (
        (sec.footer, WD_ALIGN_PARAGRAPH.RIGHT),
        (sec.even_page_footer, WD_ALIGN_PARAGRAPH.LEFT),
    ):
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        clear_paragraph(fp)
        fp.alignment = alignment
        fp.paragraph_format.space_before = Pt(0)
        fp.paragraph_format.space_after = Pt(0)
        fp.paragraph_format.line_spacing = 1.0
        set_first_line_chars(fp, 0)
        if alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            set_right_indent_chars(fp, 1)
        else:
            set_left_indent_chars(fp, 1)

        left, right = ("— ", " —") if profile == "meeting" else ("—", "—")
        r1 = fp.add_run(left)
        set_run_font(r1, FONT_PAGENO, SIZE_PAGENO)
        add_page_field(fp)
        r2 = fp.add_run(right)
        set_run_font(r2, FONT_PAGENO, SIZE_PAGENO)


def add_page_field(p):
    run = p.add_run()
    set_run_font(run, FONT_PAGENO, SIZE_PAGENO)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE  \\* MERGEFORMAT "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    txt = OxmlElement("w:t")
    txt.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, txt, end):
        run._r.append(el)


# ---------- 编号 ----------
CN_DIGITS = "零一二三四五六七八九"


def cn_number(n):
    if n <= 0:
        return str(n)
    if n < 10:
        return CN_DIGITS[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + CN_DIGITS[n - 10]
    if n < 100:
        tens, ones = divmod(n, 10)
        return CN_DIGITS[tens] + "十" + (CN_DIGITS[ones] if ones else "")
    return str(n)


def strip_h1_number(text):
    return re.sub(r"^[一二三四五六七八九十百]+、\s*", "", text)


def strip_h2_number(text):
    return re.sub(r"^（[一二三四五六七八九十百]+）\s*", "", text)


def strip_h3_number(text):
    return re.sub(r"^\d+[.、]\s*", "", text)


def strip_h4_number(text):
    return re.sub(r"^（\d+）\s*", "", text)


def strip_chapter_number(text):
    return re.sub(r"^第[一二三四五六七八九十百\d]+章\s*", "", text)


def strip_section_number(text):
    return re.sub(r"^第[一二三四五六七八九十百\d]+节\s*", "", text)


def strip_article_number(text):
    return re.sub(r"^第[一二三四五六七八九十百\d]+条\s*", "", text)


# ---------- 正文 ----------
def add_title(doc, title):
    lines = title if isinstance(title, list) else str(title).split("\n")
    lines = [x.strip() for x in lines if x and x.strip()]
    if not lines:
        return
    for i, line in enumerate(lines):
        p = doc.add_paragraph(style="Title")
        apply_para(p, LINE_TITLE, 0, WD_ALIGN_PARAGRAPH.CENTER,
                   keep_next=(i < len(lines) - 1), keep_together=True)
        r = p.add_run(line)
        set_run_font(r, FONT_TITLE, SIZE_TITLE)


def add_recipient(doc, recipient):
    if not recipient:
        return
    text = recipient.strip()
    if not text.endswith(("：", ":")):
        text += "："
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.JUSTIFY)
    r = p.add_run(text)
    set_run_font(r, FONT_BODY, SIZE_BODY)


def add_h1(doc, text, number):
    p = doc.add_paragraph(style="Heading 1")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY, keep_next=True, keep_together=True)
    r = p.add_run(f"{cn_number(number)}、{text}")
    set_run_font(r, FONT_H1, SIZE_BODY)


def add_h2(doc, text, number):
    p = doc.add_paragraph(style="Heading 2")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY, keep_next=True, keep_together=True)
    r = p.add_run(f"（{cn_number(number)}）{text}")
    set_run_font(r, FONT_H2, SIZE_BODY)


def split_h3(text):
    m = re.match(r"^(.+?[。：:])(.*)$", text, re.S)
    if m and m.group(2).strip():
        return m.group(1), m.group(2)
    return text, ""


def add_h3(doc, text, number):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY, keep_next=True, keep_together=True)
    r = p.add_run(f"{number}.{text}")
    set_run_font(r, FONT_BODY, SIZE_BODY)


def add_h4(doc, text, number):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY)
    r = p.add_run(f"（{number}）{text}")
    set_run_font(r, FONT_BODY, SIZE_BODY)


def chapter_display_name(text):
    compact = re.sub(r"\s+", "", text)
    if compact in ("总则", "附则"):
        return f"{compact[0]}　　{compact[1]}"
    return text.strip()


def add_chapter(doc, text, number):
    p = doc.add_paragraph(style="Heading 1")
    apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.CENTER, keep_next=True, keep_together=True)
    r = p.add_run(f"第{cn_number(number)}章　　{chapter_display_name(text)}")
    set_run_font(r, FONT_H1, SIZE_BODY)


def add_section(doc, text, number):
    p = doc.add_paragraph(style="Heading 2")
    apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.CENTER, keep_next=True, keep_together=True)
    r = p.add_run(f"第{cn_number(number)}节　　{text.strip()}")
    set_run_font(r, FONT_H2, SIZE_BODY)


def add_article(doc, text, number):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY)
    r = p.add_run(f"第{cn_number(number)}条　　")
    set_run_font(r, FONT_H1, SIZE_BODY)
    add_rich_text(p, text, FONT_BODY, SIZE_BODY)


def add_body(doc, text):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_rich_text(p, text, FONT_BODY, SIZE_BODY)


def add_blank(doc):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY)
    return p


# ---------- 表格 ----------
PPR_ORDER = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
    "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
]

RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
    "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
    "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
    "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
    "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
]

TCPR_ORDER = [
    "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
    "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge",
    "w:tcPrChange",
]

TBLPR_ORDER = [
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
    "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
    "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
    "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
]


def insert_ordered(parent, child, order):
    tag = child.tag.split("}")[1]
    key = "w:" + tag
    if key not in order:
        parent.append(child)
        return
    idx = order.index(key)
    for existing in list(parent):
        etag = existing.tag.split("}")[1]
        ekey = "w:" + etag
        if ekey in order and order.index(ekey) > idx:
            existing.addprevious(child)
            return
    parent.append(child)


def set_table_borders_none(table):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "none")
        el.set(qn("w:color"), "auto")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        borders.append(el)
    insert_ordered(tblPr, borders, TBLPR_ORDER)


def set_table_cell_margins(table):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblCellMar"))
    if old is not None:
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side, w in (("top", 0), ("left", 108), ("bottom", 0), ("right", 108)):
        el = OxmlElement("w:" + side)
        el.set(qn("w:w"), str(w))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    insert_ordered(tblPr, mar, TBLPR_ORDER)


def set_cell_border(cell, sz=4, color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None:
        tcPr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    insert_ordered(tcPr, borders, TCPR_ORDER)


def style_cell(cell, text, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    set_first_line_chars(p, 0)
    r = p.add_run(str(text))
    set_run_font(r, FONT_TABLE, SIZE_TABLE, bold)
    set_cell_border(cell)


def normalize_widths(widths, ncols):
    if not widths:
        if ncols == 4:
            return DEFAULT_COL_WIDTHS_CM[:]
        # 默认：首列窄，末列宽，其余均分，总宽不超过版心。
        if ncols == 1:
            return [TABLE_WIDTH_CM]
        first, last = 1.4, 4.0
        rest = max(0.1, (TABLE_WIDTH_CM - first - last) / max(1, ncols - 2))
        return [first] + [rest] * (ncols - 2) + [last]
    if len(widths) != ncols:
        raise ValueError(f"widths_cm 数量 {len(widths)} 与列数 {ncols} 不一致")
    total = sum(widths)
    if total > TABLE_WIDTH_CM:
        widths = [w * TABLE_WIDTH_CM / total for w in widths]
    return widths


def add_table(doc, spec):
    columns = spec["columns"]
    rows = spec.get("rows", [])
    ncols = len(columns)
    nrows = 1 + len(rows) + (1 if spec.get("total") else 0)
    widths = normalize_widths(spec.get("widths_cm"), ncols)

    table = doc.add_table(rows=nrows, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"

    tblPr = table._tbl.tblPr
    old_w = tblPr.find(qn("w:tblW"))
    if old_w is not None:
        tblPr.remove(old_w)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(TABLE_WIDTH_TWIPS))
    tblW.set(qn("w:type"), "dxa")
    insert_ordered(tblPr, tblW, TBLPR_ORDER)

    old_layout = tblPr.find(qn("w:tblLayout"))
    if old_layout is not None:
        tblPr.remove(old_layout)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    insert_ordered(tblPr, layout, TBLPR_ORDER)

    set_table_borders_none(table)
    set_table_cell_margins(table)

    # 网格列宽
    grid_cols = table._tbl.tblGrid.findall(qn("w:gridCol"))
    for gc, w in zip(grid_cols, widths):
        gc.set(qn("w:w"), str(int(w / 2.54 * 1440)))

    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Twips(int(widths[i] / 2.54 * 1440))

    # 表头
    header = table.rows[0]
    header.height = Cm(1.0)
    header.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    trPr = header._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    trPr.append(th)
    for i, name in enumerate(columns):
        style_cell(header.cells[i], name, bold=True,
                   align=WD_ALIGN_PARAGRAPH.CENTER)

    # 数据行
    for ri, row_data in enumerate(rows, start=1):
        row = table.rows[ri]
        row.height = Cm(0.75)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for ci in range(ncols):
            val = row_data[ci] if ci < len(row_data) else ""
            align = WD_ALIGN_PARAGRAPH.LEFT if ci == ncols - 1 else WD_ALIGN_PARAGRAPH.CENTER
            style_cell(row.cells[ci], val, bold=False, align=align)

    # 合计行
    if spec.get("total"):
        ri = nrows - 1
        row = table.rows[ri]
        row.height = Cm(0.8)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        total = spec["total"]
        if len(total) == 2 and ncols >= 2:
            half = ncols // 2
            groups = [(0, half), (half, ncols)]
            for (start, end), val in zip(groups, total):
                cell = table.cell(ri, start)
                if end - start > 1:
                    cell = cell.merge(table.cell(ri, end - 1))
                style_cell(cell, val, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            for ci in range(ncols):
                val = total[ci] if ci < len(total) else ""
                style_cell(row.cells[ci], val, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    # 自定义合并（可选）
    for m in spec.get("merges", []):
        r = m["row"]
        start, end = m["start"], m["end"]
        cell = table.cell(r, start)
        if end - start > 1:
            cell = cell.merge(table.cell(r, end - 1))
        if m.get("text") is not None:
            style_cell(cell, m["text"], bold=m.get("bold", True), align=WD_ALIGN_PARAGRAPH.CENTER)

    # 表格整体居中：加一个空段落会破坏间距，直接依赖 jc=center
    return table


def add_body_table(doc, spec):
    # 正文表格前后各留半行，避免与文字粘连
    add_table(doc, spec)


# ---------- 附件 ----------
def add_attachment_note(doc, attachments):
    if not attachments:
        return
    first = True
    for i, att in enumerate(attachments, start=1):
        p = doc.add_paragraph(style="Body Text")
        apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY)
        set_hanging_chars(p, left_chars=5, hanging_chars=3)
        prefix = "附件：" if first else ""
        text = f"{prefix}{i}.{att.get('title', '')}"
        r = p.add_run(text)
        set_run_font(r, FONT_BODY, SIZE_BODY)
        first = False


def add_attachment(doc, att, index):
    marker = doc.add_paragraph(style="附录标题")
    apply_para(marker, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.LEFT,
               keep_next=True, keep_together=True)
    marker.paragraph_format.page_break_before = True
    r = marker.add_run(f"附件{index}")
    set_run_font(r, FONT_H1, SIZE_BODY)

    if att.get("title"):
        tp = doc.add_paragraph(style="Title")
        apply_para(tp, LINE_TITLE, 0, WD_ALIGN_PARAGRAPH.CENTER,
                   keep_next=True, keep_together=True)
        tr = tp.add_run(att["title"])
        set_run_font(tr, FONT_TITLE, SIZE_TITLE)

    images = att.get("images") or []
    for ii, img in enumerate(images):
        p = doc.add_paragraph()
        apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.CENTER, keep_together=True)
        if ii > 0:
            p.paragraph_format.page_break_before = True
        run = p.add_run()
        width = img.get("width_cm") if isinstance(img, dict) else None
        path = img.get("path") if isinstance(img, dict) else img
        if not path or not os.path.exists(path):
            raise FileNotFoundError(f"附件图片不存在：{path}")
        if width:
            run.add_picture(path, width=Cm(width))
        else:
            shape = run.add_picture(path)
            max_w, max_h = Cm(TABLE_WIDTH_CM), Cm(18.0)
            if shape.width > max_w:
                ratio = max_w / shape.width
                shape.width = max_w
                shape.height = int(shape.height * ratio)
            if shape.height > max_h:
                ratio = max_h / shape.height
                shape.height = max_h
                shape.width = int(shape.width * ratio)
        # 图片段落固定行距会裁切，改为单倍行距
        p.paragraph_format.line_spacing = 1.0

    if att.get("table"):
        add_table(doc, att["table"])


# ---------- 落款 ----------
def estimate_width(text):
    """粗略估算文字宽度（以 16pt 中文字符宽为 1）。"""
    w = 0.0
    prev_cjk = None
    for ch in text:
        is_cjk = ord(ch) > 127
        w += 1.0 if is_cjk else 0.635
        if prev_cjk is not None and prev_cjk != is_cjk:
            w += 0.25
        prev_cjk = is_cjk
    return w


def add_signature(doc, signer, date_text):
    if not signer and not date_text:
        return
    add_blank(doc)
    add_blank(doc)
    if signer:
        p = doc.add_paragraph(style="Body Text")
        apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.RIGHT)
        extra = 0.0
        if date_text:
            diff = estimate_width(date_text) - estimate_width(signer)
            extra = max(0.0, diff / 2.0)
        set_right_indent_chars(p, 4 + extra)
        r = p.add_run(signer)
        set_run_font(r, FONT_BODY, SIZE_BODY)
    if date_text:
        p = doc.add_paragraph(style="Body Text")
        apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.RIGHT)
        set_right_indent_chars(p, 4)
        r = p.add_run(date_text)
        set_run_font(r, FONT_BODY, SIZE_BODY)


# ---------- 主流程 ----------
def build(data, out_path):
    global LINE_BODY
    profile = data.get("profile", "meeting")
    if profile not in ("meeting", "regulation"):
        raise ValueError("profile 仅支持 meeting 或 regulation")
    LINE_BODY = LINE_REGULATION if profile == "regulation" else 28.95

    doc = Document()
    setup_page(doc)
    build_styles(doc)
    ensure_appendix_style(doc)
    setup_header_footer(doc, profile, data.get("sensitivity", ""))

    body = doc.element.body
    for p in list(body.findall(qn("w:p"))):
        body.remove(p)

    title = data.get("title", "")
    add_title(doc, title)
    add_blank(doc)
    add_recipient(doc, data.get("recipient"))

    h1 = h2 = h3 = h4 = 0
    chapter = section = article = 0
    numbering = data.get("numbering", "auto")
    for block in data.get("body", []):
        if isinstance(block, str):
            block = {"type": "p", "text": block}
        typ = block.get("type", "p")
        text = (block.get("text") or "").strip()
        if typ == "h1":
            h1 += 1
            h2 = h3 = h4 = 0
            if numbering == "auto":
                add_h1(doc, strip_h1_number(text), h1)
            else:
                add_h1_manual(doc, text)
        elif typ == "h2":
            h2 += 1
            h3 = h4 = 0
            if numbering == "auto":
                add_h2(doc, strip_h2_number(text), h2)
            else:
                add_h2_manual(doc, text)
        elif typ == "h3":
            h3 += 1
            h4 = 0
            if numbering == "auto":
                add_h3(doc, strip_h3_number(text), h3)
            else:
                add_h3_manual(doc, text)
        elif typ == "h4":
            h4 += 1
            if numbering == "auto":
                add_h4(doc, strip_h4_number(text), h4)
            else:
                add_h4_manual(doc, text)
        elif typ == "chapter":
            chapter += 1
            section = 0
            if numbering == "auto":
                add_chapter(doc, strip_chapter_number(text), chapter)
            else:
                add_chapter_manual(doc, text)
        elif typ == "section":
            section += 1
            if numbering == "auto":
                add_section(doc, strip_section_number(text), section)
            else:
                add_section_manual(doc, text)
        elif typ == "article":
            article += 1
            if numbering == "auto":
                add_article(doc, strip_article_number(text), article)
            else:
                add_article_manual(doc, text)
        elif typ in ("p", "body"):
            add_body(doc, text)
        elif typ == "blank":
            add_blank(doc)
        elif typ == "table":
            add_body_table(doc, block["table"])
        elif typ == "pagebreak":
            p = doc.add_paragraph()
            p.add_run().add_break(WD_BREAK.PAGE)
        else:
            raise ValueError(f"未知正文块类型：{typ}")

    attachments = data.get("attachments") or []
    if attachments:
        add_blank(doc)
        add_attachment_note(doc, attachments)

    date_text = data.get("date")
    if date_text == "auto":
        today = _dt.date.today()
        date_text = f"{today.year}年{today.month}月{today.day}日"
    if data.get("signer") or date_text:
        add_signature(doc, data.get("signer"), date_text)

    for i, att in enumerate(attachments, start=1):
        add_attachment(doc, att, i)

    settings = doc.settings.element
    old_update = settings.find(qn("w:updateFields"))
    if old_update is not None:
        settings.remove(old_update)
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    settings.append(update)

    cp = doc.core_properties
    cp.title = "".join(title) if isinstance(title, list) else str(title)
    cp.author = ""
    cp.last_modified_by = ""
    cp.comments = ""
    cp.keywords = profile

    doc.save(out_path)
    return out_path


def add_h1_manual(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY, keep_next=True, keep_together=True)
    r = p.add_run(text)
    set_run_font(r, FONT_H1, SIZE_BODY)


def add_h2_manual(doc, text):
    p = doc.add_paragraph(style="Heading 2")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY, keep_next=True, keep_together=True)
    r = p.add_run(text)
    set_run_font(r, FONT_H2, SIZE_BODY)


def add_h3_manual(doc, text):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY, keep_next=True, keep_together=True)
    r = p.add_run(text)
    set_run_font(r, FONT_BODY, SIZE_BODY)


def add_h4_manual(doc, text):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY)
    r = p.add_run(text)
    set_run_font(r, FONT_BODY, SIZE_BODY)


def add_chapter_manual(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.CENTER, keep_next=True, keep_together=True)
    r = p.add_run(text)
    set_run_font(r, FONT_H1, SIZE_BODY)


def add_section_manual(doc, text):
    p = doc.add_paragraph(style="Heading 2")
    apply_para(p, LINE_BODY, 0, WD_ALIGN_PARAGRAPH.CENTER, keep_next=True, keep_together=True)
    r = p.add_run(text)
    set_run_font(r, FONT_H2, SIZE_BODY)


def add_article_manual(doc, text):
    p = doc.add_paragraph(style="Body Text")
    apply_para(p, LINE_BODY, 2, WD_ALIGN_PARAGRAPH.JUSTIFY)
    match = re.match(r"^(第[一二三四五六七八九十百\d]+条)(　　|\s+)?(.*)$", text, re.S)
    if match:
        r = p.add_run(match.group(1) + "　　")
        set_run_font(r, FONT_H1, SIZE_BODY)
        add_rich_text(p, match.group(3), FONT_BODY, SIZE_BODY)
    else:
        add_rich_text(p, text, FONT_BODY, SIZE_BODY)


def main():
    ap = argparse.ArgumentParser(description="按公文报告标准版式生成 docx")
    ap.add_argument("json_file", help="内容 JSON 路径")
    ap.add_argument("-o", "--output", required=True, help="输出 docx 路径")
    args = ap.parse_args()
    with open(args.json_file, encoding="utf-8") as f:
        data = json.load(f)
    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    build(data, args.output)
    print(f"已生成：{args.output}")


if __name__ == "__main__":
    main()
