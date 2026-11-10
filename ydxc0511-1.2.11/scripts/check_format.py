#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""检查 docx 是否符合茅台数科公司公文/制度版式。

用法：python check_format.py 报告.docx
"""

import argparse
import re
import sys

from docx import Document
from docx.oxml.ns import qn

TITLE_FONTS = {"方正小标宋简体", "国标小标宋-GB/T 2312", "方正小标宋_GBK", "方正小标宋简体_GBK"}
FONT_BODY = "仿宋_GB2312"
FONT_H1 = "黑体"
FONT_H2 = "楷体_GB2312"
FONT_TABLE = "宋体"
TABLE_WIDTH_TWIPS = str(int(round(15.6 / 2.54 * 1440)))
CN_DIGITS = "零一二三四五六七八九"


def cn_number(n):
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


def emu_to_pt(v):
    return None if v is None else v / 12700.0


def style_eastasia(style):
    rPr = style.element.rPr
    if rPr is None or rPr.rFonts is None:
        return None
    return rPr.rFonts.get(qn("w:eastAsia"))


def run_eastasia(run):
    rPr = run._element.rPr
    if rPr is None or rPr.rFonts is None:
        return None
    return rPr.rFonts.get(qn("w:eastAsia"))


def ind_attr(p, name):
    pPr = p._p.pPr
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    return None if ind is None else ind.get(qn(name))


def p_bdr(style):
    pPr = style.element.pPr
    return None if pPr is None else pPr.find(qn("w:pBdr"))


class Checker:
    def __init__(self, path):
        self.doc = Document(path)
        profile = (self.doc.core_properties.keywords or "").strip()
        self.profile = profile if profile in ("meeting", "regulation") else "meeting"
        self.body_line = 31.0 if self.profile == "regulation" else 28.95
        self.pass_n = 0
        self.warn_n = 0
        self.fail_n = 0

    def ok(self, msg):
        self.pass_n += 1
        print(f"[PASS] {msg}")

    def warn(self, msg):
        self.warn_n += 1
        print(f"[WARN] {msg}")

    def fail(self, msg):
        self.fail_n += 1
        print(f"[FAIL] {msg}")

    def near(self, a, b, tol=0.05):
        return a is not None and abs(a - b) <= tol

    def check_page(self):
        s = self.doc.sections[0]
        vals = [
            ("纸张 A4", s.page_width.cm, 21.0, 0.1),
            ("纸张高度 A4", s.page_height.cm, 29.7, 0.1),
            ("上边距 3.7cm", s.top_margin.cm, 3.7, 0.05),
            ("下边距 3.5cm", s.bottom_margin.cm, 3.5, 0.05),
            ("左边距 2.8cm", s.left_margin.cm, 2.8, 0.05),
            ("右边距 2.6cm", s.right_margin.cm, 2.6, 0.05),
            ("页眉距 1.5cm", s.header_distance.cm, 1.5, 0.05),
            ("页脚距 1.75cm", s.footer_distance.cm, 1.75, 0.05),
        ]
        for name, value, expect, tol in vals:
            if self.near(value, expect, tol):
                self.ok(f"{name}（{value:.2f}cm）")
            else:
                self.fail(f"{name}：实际 {value:.2f}cm，应为 {expect}cm")

    def check_footer(self):
        if "evenAndOddHeaders" not in self.doc.settings.element.xml:
            self.fail("未启用奇偶页不同页脚")
        else:
            self.ok("已启用奇偶页不同页脚")
        sec = self.doc.sections[0]
        expected_space = self.profile == "meeting"
        for label, footer, align, side in (
            ("奇数页", sec.footer, 2, "right"),
            ("偶数页", sec.even_page_footer, 0, "left"),
        ):
            text = "".join(p.text for p in footer.paragraphs)
            has_field = "PAGE" in footer._element.xml
            if not footer.paragraphs:
                self.fail(f"{label}页脚为空")
                continue
            p = footer.paragraphs[0]
            problems = []
            if not has_field or "—" not in text:
                problems.append("缺PAGE域或一字线")
            if expected_space and not text.startswith("— "):
                problems.append("应为“— 1 —”样式")
            if not expected_space and text.startswith("— "):
                problems.append("制度页码应为“—1—”样式")
            if p.alignment is None or int(p.alignment) != align:
                problems.append("奇右偶左位置错误")
            if ind_attr(p, f"w:{side}Chars") != "100":
                problems.append("未空一字")
            if problems:
                self.fail(f"{label}页码异常：" + "；".join(problems))
            else:
                self.ok(f"{label}页码（{text.strip()}，位置及空字正确）")

    def check_title(self):
        paras = []
        for p in self.doc.paragraphs:
            if p.style.name == "Title":
                paras.append(p)
            elif paras:
                break
        if not paras:
            self.fail("未找到 Title 样式标题")
            return
        bad = []
        for p in paras:
            size = p.runs[0].font.size.pt if p.runs and p.runs[0].font.size else (
                p.style.font.size.pt if p.style.font.size else None)
            line = emu_to_pt(p.paragraph_format.line_spacing)
            if size != 22.0:
                bad.append(f"字号 {size}")
            if p.alignment is None or int(p.alignment) != 1:
                bad.append("未居中")
            if line is None or abs(line - 31.0) > 0.2:
                bad.append(f"行距 {line}")
            if p_bdr(self.doc.styles["Title"]) is not None:
                bad.append("Title 样式仍有边框")
            for r in p.runs:
                if not r.text.strip():
                    continue
                ea = run_eastasia(r) or style_eastasia(p.style)
                if ea not in TITLE_FONTS:
                    bad.append(f"标题字体 {ea}")
        if bad:
            self.fail("标题格式异常：" + "；".join(sorted(set(bad))))
        else:
            self.ok(f"标题 {len(paras)} 行，二号小标宋、居中、行距31磅、无边框")

    def check_style(self, name, font, size):
        st = self.doc.styles[name]
        ea = style_eastasia(st)
        sz = st.font.size.pt if st.font.size else None
        color = st.font.color.rgb if st.font.color and st.font.color.rgb else None
        problems = []
        expect_fonts = {font} if isinstance(font, str) else set(font)
        if ea not in expect_fonts:
            problems.append(f"中文字体 {ea}")
        if sz != size:
            problems.append(f"字号 {sz}")
        if color is not None and str(color) != "000000":
            problems.append(f"颜色 {color}")
        if problems:
            self.fail(f"{name} 样式异常：" + "、".join(problems))
        else:
            self.ok(f"{name} 样式（{ea} {size}pt，黑色）")

    def check_body_paragraphs(self):
        total = 0
        bad = []
        seen_heading = False
        allowed_fonts = {FONT_BODY, FONT_H1} if self.profile == "regulation" else {FONT_BODY}
        for p in self.doc.paragraphs:
            if p.style.name in ("Heading 1", "Heading 2"):
                seen_heading = True
            if p.style.name != "Body Text" or not p.text.strip():
                continue
            total += 1
            line = emu_to_pt(p.paragraph_format.line_spacing)
            if line is None or abs(line - self.body_line) > 0.2:
                bad.append(f"行距 {line}")
            if p.alignment is not None and int(p.alignment) in (1, 2):
                continue
            hanging = ind_attr(p, "w:hangingChars")
            first_chars = ind_attr(p, "w:firstLineChars")
            if hanging == "300" or first_chars == "500":
                continue
            if not seen_heading and first_chars == "0" and p.text.strip().endswith("："):
                continue
            if first_chars != "200":
                bad.append(f"首行缩进 {first_chars}")
            for r in p.runs:
                if not r.text.strip():
                    continue
                ea = run_eastasia(r) or FONT_BODY
                if ea not in allowed_fonts:
                    bad.append(f"正文字体 {ea}")
                if r.font.size and r.font.size.pt != 16.0:
                    bad.append(f"正文字号 {r.font.size.pt}")
        if bad:
            self.fail(f"正文段落 {total} 段存在异常：" + "；".join(sorted(set(bad))[:8]))
        else:
            self.ok(f"正文 {total} 段，行距 {self.body_line} 磅、首行缩进2字符、三号字体")

    def check_heading_paragraphs(self):
        h1 = [p for p in self.doc.paragraphs if p.style.name == "Heading 1"]
        h2 = [p for p in self.doc.paragraphs if p.style.name == "Heading 2"]
        problems = []
        for p in h1:
            for r in p.runs:
                if (run_eastasia(r) or style_eastasia(p.style)) != FONT_H1:
                    problems.append("一级标题字体")
            if self.profile == "regulation" and (p.alignment is None or int(p.alignment) != 1):
                problems.append("章标题未居中")
        for p in h2:
            for r in p.runs:
                if (run_eastasia(r) or style_eastasia(p.style)) != FONT_H2:
                    problems.append("二级标题字体")
            if self.profile == "regulation" and (p.alignment is None or int(p.alignment) != 1):
                problems.append("节标题未居中")
        if problems:
            self.fail("标题异常：" + "、".join(sorted(set(problems))))
        else:
            self.ok(f"一级标题 {len(h1)} 个（黑体）、二级标题 {len(h2)} 个（楷体）")

    def check_regulation(self):
        if self.profile != "regulation":
            return
        articles = []
        for p in self.doc.paragraphs:
            match = re.match(r"^(第[一二三四五六七八九十百\d]+条)", p.text.strip())
            if match and p.style.name == "Body Text":
                articles.append((match.group(1), p))
        if not articles:
            self.fail("regulation 配置下未识别到制度条文")
            return
        problems = []
        for i, (actual, p) in enumerate(articles, start=1):
            expected = f"第{cn_number(i)}条"
            if actual != expected:
                problems.append(f"第{i}个条号为{actual}，应为{expected}")
            if not p.runs or run_eastasia(p.runs[0]) != FONT_H1:
                problems.append(f"{actual}条号未用黑体")
        if problems:
            self.fail("制度条文异常：" + "；".join(problems[:8]))
        else:
            self.ok(f"制度条文 {len(articles)} 条，全文连续编号、条号黑体")

    def check_tables(self):
        if not self.doc.tables:
            self.ok("无表格")
            return
        problems = []
        for ti, table in enumerate(self.doc.tables, start=1):
            tblW = table._tbl.tblPr.find(qn("w:tblW"))
            w = tblW.get(qn("w:w")) if tblW is not None else None
            if w != TABLE_WIDTH_TWIPS:
                problems.append(f"表{ti} 宽度 {w}")
            run = None
            for p in table.rows[0].cells[0].paragraphs:
                if p.runs:
                    run = p.runs[0]
                    break
            if run is None:
                problems.append(f"表{ti} 表头无文字")
            else:
                if run.font.size is None or run.font.size.pt != 12.0:
                    problems.append(f"表{ti} 字号 {run.font.size.pt if run.font.size else None}")
                if run_eastasia(run) != FONT_TABLE:
                    problems.append(f"表{ti} 字体 {run_eastasia(run)}")
                if not run.font.bold:
                    problems.append(f"表{ti} 表头未加粗")
            tcPr = table.rows[0].cells[0]._tc.tcPr
            borders = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
            if borders is None or borders.find(qn("w:top")) is None:
                problems.append(f"表{ti} 缺框线")
        if problems:
            self.fail("表格异常：" + "；".join(problems))
        else:
            self.ok(f"表格 {len(self.doc.tables)} 个，版心宽、宋体小四、表头加粗、有框线")

    def run(self):
        print(f"=== 公司公文格式检查（profile={self.profile}）===")
        self.check_page()
        self.check_footer()
        self.check_style("Title", TITLE_FONTS, 22.0)
        self.check_style("Heading 1", FONT_H1, 16.0)
        self.check_style("Heading 2", FONT_H2, 16.0)
        self.check_style("Body Text", FONT_BODY, 16.0)
        self.check_title()
        self.check_heading_paragraphs()
        self.check_body_paragraphs()
        self.check_regulation()
        self.check_tables()
        print(f"\n汇总：PASS {self.pass_n}，WARN {self.warn_n}，FAIL {self.fail_n}")
        return 0 if self.fail_n == 0 else 1


def main():
    ap = argparse.ArgumentParser(description="检查茅台数科公司公文/制度版式")
    ap.add_argument("docx_file")
    args = ap.parse_args()
    sys.exit(Checker(args.docx_file).run())


if __name__ == "__main__":
    main()
