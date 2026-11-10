#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""提取 docx 结构，供改排前盘点内容。

用法：
    python extract_docx.py 来源.docx -o work/media
"""

import argparse
import os
import zipfile

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def emu_to_pt(v):
    return None if v is None else round(v / 12700.0, 2)


def ind_chars(p, name):
    pPr = p._p.pPr
    if pPr is None:
        return None
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return None
    return ind.get(qn(name))


def run_info(r, style=None):
    ea = None
    rPr = r._element.rPr
    if rPr is not None and rPr.rFonts is not None:
        ea = rPr.rFonts.get(qn("w:eastAsia"))
    size = r.font.size.pt if r.font.size else None
    if ea is None and style is not None:
        spPr = style.element.rPr
        if spPr is not None and spPr.rFonts is not None:
            ea = spPr.rFonts.get(qn("w:eastAsia"))
    if size is None and style is not None and style.font.size:
        size = style.font.size.pt
    return {
        "text": r.text,
        "font": ea,
        "size": size,
        "bold": r.font.bold,
    }


def paragraph_info(p):
    pf = p.paragraph_format
    return {
        "style": p.style.name,
        "align": str(p.alignment) if p.alignment is not None else None,
        "first_line_chars": ind_chars(p, "w:firstLineChars"),
        "hanging_chars": ind_chars(p, "w:hangingChars"),
        "line_pt": emu_to_pt(pf.line_spacing),
        "text": p.text,
        "runs": [run_info(r, p.style) for r in p.runs if r.text],
    }


def main():
    ap = argparse.ArgumentParser(description="提取 docx 段落、表格、图片")
    ap.add_argument("docx_file")
    ap.add_argument("-o", "--outdir", default="work/media", help="图片输出目录")
    args = ap.parse_args()

    doc = Document(args.docx_file)
    s = doc.sections[0]
    print("=== 页面 ===")
    print(f"纸张 {s.page_width.cm:.2f} x {s.page_height.cm:.2f} cm")
    print(f"页边距 上{s.top_margin.cm:.2f} 下{s.bottom_margin.cm:.2f} "
          f"左{s.left_margin.cm:.2f} 右{s.right_margin.cm:.2f} cm")

    print("\n=== 正文结构 ===")
    idx = 0
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            p = Paragraph(child, doc)
            if not p.text.strip():
                continue
            info = paragraph_info(p)
            print(f"\n[{idx}] P style={info['style']} align={info['align']} "
                  f"first={info['first_line_chars']} hanging={info['hanging_chars']} "
                  f"line={info['line_pt']}pt")
            print(f"    {info['text'][:160]}")
            for r in info["runs"][:4]:
                print(f"    run: {r['text'][:40]!r} font={r['font']} "
                      f"size={r['size']} bold={r['bold']}")
            idx += 1
        elif tag == "tbl":
            t = Table(child, doc)
            print(f"\n[{idx}] TABLE {len(t.rows)}行 x {len(t.columns)}列")
            for row in t.rows:
                cells = [c.text.replace("\n", " ")[:30] for c in row.cells]
                print("    | " + " | ".join(cells))
            idx += 1

    # 导出图片
    os.makedirs(args.outdir, exist_ok=True)
    count = 0
    with zipfile.ZipFile(args.docx_file) as z:
        for name in z.namelist():
            if name.startswith("word/media/"):
                out = os.path.join(args.outdir, os.path.basename(name))
                with open(out, "wb") as f:
                    f.write(z.read(name))
                print(f"\n图片：{out}")
                count += 1
    print(f"\n共 {idx} 个正文块，导出 {count} 张图片到 {args.outdir}")


if __name__ == "__main__":
    main()
