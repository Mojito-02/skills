# -*- coding: utf-8 -*-
"""把 问题一论文正文_可直接插入.md 转换为排版规范的 Word (.docx)。

- 标题层级 -> Heading 1 / 2
- 正文 -> 首行缩进2字符、1.5倍行距、宋体小四
- 行内/独立公式 -> 原生 OMML（LaTeX -> MathML -> OMML，经 Word 的 MML2OMML.XSL）
- Markdown 表格 -> 真实 Word 表格（带边框、表头加粗、按对齐标记设置列对齐）
- 表题 / 图注 -> 居中
"""
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree
import latex2mathml.converter as l2m

SRC = Path(sys.argv[1])
OUT = Path(sys.argv[2])
XSL = r"C:/Program Files/Microsoft Office/root/Office16/MML2OMML.XSL"

CN_BODY = "宋体"
CN_HEAD = "黑体"
LATIN = "Times New Roman"
MONO = "Consolas"

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M = f"{{{M_NS}}}"

_ts = etree.XSLT(etree.parse(XSL))


# ---------------------------------------------------------------- 字体 / 样式
def set_run_font(run, cn=CN_BODY, latin=LATIN, size=None, bold=None, mono=False):
    if mono:
        latin = MONO
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), cn)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def style_font(style, cn, latin=LATIN, size=None, bold=None):
    style.font.name = latin
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), cn)
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold


def configure_styles(doc):
    st = doc.styles
    style_font(st["Normal"], CN_BODY, LATIN, 12)          # 小四
    pf = st["Normal"].paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)

    # 一、二、三级标题：中文黑体
    style_font(st["Heading 1"], CN_HEAD, LATIN, 16, True)
    style_font(st["Heading 2"], CN_HEAD, LATIN, 14, True)
    style_font(st["Heading 3"], CN_HEAD, LATIN, 12, True)
    for name, before, after in (("Heading 1", 12, 12), ("Heading 2", 10, 8), ("Heading 3", 8, 6)):
        p = st[name].paragraph_format
        p.space_before = Pt(before)
        p.space_after = Pt(after)
        p.line_spacing = 1.5
        p.first_line_indent = Pt(0)
        st[name].font.color.rgb = RGBColor(0, 0, 0)

    # 表格网格样式
    tgs = st["Table Grid"]
    style_font(tgs, CN_BODY, LATIN, 10.5)                  # 五号


# ---------------------------------------------------------------- 公式
def latex_to_omml(latex: str):
    """LaTeX -> OMML。失败则做转义重试，仍失败返回 None。"""
    cands = [latex]
    esc = latex.replace("<", r"\lt ").replace(">", r"\gt ")
    if esc != latex:
        cands.append(esc)
    for c in cands:
        try:
            mml = l2m.convert(c)
            node = _ts(etree.fromstring(mml.encode("utf-8")))
            s = etree.tostring(node, encoding="unicode")
            if "<m:oMath" in s:
                return s[s.index("<m:oMath"):]
        except Exception:
            continue
    return None


def plain_from_texttt(latex: str) -> str | None:
    """若整个公式就是 \\texttt{...}，还原为普通等宽文本（代码/路径/文件名）。

    这类内容是标识符而非数学符号，用数学公式渲染会得到错误的字形与转义，
    因此统一按等宽正文输出。返回 None 表示不是纯 \\texttt。
    """
    s = latex.strip()
    m = re.fullmatch(r"\\texttt\{(.*)\}", s, re.S)
    if not m:
        return None
    t = m.group(1)
    t = t.replace(r"\_", "_").replace(r"\$", "$").replace(r"\&", "&")
    t = t.replace(r"\{", "{").replace(r"\}", "}").replace(r"\%", "%").replace(r"\#", "#")
    return t.strip()


def add_inline_math_or_code(p, latex, size=12):
    """行内公式：纯 \\texttt 走等宽文本，其余走 OMML；转换失败再兜底。"""
    code = plain_from_texttt(latex)
    if code is not None:
        r = p.add_run(code)
        set_run_font(r, mono=True, size=size)
        return
    omml = latex_to_omml(latex)
    if omml is not None:
        p._p.append(parse_xml(omml))
        return
    r = p.add_run(latex.strip())
    set_run_font(r, mono=True, size=size)


def add_math_paragraph(doc, latex, display=False):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    if display:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
    code = plain_from_texttt(latex)
    if code is not None:
        r = p.add_run(code)
        set_run_font(r, mono=True, size=12)
        return p
    omml = latex_to_omml(latex)
    if omml is None:
        r = p.add_run(latex.strip())
        set_run_font(r, mono=True, size=12)
        return p
    wrapped = f'<m:oMathPara {nsdecls("m", "w")}>{omml}</m:oMathPara>' if display else omml
    p._p.append(parse_xml(wrapped))
    return p


# ---------------------------------------------------------------- 行内富文本
INLINE_TOKEN = re.compile(
    r"(\\\(.*?\\\)"          # 行内公式
    r"|\*\*.+?\*\*"          # 加粗
    r"|`[^`]+`"              # 行内代码
    r")",
    re.S,
)


def add_rich_runs(p, text, size=12):
    for tok in INLINE_TOKEN.split(text):
        if not tok:
            continue
        if tok.startswith(r"\(") and tok.endswith(r"\)"):
            add_inline_math_or_code(p, tok[2:-2], size=size)
        elif tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            r = p.add_run(tok[2:-2])
            set_run_font(r, size=size, bold=True)
        elif tok.startswith("`") and tok.endswith("`") and len(tok) > 2:
            r = p.add_run(tok[1:-1])
            set_run_font(r, mono=True, size=size)
        else:
            r = p.add_run(tok)
            set_run_font(r, size=size)


def add_body(doc, text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Pt(24)     # 首行缩进 2 字符（12pt * 2）
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    add_rich_runs(p, text, size=12)
    return p


def add_caption(doc, text, bold=False, size=10.5):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.first_line_indent = Pt(0)
    pf.space_before = Pt(6)
    pf.space_after = Pt(3)
    r = p.add_run(text)
    set_run_font(r, size=size, bold=bold)
    return p


# ---------------------------------------------------------------- 表格
def split_row(line):
    line = line.strip().strip("|")
    return [c.strip() for c in line.split("|")]


def add_table(doc, block):
    lines = [l for l in block.split("\n") if l.strip()]
    header = split_row(lines[0])
    aligns = split_row(lines[1])
    body = [split_row(l) for l in lines[2:]]
    ncol = len(header)

    t = doc.add_table(rows=1, cols=ncol)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True

    def render(cell, text, bold, align):
        cell.text = ""
        p = cell.paragraphs[0]
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        if align == "r":
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align == "c":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_rich_runs(p, text, size=10.5)
        for r in p.runs:
            r.font.bold = bold

    for j, h in enumerate(header):
        a = "c" if aligns[j] in (":---:", ":--:") else ("r" if aligns[j].endswith(":") else "l")
        render(t.rows[0].cells[j], h, True, a)

    for row in body:
        cells = t.add_row().cells
        for j in range(ncol):
            a = "c" if aligns[j] in (":---:", ":--:") else ("r" if aligns[j].endswith(":") else "l")
            render(cells[j], row[j] if j < len(row) else "", False, a)
    return t


# ---------------------------------------------------------------- 主流程
def main():
    text = SRC.read_text(encoding="utf-8")
    doc = Document()
    configure_styles(doc)

    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2.8)
    sec.top_margin = sec.bottom_margin = Cm(2.5)

    blocks = re.split(r"\n[ \t]*\n", text)
    for b in blocks:
        s = b.strip()
        if not s:
            continue

        # 标题
        m = re.match(r"^(#{1,4})\s+(.*)$", s, re.S)
        if m:
            lvl = len(m.group(1))
            p = doc.add_heading("", level=min(lvl, 3))
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            add_rich_runs(p, m.group(2).strip(), size={1: 16, 2: 14, 3: 12}[min(lvl, 3)])
            for r in p.runs:
                r.font.bold = True
            if lvl == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(12)
            continue

        # 表格
        if all(l.strip().startswith("|") for l in s.split("\n") if l.strip()):
            add_table(doc, s)
            continue

        # 独立公式
        if s.startswith(r"\[") and s.rstrip().endswith(r"\]"):
            latex = s.strip()[2:].rstrip()[:-2].strip()
            latex = re.sub(r"\s+", " ", latex)
            add_math_paragraph(doc, latex, display=True)
            continue

        # 图片 -> 仅保留占位说明 + 图注
        mi = re.match(r"^!\[(.*?)\]\((.*?)\)$", s, re.S)
        if mi:
            ph = doc.add_paragraph()
            ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ph.paragraph_format.first_line_indent = Pt(0)
            r = ph.add_run(f"［此处插入图片：{mi.group(2)}］")
            set_run_font(r, size=10.5)
            r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
            r.font.italic = True
            continue

        # 图注（*...*）
        if re.fullmatch(r"\*[^*].*\*", s, re.S):
            add_caption(doc, s.strip()[1:-1].strip(), bold=False)
            continue

        # 表题（**...**）
        if re.fullmatch(r"\*\*[^*].*\*\*", s, re.S):
            add_caption(doc, s.strip()[2:-2].strip(), bold=True)
            continue

        add_body(doc, s)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT))
    print("SAVED", OUT)


if __name__ == "__main__":
    main()