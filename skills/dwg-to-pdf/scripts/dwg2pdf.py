#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""DWG / DXF  ->  PDF  (no CAD software required)

Pipeline:
  DWG --(ODA File Converter, portable)--> DXF --(ezdxf + matplotlib)--> PDF

Usage:
  python dwg2pdf.py <input.dwg|input.dxf> [output.pdf] [--color] [--keep-dxf]

Notes:
  * The original DWG text styles reference AutoCAD SHX fonts (isocp.shx /
    hztxt.shx / gbcbig.shx ...) which are NOT shipped on this machine.
    Chinese is usually carried by MTEXT inline overrides like
        {\\fFangSong_GB2312|b0|i0|c134|p49;设 计}
    Those must be stripped, otherwise ezdxf resolves nothing and CJK turns
    into garbage. After stripping, every text style is pointed at a real
    system CJK TTF (simhei.ttf) so both CJK and Latin render.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile

ODA_EXE = r"C:\Users\ASUS\.workbuddy\tools\ODAFileConverter\ODAFileConverter.exe"
CJK_TTF = "simhei.ttf"          # fallback candidates tried in order
CJK_CANDIDATES = ["simhei.ttf", "simsun.ttc", "msyh.ttc", "simfang.ttf"]
PAGE_H_IN = 23.39               # 594 mm -> A1 series height

FCODE = re.compile(r"\\f[^;]*;")


def dwg_to_dxf(dwg_path, workdir):
    """Run the portable ODA File Converter. Returns produced .dxf path."""
    ind = os.path.join(workdir, "in")
    outd = os.path.join(workdir, "out")
    os.makedirs(ind, exist_ok=True)
    os.makedirs(outd, exist_ok=True)
    # ASCII staging name: ODA is happier with plain paths
    staged = os.path.join(ind, "drawing.dwg")
    shutil.copy2(dwg_path, staged)
    cmd = [ODA_EXE, ind, outd, "ACAD2018", "DXF", "0", "1", "*.DWG"]
    subprocess.run(cmd, check=True, timeout=600)
    made = [os.path.join(outd, f) for f in os.listdir(outd)
            if f.lower().endswith(".dxf")]
    if not made:
        raise RuntimeError("ODA produced no DXF (check the DWG is readable)")
    return made[0]


def strip_inline_fonts(doc):
    """Remove MTEXT inline \\f<font>; overrides so style fonts apply."""
    n = 0
    for layout in doc.layouts:
        for e in layout:
            if e.dxftype() == "MTEXT" and FCODE.search(e.text):
                e.text = FCODE.sub("", e.text)
                n += 1
    for blk in doc.blocks:
        for e in blk:
            if e.dxftype() == "MTEXT" and FCODE.search(e.text):
                e.text = FCODE.sub("", e.text)
                n += 1
    return n


def render(dxf_path, pdf_path, color=False):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import ezdxf
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    from ezdxf.addons.drawing.config import (
        Configuration, BackgroundPolicy, ColorPolicy)
    from ezdxf.fonts import fonts

    face = next((f for f in CJK_CANDIDATES if fonts.font_manager.has_font(f)),
                None)
    if face is None:
        print("WARNING: no CJK font found; Chinese may render incorrectly")

    doc = ezdxf.readfile(dxf_path)
    strip_inline_fonts(doc)
    if face:
        for st in doc.styles:
            try:
                st.dxf.font = face
                st.dxf.bigfont = ""
            except Exception:
                pass

    msp = doc.modelspace()
    fig = plt.figure()
    ax = fig.add_axes([0, 0, 1, 1])
    ctx = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    cp = ColorPolicy.COLOR if color else ColorPolicy.MONOCHROME_LIGHT_BG
    Frontend(ctx, backend,
             config=Configuration(background_policy=BackgroundPolicy.WHITE,
                                  color_policy=cp)).draw_layout(msp, finalize=True)

    xl, yl = ax.get_xlim(), ax.get_ylim()
    w, h = abs(xl[1] - xl[0]), abs(yl[1] - yl[0])
    fig.set_size_inches(PAGE_H_IN * (w / h), PAGE_H_IN)
    ax.set_facecolor("white")
    fig.patch.set_facecolor("white")

    # write to an ASCII temp path first, then move to the (possibly CJK) target
    tmp = pdf_path + ".tmp.pdf"
    fig.savefig(tmp, format="pdf", facecolor="white")
    plt.close(fig)
    shutil.move(tmp, pdf_path)
    return PAGE_H_IN * (w / h), PAGE_H_IN


def main():
    args = [a for a in sys.argv[1:]]
    color = "--color" in args
    keep = "--keep-dxf" in args
    args = [a for a in args if not a.startswith("--")]
    if not args:
        print(__doc__)
        return 2
    src = os.path.abspath(args[0])
    dst = os.path.abspath(args[1]) if len(args) > 1 else os.path.splitext(src)[0] + ".pdf"

    work = tempfile.mkdtemp(prefix="dwg2pdf_")
    try:
        dxf = src if src.lower().endswith(".dxf") else dwg_to_dxf(src, work)
        w, h = render(dxf, dst, color=color)
        print("OK  %s" % dst)
        print("    page %.1f x %.1f in  (%.0f x %.0f mm)" % (w, h, w * 25.4, h * 25.4))
        if keep:
            kd = os.path.splitext(dst)[0] + ".dxf"
            shutil.copy2(dxf, kd)
            print("    dxf kept: %s" % kd)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
