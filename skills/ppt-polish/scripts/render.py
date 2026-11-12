#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
把 PPT 渲染成 PNG —— 通过 WPS / PowerPoint 的 COM 接口，无需 LibreOffice。

用法:
    python render.py <pptx路径> <输出目录> [宽] [高]

例:
    python render.py "D:/deck.pptx" "./render_out" 1280 720

依赖: pip install pywin32
说明:
    - 自动探测 KWPP.Application (WPS) -> PowerPoint.Application
    - 用 pres.Export(dir, "PNG", w, h)，不要用 SaveAs（WPS 下会报"发生意外"）
    - 输出目录已存在时直接复用（safe-delete 环境下不要删旧目录，换目录更安全）
"""
import sys, os, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

CANDIDATES = ["KWPP.Application", "PowerPoint.Application"]


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    src = os.path.abspath(sys.argv[1])
    out = os.path.abspath(sys.argv[2])
    width = int(sys.argv[3]) if len(sys.argv) > 3 else 1280
    height = int(sys.argv[4]) if len(sys.argv) > 4 else 720

    if not os.path.exists(src):
        print(f"[错误] 源文件不存在: {src}")
        sys.exit(1)

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("[错误] 缺少 pywin32。请先安装: pip install pywin32")
        sys.exit(1)

    pythoncom.CoInitialize()

    app = None
    progid = None
    for prog in CANDIDATES:
        try:
            app = win32com.client.DispatchEx(prog)
            progid = prog
            break
        except Exception:
            continue

    if app is None:
        print("[错误] 没有可用的 COM 渲染接口。")
        print("       运行环境需要安装 WPS Office 或 Microsoft PowerPoint（Windows）。")
        print("       （若确实没有，可安装 LibreOffice 后用 soffice --convert-to pdf 转图）")
        sys.exit(1)

    os.makedirs(out, exist_ok=True)
    print(f"渲染引擎: {progid}")
    print(f"源文件  : {src}")
    print(f"输出目录: {out}")

    try:
        pres = app.Presentations.Open(src, ReadOnly=True, WithWindow=False)
        n = pres.Slides.Count
        print(f"页数    : {n}")
        pres.Export(out, "PNG", width, height)
        files = [f for f in os.listdir(out) if f.lower().endswith(".png")]
        print(f"完成    : 导出 {len(files)} 张 PNG")
        pres.Close()
    except Exception as e:
        print(f"[错误] 渲染失败: {type(e).__name__}: {e}")
        sys.exit(1)
    finally:
        try:
            app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
