#!/usr/bin/env python3
"""create_report.py — 生成带专业格式的 Excel 报表（三色规则）
====================================================

用法：
    python3 create_report.py demo --output 销售月报.xlsx
    python3 create_report.py from-csv data.csv --title "销售月报" --output 报表.xlsx
    python3 create_report.py from-json spec.json --output 报表.xlsx
    python3 create_report.py verify 报表.xlsx

三色规则：蓝色=输入项 / 黑色=公式计算值 / 绿色=跨表引用。
依赖：pip install openpyxl
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write(
        "[X] 缺少依赖 openpyxl：请先执行 pip install openpyxl 再重试。\n"
        "    （parse_xlsm.py 无此依赖，可正常使用）\n")
    sys.exit(2)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsx_style import (  # noqa: E402
    CALC_BLACK, FORMATS, INPUT_BLUE, LINK_GREEN,
    auto_width, band_rows, build_formula, f_cell, freeze, header_row, legend,
    numfmt_for, put, safe_value, sheet_defaults, summary_row, title,
)


def _stdout_utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass


# ================================================================
# 公共构建逻辑
# ================================================================

def _finish_sheet(ws, *, header_row_num: int, has_title: bool) -> None:
    sheet_defaults(ws)
    auto_width(ws, header_row=header_row_num)
    freeze(ws, header_row_num=header_row_num)


def _apply_cell_style(cell, *, color: str, numfmt: str | None) -> None:
    cell.font = f_cell(color=color)
    if numfmt:
        cell.number_format = numfmt


# ================================================================
# 子命令 1：demo —— 示例销售月报（完整演示三色规则）
# ================================================================

DEMO_ROWS = [
    ("张伟",   120000, 139500),
    ("王芳",   120000, 126800),
    ("李娜",   100000, 118200),
    ("刘洋",   100000,  97600),
    ("陈静",    90000,  95400),
    ("杨帆",    90000,  72300),
    ("赵磊",    80000,  83100),
    ("孙丽",    80000,  66400),
]


def cmd_demo(args) -> int:
    out = Path(args.output)
    wb = Workbook()

    # ---- 参数表（蓝色输入项）----
    ws_p = wb.active
    ws_p.title = "参数"
    sheet_defaults(ws_p)
    title(ws_p, "提成参数（蓝字可直接修改）", row=1, col=1, span=2)
    params = [("最低完成率门槛", 1.0, FORMATS["percent"]),
              ("低档提成率（未达标）", 0.03, FORMATS["percent"]),
              ("高档提成率（达标）",   0.05, FORMATS["percent"])]
    header_row(ws_p, 3, ["参数名", "取值"])
    for i, (name, val, nf) in enumerate(params):
        r = 4 + i
        put(ws_p, r, 1, name, color=CALC_BLACK, numfmt=FORMATS["text"])
        put(ws_p, r, 2, val, color=INPUT_BLUE, numfmt=nf)
    auto_width(ws_p, header_row=3)
    legend(ws_p, 8, 1, span=2,
           text="本表为输入表：修改蓝字数值，销售明细表的提成会自动重算。")

    # ---- 销售明细表 ----
    ws = wb.create_sheet("销售明细")
    sheet_defaults(ws)
    headers = ["销售员", "目标销售额", "实际销售额", "完成率", "排名", "提成"]
    title(ws, args.title or "销售月报（示例数据）", row=1, col=1, span=len(headers))
    legend(ws, 2, 1, span=len(headers))
    HR = 3                                   # 表头行号
    header_row(ws, HR, headers)
    first = HR + 1
    last = first + len(DEMO_ROWS) - 1

    key2col = {"name": 1, "target": 2, "actual": 3, "rate": 4,
               "rank": 5, "bonus": 6}
    key2letter = {k: get_column_letter(v) for k, v in key2col.items()}

    for i, (name, target, actual) in enumerate(DEMO_ROWS):
        r = first + i
        put(ws, r, 1, name, numfmt=FORMATS["text"])
        put(ws, r, 2, target, color=INPUT_BLUE, numfmt=FORMATS["currency"])
        put(ws, r, 3, actual, color=INPUT_BLUE, numfmt=FORMATS["currency"])
        # 完成率：黑字公式（除法包 IFERROR）
        put(ws, r, 4,
            f"=IFERROR({key2letter['actual']}{r}/{key2letter['target']}{r},\"-\")",
            color=CALC_BLACK, numfmt=FORMATS["percent"])
        # 排名：黑字公式（RANK 兼容 Excel/WPS）
        put(ws, r, 5,
            f"=RANK({key2letter['actual']}{r},"
            f"{key2letter['actual']}${first}:{key2letter['actual']}${last})",
            color=CALC_BLACK, numfmt=FORMATS["int"])
        # 提成：绿字跨表引用（引用“参数”表的三个输入项）
        put(ws, r, 6,
            f"=IF({key2letter['rate']}{r}>='参数'!$B$4,"
            f"{key2letter['actual']}{r}*'参数'!$B$6,"
            f"{key2letter['actual']}{r}*'参数'!$B$5)",
            color=LINK_GREEN, numfmt=FORMATS["currency"])

    band_rows(ws, first, last)

    summary_row(
        ws, last + 1,
        ["合计", f"=SUM(B{first}:B{last})", f"=SUM(C{first}:C{last})",
         f"=IFERROR(C{last + 1}/B{last + 1},\"-\")", "", f"=SUM(F{first}:F{last})"],
        colors=[CALC_BLACK] * 6,
        numfmts=[FORMATS["text"], FORMATS["currency"], FORMATS["currency"],
                 FORMATS["percent"], None, FORMATS["currency"]],
    )
    _finish_sheet(ws, header_row_num=HR, has_title=True)
    ws.column_dimensions["A"].width = 14

    wb.save(out)
    print(f"[OK] 已生成 {out}")
    print("     工作表：销售明细（三色规则示例）、参数（蓝色输入项）")
    print("     颜色图例：蓝色=输入项 / 黑色=公式计算 / 绿色=跨表引用")
    return 0


# ================================================================
# 子命令 2：from-json —— 声明式构建
# ================================================================

def _as_cell_style(spec_col) -> str:
    style = (spec_col.get("style") or "calc").lower()
    return {"input": INPUT_BLUE, "calc": CALC_BLACK, "link": LINK_GREEN}.get(
        style, CALC_BLACK)


def cmd_from_json(args) -> int:
    spec_path = Path(args.spec)
    spec = json.loads(spec_path.read_text(encoding="utf-8-sig"))
    out = Path(args.output)
    wb = Workbook()
    first_sheet = True

    for si, sh in enumerate(spec.get("sheets", [])):
        name = sh.get("name", "Sheet")
        ws = wb.active if first_sheet else wb.create_sheet()
        ws.title = name
        first_sheet = False
        sheet_defaults(ws)

        cols = sh.get("columns", [])
        rows = sh.get("rows", [])
        n = len(cols)
        # sheet 级 title 优先；否则首个工作表继承 spec 顶层的 title（与文档示例一致）
        sheet_title = sh.get("title") or (
            spec.get("title") if si == 0 else None)
        if sheet_title:
            title(ws, sheet_title, row=1, col=1, span=n)
            legend(ws, 2, 1, span=n)
            hr = 3
        else:
            legend(ws, 1, 1, span=n)
            hr = 2

        header_row(ws, hr, [c.get("header", c.get("key", "")) for c in cols])

        # key -> 列号/字母
        key2col, key2letter = {}, {}
        for idx, c in enumerate(cols, start=1):
            k = c.get("key")
            if k:
                key2col[k] = idx
                key2letter[k] = get_column_letter(idx)
        # 无 key 的公式列也允许被 {col.header} 以表头名引用
        for idx, c in enumerate(cols, start=1):
            h = c.get("header")
            if h and h not in key2letter:
                key2letter[h] = get_column_letter(idx)

        first = hr + 1
        last = first + len(rows) - 1 if rows else first
        for i, row_data in enumerate(rows):
            r = first + i
            for idx, c in enumerate(cols, start=1):
                nf = numfmt_for(c.get("type_out") or c.get("type"))
                color = _as_cell_style(c)
                if "formula" in c:
                    val = build_formula(c["formula"], r, key2letter, first, last)
                else:
                    k = c.get("key")
                    if isinstance(row_data, dict):
                        val = row_data.get(k, "") if k else ""
                    else:                      # 行也允许是数组
                        val = row_data[idx - 1] if idx - 1 < len(row_data) else ""
                    # 数据行没有 style 的默认按输入项（蓝字）
                    if "style" not in c and "formula" not in c:
                        color = INPUT_BLUE
                put(ws, r, idx, val, color=color, numfmt=nf)
        if rows:
            band_rows(ws, first, last)

        # 汇总行
        sum_specs = sh.get("summary", [])
        if sum_specs and rows:
            vals = [""] * n
            fmts: list[str | None] = [None] * n
            for j, s in enumerate(sum_specs):
                label = s.get("label")
                col_key = s.get("col") or s.get("header")
                func = s.get("func", "SUM").upper()
                if label is not None and j == 0:
                    vals[0] = label
                if col_key:
                    letter = key2letter.get(col_key)
                    if not letter:
                        raise SystemExit(f"[X] summary 引用未知列 '{col_key}'")
                    idx = next(
                        (k - 1 for k, c in enumerate(cols, start=1)
                         if c.get("key") == col_key or c.get("header") == col_key), 0)
                    vals[idx] = f"={func}({letter}{first}:{letter}{last})"
                    cdef = cols[idx]
                    disp = cdef.get("type_out") or cdef.get("type")
                    is_pct = "percent" in (cdef.get("type"), cdef.get("type_out"))
                    fmts[idx] = FORMATS["percent"] if is_pct else numfmt_for(disp)
                    # 百分比列默认求平均（求和会超过 100%，没有业务意义）
                    if is_pct and "func" not in s:
                        vals[idx] = f"=AVERAGE({letter}{first}:{letter}{last})"
            summary_row(ws, last + 1, vals, numfmts=fmts)

        _finish_sheet(ws, header_row_num=hr, has_title=bool(sh.get("title")))

    wb.save(out)
    n_sheets = len(spec.get("sheets", []))
    print(f"[OK] 已生成 {out}（{n_sheets} 个工作表，样式与三色规则已应用）")
    return 0


# ================================================================
# 子命令 3：from-csv —— CSV 一键转报表
# ================================================================

def _read_csv_smart(path: Path):
    """读 CSV：先试 UTF-8（含 BOM），失败回退 GBK（Windows 中文常见）。"""
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8"):
        try:
            return list(csv.reader(raw.decode(enc).splitlines()))
        except UnicodeDecodeError:
            continue
    return list(csv.reader(raw.decode("gbk", errors="replace").splitlines()))


def _looks_number(s: str) -> bool:
    core = s.replace(",", "")
    # 前导零整数（工号 007）与 11 位以上纯数字（手机号/身份证）保持文本
    if re.fullmatch(r"0\d+", core) or (core.isdigit() and len(core) >= 11):
        return False
    try:
        float(core)
        return True
    except ValueError:
        return False


def _looks_int(s: str) -> bool:
    try:
        return float(s.replace(",", "")).is_integer()
    except ValueError:
        return False


def _looks_date(s: str):
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def cmd_from_csv(args) -> int:
    src = Path(args.input)
    out = Path(args.output)
    table = _read_csv_smart(src)
    if not table:
        raise SystemExit("[X] CSV 为空")
    header_idx = max(0, (args.header_row or 1) - 1)
    headers = [h.strip() or f"列{i + 1}" for i, h in enumerate(table[header_idx])]
    data = table[header_idx + 1:]
    # 去掉完全为空的行
    data = [row for row in data if any(str(c).strip() for c in row)]

    # 列类型探测（全列扫描）
    kinds = []
    for j in range(len(headers)):
        vals = [row[j] for row in data if j < len(row) and str(row[j]).strip()]
        if not vals:
            kinds.append("text")
            continue
        if all(_looks_number(v) for v in vals):
            kinds.append("int" if all(_looks_int(v) for v in vals) else "number")
        elif all(_looks_date(v) is not None for v in vals):
            kinds.append("date")
        else:
            kinds.append("text")

    wb = Workbook()
    ws = wb.active
    ws.title = args.sheet_name or "数据"
    sheet_defaults(ws)
    title(ws, args.title or src.stem, row=1, col=1, span=len(headers))
    legend(ws, 2, 1, span=len(headers))
    HR = 3
    header_row(ws, HR, headers)
    first = HR + 1

    for i, row in enumerate(data):
        r = first + i
        for j in range(len(headers)):
            raw_v = row[j] if j < len(row) else ""
            s = str(raw_v).strip()
            if s == "":
                put(ws, r, j + 1, "", color=INPUT_BLUE)
                continue
            if kinds[j] == "int":
                put(ws, r, j + 1, int(s.replace(",", "")),
                    color=INPUT_BLUE, numfmt=FORMATS["int"])
            elif kinds[j] == "number":
                put(ws, r, j + 1, float(s.replace(",", "")),
                    color=INPUT_BLUE, numfmt=FORMATS["number"])
            elif kinds[j] == "date":
                put(ws, r, j + 1, _looks_date(s),
                    color=INPUT_BLUE, numfmt=FORMATS["datetime"]
                    if " " in s else FORMATS["date"])
            else:
                put(ws, r, j + 1, s, color=INPUT_BLUE, numfmt=FORMATS["text"])

    last = first + len(data) - 1
    if data:
        band_rows(ws, first, last)
    if data and not args.no_total:
        # 对数值列追加 SUM 合计（黑字公式）
        vals: list = ["合计"] + [""] * (len(headers) - 1)
        fmts: list[str | None] = [FORMATS["text"]] + [None] * (len(headers) - 1)
        for j, k in enumerate(kinds):
            if k in ("number", "int"):
                letter = get_column_letter(j + 1)
                vals[j] = f"=SUM({letter}{first}:{letter}{last})"
                fmts[j] = FORMATS["int"] if k == "int" else FORMATS["number"]
        summary_row(ws, last + 1, vals, numfmts=fmts)

    _finish_sheet(ws, header_row_num=HR, has_title=True)
    wb.save(out)
    print(f"[OK] 已生成 {out}（{len(data)} 行 × {len(headers)} 列，"
          f"来源：{src.name}，数值列已加合计公式）")
    return 0


# ================================================================
# 子命令 4：verify —— 交付前自检
# ================================================================

def cmd_verify(args) -> int:
    path = Path(args.file)
    wb = load_workbook(path)
    print(f"文件：{path.name}（{path.stat().st_size:,} bytes）")
    for ws in wb.worksheets:
        n_formula = sum(
            1 for row in ws.iter_rows()
            for c in row
            if isinstance(c.value, str) and c.value.startswith("=")
        )
        merged = len(ws.merged_cells.ranges)
        frozen = "是" if ws.freeze_panes else "否"
        print(f"  - 工作表「{ws.title}」：{ws.max_row} 行 × {ws.max_column} 列，"
              f"公式 {n_formula} 个，合并单元格 {merged}，冻结表头：{frozen}")
    print("提示：请人工抽查 2-3 个单元格的值与公式；"
          "蓝字=可改输入，黑字=公式，绿字=跨表引用。")
    return 0


# ================================================================
# CLI
# ================================================================

def main() -> int:
    _stdout_utf8()
    p = argparse.ArgumentParser(
        description="生成带专业格式的 Excel 报表（三色规则：蓝=输入/黑=计算/绿=跨表）")
    sub = p.add_subparsers(dest="cmd", required=True)

    d = sub.add_parser("demo", help="生成示例销售月报（演示三色规则）")
    d.add_argument("--output", default="销售月报示例.xlsx")
    d.add_argument("--title", default=None)
    d.set_defaults(func=cmd_demo)

    j = sub.add_parser("from-json", help="按 JSON spec 生成报表")
    j.add_argument("spec")
    j.add_argument("--output", default="报表.xlsx")
    j.set_defaults(func=cmd_from_json)

    c = sub.add_parser("from-csv", help="把 CSV 转成格式化报表")
    c.add_argument("input")
    c.add_argument("--output", default="报表.xlsx")
    c.add_argument("--title", default=None)
    c.add_argument("--header-row", type=int, default=1)
    c.add_argument("--sheet-name", default=None)
    c.add_argument("--no-total", action="store_true", help="不加合计行")
    c.set_defaults(func=cmd_from_csv)

    v = sub.add_parser("verify", help="自检生成的报表结构")
    v.add_argument("file")
    v.set_defaults(func=cmd_verify)

    args = p.parse_args()
    try:
        return args.func(args)
    except SystemExit:
        raise
    except Exception as e:  # noqa: BLE001
        print(f"[X] 失败：{type(e).__name__}: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
