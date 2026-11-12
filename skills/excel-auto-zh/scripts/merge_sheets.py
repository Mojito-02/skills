#!/usr/bin/env python3
"""merge_sheets.py — 批量合并多个 Excel/CSV 工作簿与工作表
====================================================

场景：考勤汇总、销售数据整理、人力资源批量处理。

用法：
    # 合并目录下所有 xlsx 的全部工作表（自动加“来源文件/来源工作表”两列 + 汇总页）
    python3 merge_sheets.py --input "考勤表/*.xlsx" --output 月度汇总.xlsx

    # 只合并指定工作表；显式表头行；不生成汇总页
    python3 merge_sheets.py --input a.xlsx b.xlsx --sheets 4月 5月 --header-row 1 \
        --no-summary --output 汇总.xlsx

    # 输出 CSV（UTF-8 BOM，Excel/WPS 双击打开不乱码）
    python3 merge_sheets.py --input "data/*.csv" --format csv --output 汇总.csv

依赖：pip install openpyxl
"""

from __future__ import annotations

import argparse
import csv
import glob as globmod
import re
import sys
from datetime import date, datetime
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write(
        "[X] 缺少依赖 openpyxl：请先执行 pip install openpyxl 再重试。\n")
    sys.exit(2)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from xlsx_style import (  # noqa: E402
    AL_LEFT, CALC_BLACK, FORMATS, f_cell, header_row, sheet_defaults,
    summary_row, title, legend, auto_width, band_rows, freeze,
)

SOURCE_COLS = ["来源文件", "来源工作表"]


def _stdout_utf8() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass


# ================================================================
# 输入收集
# ================================================================

def expand_inputs(patterns: list[str]) -> list[Path]:
    files: list[Path] = []
    for pat in patterns:
        hits = sorted(globmod.glob(pat))
        if not hits:
            print(f"  ! 未匹配到文件：{pat}")
        for h in hits:
            p = Path(h)
            if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".csv") \
                    and not p.name.startswith("~$"):
                files.append(p)
    # 去重保序
    seen, uniq = set(), []
    for f in files:
        if f.resolve() not in seen:
            seen.add(f.resolve())
            uniq.append(f)
    return uniq


def read_csv_smart(path: Path) -> list[list]:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8"):
        try:
            return list(csv.reader(raw.decode(enc).splitlines()))
        except UnicodeDecodeError:
            continue
    return list(csv.reader(raw.decode("gbk", errors="replace").splitlines()))


def _parse_datetime(s: str):
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _coerce_csv_value(s):
    """CSV 单元格类型还原：数字→int/float、日期→datetime，与 xlsx 来源对齐，
    保证合并后数值列有千分位、日期列有日期格式。
    前导零整数与 11 位以上纯数字保持文本（工号/手机号/身份证），避免丢零与误套千分位。"""
    t = str(s).strip() if s is not None else ""
    if t == "":
        return ""
    core = t.replace(",", "")
    if re.fullmatch(r"0\d+", core) or (core.isdigit() and len(core) >= 11):
        return t
    try:
        f = float(core)
        return int(f) if f.is_integer() and abs(f) < 1e15 else f
    except ValueError:
        pass
    return _parse_datetime(t) or t


def detect_header_row(rows: list[list]) -> int:
    """自动找表头：第一行满足“非空单元格 >= 2 且下一行也有内容”。"""
    for i, row in enumerate(rows[:10]):
        nonempty = sum(1 for c in row if str(c or "").strip())
        if nonempty >= 2 and i + 1 < len(rows) and \
                any(str(c or "").strip() for c in rows[i + 1]):
            return i
    return 0


def read_table(path: Path, only_sheets: set[str] | None,
               header_row_num: int | None):
    """返回 [(sheet名, 表头list, 行list-of-list), ...]，行已对齐到表头宽度。"""
    tables = []
    if path.suffix.lower() == ".csv":
        rows = read_csv_smart(path)
        if not rows:
            return tables
        hr = (header_row_num or detect_header_row(rows) + 1) - 1
        headers = [str(h).strip() or f"列{i + 1}" for i, h in enumerate(rows[hr])]
        width = len(headers)
        data = []
        for r in rows[hr + 1:]:
            if not any(str(c or "").strip() for c in r):
                continue
            padded = list(r) + [""] * (width - len(r))
            data.append([_coerce_csv_value(c) for c in padded[:width]])
        tables.append(("CSV", headers, data))
        return tables

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            if only_sheets and ws.title not in only_sheets:
                continue
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            hr = (header_row_num or detect_header_row(rows) + 1) - 1
            head_row = rows[hr]
            headers, used = [], []
            for i, h in enumerate(head_row):
                name = str(h).strip() if h is not None else ""
                if not name:
                    name = f"列{i + 1}"
                # 重复表头改名
                base, k = name, 2
                while name in used:
                    name = f"{base}_{k}"
                    k += 1
                used.append(name)
                headers.append(name)
            width = len(headers)
            data = [list(r[:width]) + [None] * max(0, width - len(r))
                    for r in rows[hr + 1:]
                    if any(c is not None and str(c).strip() for c in r)]
            tables.append((ws.title, headers, data))
    finally:
        wb.close()
    return tables


# ================================================================
# 合并
# ================================================================

def merge(inputs: list[Path], only_sheets, header_row_num):
    """返回 (合并列名list, 数据行list, 每来源统计list)。"""
    all_rows = []
    stats = []
    header_order: list[str] = []
    header_seen = set()

    for path in inputs:
        try:
            tables = read_table(path, only_sheets, header_row_num)
        except Exception as e:  # noqa: BLE001
            print(f"  ! 无法读取 {path.name}：{type(e).__name__}: {e}")
            continue
        for sheet_name, headers, rows in tables:
            for h in headers:
                if h not in header_seen:
                    header_seen.add(h)
                    header_order.append(h)
            stats.append({"file": path.name, "sheet": sheet_name,
                          "rows": len(rows), "headers": headers})
            for r in rows:
                all_rows.append((path.name, sheet_name, headers, r))

    columns = SOURCE_COLS + header_order
    out_rows = []
    for fname, sname, headers, r in all_rows:
        row = {h: "" for h in header_order}
        for h, v in zip(headers, r):
            row[h] = v
        out_rows.append([fname, sname] + [row[h] for h in header_order])
    return columns, out_rows, stats


def column_kinds(columns, rows) -> dict:
    """每列类型：number / date / text（采样判定）。"""
    kinds = {}
    for j, name in enumerate(columns):
        if j < len(SOURCE_COLS):
            kinds[name] = "text"
            continue
        vals = [r[j] for r in rows
                if j < len(r) and r[j] not in (None, "")]
        if not vals:
            kinds[name] = "text"
            continue
        if all(isinstance(v, int) and not isinstance(v, bool) for v in vals):
            kinds[name] = "int"
            continue
        if all(isinstance(v, (int, float)) and not isinstance(v, bool)
               for v in vals):
            kinds[name] = "number"
            continue
        datelike = all(isinstance(v, (date, datetime)) for v in vals)
        kinds[name] = "date" if datelike else "text"
    return kinds


# ================================================================
# 输出
# ================================================================

def write_xlsx(columns, rows, stats, out: Path, with_summary: bool):
    wb = Workbook()
    ws = wb.active
    ws.title = "合并数据"
    sheet_defaults(ws)
    n = len(columns)
    title(ws, "数据汇总", row=1, col=1, span=n)
    legend(ws, 2, 1, span=n,
           text="前两列为来源追溯（来源文件/来源工作表），可直接用筛选查看单表数据。")
    HR = 3
    header_row(ws, HR, columns)
    kinds = column_kinds(columns, rows)

    r = HR
    for row in rows:
        r += 1
        for j, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=j,
                           value=v if v is not None else "")
            kind = kinds.get(columns[j - 1], "text")
            if kind in ("number", "int") and isinstance(v, (int, float)):
                cell.number_format = (FORMATS["int"] if kind == "int"
                                      else FORMATS["number"])
            elif kind == "date" and isinstance(v, (date, datetime)):
                cell.number_format = (FORMATS["datetime"]
                                      if isinstance(v, datetime)
                                      and (v.hour or v.minute or v.second)
                                      else FORMATS["date"])
            elif kind == "text" and isinstance(v, str):
                cell.number_format = FORMATS["text"]   # 工号/手机号列用 @
            if j <= len(SOURCE_COLS):
                cell.font = f_cell(color=CALC_BLACK)
                cell.alignment = AL_LEFT

    last = r
    band_rows(ws, HR + 1, last, col_start=3)   # 数据区隔行浅色（跳过来源两列）
    if with_summary:
        # 汇总页
        num_cols = [c for c in columns[2:]
                    if kinds[c] in ("number", "int")][:6]
        ws2 = wb.create_sheet("汇总")
        sheet_defaults(ws2)
        heads = ["来源文件", "来源工作表", "数据行数"] + \
            [f"{c} 合计" for c in num_cols]
        title(ws2, "各来源汇总", row=1, col=1, span=len(heads))
        HR2 = 3
        header_row(ws2, HR2, heads)
        rr = HR2
        first_data = HR2 + 1
        for s in stats:
            rr += 1
            ws2.cell(row=rr, column=1, value=s["file"])
            ws2.cell(row=rr, column=2, value=s["sheet"])
            ws2.cell(row=rr, column=3, value=s["rows"])
            for k, col in enumerate(num_cols):
                j = columns.index(col) + 1
                letter = get_column_letter(j)
                ws2.cell(row=rr, column=4 + k,
                         value=f"=SUMIF('合并数据'!$A:$A,A{rr},"
                               f"'合并数据'!{letter}:{letter})")
                ws2.cell(row=rr, column=4 + k).number_format = FORMATS["number"]
        # 总计行
        rr += 1
        vals = ["总计", "", f"=SUM(C{first_data}:C{rr - 1})"] + \
            [f"=SUM({get_column_letter(4 + k)}{first_data}:"
             f"{get_column_letter(4 + k)}{rr - 1})" for k in range(len(num_cols))]
        summary_row(ws2, rr, vals,
                    numfmts=[None, None, FORMATS["int"]] +
                    [FORMATS["number"]] * len(num_cols))
        auto_width(ws2, header_row=HR2)
        freeze(ws2, header_row_num=HR2)
    auto_width(ws, header_row=HR)
    freeze(ws, header_row_num=HR)
    wb.save(out)


def write_csv(columns, rows, out: Path):
    """CSV 输出（UTF-8 BOM）。
    源文件里的公式单元格导出为 '=公式文本（加 ' 前缀），避免 Excel/WPS
    打开 CSV 时在合并后的新上下文里误算公式；程序处理时去掉前缀即可。"""
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(columns)
        for row in rows:
            w.writerow([_csv_cell(v) for v in row])


def _csv_cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str) and v.startswith("="):
        return "'" + v
    if isinstance(v, datetime):
        s = v.isoformat(sep=" ")
        return s[:-9] if s.endswith(" 00:00:00") else s
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# ================================================================
# CLI
# ================================================================

def main() -> int:
    _stdout_utf8()
    p = argparse.ArgumentParser(description="批量合并多个 Excel/CSV 工作簿与工作表")
    p.add_argument("--input", nargs="+", required=True,
                   help="文件或通配模式（可多个，如 \"考勤表/*.xlsx\"）")
    p.add_argument("--sheets", nargs="*", default=None,
                   help="只合并指定名称的工作表")
    p.add_argument("--header-row", type=int, default=None,
                   help="表头所在行（1 起，默认自动探测）")
    p.add_argument("--output", default="汇总.xlsx")
    p.add_argument("--format", choices=["xlsx", "csv"], default="xlsx")
    p.add_argument("--no-summary", action="store_true", help="不生成汇总页")
    args = p.parse_args()

    inputs = expand_inputs(args.input)
    if not inputs:
        raise SystemExit("[X] 没有匹配到任何输入文件")
    print(f"读取 {len(inputs)} 个文件：{', '.join(f.name for f in inputs[:8])}"
          f"{' …' if len(inputs) > 8 else ''}")

    only_sheets = set(args.sheets) if args.sheets else None
    columns, rows, stats = merge(inputs, only_sheets, args.header_row)
    if not rows:
        raise SystemExit("[X] 输入文件中没有读到数据行")
    print(f"合并完成：{len(stats)} 个工作表，共 {len(rows)} 行，"
          f"{len(columns) - len(SOURCE_COLS)} 个数据列（缺失列已自动补空）")

    out = Path(args.output)
    if args.format == "csv":
        write_csv(columns, rows, out)
        print(f"[OK] 已输出 {out}（UTF-8 BOM，Excel/WPS 可直接打开）")
    else:
        write_xlsx(columns, rows, stats, out, with_summary=not args.no_summary)
        extra = "（含「合并数据」+「汇总」两个工作表）" if not args.no_summary else ""
        print(f"[OK] 已输出 {out}{extra}")
    for s in stats:
        print(f"  · {s['file']} / {s['sheet']}：{s['rows']} 行")
    return 0


if __name__ == "__main__":
    sys.exit(main())
