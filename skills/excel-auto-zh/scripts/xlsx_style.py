"""excel-auto-zh 共享样式模块
=================================
三个脚本（create_report / merge_sheets / 手写代码）共用的设计基元，
单一事实来源：颜色、字体、数字格式、表头/合计样式、列宽自适应。
手写报表代码时直接 `from xlsx_style import ...` 复用，不要散落硬编码颜色。

三色规则（财务建模惯例）：
    蓝色 0000FF = 输入项（用户可修改的原始数值）
    黑色 000000 = 计算值（公式结果）
    绿色 008000 = 跨表引用（引用其他工作表）
"""

from __future__ import annotations

import platform
import unicodedata

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------
# §1 三色规则与基础配色
# ---------------------------------------------------------------
INPUT_BLUE = "0000FF"    # 输入项字体
CALC_BLACK = "000000"    # 计算/公式字体
LINK_GREEN = "008000"    # 跨表引用字体

HEADER_BG = "1F4E78"     # 表头深蓝底
HEADER_FG = "FFFFFF"     # 表头白字
TITLE_COLOR = "1F4E78"   # 大标题
BAND_BG = "EEF3F9"       # 隔行浅色
SUMMARY_BG = "DDEBF7"    # 合计行底色
LEGEND_GRAY = "808080"   # 图例/脚注
BORDER_GRAY = "BFBFBF"   # 细边框

# ---------------------------------------------------------------
# §2 字体解析（Windows / macOS / Linux 回退链）
# ---------------------------------------------------------------
_FONT_HINTS = {
    "Windows": ["Microsoft YaHei", "SimHei"],
    "Darwin": ["PingFang SC", "Hiragino Sans GB"],
    "Linux": ["Noto Sans CJK SC", "WenQuanYi Micro Hei", "Source Han Sans SC"],
}


def resolve_font() -> str:
    """返回当前操作系统最可能存在的中文字体名。"""
    candidates = _FONT_HINTS.get(platform.system(), _FONT_HINTS["Linux"])
    return candidates[0]


FONT_NAME = resolve_font()

# ---------------------------------------------------------------
# §3 数字格式
# ---------------------------------------------------------------
FORMATS = {
    "text": "@",
    "int": "#,##0",
    "number": "#,##0.00",
    "currency": "¥#,##0.00",
    "percent": "0.0%",
    "date": "YYYY-MM-DD",
    "datetime": "YYYY-MM-DD HH:MM",
}

# type 别名 → FORMATS 键
_TYPE_ALIASES = {
    "text": "text", "str": "text", "string": "text",
    "int": "int", "integer": "int",
    "number": "number", "float": "number", "decimal": "number",
    "currency": "currency", "money": "currency", "cny": "currency",
    "percent": "percent", "pct": "percent",
    "date": "date", "datetime": "datetime",
}


def numfmt_for(type_name: str | None) -> str | None:
    """按类型名返回数字格式码；None/未知返回 None（用常规格式）。"""
    if not type_name:
        return None
    key = _TYPE_ALIASES.get(str(type_name).lower())
    return FORMATS.get(key) if key else None


# ---------------------------------------------------------------
# §4 边框 / 对齐 / 字体工厂
# ---------------------------------------------------------------
_THIN = Side(style="thin", color=BORDER_GRAY)
BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
BORDER_SUMMARY = Border(
    left=_THIN, right=_THIN, bottom=_THIN,
    top=Side(style="medium", color=HEADER_BG),
)


def f_cell(color=CALC_BLACK, bold=False, size=11) -> Font:
    return Font(name=FONT_NAME, size=size, bold=bold, color=color)


def f_header() -> Font:
    return Font(name=FONT_NAME, size=11, bold=True, color=HEADER_FG)


def f_title() -> Font:
    return Font(name=FONT_NAME, size=14, bold=True, color=TITLE_COLOR)


def f_legend() -> Font:
    return Font(name=FONT_NAME, size=9, color=LEGEND_GRAY)


AL_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_LEFT = Alignment(horizontal="left", vertical="center")
AL_RIGHT = Alignment(horizontal="right", vertical="center")


# ---------------------------------------------------------------
# §5 写入助手（带三色规则）
# ---------------------------------------------------------------
def safe_value(value):
    """文本以 = 开头时加 ' 前缀，防止被 Excel/WPS 当成公式报错。
    仅在写入“确实是文本”的值时使用；put() 默认把 = 开头当公式原样写入。"""
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def put(ws, row, col, value, color=CALC_BLACK, numfmt=None,
        align=None, bold=False, border=True, fill=None):
    """写一个单元格并套样式。color 用 INPUT_BLUE/CALC_BLACK/LINK_GREEN。
    以 = 开头的字符串视为公式原样写入；纯文本左对齐；
    数字/公式按 numfmt 决定对齐（text 格式左对齐，其余右对齐）。"""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = f_cell(color=color, bold=bold)
    if numfmt:
        cell.number_format = numfmt
    if align is not None:
        cell.alignment = align
    elif isinstance(value, str):
        is_formula = value.startswith("=")
        if not is_formula or numfmt in (None, FORMATS["text"]):
            cell.alignment = AL_LEFT
        else:
            cell.alignment = AL_RIGHT
    else:
        cell.alignment = AL_RIGHT if numfmt not in (None, FORMATS["text"]) else AL_LEFT
    if border:
        cell.border = BORDER_ALL
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    return cell


def header_row(ws, row, headers, col_start=1, height=24):
    """写表头行：深蓝底白字居中 + 细边框。返回结束列号。"""
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col_start + i, value=h)
        c.font = f_header()
        c.fill = PatternFill("solid", fgColor=HEADER_BG)
        c.alignment = AL_CENTER
        c.border = BORDER_ALL
    ws.row_dimensions[row].height = height
    return col_start + len(headers) - 1


def title(ws, text, row=1, col=1, span=None, size=14):
    """合并单元格写大标题（不跨边框线）。"""
    if span and span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=FONT_NAME, size=size, bold=True, color=TITLE_COLOR)
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 28


def legend(ws, row, col, span=None, text=None):
    """写三色图例说明行（小灰字）。"""
    if text is None:
        text = ("颜色图例：蓝色 = 输入项（可直接修改）｜黑色 = 公式计算结果"
                "｜绿色 = 跨表引用")
    if span and span > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + span - 1)
    c = ws.cell(row=row, column=col, value=text)
    c.font = f_legend()
    c.alignment = Alignment(horizontal="left", vertical="center")


def summary_row(ws, row, values, col_start=1, colors=None, numfmts=None,
                height=26):
    """写合计/平均行：浅蓝底、上粗边框、加粗。colors/numfmts 可传列表或单值。"""
    n = len(values)

    def pick(spec, i, default):
        if spec is None:
            return default
        if isinstance(spec, (list, tuple)):
            return spec[i] if i < len(spec) else default
        return spec

    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col_start + i, value=v)
        c.font = f_cell(color=pick(colors, i, CALC_BLACK), bold=True)
        nf = pick(numfmts, i, None)
        if nf:
            c.number_format = nf
        c.alignment = AL_RIGHT if nf not in (None, FORMATS["text"]) else AL_LEFT
        c.border = BORDER_SUMMARY
        c.fill = PatternFill("solid", fgColor=SUMMARY_BG)
    ws.row_dimensions[row].height = height
    return col_start + n - 1


# ---------------------------------------------------------------
# §6 布局（冻结、筛选、列宽）
# ---------------------------------------------------------------
def freeze(ws, header_row_num=1, last_col=None, autofilter=True, last_row=None):
    """冻结表头并加自动筛选。last_col/last_row 缺省用当前数据区。"""
    last_col = last_col or ws.max_column
    ws.freeze_panes = ws.cell(row=header_row_num + 1, column=1)
    if autofilter:
        last_row = last_row or ws.max_row
        ref = f"A{header_row_num}:{get_column_letter(last_col)}{last_row}"
        ws.auto_filter.ref = ref


def display_width(text) -> float:
    """估算显示宽度：CJK 全角字符按 1.7，其余按 1。"""
    if text is None:
        return 0
    if hasattr(text, "strftime"):            # date/datetime
        return 17
    s = str(text)
    w = 0.0
    for ch in s:
        w += 1.7 if unicodedata.east_asian_width(ch) in ("W", "F") else 1.0
    return w


def auto_width(ws, min_w=8, max_w=42, extra=3, header_row=None):
    """按内容自适应列宽（中文友好）。
    header_row 给定时只从该行开始测宽——大标题、图例这类合并长文本
    不会把第一列撑到上限；合并单元格的锚点格同样跳过。"""
    first_row = header_row if header_row is not None else 1
    anchors = {(r.min_row, r.min_col) for r in ws.merged_cells.ranges}
    for col_cells in ws.iter_cols(min_col=1, max_col=ws.max_column,
                                  min_row=first_row):
        letter = get_column_letter(col_cells[0].column)
        w = 0.0
        for c in col_cells:
            if (c.row, c.column) in anchors:
                continue                      # 标题/图例等合并长文本
            v = c.value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                w = max(w, 12)                # 公式列给个基础宽度
            else:
                w = max(w, display_width(v))
        ws.column_dimensions[letter].width = min(max_w, max(min_w, w + extra))


def band_rows(ws, first_row, last_row, col_start=1, col_end=None,
              color=BAND_BG):
    """数据区隔行浅色（从第 2 行数据开始隔行着色，不覆盖已有填充）。"""
    if last_row < first_row:
        return
    col_end = col_end or ws.max_column
    for r in range(first_row, last_row + 1):
        if (r - first_row) % 2 == 1:
            for cidx in range(col_start, col_end + 1):
                cell = ws.cell(row=r, column=cidx)
                if cell.fill is None or cell.fill.fill_type in (None, "none"):
                    cell.fill = PatternFill("solid", fgColor=color)


def sheet_defaults(ws):
    """通用观感：隐藏网格线。"""
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------
# §7 公式助手
# ---------------------------------------------------------------
def col_letter_of(key: str, key_to_letter: dict) -> str:
    """从 key→列字母 映射表取列字母；找不到时抛出带提示的 KeyError。"""
    if key in key_to_letter:
        return key_to_letter[key]
    raise KeyError(
        f"spec 中引用了未知列键 '{key}'；已知键：{sorted(key_to_letter)}"
    )


def build_formula(template: str, row_idx: int, key_to_letter: dict,
                  first_data_row: int, last_data_row: int | None = None) -> str:
    """把公式模板渲染成实际公式。
    支持 {r}（当前行号）、{col.键名}（列字母，键名也可以是表头名）、
    {first}/{last}（数据区首/末行号，用于 RANK 等固定区域引用）。"""
    import re
    out = template.replace("{r}", str(row_idx))
    out = re.sub(r"\{col\.([^}]+)\}",
                 lambda m: col_letter_of(m.group(1), key_to_letter), out)
    out = out.replace(
        "{last}", str(last_data_row if last_data_row is not None else row_idx))
    out = out.replace("{first}", str(first_data_row))
    return out
