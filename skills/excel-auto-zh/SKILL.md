---
name: excel-auto-zh
version: "1.2.0"
agent_created: true
author: WJ
publisher: WJ（个人开发者）
display_name: WJ 的 Excel / WPS 表格自动化工具
display_name_en: Excel / WPS Spreadsheet Automation by WJ
description: "Excel / WPS 表格自动化技能，由开发者 WJ 出品维护。当用户需要生成格式专业的 Excel 报表（财务报表、销售月报、考勤汇总、人力资源批量处理）、批量合并多个工作表或工作簿、解析 openpyxl 打不开的 xlsm/xlsx 文件（含宏、超大、关系受损）时使用本技能；触发语包括：Excel自动化、WPS表格、批量处理表格、xlsm解析、openpyxl、报表生成、多表合并。产物在 Microsoft Excel 与 WPS Office 中均可正常打开，兼容 Windows 与 macOS。"
description_zh: "Excel / WPS 表格自动化技能：openpyxl 生成专业格式报表、零依赖解析 xlsm/xlsx、批量合并多工作表，适用于财务、销售、考勤、人力场景，双办公软件兼容。"
description_en: "Excel / WPS spreadsheet automation skill by developer WJ. Use when the user needs professionally formatted Excel reports (finance, sales, attendance, HR), batch-merging multiple worksheets or workbooks, or parsing xlsm/xlsx files that openpyxl cannot open (macro-enabled, very large, or damaged). Triggers: Excel automation, WPS spreadsheet, batch table processing, xlsm parsing, openpyxl. Output opens correctly in both Microsoft Excel and WPS Office on Windows and macOS."
tags: [excel自动化, WPS, 表格处理, openpyxl, 数据整理, office自动化, excel-automation]
tools: [exec]
metadata:
  requires: "Python 3.8+；仅 create_report.py 与 merge_sheets.py 需要 pip install openpyxl（parse_xlsm.py 零依赖）"
  os: [windows, macos, linux]
  apps: [Microsoft Excel, WPS Office]
---

# WJ 的 Excel / WPS 表格自动化工具

> 开发者：WJ（个人开发者，本技能由 WJ 独立开发、维护并承担全部责任）

用 Python 自动创建、解析和批量处理 Excel / WPS 表格文件，告别重复手动操作。
本技能自带 3 个可执行的命令行脚本和一个共享样式模块：优先直接运行脚本，脚本不覆盖的场景再手写代码（可 `import` 同目录的 `xlsx_style.py` 复用配色与版式）。

## Tools Required

exec（需要能执行 shell 命令、读写文件）

## Setup

```bash
pip install openpyxl
```

验证安装：

```bash
python3 -c "import openpyxl; print(openpyxl.__version__)"
```

- **命令名**：macOS / Linux 用 `python3`；**Windows 下把下文所有 `python3` 换成 `python`**（若无 `python` 则用 `py`）。
- **依赖范围**：只有 `create_report.py` 和 `merge_sheets.py` 需要 openpyxl；`parse_xlsm.py` 纯标准库，**零依赖可直接运行**。缺 openpyxl 时脚本会给出中文提示而不是报错堆栈。
- **技能目录**：下文 `<技能目录>` 指本技能的安装位置（如 `~/.workbuddy/skills/excel-auto-zh`），调用脚本时一律用绝对路径。

- **WPS Office 用户**：本技能生成的 `.xlsx` 文件可直接用 WPS 打开，格式完全兼容（仅使用 WPS/Excel 双端支持的函数与特性）。
- **WPS 自动化另解**：若需驱动"正在运行的 WPS 程序"（打开窗口、点按钮），属于 UI 自动化，不在本技能范围；本技能通过直接读写文件完成自动化，无需打开软件。

## 常见报错速查

| 现象 | 原因 | 处理 |
|---|---|---|
| `ModuleNotFoundError: openpyxl` | 未装依赖 | `pip install openpyxl` 后重试 |
| `PermissionError` 写入失败 | 目标文件正被 Excel/WPS 打开 | 让用户关闭文件后重试，或换输出文件名 |
| `python3: command not found`（Windows） | Windows 无 python3 命令 | 改用 `python` 或 `py` |
| 打开的文件是 `.xls`/加密文件 | OLE 二进制容器，非 zip | 用 parse_xlsm.py 会给出明确中文提示；让用户另存为 `.xlsx` |
| CSV 中文乱码 | 已内置处理 | 输出统一 UTF-8 BOM；输入自动探测 UTF-8/GBK，无需干预 |

## 核心能力与脚本路由

| 用户需求 | 使用脚本 | 深入文档 |
|---|---|---|
| 生成格式化报表（财务/销售/考勤模板） | `scripts/create_report.py` | `references/formatting-standards.md` |
| 解析 openpyxl 打不开/很大的 xlsm、列出宏 | `scripts/parse_xlsm.py` | `references/xlsm-parsing-guide.md` |
| 合并多个工作表/工作簿、生成汇总 | `scripts/merge_sheets.py` | — |

脚本路径写法：调用时使用**绝对路径**，如
`python3 "<技能目录>/scripts/create_report.py" demo --output 销售月报.xlsx`。

---

### 能力一：创建格式化报表

用 openpyxl 生成带专业格式的 Excel 文件，默认遵守财务建模通用的**三色规则**：

| 字体颜色 | 含义 | 示例 |
|---|---|---|
| 蓝色 `0000FF` | 输入项（用户可修改的原始数值） | 目标销售额、汇率、人数 |
| 黑色 `000000` | 计算值（公式结果） | 完成率、合计、排名 |
| 绿色 `008000` | 跨表引用（引用其他工作表） | `='参数'!$B$2` |

快速开始：

```bash
# 1) 一分钟出效果：生成示例销售月报（含参数表、三色规则、排名与提成公式）
python3 "<技能目录>/scripts/create_report.py" demo --output 销售月报.xlsx

# 2) 从 CSV 生成（自动识别数字/日期，自动加合计行；兼容 UTF-8 / GBK 编码）
python3 "<技能目录>/scripts/create_report.py" from-csv 考勤.csv --title "3月考勤" --output 考勤报表.xlsx

# 3) 用 JSON 描述表结构（声明式，无需写代码）
python3 "<技能目录>/scripts/create_report.py" from-json spec.json --output 报表.xlsx

# 4) 交付前自检：工作表清单、行列数、公式数量、颜色图例是否齐全
python3 "<技能目录>/scripts/create_report.py" verify 报表.xlsx
```

JSON spec 最小示例（`style` 决定颜色：`input`=蓝 / `calc`=黑 / `link`=绿；公式里 `{r}`=当前行号、`{col.键名}`=列字母（键名或表头名）、`{first}`/`{last}`=数据区首末行；公式列用 `type_out` 指定结果显示格式）：

```json
{
  "title": "销售月报",
  "sheets": [{
    "name": "销售明细",
    "columns": [
      {"header": "销售员", "key": "name", "type": "text"},
      {"header": "目标",   "key": "target", "type": "currency", "style": "input"},
      {"header": "实际",   "key": "actual", "type": "currency", "style": "input"},
      {"header": "完成率", "type": "percent",
       "formula": "=IFERROR({col.actual}{r}/{col.target}{r},\"-\")"},
      {"header": "排名",   "type": "int",
       "formula": "=RANK({col.actual}{r},{col.actual}${first}:{col.actual}${last})"}
    ],
    "rows": [
      {"name": "张三", "target": 100000, "actual": 123000},
      {"name": "李四", "target": 100000, "actual": 87000}
    ],
    "summary": [
      {"label": "合计", "col": "target", "func": "SUM"},
      {"col": "actual", "func": "SUM"}
    ]
  }]
}
```

手写代码时的硬规则（保证 WPS 与 Excel 双端可用、可读性专业）：

1. 除法一律包 `IFERROR(...,"-")`，共享分母用绝对引用 `$D$2`。
2. 只用兼容函数：`SUM/AVERAGE/IF/COUNTIF(S)/SUMIF(S)/RANK/VLOOKUP/INDEX/MATCH/TEXT/ROUND/EOMONTH`。
   禁用动态数组函数：`XLOOKUP/FILTER/UNIQUE/SORT/LET/LAMBDA/SEQUENCE`。
3. 文本若以 `=` 开头，写入前加 `'` 前缀，防止被当成公式。
4. 表头深蓝底白字、冻结表头、数据区隔行浅色、数字右对齐并套千分位/货币/百分比格式、列宽按内容自适应（中文按 1.7 倍宽度估算）。
5. 表尾加"合计/平均"行（黑字公式），标题下方加三色图例说明。
6. 这些样式已封装在 `scripts/xlsx_style.py`，手写代码直接 `import` 复用，不要散落硬编码颜色。

完整配色、字体、数字格式、打印设置规范 → **读 `references/formatting-standards.md`**。

---

### 能力二：解析复杂 xlsm 文件

openpyxl 打不开或读得很慢的文件（含 VBA 宏、体积大、被其他工具改坏过关系文件的财务模型），用**标准库 `zipfile + xml.etree`** 解析，零第三方依赖，流式读取可处理 100MB+ 文件：

```bash
# 列出工作表、行列数、定义名称、VBA 宏模块、zip 内部结构
python3 "<技能目录>/scripts/parse_xlsm.py" 模型.xlsm --list

# 预览某表前 20 行（自动转换日期、识别共享字符串）
python3 "<技能目录>/scripts/parse_xlsm.py" 模型.xlsm --sheet Sheet1 --head 20

# 全部数据导出：每张表一个 CSV（UTF-8 BOM，Excel/WPS 直接打开不乱码）+ 一个 workbook.json
python3 "<技能目录>/scripts/parse_xlsm.py" 模型.xlsm --all --outdir 解析结果/

# 只要 JSON（含公式文本请加 --formulas）
python3 "<技能目录>/scripts/parse_xlsm.py" 模型.xlsm --json --formulas > data.json

# 列出 VBA 宏模块名（只读解析，绝不启用宏）
python3 "<技能目录>/scripts/parse_xlsm.py" 模型.xlsm --vba
```

- `.xlsx` / `.xlsm` / `.xltm` 通用；输出用 `--head` 快速预览，用 `--all` 拿全量数据。
- 遇到**加密文件或旧版 .xls 二进制**会给出明确提示（建议用 Excel/WPS 另存为 `.xlsx` 再解析）。
- 安全红线：解析是只读的；绝不自动启用宏，不执行 VBA。

OOXML 内部结构、日期序列号换算、常见 openpyxl 报错对照表 → **读 `references/xlsm-parsing-guide.md`**。

---

### 能力三：批量数据处理

自动读取多工作表/多文件、对齐列、合并数据、生成汇总报告。适合考勤汇总、销售数据整理、人力资源批量处理：

```bash
# 合并目录下所有 xlsx 的所有工作表，自动加"来源文件/来源工作表"两列 + 汇总页
python3 "<技能目录>/scripts/merge_sheets.py" --input "考勤表/*.xlsx" --output 月度汇总.xlsx

# 只合并指定名称的工作表；指定表头所在行；只追加不汇总
python3 "<技能目录>/scripts/merge_sheets.py" --input a.xlsx b.xlsx --sheets 4月 5月 6月 --header-row 1 --no-summary --output 汇总.xlsx

# 也可以输出 CSV（UTF-8 BOM）
python3 "<技能目录>/scripts/merge_sheets.py" --input "data/*.csv" --format csv --output 汇总.csv
```

- 列按"首次出现的表头顺序"对齐，缺失列自动补空；数字列自动千分位，日期列自动日期格式。
- 默认生成 `汇总` 页：每个来源文件的行数、各数值列合计，末行总计。

---

## 推荐工作流程

1. **对齐需求**：输出文件名、列结构、颜色语义（默认三色规则）、是否需要公式（活公式）还是纯值。
2. **选路径**：标准报表 → `from-json`；多表合并 → `merge_sheets.py`；复杂/非标结构 → 手写代码并 `import xlsx_style`。
3. **运行脚本**，用 `create_report.py verify` 或重新用 openpyxl 打开抽查单元格值、公式、颜色。
4. **交付**：报告生成了什么文件、颜色图例含义、公式在哪些列、用户可以直接改哪些蓝字单元格。

## Usage

- “帮我用 openpyxl 生成一份财务报表，区分输入项和计算项的颜色”
- “这个 xlsm 文件 openpyxl 打不开，帮我用 zipfile 解析里面的数据”
- “帮我写脚本把 20 个考勤表自动合并成一份月度汇总表”

## Examples

**输入**：帮我生成一份销售月报，包含每个销售员的业绩、完成率、排名，蓝色标注目标值，黑色标注实际值

**输出**：

- 生成 `sales_monthly_report.xlsx`，含格式化表头、自动排名列、汇总行（SUM/AVERAGE）
- 颜色规则：蓝色=目标，黑色=实际，绿色=完成率公式
- 附带可直接运行的 Python 代码，Windows/WPS 环境直接可用

## 兼容性说明

- 操作系统：Windows / macOS / Linux（脚本纯 Python，无系统调用）。
- 办公软件：Microsoft Excel 2010+ 与 WPS Office 均可打开本技能产物；函数仅使用双端兼容子集。
- 编码：所有 CSV 输出带 UTF-8 BOM，Windows Excel/WPS 双击打开中文不乱码；CSV 输入自动探测 UTF-8/GBK。
