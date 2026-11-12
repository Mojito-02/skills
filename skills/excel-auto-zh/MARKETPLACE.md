# WORKBUDDY 市场上架信息（wj-excel-auto-zh）

> 本文件是上架表单的填写内容，直接复制对应字段即可；技能本体见 SKILL.md。

---

**name**

```
wj-excel-auto-zh
```

**version**

```
1.2.0
```

**display_name（中文展示名）**

```
WJ 的 Excel / WPS 表格自动化工具
```

**display_name_en（英文展示名）**

```
Excel / WPS Spreadsheet Automation by WJ
```

**author / 开发者署名**

```
WJ（个人开发者）
```

**description**

```
Excel / WPS 表格自动化技能，由开发者 WJ 出品维护。当用户需要生成格式专业的 Excel 报表（财务报表、销售月报、考勤汇总、人力资源批量处理）、批量合并多个工作表或工作簿、解析 openpyxl 打不开的 xlsm/xlsx 文件（含宏、超大、关系受损）时使用本技能；触发语包括：Excel自动化、WPS表格、批量处理表格、xlsm解析、openpyxl、报表生成、多表合并。产物在 Microsoft Excel 与 WPS Office 中均可正常打开，兼容 Windows 与 macOS。
```

**description_zh（中文描述）**

```
同上 description。
```

**description_en（英文描述）**

```
Excel / WPS spreadsheet automation skill by developer WJ. Use when the user needs professionally formatted Excel reports (finance, sales, attendance, HR), batch-merging multiple worksheets or workbooks, or parsing xlsm/xlsx files that openpyxl cannot open (macro-enabled, very large, or damaged). Triggers: Excel automation, WPS spreadsheet, batch table processing, xlsm parsing, openpyxl. Output opens correctly in both Microsoft Excel and WPS Office on Windows and macOS.
```

**tags**

```
excel自动化, WPS, 表格处理, openpyxl, 数据整理, office自动化, excel-automation
```

**Tools Required**

```
exec
```

---

## 申请主体归属说明（认证备注，直接复制到上架表单备注/说明栏）

**一、名称归属说明**

本技能名称为「WJ 的 Excel / WPS 表格自动化工具」，技术标识名为 `wj-excel-auto-zh`。名称中的「WJ」即申请主体（个人开发者 WJ）的署名标识，符合"个人主体：昵称/姓名 + 功能标识"的命名规则，且应用名称、frontmatter author 署名、技能包目录名、技术标识名四处完全一致，归属关系清晰可查。

**二、主体与技能的关联关系说明**

1. 本技能由申请主体 WJ 个人独立开发、维护，并承担全部运营与责任，无第三方版权内容。
2. 技能提供的具体服务内容：基于 Python/openpyxl 的 Excel/WPS 表格自动化处理，包含三项功能——①生成格式专业的报表（三色规则/公式/汇总）；②零依赖解析 xlsm/xlsx 文件（含宏清单、大文件流式读取）；③批量合并多工作簿/多工作表并生成汇总。
3. 服务性质为通用办公效率工具，不涉及医疗、金融、教育等特殊行业资质，不采集任何用户数据，全部计算在本地完成。
4. 开发者承诺遵守平台规范与中华人民共和国法律法规，按照申请主体的实际能力提供本技能服务。

---

**核心能力**

1. 创建格式化报表：用 openpyxl 生成带专业格式的 Excel 文件，支持自定义颜色（蓝色=输入项、黑色=计算值、绿色=跨表引用），适合财务报表、数据汇总等场景。
2. 解析复杂 xlsm 文件：用标准库 zipfile + xml.etree 解析 openpyxl 无法处理的复杂 xlsm 文件（含 VBA 宏、大型财务模型），提取数据无需安装额外依赖。
3. 批量数据处理：自动读取多工作表、合并数据、生成汇总报告，适合考勤汇总、销售数据整理、人力资源批量处理等重复性工作。

**Setup**

```
pip install openpyxl
```

仅 create_report.py 与 merge_sheets.py 需要 openpyxl；parse_xlsm.py 零依赖可直接运行。WPS Office 用户：openpyxl 生成的 .xlsx 文件可直接用 WPS 打开，格式完全兼容。

**Usage**

- 帮我用 openpyxl 生成一份财务报表，区分输入项和计算项的颜色
- 这个 xlsm 文件 openpyxl 打不开，帮我用 zipfile 解析里面的数据
- 帮我写脚本把 20 个考勤表自动合并成一份月度汇总表

**Examples**

输入：帮我生成一份销售月报，包含每个销售员的业绩、完成率、排名，蓝色标注目标值，黑色标注实际值

输出：生成 sales_monthly_report.xlsx，含格式化表头、自动排名列、汇总行（SUM/AVERAGE）；颜色规则：蓝色=目标，黑色=实际，绿色=完成率公式；附带可直接运行的 Python 代码，Windows/WPS 环境直接可用。

---

## 包内容清单

```
wj-excel-auto-zh/
├── SKILL.md                          技能主文件（frontmatter + 归属署名 + 能力路由 + 报错速查）
├── MARKETPLACE.md                    本文件：市场上架文案 + 申请主体归属说明
├── CHANGELOG.md                      版本变更记录
├── scripts/
│   ├── xlsx_style.py                 共享样式模块（三色规则/表头/列宽/隔行浅色/合计行）
│   ├── create_report.py              报表生成：demo / from-json / from-csv / verify
│   ├── parse_xlsm.py                 零依赖解析 xlsx/xlsm（含 VBA 模块清单）
│   └── merge_sheets.py               批量合并多工作簿/多工作表 + 汇总页
└── references/
    ├── formatting-standards.md       格式规范：三色规则/数字格式/公式兼容/自检
    └── xlsm-parsing-guide.md         OOXML 结构/日期序列号/VBA/大文件流式解析
```

## 版本变更记录

- **v1.2.0（2026-09-08）**：针对"应用名称与申请主体无明确归属"审核意见——应用更名为「WJ 的 Excel / WPS 表格自动化工具」（name: `wj-excel-auto-zh`），frontmatter 增加 author/publisher 署名，description 注明开发者，新增申请主体归属说明（名称归属 + 主体关联关系），四处归属标识完全一致。
- **v1.1.0（2026-09-07）**：修复沙箱评测问题——frontmatter 规范化、依赖缺失降级提示、from-json 顶层 title、SUMIF 跨表引用引号、parse_xlsm OLE 报错崩溃、列宽自适应、隔行浅色、CSV 类型还原、前导零/手机号文本保护，全链路实测通过。
- **v1.0.0（2026-09-06）**：首版。
