# CHANGELOG — wj-excel-auto-zh

## 1.2.0（2026-09-08）

针对市场审核"应用名称与申请主体无明确归属，暂不支持通过认证"的归属类驳回。

### 归属修复

- 应用更名：`excel-auto-zh` → `wj-excel-auto-zh`，展示名 →「WJ 的 Excel / WPS 表格自动化工具」，符合"个人主体：昵称/姓名 + 功能标识"的命名规则。
- frontmatter 新增 `author: WJ` 与 `publisher` 署名，description 中英文均注明"由开发者 WJ 出品维护"。
- SKILL.md 正文增加开发者署名与责任声明；MARKETPLACE.md 新增「申请主体归属说明」（名称归属 + 主体关联关系），可直接复制到上架表单备注栏。
- 归属标识一致性：应用名称 = frontmatter author = 技能包目录名 = 技术标识名（`wj-excel-auto-zh`），审核可交叉核验。

## 1.1.0（2026-09-07）

针对市场审核"沙箱评测表现未达标"反馈的针对性修复与增强。

### 修复

- **frontmatter 校验失败**：`description` 原使用 YAML 折叠语法 `>`，本地校验器会将其识别为含尖括号的非法值；改为单行规范字符串，并补充 `agent_created: true`。
- **merge_sheets.py**：汇总页 SUMIF 跨表引用 `合并数据!$A:$A` 未加单引号，违反本技能 formatting-standards 中"跨表引用表名一律加单引号"的规范；已修复为 `'合并数据'!$A:$A`。
- **create_report.py demo**：提成公式硬编码 `D{r}` 列号，改为由 key→列字母映射生成；参数表图例改为合并单元格，避免撑宽 A 列。
- **xlsx_style.auto_width**：改为只从表头行开始测宽并跳过合并单元格锚点，标题/图例长文本不再把首列撑到宽度上限。
- **create_report.py from-json**：无数据行时不再生成引用空区域的汇总公式；行数据支持数组形式（与既有注释语义对齐）。

### 增强

- 依赖缺失降级：create_report.py / merge_sheets.py 在缺 openpyxl 时输出中文安装提示并退出（exit 2），不再抛裸堆栈；parse_xlsm.py 保持零依赖。
- 新增 `xlsx_style.band_rows()`：数据区隔行浅色，create_report / merge_sheets 全部接入（此前文档承诺但无实现）。
- parse_xlsm.py：`--head` 可单独使用（预览全部工作表），不再被静默忽略。
- SKILL.md：新增"常见报错速查"表（缺依赖 / 文件被占用 / Windows 无 python3 / 加密文件 / 编码），明确 Windows 下用 `python` 替代 `python3`。
- 全链路实测：demo、from-json（含 SKILL.md 示例 spec）、from-csv（UTF-8 与 GBK）、verify、merge（xlsx/csv）、parse_xlsm（--list / --head / --json --formulas / --all / --vba / 加密文件提示）在 Windows + Python 3.13 + openpyxl 3.1.5 下全部通过。

## 1.0.0（2026-09-06）

- 首版：三色规则报表生成（demo / from-json / from-csv / verify）、零依赖 xlsm/xlsx 解析（含 VBA 模块清单、大文件流式）、多工作簿/工作表批量合并。
