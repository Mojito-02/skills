# 报表格式规范（三色规则与 Excel/WPS 兼容）

生成专业报表时按本规范执行；`scripts/xlsx_style.py` 是这些规范的代码化实现，
手写代码直接复用，不要散落硬编码颜色。

## 1. 三色规则（财务建模惯例）

| 颜色 | 色值 | 含义 | 例子 |
| --- | --- | --- | --- |
| 蓝色 | 0000FF | 输入项：用户可修改的原始数值 | 目标值、单价、考勤天数、假设参数 |
| 黑色 | 000000 | 计算值：本表公式算出的结果 | 完成率、排名、合计行、同比 |
| 绿色 | 008000 | 跨表引用：公式引用了其他工作表/文件 | 引用"参数"表的提成率 |

- 字体其余属性一致（默认 11pt），只变颜色，保持版面统一。
- 每张表第一行下方放一行小字图例：`图例：蓝色=输入项 / 黑色=公式计算 / 绿色=跨表引用`（用 `legend()`）。
- 交付时向用户说明"蓝字单元格可以直接改，公式会自动重算"。

## 2. 版式

- 大标题：合并首行，14pt 加粗、深蓝 1F4E78 字色、左对齐。
- 表头：深蓝底 1F4E78 + 白字加粗、居中、细边框、行高 24。
- 数据区：细边框 BFBFBF、隔行浅色 EEF3F9（`band_rows()` 已实现）；文本列左对齐，数字右对齐。
- 合计行：加粗、顶部 medium 粗线 + 细底边、浅蓝底 DDEBF7、公式用黑字（`summary_row()` 已实现）。
- 冻结表头（freeze_panes = 表头下一行 A 列）+ 自动筛选（表头行到数据末行）。
- 列宽按内容自适应，中日韩字符按 1.7 倍宽度估算；过窄时表头开 `wrap_text`。

## 3. 数字格式速查

| 场景 | 格式码 | 示例 |
| --- | --- | --- |
| 金额 | `¥#,##0.00` 或 `#,##0.00` | ¥12,345.00 |
| 整数/人数/排名 | `#,##0` 或 `0` | 1,234 |
| 百分比 | `0.0%` | 87.5% |
| 日期 | `yyyy-mm-dd` | 2026-09-06 |
| 日期时间 | `yyyy-mm-dd hh:mm` | 2026-09-06 22:28 |
| 编码/工号/手机号 | `@`（文本） | 0012、13800000000 不丢前导零 |

写值时同步 `number_format`；不要把手机号、工号写成数字。

## 4. 公式兼容性（Excel 与 WPS 双端）

✅ 放心用：`SUM AVERAGE COUNT COUNTA MAX MIN ROUND IF IFERROR AND OR
VLOOKUP HLOOKUP INDEX MATCH RANK SUMIF SUMIFS COUNTIF COUNTIFS SUMPRODUCT
TEXT LEFT MID RIGHT LEN TRIM TODAY DATE EOMONTH SUBTOTAL OFFSET INDIRECT`

❌ 不要用（老版 Excel/WPS 部分版本不支持，WPS 兼容模式报错）：
`XLOOKUP FILTER SORT UNIQUE LET LAMBDA TEXTJOIN IFS MAXIFS MINIFS CONCAT`
以及一切动态数组溢出写法。替代：XLOOKUP→INDEX/MATCH；TEXTJOIN→&拼接；
IFS→嵌套 IF；MAXIFS→MAX(IF()) 数组式改 SUMPRODUCT/MAX+IF。

固定写法：

- 除法必须防零：`=IFERROR(实际/目标, "-")`
- 排名：`=RANK(C4,$C$4:$C$11)`（RANK.number 参数即当前单元格）
- 跨表引用：表名一律加单引号，中文/空格/数字开头都安全：`='参数'!$B$4`
- 文本若以 `=`、`+`、`-`、`@` 开头，前置单引号 `'` 写入，避免被当公式（`safe_value()`）
- 公式只写在首行数据之后的每一行，不要用数组溢出

## 5. WPS 与跨平台注意事项

- 字体优先级：微软雅黑 → 苹方 → Noto Sans CJK（`xlsx_style.FONT_NAME` 自动解析），三家都装齐靠回退链。
- `.xlsx` 双端完全兼容；`.xlsm` 只在确需宏时交付，openpyxl 写宏需 `keep_vba` 读模板改造。
- CSV 交付一律 UTF-8 带 BOM（`utf-8-sig`），否则 Windows Excel 打开中文乱码。
- 打印设置：A4、横向（宽表）、缩放适配一页宽、顶端标题行重复：`ws.print_title_rows = '1:3'`。

## 6. 交付前自检清单

1. `python3 scripts/create_report.py verify 输出.xlsx`：工作表数、行列数、公式个数、冻结是否生效。
2. 用 openpyxl 重新打开抽查 2-3 个单元格：值、公式文本、`font.color` 是否符合三色规则。
3. 合计行与明细行数量一致；百分比列格式是 `0.0%` 而非小数。
4. 文件名有业务含义（如 `销售月报_2026-08.xlsx`），不含"新建/最终版v2"。
5. 回复用户：文件里哪些列是蓝字可改、公式在哪些列、颜色图例含义。
