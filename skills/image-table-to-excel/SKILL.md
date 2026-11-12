---
name: image-table-to-excel
display_name: 图片表格转 Excel
display_name_en: Image Table to Excel
description: "把截图或照片里的表格（WPS/Excel 界面截图、纸质表格拍照、微信图片）准确转录为结构化 Excel。当用户提供表格图片并要求「整理到 Excel / 整理成表格 / 确保准确」时使用。核心方法：整图读结构 → 按列裁切放大精读 → openpyxl 生成带样式表格 → 存疑点回报用户。"
description_zh: "把截图或照片里的表格准确转录为结构化 Excel，适用于 WPS/Excel 界面截图、纸质表格拍照、微信聊天图片等来源。核心方法：整图先读表格结构，再按列裁切放大逐格精读，用 openpyxl 生成带样式的表格，最后把存疑单元格整理成清单回报用户核对。"
description_en: "Transcribes tables from screenshots and photos (spreadsheet UI captures, phone photos of paper forms, chat images) into structured Excel workbooks. Core method: read the whole-image table structure first, crop and magnify column by column for close reading, generate a styled workbook with openpyxl, and report any uncertain cells back to the user for verification."
category: document-processing
version: 1.0.0
agent_created: true
---

# 图片表格 → Excel 转录

## 适用场景
- 用户给几张表格截图/照片，要求整理成 Excel
- 常见来源：WPS/Excel 界面手机拍屏、微信转发的截图、纸质表格拍照
- 特点是分辨率低、有摩尔纹、水印遮挡、透视畸变，**不能靠印象填内容**

## 核心原则：两遍读图

**第一遍——整图读结构（每图 1 次 Read）**
目标只有三个：确认列头文字、确认行号/序号体系、确认哪些列有数据。
不要在这一遍纠结具体文字。

**第二遍——按列裁切放大精读**
关键公式（决定性）：

```
文字在最终渲染中的像素高度 ≈ 原图文字高度 × 1240 / 裁切宽度
```

Read 工具按长边缩放显示，所以**裁得越窄，字越清楚**。全宽切条带（只切高度）几乎没用。
- 裁切宽度 940–1200px → 可读性最佳
- 裁切宽度 1700px（全宽）→ 中文字约 12px，基本误读率很高

**预处理参数（Pillow）**：
```python
c = im.crop(box)
c = c.resize((int(c.width*2.4), int(c.height*2.4)), Image.LANCZOS)
c = ImageEnhance.Contrast(c).enhance(1.5)
c = ImageEnhance.Sharpness(c).enhance(2.0)
```
放大倍率 2.2–2.6 均可，再大不增益（Read 会缩回去）。

## 必须避开的坑

### 坑 1：把 WPS 行号当成 A 列序号（最常见）
WPS 表格最左侧有**行号列**，紧挨着它是 **A 列**。两者都是纯数字，在截图里极易混淆。
- 行号从 1 开始且第 1 行通常是表头 → **行号 = 序号 + 1**
- 判断方法：放大列标区域，看"A / B / C"字母标在哪一列上方；字母标下方的才是 A 列
- 后果：整套序号错位，且多张图之间会"看起来不连续"

### 坑 2：多张图的序号重叠
同一份表的分屏截图，相邻两张会有一段重叠区（重复行）。
必须核对**重叠行的内容是否一致**——如果同序号不同内容，说明这几张图可能来自
不同版本或不同工作表，**不要硬拼**，要回报用户确认。

### 坑 3：靠像素投影检测表格线
手机拍屏的表格线很浅、有水印干扰，水平/垂直投影找线基本不可靠。
不如直接放大目视。

### 坑 4：合并单元格
大项/小项列多为合并单元格，文字垂直居中。所以同一个"大项"名可能出现在
它覆盖行区间的中间位置，而不是第一行。按覆盖区间推断，不要按显示位置。

### 坑 5：Python 环境
`openpyxl` 常不在默认 Python 里。先探测：
```bash
python -c "import openpyxl; print(openpyxl.__version__)"
```
本机可用环境：`~/.workbuddy/binaries/python/envs/default/Scripts/python.exe`（3.1.5）、
系统 Python 3.11（3.1.0）。**管理版 3.13.12 没有 openpyxl。**

## 生成 Excel 的模板

按原表列结构建表，保留原表的合并语义（用重复填值代替合并，便于筛选）：

```python
HEADERS = ['序号', '大项', '小项', '安全需求', '裁剪结果', '概设实现描述', '详设实现描述', '备注']
WIDTHS  = [6, 16, 22, 95, 16, 20, 20, 14]

THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

for c, (h, w) in enumerate(zip(HEADERS, WIDTHS), 1):
    cell = ws.cell(row=1, column=c, value=h)
    cell.font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='1F4E79')
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.column_dimensions[cell.column_letter].width = w
```
要点：`wrap_text=True` 必须开（需求条文很长）；`freeze_panes='A2'`；
`auto_filter.ref` 便于用户自查；隔行浅色填充提升可读性。

## 交付前自检
1. 读回 xlsx，打印总行数 + 表头 + 按大项分组计数
2. 抽查首行、末行、每个大项的首行
3. 核对大项计数之和 == 总行数（能发现漏行）

## 必须回报用户的存疑点
转录任务里"确保准确"是硬要求。以下必须在最终回复中列出，不要默默按猜测填：
- 序号体系冲突或重叠
- 截图中为空的列（说明是原表没填，不是我漏了）
- 模糊处的多个候选读法（如"图形验证码"vs"图形认证码"、"重设"vs"重置"）
- 疑似缺字、重复字、语句不通处

临时裁切图放在 `<源目录>/_crops/`，交付时告知用户可删。

---

## 后续：往已存在的本地 xlsx 批量回写单元格

用户常见后续动作是"在这个表里再补几列"（例如按某一列的值填写另外两列）。此时**不要**用 openpyxl 直接改盘上的文件（用户很可能正开着它，会冲突），走本地编辑通道：

### 通道选择（本机实测结论）
| 通道 | 是否可用 | 说明 |
|---|---|---|
| `sheet-agent` 子代理 | ❌ **不可用** | 无法解析 Windows 盘符路径（`INVALID_LOCAL_PATH`）；传 `/WorkBuddy/...` 虽过存在性校验，但 editorsdk 侧返回 `No workbook open` |
| `tencent-local-office-edit` 的 edsdk | ✅ **走这条** | file_id **就是文件的完整 Windows 路径** |

### 操作要点
1. `get_pool_status` 看是否已打开；未打开则 `open_file(file_path=..., file_type='sheet')`
2. `sheet_get_sheet_info` 拿 `sheet_id`（通常是 `000001`），`sheet_get_used_range` 拿数据边界
3. `sheet_get_cell_data(..., return_csv=True)` 一次读全表，落盘成临时 csv 再 Read，避免撑爆上下文
4. 批量写多列多行用 **`sheet_set_range_value_by_csv`**（`start_row`/`start_col` 均 **0-based**，空单元格用连续逗号，单次 ≤20000 格）——比逐个 `sheet_set_cell_value` 快一个数量级
5. **CSV 文本里不要出现半角双引号**，会与 CSV 引号转义打架；中文引号用「」或干脆不加
6. 写完必须显式调 `save_file`，否则只改内存
7. **保存后编辑器实例可能被回收**，后续再读会报 `workbook is not open`——重新 `open_file` 等 2–3 秒再读即可
8. 回读校验：按条件分类统计（如"执行且已填 / 不执行且填 / 异常"），异常行数为 0 才算通过

### 填表内容的组织方式
把内容写成 `{行号: (甲列文本, 乙列文本)}` 的字典放在脚本里，脚本负责校验"无遗漏行"再写入。
这样后续要改文案，改脚本重跑即可，不必手工点表格。
