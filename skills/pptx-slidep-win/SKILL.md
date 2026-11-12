---
name: pptx-slidep-win
description: 在 Windows 且 Bash 残缺（缺 coreutils）的机器上，用 slidep 生成/编辑 PPT 页面，并用 PowerPoint COM 渲染成 PNG 做版面校验。当需要产出 .pptx 页面、或需要确认 PPT 实际渲染效果时使用。
agent_created: true
---

# Windows 下用 slidep 生成 PPT 页面

## 何时用

- 需要产出 `.pptx`（整篇或单页）。
- 本机 Bash 缺 coreutils（`mkdir`／`ls`／`cat`／`dirname`／`rm` 全部 `command not found`），shell 脚本跑不动。
- 需要**目视确认** PPT 实际渲染效果（lint 只查溢出，不查美观）。

## 工具位置

```
NODE   = C:\Users\ASUS\.workbuddy\binaries\node\versions\22.22.2-3\node.exe
SLIDEP = C:\Users\ASUS\.workbuddy\binaries\node\versions\22.22.2-3\node_modules\@tencent\slidep\dist\slidep.js
```

`.cmd` shim 虽在 PATH，但 Bash 调不动 → **一律用 node 直接跑 slidep.js**。

## 流程

### 0. 先查目标 PPT 的尺寸与版式

若页面要并入既有 PPT，先读它的 `slide_width` / `slide_height`（python-pptx，Anaconda 解释器有），确保新页同尺寸。

### 1. 建项目目录（用 Python，别用 mkdir）

```
<workspace>/<project_dir>/
├── DESIGN.md / STORY.md
├── slides/01.slide
└── <project_dir>.pptx
```

### 2. 创建空白 PPTX

```bash
"$NODE" "$SLIDEP" create "F:/path/文件名.pptx"
# → {"ok":true,"file_id":"...","latest_version":0}
```

创建后**立即** `present_files` 打开预览（流程硬约束）。

### 3. 写 .slide

完整语法见 `tencent-pptx` 插件的 `references/component-*.md`（box / text / table / slide）。要点：

- 画布固定 **1280×720**，内容**绝不允许溢出**，动笔前先算总高度。
- `<Slide>` 是容器（继承 Box），文件最后一个表达式必须是它；**禁止** `import` / `export`。
- 换行用 `<br />`，局部样式用 `<span style={{...}}>`；**没有** `<strong>` / `<u>`。
- Box **只支持 flex**，禁 `display:grid`，**不支持 `calc()`**。
- 表格的 `style` / `defaultTextStyle` / `defaultCellStyle` 都是必填。
- **表格式布局建议手工搭行**（`Box flexDirection:'row'` + 固定 `width` 的单元格），比 `Table` 更好控列宽。
- **不要用 emoji**（✅⚠️ 在 PPT 字体里易成豆腐块）→ 改用 `●` `◐` `○` `△`。

### 4. 校验 → 写入（循环）

```bash
"$NODE" "$SLIDEP" lint "$PPTX" --dsl-file "slides/01.slide"       # 需 ok:true, diagnostics:[]
"$NODE" "$SLIDEP" upsert-dsl "$PPTX" --dsl-file "slides/01.slide" --page-index 0
# 替换已有页用 --page-index N（0-based）；追加则省略该参数
```

lint 不通过 → 改 `.slide` → 重跑，直到 `diagnostics: []`。
改页面时**直接重写整个 .slide 文件**，比打补丁式 Edit 安全（曾因 Edit 少写一个开标签导致结构损坏）。

### 5. 渲染成 PNG 目视校验（关键，别跳过）

本机**无 LibreOffice**（`pptx2image.py` 用不了），但**有 PowerPoint（Office16）**。用 COM 导出：

```python
import win32com.client as wc
app = wc.Dispatch("PowerPoint.Application")
pres = app.Presentations.Open(PPTX, ReadOnly=True, WithWindow=False)
pres.Slides(1).Export(r"...\slide01.png", "PNG", 1920, 1080)
pres.Close(); app.Quit()
```

用 **Anaconda python**（`F:\Anaconda3\python.exe`，含 pywin32）。`WithWindow=False` 不弹窗。
导出后 Read 图片，逐项核对：列宽、行高、文字是否溢出／断行难看、配色对比。
**先确认目标 PPT 没被用户打开**（同目录有 `~$` 临时文件即为打开中）。

## 坑

- 所有命令的 stdout 常被吞 → **一律重定向到文件再 Read**。
- PowerShell 的 `Remove-Item` 退出码偶发为 1（实际删成功），需用 Python 列目录复核。
- `slidep lint` 通过 ≠ 好看。必须渲染看图。
- 一页独立生成的 PPT 页，并入主 PPT 时页码／页脚由主 PPT 版式接管，**不要在单页里写死页码**。
