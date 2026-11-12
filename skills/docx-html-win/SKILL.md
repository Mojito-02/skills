---
name: docx-html-win
description: 在 Windows 且 Bash 残缺（缺 coreutils）的机器上生产 .docx：①用 tencent-docx 的 html-to-docx 把 HTML 转成 .docx（并修正转换器在 CJK/拉丁边界插入的多余空格）；②往**已排好版的成品 docx 里定点插图 + 图注 + 定向删字**（不必重跑 HTML 流水线）。当需要生成、转换或给本地 Word 文档插图而 setup 脚本跑不动时使用。
agent_created: true
---

# Windows 下 DOCX 的生产与二次加工

## 何时用

- 需要产出 `.docx`（声明函、通知、简介等短篇文书或正式排版文档）。
- 需要**往已定稿的 docx 里插图片**（说明书、申报书配图），且不能重排版、不能动字数账。
- 本机 Bash 缺 coreutils（`ls` / `cat` / `grep` / `dirname` / `rm` / **`cp`** 都报 `command not found`），shell 脚本跑不动。
  复制文件用 PowerShell：`Copy-Item -LiteralPath <src> -Destination <dst> -Force`。

## 判定：走 brief-compose 还是 full_pipeline

`tencent-docx` 插件的 `skills/tdoc-orchestrator/SKILL.md` 是入口。命中 **短篇创作**（无已有文档 ＋ 创作动词 ＋ 目标篇幅 < 1000 字，或体裁为简介/通知/备忘录/声明）→ 直接读 `skills/brief-compose/SKILL.md`，按它执行，**不要再走 orchestrator 的子 Agent 编排**。

## 步骤

### 1. 定落盘路径

```bash
"C:/Users/ASUS/.workbuddy/binaries/python/versions/3.13.12/python.exe" \
  "<plugin_root>/scripts/wb/local/resolve_output_docx_path.py" \
  --workspace "<workspace>" \
  --user-path "<目标的绝对路径.docx>" > <日志文件> 2>&1
```

`<plugin_root>` 形如 `C:\Users\ASUS\.workbuddy\plugins\cache\workbuddy-builtin\tencent-docx\<版本号>`。
返回 JSON 里 `output_docx_path` 即最终路径。**stdout 会被吞时务必重定向到文件再 Read。**

### 2. 写 HTML

- 用 `pt` 字号，不用 `px`。正文默认 12pt，标题 18–22pt。
- 转换器**不做块级样式继承**：`font-family` / `font-size` / `text-align` 必须写在每个 `<p>`/`<h1>`/`<td>` 上，写在 `<body>` 无效。
- 禁 `display:grid`、禁 `<dl>/<dt>/<dd>`；键值对用两列 `<table>`。
- 不要用 `<hr>` 做装饰线（只能变成满版下边框）。
- 文件与目标 docx 同目录、同名。

### 3. 转换（绕开 setup 脚本）

**不要**执行 `setup-html-to-docx.sh` —— 本机 Bash 跑不动（exit 127）。venv 通常已就绪：

```
C:\Users\ASUS\.venv-html-to-docx\Scripts\python.exe
```

（Windows 是 `Scripts\`，不是 `bin/`。）直接转换：

```bash
PY="C:/Users/ASUS/.venv-html-to-docx/Scripts/python.exe"
PLUGIN="C:/Users/ASUS/.workbuddy/plugins/cache/workbuddy-builtin/tencent-docx/<版本号>"
cd "$PLUGIN/skills/html-to-docx/scripts" && \
  "$PY" -m html_to_docx convert "<in.html>" -o "<out.docx>" --page-size A4 \
  >> <日志文件> 2>&1
cat <日志文件>   # 看 {"success": true, ...}
```

成功输出：`{"success": true, "docx_path": "...", "warnings": [...]}`。
遇到 `ModuleNotFoundError: htmldocx` 可忽略 —— 冒烟测试里该模块缺失不影响转换。

### 4. 修补 CJK 空格（正式文书必做）

转换器按 CJK/拉丁边界拆 run，会在全角标点（：、（））前多插空格，例如：
`COS-SafeLink ：国产…`、`学府路  301  号`。

**名称、地址、编号类字段必须与原始材料逐字一致**，用 python-docx 定点修正：

```python
from docx import Document
d = Document(P)
for row in d.tables[0].rows:
    if row.cells[0].text.strip() == "作品名称":
        p = row.cells[1].paragraphs[0]
        p.runs[0].text = "COS-SafeLink：国产操作系统驱动的智慧工厂安全感知与联动控制平台"
        for r in p.runs[1:]:
            r.text = ""
        for extra in list(row.cells[1].paragraphs[1:]):
            extra._element.getparent().remove(extra._element)
d.save(P)
```

改前先 `shutil.copyfile` 备份，改完删掉 `.bak`（本项目要求提交目录干净）。

### 5. 校验 + 交付

读回 docx 确认：表格单元格文本、段落字体字号（`run.font.name` / `run.font.size.pt`）、汉字数、页数。然后 **必须** `present_files` 打开 `.docx` 预览。

## 坑

- venv 里的 python-docx 在 `Scripts\Lib\site-packages`，用 venv 的 python 执行即可，不要另建环境。
- PowerShell `Remove-Item` 退出码偶发为 1（实际删成功），需用 Python 列目录复核。
- 扫描版 PDF（无文本层）取信息：fitz 渲染成 PNG 后**读图**，不要指望 `get_text()`。

---

# 往成品 docx 定点插图（不重跑 HTML 流水线）

场景：docx 已定稿、字数卡着上限、排版已对好，现在要插 N 张图 + 图注。
**不要**回头改 HTML 重转——重转会重排正文、冲爆字数、丢弃已清干净的元数据。
直接改成品 docx，保留一切原有格式。

## 步骤

### 1. 先巡检，拿到"锚点位"

先 dump 段落序（index / style / first_line_indent / alignment / img 数 / 前 70 字），
用**句尾严格匹配 + 断言唯一**定位锚点，不要用硬编码 index（插一张就全错位）。

```python
cand = [p for p in doc.paragraphs if p.text.strip().endswith(suffix)]
assert len(cand) == 1, "锚点不唯一: %s -> %d" % (suffix, len(cand))
```

### 2. 建块 → 搬到锚点后面

python-docx 只能 append，所以"先 append 再搬家"：

```python
def insert_after(anchor_p, new_ps):          # new_ps 顺序 = 想要的顺序
    for p in reversed(new_ps):
        anchor_p._p.addnext(p._p)
```

要**顶替**某段（例如删掉"技术路线图见附图。"并把图放在它原位）：
同样 `addnext` 后 `anchor._p.getparent().remove(anchor._p)`。

### 3. 图段 / 图注段的格式（三个必须）

```python
# 图段
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
pf = p.paragraph_format
pf.first_line_indent = Cm(0)   # ★ 必须显式清 0：正文有 2 字符缩进，不清图会歪
pf.line_spacing = 1.0
p.add_run().add_picture(path, width=Cm(w))

# 图注（宋体 10.5pt 居中）
r = p.add_run(caption); r.font.name = "宋体"; r.font.size = Pt(10.5)
rf = r._element.get_or_add_rPr().get_or_add_rFonts()
rf.set(qn("w:ascii"), "宋体"); rf.set(qn("w:eastAsia"), "宋体"); rf.set(qn("w:hAnsi"), "宋体")  # ★ 三处都要设
```

### 4. 宽度钳制（★ 竖构图最容易翻车）

A4 版心约 14.66cm，正文常用 14cm。但**竖构图按 14cm 宽会纵向溢出页面**：

```python
MAXW, MAXH = 14.0, 11.0
w = min(MAXW, MAXH * (px_w / px_h))     # 用 PIL 读原始像素比
```

### 5. 定向删字（要压字数时）

挑"同一内容在别处已展开"的冗余句。**先确认待删子串完整落在单个 run 内**，
命中就 `run.text = run.text.replace(drop, "", 1)`——零格式损伤；
跨 run 则需重建 run，别硬替换（会把 `“ ”` 引号拆到不同 run 上）。

### 6. 保存（防预览面板占用）

内存重建 zip + 带重试覆写：

```python
buf = io.BytesIO(); doc.save(buf); data = buf.getvalue()
for i in range(8):
    try:
        open(DOCX, "wb").write(data); break
    except PermissionError:
        time.sleep(1.5)
```

### 7. 防重复执行闸门（★ 必加）

脚本会被复用/误跑。开跑前先拒绝：

```python
if any(p.text.strip().startswith("图1 ") for p in doc.paragraphs): sys.exit("ABORT: 已插过图")
if any(p._p.findall(DRAWING_BLIP) for p in doc.paragraphs):      sys.exit("ABORT: 已有图片")
```

### 8. 终检（缺一项都不算完）

| 检查 | 判据 |
|---|---|
| 字数 | 从**最终文件**直读，按章核上限 |
| 图纸 | `wp:inline` 数 == 图数；`wp:anchored` 应为 0（浮动图会乱位） |
| 尺寸 | 读 `wp:extent` 的 cx/cy → 换算 cm，确认高不超版心 |
| 关系 | `r:embed` 的每个 rId 都在 `document.xml.rels` 里；扩展名已在 `[Content_Types].xml` 登记 |
| 良构 | 每个 `.xml`/`.rels` 用 `xml.dom.minidom.parseString` 过一遍 |
| 元数据 | `dc:creator` / `cp:lastModifiedBy` / `app.Application` 为空；禁词零命中 |
| md 镜像 | 若同时维护可编辑版 md，脚本化同步后**逐段比对**，别靠肉眼 |

### 9. 图内文字也要查（★ 最容易漏）

匿名/合规只扫 docx 的 XML 是**不够**的——PNG 图里的文字扫不到。
插图后要**逐张目视**（Read 工具读图）复核两件事：
① 是否含身份信息 / 主机名 / `/home/<user>/` / GitHub 用户名；
② 图内标题与**当前**作品定位语是否一致（旧图常留旧定位语，评审会读成定位漂移）。

