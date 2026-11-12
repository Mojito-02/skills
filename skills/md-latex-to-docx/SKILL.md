---
name: md-latex-to-docx
description: 把「带 LaTeX 公式的 Markdown」高保真转换为排版规范的 Word (.docx)：行内/独立公式转原生 OMML 可编辑公式，Markdown 表格转真实 Word 表格，标题层级/首行缩进/行距按论文规范排。当用户要求把 md 论文、技术文档、报告插入 Word / 输出 docx / 生成 Word 版，且文中含 \\(...\\)、\\[...\\] 公式或 Markdown 表格时使用。触发词：md 转 word、markdown 转 docx、插入word、公式转word、论文排版、生成docx。
description_en: Convert LaTeX-flavored Markdown into a well-formatted Word (.docx) with native OMML equations and real Word tables.
agent_created: true
---

# Markdown（含 LaTeX 公式）→ Word docx

把含公式和表格的 Markdown 转成**可编辑**的 Word：公式是原生 OMML（在 Word 里点开即公式编辑器），
表格是真实 Word 表格（不是文本对齐）。Windows 环境验证通过。

## 核心链路

```
LaTeX  --latex2mathml-->  MathML  --Word 自带 MML2OMML.XSL-->  OMML  --python-docx-->  .docx
```

`MML2OMML.XSL` 随 Office 安装，路径：
`C:/Program Files/Microsoft Office/root/Office16/MML2OMML.XSL`
（也检查 `Program Files (x86)`）。**这是最高保真的写法**——比 pandoc 更可控，
比"公式转图片"更好（可编辑、可缩放、能被 Word 公式编号引用）。

## 环境准备（必须显式官方源）

```bash
MVP="C:/Users/ASUS/.workbuddy/binaries/python/envs/default"
"$MVP/Scripts/python.exe" -m pip install python-docx latex2mathml lxml -i https://pypi.org/simple
```

> 本机 pip 默认走清华镜像且缺包，**必须** `-i https://pypi.org/simple`。
> `numpy` 在托管 venv 里没有；纯 docx 转换用不到。

## 用法

```bash
cd <工作目录>
"$MVP/Scripts/python.exe" <技能目录>/build_docx.py <输入.md> <输出.docx>
```

脚本已参数化，输入输出任意路径；改排版基线直接改脚本里的 `configure_styles` / `add_body`。

## 铁律（踩过的坑）

1. **`\texttt{...}` 不能当公式渲染**。它装的是代码/路径/文件名/变量名（`video\_id`、
   `allow_pickle=False`、`run_config.json`），当 OMML 渲染会得到斜体数学变量、下划线转义错乱。
   规则：**整个公式恰为纯 `\texttt{...}`** → 还原为普通**等宽文本**（Consolas）；
   其余才走 OMML。脚本里 `plain_from_texttt()` 负责这个判定。
2. **行内公式要与 `w:t` 混排**：直接把 `<m:oMath>` append 到段落 XML 即可，
   Word 会按文档顺序自动排版，**不要**为行内公式单独建段落。
3. **独立公式包 `<m:oMathPara>`**（居中）；命名空间只需 `xmlns:m`（`nsdecls("m","w")` 更稳）。
4. **`<` `>` 会破坏 MathML 解析**（如 `features/<video\_id>/<clip\_id>.npz`）。
   脚本已有兜底：失败时把 `<`/`>` 换成 `\lt`/`\gt` 重试，终归最后退回等宽文本。
5. **表格单元格里的公式**同样走行内 OMML；表头加粗要靠**逐 run** 设 `bold`，
   设单元格整体样式会被 run 覆盖。
6. **Office 未安装时**：`MML2OMML.XSL` 不存在。降级方案——公式退化为等宽文本，
   或改用 pandoc（若有）输出 docx `--from markdown+tex_math_dollars`。

## 排版基线（论文/报告，可改）

| 元素 | 规格 |
|---|---|
| H1 | 黑体 16pt、居中、段后 12pt |
| H2 / H3 | 黑体 14pt / 12pt、左对齐 |
| 正文 | 宋体 + Times New Roman 12pt（小四）、**首行缩进 2 字符**、1.5 倍行距 |
| 表格 | Table Grid、整体居中、表头加粗 10.5pt（五号）、按 md 对齐标记设列对齐 |
| 表题 / 图注 | 宋体 10.5pt、居中 |
| 页边距 | 上下 2.5cm、左右 2.8cm |

## 转换后的自检（必做）

断言以下四项，全部通过才算成功：

```python
from docx import Document
d = Document(out)
xml = d.element.xml
assert xml.count('<m:oMath') >= <源文公式数>, '公式丢失'
assert len(d.tables) == <源文表格数>, '表格数量不符'
# 表格行列数逐张比对
assert '\\' not in ''.join(<所有 w:t 文本>), '有未还原的 LaTeX 转义残留'
```

再抽检：`oMathPara` 数 = 独立公式数；各表 `len(rows) x len(columns)` 与 md 一致；
`w:t` 拼接文本里关键字（数字、专有名词）无缺失。

## 注意

- **先查目标文档再动手**：主论文文件名常与"赛题原文/模板"混淆（本次 E 题根目录的
  `复杂场景下多模态情感识别的数学建模与算法设计.docx` 内容其实是赛题原文，不含问题一正文）。
  插入前用 `docx` 读一遍段落确认落点，或直接问用户。
- 图片（`![...](...)`）：默认**只留居中占位 + 图注**，避免路径失效或体积膨胀；
  需要真图时改用 `doc_insert_image`（本地编辑通道）或 python-docx `add_picture`。