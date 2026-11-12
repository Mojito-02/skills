---
name: ppt-polish
description: 美化、体检与批量精修一份已有的 PPT / .pptx 演示文稿。当用户拿来一份现成 PPT 说"美化一下""优化排版""检查这份 PPT 有什么问题""统一风格""批量改配色""PPT 太丑/AI 味太重/颜色太多/风格不统一"时用。能力：① python-pptx 几何与配色体检（越界/文字溢出/元素重叠/色板合规/字体字号/页脚页码/文案自相矛盾）② 用 WPS Office 或 Microsoft PowerPoint 的 COM 接口把 PPT 渲染成 PNG 做真正的视觉复核，无需安装 LibreOffice ③ 按页号范围批量收敛配色、重编页码、修正文案，并自动处理"AI 生成 PPT 色彩过多"类问题。触发词：PPT美化、美化PPT、PPT优化、PPT体检、PPT检查、PPT排版、PPT精修、统一PPT风格、批量改PPT、pptx精修、去AI味、AI味太浓、颜色太多、风格不统一、检查PPT有什么问题、PPT视觉复核。
agent_created: true
version: 1.0.0
display_name: PPT 美化体检
display_name_en: PPT Polish & Audit
description_zh: 给已有的 PPT 做「美化 + 体检」的批量精修：先全量检查版式毛病（元素越界、文字溢出、互相压盖、字体字号、页脚页码、文案矛盾），再把 PPT 渲染成图片做真实视觉复核，最后按页批量统一配色、修正文案。特别擅长去掉 AI 生成 PPT 的「AI 味」（颜色太多、同色族难区分、占位框刺眼），让 PPT 更好看、更专业、风格统一。适合：美化PPT、优化排版、统一风格、检查PPT问题、批量改配色。运行需要 Windows + WPS Office 或 Microsoft PowerPoint。搜索关键词：PPT美化、PPT优化、PPT体检、PPT精修、去AI味、PPT排版。
description_en: Polish, audit and batch-refine an existing PPT/.pptx deck. Runs full layout and palette checks (out-of-bounds elements, text overflow, overlapping shapes, fonts, footers and page numbers, conflicting copy), renders the deck to PNG via WPS Office or Microsoft PowerPoint COM (no LibreOffice needed) for true visual review, then batch-unifies colors and fixes text per page range. Especially good at removing the "AI look" from AI-generated decks (too many colors, indistinguishable shades, noisy placeholder boxes) to make slides cleaner, more professional and consistent. Great for polishing an old deck, unifying style, checking for layout problems, or batch color cleanup. Requires Windows with WPS Office or Microsoft PowerPoint.
---

# PPT 美化体检

把「一份已经存在的 PPT」变成「排版合格、风格统一、看不出 AI 味」的成品。

## 这个 skill 是干什么的

- **输入**：一个现成的 `.pptx` 文件（人做的或 AI 生成的都行）
- **流程**：先体检找硬伤 → 渲染成图看真实效果 → 按规则批量精修 → 复检对比
- **输出**：修好的 `.pptx`（默认另存新文件，不覆盖原件）
- **特点**：不只靠数据下判断，**渲染出来用眼睛复核**；颜色处理**先问用途再决定去留**（语义色 ≠ 装饰色）

## 适用边界

- **用**：已有 `.pptx` 需要体检 / 优化 / 统一风格 / 修硬伤 / 出视觉复核 / 去 AI 味
- **不用**：
  - 从素材**从零生成** PPT（那是专门的 PPT 生成能力，不在本 skill 范围）
  - 需要逐页在编辑器里精细交互微调（请用能所见即所得编辑 PPT 的工具）
  - 非 Windows 环境、且没有装 WPS Office 或 Microsoft PowerPoint（COM 渲染依赖它们）

## 四条核心原则（都是踩过的坑）

1. **先渲染，再动手。** 不要只靠数据下判断。实测案例里，配色方案在数据层面「0 违规」，渲染出来才发现第 4 模块的深蓝跟第 1 模块的蓝**几乎同色**。几何检查（越界 / 溢出 / 重叠）永远替代不了看一眼。
2. **颜色先问用途，再决定去留。** 红 / 绿常常是**语义色**（反面示范、✅ 解法、红线警示），不是装饰色。先跑 `context.py` 看它长在什么元素上，再决定是收敛还是保留。
3. **区分模块要用不同色相，不能靠明度。** 同色族两档（深蓝 vs 蓝）在小尺寸、投影、远看时区分不开。要几个模块就给几个色相。
4. **原文件被占用就输出新文件。** 别硬覆盖，用户可能正开着它。

## 环境要求

- Windows + 装有 **WPS Office** 或 **Microsoft PowerPoint**（用于渲染复核，二选一即可）
- Python 3.9+，`pip install python-pptx`（几何检查与批量修改用；渲染走 COM 不需要 LibreOffice）

## 标准流程

### 步骤 0 · 前置（必做）

```bash
# 备份：改之前一定留后路
cp "原文件.pptx" "原文件_备份_$(date +%Y%m%d).pptx"

# 检查占用：目录里出现 ~$xxx.pptx 锁文件 = 编辑器正开着
ls "目标目录/" | grep '^~\$'
```

有锁文件 → 输出到新文件（如 `xxx_优化版.pptx`），**不要覆盖原文件**。

### 步骤 1 · 全量体检

```bash
python scripts/audit.py "<pptx路径>"
```

输出：页面尺寸与页数、填充色分布、字体、字号分布、文字色、越界元素、溢出风险、页脚样本。
**先读输出再决定改什么** —— 常见硬伤一眼可见：页脚分母跟实际页数不符、封面时长跟目录时长矛盾、某个模块自成一套配色。

### 步骤 2 · 上下文采样（不可跳过）

```bash
python scripts/context.py "<pptx路径>" "8B5CF6,4F46E5,059669,EC4899"
```

列出每个嫌疑色**用在什么元素上**（页号 / 坐标 / 尺寸 / 相邻文本），据此判断它是：

| 判断 | 处理 |
|---|---|
| 模块标识色（目录编号块、章节页进度条、章节徽章） | 保留「每模块一色」机制，只换色相 |
| 卡片装饰色（无语义） | 收敛到规范色板 |
| 语义色（❌ 错误示范 / ✅ 解法 / 红线警示） | **保留语义**，只统一到同一个色值 |

### 步骤 3 · 渲染基线

```bash
python scripts/render.py "<pptx路径>" "<输出目录>" [宽] [高]
```

自动探测 `KWPP.Application`（WPS）→ `PowerPoint.Application`，把每一页导出成 PNG。
**不需要 LibreOffice / Office 全家桶**。

然后用多模态读图能力看几张关键页（封面、目录、各章节页、内容页样本、结束页），建立视觉基线。

### 步骤 4 · 批量精修

写规则 JSON（见下），然后：

```bash
python scripts/refine.py "<pptx路径>" rules.json
```

```json
{
  "fill": [
    {"from": "8B5CF6", "to": "6366F1", "pages": [1, 7, 21, 33, 34, 35]},
    {"from": "EF4444", "to": "DC2626"}
  ],
  "font": [
    {"from": "7C3AED", "to": "4338CA", "pages": [34, 35]},
    {"from": "EC4899", "to": "2563EB"}
  ]
}
```

`pages` 省略 = 全局替换；`pages` 用 0-based 页索引数组。
**按页号限定是刚需** —— 同一个颜色在不同模块里可能是不同用途。

### 步骤 5 · 复检 + 再渲染对比

```bash
python scripts/verify.py "<pptx路径>" palette.json   # 色板合规 / 页脚连续 / 越界 / 溢出
python scripts/overlap.py "<pptx路径>"               # 文本框互相压盖
python scripts/render.py "<pptx路径>" "<新输出目录>"   # 再渲染，读图对比
```

## 常见硬伤清单（体检时逐项扫）

- [ ] 页脚分母 ≠ 实际页数（增删页后必错位，硬编码页码要重排）
- [ ] 封面时长 vs 目录时长 vs 各模块时长之和，三者是否自洽
- [ ] 结束页 / 过渡页是否卡在中间（PPT 分几批追加时极常见）
- [ ] 各模块是否各自一套配色（分批制作的典型痕迹）
- [ ] 页脚页码是否在移动页之后连续

## 已知坑

| 坑 | 解法 |
|---|---|
| `from pptx.util import RGBColor` 报 ImportError | 正确路径是 `from pptx.dml.color import RGBColor` |
| 改 `run.text` 后格式丢失 | 同 run 内替换会保留格式；跨 run 的文本把**第一个 run 设为全文、其余清空** |
| 改长文本后文字横向撑出框 | `word_wrap=False` 的框不换行，改完要加宽 |
| 重新渲染时删除旧图被拦截 | 部分 Agent 环境带文件安全防护会拦 `os.remove`/`rm`，**换输出目录**而不是删旧的 |
| WPS `SaveAs` 导 PNG 报「发生意外」 | 用 `pres.Export(dir, "PNG", w, h)`，不要用 SaveAs |

## 去 AI 味色系整改（全局收敛 47+→8 色）

当用户说"PPT AI 味太浓""颜色太多""风格不统一"时，走这套流程。

### 核心思路

AI 生成 PPT 的典型痕迹：① 紫色系（6D28D9/8B5CF6）泛滥；② 每个模块一种色相形成彩虹效果；③ 黄/绿/红底色卡片到处都是；④ 占位框用醒目橙色虚线。统一方案：杀紫色 → teal 品牌 + 中性灰 + amber 警示，但**保留设计意图色**（封面渐变、目录导航色谱、深色高亮框）。

### 执行步骤（4 阶段）

**Phase 1 · 全局色映射**：定义 COLOR_MAP（47→8），遍历所有 shape 的 `a:srgbClr` 元素替换 `val` 属性。详见 `references/redesign-methodology.md` 的 COLOR_MAP 完整定义。

**Phase 2 · 还原设计意图色**：Phase 1 会把不该统一的也统一了。还原三类：
1. 封面渐变 → 从备份文件读取原始颜色按 shape 顺序写回
2. 目录导航徽章 → 按文字内容匹配，还原为各自的 section 色
3. 深色高亮框 → `is_dark_block_shape()`（无文字+大块）判断，#0F172A 还原为 #1E3A8A

**Phase 3 · 占位框统一**：① 含"截图位/待补"文字 → slate-400；② 虚线边框 `prstDash` + amber 色 → slate-400；③ "截图待补"徽章文字 → 白色

**Phase 3.7 · 占位药丸精准修复**：先找所有有 `prstDash` 虚线 + 宽>4" 的占位框，在每个框右上角 (box_w-1, box_y+0.4) 范围内找小矩形（w<1, h<0.4），改 fill 为 slate-400。**位置匹配法比尺寸+颜色匹配法更精准**，不会误改编号/分类药丸。

### 踩坑教训

| 坑 | 解法 |
|---|---|
| 全局色映射把封面渐变也改成纯黑 | Phase 2 从备份还原，封面是品牌门面不参与统一 |
| 深色高亮框变黑（#1E3A8A 被映射到 #0F172A） | Phase 2 用 `is_dark_block_shape()` 判断大块深色 shape 还原 |
| 按尺寸+颜色匹配改药丸 → 误改编号/分类药丸 | 改用**位置匹配**：先找占位框，再在框右上角找小矩形 |
| `sh.text` 在 Picture 对象上抛 AttributeError | 先 `if not sh.has_text_frame: continue` 过滤 |
| `has_text_frame=False` 时 `sh.text` 报错 | 拆两步：先判 has_text_frame，再判 text.strip() |

## 参考文件

- 配色收敛方法论（语义色判定、同族两档的坑、多模块色相方案）：`references/color-strategy.md`
- 去 AI 味色系整改方法论（COLOR_MAP 完整定义、4 阶段流程、设计意图色判定）：`references/redesign-methodology.md`

## 目录

```
SKILL.md                       本文件（流程 + 原则 + 坑）
scripts/audit.py               全量体检：几何/配色/字体/页脚/越界/溢出
scripts/context.py             嫌疑色的元素上下文采样（先问用途再处理）
scripts/render.py              渲染成 PNG（自动探测 WPS → PowerPoint）
scripts/refine.py              按 rules.json 批量替换填充色/字体色/文本
scripts/verify.py              修后复检：色板合规/页脚连续/越界/溢出
scripts/overlap.py             文本框互相压盖检测
references/color-strategy.md   配色收敛方法论
references/redesign-methodology.md  去 AI 味 4 阶段方法论 + COLOR_MAP
```
