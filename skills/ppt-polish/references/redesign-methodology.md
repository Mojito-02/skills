# 去 AI 味色系整改方法论

> 从 47 色统一到 8 色的完整流程，含 COLOR_MAP 定义、设计意图色判定、占位框精准修复策略。

## 最终色板（8 色）

| 名称 | Hex | 用途 |
|---|---|---|
| slate-900 | #0F172A | 标题、深色背景 |
| slate-700 | #334155 | 正文、次级标题 |
| slate-400 | #94A3B8 | 辅助文字、占位框、次要信息 |
| white | #FFFFFF | 反色文字、白底 |
| slate-50 | #F8FAFC | 浅底背景 |
| slate-200 | #E2E8F0 | 分隔线、边框 |
| teal-700 | #0F766E | 品牌强调色（统一替代所有蓝/紫/绿装饰色） |
| amber-700 | #B45309 | 警示/反面强调（替代红色、弱化橙黄） |

## 保留的设计意图色（不参与统一）

| 色彩 | 出现位置 | 保留理由 |
|---|---|---|
| 封面渐变原始色 | P1 封面 | 品牌门面，渐变是设计意图 |
| 目录 7 色导航色谱 | P2 目录徽章 | 导航功能需要多色区分 |
| #1E3A8A 深蓝 | P22/P28/P36/P37 深色高亮框 | 大块深色背景的强调效果 |

## COLOR_MAP 完整定义（47→8）

```python
COLOR_MAP = {
    # ── 紫色系（全杀 → teal，AI 味最重）
    '6D28D9': '0F766E',
    '8B5CF6': '0F766E',
    '7C3AED': '0F766E',
    'A78BFA': '0F766E',
    'C4B5FD': '0F766E',
    'EDE9FE': '0F766E',
    '5B21B6': '0F766E',
    '4C1D95': '0F766E',
    'DDD6FE': '0F766E',

    # ── 蓝色系 → teal（统一品牌色）
    '2563EB': '0F766E',
    '1D4ED8': '0F766E',
    '3B82F6': '0F766E',
    '60A5FA': '0F766E',
    '93C5FD': '0F766E',
    'DBEAFE': '0F766E',
    '1E40AF': '0F766E',
    '1E3A8A': '0F766E',  # 注意：深色高亮框在 Phase 2 还原

    # ── 绿色系 → teal
    '059669': '0F766E',
    '10B981': '0F766E',
    '34D399': '0F766E',
    'A7F3D0': '0F766E',
    'D1FAE5': '0F766E',
    '064E3B': '0F766E',
    '047857': '0F766E',

    # ── 红色系 → amber（减弱警示感）
    'DC2626': 'B45309',
    'EF4444': 'B45309',
    'F87171': 'B45309',
    'FCA5A5': 'B45309',
    'FEE2E2': 'B45309',
    '991B1B': 'B45309',
    'B91C1C': 'B45309',

    # ── 橙黄系 → amber-700
    'F59E0B': 'B45309',
    'D97706': 'B45309',
    'FBBF24': 'B45309',
    'FCD34D': 'B45309',
    'FEF3C7': 'B45309',
    'FDE68A': 'B45309',
    '92400E': 'B45309',
    '78350F': 'B45309',

    # ── 粉色系 → teal
    'EC4899': '0F766E',
    'F472B6': '0F766E',
    'FBCFE8': '0F766E',
    'BE185D': '0F766E',
    '831843': '0F766E',

    # ── 青色系 → teal
    '06B6D4': '0F766E',
    '0891B2': '0F766E',
    '22D3EE': '0F766E',

    # ── 卡片底色 → 中性底
    'FFFBEB': 'F8FAFC',  # 黄底
    'F0FDFA': 'F8FAFC',  # 绿底
    'FEF2F2': 'F8FAFC',  # 红底
    'FDF4FF': 'F8FAFC',  # 粉底
    'EFF6FF': 'F8FAFC',  # 蓝底
    'ECFDF5': 'F8FAFC',  # 绿底变体
    'FFF7ED': 'F8FAFC',  # 橙底

    # ── 深色背景 → slate-900（Phase 2 会还原设计意图色）
    '1E1B4B': '0F172A',  # indigo 深
    '111827': '0F172A',
    '0F172A': '0F172A',  # 自身
}
```

## 设计意图色判定方法

### `is_dark_block_shape()` — 深色高亮框

```python
def is_dark_block_shape(shape):
    """判断是否为深色高亮块（大尺寸 + 无文字/少文字）"""
    w = shape.width / 914400  # EMU → inches
    h = shape.height / 914400
    has_text = shape.has_text_frame and shape.text_frame.text.strip()
    return w > 3 and h > 1.5 and not has_text
```

匹配后：把 #0F172A 改回 #1E3A8A（深蓝，保留视觉冲击力）。

### 目录徽章颜色还原

按文字内容匹配：
```python
TOC_BADGE_COLORS = {
    '第1讲': '2563EB',  # 蓝
    '第2讲': '0F766E',  # teal
    '第3讲': 'B45309',  # amber
    '第4讲': '2563EB',  # 蓝
    '第5讲': '0F766E',  # teal
    '附录': '475569',   # 灰
    '收尾': 'B91C1C',   # 红
}
```

### 封面渐变还原

从备份文件读取原始 XML 中 `a:gradFill` 的 `a:gs`（渐变停止点），按 shape 在幻灯片中的顺序写回当前文件。不能硬编码颜色值——每份 PPT 的封面渐变可能不同。

## 占位药丸精准修复（Phase 3.7）

### 问题

Phase 3.5 用「尺寸 + 颜色」匹配（w<1, h<0.3, fill=B45309）识别占位药丸，但误改了：
- P27 的步骤编号药丸（1,2,3,4,5）
- P31 的 A/B/C 分类药丸
- P34 的运营/销售分类药丸

### 正确方法：位置匹配

```python
# Step 1: 找所有占位框（有虚线边框 + 宽>4"）
placeholder_boxes = []
for shape in slide.shapes:
    if has_dashed_border(shape) and shape.width / 914400 > 4:
        placeholder_boxes.append(shape)

# Step 2: 在每个占位框的右上角范围找小矩形
for box in placeholder_boxes:
    box_x = box.left / 914400
    box_y = box.top / 914400
    box_w = box.width / 914400
    box_h = box.height / 914400

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if shape.text.strip():
            continue  # 有文字的不是纯背景药丸

        sx = shape.left / 914400
        sy = shape.top / 914400
        sw = shape.width / 914400
        sh = shape.height / 914400

        # 判断是否在占位框右上角范围
        in_right_top = (
            sx > box_x + box_w - 1 and sx < box_x + box_w and
            sy > box_y and sy < box_y + 0.4 and
            sw < 1 and sh < 0.4
        )
        if in_right_top:
            # 这是占位药丸，改灰
            set_fill(shape, '94A3B8')
```

### 效果

- Phase 3.5（尺寸+颜色匹配）：22 pills 改灰，**全是误改**（编号/分类药丸）
- Phase 3.6：22 pills 全回滚 teal
- Phase 3.7（位置匹配）：9 pills 精准改灰，**全部正确**（真正的占位药丸）

## 通用脚本模板

```python
from pptx import Presentation
from lxml import etree

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

def remap_colors(prs, color_map):
    """遍历所有 shape 的 srgbClr 元素，按 color_map 替换 val 属性"""
    count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            xml = shape._element.xml
            for srgb in etree.XPath('//a:srgbClr')(etree.fromstring(xml)):
                old_val = srgb.get('val', '').upper()
                if old_val in color_map:
                    srgb.set('val', color_map[old_val])
                    count += 1
            # 写回修改后的 XML
            shape._element = etree.fromstring(xml_modified)
    return count
```

> 注意：python-pptx 的 shape._element 是 lxml Element，可以直接操作子节点，但需要用 `shape._element` 的 `xpath` 或 `iter` 方法遍历 `a:srgbClr`。
