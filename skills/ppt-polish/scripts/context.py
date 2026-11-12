#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
嫌疑色上下文采样 —— 看一个颜色到底用在什么元素上，再决定是收敛还是保留。

用法:
    python context.py <pptx路径> ["8B5CF6,4F46E5,059669"]
    python context.py <pptx路径>            # 不传颜色时，自动列出出现最多的 8 个彩色

判断指引:
    - 42x42 / 56x5 / 6x34 这类小色块 + 分布在目录页和每个章节页 -> 模块标识色
    - 大卡片底 (1136xNN) + 集中在某几个连续页                  -> 模块主题色
    - 伴随 "反面示范 / 错误 / 红线" 文案                        -> 语义色，保留
    - 伴随 "✅ / 解法 / 正确"                                  -> 语义色，保留
"""
import sys, os, io
from collections import defaultdict, Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
EMU_PX = 9525.0

# 中性灰阶不算"嫌疑色"
NEUTRAL = {"FFFFFF", "000000", "0F172A", "1E293B", "334155", "475569",
           "64748B", "94A3B8", "CBD5E1", "E2E8F0", "F1F5F9", "F8FAFC"}


def fill_hex(sh):
    try:
        f = sh.fill
        if f.type is not None and int(f.type) == 1:
            return str(f.fore_color.rgb).upper()
    except Exception:
        pass
    return None


def run(path, targets=None):
    from pptx import Presentation

    prs = Presentation(path)
    ctx = defaultdict(list)
    counter = Counter()

    def walk(shapes, sidx):
        for sh in shapes:
            fh = fill_hex(sh)
            if fh:
                counter[fh] += 1
                if targets and fh in targets:
                    txt = ""
                    if sh.has_text_frame:
                        txt = "".join(p.text for p in sh.text_frame.paragraphs)
                    try:
                        pos = (round(sh.left / EMU_PX), round(sh.top / EMU_PX))
                        size = (round(sh.width / EMU_PX), round(sh.height / EMU_PX))
                    except Exception:
                        pos, size = (0, 0), (0, 0)
                    ctx[fh].append({"p": sidx, "pos": pos, "size": size, "t": txt[:36]})
            if hasattr(sh, "shapes"):
                try:
                    walk(sh.shapes, sidx)
                except Exception:
                    pass

    for i, s in enumerate(prs.slides):
        walk(s.shapes, i)

    if not targets:
        targets = [c for c, _ in counter.most_common() if c not in NEUTRAL][:8]
        print(f"（未指定颜色，自动取出现最多的 {len(targets)} 个彩色）\n")

    for color in targets:
        items = ctx.get(color, [])
        if not items:
            print(f"\n#{color}: 未出现")
            continue
        pages = sorted({it["p"] for it in items})
        print(f"\n{'=' * 66}")
        print(f"#{color}   共 {len(items)} 处，分布在 {len(pages)} 页")
        print(f"  页: {pages[:30]}{' ...' if len(pages) > 30 else ''}")
        seen, shown = set(), 0
        for it in items:
            key = (it["size"], it["t"][:8])
            if key in seen:
                continue
            seen.add(key)
            print(f"   P{it['p']:02d} pos={it['pos']} size={it['size']}  '{it['t']}'")
            shown += 1
            if shown >= 6:
                break


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    tg = None
    if len(sys.argv) > 2 and sys.argv[2].strip():
        tg = [c.strip().upper().lstrip("#") for c in sys.argv[2].split(",")]
    run(sys.argv[1], tg)
