# -*- coding: utf-8 -*-
"""按固定间隔抽帧 + 生成总览拼版。

需用带 OpenCV 的解释器运行（本机: AppData\\Local\\Programs\\Python\\Python39\\python.exe）

用法:
  python extract_frames.py <视频文件> <输出目录> [间隔秒=60] [每版帧数=12]

说明:
  单帧(如 640x360)看文档正文清晰可读；拼成 4x3 大图后正文不可读、仅能看标题。
  总览图用于“定位”，要读内容请回看单帧文件。
"""
import os
import sys

import cv2
import numpy as np

VID = sys.argv[1]
OUT = sys.argv[2]
STEP = int(sys.argv[3]) if len(sys.argv) > 3 else 60
PER = int(sys.argv[4]) if len(sys.argv) > 4 else 12

FR = os.path.join(OUT, "frames")
os.makedirs(FR, exist_ok=True)

cap = cv2.VideoCapture(VID)
fps = cap.get(cv2.CAP_PROP_FPS) or 25
n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
dur = n / fps if fps else 0
W = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
H = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
print(f"fps={fps:.3f} frames={n} dur={dur:.1f}s size={W}x{H}")

saved = []
for t in range(0, int(dur), STEP):
    cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000.0)
    ok, fr = cap.read()
    if not ok:
        continue
    label = f"{t // 60:02d}{t % 60:02d}"
    cv2.imwrite(os.path.join(FR, label + ".jpg"), fr,
                [int(cv2.IMWRITE_JPEG_QUALITY), 92])
    saved.append((label, fr))
cap.release()
print("frames saved:", len(saved))

# 拼版
COLS = 4
ROWS = max(1, PER // COLS)
TILE_W = W
TILE_H = H
sheets = []
for i in range(0, len(saved), COLS * ROWS):
    grp = saved[i:i + COLS * ROWS]
    sheet = np.full((ROWS * (TILE_H + 26), COLS * TILE_W, 3), 24, dtype=np.uint8)
    for j, (label, fr) in enumerate(grp):
        r, c = divmod(j, COLS)
        y0 = r * (TILE_H + 26)
        sheet[y0:y0 + TILE_H, c * TILE_W:(c + 1) * TILE_W] = fr
        cv2.putText(sheet, label, (c * TILE_W + 8, y0 + TILE_H + 19),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.62, (90, 220, 255), 2, cv2.LINE_AA)
    sp = os.path.join(FR, f"sheet_{i // (COLS * ROWS) + 1:02d}.jpg")
    cv2.imwrite(sp, sheet, [int(cv2.IMWRITE_JPEG_QUALITY), 88])
    sheets.append(sp)

print("sheets:", len(sheets))
for s in sheets:
    print(s)
print("DONE")
