# -*- coding: utf-8 -*-
"""faster-whisper 转写，GPU 优先、自动降级 CPU。

用法:
  python transcribe.py <媒体文件> <输出目录> [模型名或本地模型目录] [initial_prompt]

产出:
  <输出目录>/transcript.json   {model, device, compute_type, language, duration, segments[]}
  <输出目录>/transcript.txt    [分:秒] 文本
"""
import glob
import json
import os
import sys
import time

MEDIA = sys.argv[1]
OUT = sys.argv[2]
MODEL = sys.argv[3] if len(sys.argv) > 3 else "large-v3-turbo"
PROMPT = sys.argv[4] if len(sys.argv) > 4 else None

# --- 注入 pip 安装的 CUDA 运行库 ---
# 注意: nvidia 是 namespace package, __file__ 为 None, 不能用 dirname(nvidia.__file__)
base = os.path.join(sys.prefix, "Lib", "site-packages", "nvidia")
for p in glob.glob(os.path.join(base, "*", "bin")):
    try:
        os.add_dll_directory(p)
        os.environ["PATH"] = p + os.pathsep + os.environ.get("PATH", "")
    except Exception:
        pass

from faster_whisper import WhisperModel  # noqa: E402

# 按顺序尝试：GPU float16 -> CPU int8 -> CPU int8 + 小模型
CANDS = [
    (MODEL, "cuda", "float16"),
    (MODEL, "cpu", "int8"),
    ("medium", "cpu", "int8"),
]

model = used = None
for name, dev, ct in CANDS:
    try:
        print(f"[try] {name} / {dev} / {ct}", flush=True)
        t0 = time.time()
        m = WhisperModel(name, device=dev, compute_type=ct,
                         cpu_threads=os.cpu_count())
        print(f"[ok] loaded in {time.time()-t0:.1f}s", flush=True)
        model, used = m, (name, dev, ct)
        break
    except Exception as e:
        print(f"[fail] {repr(e)[:200]}", flush=True)

if model is None:
    raise SystemExit("ALL_CANDIDATES_FAILED")

t0 = time.time()
segments, info = model.transcribe(
    MEDIA,
    language="zh",
    beam_size=5,
    vad_filter=True,
    vad_parameters=dict(min_silence_duration_ms=400),
    initial_prompt=PROMPT,
    condition_on_previous_text=False,   # 防幻觉复读循环
    temperature=[0.0, 0.2, 0.4],
)

lines, plain = [], []
for s in segments:
    mm, ss = int(s.start // 60), int(s.start % 60)
    stamp = f"[{mm:02d}:{ss:02d}]"
    txt = s.text.strip()
    lines.append({"start": round(s.start, 2), "end": round(s.end, 2),
                  "stamp": stamp, "text": txt})
    plain.append(f"{stamp} {txt}")
    if len(lines) % 50 == 0:
        el = time.time() - t0
        rtf = el / max(s.end, 1)
        print(f"  seg={len(lines)} at={s.end:.0f}s elapsed={el:.0f}s "
              f"rtf={rtf:.3f} eta={(info.duration - s.end) * rtf:.0f}s", flush=True)

os.makedirs(OUT, exist_ok=True)
with open(os.path.join(OUT, "transcript.json"), "w", encoding="utf-8") as f:
    json.dump({"model": str(used[0]), "device": used[1], "compute_type": used[2],
               "language": info.language, "duration": info.duration,
               "segments": lines}, f, ensure_ascii=False, indent=1)
with open(os.path.join(OUT, "transcript.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(plain))

print(f"DONE segs={len(lines)} chars={sum(len(l['text']) for l in lines)} "
      f"elapsed={time.time()-t0:.0f}s model={used}")
