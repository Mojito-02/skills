---
name: bilibili-content-extract
description: 提取 B 站（bilibili.com）视频内容。适用于用户给出 BV 号或视频链接并要求"提取内容/总结/转写/拿文案/看讲了什么"。流程：接口取元信息 → 检测字幕轨 → 无字幕则下载音视频流 → faster-whisper GPU 转写 → 抽帧 + 图文对照报告。也适用于需要 B 站视频逐字稿、关键截图、结构化内容报告的场景。
agent_created: true
---

# B 站视频内容提取

## 适用场景

- 用户给出 `BVxxxxxxxxxx` 或 `https://www.bilibili.com/video/BV...` 链接，要求提取/总结/转写内容。
- 需要 B 站视频的**逐字稿**、**关键截图**、**结构化内容报告**。

## 关键结论（先看这条）

B 站 UGC 视频**绝大多数没有可用的 CC 字幕**（`subtitle.list` 为空；有字幕轨也需要登录 Cookie）。
因此默认路线是：**下载音轨 → 本地 faster-whisper 转写**。不要一上来就承诺"直接取字幕"。

## 环境事实（本机已验证）

| 项 | 值 |
|---|---|
| 托管 Python（3.13，无 CUDA 库时用它） | `C:\Users\ASUS\.workbuddy\binaries\python\versions\3.13.12\python.exe` |
| 托管 env（**含 faster-whisper + nvidia cublas/cudnn**） | `C:\Users\ASUS\.workbuddy\binaries\python\envs\default\Scripts\python.exe` |
| 备用 env（仅 faster-whisper，无 CUDA） | `C:\Users\ASUS\.workbuddy\binaries\python\envs\whisper\Scripts\python.exe` |
| 本机 Python3.9（**有 cv2 4.4**，用于抽帧） | `C:\Users\ASUS\AppData\Local\Programs\Python\Python39\python.exe` |
| 已缓存的 whisper 模型 | `C:\Users\ASUS\.cache\huggingface\hub\` 与 `<工作区>\models\turbo\` |
| GPU | RTX 2060 6GB，`float16` 可用，实测 RTF ≈ **0.055**（35 分钟音频约 2 分钟） |

## 执行流程

### Step 1 · 取元信息 + 探测字幕

```bash
python scripts/fetch_bili.py info BV1xxxxxxxxx <工作目录>
```

输出 `bili_info.json`：标题、UP 主、aid/cid、时长、分辨率、播放/点赞/投币/收藏、**字幕轨列表**。
若 `subtitle_list` 非空 → 直接下字幕（见下），否则进入音频转写。

**接口备忘**
- 元信息：`https://api.bilibili.com/x/web-interface/view?bvid={bvid}`
- 字幕轨：`https://api.bilibili.com/x/player/wbi/v2?aid={aid}&cid={cid}`
  → `data.subtitle.subtitles[].subtitle_url`（**通常为空**）
- 播放流：`https://api.bilibili.com/x/player/playurl?bvid={bvid}&cid={cid}&qn={qn}&fnval={fnval}`
  - `fnval=0&qn=32` → `data.durl[0].url`：**m4s/mp4，含音轨**，360P，体积小，首选
  - `fnval=16` → `data.dash.audio[]`：纯音轨（按 `bandwidth` 挑最大），无音轨损失时更省流量

> **必须带请求头** `Referer: https://www.bilibili.com/video/{bvid}/`，否则 CDN 返回 403。
> `qn=32` 的 mp4 同时含画面与声音，**抽帧和转写可以只下这一个文件**。
> 注意：脚本里不要把 URL 截断保存后再复用，查询串被截断会导致下载失败——**要么存完整 URL，要么在下载脚本内重新拉取**。

### Step 2 · 下载媒体

```bash
python scripts/fetch_bili.py dl BV1xxxxxxxxx <工作目录>
```

产出 `video.mp4`（360P 含音轨）与 `audio.m4s`（最高码率音轨）。

### Step 3 · 转写

```bash
# GPU（推荐）
"<工作目录 或 envs/default>" ... envs/default/Scripts/python.exe scripts/transcribe.py <音频或视频路径> <工作目录> [模型目录或名称]
```

- 模型用 `large-v3-turbo`：中文效果好、速度快。
- 模型名传 `large-v3-turbo` 时 faster-whisper 会从 HF 下载；**若卡住见下方"模型下载兜底"**。
- 脚本已内置：领域提示词 `initial_prompt`、`vad_filter`、`condition_on_previous_text=False`（防幻觉复读循环）。
- 产出 `transcript.json`（起止秒 + 时间戳 + 文本）与 `transcript.txt`。

**必做：为领域定制 `initial_prompt`**。把视频主题、专有名词、常见英文术语写进去，能显著降低技术名词错字率。例如数模视频写「中国研究生数学建模竞赛、MFCC、eGeMAPS、BiLSTM、Transformer、消融实验」。

### Step 4 · 抽帧（可选，但对"讲解类"视频价值很高）

```bash
"C:/Users/ASUS/AppData/Local/Programs/Python/Python39/python.exe" scripts/extract_frames.py <视频> <输出目录> [间隔秒=60]
```

- 每 N 秒抽一帧存 JPEG，并拼 4×3 总览图。
- **可读性经验**：单帧（640×360）看论文/文档正文**清晰可读**；拼成 4×3 大图后**正文不可读、只能看标题**。所以总览图只用于**定位**，要读内容就回头单独看某一帧。

### Step 5 · 出报告

按需产出（不要无脑全做，先看用户要什么）：
- 只要内容 → `逐字稿_带时间戳.txt`
- 要结构化 → Markdown 报告（内容地图 / 分节要点 / 关键数据 / 口径提示）
- 要图文对照 → 单文件 HTML，关键截图 **base64 内嵌**，保证脱离目录也能打开

## 模型下载兜底（HF 卡死的处理）

HF 直连下大模型经常**中途停滞且不报错**。处理办法：

1. 检查 `~/.cache/huggingface/hub/models--*/blobs/*.incomplete` 是否在增长（隔 20s 比两次 `stat -c %s`）。`delta=0` 即卡死。
2. 杀掉卡死进程（后台任务用 TaskStop），**保留 `.incomplete` 文件**。
3. 把它复制到工作区当 `model.bin`，用 `curl -L -C -` 从镜像续传：
   ```bash
   R="https://hf-mirror.com/<repo>/resolve/main"
   curl -sL -o config.json "$R/config.json"      # 同样取 tokenizer.json / vocabulary.json / preprocessor_config.json
   curl -L -C - -o model.bin "$R/model.bin"
   ```
4. `WhisperModel("<本地模型目录>", ...)` —— faster-whisper 支持直接传目录。
5. 校验：本地文件大小 == `Content-Length`（`curl -sIL` 取）。

## CUDA 环境要点

```python
# nvidia 是 namespace package，__file__ 为 None，不能用 os.path.dirname(nvidia.__file__)
import os, glob, sys
base = os.path.join(sys.prefix, "Lib", "site-packages", "nvidia")
for p in glob.glob(os.path.join(base, "*", "bin")):
    os.add_dll_directory(p)                       # 必须，否则 ct2 找不到 cudnn_ops_infer64_9.dll
    os.environ["PATH"] = p + os.pathsep + os.environ.get("PATH", "")
import ctranslate2
print(ctranslate2.get_cuda_device_count())        # 期望 1
print(ctranslate2.get_supported_compute_types("cuda"))  # 期望含 float16
```

依赖：`pip install faster-whisper nvidia-cublas-cu12 nvidia-cudnn-cu12 -i https://pypi.org/simple`
（本机 pip 默认源是清华镜像且缺包，**必须显式指定 pypi.org**；cuDNN 包约 740MB，网络不稳时下载会中断。）

## 坑位清单

- **pip 装包中途失败会原子回滚**：整条 `pip install a b c` 里任一包下载失败，之前下好的也不会装上。宁可分开装、先装小的。
- **`wmic` 被安全策略拉黑**；Bash 里不能调 PowerShell（需用 PowerShell 工具）。
- **Bash 的 `cp` / `rm` / `mkdir -p` 状态不稳定**，重要文件操作用 Python `shutil` 或 PowerShell。
- 转写进度打印要每 N 段 flush 一次，否则长任务看不到任何输出，会误判成卡死。
- 视频里的**口述数字可能与字幕/论文原文矛盾**（尤其技术维度、指标）。报告里要**标注差异并说明以哪个为准**，不要悄悄取一个。

## 产出目录约定

```
<工作目录>/
  video.mp4 / audio.m4s        源文件
  bili_info.json               元信息 + 字幕探测结果
  transcript.json / .txt       逐字稿
  frames/                      抽帧原图 + 总览拼版
  关键截图/                     按内容命名的精选图
  <主题>_视频内容提取.md         结构化报告
  视频内容提取报告.html          图文对照版
```
