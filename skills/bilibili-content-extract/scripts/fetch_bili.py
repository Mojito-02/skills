# -*- coding: utf-8 -*-
"""B 站视频元信息 / 字幕探测 / 媒体下载。

用法:
  python fetch_bili.py info <BV号或URL> <输出目录>
  python fetch_bili.py dl   <BV号或URL> <输出目录> [qn]

产出:
  info  -> <输出目录>/bili_info.json
  dl    -> <输出目录>/video.mp4 与 <输出目录>/audio.m4s
"""
import json
import os
import re
import sys
import urllib.request

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def parse_bvid(s):
    m = re.search(r"(BV[0-9A-Za-z]{10})", s)
    if not m:
        raise SystemExit("未找到 BV 号: " + s)
    return m.group(1)


def headers(bvid):
    return {
        "User-Agent": UA,
        "Referer": f"https://www.bilibili.com/video/{bvid}/",
        "Origin": "https://www.bilibili.com",
        "Accept": "*/*",
    }


def jget(url, bvid):
    req = urllib.request.Request(url, headers=headers(bvid))
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def download(url, dst, bvid, reporthook=None):
    req = urllib.request.Request(url, headers=headers(bvid))
    got = 0
    with urllib.request.urlopen(req, timeout=60) as r, open(dst, "wb") as f:
        total = int(r.headers.get("Content-Length") or 0)
        while True:
            chunk = r.read(1024 * 256)
            if not chunk:
                break
            f.write(chunk)
            got += len(chunk)
            if reporthook:
                reporthook(got, total)
    return got


def cmd_info(bvid, outdir):
    os.makedirs(outdir, exist_ok=True)
    v = jget(f"https://api.bilibili.com/x/web-interface/view?bvid={bvid}", bvid)
    if v.get("code") != 0:
        raise SystemExit(f"取元信息失败: {v.get('code')} {v.get('message')}")
    d = v["data"]
    aid, cid = d["aid"], d["cid"]

    info = {
        "bvid": d["bvid"], "aid": aid, "cid": cid,
        "title": d["title"], "desc": d.get("desc", ""),
        "duration_sec": d["duration"],
        "pubdate": d.get("pubdate"),
        "owner": d.get("owner", {}).get("name"),
        "owner_mid": d.get("owner", {}).get("mid"),
        "dimension": d.get("dimension"),
        "stat": d.get("stat"),
        "pages": [{"cid": p["cid"], "part": p["part"], "duration": p["duration"]}
                  for p in d.get("pages", [])],
        "subtitle_tracks": [],
    }

    # 逐分P探测字幕轨
    for p in info["pages"]:
        try:
            s = jget(f"https://api.bilibili.com/x/player/wbi/v2?aid={aid}&cid={p['cid']}", bvid)
            subs = (s.get("data", {}).get("subtitle", {}) or {}).get("subtitles", []) or []
            for sub in subs:
                info["subtitle_tracks"].append({
                    "cid": p["cid"], "lan": sub.get("lan"), "lan_doc": sub.get("lan_doc"),
                    "subtitle_url": sub.get("subtitle_url"),
                })
        except Exception as e:
            info.setdefault("subtitle_errors", []).append({"cid": p["cid"], "err": repr(e)})

    info["has_subtitle"] = bool(info["subtitle_tracks"])
    p = os.path.join(outdir, "bili_info.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(info, f, ensure_ascii=False, indent=2)

    m, s = divmod(info["duration_sec"], 60)
    print(f"标题: {info['title']}")
    print(f"UP主: {info['owner']}  时长: {m}分{s}秒  分辨率: {info['dimension']}")
    print(f"字幕轨: {'有 ' + str(len(info['subtitle_tracks'])) + ' 条' if info['has_subtitle'] else '无（需走音频转写）'}")
    print("written:", p)


def cmd_dl(bvid, outdir, qn="32"):
    os.makedirs(outdir, exist_ok=True)
    # 重新拉取元信息，避免依赖被截断的旧 URL
    v = jget(f"https://api.bilibili.com/x/web-interface/view?bvid={bvid}", bvid)
    cid = v["data"]["cid"]

    # 1) mp4（含音轨，体积小）
    pj = jget(f"https://api.bilibili.com/x/player/playurl?bvid={bvid}&cid={cid}"
              f"&qn={qn}&fnval=0&fnver=0&fourk=1", bvid)
    durl = pj.get("data", {}).get("durl") or []
    if durl:
        dst = os.path.join(outdir, "video.mp4")
        n = download(durl[0]["url"], dst, bvid,
                     lambda g, t: print(f"\r  video.mp4 {g/1e6:.1f}MB"
                                        + (f" / {t/1e6:.1f}MB" if t else ""), end="", flush=True))
        print(f"\nMP4 OK {n/1e6:.1f}MB")

    # 2) DASH 最高码率纯音轨
    pj2 = jget(f"https://api.bilibili.com/x/player/playurl?bvid={bvid}&cid={cid}"
               f"&qn=0&fnval=16&fnver=0&fourk=1", bvid)
    audios = (pj2.get("data", {}).get("dash") or {}).get("audio") or []
    if audios:
        a = max(audios, key=lambda x: x["bandwidth"])
        dst = os.path.join(outdir, "audio.m4s")
        try:
            n = download(a["baseUrl"], dst, bvid)
            print(f"AUDIO OK {n/1e6:.1f}MB  id={a['id']} bw={a['bandwidth']}")
        except Exception as e:
            # 纯音轨下载失败（超时/断连）不致命：video.mp4 已含音轨，可直接用于转写与抽帧
            print(f"AUDIO FAIL (可忽略，转写用 video.mp4 即可): {repr(e)[:120]}")
            try:
                if os.path.exists(dst) and os.path.getsize(dst) == 0:
                    os.remove(dst)
            except Exception:
                pass
    print("DONE")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    mode, target, out = sys.argv[1], parse_bvid(sys.argv[2]), sys.argv[3]
    if mode == "info":
        cmd_info(target, out)
    elif mode == "dl":
        cmd_dl(target, out, sys.argv[4] if len(sys.argv) > 4 else "32")
    else:
        print(__doc__)
        sys.exit(1)
