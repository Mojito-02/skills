---
name: dwg-to-pdf
description: 在没有安装任何 CAD 软件（AutoCAD/中望/浩辰等）的 Windows 机器上，把 DWG/DXF 图纸转成高质量矢量 PDF。当用户要求「DWG 转 PDF / 图纸导出 PDF / CAD 文件转 PDF / 出图」时使用。内置便携版 ODA File Converter，无需系统安装。
agent_created: true
---

# DWG / DXF → PDF（零安装）

本机**没有装任何 CAD 软件**，也不走在线转换（图纸外传有风险）。
链路固定为：

```
DWG --(ODA File Converter 便携版)--> DXF --(ezdxf + matplotlib)--> 矢量 PDF
```

## 资产位置（已就位，不要重复下载）

| 用途 | 路径 |
|---|---|
| DWG→DXF 转换器 | `C:\Users\ASUS\.workbuddy\tools\ODAFileConverter\ODAFileConverter.exe` |
| Python 环境（已装 ezdxf / matplotlib / pymupdf） | `C:\Users\ASUS\.workbuddy\binaries\python\envs\default\Scripts\python.exe` |
| 本技能脚本 | `C:\Users\ASUS\.workbuddy\skills\dwg-to-pdf\scripts\dwg2pdf.py` |

## 用法

```bash
"C:/Users/ASUS/.workbuddy/binaries/python/envs/default/Scripts/python.exe" \
  "C:/Users/ASUS/.workbuddy/skills/dwg-to-pdf/scripts/dwg2pdf.py" \
  "路径/图纸.dwg" "输出.pdf" [--color] [--keep-dxf]
```

- 默认输出**黑白打印版**（白底黑线，A1 比例）；`--color` 出彩色版（还原 CAD 屏幕配色）。
- 输出页高固定 594 mm（A1 系列），宽按图纸包围盒比例自适应。
- 图纸里的中文路径/文件名不影响，脚本内部会先落到 ASCII 临时目录再处理。

## 关键坑（踩过的，别再踩）

1. **中文变乱码 = MTEXT 内联字体**
   图纸中文常写成 `{\fFangSong_GB2312|b0|i0|c134|p49;设 计}`。
   只改文字样式（`doc.styles`）**无效**，必须用 `\f[^;]*;` 正则把内联字体码**删掉**，
   再把所有 style 的 font 指向真实中文字体 `simhei.ttf`。脚本已内置。

2. **SHX 字体本机没有**：`isocp.shx / hztxt.shx / gbcbig.shx / hzfs.shx` 都不存在，
   ezdxf 解析不到会静默回退到拉丁字体。所以必须换成系统 TTF（simhei/simsun/msyh/simfang）。

3. **ezdxf 的字体解析函数行为不一致**，用之前先验证：
   - `find_best_match("SimHei")` → `simhei.ttf` ✅（按字体族名查）
   - `find_best_match("simhei.ttf")` → `None` ❌（按文件名查不行）
   - `has_font("simhei.ttf")` → `True` ✅
   结论：**用文件名配 `has_font` 判断可用性，用字体族名做匹配**。

4. **matplotlib 背景**：必须传 `Configuration(background_policy=BackgroundPolicy.WHITE)`。
   否则 ezdxf 会按深色主题渲染，白色线（ACI 7）在深底上才看得见，白底就"消失"。
   黑白版配 `ColorPolicy.MONOCHROME_LIGHT_BG`。

5. **PDF 里的文字是矢量轮廓，不可选中/搜索**。这是 ezdxf matplotlib 后端的固有行为，
   视觉上完全正确、打印无差异；若用户明确要求可复制文本，得另找方案。

6. **aspose-cad 不要用**：免费版会在图上打巨大 "Evaluation only" 横幅 + 对角斜叉水印，
   不能作为交付件。（PyPI 上 Python 3.13 也装不上。）

## 环境注意事项（本机特性）

- PowerShell 工具**禁止 COM 实例化**（`New-Object -ComObject` 直接报安全拦截），
  所以读 MSI 文件表要用 Python 的 `msilib`（Python 3.12 有，3.13 已移除）。
- Bash 里**不能调用 powershell**（安全策略拦截）。
- **pip 默认源是清华镜像，缺包**（ezdxf / aspose-cad 都报 "from versions: none"）。
  装包时必须显式指定：
  `-i https://pypi.org/simple`。
- ODA 官方直链（免注册、可直接 curl）：
  `https://www.opendesign.com/guestfiles/get?filename=ODAFileConverter_QT6_vc16_amd64dll_27.1.msi`
  取回的是 MSI，**不需要安装**：用 Bandizip 解出内嵌 CAB，再用 Python `msilib`
  读 `File` / `Component` / `Directory` 表还原真实文件名即可（这就是 tools 目录的来历）。

## 验收

出图后**必须**做两件事，再交付：
1. 渲染首页 PNG（`pymupdf`）肉眼看一遍，确认图元完整、标题栏中文正常；
2. 报告页面尺寸（mm）和文件大小。
