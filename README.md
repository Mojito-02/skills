# Magpie Skills Backup

这是 `Magpie` 技能库在 **2026-11-12** 的快照，来源为本机：

```text
~/.config/magpie/library/skills
```

## 快照概况

| 项目 | 数量 |
| --- | ---: |
| Magpie 注册技能 | 351 |
| 有本地 SKILL.md 且定位到上游仓库 | 296 |
| 其中 SKILL.md 内容完全一致 | 172 |
| 其中同名但内容有差异（后续版本/本地修改） | 124 |
| 本地缺失 SKILL.md 的残缺条目 | 42 |
| 其中可定位上游候选 | 39 |
| 完全未找到上游 | 16 |

> `exact` 表示本地 `SKILL.md` 与上游当前 `main` 分支（少数历史版本为对应 commit）SHA-256 一致。
> `name_only` 表示目录名和描述吻合，但本地版本与上游当前版本不同。

## 目录结构

```text
skills/<name>/           # 347 个技能目录（351 条中排除 4 个受限技能）
skills/<name>/UPSTREAM.md # 该技能的原仓库、原文件、commit 和许可证
sources.json             # 机器可读的来源清单（351 条）
SKILLS.md                # 按技能名索引，直接链接原 GitHub 仓库
SOURCES.md               # 按上游仓库分组的可读清单
INCOMPLETE.md            # 42 个缺失 SKILL.md 的残缺条目
RESTRICTED.md            # 因许可证原因未放入公开镜像的 4 个技能
NOTICE.md                # 第三方许可证与归属说明
licenses/                # 上游仓库许可证原文
```

## 主要上游

| 上游仓库 | 技能数 | 许可证 |
| --- | ---: | --- |
| [`K-Dense-AI/claude-scientific-skills`](https://github.com/K-Dense-AI/claude-scientific-skills) | 66 | MIT |
| [`qodex-ai/ai-agent-skills`](https://github.com/qodex-ai/ai-agent-skills) | 66 | Apache-2.0 |
| [`wanshuiyin/Auto-claude-code-research-in-sleep`](https://github.com/wanshuiyin/Auto-claude-code-research-in-sleep) | 62 | MIT |
| [`mattpocock/skills`](https://github.com/mattpocock/skills) | 46 | MIT |
| [`Yuan1z0825/nature-skills`](https://github.com/Yuan1z0825/nature-skills) | 22 | Apache-2.0 |
| [`obra/superpowers`](https://github.com/obra/superpowers) | 14 | MIT |
| [`DietrichGebert/ponytail`](https://github.com/DietrichGebert/ponytail) | 6 | MIT |
| [`jihe520/MathModelAgent`](https://github.com/jihe520/MathModelAgent) | 4 | Custom: free personal use, no commercial use, no closed-source redistribution |
| [`anthropics/skills`](https://github.com/anthropics/skills) | 2 | Apache-2.0 (per upstream README) |
| [`AIScientists-Dev/academic-humanizer`](https://github.com/AIScientists-Dev/academic-humanizer) | 1 | MIT |
| [`Haojae/scipilot-figure-skill`](https://github.com/Haojae/scipilot-figure-skill) | 1 | MIT |
| [`OthmanAdi/planning-with-files`](https://github.com/OthmanAdi/planning-with-files) | 1 | MIT |
| [`hugohe3/ppt-master`](https://github.com/hugohe3/ppt-master) | 1 | MIT |
| [`lc2panda/wps-skills`](https://github.com/lc2panda/wps-skills) | 1 | MIT |
| [`multica-ai/andrej-karpathy-skills`](https://github.com/multica-ai/andrej-karpathy-skills) | 1 | MIT (declared in SKILL.md) |
| [`vercel-labs/skills`](https://github.com/vercel-labs/skills) | 1 | MIT |
| [`zyhsechub/chinese-thesis-workbench-skill`](https://github.com/zyhsechub/chinese-thesis-workbench-skill) | 1 | MIT |

其余 16 个条目未找到可靠上游，标记为 `local`；详见 `sources.json`。

## 使用与同步

把某个技能复制回 Magpie 库即可：

```bash
rsync -a skills/<name>/ ~/.config/magpie/library/skills/<name>/
```

批量恢复前建议先备份现有目录。部分技能包含 `scripts/`、模板和字体资源，不要只复制 `SKILL.md`。

## 许可证

本仓库是**混合许可证**的个人快照，不是单一开源项目。每个技能保留其原始 `LICENSE`/`LICENSE.txt`；上游仓库的许可证见 `SOURCES.md` 与 `NOTICE.md`。

其中 4 个 Anthropic 文档类技能为 **source-available / proprietary**，未放入本公开镜像；如果你需要完整本地备份，请把仓库改为私有后再加入，或单独保存到私有位置。

## 已知问题

- 42 个 Magpie 条目在本地就没有 `SKILL.md`，属于历史导入残留；清单见 `INCOMPLETE.md`。
- 同名技能可能来自不同上游（例如 `exa-search` 同时出现在 K-Dense 与 ARIS 仓库），本快照按本地来源目录和内容哈希择优归属。
- 本快照只代表 2026-11-12 的本机状态，不代表上游最新版本。
