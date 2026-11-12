# Restricted skills

以下技能未放入本公开镜像。原因是其 `LICENSE.txt` 标注为 Anthropic 的
**source-available / proprietary** 材料，不是 Apache-2.0 开源内容。
它们在本地 Magpie 库中仍然存在；如需完整备份，请使用私有仓库或私有存储。

| Skill | Upstream | License |
| --- | --- | --- |
| `portable-document-handler` | [`qodex-ai/ai-agent-skills`](https://github.com/qodex-ai/ai-agent-skills/blob/main/skills/portable-document-handler/SKILL.md) | Source-available / proprietary |
| `presentation-builder` | [`qodex-ai/ai-agent-skills`](https://github.com/qodex-ai/ai-agent-skills/blob/main/skills/presentation-builder/SKILL.md) | Source-available / proprietary |
| `spreadsheet-processor` | [`qodex-ai/ai-agent-skills`](https://github.com/qodex-ai/ai-agent-skills/blob/main/skills/spreadsheet-processor/SKILL.md) | Source-available / proprietary |
| `word-document-processor` | [`qodex-ai/ai-agent-skills`](https://github.com/qodex-ai/ai-agent-skills/blob/main/skills/word-document-processor/SKILL.md) | Source-available / proprietary |

恢复方式（本机）：

```bash
rsync -a ~/.config/magpie/library/skills/portable-document-handler/ <private-destination>/portable-document-handler/
rsync -a ~/.config/magpie/library/skills/presentation-builder/ <private-destination>/presentation-builder/
rsync -a ~/.config/magpie/library/skills/spreadsheet-processor/ <private-destination>/spreadsheet-processor/
rsync -a ~/.config/magpie/library/skills/word-document-processor/ <private-destination>/word-document-processor/
```
