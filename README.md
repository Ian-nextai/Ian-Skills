<div align="center">

# 🛠 Ian Skills

#### Ian 的 Skill 集合，自己在用的都开源在这里

[![AgentSkills](https://img.shields.io/badge/AgentSkills-Standard-8B5CF6?style=for-the-badge)](https://agentskills.io)
![Codex](https://img.shields.io/badge/Codex-Skill-10B981?style=flat-square&logo=openai&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-3B82F6?style=flat-square)

</div>

这里的每个 Skill 都是 Agent 能直接加载的结构化指令集，遵循 [Agent Skills](https://agentskills.io) 开放标准。

---

## ⚠️ 运行环境

**这些 Skill 是在 Codex 里开发并日常使用的，按 Codex 的环境假设编写。**

`Amz-Excess-Inventory-Check` 的分析脚本 `scripts/analyze_inventory.mjs` 依赖 Codex 运行时内置的 `@oai/artifact-tool`（用来生成 Excel 工作簿），以及两个环境变量：

| 变量 | 用途 |
|---|---|
| `$CODEX_PRIMARY_RUNTIME_NODE` | Node 可执行文件路径 |
| `$CODEX_PRIMARY_RUNTIME_NODE_MODULES` | 预置依赖目录，脚本通过软链使用它 |

**在 Codex 之外（Claude Code、Cursor 等其他 Agent）脚本不会直接跑通**——不是 Skill 标准不兼容，只是缺这个包和这些变量。这种情况下**不用去装它**，让执行任务的 Agent 自己换一套方案就行，规则部分完全不受影响：

1. 照 `SKILL.md` 的工作流与判定规则，用当前环境能装上的表格库重写导出（Python 的 `openpyxl`、Node 的 `exceljs` / `xlsx` 都可以）；
2. 或者干脆不写 xlsx——先只产出分析与结论，导出格式由你或 Agent 另行决定。

两条底线不能变：**四个 sheet 的顺序与字段**（照 [references/output-schema.md](./Amz-Excess-Inventory-Check/references/output-schema.md) 对齐），以及**冗余优先级判定规则**。至于表格配色、列宽、冻结窗格这些排版细节，属于 Codex 那套运行时的呈现效果，换环境后做不到也没关系，不影响结论。

预览图（PNG）依赖 `@oai/artifact-tool` 的渲染能力，换环境后同样没有。按 `SKILL.md` 的约定，预览只是导出后的目视检查，不是产出工作簿的前提，跳过即可。

---

## 📦 安装方式

在 Codex 等支持 Agent Skills 的工具里，直接说：

```
帮我安装这个 skill：https://github.com/Ian-nextai/Ian-Skills/tree/main/Amz-Excess-Inventory-Check
```

你的 Agent 不支持 Skill 也没关系：把对应目录的 `SKILL.md` 全文下载下来，当成项目规则文件（或直接贴进对话）让 Agent 照着执行。

---

## ✨ Skills

<table>
<tr><td>

### 📦 Amz-Excess-Inventory-Check（FBA 冗余库存检查）

> *"亚马逊判定冗余、销量、库龄、仓储费，四张表一次看清该清哪些货。"*

分析 Amazon FBA 库存健康度，识别冗余库存与库龄风险，导出固定结构的 Excel 工作簿（分析总览 / 冗余库存 / 库存全量 / 数据检查）。

**三种数据来源**

| Mode | 来源 | 费用列 |
|---|---|---|
| A | Amazon 后台导出的三份 CSV | 完整 |
| B | Amazon MCP 报表网关（同样三份报表） | 完整 |
| C | 本地领星 MCP 网关 | **留空并标注**（领星不提供历史费用明细） |

**冗余优先级怎么定的**

先看 Amazon 的 `estimated-excess-quantity`，再结合销量与库龄分级：

| 级别 | 条件 |
|---|---|
| 🔴 P0 紧急 | 存在 366 天以上库存 |
| 🟠 P1 高 | 60 天零销，或存在 181 天以上库存 |
| 🔵 P2 中 | 30 天零销，或冗余量 ≥ 可售量的 50% |
| 🟢 P3 低 | 其余 Amazon 判定为冗余的库存 |

两条容易踩的口径：

- 超龄附加费是**历史证据**，单凭它不构成 P0。
- 月度仓储费与超龄附加费的**期间往往不同**，不要相加当成"本月费用"。

**输出**

文件按数据快照日期命名：`amazon_fba_inventory_check_<YYYYMMDD>.xlsx`。日期取报表里的 `snapshot-date`，不是运行当天。

→ [SKILL.md](./Amz-Excess-Inventory-Check/SKILL.md) · [字段定义](./Amz-Excess-Inventory-Check/references/output-schema.md)

</td></tr>
</table>

---

## 🌟 关于

我是 Ian，这些 Skill 都是自己业务里跑通之后才搬出来的。

开源出来如果对你有帮助，给个 ⭐ 就行。有问题或建议，欢迎在 Issues / Discussions 里说一声。

---

<div align="center">

[MIT License](./LICENSE) · 自由使用 / 修改 / 再分发

Made by [@Ian-nextai](https://github.com/Ian-nextai)

</div>
