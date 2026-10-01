<div align="center">

# 🛠 Ian Skills

#### Ian 的 Skill 集合，自己在用的都开源在这里

[![AgentSkills](https://img.shields.io/badge/AgentSkills-Standard-8B5CF6?style=flat-square)](https://agentskills.io)
![Codex](https://img.shields.io/badge/Codex-Skill-10B981?style=flat-square&logo=openai&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-3B82F6?style=flat-square)

</div>

这里的每个 Skill 都是 Agent 能直接加载的结构化指令集，遵循 [Agent Skills](https://agentskills.io) 开放标准。

---

## 📦 安装方式

在 Codex 等支持 Agent Skills 的工具里，直接说：

```
帮我安装这个 skill：https://github.com/Ian-nextai/Ian-Skills/tree/main/AMZ-Excess-Inventory-Check
```

你的 Agent 不支持 Skill 也没关系：把对应目录的 `SKILL.md` 全文下载下来，当成项目规则文件（或直接贴进对话）让 Agent 照着执行。

---

## ✨ Skills

<table>
<tr><td>

### 📦 AMZ-Excess-Inventory-Check（FBA 冗余库存检查）

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

> ⚠️ 脚本依赖 Codex 内置的 `@oai/artifact-tool`。非 Codex 环境下不用装它，让 Agent 换个表格库重写导出、或直接只交付分析结论即可；判定规则和四个 sheet 的结构不变，排版细节可省。

→ [SKILL.md](./AMZ-Excess-Inventory-Check/SKILL.md) · [字段定义](./AMZ-Excess-Inventory-Check/references/output-schema.md)

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
