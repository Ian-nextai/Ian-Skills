# Ian-Skills

Public skill collection.

## 技能清单

| 技能 | 说明 |
| --- | --- |
| [Amz-Excess-Inventory-Check](Amz-Excess-Inventory-Check/) | Amazon FBA 冗余库存检查：从 Amazon CSV、Amazon MCP 报表网关或本地领星 MCP 读取库存，导出固定结构的 Excel 工作簿（分析总览 / 冗余库存 / 库存全量 / 数据检查）。 |

## Amz-Excess-Inventory-Check

分析 Amazon FBA 库存健康度，识别冗余库存与库龄风险，导出固定 schema 的 Excel 工作簿。

三种数据来源：

- **Mode A** — Amazon 后台导出的三份 CSV（FBA 库存、月度仓储费、超龄附加费）
- **Mode B** — Amazon MCP 报表网关（拉同样的三份报表）
- **Mode C** — 本地领星 MCP 网关（库存与销量；无历史费用明细，相关列留空并标注）

冗余优先级按 Amazon 的 `estimated-excess-quantity` 并结合销量与库龄判定：

| 级别 | 条件 |
| --- | --- |
| P0 紧急 | 存在 366 天以上库存 |
| P1 高 | 60 天零销，或存在 181 天以上库存 |
| P2 中 | 30 天零销，或冗余量 ≥ 可售量的 50% |
| P3 低 | 其余 Amazon 判定为冗余的库存 |

输出文件按数据快照日期命名：`amazon_fba_inventory_check_<YYYYMMDD>.xlsx`。

详细字段定义见 [references/output-schema.md](Amz-Excess-Inventory-Check/references/output-schema.md)。
