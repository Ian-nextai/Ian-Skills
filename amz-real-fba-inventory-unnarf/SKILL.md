---
name: amz-real-fba-inventory-unnarf
description: Check REAL local FBA inventory per Amazon marketplace by reading the Inventory Ledger Summary report and filtering on the physical Location column, excluding NARF remote-fulfillment shared stock. Exports one sheet per marketplace plus a summary workbook. Use for cross-site FBA inventory checks, NARF-adjusted stock, and local-warehouse reconciliation.
---

# Amazon 真实 FBA 库存（排除远程配送 / NARF）

按站点导出**物理存放在该国 FBA 仓**的库存，导出 Excel：汇总页 + 每个站点一个明细 sheet + 数据检查页。

## 为什么不能用 FBA Inventory API / All Inventory 报告

`amazon_sp_get_inventory` 和 `GET_FBA_MYI_ALL_INVENTORY_DATA` 给的是“**可发往该站点**”的口径：

- 加拿大、墨西哥站会把 NARF 远程配送可用的**美国仓库存**计入 fulfillable；
- 在测试账户上可以观察到，多个 SKU 的 CA 站可售数量与美国站完全一致，说明这是远程配送共享口径；
- 因此这些数字**不能**当作加拿大/墨西哥本地 FBA 库存。

本 skill 使用 `GET_LEDGER_SUMMARY_VIEW_DATA`（Inventory Ledger Summary），按 `Location` 列（FC 所在国家）过滤，只保留物理在站点所在国的结存。

## 数据来源

- 主数据：`GET_LEDGER_SUMMARY_VIEW_DATA`
  - 账户级报告，一次覆盖 US/CA/MX/BR；
  - 列：`Date`、`FNSKU`、`ASIN`、`MSKU`、`Title`、`Disposition`、`Ending Warehouse Balance`、`Location`、`Store` 等；
  - `Location` 是仓库所在国家代码（US、CA、MX、BR…），这是区分本地仓与 NARF 的唯一可靠字段；
  - 行按**日期倒序**排列，最新日期在最前。
- 站点列表：`amazon_sp_list_marketplaces`，只保留 `Amazon.*` 真实站点，排除 `Non-Amazon` 与 `Invoicing/Shadow`。
- 报告选择：`amazon_sp_list_reports` 列出近 N 天 DONE 的 Ledger 报告，每个站点选 `dataEndTime` 最新且 `dataStartTime < dataEndTime` 的报告（排除空报告）。

## 执行

### 环境变量

```bash
export AMAZON_GATEWAY_MCP_URL="https://<your-amazon-gateway>/mcp"
export AMAZON_GATEWAY_MCP_TOKEN="Bearer <token>"
```

### 1. 拉取并解析库存

```bash
node scripts/fetch_real_fba_inventory.mjs \
  --output outputs/real_fba_inventory.json
```

可选参数：

- `--sites US,CA,MX,BR` 只处理指定国家；
- `--lookback-days 45` 报告回溯天数（默认 45）；
- `--max-rows 1000` 单次读取行数上限（默认 1000，接口硬上限）。

脚本会：

1. 列出参与站点；
2. 为每个站点选择最新可用的 Ledger 报告；
3. 读取报告文档（最多 1000 行）；
4. 取报告中最新的**完整**日期分组；
5. 按 `Location` 过滤站点所在国；
6. 按 ASIN + MSKU + Disposition 汇总 `Ending Warehouse Balance`；
7. 输出 JSON 并打印各站点合计。

### 2. 生成 Excel

```bash
python scripts/build_workbook.py \
  --input outputs/real_fba_inventory.json \
  --output amz_real_fba_inventory_unnarf_<YYYYMMDD>.xlsx
```

若本机没有 `openpyxl`：

```bash
pip install openpyxl
# 或
uv run scripts/build_workbook.py --input ... --output ...
```

脚本带 PEP 723 头，`uv run` 会自动安装依赖。

## 输出结构

工作表顺序固定：

1. `汇总` — 每个站点一行：ASIN 数、结存总数、可售、不可售、可售占比、报告 ID、台账日期与口径说明；
2. 每个站点一个 sheet：`美国 US`、`加拿大 CA`、`墨西哥 MX`、`巴西 BR`（按实际参与站点生成）；
3. `数据检查` — 报告范围、行数、截断与日期完整性、Location 汇总、警告。

`数据检查` 只展示生成时间、快照日期和数据来源；**不写入**网关地址、账户名称、Seller ID、区域等敏感信息。

站点 sheet 列：

```
ASIN | MSKU | FNSKU | 产品名称 | 结存总数 | 可售 |
不可售合计 | 仓库损坏 | 客户损坏 | 承运损坏 | 分销商损坏 |
残次 | 过期 | 其他不可售 | 台账日期
```

完整字段定义见 [references/output-schema.md](references/output-schema.md)。

## 判定规则

### 最新日期

- 只在报告包含的日期中取**最大日期**；
- Ledger 行按日期倒序排列，因此最新日期分组在文档最前面；
- 如果最新日期分组的行数等于接口返回的总行数（1000），说明该分组可能被截断，**必须**在 `数据检查` 中标红并向用户说明，不能把不完整的数据当作完整库存；
- 周/月粒度报告按期末结存计算，`台账粒度` 列会标注“日/月”。

### 本地仓判定

- `Location` 必须等于站点国家代码，例如 CA 站只取 `Location = CA`；
- `Location = US` 的库存**永远不**计入 CA/MX/BR（NARF 共享库存）；
- 某站点完全没有对应 `Location` 行时，表示该国 FBA 仓无库存，不是数据缺失；在 sheet 中写一条说明并在汇总标“无本地仓库存记录”。

### 不可售

`Disposition` 不为 `SELLABLE` 的全部计入不可售：

- `WAREHOUSE_DAMAGED` → 仓库损坏
- `CUSTOMER_DAMAGED` → 客户损坏
- `CARRIER_DAMAGED` → 承运损坏
- `DISTRIBUTOR_DAMAGED` → 分销商损坏
- `DEFECTIVE` → 残次
- `EXPIRED` → 过期
- 其他未知值 → 其他不可售（`数据检查` 中列出名称）

## 文件命名

```
amz_real_fba_inventory_unnarf_<YYYYMMDD>.xlsx
```

日期用**台账快照日期**（报告中最新的完整日期），不是运行日期。多站点日期不一致时，文件名用最新日期，`汇总` 与 `数据检查` 中逐站标注各自日期。

## 校验清单

交付前必须确认：

- 工作表顺序为 `汇总` → 各站点 → `数据检查`；
- 每个站点 sheet 的结存总数 = 该站 `Location` 在最新日期的 `Ending Warehouse Balance` 之和；
- 各站点合计 = 最新日期所有已映射 Location 之和，未映射 Location 已在 `数据检查` 列出；
- `数据检查` 中所有报告均 `日期倒序=是`、`最新日完整=是`；否则不得声称数据完整；
- 汇总页“可售占比”分母为结存总数，不是 FBA 可售口径；
- 明确告知用户 Ledger 有 **2–7 天延迟**，不是实时库存；
- 不要把 FBA Inventory API / All Inventory 的 fulfillable 数字混入本表。

## 已知限制

- `amazon_sp_get_report_document` 单次最多 1000 行且没有 `offset` 参数。Ledger 行按日期倒序，所以通常最新日期分组可以完整读取；脚本会校验并在截断风险时告警。
- `amazon_sp_export_report` 无法指定 `reportId`，只会按自己的“最新报告”逻辑选取，结果不可控，因此本 skill 不使用它。
- 台账报告是账户级：US/CA/MX/BR 在同一份报告中，必须按 `Location` 分流。
- 站点本地仓库存为 0 时不会凭空补行；空 sheet 加说明是正确输出。
