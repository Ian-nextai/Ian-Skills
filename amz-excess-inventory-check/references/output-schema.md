# Fixed Excel output schema

The workbook always contains four worksheets in this order. Field names remain stable when report periods and currencies change; both are shown in dedicated fields or labels.

The delivered file name always carries the snapshot date —
`amazon_fba_inventory_check_<YYYYMMDD>.xlsx`. See
[File naming](../SKILL.md#file-naming).

## 分析总览

Contains a compact KPI table, counts by excess-stock priority, and the highest-cost excess SKUs. Monetary KPI labels use the report currency.

## Shared blocks

Both data sheets open and close with the same columns. The fee block between
them appears only in the Amazon CSV modes — see [Fee columns](#fee-columns).

Leading block, in order:

1. 站点
2. MSKU
3. FNSKU
4. ASIN

`负责人` and `标签` follow ASIN **only in the LingXing MCP mode** — see
[Owner and tag columns](#owner-and-tag-columns). In every Amazon mode they are
dropped from the sheets entirely: no column, no `未提供` placeholder.

Neither data sheet carries a `成色` (item condition) column. The source
`condition` value is read only for the non-New row count on `数据检查`; it is
not exported per SKU.

Trailing block, in order:

- 冗余优先级
- 建议动作

## 冗余库存

Fields between the leading and trailing blocks, in order:

1. 可售数量
2. 预留数量
3. 不可售数量
4. 在途数量
5. 亚马逊预计冗余数量
6. 冗余占可售比
7. 近30天销量
8. 近60天销量
9. 30天日均销量
10. 60天日均销量
11. 亚马逊可售天数
12. 售罄率
13. 库龄0-90天
14. 库龄91-180天
15. 库龄181-270天
16. 库龄271-365天
17. 库龄366天以上
18. 181天以上库存
19. 366天以上库存
20. 最老库龄区间

> `库龄366天以上` = source `inv-age-366-to-455-days` + `inv-age-456-plus-days`.
> There is no separate `456天以上` column; the older bucket is merged in, so no
> quantity is lost.

## 库存全量

Fields between the leading and trailing blocks, in order:

1. 可售数量
2. 预留数量
3. 不可售数量
4. FBA总库存
5. 在途数量
6. 近30天销量
7. 近60天销量
8. 亚马逊预计冗余数量
9. 冗余占可售比
10. 库龄0-90天
11. 库龄91-180天
12. 库龄181-270天
13. 库龄271-365天
14. 库龄366天以上
15. 181天以上库存
16. 366天以上库存
17. 售罄率
18. 亚马逊可售天数

## Fee columns

These sit between the leading block and the trailing block, and exist only when
the source carries the historical fee reports — that is, the Amazon CSV modes.
The LingXing MCP gateway does not supply them, so those runs **drop the columns
entirely** rather than writing a column of blanks. `分析总览` drops the
`费用合计` column from its top-10 table in the same runs. Both runs still name
both fee families as `未提供` in the `分析总览` KPI block and on `数据检查`: the
disclosure is required, the empty per-SKU columns are not.

`冗余库存`:

- 月度仓储费期间
- 基础仓储费（原币种）
- 仓储利用率附加费（原币种）
- 月度仓储费（原币种）
- 超龄附加费日期
- 超龄计费数量
- 超龄附加费（原币种）
- 两项费用合计（原币种）

`库存全量`:

- 月度仓储费期间
- 月度仓储费（原币种）
- 超龄附加费日期
- 超龄附加费（原币种）
- 两项费用合计（原币种）

## 数据检查

Contains fixed check/result/note fields for source filenames, row counts, report periods, duplicate keys, unmatched fee rows, missing sales values, and reconciliation totals.

## Source mapping and calculations

- Current stock, inbound quantity, sales, aging buckets, sell-through, Amazon excess quantity, and Amazon days of supply come from FBA Inventory.
- Monthly storage amounts are aggregated across fulfillment centers by marketplace and FNSKU.
- Aged surcharge amounts are aggregated by marketplace, MSKU, and FNSKU.
- `181天以上库存` is the sum of the four mutually exclusive FBA age buckets from 181 days onward.
- `30天日均销量 = units-shipped-t30 / 30`; `60天日均销量 = units-shipped-t60 / 60`.
- `两项费用合计 = 月度仓储费 + 超龄附加费`. The two sources may describe different historical reference dates; the workbook displays both periods.
- `库龄366天以上` = `inv-age-366-to-455-days` + `inv-age-456-plus-days`. The 456+ bucket is merged into the 366+ column, so `366天以上库存` and the P0 rule still see every unit.
- `最老库龄区间` therefore tops out at `366天以上`; it never reports `456天以上`.

## Owner and tag columns

`负责人` (ASIN-level listing owner) and `标签` (LingXing Listing tag set) exist **only** in the LingXing MCP mode:

- `负责人` — several owners are joined with `, ` in the order LingXing returns them; listings with no assigned owner show `未分配`.
- `标签` — from the separate `sales_relation_tag_list` tool (addressed by shop id + MSKU, 100 listings per call); several tags are joined with `, `; listings with no tag show `无标签`. Use `--skip-tags` to omit the tag lookup; the column then reads `无标签`.

In every Amazon mode (CSV or MCP report) the two columns are **dropped entirely**
from both data sheets. Do not emit them with a `未提供` placeholder — a column of
placeholders is noise, and the absence of the column is self-explanatory. The
disclosure rule that applies to unavailable *fees* does not apply here.

## LingXing MCP mode

The local LingXing MCP adapter supplies the current FBA inventory view from `get_fba_stock_list`, may add 30/60-day sales from `query_product_performance_asin_lists`, the listing owner from `asin_principal_list`, and Listing tags from `sales_relation_tag_list`. LingXing's current gateway response does not supply the historical Monthly Storage Fees or Aged Inventory Surcharge detail reports used by the Amazon CSV mode.

In LingXing MCP mode:

- The fee columns are **not present at all** on `冗余库存` and `库存全量` (see [Fee columns](#fee-columns)), and `分析总览`'s top-10 table has no `费用合计` column.
- The summary and data-check sheets show `未提供` for both fee totals and explain that the values were unavailable; they are never converted to zero.
- `estimated_storage_cost_next_month` from LingXing, when present, is retained only in the normalized source payload and is not mapped to the historical `月度仓储费` column.
- LingXing's combined `365+` age bucket is mapped to the `库龄366天以上` column for continuity. Since 456+ is merged into that same column everywhere, no separate data-check note about `456+` is required.
- `负责人` and `标签` **are** present in this mode — it is the only mode that has them.
