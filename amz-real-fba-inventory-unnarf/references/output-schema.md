# Output schema

> 示例中的所有 ID、日期和数量均为占位符，不代表任何真实账户或业务数据。

## JSON produced by fetch_real_fba_inventory.mjs

```jsonc
{
  "generatedAt": "ISO timestamp",
  "gatewayUrl": "https://.../mcp",
  "account": { "name": "...", "sellerId": "...", "region": "NA" },
  "reportType": "GET_LEDGER_SUMMARY_VIEW_DATA",
  "snapshotDate": "MM/DD/YYYY",
  "marketplaces": [
    { "id": "ATVPDKIKX0DER", "name": "Amazon.com", "countryCode": "US", "storeName": "..." }
  ],
  "sites": {
    "US": {
      "countryCode": "US",
      "marketplaceId": "ATVPDKIKX0DER",
      "marketplaceName": "Amazon.com",
      "sheetName": "美国 US",
      "reportId": "12345678901",
      "snapshotDate": "MM/DD/YYYY",
      "dateKind": "daily | monthly",
      "asinCount": 0,
      "totalUnits": 0,
      "sellableUnits": 0,
      "nonSellableUnits": 0,
      "items": [
        {
          "asin": "B0XXXXXXXX",
          "msku": "EXAMPLE-SKU",
          "fnsku": "X00EXAMPLE",
          "title": "...",
          "quantity": 0,
          "dispositions": { "SELLABLE": 0 }
        }
      ]
    }
  },
  "reports": [
    {
      "reportId": "12345678901",
      "dataStartTime": "YYYY-MM-DDT00:00:00+00:00",
      "dataEndTime": "YYYY-MM-DDT23:59:59+00:00",
      "marketplaceIds": ["ATVPDKIKX0DER", "..."],
      "rowCount": 1234,
      "returnedRows": 1000,
      "truncated": true,
      "orderingVerified": true,
      "latestDate": "MM/DD/YYYY",
      "latestDateKind": "daily",
      "latestDateRows": 123,
      "latestDateComplete": true,
      "locationTotals": { "US": 0, "CA": 0 },
      "locationRows": { "US": 0, "CA": 0 },
      "unmappedLocations": []
    }
  ],
  "warnings": [],
  "method": { "source": "...", "localStockRule": "...", "dateRule": "..." }
}
```

## Excel workbook

### 汇总

| Column | Meaning |
|---|---|
| 站点 | Sheet name, e.g. `加拿大 CA` |
| 国家/地区 | Country code used against `Location` |
| Marketplace ID | SP-API marketplace id |
| 台账日期 | Latest complete date in the chosen report |
| 台账粒度 | 日 / 月 |
| ASIN 数 | Distinct ASIN + MSKU rows with non-zero balance |
| 结存总数 | Sum of `Ending Warehouse Balance` (sellable + unsellable) |
| 可售 | Sum of `SELLABLE` |
| 不可售 | 结存总数 − 可售 |
| 可售占比 | 可售 ÷ 结存总数 |
| 报告 ID | Ledger report used |
| 备注 | `仅本地仓` / `台账中无该站点本地仓库存` / completeness warning |

### One sheet per marketplace

| Column | Meaning |
|---|---|
| ASIN | Amazon ASIN |
| MSKU | Seller SKU from the ledger |
| FNSKU | Fulfillment network SKU |
| 产品名称 | Title from the ledger |
| 结存总数 | All dispositions at the latest date |
| 可售 | `SELLABLE` |
| 不可售合计 | 结存总数 − 可售 |
| 仓库损坏 | `WAREHOUSE_DAMAGED` |
| 客户损坏 | `CUSTOMER_DAMAGED` |
| 承运损坏 | `CARRIER_DAMAGED` |
| 分销商损坏 | `DISTRIBUTOR_DAMAGED` |
| 残次 | `DEFECTIVE` |
| 过期 | `EXPIRED` |
| 其他不可售 | Other/unknown dispositions, including `UNKNOWN` |
| 台账日期 | Site snapshot date |

### 数据检查

- Account and gateway identity (no token);
- One row per ledger report: id, range, total rows, rows read, truncated, date ordering verified, latest date, latest-date row count, latest-date completeness;
- `Location` totals at the latest date, plus any Location not mapped to a requested marketplace;
- Per-site reconciliation: ASIN count, total units, note;
- All warnings emitted by the fetch script.
