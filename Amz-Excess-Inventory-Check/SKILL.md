---
name: Amz-Excess-Inventory-Check
description: Analyze Amazon FBA inventory from Amazon CSV reports, the existing Amazon MCP report gateway, or a local LingXing MCP gateway, then export a fixed-schema Excel workbook for excess inventory and aging stock. Use for recurring Amazon inventory health, excess-stock, and carrying-cost analysis.
---

# Amazon FBA 冗余库存检查

Create a repeatable Excel analysis from one of these supported input modes:

### Mode A: Amazon CSV exports

Use the three original Amazon CSV reports:

1. FBA Inventory Report
2. Monthly Storage Fees Report
3. Aged Inventory Surcharge Report

### Mode B: Existing Amazon MCP report gateway

Keep the existing `scripts/fetch_mcp_reports.mjs` flow. It fetches the same three report types and then uses the normal CSV analyzer.

### Mode C: Local LingXing MCP gateway

Use `scripts/fetch_lingxing_mcp.mjs` with the local gateway that exposes `help -> search -> action`. It reads `get_fba_stock_list` (which also carries the ASIN-level listing owner in `asin_principal_list`), optionally reads `query_product_performance_asin_lists` for 30/60-day sales, optionally reads `sales_relation_tag_list` for Listing tags, and writes a normalized JSON payload for the analyzer. LingXing does not provide the historical Monthly Storage Fees or Aged Inventory Surcharge detail needed by this workbook, so in this mode the analyzer drops those fee columns from the two data sheets entirely and marks both fee families as unavailable in the summary and data-check sheets. Never substitute `estimated_storage_cost_next_month` for historical monthly storage fees, and never substitute 0 for an unavailable fee.

Read [references/output-schema.md](references/output-schema.md) when the user asks what fields are exported, wants the schema changed, or when source headers have changed.

## Workflow

1. For Mode A and Mode B, confirm that all three Amazon report types are present by their headers, not their filenames. If a type is missing or appears more than once, stop and identify the exact issue instead of mixing periods.
2. For Mode C, confirm that the normalized JSON contains FBA rows. Missing LingXing fee data is expected and must not stop the workbook.
3. Treat FBA Inventory as the authoritative current-inventory snapshot. Use its `available` value for sellable inventory and its `inbound-quantity` value for inbound inventory. Match Amazon fee reports by marketplace and FNSKU/MSKU only in Mode A or Mode B.
4. Keep marketplaces separate. Never merge US, CA, or MX inventory into one SKU row.
5. Keep current sellable inventory and inbound inventory separate. Inbound is not sellable inventory.
6. Retain non-New and zero-sales rows in the full-inventory output and data-quality checks.
7. Label fee periods explicitly. Amazon monthly storage fees and aged-inventory surcharges are historical charges. In Mode C, drop the fee columns from both data sheets and display both fee families as `未提供` in the summary and data-check sheets; do not describe them as zero.
7b. `负责人` and `标签` are LingXing-only. In Mode A and Mode B, drop both columns from both data sheets entirely — never write them as a `未提供` placeholder column. Do not export a `成色` column either; the source `condition` is used only for the non-New row count on `数据检查`. The 456+ age bucket is never a separate column in any mode; merge it into `库龄366天以上`.
8. Run the bundled script and deliver only the validated `.xlsx` workbook. Do not convert the output to CSV because the fixed workbook contains multiple views and formatting.
9. Treat preview rendering as a post-export visual check, not as a prerequisite for producing the workbook. If preview rendering fails, report the limitation explicitly and continue delivering the validated `.xlsx`.
10. **The delivered file name must carry the run date** — see [File naming](#file-naming). Never deliver a bare `amazon_fba_inventory_check.xlsx`: repeated runs overwrite each other and the reader cannot tell which snapshot a file holds.

## File naming

Name the workbook:

```
amazon_fba_inventory_check_<YYYYMMDD>.xlsx
```

The date is the **data snapshot date** (`snapshot-date` from FBA Inventory), not
the day you happen to run the analysis — a workbook re-generated on Monday from
Sunday's snapshot still carries Sunday's date. When the source carries no
snapshot date (or several marketplaces disagree), fall back to the generation
date in the report's own timezone and say so on `数据检查`.

Examples: `amazon_fba_inventory_check_20261001.xlsx`,
`amazon_fba_inventory_check_20260930.xlsx`.

The same rule applies to any other artifact this skill produces (JSON dumps,
preview folders): append the snapshot date rather than reusing a fixed name.

## Execution

Use the spreadsheet runtime. Immediately before the first workbook creation command in the turn, run the required spreadsheet artifact-operation marker once.

Create a conversation-specific temporary directory, copy `scripts/analyze_inventory.mjs` into it, and create a `node_modules` symlink to `$CODEX_PRIMARY_RUNTIME_NODE_MODULES`. Run the copied script with `$CODEX_PRIMARY_RUNTIME_NODE`:

```bash
$CODEX_PRIMARY_RUNTIME_NODE scripts/analyze_inventory.mjs \
  --input-dir /absolute/path/to/csv-folder \
  --output /absolute/path/to/outputs/amazon_fba_inventory_check_20261001.xlsx \
  --preview-dir /absolute/path/to/temp/previews
```

When the configured Amazon MCP gateway is available, it is an optional alternative to upload. Fetch the same three reports first, then pass its output folder to the analyzer:

```bash
$CODEX_PRIMARY_RUNTIME_NODE scripts/fetch_mcp_reports.mjs \
  --url "$AMAZON_GATEWAY_MCP_URL" \
  --auth "$AMAZON_GATEWAY_MCP_TOKEN" \
  --output-dir /absolute/path/to/mcp-reports
```

The fetch helper expects `--auth` to contain the complete authorization header value, such as `Bearer …`. The analyzer auto-detects report types, joins records, calculates excess-stock priorities, creates the four fixed worksheets, scans for formula errors, and exports the workbook. When `--preview-dir` is provided, each preview is rendered in an isolated child process after export; a preview failure is reported as a warning and does not invalidate the workbook.

For a local LingXing MCP gateway, keep the endpoint and authorization value in environment variables or pass them as arguments without printing them:

```bash
$CODEX_PRIMARY_RUNTIME_NODE scripts/fetch_lingxing_mcp.mjs \
  --url "$LINGXING_MCP_URL" \
  --auth "$LINGXING_MCP_AUTH" \
  --output /absolute/path/to/temp/lingxing_mcp_data.json \
  --end-date 2026-09-20

$CODEX_PRIMARY_RUNTIME_NODE scripts/analyze_inventory.mjs \
  --source lingxing \
  --input-json /absolute/path/to/temp/lingxing_mcp_data.json \
  --output /absolute/path/to/outputs/amazon_fba_inventory_check_20261001.xlsx \
  --preview-dir /absolute/path/to/temp/previews
```

The LingXing adapter defaults to `http://127.0.0.1:3211/mcp` when `--url` and `LINGXING_MCP_URL` are absent. It uses read-only tools, follows the gateway's `help -> search -> action` sequence, paginates FBA inventory, sweeps Listing tags in batches of 100, and continues with inventory output if the sales or tag call is unavailable. Use `--skip-sales` or `--skip-tags` only when the caller explicitly wants to omit that lookup. If multiple sites are returned, pass `--allow-multi-marketplace` to the analyzer or run one site at a time.

## Decision rules

### Excess inventory

Include a SKU when `estimated-excess-quantity > 0`.

- `P0 紧急`: current 366+ day inventory.

`库龄366天以上` (and therefore the P0 test) includes the 456+ bucket — 456 天以上 is
merged into it rather than exported as its own column.
- `P1 高`: no sales in 60 days or any 181+ day inventory.
- `P2 中`: no sales in 30 days or estimated excess is at least 50% of available inventory.
- `P3 低`: all other Amazon-identified excess inventory.

An aged-inventory charge is historical evidence and does not alone make a SKU `P0`.

## Validation

Before delivery:

- Confirm the file name carries the snapshot date: `amazon_fba_inventory_check_<YYYYMMDD>.xlsx`.
- Confirm the workbook has exactly these sheets in order: `分析总览`, `冗余库存`, `库存全量`, `数据检查`.
- Reconcile total available inventory and excess quantity to FBA Inventory.
- In Mode A or Mode B, reconcile total monthly storage fee and aged surcharge to their source reports. In Mode C, report both as `未提供` and confirm the fee columns are absent from both data sheets rather than present and blank or zero.
- Review rendered previews for every sheet and fix clipped headers, unreadable widths, or broken formatting. If the renderer is unavailable, inspect the workbook contents directly and report that visual rendering could not be completed.
- Report material unmatched rows, duplicate keys, inventory discrepancies, missing dates, or missing sales fields. In Mode C, explicitly report the missing snapshot date, the combined 365+ age-band limitation, and any unmatched sales rows. Do not silently coerce missing data to zero when it changes a risk decision.
