#!/usr/bin/env node
/*
 * Read-only adapter for the local LingXing MCP gateway.
 *
 * The gateway exposes a help -> search -> action chain. This helper keeps that
 * protocol detail out of the workbook analyzer and writes a small normalized
 * JSON payload. Historical monthly storage fees and aged-inventory surcharges
 * are intentionally not fabricated: the payload marks both sources absent so
 * the analyzer leaves those workbook cells blank.
 */

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const BOOLEAN_ARGS = new Set(["skip-sales", "skip-tags"]);
const DEFAULT_URL = "http://127.0.0.1:3211/mcp";

function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument: ${token}`);
    const key = token.slice(2);
    if (BOOLEAN_ARGS.has(key)) {
      out[key] = true;
      continue;
    }
    const value = argv[i + 1];
    if (!value || value.startsWith("--")) throw new Error(`Missing value for --${key}`);
    out[key] = value;
    i += 1;
  }
  return out;
}

function clean(value) {
  return String(value ?? "").trim();
}

function normalizeKey(value) {
  const key = clean(value).toUpperCase();
  return ["", "-", "NONE", "NULL"].includes(key) ? "" : key;
}

function maybeNumber(value) {
  const text = clean(value).replace(/[$,]/g, "");
  if (!text || text === "--" || text.toLowerCase() === "none") return null;
  const number = Number(text.replace(/\+$/, ""));
  return Number.isFinite(number) ? number : null;
}

function firstNonEmpty(...values) {
  return values.map(clean).find(Boolean) || "";
}

function firstNumber(...values) {
  for (const value of values) {
    const number = maybeNumber(value);
    if (number !== null) return number;
  }
  return null;
}

function sumNumbers(...values) {
  return values.reduce((total, value) => total + (maybeNumber(value) ?? 0), 0);
}

function parseDate(text) {
  const value = clean(text);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return null;
  const date = new Date(`${value}T00:00:00Z`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDate(date) {
  return date.toISOString().slice(0, 10);
}

function shiftDate(dateText, days) {
  const date = parseDate(dateText) || new Date();
  date.setUTCDate(date.getUTCDate() + days);
  return formatDate(date);
}

function countryFromRow(row) {
  const direct = firstNonEmpty(row.country, row.marketplace, row.country_code, row.site);
  if (direct && !/^AMAZON_/i.test(direct)) return direct;
  const url = firstNonEmpty(row.site_url, row.amazon_url);
  const host = url.match(/amazon\.([a-z.]+)/i)?.[1]?.toLowerCase() || "";
  const countryByHost = {
    "com": "US",
    "ca": "CA",
    "com.mx": "MX",
    "co.uk": "UK",
    "de": "DE",
    "fr": "FR",
    "it": "IT",
    "es": "ES",
    "co.jp": "JP",
    "com.au": "AU",
    "nl": "NL",
    "se": "SE",
    "pl": "PL",
    "be": "BE",
    "sg": "SG",
    "in": "IN",
  };
  if (countryByHost[host]) return countryByHost[host];
  const text = `${clean(row.name)} ${clean(row.seller_name)} ${clean(row.seller_group_name)}`.toUpperCase();
  if (text.includes("美国") || text.includes("US")) return "US";
  if (text.includes("加拿大") || text.includes("CA")) return "CA";
  if (text.includes("墨西哥") || text.includes("MX")) return "MX";
  return "";
}

function rowKeys(row) {
  return [row.asin, row.msku, row.sku, row.seller_sku, row.local_sku]
    .map(normalizeKey)
    .filter(Boolean);
}

// LingXing exposes the ASIN-level listing owner as `asin_principal_list`
// (an array — a listing can have several owners). `asin_principal_arr` is
// always empty in the current gateway response and is only a fallback.
// Listings with no assigned owner come back as an empty list; surface that as
// 未分配 instead of blanking the cell, so the gap stays visible.
function principalNames(row) {
  for (const value of [row.asin_principal_list, row.asin_principal_arr, row.asin_principal]) {
    const list = Array.isArray(value)
      ? value.map(clean).filter(Boolean)
      : clean(value)
        ? clean(value).split(/[,，]/).map(clean).filter(Boolean)
        : [];
    if (list.length) return [...new Set(list)].join(", ");
  }
  return "未分配";
}

// Listing tags live behind a separate read tool (`sales_relation_tag_list`),
// which is addressed by shop id + MSKU rather than ASIN, so the inventory rows
// have to be echoed back to it in batches. `get_fba_stock_list` itself returns
// no tag field.
function rowSid(row) {
  return clean(row.sid ?? row.group_by_sid ?? row.group_by_seller_id);
}

function rowRelationId(row) {
  return firstNonEmpty(row.seller_sku, row.msku, row.sku);
}

function tagLookupKey(sid, relationId) {
  return sid && relationId ? `${sid}\u001F${relationId}` : "";
}

const TAG_BATCH_SIZE = 100; // gateway cap for bindDetail

// The gateway sits behind a proxy that intermittently answers 502 on request
// bursts — a tag sweep is ~34 sequential calls, so one bad answer used to drop
// the whole column. Back off briefly and retry.
async function withRetry(label, fn, attempts = 3) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;
      if (attempt < attempts) await new Promise((resolve) => setTimeout(resolve, 500 * attempt));
    }
  }
  throw new Error(`${label}（重试 ${attempts} 次后仍失败）：${lastError.message}`);
}

async function fetchTagMap(url, auth, rawRows) {
  const map = new Map();
  const detail = [];
  const seen = new Set();
  for (const row of rawRows) {
    const key = tagLookupKey(rowSid(row), rowRelationId(row));
    if (!key || seen.has(key)) continue;
    seen.add(key);
    detail.push({ relationId: rowRelationId(row), sid: rowSid(row) });
  }
  let id = 500;
  for (let i = 0; i < detail.length; i += TAG_BATCH_SIZE) {
    const payload = await withRetry(`Listing标签批次 ${i / TAG_BATCH_SIZE + 1}/${Math.ceil(detail.length / TAG_BATCH_SIZE)}`, () =>
      gatewayAction(url, auth, (id += 1), "sales_relation_tag_list", {
        bindDetail: detail.slice(i, i + TAG_BATCH_SIZE),
      }),
    );
    // Listings without tags are omitted from the response entirely.
    const list = Array.isArray(payload) ? payload : rowsFromPayload(payload);
    for (const item of list) {
      const key = tagLookupKey(clean(item.sid), clean(item.relationId));
      const names = (item.tagInfos || []).map((tag) => clean(tag.tagName)).filter(Boolean);
      if (key && names.length) map.set(key, [...new Set(names)]);
    }
  }
  return map;
}

function pickSalesValue(row, preferredKeys) {
  let zero = null;
  for (const key of preferredKeys) {
    const value = maybeNumber(row[key]);
    if (value === null) continue;
    if (value > 0) return value;
    zero ??= value;
  }
  return zero;
}

function makeSalesMap(rows, windowDays) {
  const preferredKeys = windowDays === 30
    ? ["volume", "volume_30d", "order_items"]
    : ["volume", "volume_60d", "order_items"];
  const map = new Map();
  for (const row of rows || []) {
    const value = pickSalesValue(row, preferredKeys);
    if (value === null) continue;
    for (const key of rowKeys(row)) {
      const previous = map.get(key);
      // Product-performance may return both a summary row and a date row.
      // Keep the largest non-zero value rather than double-counting duplicates.
      map.set(key, previous === undefined ? value : Math.max(previous, value));
    }
  }
  return map;
}

function lookupSales(map, keys) {
  for (const key of keys) {
    if (map.has(key)) return map.get(key);
  }
  return "";
}

function normalizeFbaRow(row, args, sales30, sales60, tagMap) {
  const sku = firstNonEmpty(row.seller_sku, row.msku, row.sku);
  const fnsku = firstNonEmpty(row.fnsku, row.fn_sku);
  const asin = firstNonEmpty(row.asin, row.parent_asin);
  const country = countryFromRow(row);
  const tags = tagMap.get(tagLookupKey(rowSid(row), rowRelationId(row))) || [];
  const available = firstNumber(row.afn_fulfillable_quantity, row.available_total, row.total_fulfillable_quantity, row.quantity) ?? 0;
  const reserved = firstNumber(row.afn_reserved_quantity, row.reserved_customerorders, row.reserved_fc_processing) ?? 0;
  const unfulfillable = firstNumber(row.afn_unsellable_quantity, row.total_unfulfillable_quantity) ?? 0;
  const inbound = sumNumbers(row.afn_inbound_working_quantity, row.afn_inbound_shipped_quantity, row.afn_inbound_receiving_quantity);
  const fbaTotal = firstNumber(row.available_total, row.total_fulfillable_quantity, row.total_onhand_quantity) ?? (available + reserved + unfulfillable);
  const age365Plus = firstNumber(row.inv_age_365_plus_days, row.inv_age_365_plus) ?? 0;
  const key = rowKeys(row);
  const metadataSnapshot = clean(args["snapshot-date"]);
  return {
    "snapshot-date": metadataSnapshot,
    sku,
    fnsku,
    asin,
    principal: principalNames(row),
    tags: tags.length ? tags.join(", ") : "无标签",
    "product-name": firstNonEmpty(row.product_name, row.item_name),
    condition: firstNonEmpty(row.condition, "New"),
    available,
    marketplace: country,
    "Total Reserved Quantity": reserved,
    "unfulfillable-quantity": unfulfillable,
    "Inventory Supply at FBA": fbaTotal,
    "units-shipped-t30": lookupSales(sales30, key),
    "units-shipped-t60": lookupSales(sales60, key),
    alert: firstNonEmpty(row.warn_status_name, row.warn_status),
    "sell-through": firstNonEmpty(row.sell_through),
    "days-of-supply": firstNonEmpty(row.historical_days_of_supply, row.long_term_historical_days_of_supply),
    "estimated-excess-quantity": firstNonEmpty(row.estimated_excess_quantity),
    "inv-age-0-to-90-days": firstNumber(row.inv_age_0_to_90_days, row.inv_age_0_to_90) ?? 0,
    "inv-age-91-to-180-days": firstNumber(row.inv_age_91_to_180_days, row.inv_age_91_to_180) ?? 0,
    "inv-age-181-to-270-days": firstNumber(row.inv_age_181_to_270_days, row.inv_age_181_to_270) ?? 0,
    "inv-age-271-to-365-days": firstNumber(row.inv_age_271_to_365_days, row.inv_age_271_to_365) ?? 0,
    // LingXing exposes a combined 365+ bucket. Keep it visible in the
    // existing 366-455 column and record the approximation in metadata.
    "inv-age-366-to-455-days": age365Plus,
    "inv-age-456-plus-days": 0,
    "inbound-quantity": inbound,
    currency: firstNonEmpty(row.currency, args.currency),
    "lingxing-estimated-storage-cost-next-month": firstNonEmpty(row.estimated_storage_cost_next_month),
    "lingxing-source-key": key[0] || "",
  };
}

function unwrapActionPayload(payload) {
  return payload?.data?.data ?? payload?.data ?? payload;
}

async function postMcp(url, auth, body) {
  const timeoutMs = Number(process.env.LINGXING_MCP_TIMEOUT_MS || 15000);
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: auth,
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
      "MCP-Protocol-Version": "2025-03-26",
    },
    signal: AbortSignal.timeout(Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 15000),
    body: JSON.stringify(body),
  });
  const text = await response.text();
  if (!response.ok) throw new Error(`MCP HTTP ${response.status}: ${text.slice(0, 500)}`);
  const candidates = text.split(/\r?\n/)
    .filter((line) => line.startsWith("data:"))
    .map((line) => line.slice(5).trim())
    .filter(Boolean)
    .reverse();
  for (const candidate of candidates) {
    try {
      return JSON.parse(candidate);
    } catch {
      // Continue in case the stream included a non-JSON event before the result.
    }
  }
  try {
    return JSON.parse(text.trim());
  } catch {
    throw new Error(`MCP returned invalid JSON: ${text.slice(0, 500)}`);
  }
}

async function callGatewayTool(url, auth, id, name, argumentsValue) {
  const response = await postMcp(url, auth, {
    jsonrpc: "2.0",
    id,
    method: "tools/call",
    params: { name, arguments: argumentsValue },
  });
  if (response.error) throw new Error(`${name}: ${response.error.message || JSON.stringify(response.error)}`);
  const text = (response.result?.content ?? [])
    .filter((item) => item.type === "text")
    .map((item) => item.text)
    .join("\n");
  if (!text) throw new Error(`${name}: MCP returned no text content`);
  const payload = JSON.parse(text);
  if (payload.success === false || payload.code === 0) {
    throw new Error(`${name}: ${payload.msg || JSON.stringify(payload)}`);
  }
  return payload;
}

async function gatewayHelp(url, auth) {
  const tools = [];
  // The gateway's full catalog is paginated and can be slow. Query the two
  // relevant capability groups instead of walking every tool on every run.
  for (const [index, query] of ["库存", "产品表现"].entries()) {
    try {
      const payload = await callGatewayTool(url, auth, 1 + index, "help", { limit: 50, offset: 0, query });
      tools.push(...(payload.data?.tools || []));
    } catch {
      // Search/action below remains the source of truth; an unavailable help
      // page should not prevent the inventory adapter from trying read-only calls.
    }
  }
  return tools;
}

async function gatewaySearch(url, auth, id, toolId) {
  return callGatewayTool(url, auth, id, "search", { toolId });
}

async function gatewayAction(url, auth, id, toolId, params) {
  const payload = await callGatewayTool(url, auth, id, "action", { toolId, params });
  return unwrapActionPayload(payload);
}

function rowsFromPayload(payload) {
  if (Array.isArray(payload?.list)) return payload.list;
  if (Array.isArray(payload?.data?.list)) return payload.data.list;
  return [];
}

async function fetchFbaRows(url, auth, args) {
  const rows = [];
  const length = 5000;
  let offset = 0;
  let total = null;
  do {
    const payload = await gatewayAction(url, auth, 10 + offset, "get_fba_stock_list", {
      search_value: "",
      is_hide_zero_stock: "0",
      offset,
      length,
      fulfillment_channel_type: "FBA",
      query_fba_storage_quantity_list: false,
      ...(args.sid ? { sid: args.sid } : {}),
    });
    const page = rowsFromPayload(payload);
    rows.push(...page);
    total = maybeNumber(payload?.total) ?? maybeNumber(payload?.data?.total) ?? rows.length;
    if (!page.length || page.length < length) break;
    offset += page.length;
  } while (rows.length < total);
  return { rows, total: total ?? rows.length };
}

async function fetchSalesRows(url, auth, args, windowDays, id) {
  const endDate = clean(args["end-date"]) || formatDate(new Date());
  const params = {
    offset: 0,
    length: 5000,
    start_date: shiftDate(endDate, -(windowDays - 1)),
    end_date: endDate,
    date_type: "purchase",
    summary_field: "asin",
    summary_field_level1: "asin",
    date_view_order_type: 0,
    turn_on_summary: 1,
    sort_field: "volume",
    sort_type: "desc",
    search_value: [],
    ...(args.sid ? { sids: args.sid } : {}),
    ...(args.mid ? { mids: args.mid } : {}),
  };
  const payload = await gatewayAction(url, auth, id, "query_product_performance_asin_lists", params);
  return rowsFromPayload(payload);
}

const args = parseArgs(process.argv);
const url = args.url || process.env.LINGXING_MCP_URL || DEFAULT_URL;
const auth = args.auth || process.env.LINGXING_MCP_AUTH || process.env.LINGXING_MCP_TOKEN;
const output = path.resolve(args.output || "./lingxing_mcp_data.json");
if (!auth) throw new Error("缺少领星 MCP 鉴权。请通过 --auth 或 LINGXING_MCP_AUTH 提供完整 Authorization 值。");

const tools = await gatewayHelp(url, auth);
const toolIds = new Set(tools.map((tool) => tool.toolId));
await gatewaySearch(url, auth, 2, "get_fba_stock_list");
const fba = await withRetry("FBA库存", () => fetchFbaRows(url, auth, args));

const warnings = [
  "领星 MCP 未提供历史月度仓储费明细；相关字段留空，不按0处理。",
  "领星 MCP 未提供历史超龄库存附加费明细；相关字段留空，不按0处理。",
  "领星 MCP 仅提供合并365+库龄；已映射至现有366-455天列，456+无法进一步拆分。",
];
if (!clean(args["snapshot-date"])) {
  warnings.push("领星 MCP 库存接口未返回快照日期；库存快照日期显示未提供。可通过 --snapshot-date 补充。");
}
let sales30 = new Map();
let sales60 = new Map();
if (!args["skip-sales"]) {
  try {
    await gatewaySearch(url, auth, 3, "query_product_performance_asin_lists");
    const rows30 = await withRetry("产品表现30天", () => fetchSalesRows(url, auth, args, 30, 30));
    const rows60 = await withRetry("产品表现60天", () => fetchSalesRows(url, auth, args, 60, 31));
    sales30 = makeSalesMap(rows30, 30);
    sales60 = makeSalesMap(rows60, 60);
    if (!sales30.size && !sales60.size) warnings.push("领星产品表现接口未返回可匹配销量，销量字段留空。");
  } catch (error) {
    warnings.push(`领星产品表现接口未成功返回销量：${error.message}；销量字段留空。`);
  }
} else {
  warnings.push("未调用领星产品表现接口；销量字段留空。");
}

let tagMap = new Map();
if (!args["skip-tags"]) {
  try {
    await gatewaySearch(url, auth, 4, "sales_relation_tag_list");
    tagMap = await fetchTagMap(url, auth, fba.rows);
    if (!tagMap.size) warnings.push("领星 Listing 标签接口未返回任何标签；标签列显示无标签。");
  } catch (error) {
    warnings.push(`领星 Listing 标签接口未成功返回标签：${error.message}；标签列显示无标签。`);
  }
} else {
  warnings.push("未调用领星 Listing 标签接口；标签字段留空。");
}

const fbaRows = fba.rows.map((row) => normalizeFbaRow(row, args, sales30, sales60, tagMap));
const unassignedPrincipal = fbaRows.filter((row) => row.principal === "未分配").length;
if (unassignedPrincipal) {
  warnings.push(
    `领星未返回 ${unassignedPrincipal} 行的 Listing 负责人（asin_principal_list 为空）；这些行在负责人列显示为未分配。`,
  );
}
const marketplaces = [...new Set(fbaRows.map((row) => clean(row.marketplace)).filter(Boolean))].sort();
const currencies = [...new Set(fbaRows.map((row) => clean(row.currency)).filter(Boolean))].sort();
const payload = {
  schemaVersion: 1,
  source: "lingxing-mcp",
  sourceLabel: "领星 MCP（本地网关）",
  retrievedAt: new Date().toISOString(),
  metadata: {
    endpoint: url,
    tools: [
      "get_fba_stock_list",
      ...(sales30.size || sales60.size ? ["query_product_performance_asin_lists"] : []),
      ...(tagMap.size ? ["sales_relation_tag_list"] : []),
    ],
    rows: fbaRows.length,
    totalFromGateway: fba.total,
    markets: marketplaces,
    currencies,
    currency: currencies.length === 1 ? currencies[0] : "",
    snapshotDate: clean(args["snapshot-date"]),
    salesMatchedRows: fbaRows.filter((row) => row["units-shipped-t30"] !== "" || row["units-shipped-t60"] !== "").length,
    taggedRows: fbaRows.filter((row) => row.tags !== "无标签").length,
    distinctTags: [...new Set(tagMap.size ? [...tagMap.values()].flat() : [])].sort(),
    warnings,
  },
  fbaRows,
};

await fs.mkdir(path.dirname(output), { recursive: true });
await fs.writeFile(output, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
console.log(JSON.stringify({
  output,
  source: payload.source,
  rows: fbaRows.length,
  totalFromGateway: fba.total,
  salesMatchedRows: payload.metadata.salesMatchedRows,
  taggedRows: payload.metadata.taggedRows,
  distinctTags: payload.metadata.distinctTags,
  fees: { monthlyStorage: "unavailable", agedSurcharge: "unavailable" },
  warnings,
}, null, 2));
