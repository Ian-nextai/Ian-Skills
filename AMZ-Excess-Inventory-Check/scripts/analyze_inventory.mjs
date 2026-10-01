import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const FONT = "Arial";
const COLUMN_MAPPING = JSON.parse(
  await fs.readFile(new URL("../references/column-mapping.json", import.meta.url), "utf8"),
);
const COLORS = {
  navy: "#163A5F",
  blue: "#2F75B5",
  lightBlue: "#DCE6F1",
  paleBlue: "#EAF2F8",
  amber: "#FFF2CC",
  red: "#FCE4D6",
  redText: "#9C0006",
  green: "#E2F0D9",
  greenText: "#375623",
  gray: "#F2F2F2",
  border: "#D9E2F3",
  text: "#1F2937",
};

const BOOLEAN_ARGS = new Set([
  "render-preview-only",
  "allow-multi-marketplace",
  "allow-multi-currency",
]);

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

function decodeCsv(buffer) {
  const bytes = buffer[0] === 0xEF && buffer[1] === 0xBB && buffer[2] === 0xBF ? buffer.subarray(3) : buffer;
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("gb18030", { fatal: true }).decode(bytes);
  }
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  const source = text.replace(/^\uFEFF/, "");
  for (let i = 0; i < source.length; i += 1) {
    const ch = source[i];
    if (quoted) {
      if (ch === '"') {
        if (source[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n") {
      row.push(field.replace(/\r$/, ""));
      if (row.some((cell) => cell !== "")) rows.push(row);
      row = [];
      field = "";
    } else {
      field += ch;
    }
  }
  if (field !== "" || row.length > 0) {
    row.push(field.replace(/\r$/, ""));
    if (row.some((cell) => cell !== "")) rows.push(row);
  }
  if (quoted) throw new Error("CSV contains an unclosed quoted field.");
  return rows;
}

function toRecords(rows) {
  if (rows.length === 0) return [];
  const headers = rows[0].map((v) => v.trim());
  return rows.slice(1).map((row) => Object.fromEntries(headers.map((header, i) => [header, row[i] ?? ""])));
}

function normalizeHeader(value) {
  return clean(value).toLowerCase().replace(/[\s_\-()/]+/g, "");
}

function buildHeaderResolver(headers, aliases) {
  const byNormalized = new Map();
  for (const header of headers) {
    const normalized = normalizeHeader(header);
    if (!byNormalized.has(normalized)) byNormalized.set(normalized, header);
  }
  const resolved = new Map();
  const missing = [];
  for (const [standard, candidates] of Object.entries(aliases)) {
    const match = [standard, ...candidates].find((candidate) => byNormalized.has(normalizeHeader(candidate)));
    if (match) resolved.set(standard, byNormalized.get(normalizeHeader(match)));
    else missing.push(standard);
  }
  return { resolved, missing };
}

function applyHeaderMapping(records, mapping) {
  return records.map((record) => {
    const out = { ...record };
    for (const [standard, source] of mapping) out[standard] = record[source] ?? "";
    return out;
  });
}

function detectReport(headers) {
  for (const type of ["fba", "storage", "aged"]) {
    const { missing } = buildHeaderResolver(headers, COLUMN_MAPPING[type].aliases);
    if (missing.length === 0) return type;
  }
  return null;
}

const clean = (v) => String(v ?? "").trim();
const norm = (v) => clean(v).toUpperCase();
function num(v, fallback = 0) {
  const s = clean(v).replace(/[$,]/g, "");
  if (s === "" || s === "--") return fallback;
  const n = Number(s.replace(/\+$/, ""));
  return Number.isFinite(n) ? n : fallback;
}
function maybeNum(v) {
  const s = clean(v).replace(/[$,]/g, "");
  if (s === "" || s === "--" || s.toLowerCase() === "none") return null;
  const n = Number(s.replace(/\+$/, ""));
  return Number.isFinite(n) ? n : null;
}
function parseDate(v) {
  const s = clean(v);
  if (!s || s.toLowerCase() === "none") return null;
  let d;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) d = new Date(s.slice(0, 10) + "T00:00:00Z");
  else if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    const [m, day, y] = s.split("/").map(Number);
    d = new Date(Date.UTC(y, m - 1, day));
  } else d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d;
}
function isoDate(d) {
  return d ? d.toISOString().slice(0, 10) : "";
}
function round(v, digits = 2) {
  if (v === null || v === undefined || !Number.isFinite(v)) return null;
  const factor = 10 ** digits;
  return Math.round((v + Number.EPSILON) * factor) / factor;
}
function sum(values) {
  return values.reduce((total, value) => total + (Number.isFinite(value) ? value : 0), 0);
}
function exactKey(country, sku, fnsku) {
  return `${norm(country)}\u001F${norm(sku)}\u001F${norm(fnsku)}`;
}
function fnskuKey(country, fnsku) {
  return `${norm(country)}\u001F${norm(fnsku)}`;
}
function excelColumn(index) {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function currencyLabel(context, suffix = "") {
  return `${context.currency || "未提供"}${suffix}`;
}
function rate(value, days) {
  return value === null ? null : round(value / days, 3);
}

function unavailableReport(label, reason) {
  return {
    file: "",
    headers: [],
    rows: [],
    available: false,
    label,
    unavailableReason: reason,
  };
}

function normalizeLingxingPayload(payload, inputJson) {
  const fbaRows = Array.isArray(payload?.fbaRows) ? payload.fbaRows : [];
  if (!fbaRows.length) throw new Error(`领星 MCP 数据为空：${inputJson}`);
  const metadata = payload.metadata || {};
  const sourceLabel = payload.sourceLabel || "领星 MCP";
  return {
    source: "lingxing-mcp",
    sourceLabel,
    metadata,
    feesAvailable: { storage: false, aged: false },
    ownerTagsAvailable: true,
    fba: {
      file: inputJson,
      headers: Object.keys(fbaRows[0]),
      rows: fbaRows,
      available: true,
      label: "领星 MCP FBA库存",
    },
    storage: unavailableReport("月度仓储费", "领星 MCP 未提供历史月度仓储费明细"),
    aged: unavailableReport("超龄附加费", "领星 MCP 未提供历史超龄库存附加费明细"),
  };
}

async function loadLingxingReports(inputJson) {
  const payload = JSON.parse(await fs.readFile(path.resolve(inputJson), "utf8"));
  return normalizeLingxingPayload(payload, path.resolve(inputJson));
}

async function loadReports(inputDir, options = {}) {
  if (options.source === "lingxing") {
    if (!options.inputJson) throw new Error("领星 MCP 模式需要 --input-json <normalized-json>");
    return loadLingxingReports(options.inputJson);
  }
  const files = (await fs.readdir(inputDir, { withFileTypes: true }))
    .filter((item) => item.isFile() && item.name.toLowerCase().endsWith(".csv"))
    .map((item) => path.join(inputDir, item.name));
  const found = new Map();
  const unknown = [];
  for (const file of files) {
    const text = decodeCsv(await fs.readFile(file));
    const rows = parseCsv(text);
    if (rows.length === 0) continue;
    const headers = rows[0].map((h) => h.trim());
    const type = detectReport(headers);
    if (!type) {
      unknown.push(path.basename(file));
      continue;
    }
    if (found.has(type)) {
      throw new Error(
        `检测到多份${COLUMN_MAPPING[type].label}：${path.basename(found.get(type).file)}, ${path.basename(file)}`,
      );
    }
    const { resolved } = buildHeaderResolver(headers, COLUMN_MAPPING[type].aliases);
    found.set(type, {
      file,
      headers,
      rows: applyHeaderMapping(toRecords(rows), resolved),
    });
  }
  const missing = ["fba", "storage", "aged"].filter((type) => !found.has(type));
  if (missing.length) {
    throw new Error(
      `缺少必需报表：${missing.map((type) => COLUMN_MAPPING[type].label).join("、")}。未识别文件：${unknown.join("、") || "无"}`,
    );
  }
  return {
    source: "amazon-csv",
    sourceLabel: "Amazon CSV",
    feesAvailable: { storage: true, aged: true },
    ownerTagsAvailable: false,
    ...Object.fromEntries(found),
  };
}

function aggregateFees(storageRows, agedRows) {
  const storage = new Map();
  for (const row of storageRows) {
    const country = clean(row.country);
    const fnsku = clean(row.fnsku);
    if (!fnsku) continue;
    const key = fnskuKey(country, fnsku);
    if (!storage.has(key)) storage.set(key, {
      country, fnsku, periods: new Set(), currencies: new Set(),
      base: 0, utilization: 0, monthly: 0, averageQty: 0, rows: 0,
    });
    const entry = storage.get(key);
    if (clean(row.month_of_charge)) entry.periods.add(clean(row.month_of_charge));
    if (clean(row.currency)) entry.currencies.add(clean(row.currency));
    entry.base += num(row.est_base_msf);
    entry.utilization += num(row.est_sus);
    entry.monthly += num(row.estimated_monthly_storage_fee);
    entry.averageQty += num(row.average_quantity_on_hand);
    entry.rows += 1;
  }

  const agedExact = new Map();
  const agedByFnskuGroups = new Map();
  for (const row of agedRows) {
    const country = clean(row.country);
    const sku = clean(row.sku);
    const fnsku = clean(row.fnsku);
    if (!fnsku) continue;
    const key = exactKey(country, sku, fnsku);
    if (!agedExact.has(key)) agedExact.set(key, {
      country, sku, fnsku, dates: new Set(), currencies: new Set(),
      qty: 0, amount: 0, tiers: new Set(), rows: 0,
    });
    const entry = agedExact.get(key);
    const d = parseDate(row["snapshot-date"]);
    if (d) entry.dates.add(isoDate(d));
    if (clean(row.currency)) entry.currencies.add(clean(row.currency));
    if (clean(row["surcharge-age-tier"])) entry.tiers.add(clean(row["surcharge-age-tier"]));
    entry.qty += num(row["qty-charged"]);
    entry.amount += num(row["amount-charged"]);
    entry.rows += 1;
  }
  for (const entry of agedExact.values()) {
    const key = fnskuKey(entry.country, entry.fnsku);
    if (!agedByFnskuGroups.has(key)) agedByFnskuGroups.set(key, []);
    agedByFnskuGroups.get(key).push(entry);
  }
  const agedUniqueFnsku = new Map();
  for (const [key, values] of agedByFnskuGroups) if (values.length === 1) agedUniqueFnsku.set(key, values[0]);
  return { storage, agedExact, agedUniqueFnsku, agedByFnskuGroups };
}

function validateUniqueness(reports, options) {
  const markets = new Set();
  const currencies = new Set();
  const periods = new Set();
  const agedDates = new Set();
  const duplicates = [];
  const fbaKeys = new Map();
  const agedKeys = new Map();

  for (const row of reports.fba.rows) {
    markets.add(clean(row.marketplace));
    currencies.add(clean(row.currency));
    const key = exactKey(row.marketplace, row.sku, row.fnsku);
    fbaKeys.set(key, (fbaKeys.get(key) || 0) + 1);
  }
  for (const row of reports.storage.rows) {
    markets.add(clean(row.country));
    currencies.add(clean(row.currency));
    periods.add(clean(row.month_of_charge));
  }
  for (const row of reports.aged.rows) {
    markets.add(clean(row.country));
    currencies.add(clean(row.currency));
    const d = parseDate(row["snapshot-date"]);
    if (d) agedDates.add(isoDate(d));
    const key = `${exactKey(row.country, row.sku, row.fnsku)}\u001F${clean(row["surcharge-age-tier"])}\u001F${clean(row["snapshot-date"])}`;
    agedKeys.set(key, (agedKeys.get(key) || 0) + 1);
  }
  for (const market of reports.metadata?.markets || []) if (clean(market)) markets.add(clean(market));
  if (clean(reports.metadata?.currency)) currencies.add(clean(reports.metadata.currency));
  for (const [label, map] of [["FBA", fbaKeys], ["超龄附加费", agedKeys]]) {
    const duplicateCount = [...map.values()].filter((count) => count > 1).reduce((a, b) => a + b - 1, 0);
    if (duplicateCount > 0) duplicates.push(`${label}${duplicateCount}条`);
  }

  const errors = [];
  if (duplicates.length) errors.push(`重复主键：${duplicates.join("，")}`);
  if (markets.size > 1 && !options.allowMultiMarketplace) {
    errors.push(`检测到多个站点（${[...markets].filter(Boolean).sort().join(", ")}）；请拆分报表或显式传 --allow-multi-marketplace`);
  }
  if (currencies.size > 1 && !options.allowMultiCurrency) {
    errors.push(`检测到多种币种（${[...currencies].filter(Boolean).sort().join(", ")}）；请拆分报表或显式传 --allow-multi-currency`);
  }
  if (periods.size > 1) errors.push(`月度仓储费包含多个期间：${[...periods].filter(Boolean).sort().join(", ")}`);
  if (agedDates.size > 1) errors.push(`超龄附加费包含多个日期：${[...agedDates].sort().join(", ")}`);
  if (errors.length) throw new Error(errors.join("；"));

  const currency = [...currencies].find(Boolean) || "";
  return {
    markets: [...markets].filter(Boolean).sort(),
    currencies: [...currencies].filter(Boolean).sort(),
    currency: currency || clean(reports.metadata?.currency),
    periods: [...periods].filter(Boolean).sort(),
    agedDates: [...agedDates].sort(),
    source: reports.source || "amazon-csv",
    sourceLabel: reports.sourceLabel || "Amazon CSV",
    feesAvailable: reports.feesAvailable || { storage: true, aged: true },
    ownerTagsAvailable: reports.ownerTagsAvailable === true,
    metadata: reports.metadata || {},
  };
}

function classifyExcess(row) {
  if (row.age366 > 0) return "P0 紧急";
  if (row.agedQty > 0 || row.age181 > 0) return "P1 高";
  if (row.sales60 === 0) return "P1 高";
  if (row.sales30 === 0 || row.excessRatio >= 0.5) return "P2 中";
  return "P3 低";
}

function excessAction(row) {
  if (row.age366 > 0) return "优先清理：停止补货，核查降价、促销或移除";
  if (row.agedQty > 0) return "核查历史超龄计费；停止补货并评估清仓或移除";
  if (row.sales60 === 0) return "停止补货；排查Listing与兼容性，评估清仓或移除";
  if (row.age181 > 0) return "停止或减少补货；加速销售，避免继续产生超龄费";
  if (row.sales30 === 0) return "暂停补货；检查流量、转化及广告";
  return "降低补货量；按预计冗余数量制定促销计划";
}

function oldestAgeBand(row) {
  if (row.age366 > 0) return "366天以上";
  if (row.age271365 > 0) return "271-365天";
  if (row.age181270 > 0) return "181-270天";
  if (row.age91180 > 0) return "91-180天";
  if (row.age090 > 0) return "0-90天";
  return "无库龄数量";
}

function buildAnalysis(reports, context) {
  const fbaRows = reports.fba.rows;
  const fee = aggregateFees(reports.storage.rows, reports.aged.rows);
  const feesAvailable = reports.feesAvailable || { storage: true, aged: true };

  const matchedStorageKeys = new Set();
  const matchedAgedKeys = new Set();
  const enriched = [];

  for (const fba of fbaRows) {
    const country = clean(fba.marketplace);
    const sku = clean(fba.sku);
    const fnsku = clean(fba.fnsku);
    const eKey = exactKey(country, sku, fnsku);
    const fKey = fnskuKey(country, fnsku);
    const storage = fee.storage.get(fKey) || null;
    if (storage) matchedStorageKeys.add(fKey);
    let aged = fee.agedExact.get(eKey) || null;
    let agedMatch = "exact";
    if (!aged && fee.agedUniqueFnsku.has(fKey)) {
      aged = fee.agedUniqueFnsku.get(fKey);
      agedMatch = "unique-fnsku";
    }
    if (aged) matchedAgedKeys.add(exactKey(aged.country, aged.sku, aged.fnsku));

    const available = num(fba.available);
    const reserved = num(fba["Total Reserved Quantity"]);
    const unfulfillable = num(fba["unfulfillable-quantity"]);
    const fbaTotal = maybeNum(fba["Inventory Supply at FBA"]) ?? (available + reserved + unfulfillable);
    const sales30Value = maybeNum(fba["units-shipped-t30"]);
    const sales60Value = maybeNum(fba["units-shipped-t60"]);
    const sales30 = sales30Value;
    const sales60 = sales60Value;
    const inbound = num(fba["inbound-quantity"]);
    const snapshot = parseDate(fba["snapshot-date"]);
    const age090 = num(fba["inv-age-0-to-90-days"]);
    const age91180 = num(fba["inv-age-91-to-180-days"]);
    const age181270 = num(fba["inv-age-181-to-270-days"]);
    const age271365 = num(fba["inv-age-271-to-365-days"]);
    const age366455 = num(fba["inv-age-366-to-455-days"]);
    const age456 = num(fba["inv-age-456-plus-days"]);
    const age181 = age181270 + age271365 + age366455 + age456;
    const age366 = age366455 + age456;
    const excess = num(fba["estimated-excess-quantity"]);
    const excessRatio = available > 0 ? excess / available : excess > 0 ? 1 : null;
    // In Lingxing mode these historical fee sources do not exist. Keep them
    // null so the workbook cell stays blank; never coerce unavailable fees to 0.
    const monthlyStorage = feesAvailable.storage ? (storage?.monthly ?? 0) : null;
    const agedSurcharge = feesAvailable.aged ? (aged?.amount ?? 0) : null;
    const baseStorage = feesAvailable.storage ? (storage?.base ?? 0) : null;
    const utilizationSurcharge = feesAvailable.storage ? (storage?.utilization ?? 0) : null;
    const agedQty = feesAvailable.aged ? (aged?.qty ?? 0) : null;
    const totalFees = Number.isFinite(monthlyStorage) || Number.isFinite(agedSurcharge)
      ? (monthlyStorage ?? 0) + (agedSurcharge ?? 0)
      : null;

    const row = {
      country, sku, fnsku, asin: clean(fba.asin), productName: clean(fba["product-name"]), condition: clean(fba.condition),
      // Only the LingXing adapter supplies a listing owner; the Amazon CSV
      // modes have no such column, so they read 未提供 rather than a blank.
      principal: clean(fba.principal) || "未提供",
      // Same for LingXing's Listing tags, which come from a separate tool.
      tags: clean(fba.tags) || "未提供",
      missingSales30: sales30Value === null,
      missingSales60: sales60Value === null,
      snapshot, available, reserved, unfulfillable, fbaTotal, inbound,
      sales30, sales60,
      alert: clean(fba.alert),
      age090, age91180, age181270, age271365, age366455, age456, age181, age366,
      oldestAgeBand: "", excess, excessRatio, sellThrough: maybeNum(fba["sell-through"]), amazonDaysSupply: maybeNum(fba["days-of-supply"]),
      storagePeriod: storage ? [...storage.periods].sort().join(", ") : "",
      baseStorage, utilizationSurcharge, monthlyStorage,
      agedDates: aged ? [...aged.dates].sort().join(", ") : "", agedQty, agedSurcharge,
      totalFees,
      excessPriority: "", action: "",
      agedMatch: aged ? agedMatch : "unmatched", sourceStorage: Boolean(storage), sourceAged: Boolean(aged),
    };
    row.oldestAgeBand = oldestAgeBand(row);
    if (row.excess > 0) row.excessPriority = classifyExcess(row);
    const actions = [];
    if (row.excessPriority) actions.push(excessAction(row));
    row.action = actions.join("；");
    enriched.push(row);
  }

  const excessOrder = new Map([["P0 紧急", 0], ["P1 高", 1], ["P2 中", 2], ["P3 低", 3]]);
  const excess = enriched.filter((r) => r.excess > 0).sort((a, b) =>
    (excessOrder.get(a.excessPriority) - excessOrder.get(b.excessPriority)) ||
    (b.agedSurcharge - a.agedSurcharge) || (b.age181 - a.age181) || (b.excess - a.excess)
  );
  const excludedNonNew = enriched.filter((r) => norm(r.condition) !== "NEW" && ((r.sales30 ?? 0) > 0 || (r.sales60 ?? 0) > 0)).length;
  const unmatchedStorage = [...fee.storage.keys()].filter((key) => !matchedStorageKeys.has(key)).length;
  const unmatchedAged = [...fee.agedExact.values()].filter((r) => !matchedAgedKeys.has(exactKey(r.country, r.sku, r.fnsku))).length;
  const missingSalesFields = enriched.filter((r) => r.missingSales30 || r.missingSales60).length;
  const missingSales30 = enriched.filter((r) => r.missingSales30).length;
  const missingSales60 = enriched.filter((r) => r.missingSales60).length;
  const validSales60 = enriched.length - missingSales60;
  const missingCoverage = enriched.length ? missingSales60 / enriched.length : 0;

  return {
    enriched, excess, context,
    checks: {
      ambiguousAgedFnsku: [...fee.agedByFnskuGroups.values()].filter((v) => v.length > 1).length,
      excludedNonNew, unmatchedStorage, unmatchedAged,
      missingSalesFields, missingSales30, missingSales60, validSales60, missingCoverage,
    },
  };
}

// Historical fee columns only exist when the source actually carries the
// Monthly Storage Fee / Aged Inventory Surcharge reports. The LingXing MCP
// gateway does not, so in that mode the columns are dropped from the sheet
// entirely instead of being written as a column of blanks. The summary and
// data-check sheets still name both families as 未提供 — the disclosure is
// required, the empty per-SKU columns are not.
// The listing owner (负责人) and Listing tags (标签) come only from the LingXing
// adapter. In every Amazon mode the columns are dropped from the sheets
// entirely instead of being written as a column of 未提供 blanks.
const OWNER_TAG_HEADERS = ["负责人", "标签"];
const STORAGE_FEE_HEADERS = ["月度仓储费期间", "基础仓储费", "仓储利用率附加费", "月度仓储费"];
const AGED_FEE_HEADERS = ["超龄附加费日期", "超龄计费数量", "超龄附加费"];
const TOTAL_FEE_HEADER = "两项费用合计";

function activeHeaders(headers, feesAvailable, ownerTagsAvailable = false) {
  return headers.filter((header) => {
    if (OWNER_TAG_HEADERS.includes(header)) return ownerTagsAvailable;
    if (STORAGE_FEE_HEADERS.includes(header)) return feesAvailable.storage;
    if (AGED_FEE_HEADERS.includes(header)) return feesAvailable.aged;
    if (header === TOTAL_FEE_HEADER) return feesAvailable.storage || feesAvailable.aged;
    return true;
  });
}

const EXCESS_HEADERS = [
  "站点", "MSKU", "FNSKU", "ASIN", "负责人", "标签", "可售数量", "预留数量", "不可售数量", "在途数量",
  "亚马逊预计冗余数量", "冗余占可售比", "近30天销量", "近60天销量", "30天日均销量", "60天日均销量", "亚马逊可售天数", "售罄率",
  "库龄0-90天", "库龄91-180天", "库龄181-270天", "库龄271-365天", "库龄366天以上", "181天以上库存", "366天以上库存", "最老库龄区间",
  ...STORAGE_FEE_HEADERS, ...AGED_FEE_HEADERS, TOTAL_FEE_HEADER, "冗余优先级", "建议动作",
];

const FULL_HEADERS = [
  "站点", "MSKU", "FNSKU", "ASIN", "负责人", "标签", "可售数量", "预留数量", "不可售数量", "FBA总库存", "在途数量", "近30天销量", "近60天销量",
  "亚马逊预计冗余数量", "冗余占可售比",
  "库龄0-90天", "库龄91-180天", "库龄181-270天", "库龄271-365天", "库龄366天以上", "181天以上库存", "366天以上库存", "售罄率", "亚马逊可售天数",
  "月度仓储费期间", "月度仓储费", "超龄附加费日期", "超龄附加费", TOTAL_FEE_HEADER, "冗余优先级", "建议动作",
];

function excessRecord(r) {
  return {
    "站点": r.country,
    "MSKU": r.sku,
    "FNSKU": r.fnsku,
    "ASIN": r.asin,
    "负责人": r.principal,
    "标签": r.tags,
    "可售数量": r.available,
    "预留数量": r.reserved,
    "不可售数量": r.unfulfillable,
    "在途数量": r.inbound,
    "亚马逊预计冗余数量": r.excess,
    "冗余占可售比": r.excessRatio,
    "近30天销量": r.sales30,
    "近60天销量": r.sales60,
    "30天日均销量": rate(r.sales30, 30),
    "60天日均销量": rate(r.sales60, 60),
    "亚马逊可售天数": r.amazonDaysSupply,
    "售罄率": r.sellThrough,
    "库龄0-90天": r.age090,
    "库龄91-180天": r.age91180,
    "库龄181-270天": r.age181270,
    "库龄271-365天": r.age271365,
    // 456+ 不再单列：并入“库龄366天以上”，数量不丢失。
    "库龄366天以上": r.age366455 + r.age456,
    "181天以上库存": r.age181,
    "366天以上库存": r.age366,
    "最老库龄区间": r.oldestAgeBand,
    "月度仓储费期间": r.storagePeriod,
    "基础仓储费": round(r.baseStorage),
    "仓储利用率附加费": round(r.utilizationSurcharge),
    "月度仓储费": round(r.monthlyStorage),
    "超龄附加费日期": r.agedDates,
    "超龄计费数量": r.agedQty,
    "超龄附加费": round(r.agedSurcharge),
    "两项费用合计": round(r.totalFees),
    "冗余优先级": r.excessPriority,
    "建议动作": excessAction(r),
  };
}

function fullRecord(r) {
  return {
    "站点": r.country,
    "MSKU": r.sku,
    "FNSKU": r.fnsku,
    "ASIN": r.asin,
    "负责人": r.principal,
    "标签": r.tags,
    "可售数量": r.available,
    "预留数量": r.reserved,
    "不可售数量": r.unfulfillable,
    "FBA总库存": r.fbaTotal,
    "在途数量": r.inbound,
    "近30天销量": r.sales30,
    "近60天销量": r.sales60,
    "亚马逊预计冗余数量": r.excess,
    "冗余占可售比": r.excessRatio,
    "库龄0-90天": r.age090,
    "库龄91-180天": r.age91180,
    "库龄181-270天": r.age181270,
    "库龄271-365天": r.age271365,
    // 456+ 不再单列：并入“库龄366天以上”，数量不丢失。
    "库龄366天以上": r.age366455 + r.age456,
    "181天以上库存": r.age181,
    "366天以上库存": r.age366,
    "售罄率": r.sellThrough,
    "亚马逊可售天数": r.amazonDaysSupply,
    "月度仓储费期间": r.storagePeriod,
    "月度仓储费": round(r.monthlyStorage),
    "超龄附加费日期": r.agedDates,
    "超龄附加费": round(r.agedSurcharge),
    "两项费用合计": round(r.totalFees),
    "冗余优先级": r.excessPriority,
    "建议动作": r.action,
  };
}

// Project records through the active column list so headers and values can
// never drift out of step when a column is dropped.
function projectRows(records, headers) {
  return records.map((record) => headers.map((header) => record[header] ?? null));
}

function setColumnWidths(sheet, headers, totalRows) {
  headers.forEach((header, i) => {
    let width = 13;
    if (["MSKU", "FNSKU", "ASIN"].includes(header)) width = 18;
    if (header === "负责人") width = 22;
    if (header === "标签") width = 24;
    if (header === "建议动作") width = 40;
    if (["月度仓储费期间", "超龄附加费日期", "最老库龄区间"].includes(header)) width = 20;
    if (header.includes("日期")) width = Math.max(width, 14);
    sheet.getRangeByIndexes(0, i, totalRows, 1).format.columnWidth = width;
  });
}

function formatDataSheet(sheet, title, subtitle, headers, rows, tableName) {
  sheet.showGridLines = false;
  sheet.tabColor = COLORS.blue;
  sheet.getRange("A2").values = [[title]];
  sheet.getRange("A2").format.font = { name: FONT, size: 15, bold: true, color: COLORS.navy };
  sheet.getRange("A3").values = [[subtitle]];
  sheet.getRange("A3").format.font = { name: FONT, size: 9, italic: true, color: "#5B6573" };
  const lastCol = excelColumn(headers.length - 1);
  sheet.getRange(`A5:${lastCol}5`).values = [headers];
  if (rows.length) sheet.getRange(`A6:${lastCol}${rows.length + 5}`).values = rows;
  const usedRows = Math.max(rows.length + 5, 6);
  const used = sheet.getRange(`A2:${lastCol}${usedRows}`);
  used.format.font = { name: FONT, size: 9, color: COLORS.text };
  used.format.verticalAlignment = "center";
  sheet.getRange(`A5:${lastCol}5`).format = {
    fill: COLORS.navy,
    font: { name: FONT, size: 9, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#FFFFFF" },
  };
  sheet.getRange(`A5:${lastCol}${usedRows}`).format.borders = { preset: "outside", style: "thin", color: COLORS.border };
  if (rows.length) {
    const table = sheet.tables.add(`A5:${lastCol}${rows.length + 5}`, true, tableName);
    table.style = "TableStyleMedium2";
    table.showBandedRows = true;
  }
  sheet.freezePanes.freezeRows(5);
  sheet.freezePanes.freezeColumns(Math.min(5, headers.length));
  setColumnWidths(sheet, headers, usedRows);
  // Validation fix (SKILL.md "Validation"): several 9pt headers wrap to two
  // lines inside 13-wide columns ("亚马逊预计冗余数量" etc.). Excel auto-fits,
  // but renderers that do not auto-fit clipped the second line. Pin the header
  // row height so every viewer shows the full header.
  sheet.getRange(`A5:${lastCol}5`).format.rowHeight = 28;
  if (headers.includes("建议动作") && rows.length) {
    const c = excelColumn(headers.indexOf("建议动作"));
    sheet.getRange(`${c}6:${c}${rows.length + 5}`).format.wrapText = true;
  }
  headers.forEach((header, i) => {
    if (!rows.length) return;
    const col = excelColumn(i);
    const range = sheet.getRange(`${col}6:${col}${rows.length + 5}`);
    if (header.includes("日期") && !["超龄附加费日期"].includes(header)) range.format.numberFormat = "yyyy-mm-dd";
    if (["基础仓储费", "仓储利用率附加费", "月度仓储费", "超龄附加费", "两项费用合计"].includes(header)) {
      range.format.numberFormat = '#,##0.00';
    }
    if (header.includes("占可售比") || header === "售罄率") range.format.numberFormat = "0.0%";
    if (header.includes("日均") || header.includes("可售天数")) range.format.numberFormat = "0.0";
  });
  for (const priorityHeader of ["冗余优先级"]) {
    const index = headers.indexOf(priorityHeader);
    if (index < 0 || !rows.length) continue;
    const col = excelColumn(index);
    const range = sheet.getRange(`${col}6:${col}${rows.length + 5}`);
    range.conditionalFormats.add("containsText", { text: "P0", format: { fill: COLORS.red, font: { bold: true, color: COLORS.redText } } });
    range.conditionalFormats.add("containsText", { text: "P1", format: { fill: COLORS.amber, font: { bold: true, color: "#9C6500" } } });
    range.conditionalFormats.add("containsText", { text: "P2", format: { fill: COLORS.paleBlue, font: { bold: true, color: COLORS.navy } } });
    range.conditionalFormats.add("containsText", { text: "P3", format: { fill: COLORS.green, font: { bold: true, color: COLORS.greenText } } });
  }
  return sheet;
}

function countBy(rows, field, values) {
  return values.map((value) => [value, rows.filter((row) => row[field] === value).length]);
}

function buildSummary(sheet, analysis, reports) {
  sheet.showGridLines = false;
  sheet.tabColor = COLORS.navy;
  sheet.getRange("A2").values = [["Amazon FBA 冗余库存检查"]];
  sheet.getRange("A2").format.font = { name: FONT, size: 16, bold: true, color: COLORS.navy };
  const snapshots = [...new Set(analysis.enriched.map((r) => isoDate(r.snapshot)).filter(Boolean))].sort();
  sheet.getRange("A3").values = [[`库存快照：${snapshots.join(", ") || "未提供"}`]];
  sheet.getRange("A3").format.font = { name: FONT, size: 9, italic: true, color: "#5B6573" };

  const currency = currencyLabel(analysis.context);
  const storageTotal = analysis.context.feesAvailable.storage
    ? round(sum(reports.storage.rows.map((r) => num(r.estimated_monthly_storage_fee))))
    : "未提供";
  const agedTotal = analysis.context.feesAvailable.aged
    ? round(sum(analysis.enriched.map((r) => r.agedSurcharge)))
    : "未提供";
  const metrics = [
    ["指标", "结果"],
    ["库存SKU数", analysis.enriched.length],
    ["可售库存数量", sum(analysis.enriched.map((r) => r.available))],
    ["预计冗余数量", sum(analysis.excess.map((r) => r.excess))],
    ["冗余SKU数", analysis.excess.length],
    ["181天以上库存", sum(analysis.enriched.map((r) => r.age181))],
    [`月度仓储费总额（${currency}）`, storageTotal],
    [`超龄附加费（${currency}）`, agedTotal],
  ];
  sheet.getRange("A5:B12").values = metrics;
  sheet.getRange("A5:B5").format = { fill: COLORS.navy, font: { name: FONT, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
  sheet.getRange("A6:B12").format.font = { name: FONT, size: 10, color: COLORS.text };
  sheet.getRange("A5:B12").format.borders = { preset: "outside", style: "thin", color: COLORS.border };
  sheet.getRange("B11:B12").format.numberFormat = '#,##0.00';

  const excessCounts = [["冗余优先级", "SKU数"], ...countBy(analysis.excess, "excessPriority", ["P0 紧急", "P1 高", "P2 中", "P3 低"])];
  sheet.getRange("D5:E9").values = excessCounts;
  sheet.getRange("D5:E5").format = { fill: COLORS.navy, font: { name: FONT, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
  sheet.getRange("D5:E9").format.borders = { preset: "outside", style: "thin", color: COLORS.border };

  // The top-10 table drops its fee column in the same mode the data sheets do.
  const showFees = analysis.context.feesAvailable.storage || analysis.context.feesAvailable.aged;
  const topHeaders = ["MSKU", "可售", "冗余", "30天销量", "60天销量", "181天以上", ...(showFees ? ["费用合计"] : []), "优先级"];
  const topExcess = analysis.excess.slice(0, 10).map((r) => {
    const record = {
      "MSKU": r.sku, "可售": r.available, "冗余": r.excess, "30天销量": r.sales30,
      "60天销量": r.sales60, "181天以上": r.age181, "费用合计": round(r.totalFees), "优先级": r.excessPriority,
    };
    return topHeaders.map((header) => record[header] ?? null);
  });
  const lastTopCol = excelColumn(topHeaders.length - 1);
  sheet.getRange(`A16:${lastTopCol}16`).values = [topHeaders];
  if (topExcess.length) sheet.getRange(`A17:${lastTopCol}${16 + topExcess.length}`).values = topExcess;
  sheet.getRange("A15").values = [["优先处理的冗余库存"]];
  sheet.getRange("A15").format.font = { name: FONT, size: 12, bold: true, color: COLORS.navy };
  sheet.getRange(`A16:${lastTopCol}16`).format = { fill: COLORS.blue, font: { name: FONT, size: 9, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
  const topFeeIndex = topHeaders.indexOf("费用合计");
  if (topFeeIndex >= 0) {
    const feeCol = excelColumn(topFeeIndex);
    sheet.getRange(`${feeCol}17:${feeCol}${Math.max(17, 16 + topExcess.length)}`).format.numberFormat = '"$"#,##0.00';
  }

  sheet.getRange("A28").values = [["口径说明"]];
  sheet.getRange("A28").format.font = { name: FONT, size: 11, bold: true, color: COLORS.navy };
  const feeNote = analysis.context.feesAvailable.storage && analysis.context.feesAvailable.aged
    ? "月度仓储费与超龄附加费来自历史费用报表"
    : "领星 MCP 未提供月度仓储费和超龄附加费，相关字段留空且不按0处理";
  sheet.getRange("A29").values = [[`${feeNote}；库存、销量和库龄来自${reports.sourceLabel || "当前FBA库存快照"}。金额币种：${currency}。`]];
  sheet.getRange("A29:Q29").format.font = { name: FONT, size: 9, italic: true, color: "#5B6573" };
  sheet.getRange("A29:Q29").format.wrapText = true;
  sheet.getRange("A1:Q30").format.verticalAlignment = "center";
  sheet.getRange("A1:Q30").format.font = { name: FONT };
  ["A", "D"].forEach((c) => { sheet.getRange(`${c}1:${c}30`).format.columnWidth = 18; });
  sheet.getRange("A1:A30").format.columnWidth = 23;
  ["B", "E"].forEach((c) => { sheet.getRange(`${c}1:${c}30`).format.columnWidth = 13; });
  return sheet;
}

function buildChecks(sheet, analysis, reports) {
  sheet.showGridLines = false;
  sheet.tabColor = "#A5A5A5";
  sheet.getRange("A2").values = [["数据检查与计算口径"]];
  sheet.getRange("A2").format.font = { name: FONT, size: 15, bold: true, color: COLORS.navy };
  const currency = currencyLabel(analysis.context);
  const sourceName = (report) => report.available
    ? (reports.source === "lingxing-mcp"
      ? (report.label || reports.sourceLabel || "已提供")
      : (report.file ? path.basename(report.file) : (report.label || reports.sourceLabel || "已提供")))
    : "未提供";
  const storageSourceNote = reports.feesAvailable.storage
    ? `${reports.storage.rows.length}行；必需`
    : "领星 MCP 未提供；费用列留空，不按0处理";
  const agedSourceNote = reports.feesAvailable.aged
    ? `${reports.aged.rows.length}行；必需`
    : "领星 MCP 未提供；费用列留空，不按0处理";
  const storageTotal = reports.feesAvailable.storage
    ? round(sum(reports.storage.rows.map((r) => num(r.estimated_monthly_storage_fee))))
    : "未提供";
  const storageMatched = reports.feesAvailable.storage
    ? round(sum(analysis.enriched.map((r) => r.monthlyStorage)))
    : "未提供";
  const agedTotal = reports.feesAvailable.aged
    ? round(sum(reports.aged.rows.map((r) => num(r["amount-charged"]))))
    : "未提供";
  const agedMatched = reports.feesAvailable.aged
    ? round(sum(analysis.enriched.map((r) => r.agedSurcharge)))
    : "未提供";
  const sourceWarnings = (reports.metadata?.warnings || []).join("；");
  const fileRows = [
    ["检查项", "结果", "说明"],
    ["数据来源", reports.sourceLabel || "Amazon CSV", reports.source === "lingxing-mcp" ? "领星 MCP 只读库存适配" : "Amazon CSV或Amazon MCP报表"],
    ["来源限制", sourceWarnings || "无", sourceWarnings ? "请结合下方字段和口径解释阅读结果" : "无"],
    ["FBA Inventory来源", sourceName(reports.fba), `${reports.fba.rows.length}行；${reports.sourceLabel || "必需"}`],
    ["月度仓储费来源", sourceName(reports.storage), storageSourceNote],
    ["超龄附加费来源", sourceName(reports.aged), agedSourceNote],
    ["站点范围", analysis.context.markets.join(", "), "单站点校验已通过"],
    ["币种范围", analysis.context.currencies.join(", ") || "未提供", "单币种校验已通过；领星未返回币种时保留未提供"],
    ["月度仓储费期间", analysis.context.periods.join(", ") || "未提供", reports.feesAvailable.storage ? "多期间会阻止生成" : "领星 MCP 未提供历史费用期间"],
    ["超龄附加费日期", analysis.context.agedDates.join(", ") || "未提供", reports.feesAvailable.aged ? "多日期会阻止生成" : "领星 MCP 未提供历史费用日期"],
    ["重复主键", "0", "重复时停止运行，不继续汇总"],
    ["FNSKU超龄费歧义", analysis.checks.ambiguousAgedFnsku, "同站点同FNSKU多SKU时不做FNSKU兜底匹配"],
    ["未匹配月度仓储费FNSKU", analysis.checks.unmatchedStorage, "可能是历史费用对应当前已无库存的FNSKU"],
    ["未匹配超龄附加费行", analysis.checks.unmatchedAged, "可能是历史费用对应当前已无库存的SKU"],
    ["近30天销量缺失SKU", analysis.checks.missingSales30, "空白销量视为缺失，不按零销量处理；缺失时不做零销量判定"],
    ["近60天销量缺失SKU", analysis.checks.missingSales60, "空白销量视为缺失，不按零销量处理；缺失时不做零销量判定"],
    ["60天销量有效率", round(1 - analysis.checks.missingCoverage, 4), "非空销量 / 全部SKU"],
    ["FBA可售数量合计", sum(analysis.enriched.map((r) => r.available)), "应与FBA Inventory的available合计一致"],
    ["预计冗余数量合计", sum(analysis.enriched.map((r) => r.excess)), "应与FBA Inventory的estimated-excess-quantity合计一致"],
    [`月度仓储费合计（${currency}）`, storageTotal, reports.feesAvailable.storage ? "源报表合计" : "领星 MCP 未提供"],
    [`已匹配月度仓储费（${currency}）`, storageMatched, reports.feesAvailable.storage ? "未匹配历史FNSKU不会进入SKU分析" : "领星 MCP 未提供"],
    [`超龄附加费合计（${currency}）`, agedTotal, reports.feesAvailable.aged ? "源报表合计" : "领星 MCP 未提供"],
    [`已匹配超龄附加费（${currency}）`, agedMatched, reports.feesAvailable.aged ? "未匹配历史SKU不会进入SKU分析" : "领星 MCP 未提供"],
  ];
  const checkRows = fileRows.map((row) => [row[0], row[1], row[2] ?? row[3] ?? ""]);
  sheet.getRange(`A5:C${checkRows.length + 4}`).values = checkRows;
  sheet.getRange("A5:C5").format = { fill: COLORS.navy, font: { name: FONT, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
  sheet.getRange(`A5:C${checkRows.length + 4}`).format.borders = { preset: "outside", style: "thin", color: COLORS.border };
  sheet.getRange(`A6:C${checkRows.length + 4}`).format.font = { name: FONT, size: 9, color: COLORS.text };
  sheet.getRange(`A6:C${checkRows.length + 4}`).format.wrapText = true;
  const start = checkRows.length + 7;
  const rules = [
    ["规则", "定义"],
    ["冗余范围", "estimated-excess-quantity > 0"],
    ["费用口径", reports.feesAvailable.storage && reports.feesAvailable.aged
      ? "月度仓储费与超龄附加费按各自历史期间汇总；不当作当前库存的即时成本"
      : "领星 MCP 未提供月度仓储费和超龄附加费；相关单元格留空，不按0处理"],
  ];
  sheet.getRange(`A${start}:B${start + rules.length - 1}`).values = rules;
  sheet.getRange(`A${start}:B${start}`).format = { fill: COLORS.blue, font: { name: FONT, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
  sheet.getRange(`A${start}:B${start + rules.length - 1}`).format.borders = { preset: "outside", style: "thin", color: COLORS.border };
  sheet.getRange(`A${start + 1}:B${start + rules.length - 1}`).format.wrapText = true;
  sheet.getRange("A1:A60").format.columnWidth = 28;
  sheet.getRange("B1:B60").format.columnWidth = 48;
  sheet.getRange("C1:C60").format.columnWidth = 50;
  sheet.getRange("A1:C60").format.verticalAlignment = "center";
  return sheet;
}

async function createWorkbook(analysis, reports, output) {
  const workbook = Workbook.create();
  const summary = workbook.worksheets.add("分析总览");
  const excessSheet = workbook.worksheets.add("冗余库存");
  const fullSheet = workbook.worksheets.add("库存全量");
  const checks = workbook.worksheets.add("数据检查");

  buildSummary(summary, analysis, reports);
  const snapshotLabel = [...new Set(analysis.enriched.map((r) => isoDate(r.snapshot)).filter(Boolean))].sort().join(", ") || "未提供";
  const feeSubtitle = analysis.context.feesAvailable.storage && analysis.context.feesAvailable.aged
    ? "费用期间见对应字段。"
    : "领星 MCP 未提供月度仓储费和超龄附加费，相关字段留空。";
  const excessCols = activeHeaders(EXCESS_HEADERS, analysis.context.feesAvailable, analysis.context.ownerTagsAvailable);
  const fullCols = activeHeaders(FULL_HEADERS, analysis.context.feesAvailable, analysis.context.ownerTagsAvailable);
  formatDataSheet(excessSheet, "冗余库存", `库存快照：${snapshotLabel}。${feeSubtitle}`, excessCols, projectRows(analysis.excess.map(excessRecord), excessCols), "ExcessInventoryTable");
  formatDataSheet(fullSheet, "库存全量", `库存快照：${snapshotLabel}。${feeSubtitle} 包含全部风险状态。`, fullCols, projectRows(analysis.enriched.map(fullRecord), fullCols), "FullInventoryTable");
  buildChecks(checks, analysis, reports);

  workbook.recalculate();
  const summaryInspect = await workbook.inspect({ kind: "table", range: "分析总览!A1:Q30", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 17, maxChars: 12000 });
  const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan", maxChars: 5000 });
  console.log(summaryInspect.ndjson);
  console.log(errors.ndjson);

  await fs.mkdir(path.dirname(output), { recursive: true });
  const file = await SpreadsheetFile.exportXlsx(workbook);
  await file.save(output);
}

async function renderPreviewOnly(input, output, previewDir) {
  const sheetRange = {
    "分析总览": "A1:Q30",
    "冗余库存": "A1:M25",
    "库存全量": "A1:N25",
    "数据检查": "A1:C40",
  };
  if (!sheetRange[input]) throw new Error(`Unknown preview sheet: ${input}`);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(output));
  await fs.mkdir(previewDir, { recursive: true });
  const blob = await workbook.render({ sheetName: input, range: sheetRange[input], scale: 1.3, format: "png" });
  await fs.writeFile(path.join(previewDir, `${input}.png`), new Uint8Array(await blob.arrayBuffer()));
}

function renderPreviews(output, previewDir) {
  const sheets = ["分析总览", "冗余库存", "库存全量", "数据检查"];
  const scriptPath = fileURLToPath(import.meta.url);
  const failures = [];
  for (const sheetName of sheets) {
    const result = spawnSync(process.execPath, [
      scriptPath,
      "--render-preview-only",
      "--input",
      sheetName,
      "--output",
      output,
      "--preview-dir",
      previewDir,
    ], { encoding: "utf8" });
    if (result.error || result.status !== 0) {
      const detail = result.error?.message || result.stderr?.trim() || result.signal || `exit ${result.status}`;
      failures.push(`${sheetName}: ${detail}`);
      process.stderr.write(`[preview-warning] Failed to render ${sheetName}: ${detail}\n`);
    } else {
      process.stdout.write(`[preview-ok] ${sheetName}\n`);
    }
  }
  if (failures.length) process.stderr.write(`[preview-warning] ${failures.length} preview(s) unavailable; workbook export remains valid.\n`);
}

const args = parseArgs(process.argv);
if (args["render-preview-only"]) {
  if (!args.input || !args.output || !args["preview-dir"]) {
    throw new Error("--render-preview-only requires --input, --output, and --preview-dir.");
  }
  await renderPreviewOnly(args.input, path.resolve(args.output), path.resolve(args["preview-dir"]));
  process.exit(0);
}
if (!args.output || (args.source === "lingxing" ? !args["input-json"] : !args["input-dir"])) {
  throw new Error(
    "Usage: analyze_inventory.mjs --input-dir <folder> --output <file.xlsx> [--source amazon-csv] [--preview-dir <folder>] [--allow-multi-marketplace] [--allow-multi-currency] OR --source lingxing --input-json <file.json> --output <file.xlsx>",
  );
}
const source = args.source || "amazon-csv";
if (!["amazon-csv", "lingxing"].includes(source)) {
  throw new Error(`不支持的数据来源：${source}；可选 amazon-csv 或 lingxing`);
}
const reports = await loadReports(args["input-dir"] ? path.resolve(args["input-dir"]) : "", {
  source,
  inputJson: args["input-json"],
});
const context = validateUniqueness(reports, {
  allowMultiMarketplace: Boolean(args["allow-multi-marketplace"]),
  allowMultiCurrency: Boolean(args["allow-multi-currency"]),
});
const analysis = buildAnalysis(reports, context);
const output = path.resolve(args.output);
await createWorkbook(analysis, reports, output);
if (args["preview-dir"]) renderPreviews(output, path.resolve(args["preview-dir"]));

console.log(JSON.stringify({
  output,
  source: reports.source,
  rows: analysis.enriched.length,
  excessRows: analysis.excess.length,
  available: sum(analysis.enriched.map((r) => r.available)),
  excessQuantity: sum(analysis.excess.map((r) => r.excess)),
  monthlyStorageMatched: reports.feesAvailable.storage ? round(sum(analysis.enriched.map((r) => r.monthlyStorage))) : null,
  agedSurchargeMatched: reports.feesAvailable.aged ? round(sum(analysis.enriched.map((r) => r.agedSurcharge))) : null,
  checks: analysis.checks,
}, null, 2));
