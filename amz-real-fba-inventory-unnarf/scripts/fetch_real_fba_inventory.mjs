#!/usr/bin/env node
/*
 * Pull the newest usable Inventory Ledger Summary report from the amazon_gateway
 * MCP endpoint and extract REAL local FBA inventory per marketplace.
 *
 * Why the ledger: the FBA Inventory API and the All Inventory report report a
 * "fulfillable to this marketplace" view that includes NARF remote-fulfillment
 * stock sitting in another country's fulfillment centers. The ledger's
 * Location column is the physical country of the fulfillment center, so
 * filtering Location == marketplace country excludes NARF shared stock.
 */

import fs from "node:fs/promises";
import process from "node:process";
import path from "node:path";

const REPORT_TYPE = "GET_LEDGER_SUMMARY_VIEW_DATA";
const MAX_DOC_ROWS = 1000;

function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument: ${token}`);
    const key = token.slice(2);
    if (key === "help") {
      out.help = true;
      continue;
    }
    const value = argv[i + 1];
    if (!value || value.startsWith("--")) throw new Error(`Missing value for --${key}`);
    out[key] = value;
    i += 1;
  }
  return out;
}

function usage() {
  return [
    "Usage:",
    "  node scripts/fetch_real_fba_inventory.mjs [options]",
    "",
    "Options:",
    "  --url <mcp-url>        MCP endpoint (or AMAZON_GATEWAY_MCP_URL)",
    "  --auth <header>        Full authorization header, e.g. 'Bearer ...' (or AMAZON_GATEWAY_MCP_TOKEN)",
    "  --output <file.json>   Output JSON path (default ./outputs/real_fba_inventory_<date>.json)",
    "  --sites <US,CA,...>    Limit to these country codes",
    "  --lookback-days <n>    How far back to list ledger reports (default 45)",
    "  --max-rows <n>         Max document rows to read, 1-1000 (default 1000)",
    "  --help                 Show this help",
  ].join("\n");
}

const args = parseArgs(process.argv);
if (args.help) {
  console.log(usage());
  process.exit(0);
}

const url = args.url || process.env.AMAZON_GATEWAY_MCP_URL;
const auth = args.auth || process.env.AMAZON_GATEWAY_MCP_TOKEN;
const lookbackDays = Number(args["lookback-days"] || 45);
const maxRows = Math.min(Number(args["max-rows"] || MAX_DOC_ROWS), MAX_DOC_ROWS);
const siteFilter = args.sites
  ? new Set(args.sites.split(",").map((s) => s.trim().toUpperCase()).filter(Boolean))
  : null;

if (!url || !auth) {
  console.error("Missing MCP endpoint. Set AMAZON_GATEWAY_MCP_URL and AMAZON_GATEWAY_MCP_TOKEN, or pass --url and --auth.");
  console.error(usage());
  process.exit(1);
}

// ---------------------------------------------------------------------------
// MCP client (streamable HTTP, handles JSON or SSE responses)
// ---------------------------------------------------------------------------

let sessionId = null;
let nextId = 1;

async function mcpPost(body, attempt = 0) {
  const headers = {
    "Content-Type": "application/json",
    "Accept": "application/json, text/event-stream",
    Authorization: auth,
  };
  if (sessionId) headers["Mcp-Session-Id"] = sessionId;

  let response;
  try {
    response = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });
  } catch (err) {
    if (attempt < 2) {
      await sleep(1000 * (attempt + 1));
      return mcpPost(body, attempt + 1);
    }
    throw new Error(`MCP request failed: ${err.message}`);
  }

  const sid = response.headers.get("mcp-session-id");
  if (sid) sessionId = sid;

  if (response.status === 429 || response.status >= 500) {
    if (attempt < 2) {
      await sleep(1500 * (attempt + 1));
      return mcpPost(body, attempt + 1);
    }
  }
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`MCP HTTP ${response.status}: ${text.slice(0, 400)}`);
  }

  const contentType = (response.headers.get("content-type") || "").toLowerCase();
  const text = await response.text();
  if (!text.trim()) return null;

  if (contentType.includes("text/event-stream")) {
    const dataLines = text
      .split(/\r?\n/)
      .filter((line) => line.startsWith("data:"))
      .map((line) => line.slice(5).replace(/^ /, ""));
    if (!dataLines.length) {
      if (body.id === undefined || body.id === null) return null;
      throw new Error("MCP SSE response contained no data line");
    }
    if (body.id === undefined || body.id === null) return null;
    const parsed = dataLines
      .map((line) => {
        try {
          return JSON.parse(line);
        } catch {
          return null;
        }
      })
      .filter(Boolean);
    return parsed.find((item) => item.id === body.id) ?? parsed[parsed.length - 1] ?? null;
  }

  const parsed = JSON.parse(text);
  if (body.id === undefined || body.id === null) return null;
  if (Array.isArray(parsed)) return parsed.find((item) => item.id === body.id) ?? null;
  return parsed;
}

async function rpc(method, params) {
  const id = nextId++;
  const response = await mcpPost({ jsonrpc: "2.0", id, method, params });
  if (!response) throw new Error(`No MCP response for ${method}`);
  if (response.error) throw new Error(`${method}: ${response.error.message || JSON.stringify(response.error)}`);
  return response.result;
}

async function notify(method, params) {
  await mcpPost({ jsonrpc: "2.0", method, params });
}

async function callTool(name, toolArgs) {
  const result = await rpc("tools/call", { name, arguments: toolArgs });
  if (result?.isError) {
    const message = (result.content || []).map((item) => item.text).join(" ");
    throw new Error(`${name}: ${message}`);
  }
  const text = (result?.content || [])
    .filter((item) => item.type === "text")
    .map((item) => item.text)
    .join("\n");
  if (!text) throw new Error(`${name}: empty tool result`);
  return JSON.parse(text);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function unquote(value) {
  if (typeof value !== "string") return value;
  let out = value.trim();
  while (out.length >= 2 && out.startsWith('"') && out.endsWith('"')) {
    out = out.slice(1, -1).trim();
  }
  return out;
}

function toNumber(value) {
  const n = Number.parseFloat(unquote(value));
  return Number.isFinite(n) ? n : 0;
}

// Returns { ts, kind, label }. Ledger Date is either MM/DD/YYYY (daily) or
// MM/YYYY (monthly). Monthly rows are month-end balances.
function parseLedgerDate(value) {
  const raw = String(unquote(value) || "").trim();
  const parts = raw.split("/").map((part) => part.trim());
  if (parts.length === 3) {
    const [mm, dd, yyyy] = parts.map(Number);
    if ([mm, dd, yyyy].every(Number.isFinite)) {
      return { ts: Date.UTC(yyyy, mm - 1, dd), kind: "daily", label: raw };
    }
  }
  if (parts.length === 2) {
    const [mm, yyyy] = parts.map(Number);
    if ([mm, yyyy].every(Number.isFinite)) {
      return { ts: Date.UTC(yyyy, mm - 1, 1), kind: "monthly", label: raw };
    }
  }
  return { ts: Number.NaN, kind: "unknown", label: raw };
}

function isoDateFromTs(ts) {
  return new Date(ts).toISOString().slice(0, 10);
}

function compactCompanyName(name) {
  return String(name || "").trim();
}

const SHEET_NAMES = {
  US: "美国 US",
  CA: "加拿大 CA",
  MX: "墨西哥 MX",
  BR: "巴西 BR",
  GB: "英国 UK",
  DE: "德国 DE",
  FR: "法国 FR",
  IT: "意大利 IT",
  ES: "西班牙 ES",
  JP: "日本 JP",
  AU: "澳洲 AU",
  SG: "新加坡 SG",
  AE: "阿联酋 AE",
  IN: "印度 IN",
  NL: "荷兰 NL",
  SE: "瑞典 SE",
  PL: "波兰 PL",
  BE: "比利时 BE",
  TR: "土耳其 TR",
  SA: "沙特 SA",
};

function sheetNameFor(countryCode, used) {
  const base = SHEET_NAMES[countryCode] || `站点 ${countryCode}`;
  let name = base;
  let suffix = 2;
  while (used.has(name)) {
    name = `${base} (${suffix})`;
    suffix += 1;
  }
  if (name.length > 31) name = name.slice(0, 31);
  used.add(name);
  return name;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  await rpc("initialize", {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "real-fba-inventory-unnarf", version: "1.0.0" },
  });
  await notify("notifications/initialized", {});

  const marketplacePayload = await callTool("amazon_sp_list_marketplaces", {});
  const account = marketplacePayload.account || {};
  const marketplaces = (marketplacePayload.data?.payload || [])
    .map((entry) => ({
      id: entry.marketplace?.id,
      name: entry.marketplace?.name || "",
      countryCode: String(entry.marketplace?.countryCode || "").toUpperCase(),
      storeName: entry.storeName || "",
      participating: entry.participation?.isParticipating !== false,
    }))
    .filter((entry) => entry.id && entry.countryCode && entry.participating)
    // Real Amazon marketplaces only: skip Non-Amazon stores and shadow/invoicing marketplaces.
    .filter((entry) => /^Amazon\./i.test(entry.name) || /^Amazon\.com/i.test(entry.name))
    .filter((entry) => !/Non-Amazon|Invoicing|Shadow/i.test(entry.name))
    .filter((entry) => !siteFilter || siteFilter.has(entry.countryCode));

  if (!marketplaces.length) throw new Error("No participating Amazon marketplaces matched the request.");

  const createdSince = new Date(Date.now() - lookbackDays * 24 * 60 * 60 * 1000).toISOString();
  const reports = [];
  let listToken = null;
  for (let page = 0; page < 6; page += 1) {
    const listArgs = {
      reportTypes: [REPORT_TYPE],
      processingStatuses: ["DONE"],
      createdSince,
      pageSize: 100,
    };
    if (listToken) listArgs.nextToken = listToken;
    const listed = await callTool("amazon_sp_list_reports", listArgs);
    const pageReports = listed.data?.reports || [];
    reports.push(...pageReports);
    listToken = listed.data?.nextToken || null;
    if (!listToken) break;
  }

  const usableReports = reports.filter((report) => {
    if (!report.dataStartTime || !report.dataEndTime) return false;
    const start = Date.parse(report.dataStartTime);
    const end = Date.parse(report.dataEndTime);
    // Empty ad-hoc reports (created with no range) have start == end.
    return Number.isFinite(start) && Number.isFinite(end) && end > start + 60 * 1000;
  });

  if (!usableReports.length) {
    throw new Error(`No usable DONE ${REPORT_TYPE} report found in the last ${lookbackDays} days.`);
  }

  // Choose the report with the newest dataEndTime for each marketplace.
  const chosen = new Map(); // reportId -> { report, marketplaces: [] }
  const missing = [];
  for (const marketplace of marketplaces) {
    const candidates = usableReports
      .filter((report) => (report.marketplaceIds || []).includes(marketplace.id))
      .sort((a, b) => Date.parse(b.dataEndTime) - Date.parse(a.dataEndTime));
    if (!candidates.length) {
      missing.push(marketplace);
      continue;
    }
    const report = candidates[0];
    if (!chosen.has(report.reportId)) chosen.set(report.reportId, { report, marketplaces: [] });
    chosen.get(report.reportId).marketplaces.push(marketplace);
  }

  if (!chosen.size) throw new Error("No ledger report covers any requested marketplace.");

  const warnings = [];
  for (const marketplace of missing) {
    warnings.push(`站点 ${marketplace.countryCode} (${marketplace.name}) 没有可用的 Ledger 报告。`);
  }

  const reportResults = [];
  const sites = {};

  for (const { report, marketplaces: reportMarketplaces } of chosen.values()) {
    const document = await callTool("amazon_sp_get_report_document", {
      reportId: report.reportId,
      maxRows,
    });
    const data = document.data || {};
    const columns = (data.columns || []).map(unquote);
    const rows = data.rows || [];
    const columnIndex = Object.fromEntries(columns.map((name, index) => [name, index]));
    for (const required of ["Date", "FNSKU", "ASIN", "MSKU", "Title", "Disposition", "Ending Warehouse Balance", "Location"]) {
      if (columnIndex[required] === undefined) {
        throw new Error(`Ledger report ${report.reportId} is missing column: ${required}`);
      }
    }

    const parsedRows = [];
    for (const raw of rows) {
      const date = parseLedgerDate(raw[columnIndex.Date]);
      if (!Number.isFinite(date.ts)) continue;
      parsedRows.push({
        ts: date.ts,
        dateLabel: date.label,
        dateKind: date.kind,
        asin: String(unquote(raw[columnIndex.ASIN]) || "").trim(),
        msku: String(unquote(raw[columnIndex.MSKU]) || "").trim(),
        fnsku: String(unquote(raw[columnIndex.FNSKU]) || "").trim(),
        title: String(unquote(raw[columnIndex.Title]) || "").trim(),
        disposition: String(unquote(raw[columnIndex.Disposition]) || "").trim().toUpperCase(),
        quantity: toNumber(raw[columnIndex["Ending Warehouse Balance"]]),
        location: String(unquote(raw[columnIndex.Location]) || "").trim().toUpperCase(),
      });
    }

    if (!parsedRows.length) {
      warnings.push(`报告 ${report.reportId} 未解析出任何台账行。`);
      continue;
    }

    const orderingVerified = parsedRows.every((row, index) => index === 0 || row.ts <= parsedRows[index - 1].ts);
    const maxTs = Math.max(...parsedRows.map((row) => row.ts));
    const latestRows = parsedRows.filter((row) => row.ts === maxTs);
    const latestDateLabel = latestRows[0]?.dateLabel || isoDateFromTs(maxTs);
    const dateKind = latestRows[0]?.dateKind || "unknown";

    // The document is capped at maxRows. Ledger rows are ordered newest-first,
    // so the newest date group is complete as long as it occupies fewer rows
    // than the whole returned document. If every returned row shares the same
    // date, truncation may have cut into that date's group.
    const latestGroupComplete = latestRows.length < parsedRows.length || parsedRows.length < maxRows;
    if (!orderingVerified) {
      warnings.push(`报告 ${report.reportId} 的台账行不是按日期倒序排列，最新日期分组可能不完整，已按现有行计算并标记。`);
    }
    if (!latestGroupComplete) {
      warnings.push(`报告 ${report.reportId} 的最新日期 ${latestDateLabel} 分组达到行数上限 (${maxRows})，该日期数据可能被截断。`);
    }
    if (data.truncated) {
      warnings.push(`报告 ${report.reportId} 共 ${data.rowCount} 行，单次最多读取 ${maxRows} 行；更早日期已被截断（不影响最新日期）。`);
    }

    const locationTotals = {};
    const locationRows = {};
    for (const row of latestRows) {
      locationTotals[row.location] = (locationTotals[row.location] || 0) + row.quantity;
      locationRows[row.location] = (locationRows[row.location] || 0) + 1;
    }

    const reportSiteCodes = new Set();
    for (const marketplace of reportMarketplaces) {
      const countryCode = marketplace.countryCode;
      reportSiteCodes.add(countryCode);
      const siteRows = latestRows.filter((row) => row.location === countryCode);

      const byKey = new Map();
      for (const row of siteRows) {
        const key = `${row.asin} ${row.msku}`;
        if (!byKey.has(key)) {
          byKey.set(key, {
            asin: row.asin,
            msku: row.msku,
            fnsku: row.fnsku,
            title: row.title,
            quantity: 0,
            dispositions: {},
          });
        }
        const item = byKey.get(key);
        item.quantity += row.quantity;
        if (!item.fnsku && row.fnsku) item.fnsku = row.fnsku;
        if (!item.title && row.title) item.title = row.title;
        item.dispositions[row.disposition] = (item.dispositions[row.disposition] || 0) + row.quantity;
      }

      const items = [...byKey.values()]
        .filter((item) => item.quantity !== 0 || Object.values(item.dispositions).some((q) => q !== 0))
        .sort((a, b) => b.quantity - a.quantity);

      const sellable = items.reduce((sum, item) => sum + (item.dispositions.SELLABLE || 0), 0);
      const total = items.reduce((sum, item) => sum + item.quantity, 0);

      sites[countryCode] = {
        countryCode,
        marketplaceId: marketplace.id,
        marketplaceName: marketplace.name,
        sheetName: sheetNameFor(countryCode, new Set(Object.values(sites).map((site) => site.sheetName))),
        reportId: report.reportId,
        snapshotDate: latestDateLabel,
        dateKind,
        asinCount: items.length,
        totalUnits: total,
        sellableUnits: sellable,
        nonSellableUnits: total - sellable,
        items,
      };
    }

    const unmappedLocations = Object.keys(locationTotals)
      .filter((location) => !reportSiteCodes.has(location))
      .map((location) => ({
        location,
        rows: locationRows[location],
        units: locationTotals[location],
      }));

    reportResults.push({
      reportId: report.reportId,
      reportType: report.reportType,
      dataStartTime: report.dataStartTime,
      dataEndTime: report.dataEndTime,
      marketplaceIds: report.marketplaceIds || [],
      rowCount: data.rowCount ?? null,
      returnedRows: parsedRows.length,
      truncated: Boolean(data.truncated),
      orderingVerified,
      latestDate: latestDateLabel,
      latestDateKind: dateKind,
      latestDateRows: latestRows.length,
      latestDateComplete: latestGroupComplete,
      locationTotals,
      locationRows,
      unmappedLocations,
    });

    if (unmappedLocations.length) {
      warnings.push(
        `报告 ${report.reportId} 中存在未映射到当前站点列表的 Location：` +
          unmappedLocations.map((item) => `${item.location}(${item.units}件)`).join("、"),
      );
    }
  }

  const snapshotDate = Object.values(sites)
    .map((site) => site.snapshotDate)
    .sort()
    .at(-1) || isoDateFromTs(Date.now());

  const allDispositions = new Set();
  for (const site of Object.values(sites)) {
    for (const item of site.items) {
      for (const disposition of Object.keys(item.dispositions)) allDispositions.add(disposition);
    }
  }
  const knownDispositions = new Set([
    "SELLABLE",
    "WAREHOUSE_DAMAGED",
    "CUSTOMER_DAMAGED",
    "DISTRIBUTOR_DAMAGED",
    "CARRIER_DAMAGED",
    "DEFECTIVE",
    "EXPIRED",
    "UNKNOWN",
  ]);
  const unknownDispositions = [...allDispositions].filter((name) => !knownDispositions.has(name));
  if (unknownDispositions.length) {
    warnings.push(`存在未识别的 Disposition：${unknownDispositions.join("、")}（已计入“其他不可售”）。`);
  }

  const safeDate = snapshotDate.replace(/[^0-9]/g, "");
  const outputPath = path.resolve(
    args.output || path.join("outputs", `real_fba_inventory_${safeDate}.json`),
  );

  const result = {
    generatedAt: new Date().toISOString(),
    gatewayUrl: url,
    account: {
      name: account.name || null,
      sellerId: account.sellerId || null,
      region: account.region || null,
    },
    reportType: REPORT_TYPE,
    snapshotDate,
    marketplaces: marketplaces.map((marketplace) => ({
      id: marketplace.id,
      name: marketplace.name,
      countryCode: marketplace.countryCode,
      storeName: marketplace.storeName,
    })),
    sites,
    reports: reportResults,
    warnings,
    method: {
      source: "GET_LEDGER_SUMMARY_VIEW_DATA",
      localStockRule:
        "Only rows whose Location equals the marketplace country code are counted. NARF/remote-fulfillment stock physically in another country is excluded.",
      dateRule:
        "Uses the newest date group in the newest usable ledger report. Ledger rows are newest-first; a date group is treated as complete when it occupies fewer rows than the returned document.",
    },
  };

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, JSON.stringify(result, null, 2), "utf8");

  const siteSummary = Object.values(sites)
    .map(
      (site) =>
        `${site.countryCode}: ${site.asinCount} ASIN / ${site.totalUnits} 件（可售 ${site.sellableUnits}，不可售 ${site.nonSellableUnits}）@ ${site.snapshotDate}`,
    )
    .join(" | ");
  console.log(JSON.stringify({ output: outputPath, snapshotDate, sites: siteSummary, warnings }, null, 2));
}

main().catch((err) => {
  console.error(`FAILED: ${err.message}`);
  process.exit(1);
});
