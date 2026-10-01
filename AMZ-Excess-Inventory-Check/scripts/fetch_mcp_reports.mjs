#!/usr/bin/env node
/*
 * Optional helper for pulling the three required Amazon reports from a streamable HTTP MCP
 * gateway. The analysis does not depend on this helper; it is only a convenience
 * entry point for environments that expose the amazon-mcp-gateway tools.
 */

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const REPORT_TYPES = {
  fba: "GET_FBA_INVENTORY_PLANNING_DATA",
  storage: "GET_FBA_STORAGE_FEE_CHARGES_DATA",
  aged: "GET_FBA_FULFILLMENT_LONGTERM_STORAGE_FEE_CHARGES_DATA",
};

function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument: ${token}`);
    const key = token.slice(2);
    const value = argv[i + 1];
    if (!value || value.startsWith("--")) throw new Error(`Missing value for --${key}`);
    out[key] = value;
    i += 1;
  }
  return out;
}

async function postMcp(url, auth, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Authorization": auth,
      "Content-Type": "application/json",
      "Accept": "application/json, text/event-stream",
    },
    body: JSON.stringify(body),
  });
  if (!response.ok) throw new Error(`MCP HTTP ${response.status}: ${await response.text()}`);
  const text = await response.text();
  const dataLine = text.split(/\r?\n/).find((line) => line.startsWith("data:"));
  const jsonText = dataLine ? dataLine.slice(5).trim() : text;
  return JSON.parse(jsonText);
}

async function callTool(url, auth, id, name, args) {
  const result = await postMcp(url, auth, {
    jsonrpc: "2.0",
    id,
    method: "tools/call",
    params: { name, arguments: args },
  });
  if (result.error) throw new Error(`${name}: ${result.error.message || JSON.stringify(result.error)}`);
  const text = (result.result?.content ?? [])
    .filter((item) => item.type === "text")
    .map((item) => item.text)
    .join("\n");
  if (!text) throw new Error(`${name}: MCP returned no text content`);
  return JSON.parse(text);
}

function extractCsv(payload) {
  const data = payload?.data;
  if (typeof data?.data === "string") return data.data;
  if (typeof data === "string") return data;
  throw new Error("MCP report payload does not contain a CSV body");
}

const args = parseArgs(process.argv);
const url = args.url || process.env.AMAZON_GATEWAY_MCP_URL;
const auth = args.auth || process.env.AMAZON_GATEWAY_MCP_TOKEN;
const outputDir = path.resolve(args["output-dir"] || "./mcp_reports");

if (!url || !auth) {
  throw new Error(
    "Missing MCP endpoint. Set AMAZON_GATEWAY_MCP_URL and AMAZON_GATEWAY_MCP_TOKEN, or pass --url and --auth.",
  );
}

await fs.mkdir(outputDir, { recursive: true });
const outputs = {};
for (const [key, reportType] of Object.entries(REPORT_TYPES)) {
  const payload = await callTool(url, auth, key, "amazon_sp_export_report", {
    reportType,
    format: "csv",
    maxRows: 100000,
  });
  const csv = extractCsv(payload);
  const file = path.join(outputDir, `${key}.csv`);
  await fs.writeFile(file, csv, "utf8");
  outputs[key] = file;
}

console.log(JSON.stringify({ outputDir, reports: outputs }, null, 2));
