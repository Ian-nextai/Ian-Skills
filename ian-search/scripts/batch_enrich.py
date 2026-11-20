"""Run the ian-search web-access workflow on a batch of records.

Strategy
--------
1. The agent uses the `web-access` skill to resolve one representative
   query per record and saves the raw results as JSONL.
2. This script applies deterministic extraction/filtering so that noisy
   snippets never overwrite an existing good field.

Input record format (one JSON object per line)::

    {"id": "example-1", "fields": {"name": "示例公司", "hall": "示例展区"}}

Input result format (one JSON object per line)::

    {"id": "example-1", "error": "", "results": [
      {"title": "...", "url": "...", "snippet": "...", "content": "..."}
    ]}

"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import (  # noqa: E402
    accept_product,
    dump_json,
    extract_product,
    fill_template,
    get_field,
    has,
    is_cn,
    load_json,
    norm,
)


def read_jsonl(path: str | Path) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    with Path(path).open("r", encoding="utf-8") as fh:
        for lineno, line in enumerate(fh, 1):
            line = line.strip()
            if not line:
                continue
            try:
                items.append(json.loads(line))
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid JSON on line {lineno} of {path}: {exc}") from exc
    return items


def build_queries(records: List[Dict[str, Any]], config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Create one query per record from configurable templates."""
    cn_tpl = config.get("query_cn") or "{name} 产品"
    en_tpl = config.get("query_en") or '"{name}" products manufacturer'
    out: List[Dict[str, Any]] = []
    for record in records:
        row = record.get("fields") or record
        name = str(get_field(row, config.get("name_field", "name")) or "").strip()
        if not has(name):
            continue
        template = cn_tpl if is_cn(name) else en_tpl
        query = fill_template(template, row, name=name, name_field=config.get("name_field", "name"))
        if not query:
            continue
        out.append({
            "id": record.get("id"),
            "name": name,
            "query": query,
            "language": "zh-CN" if is_cn(name) else "en",
        })
    return out


def _result_text(result: Dict[str, Any]) -> str:
    return " ".join(
        str(result.get(key) or "")
        for key in ("title", "snippet", "content", "description")
    )


def apply_results(
    records: List[Dict[str, Any]],
    results: List[Dict[str, Any]],
    config: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], Dict[str, int]]:
    """Merge extracted products back into records, keeping source provenance."""
    by_id: Dict[str, Dict[str, Any]] = {}
    for result in results:
        rid = str(result.get("id") or "")
        entry = by_id.setdefault(rid, {"results": []})
        if "results" in result:
            entry["results"].extend(result.get("results") or [])
        elif any(result.get(k) for k in ("title", "snippet", "content", "description")):
            # Also accept one flat result object per line.
            entry["results"].append(result)

    name_key = norm(config.get("name_field", "name"))
    product_field = config.get("product_field", "product")
    source_field = config.get("source_field", "product_source")
    link_field = config.get("link_field", "product_link")
    source_label = config.get("source_label") or "网络检索"
    product_terms = config.get("product_terms")
    generic_terms = config.get("generic_terms")
    min_source_len = int(config.get("min_source_len", 2) or 2)
    require_existing_source = bool(config.get("require_existing_source", False))

    stats = {"records": len(records), "results": len(results), "updated": 0, "skipped": 0}

    for record in records:
        rid = str(record.get("id") or "")
        row = record.get("fields") or record
        name = str(get_field(row, config.get("name_field", "name")) or "").strip()
        if name_key and name_key not in norm(name):
            pass  # name_field is a path; no extra filtering needed here

        # Skip rows that already have a better value.
        existing = get_field(row, product_field)
        existing_source = str(get_field(row, source_field) or "")
        if has(existing) and (not require_existing_source or len(existing_source) >= min_source_len):
            stats["skipped"] += 1
            continue

        snippets: List[str] = []
        url = ""
        entry = by_id.get(rid)
        if entry:
            key = norm(name)
            for result in entry["results"]:
                text = _result_text(result)
                if key and key not in norm(text):
                    continue
                snippets.append(text)
                if not url:
                    url = str(result.get("url") or "")

        phrase = extract_product(
            snippets,
            name,
            product_terms=product_terms,
            generic_terms=generic_terms,
            max_len=int(config.get("max_product_len", 80)),
        )
        phrase = accept_product(
            phrase,
            product_terms=product_terms,
            generic_terms=generic_terms,
            max_len=int(config.get("max_product_len", 80)),
        )
        if not phrase:
            stats["skipped"] += 1
            continue

        # Write back using the same dotted-path convention as common.get_field.
        target: Dict[str, Any] = row
        parts = product_field.split(".")
        for part in parts[:-1]:
            target = target.setdefault(part, {})
        target[parts[-1]] = phrase

        for field, value in ((source_field, source_label), (link_field, url)):
            if not field:
                continue
            t: Dict[str, Any] = row
            fp = field.split(".")
            for part in fp[:-1]:
                t = t.setdefault(part, {})
            t[fp[-1]] = value

        stats["updated"] += 1

    return records, stats


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply web-access search results to a JSONL batch.")
    parser.add_argument("--records", required=True, help="Input JSONL records")
    parser.add_argument("--results", required=True, help="JSONL results produced via web-access")
    parser.add_argument("--config", required=True, help="JSON config with field mapping and labels")
    parser.add_argument("--queries-out", help="Optional path to write generated queries as JSONL")
    parser.add_argument("--out", required=True, help="Output JSONL records")
    parser.add_argument("--stats", help="Optional path for JSON stats")
    args = parser.parse_args()

    config = load_json(args.config, default={})
    records = read_jsonl(args.records)
    results = read_jsonl(args.results)

    queries = build_queries(records, config)
    if args.queries_out:
        with Path(args.queries_out).open("w", encoding="utf-8") as fh:
            for query in queries:
                fh.write(json.dumps(query, ensure_ascii=False) + "\n")

    updated_records, stats = apply_results(records, results, config)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as fh:
        for record in updated_records:
            fh.write(json.dumps(record, ensure_ascii=False) + "\n")

    if args.stats:
        dump_json(args.stats, stats)
    print(json.dumps(stats, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
