"""Build a single-sheet workbook from a merged JSONL dataset.

Only stdlib + openpyxl are used. Missing cells remain blank.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def load_jsonl(path: str | Path) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    with Path(path).open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if line:
                rows.append(json.loads(line))
    return rows


def get(row: Dict[str, Any], field: str) -> Any:
    cur: Any = row
    for part in field.split("."):
        if not isinstance(cur, dict):
            return ""
        cur = cur.get(part, "")
    if isinstance(cur, (dict, list)):
        return json.dumps(cur, ensure_ascii=False)
    if cur is None:
        return ""
    text = str(cur).strip()
    return "" if text.lower() in {"nan", "none", "null"} else text


def auto_width(rows: Sequence[Dict[str, Any]], fields: Sequence[str], max_width: int = 48) -> List[int]:
    widths: List[int] = []
    for field in fields:
        longest = len(field)
        for row in rows[:500]:
            value = get(row, field)
            if value:
                # CJK counts roughly two columns in Excel.
                display = sum(2 if ord(c) > 127 else 1 for c in value)
                longest = max(longest, min(display, max_width * 2))
        widths.append(min(max(10, longest // 2 + 2), max_width))
    return widths


def build(
    rows: List[Dict[str, Any]],
    fields: Sequence[str],
    output: str | Path,
    title: str,
    note: str,
    highlight_fields: Sequence[str] = (),
    freeze: str = "C5",
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "总表"

    header_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    title_font = Font(name="微软雅黑", bold=True, size=14, color="1F4E78")
    note_font = Font(name="微软雅黑", size=9, italic=True, color="666666")
    body_font = Font(name="微软雅黑", size=9)
    highlight_fill = PatternFill("solid", fgColor="E2F0D9")
    thin = Side(style="thin", color="D9E1F2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    last_col_letter = get_column_letter(len(fields))
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = title
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = note
    ws["A2"].font = note_font
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 32

    # Row 3 is intentionally blank as a separator.
    for col, field in enumerate(fields, 1):
        cell = ws.cell(row=4, column=col, value=field)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[4].height = 32

    highlight_cols = {fields.index(f) + 1 for f in highlight_fields if f in fields}
    for r_idx, row in enumerate(rows, 5):
        for c_idx, field in enumerate(fields, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=get(row, field))
            cell.font = body_font
            cell.alignment = Alignment(vertical="top", wrap_text=False)
            cell.border = border
            if c_idx in highlight_cols and cell.value:
                cell.fill = highlight_fill

    for col, width in enumerate(auto_width(rows, fields), 1):
        ws.column_dimensions[get_column_letter(col)].width = width

    last_row = 4 + len(rows)
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A4:{last_col_letter}{last_row}"

    out_path = Path(output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build one sheet from JSONL with stable formatting.")
    parser.add_argument("--records", required=True)
    parser.add_argument("--config", required=True, help="JSON with title, note, fields, highlight_fields")
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    rows = load_jsonl(args.records)
    build(
        rows,
        config["fields"],
        args.out,
        title=config.get("title", "信息总表"),
        note=config.get("note", "空白单元格表示源数据中没有该项。"),
        highlight_fields=config.get("highlight_fields", []),
        freeze=config.get("freeze", "C5"),
    )
    print(json.dumps({"rows": len(rows), "columns": len(config["fields"]), "out": args.out}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
