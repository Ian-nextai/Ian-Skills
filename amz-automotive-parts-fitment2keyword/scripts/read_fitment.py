#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# dependencies = ["openpyxl>=3.1"]
# ///
"""把汽车适配表（xlsx / xlsm / csv / tsv）读成 job.json 的 fitment 数组。

Usage:
  uv run scripts/read_fitment.py fitment.xlsx
  uv run scripts/read_fitment.py fitment.xlsx --sheet "Fitment" > job-fitment.json
  uv run scripts/read_fitment.py table.csv --forward-fill Make,Model
  uv run scripts/read_fitment.py table.xlsx --map Year=Model Year,Make=Brand,Model=Car,Engine=Size

行为:
  * 自动在前 10 行里找表头行，并把中英文列名映射到 Year / Make / Model / Cyl / Engine。
  * year 列原样输出（"2009-2022" 保持不动），展开由 build_combinations.py 完成。
  * cyl 列（缸型，如 L4 / V6）原样保留，不做归一化 —— 源数据怎么写就怎么留。
  * engine 列默认清洗：只保留排量（"5.7L V8 HEMI" -> "5.7L"）。用 --no-clean-engine 关闭。
  * 输出 JSON 到 stdout，进度/警告到 stderr。

退出码:
  0 成功   2 参数或输入不合法   3 文件读写失败
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import unicodedata
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    sys.exit(
        "error: 缺少 openpyxl。请用 `uv run scripts/read_fitment.py ...`，"
        "或先执行 `pip install openpyxl`。"
    )

EXIT_OK, EXIT_INPUT, EXIT_IO = 0, 2, 3


def _force_utf8_when_not_tty() -> None:
    """管道 / 重定向时输出 UTF-8，避免中文在 cp936 等本地编码下变乱码。

    接到终端时保持 Python 默认（Windows 下走宽字符 API，任何代码页都正常）。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            if not stream.isatty():
                stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError, OSError):
            pass


FIELDS = ("Year", "Make", "Model", "Cyl", "Engine")
HEADER_SCAN_ROWS = 10

# 表头别名。EXACT 集合先做全等匹配，再做子串兜底。
HEADER_ALIASES: dict[str, set[str]] = {
    "Year": {"year", "years", "modelyear", "my", "年份", "年款", "年度", "车型年份", "年"},
    "Make": {"make", "makes", "brand", "manufacturer", "mfr", "品牌", "厂商", "制造商", "厂牌"},
    "Model": {"model", "models", "车型", "车系", "型号", "car", "vehicle", "车款", "车辆"},
    "Cyl": {
        "cyl",
        "cyls",
        "cylinder",
        "cylinders",
        "cylinderconfig",
        "engineconfig",
        "config",
        "缸型",
        "缸数",
        "气缸数",
        "发动机配置",
        "气缸配置",
    },
    "Engine": {
        "engine",
        "engines",
        "enginesize",
        "enginedisplacement",
        "displacement",
        "liter",
        "litre",
        "排量",
        "排气量",
        "引擎",
        "发动机",
    },
}

# 太短的别名只允许全等匹配，避免子串误伤（如 "l" 命中 "Model"）
MIN_SUBSTRING_LEN = 3

_SPACES = re.compile(r"[\s 　]+")
_DISPLACEMENT = re.compile(
    r"(\d+(?:\.\d+)?)\s*(?:l\b|liter|litre|liters|litres|升)", re.IGNORECASE
)


def norm(value) -> str:
    """表头归一化：NFKC + 小写 + 去掉空格和常见分隔符。"""
    value = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"[\s_\-/（）() 　]+", "", value).lower()


_NORM_ALIASES = {f: {norm(a) for a in aliases} for f, aliases in HEADER_ALIASES.items()}


def match_field(header) -> str | None:
    key = norm(header)
    if not key:
        return None
    for field in FIELDS:
        if key in _NORM_ALIASES[field]:
            return field
    for field in FIELDS:
        for alias in _NORM_ALIASES[field]:
            if len(alias) >= MIN_SUBSTRING_LEN and alias in key:
                return field
    return None


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return _SPACES.sub(" ", str(value)).strip()


def clean_engine(value: str, warnings: list[str], where: str) -> str:
    """'5.7L V8 HEMI' -> '5.7L'；找不到排量则原样返回并告警。"""
    if not value:
        return value
    match = _DISPLACEMENT.search(value)
    if not match:
        warnings.append(f"{where}: engine 值 {value!r} 里找不到排量（如 5.7L），已原样保留，请人工确认。")
        return value
    cleaned = f"{match.group(1)}L"
    if norm(cleaned) != norm(value):
        warnings.append(f"{where}: engine {value!r} -> {cleaned!r}")
    return cleaned


# --------------------------------------------------------------------------
# 读原始二维表
# --------------------------------------------------------------------------
def read_xlsx(path: Path, sheet: str | None) -> tuple[list[list[str]], str]:
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise SystemExit(f"error: 打不开 {path}：{exc}")
    try:
        if sheet:
            if sheet not in wb.sheetnames:
                raise SystemExit(
                    f"error: {path} 里没有名为 {sheet!r} 的 sheet。"
                    f"可用的 sheet：{', '.join(wb.sheetnames)}"
                )
            ws = wb[sheet]
        else:
            ws = wb[wb.sheetnames[0]]
        rows = [[cell_text(c) for c in row] for row in ws.iter_rows(values_only=True)]
        return rows, ws.title
    finally:
        wb.close()


def read_delimited(path: Path) -> tuple[list[list[str]], str]:
    raw = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "cp1252"):
        try:
            raw = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        raise SystemExit(f"error: 无法确定 {path} 的文本编码（已尝试 utf-8 / gb18030 / cp1252）。")

    sample = raw[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
    except csv.Error:
        dialect = csv.excel_tab if path.suffix.lower() == ".tsv" else csv.excel

    rows = [[cell_text(c) for c in row] for row in csv.reader(raw.splitlines(), dialect)]
    return rows, "csv"


def load_table(path: Path, sheet: str | None) -> tuple[list[list[str]], str]:
    if not path.exists():
        raise SystemExit(f"error: 找不到输入文件：{path}")
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return read_xlsx(path, sheet)
    if suffix in (".csv", ".tsv", ".txt"):
        if sheet:
            print("warning: --sheet 对分隔符文件无效，已忽略。", file=sys.stderr)
        return read_delimited(path)
    if suffix == ".xls":
        raise SystemExit(
            "error: 不支持旧版 .xls 二进制格式。请先在 Excel 里另存为 .xlsx 再试。"
        )
    raise SystemExit(f"error: 不支持的文件类型 {suffix!r}。支持 .xlsx / .xlsm / .csv / .tsv。")


# --------------------------------------------------------------------------
# 表头识别
# --------------------------------------------------------------------------
def score_header_row(row: list[str]) -> int:
    return len({f for f in (match_field(c) for c in row) if f})


def find_header(rows: list[list[str]]) -> int:
    best_index, best_score = -1, 0
    for index, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        score = score_header_row(row)
        if score > best_score:
            best_index, best_score = index, score
    if best_score >= 3:
        return best_index
    if best_score == 2:
        print(
            f"warning: 只在第 {best_index + 1} 行认出 {best_score} 个列名，"
            f"用 --map 明确指定更稳妥。",
            file=sys.stderr,
        )
        return best_index
    return -1


def map_columns(header: list[str], overrides: dict[str, str], warnings: list[str]) -> dict[str, int]:
    columns: dict[str, int] = {}
    for field, name in overrides.items():
        key = norm(name)
        for index, cell in enumerate(header):
            if norm(cell) == key:
                columns[field] = index
                break
        else:
            raise SystemExit(
                f"error: --map {field}={name} 在表头里找不到这一列。"
                f"表头是：{', '.join(repr(c) for c in header if c)}"
            )

    for index, cell in enumerate(header):
        field = match_field(cell)
        if field is None or field in columns:
            continue
        columns[field] = index

    for field in FIELDS:
        if field not in columns:
            warnings.append(f"表头里没有识别出 {field} 列，相关区块会被跳过。")
    return columns


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="read_fitment.py",
        description="读取汽车适配表，输出 build_combinations.py 能吃的 fitment JSON。",
        epilog=(
            "示例:\n"
            "  read_fitment.py fitment.xlsx > job-fitment.json\n"
            "  read_fitment.py fitment.xlsx --sheet Fitment --forward-fill Make,Model\n"
            "  read_fitment.py table.csv --map Year='Model Year',Engine=Size\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="xlsx / xlsm / csv / tsv 文件")
    parser.add_argument("--sheet", help="xlsx 的 sheet 名（默认第一个）")
    parser.add_argument("--header-row", type=int, metavar="N", help="表头所在行号，从 1 开始（默认自动识别）")
    parser.add_argument(
        "--map",
        metavar="FIELD=COLUMN",
        action="append",
        default=[],
        help="手动指定列（可重复），如 --map Engine=Size",
    )
    parser.add_argument(
        "--forward-fill",
        metavar="FIELDS",
        default="",
        help="对空白单元格向下填充，如 Make,Model（合并单元格常见）",
    )
    parser.add_argument(
        "--no-clean-engine",
        action="store_true",
        help="不要清洗 engine 列（默认会把 '5.7L V8' 归一成 '5.7L'）",
    )
    parser.add_argument("-o", "--output", type=Path, help="写入文件而不是 stdout")
    return parser


def main(argv=None) -> int:
    _force_utf8_when_not_tty()
    args = build_parser().parse_args(argv)

    overrides: dict[str, str] = {}
    for item in args.map:
        if "=" not in item:
            print(f"error: --map 需要 FIELD=COLUMN 形式，收到 {item!r}。", file=sys.stderr)
            return EXIT_INPUT
        field, column = item.split("=", 1)
        field = field.strip().capitalize()
        if field not in FIELDS:
            print(f"error: --map 的字段必须是 {', '.join(FIELDS)}，收到 {field!r}。", file=sys.stderr)
            return EXIT_INPUT
        overrides[field] = column.strip()

    forward = [f.strip().capitalize() for f in args.forward_fill.split(",") if f.strip()]
    for field in forward:
        if field not in FIELDS:
            print(f"error: --forward-fill 的字段必须是 {', '.join(FIELDS)}，收到 {field!r}。", file=sys.stderr)
            return EXIT_INPUT

    warnings: list[str] = []
    rows, sheet_name = load_table(args.input, args.sheet)
    if not rows:
        print(f"error: {args.input} 是空的。", file=sys.stderr)
        return EXIT_INPUT

    header_index = args.header_row - 1 if args.header_row else find_header(rows)
    if header_index < 0:
        print(
            "error: 在前 10 行里找不到表头。请用 --header-row N 指定表头行号，"
            "或用 --map Year=列名 手动指定列。",
            file=sys.stderr,
        )
        return EXIT_INPUT
    if header_index >= len(rows):
        print(f"error: --header-row {args.header_row} 超出了表格范围（共 {len(rows)} 行）。", file=sys.stderr)
        return EXIT_INPUT

    header = rows[header_index]
    columns = map_columns(header, overrides, warnings)
    if "Model" not in columns and "Year" not in columns:
        print(
            f"error: 表头里既没有 Model 也没有 Year 列，无法继续。"
            f"表头是：{', '.join(repr(c) for c in header if c)}。用 --map 手动指定。",
            file=sys.stderr,
        )
        return EXIT_INPUT

    fitment: list[dict] = []
    blank_rows = 0
    fill_cache: dict[str, str] = {}
    gap_counts: dict[str, int] = {}

    for offset, row in enumerate(rows[header_index + 1:], start=header_index + 2):
        values = {f: (row[columns[f]] if f in columns and columns[f] < len(row) else "") for f in FIELDS}

        if not any(values.values()):
            blank_rows += 1
            continue

        for field in forward:
            if values[field]:
                fill_cache[field] = values[field]
            elif field in fill_cache:
                values[field] = fill_cache[field]

        # 没要求 forward-fill 的字段：只统计"上一行有值、这一行空"的情况，
        # 这是合并单元格的典型特征，最后统一提示用户加 --forward-fill。
        for field in FIELDS:
            if field in forward:
                continue
            if values[field]:
                fill_cache[field] = values[field]
            elif fill_cache.get(field):
                gap_counts[field] = gap_counts.get(field, 0) + 1

        if not args.no_clean_engine and values["Engine"]:
            values["Engine"] = clean_engine(values["Engine"], warnings, f"第 {offset} 行")

        fitment.append(values)

    if gap_counts:
        fields = ", ".join(sorted(gap_counts))
        warnings.append(
            f"有 {sum(gap_counts.values())} 处单元格为空但上一行有值（{fields}），"
            f"像是合并单元格。确认是合并的话，加 --forward-fill {fields} 重跑。"
        )

    if not fitment:
        print(f"error: 表头行之后没有读到任何数据行。", file=sys.stderr)
        return EXIT_INPUT

    result = {
        "source": str(args.input),
        "sheet": sheet_name,
        "header_row": header_index + 1,
        "column_map": {f: header[columns[f]] if f in columns else None for f in FIELDS},
        "fitment": fitment,
        "rows_read": len(fitment) + blank_rows,
        "rows_skipped_blank": blank_rows,
        "warnings": warnings,
    }

    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.write_text(payload, encoding="utf-8")
        print(f"已写出 {args.output}（{len(fitment)} 行）", file=sys.stderr)
    else:
        sys.stdout.write(payload)

    for warning in warnings[:20]:
        print(f"warning: {warning}", file=sys.stderr)
    if len(warnings) > 20:
        print(f"warning: 另有 {len(warnings) - 20} 条警告已省略。", file=sys.stderr)
    return EXIT_OK


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except OSError as exc:
        print(f"error: 文件读写失败：{exc}", file=sys.stderr)
        sys.exit(EXIT_IO)
