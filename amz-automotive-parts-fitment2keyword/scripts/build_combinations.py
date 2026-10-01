#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# dependencies = ["openpyxl>=3.1"]
# ///
"""汽车配件关键词组合生成器：fitment job JSON -> 3-sheet Excel。

Usage:
  uv run scripts/build_combinations.py job.json --output keyword-combinations.xlsx
  uv run scripts/build_combinations.py job.json -o out.xlsx --blocks-file blocks.txt --preview 0

输入 job.json:
  {
    "fitment":  [{"year": "2009-2022", "make": "Dodge", "model": "Challenger",
                  "cyl": "L4", "engine": "5.7L"}],
    "keywords": ["oil filter", "brake pad"],
    "oe":       ["15400-PLM-A02"],
    "problems": ["oil leak"]
  }
  year 支持单年 "2009"、闭区间 "2009-2022"、多段 "2015, 2017-2019"；区间在本脚本内展开成逐年行。
  字段名不分大小写，并支持中英文别名（年份 / 品牌 / 车型 / 缸型 / 排量）。
  cyl（缸型，如 L4 / V6）是可选维度：整列没有就自动跳过所有含它的区块，区块数从 39 回到 22。

输出:
  stdout   JSON 统计摘要（结构化，可直接喂给 jq）
  stderr   进度、警告、区块预览
  --output 3 个 sheet 的 xlsx（格式规格见 references/output-format.md）
  --blocks-file  22 个区块的纯文本版本（普通版 + 【+号限定符版本】）

退出码:
  0 成功   2 参数或输入不合法   3 文件读写失败
"""

from __future__ import annotations

import argparse
import itertools
import json
import re
import sys
import unicodedata
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit(
        "error: 缺少 openpyxl。请用 `uv run scripts/build_combinations.py ...`，"
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


class JobError(Exception):
    """输入不合法 —— 退出码 2。"""


# --------------------------------------------------------------------------
# 区块定义（唯一事实来源）
# --------------------------------------------------------------------------
# DIMENSIONS 的顺序 = 组合里 token 的顺序，也决定 Data Format 标签。
# 加一个维度（如 Cyl 缸型）会让区块数按 2^n-1 增长：
#   4 维 -> 4 + 15 + 3 = 22 个区块
#   5 维 -> 5 + 31 + 3 = 39 个区块

DIMENSIONS = ("Year", "Make", "Model", "Cyl", "Engine")

PURE_SECTIONS = DIMENSIONS

# 每个维度子集各出一个 Keyword 区块，按 itertools.combinations 的字典序展开：
# 1 维 -> 5 组，2 维 -> 10 组，3 维 -> 10 组，4 维 -> 5 组，5 维 -> 1 组，共 31 组。
# （退回 4 维时这里是 4+6+4+1 = 15 组，区块总数回到 22。）
KEYWORD_DIM_SECTIONS = tuple(
    combo
    for size in range(1, len(DIMENSIONS) + 1)
    for combo in itertools.combinations(DIMENSIONS, size)
)

TAIL_SECTIONS = ("OE+Keyword", "Problem+Keyword", "Problem+OE")

SECTION_FORMATS = (
    list(PURE_SECTIONS)
    + ["+".join(dims) + "+Keyword" for dims in KEYWORD_DIM_SECTIONS]
    + list(TAIL_SECTIONS)
)

# 改 DIMENSIONS 或上面的表时这个断言会立刻炸掉，防止区块悄悄漂移。
assert len(SECTION_FORMATS) == len(DIMENSIONS) + (2 ** len(DIMENSIONS) - 1) + len(TAIL_SECTIONS), (
    len(SECTION_FORMATS)
)

DIM_ZH = {
    "Year": "年份",
    "Make": "品牌",
    "Model": "车型",
    "Cyl": "缸型",
    "Engine": "排量",
}


def zh_label(fmt: str) -> str:
    """'Year+Make+Keyword' -> '年份 + 品牌 + Keyword'，用于文本区块标题。"""
    if fmt == "OE+Keyword":
        return "OE + Keyword"
    if fmt == "Problem+Keyword":
        return "Problem + 词根关键词"
    if fmt == "Problem+OE":
        return "Problem + OE"
    parts = fmt.split("+")
    if parts[-1] == "Keyword":
        dims, tail = parts[:-1], ["Keyword"]
    else:
        dims, tail = parts, []
    return " + ".join([DIM_ZH.get(d, d) for d in dims] + tail)


# --------------------------------------------------------------------------
# 样式常量（规格见 references/output-format.md）
# --------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
BAND_FILL = PatternFill("solid", fgColor="DCE6F1")
VARIANT_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_FONT = Font(color="FFFFFF", bold=True)

_EDGE = Side(style="thin", color="BFBFBF")
CELL_BORDER = Border(left=_EDGE, right=_EDGE, top=_EDGE, bottom=_EDGE)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")

MAX_COL_WIDTH = 100


# --------------------------------------------------------------------------
# 年份展开
# --------------------------------------------------------------------------
_YEAR_ATOM = re.compile(r"^(\d{4})(?:-(\d{4}))?$")
_YEAR_OPEN = re.compile(r"^(\d{4})-(present|current|now|至今|现在|今)$", re.IGNORECASE)
_RANGE_WARN_THRESHOLD = 40


def expand_year_field(raw) -> tuple[list[str], list[str]]:
    """展开年份字段，返回 (逐年列表, 警告列表)。

    '2009'          -> ['2009']
    '2009-2022'     -> ['2009' ... '2022']
    '2015, 2017-2019' -> ['2015', '2017', '2018', '2019']
    """
    text = str(raw if raw is not None else "").strip()
    if not text:
        return [], []

    # 只在 year 字段上收敛连字符两侧空白（"2009 - 2022" -> "2009-2022"），
    # 这样按空白切分才不会把区间拆散。
    text = re.sub(r"\s*[-–—~]\s*", "-", text)
    atoms = [a for a in re.split(r"[,，;；/、\s]+", text) if a]

    years: list[str] = []
    warnings: list[str] = []
    for atom in atoms:
        if _YEAR_OPEN.match(atom):
            raise JobError(
                f'year 用了开放区间 "{atom}"，无法确定结束年份。'
                f'请写成明确区间（如 "2015-2024"），或只保留起始年（如 "2015"）。'
            )
        m = _YEAR_ATOM.match(atom)
        if not m:
            raise JobError(
                f'无法解析 year 值 "{atom}"（原始值 "{text}"）。'
                f'支持：2009 / 2009-2022 / "2015, 2017-2019"。'
            )
        start = int(m.group(1))
        end = int(m.group(2) or start)
        if end < start:
            raise JobError(f'year 区间 "{atom}" 的结束年小于起始年。')
        if end - start > _RANGE_WARN_THRESHOLD:
            warnings.append(
                f'year 区间 "{atom}" 展开成 {end - start + 1} 年，请确认不是笔误。'
            )
        for y in range(start, end + 1):
            if not 1900 <= y <= 2100:
                raise JobError(f"year 值 {y} 超出合理范围（1900-2100）。")
            if str(y) not in years:
                years.append(str(y))
    return years, warnings


# --------------------------------------------------------------------------
# 读取 job.json
# --------------------------------------------------------------------------
def _norm_key(value) -> str:
    value = unicodedata.normalize("NFKC", str(value)).lower()
    return re.sub(r"[\s_\-/（）()]+", "", value)


_FIELD_ALIASES = {
    "Year": {"year", "years", "年份", "年款", "年度", "年", "modelyear", "my"},
    "Make": {"make", "makes", "brand", "manufacturer", "品牌", "厂商", "制造商", "厂牌"},
    "Model": {"model", "models", "车型", "车系", "型号"},
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
        "排量",
        "排气量",
        "引擎",
        "发动机",
    },
}

_JOB_ALIASES = {
    "keywords": ("keywords", "keyword", "kw", "词根关键词", "词根", "关键词"),
    "oe": ("oe", "oes", "oenumber", "oenumbers", "oem", "partnumber", "原厂件号", "oe号"),
    "problems": ("problems", "problem", "故障语义", "故障", "症状"),
}


def _pick_field(entry: dict, field: str) -> str:
    for key, value in entry.items():
        if _norm_key(key) in _FIELD_ALIASES[field]:
            return "" if value is None else str(value).strip()
    return ""


def _pick_list(job: dict, name: str) -> list[str]:
    for key, value in job.items():
        if _norm_key(key) in _JOB_ALIASES[name]:
            if value is None:
                return []
            if isinstance(value, str):
                return [value.strip()] if value.strip() else []
            if isinstance(value, list):
                return [str(v).strip() for v in value if str(v).strip()]
            raise JobError(f'"{key}" 必须是字符串数组，收到 {type(value).__name__}。')
    return []


def load_job(path: Path, split_dims: bool) -> dict:
    try:
        job = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise JobError(f"找不到 job 文件：{path}")
    except json.JSONDecodeError as exc:
        raise JobError(f"{path} 不是合法 JSON：{exc}")
    if not isinstance(job, dict):
        raise JobError(f"{path} 顶层必须是 JSON 对象。")

    raw_rows = None
    for key, value in job.items():
        if _norm_key(key) in {"fitment", "fitments", "适配", "适配数据", "适配表"}:
            raw_rows = value
            break
    if raw_rows is None:
        raise JobError(
            f'{path} 缺少 "fitment" 数组。期望格式：\n'
            '  {"fitment": [{"year": "2009-2022", "make": "Dodge", '
            '"model": "Challenger", "engine": "5.7L"}], "keywords": ["oil filter"]}'
        )
    if not isinstance(raw_rows, list) or not raw_rows:
        raise JobError('"fitment" 必须是至少含一项的数组。')

    fitment: list[dict] = []
    seen: set[tuple[str, str, str, str]] = set()
    warnings: list[str] = []

    for idx, entry in enumerate(raw_rows, 1):
        if not isinstance(entry, dict):
            raise JobError(f"fitment[#{idx}] 必须是对象，收到 {type(entry).__name__}。")
        values = {f: _pick_field(entry, f) for f in DIMENSIONS}
        if not any(values.values()):
            warnings.append(f"fitment[#{idx}] 所有字段都为空，已跳过。")
            continue

        years, year_warnings = expand_year_field(values["Year"])
        warnings.extend(f"fitment[#{idx}] {w}" for w in year_warnings)

        if not years:
            base = dict(values, Year="")
            key = tuple(base[f] for f in DIMENSIONS)
            if key not in seen:
                seen.add(key)
                fitment.append(base)
            warnings.append(f"fitment[#{idx}] 没有 year，只参与不含年份的区块。")
            continue

        for year in years:
            row = dict(values, Year=year)
            key = tuple(row[f] for f in DIMENSIONS)
            if key in seen:
                continue
            seen.add(key)
            fitment.append(row)

    if not fitment:
        raise JobError("fitment 展开后没有任何有效行。")

    return {
        "fitment": fitment,
        "keywords": _pick_list(job, "keywords"),
        "oe": _pick_list(job, "oe"),
        "problems": _pick_list(job, "problems"),
        "split_dims": split_dims,
        "warnings": warnings,
    }


# --------------------------------------------------------------------------
# 组合生成
# --------------------------------------------------------------------------
def _dim_tokens(row: dict, dims, split_dims: bool) -> list[str] | None:
    """把若干维度的值拼成 token 列表；任一维度为空则返回 None（跳过该组合）。"""
    tokens: list[str] = []
    for dim in dims:
        value = (row.get(dim) or "").strip()
        if not value:
            return None
        tokens.extend(value.split() if split_dims else [value])
    return tokens


def _dedupe_and_variant(raw_rows: list[list[str]]) -> list[tuple[list[str], bool]]:
    """按普通版字符串去重，并在原始行后紧跟插入去连字符变体行。"""
    seen: set[str] = set()
    out: list[tuple[list[str], bool]] = []
    for tokens in raw_rows:
        normal = " ".join(tokens)
        if not normal or normal in seen:
            continue
        seen.add(normal)
        out.append((tokens, False))

        if any("-" in t for t in tokens):
            variant = [t.replace("-", "") for t in tokens]
            vnormal = " ".join(variant)
            if all(variant) and vnormal not in seen:
                seen.add(vnormal)
                out.append((variant, True))
    return out


def build_sections(job: dict) -> list[tuple[str, list[tuple[list[str], bool]]]]:
    fitment = job["fitment"]
    keywords = job["keywords"]
    oes = job["oe"]
    problems = job["problems"]
    split_dims = job["split_dims"]

    sections: list[tuple[str, list[tuple[list[str], bool]]]] = []

    for dim in PURE_SECTIONS:
        raw = [t for t in (_dim_tokens(r, (dim,), split_dims) for r in fitment) if t]
        sections.append((dim, _dedupe_and_variant(raw)))

    for dims in KEYWORD_DIM_SECTIONS:
        raw: list[list[str]] = []
        for row in fitment:
            base = _dim_tokens(row, dims, split_dims)
            if base is None:
                continue
            for kw in keywords:
                words = kw.split()
                if words:
                    raw.append(base + words)
        sections.append(("+".join(dims) + "+Keyword", _dedupe_and_variant(raw)))

    raw_oe_kw: list[list[str]] = []
    for oe in oes:
        for kw in keywords:
            words = kw.split()
            if words:
                raw_oe_kw.append([oe] + words)
    sections.append(("OE+Keyword", _dedupe_and_variant(raw_oe_kw)))

    raw_prob_kw: list[list[str]] = []
    for problem in problems:
        pwords = problem.split()
        if not pwords:
            continue
        for kw in keywords:
            words = kw.split()
            if words:
                raw_prob_kw.append(pwords + words)
    sections.append(("Problem+Keyword", _dedupe_and_variant(raw_prob_kw)))

    raw_prob_oe: list[list[str]] = []
    for problem in problems:
        pwords = problem.split()
        if not pwords:
            continue
        for oe in oes:
            raw_prob_oe.append(pwords + [oe])
    sections.append(("Problem+OE", _dedupe_and_variant(raw_prob_oe)))

    assert [fmt for fmt, _ in sections] == SECTION_FORMATS
    return sections


def render_normal(tokens) -> str:
    return " ".join(tokens)


def render_plus(tokens) -> str:
    return " ".join("+" + t for t in tokens)


# --------------------------------------------------------------------------
# Excel 输出
# --------------------------------------------------------------------------
def _display_width(value) -> int:
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(value))


def _autofit(ws, n_cols: int, min_width: int = 10) -> None:
    for col in range(1, n_cols + 1):
        letter = get_column_letter(col)
        widest = max(
            (_display_width(ws.cell(row=r, column=col).value or "") for r in range(1, ws.max_row + 1)),
            default=0,
        )
        ws.column_dimensions[letter].width = max(min_width, min(widest + 3, MAX_COL_WIDTH))


def _style_header(ws, headers: list[str]) -> None:
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = CELL_BORDER
    ws.freeze_panes = "A2"


def write_excel(path: Path, sections, fitment, stats, sheet_names) -> None:
    wb = Workbook()

    # ---- Sheet 1: 关键词组合结果 ----
    ws = wb.active
    ws.title = sheet_names[0]
    _style_header(ws, ["Data Format", "普通版本", "+号限定符版本"])

    band = 0
    for fmt, rows in sections:
        for tokens, is_variant in rows:
            ws.append([fmt, render_normal(tokens), render_plus(tokens)])
            row_idx = ws.max_row
            if is_variant:
                fill = VARIANT_FILL
            else:
                # 蓝白交替只对原始行计数，变体行固定黄色，交替看起来才连续
                fill = BAND_FILL if band % 2 else None
                band += 1
            for col in range(1, 4):
                cell = ws.cell(row=row_idx, column=col)
                cell.border = CELL_BORDER
                cell.alignment = LEFT
                if fill is not None:
                    cell.fill = fill
    _autofit(ws, 3)

    # ---- Sheet 2: 测试摘要 ----
    ws2 = wb.create_sheet(sheet_names[1])
    _style_header(ws2, ["项目", "值"])
    summary = [
        ("Total Rows", stats["total_rows"]),
        ("Sections", stats["sections_emitted"]),
        ("Original Rows", stats["original_rows"]),
        ("Hyphen Variants Added", stats["hyphen_variants"]),
        ("Total Output Rows", stats["total_output_rows"]),
    ]
    for label, value in summary:
        ws2.append([label, value])
    ws2.append([])
    for label, value in stats["diagnostics"]:
        ws2.append([label, value])
    for row in range(2, ws2.max_row + 1):
        for col in (1, 2):
            cell = ws2.cell(row=row, column=col)
            cell.border = CELL_BORDER
            cell.alignment = LEFT
            if row % 2 == 0 and cell.value is not None:
                cell.fill = BAND_FILL
    _autofit(ws2, 2)

    # ---- Sheet 3: 原始适配数据 ----
    # 列 = DIMENSIONS，但整列全空的维度不输出（没有缸型数据时就还是原来的四列）
    cols = [d for d in DIMENSIONS if any((row.get(d) or "").strip() for row in fitment)]
    ws3 = wb.create_sheet(sheet_names[2])
    _style_header(ws3, cols)
    for row in fitment:
        ws3.append([row[d] for d in cols])
    for row in range(2, ws3.max_row + 1):
        for col in range(1, len(cols) + 1):
            cell = ws3.cell(row=row, column=col)
            cell.border = CELL_BORDER
            cell.alignment = LEFT
            if row % 2 == 0:
                cell.fill = BAND_FILL
    _autofit(ws3, len(cols))

    try:
        wb.save(path)
    except PermissionError:
        raise SystemExit(
            f"error: 写不了 {path} —— 文件被占用。\n"
            f"       最常见的原因是这个文件正在 Excel / WPS 里打开。"
            f"关掉它再跑一次，或者用 -o 换个文件名。"
        )
    except OSError as exc:
        raise SystemExit(f"error: 写入 {path} 失败：{exc}")


# --------------------------------------------------------------------------
# 文本区块输出
# --------------------------------------------------------------------------
def write_blocks(path: Path, sections) -> None:
    lines: list[str] = []
    for variant_name, use_plus in (("", False), ("【+号限定符版本】", True)):
        if variant_name:
            lines.append("")
            lines.append(f"# {variant_name}")
            lines.append("")
        for index, (fmt, rows) in enumerate(sections, 1):
            lines.append("")
            lines.append(f"## 区块 {index}：{zh_label(fmt)}")
            lines.append("")
            for tokens, _ in rows:
                lines.append(render_plus(tokens) if use_plus else render_normal(tokens))
    try:
        path.write_text("\n".join(lines).lstrip("\n") + "\n", encoding="utf-8")
    except OSError as exc:
        raise SystemExit(f"error: 写入 {path} 失败：{exc}")


def preview(sections, limit: int, stream=sys.stderr) -> None:
    if limit < 0:
        return
    print("", file=stream)
    print("---- 区块预览 ----", file=stream)
    for index, (fmt, rows) in enumerate(sections, 1):
        if not rows:
            continue
        print(f"\n区块 {index}：{zh_label(fmt)}   [{fmt}]   共 {len(rows)} 行", file=stream)
        shown = rows if limit == 0 else rows[:limit]
        for tokens, is_variant in shown:
            mark = "  (变体)" if is_variant else ""
            print(f"   {render_normal(tokens)}    |    {render_plus(tokens)}{mark}", file=stream)
        if limit and len(rows) > limit:
            print(f"   ... 其余 {len(rows) - limit} 行见 Excel", file=stream)


def print_sections(stream=sys.stdout) -> None:
    print(
        f"共 {len(SECTION_FORMATS)} 个区块，{len(DIMENSIONS)} 个维度：{' / '.join(DIMENSIONS)}",
        file=stream,
    )
    print(
        f"（每加一个维度，区块数 = n + 2^n - 1 + {len(TAIL_SECTIONS)}；"
        f"4 维 = 22 个，5 维 = {len(SECTION_FORMATS)} 个）",
        file=stream,
    )
    print(file=stream)
    for index, fmt in enumerate(SECTION_FORMATS, 1):
        dep = "补充信息" if fmt in TAIL_SECTIONS else "fitment"
        print(f"{index:>3}  {fmt:<44} {zh_label(fmt):<30} {dep}", file=stream)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="build_combinations.py",
        description="从 fitment job JSON 生成 22 个区块的关键词组合，输出 3-sheet Excel。",
        epilog=(
            "示例:\n"
            "  build_combinations.py job.json -o out.xlsx\n"
            "  build_combinations.py job.json -o out.xlsx --blocks-file blocks.txt --preview 0\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("job", type=Path, nargs="?", help="job JSON 文件路径")
    parser.add_argument("-o", "--output", type=Path, help="输出的 xlsx 路径")
    parser.add_argument(
        "--list-sections",
        action="store_true",
        help="只打印区块清单（编号 / Data Format / 中文标题 / 依赖）然后退出，不需要 job 文件",
    )
    parser.add_argument(
        "--blocks-file",
        type=Path,
        help="同时写出 22 个区块的纯文本版本（含 +号限定符版本）",
    )
    parser.add_argument(
        "--preview",
        type=int,
        default=3,
        metavar="N",
        help="向 stderr 打印每个区块的前 N 行（默认 3；0=全部；-1=不打印）",
    )
    parser.add_argument(
        "--split-dims",
        action="store_true",
        help="多维度的值（如 Grand Cherokee）也逐词加 +，默认整体作为一个 token",
    )
    parser.add_argument("--sheet1-name", default="关键词组合结果", help="第 1 个 sheet 名")
    parser.add_argument("--sheet2-name", default="测试摘要", help="第 2 个 sheet 名")
    parser.add_argument("--sheet3-name", default="原始适配数据", help="第 3 个 sheet 名")
    return parser


def main(argv=None) -> int:
    _force_utf8_when_not_tty()
    argv = list(sys.argv[1:] if argv is None else argv)

    # --list-sections 是自省用的，不需要 job / -o，所以走 argparse 之前先拦掉
    if "--list-sections" in argv:
        print_sections()
        return EXIT_OK

    args = build_parser().parse_args(argv)
    if args.job is None:
        print("error: 缺少 job JSON 路径。用法：build_combinations.py job.json -o out.xlsx", file=sys.stderr)
        return EXIT_INPUT
    if args.output is None:
        print("error: 缺少 -o/--output。用法：build_combinations.py job.json -o out.xlsx", file=sys.stderr)
        return EXIT_INPUT

    sheet_names = [args.sheet1_name, args.sheet2_name, args.sheet3_name]
    if len(set(sheet_names)) != 3:
        print("error: 三个 sheet 名不能重复。", file=sys.stderr)
        return EXIT_INPUT

    try:
        job = load_job(args.job, args.split_dims)
    except JobError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return EXIT_INPUT

    sections = build_sections(job)

    total_rows = sum(len(rows) for _, rows in sections)
    original_rows = sum(1 for _, rows in sections for _, is_v in rows if not is_v)
    variants = total_rows - original_rows
    emitted = sum(1 for _, rows in sections if rows)
    skipped = [fmt for fmt, rows in sections if not rows]

    stats = {
        "total_rows": original_rows,
        "original_rows": original_rows,
        "hyphen_variants": variants,
        "total_output_rows": total_rows,
        "sections_emitted": emitted,
        "diagnostics": [
            ("适配数据行数（year 展开后）", len(job["fitment"])),
            ("词根关键词数", len(job["keywords"])),
            ("OE 号数", len(job["oe"])),
            ("故障语义数", len(job["problems"])),
            ("跳过的区块", "、".join(skipped) if skipped else "无"),
        ],
    }

    write_excel(args.output, sections, job["fitment"], stats, sheet_names)
    if args.blocks_file:
        write_blocks(args.blocks_file, sections)
    preview(sections, args.preview)

    for warning in job["warnings"]:
        print(f"warning: {warning}", file=sys.stderr)

    report = {
        "output": str(args.output),
        "blocks_file": str(args.blocks_file) if args.blocks_file else None,
        "fitment_rows": len(job["fitment"]),
        "keywords": job["keywords"],
        "oe": job["oe"],
        "problems": job["problems"],
        "sections": [
            {"format": fmt, "rows": len(rows)} for fmt, rows in sections
        ],
        "sections_emitted": emitted,
        "sections_skipped": skipped,
        "totals": {
            "total_rows": original_rows,
            "original_rows": original_rows,
            "hyphen_variants_added": variants,
            "total_output_rows": total_rows,
        },
        "warnings": job["warnings"],
    }
    json.dump(report, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
    return EXIT_OK


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except OSError as exc:
        print(f"error: 文件读写失败：{exc}", file=sys.stderr)
        sys.exit(EXIT_IO)
