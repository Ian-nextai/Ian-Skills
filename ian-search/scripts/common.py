"""Shared helpers for batch information enrichment.

No network calls and no provider-specific secrets live here.
"""

from __future__ import annotations

import html
import json
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence

DEFAULT_PRODUCT_TERMS: List[str] = [
    # Chinese examples; override with config["product_terms"] for a domain.
    "汽车", "配件", "零部件", "部件", "设备", "机械", "机器", "工具", "电子", "电器",
    "电机", "传感器", "泵", "阀", "轴承", "齿轮", "滤清", "过滤", "制动", "刹车",
    "离合", "轮胎", "轮毂", "灯具", "车灯", "橡胶", "塑料", "密封", "紧固件",
    "模具", "材料", "用品", "组件", "配件", "line", "system",
    # English examples.
    "parts", "components", "equipment", "machinery", "tools", "electronics",
    "electrical", "motor", "sensor", "pump", "valve", "bearing", "gear",
    "filter", "brake", "clutch", "tire", "tyre", "wheel", "lamp", "light",
    "rubber", "plastic", "seal", "fastener", "mould", "mold", "accessor",
]

DEFAULT_GENERIC_TERMS = {
    "汽车配件", "汽车零部件", "汽车部件", "汽车用品", "汽车用品与改装",
    "汽车改装", "汽车相关产品", "汽车配件/零部件", "零部件及组件",
    "auto parts", "automotive parts", "car accessories", "spare parts",
}

BAD_RE = re.compile(
    r"公司|地址|電話|电话|郵箱|邮箱|成立|注册|註冊|简介|簡介|招聘|认证|認證|纳税人|"
    r"新闻|新聞|媒体|媒體|协会|協會|商会|商會|展览|展覽|展会|展會|政府|"
    r"批发$|批發$|零售$|销售$|銷售$|贸易$|貿易$|"
    r"我们|我們|服务|服務|品牌|定位|价格|價格|报价|報價|专利|專利|"
    r"经营范围|經營範圍|一般项目|一般項目|技术开发|技術開發|技术咨询|技術諮詢|"
    r"技术转让|技術轉讓|进出口|進出口|货物进出口|貨物進出口|"
    r"\[|【|首页|联系我们|视频中心|加入我们|立即沟通|进入店铺",
    re.I,
)

PUNCT_RE = re.compile(
    r"[\s　\-—_·•,，.。;；:：()（）\[\]【】{}“”\"'‘’/\\|&＆+＋]+"
)

CN_SUFFIX_RE = re.compile(r"(股份有限公司|有限责任公司|有限公司|集团有限公司|集团|公司)$")
EN_SUFFIX_RE = re.compile(
    r"(coltd|co\.?ltd\.?|ltd\.?|limited|inc\.?|llc|gmbh|corp\.?|corporation|company|co\.?)$"
)
LEAD_REGION_RE = re.compile(r"^(.{2,4}?[省市])")


def has(value: Any) -> bool:
    """Return True for a meaningful cell value."""
    if value is None:
        return False
    if isinstance(value, float):
        return value == value  # not NaN
    text = str(value).strip()
    return bool(text) and text.lower() not in {"nan", "none", "null", "n/a", "-"}


def clean_text(value: Any) -> str:
    """Strip HTML and normalize whitespace without changing meaning."""
    text = html.unescape(str(value or ""))
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text.replace(" ", " ")).strip()


def norm(value: Any) -> str:
    """Canonical key for matching: NFKC, lowercase, punctuation removed."""
    return PUNCT_RE.sub("", unicodedata.normalize("NFKC", str(value or "")).lower())


def name_key(value: Any, strip_suffix: bool = True) -> str:
    """Canonical company-name key, optionally stripping legal suffixes."""
    key = norm(value)
    if not strip_suffix:
        return key
    key = CN_SUFFIX_RE.sub("", key)
    key = EN_SUFFIX_RE.sub("", key)
    key = LEAD_REGION_RE.sub(lambda m: m.group(1)[:-1], key)
    return key


def is_cn(value: Any) -> bool:
    return bool(re.search(r"[一-鿿]", str(value or "")))


def get_field(row: Dict[str, Any], field: str) -> Any:
    """Read a dotted field path from a row dict."""
    cur: Any = row
    for part in field.split("."):
        if not isinstance(cur, dict):
            return ""
        cur = cur.get(part, "")
    return cur


def fill_template(template: str, row: Dict[str, Any], name: str = "", name_field: str = "name") -> str:
    """Fill {name} and {field} placeholders in a query template."""
    def repl(match: re.Match[str]) -> str:
        key = match.group(1)
        if key == "name":
            return name
        value = get_field(row, key)
        return str(value) if has(value) else ""

    return re.sub(r"\{([^{}]+)\}", repl, str(template or "")).strip()


def load_json(path: str | Path, default: Any = None) -> Any:
    p = Path(path)
    if not p.exists():
        return default
    with p.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def dump_json(path: str | Path, obj: Any) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2)


def _dedupe_tokens(text: str) -> str:
    tokens = [t.strip() for t in re.split(r"[、;；]", text) if t.strip()]
    seen: set[str] = set()
    out: List[str] = []
    for token in tokens:
        key = token.lower()
        if key not in seen:
            seen.add(key)
            out.append(token)
    return "、".join(out)


def clean_phrase(value: Any, max_len: int = 80) -> str:
    """Remove common boilerplate around an extracted product phrase."""
    text = clean_text(value)
    text = re.sub(r"\s+", "", text)
    text = re.split(r"[.…·。；;]", text)[0]
    text = re.split(r"[（(\[]", text)[0]
    text = re.split(r"的企业|等产品的企业|等產品的企業", text)[0]
    text = re.sub(
        r"^(经营范围(?:包括|包含)?|主营业务|主要经营范围|主要经营|主营产品|主营|主要产品|产品|产品|"
        r"致力于|专业研发、生产和销售|研发、生产和销售|生产和销售|专业生产|生产销售|制造销售|"
        r"公司主要从事|speciali[sz]\w*\s+in|main products?(?:\s+are|\s+include)?|products?\s+include)[:：]?",
        "",
        text,
        flags=re.I,
    )
    text = re.sub(r"^[是为：:，,\s]+", "", text)
    text = re.sub(r"(等产品|等產品|等)$", "", text)
    text = re.split(r"等业务|等業務|，业务|，業務", text)[0]
    text = text.strip(" ：:，,、;；。.-—_|*＊")
    text = text.replace(",", "、").replace("，", "、")
    return _dedupe_tokens(text)[:max_len]


def product_like(text: str, product_terms: Optional[Sequence[str]] = None) -> bool:
    terms = list(product_terms or DEFAULT_PRODUCT_TERMS)
    low = str(text or "").lower()
    return any(str(term).lower() in low for term in terms)


def accept_product(
    value: Any,
    product_terms: Optional[Sequence[str]] = None,
    generic_terms: Optional[Iterable[str]] = None,
    max_len: int = 80,
) -> str:
    """Return a cleaned product phrase, or '' when it looks like noise."""
    text = clean_phrase(value, max_len=max_len)
    if not (2 <= len(text) <= max_len):
        return ""
    if BAD_RE.search(text):
        return ""
    if not product_like(text, product_terms):
        return ""
    generic = {str(x).lower() for x in (generic_terms or DEFAULT_GENERIC_TERMS)}
    tokens = [t for t in re.split(r"[、;；]", text) if t]
    if tokens and all(t.lower() in generic or len(t) <= 2 for t in tokens):
        return ""
    return text


def extract_product(
    snippets: Iterable[str],
    name: str,
    product_terms: Optional[Sequence[str]] = None,
    generic_terms: Optional[Iterable[str]] = None,
    max_len: int = 80,
    window: int = 240,
) -> str:
    """Extract a product phrase from snippets that mention the target name."""
    key = name_key(name)
    compact_name = re.sub(r"\s+", "", str(name or ""))
    candidates: List[str] = []

    for snippet in snippets:
        text = clean_text(snippet)
        if key and key not in name_key(text):
            continue
        compact = re.sub(r"\s+", "", text)
        windows: List[str] = []
        if compact_name and compact_name in compact:
            windows.extend(
                compact[m.end():m.end() + window]
                for m in re.finditer(re.escape(compact_name), compact)
            )
        windows.append(text)
        for chunk in windows:
            for pattern in (
                r"(?:主营产品|主营|主要产品(?:有|包括|为)?|经营范围(?:包括|包含)?|"
                r"专业生产|生产销售|制造销售|主要从事)[:：]?\s*([^。]{2,120})",
                r"(?:speciali[sz]\w*\s+in|main products?\s+(?:are|include)|products?\s+include)\s*([^.;]{3,180})",
            ):
                match = re.search(pattern, chunk, re.I)
                if match:
                    candidates.append(match.group(1))

    best = ""
    for candidate in candidates:
        phrase = accept_product(
            candidate,
            product_terms=product_terms,
            generic_terms=generic_terms,
            max_len=max_len,
        )
        if len(phrase) > len(best):
            best = phrase
    return best
