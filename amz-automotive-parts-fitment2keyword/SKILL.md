---
name: amz-automotive-parts-fitment2keyword
description: >-
  汽车配件关键词组合生成器：把汽车适配表（fitment data）展开成电商关键词矩阵并输出 Excel。只要用户提供包含
  年份 Year、品牌 Make、车型 Model、排量 Engine、缸型 Cyl（L4/V6）的适配信息就用这个 skill —— 无论数据来自
  图片截图、粘贴的文本，还是 Excel / CSV 表格，也无论用户有没有说出"关键词组合""fitment"这些词。流程会提取
  适配数据、追问词根关键词 / OE 号 / Part Number / 故障语义，然后产出 39 个分类区块的全部组合（普通版、
  +号精确匹配版、去连字符变体）和 3 个 sheet 的 Excel 文件。触发场景：适配表、适配数据、fitment、
  Year/Make/Model/Engine、keyword 组合、关键词组合、OE 号、Part Number、亚马逊汽车配件 listing、ACES 数据、
  汽配关键词矩阵。
compatibility: >-
  需要 Python 3.9+。脚本用 PEP 723 声明了 openpyxl，用 uv run 执行会自动装；直接 python 执行需先 pip install openpyxl。
metadata:
  version: "2.0"
  format: agentskills.io specification
---

# 汽车配件关键词组合生成器

把 fitment 数据展开成电商关键词矩阵，交付一个 3-sheet 的 Excel。

**组合一律由脚本生成，你只负责提取数据和驱动脚本。** 22 个区块 × 逐年展开 × 每个词根关键词，
几十行输入会膨胀到上千行输出（参考量级：975 → 1316 行）。手写既慢又会漏区块、漏变体。

## 工作流

- [ ] **Step 1** 提取适配数据 → 只保留 Year / Make / Model / Cyl / Engine 五列
- [ ] **Step 2** 用**一条消息**问齐补充信息（词根关键词 / OE 号 / 故障语义）
- [ ] **Step 3** 写 `job.json`
- [ ] **Step 4** 跑 `scripts/build_combinations.py` 生成 Excel
- [ ] **Step 5** 汇报统计摘要 + 抽样给用户核对

---

## Step 1：提取适配数据

按输入来源分流：

| 来源 | 做法 |
| --- | --- |
| Excel / CSV / TSV 文件 | 跑 `scripts/read_fitment.py`，自动映射中英文表头 |
| 图片截图 | 用视觉能力读表，只取五列 |
| 粘贴的文本 / 聊天里的表格 | 直接解析成五列 |

五列内容照抄，**不要脑补**：某个字段缺失就留空，脚本会自动跳过含该字段的区块。
`Cyl`（缸型，如 `L4`/`V6`）和 `Engine`（排量，如 `2.5L`）是**两列**，别合并 ——
挤在一格里的 `5.7L V8 HEMI` 要拆成 `Cyl=V8` + `Engine=5.7L`。

图片里通常混着马力、驱动方式、车身类型、营销技术名（`Hemi`/`VTEC`）等干扰列，一行都不要带进来。

提取细则（干扰项清单、缸型写法、列名映射表、合并单元格、engine 清洗）见
[references/extraction.md](references/extraction.md)。

---

## Step 2：收集补充信息

提取完成后，用**一条消息**问齐三项（用户已在对话里给过就跳过，不要再问）：

```
数据已提取完成 ✅ 请补充以下信息，没有的直接回复"跳过"：

1. 词根关键词（如：oil filter, brake pad, spark plug）
2. OE 号码（原厂件号，如：15400-PLM-A02）
3. 故障语义（如：oil leak, misfire, hard to start）
```

用户回"没有" / "跳过" / "不需要"时，对应区块整块跳过 —— 这是正常路径，不要追问，也不要自己编。
三个都没有时脚本会只产出区块 1–19。

---

## Step 3：写 job.json

把提取结果落成一个文件，别把数据留在对话里：

```json
{
  "fitment": [
    {"year": "2021-2022", "make": "Hyundai", "model": "Santa Fe", "cyl": "L4", "engine": "2.5L"},
    {"year": "2015, 2017-2019", "make": "Ford", "model": "F-150", "cyl": "V8", "engine": "3.5L"}
  ],
  "keywords": ["oil filter", "brake pad"],
  "oe": ["15400-PLM-A02"],
  "problems": ["oil leak"]
}
```

- `year` 原样填区间字符串（`2009-2022`、`2015, 2017-2019`），**展开由脚本做**，不要手算成逐年数组。
- `cyl` 是可选维度：整列没有就自动跳过所有含它的区块，区块数从 39 回到 22。
- 字段名不分大小写，也认中文别名（年份 / 品牌 / 车型 / 缸型 / 排量；词根关键词 / OE 号 / 故障语义）。
- `keywords` / `oe` / `problems` 没有就给空数组或直接省略。

用 `read_fitment.py` 时它输出的 `fitment` 数组可以直接拷进来：

```bash
python scripts/read_fitment.py fitment.xlsx -o job-fitment.json
```

---

## Step 4：生成 Excel

```bash
python scripts/build_combinations.py job.json -o keyword-combinations.xlsx
```

需要把组合贴到别处时，加 `--blocks-file` 同时导出文本区块：

```bash
python scripts/build_combinations.py job.json -o keyword-combinations.xlsx --blocks-file blocks.txt
```

脚本会把 JSON 统计摘要打到 stdout、进度和预览打到 stderr。
`--preview N` 控制预览行数（默认每区块 3 行，`0` 全部，`-1` 关闭）。
`--list-sections` 打印区块清单（不需要 job 文件）。

脚本自检：`--help` 看全部参数，退出码 `0` 成功 / `2` 输入不合法 / `3` 文件读写失败。

---

## Step 5：汇报

**不要把上千行组合贴进对话。** Excel 就是交付物，聊天里给摘要：

1. 文件路径
2. 统计：原始行数 / 变体行数 / 总输出行数 / 覆盖的区块数
3. 每个区块各几行（脚本 stdout 的 JSON 里 `sections` 字段直接有）
4. 挑 2–3 个区块展示前几行，让用户核对口径
5. 脚本报的 `warning`（engine 清洗、年份区间异常、跳过的区块）原样转达，不要吞掉

---

## 区块总览

有 5 个维度 `Year / Make / Model / Cyl / Engine`，区块 = 纯维度 + 每个维度子集配 Keyword + 3 个补充信息区块：

```
区块数 = n + (2^n - 1) + 3        # n = 维度个数
5 维 -> 5 + 31 + 3 = 39 个区块
```

**不要凭记忆背 39 个区块名 —— 问脚本：**

```bash
python scripts/build_combinations.py --list-sections
```

它会打印编号、`Data Format`（Sheet 1 第一列的值）、中文标题和依赖，永远跟代码一致。
加维度只需改脚本里的 `DIMENSIONS` 一处，区块数由公式断言校验。

每个区块同时产出**普通版**和 **+号限定符版**（每个独立词前加 `+`）：
`2021 Hyundai Santa Fe L4 2.5L sensor` ↔ `+2021 +Hyundai +Santa Fe +L4 +2.5L +sensor`。

Sheet 结构、配色、连字符变体规则、统计口径见
[references/output-format.md](references/output-format.md) —— 只在需要核对产出或回答格式问题时读。

---

## Gotchas

- **别手算年份。** `2009-2022` 直接把字符串交给脚本。手算 14 个年份容易错一两年，而且后面每处都要跟着错。
- **别手写区块枚举。** 5 个维度有 39 个区块，凭记忆列必然漏。用
  `--list-sections` 拿清单，区块数由 `n + 2^n - 1 + 3` 断言校验。
  > 历史坑：早先手写的枚举漏掉过 `Model+Engine+Keyword`。
- **`L4` 和 `I4` 在无衬线字体里长得一模一样** —— 都是同一个竖条。读错了不会报错，
  只会让所有含 `Cyl` 的区块措词全歪。拿不准就抽一行跟原图对一下。
  Hyundai/Kia 系多用 `L4`，北美厂商多用 `I4`/`V6`。
- **`Cyl` 和 `Engine` 是两列，不能合并。** `5.7L V8 HEMI` 要拆成 `Cyl=V8` + `Engine=5.7L`。
  把 `V8` 混进 Engine 会让含排量的区块出现 `5.7L V8` 这种词，把 `2.5L` 混进 Cyl 同理。
- **恒定维度会翻倍行数但没有区分度。** 如果 `Cyl` 整列都是同一个值（比如全 `L4`），
  含它的区块产出的行只是把现有组合各复制一份。这是允许的，但汇报时要说清楚，
  别让用户以为多出来的 192 行是新增价值。真有多缸型混排时才有意义。
- **去连字符变体不要手写。** 脚本自动在原始行后插一行 `F-150 → F150`。
  手工加一条会和自动变体撞车，被去重规则丢掉，你会以为是脚本漏了。
  注意 `L4`、`V6` 不含连字符，不产生变体 —— 全表无连字符时 `Hyphen Variants Added` 就是 0。
- **`+` 只加在词上，不加在连字符上。** `15400-PLM-A02` → `+15400-PLM-A02`（一个 `+`），
  变体才变成 `+15400PLMA02`。
- **多词车型默认整体当一个 token。** `Grand Cherokee` → `+Grand Cherokee`。
  要逐词加 `+`（`+Grand +Cherokee`）才加 `--split-dims`。
- **开放年份区间会直接报错。** `2015-present` 没有终点，脚本报错并让你改写，不会瞎猜一个结束年。
- **`.xls` 不支持。** 让用户在 Excel 里另存为 `.xlsx`。
- **输出文件被 Excel 占用会写不进去。** 报错是 `Permission denied`，脚本会提示；
  让用户关掉那个文件，或用 `-o` 换个名字重跑。
- **字段为空 = 跳过，不是补全。** 没给排量时所有含排量的区块不出行，这是设计行为。
- **Sheet 2 的 `Sections` 可能小于 39。** 没有 `cyl` 就是 22，跳过 OE / 故障语义还会更少，
  别在汇报里硬说 39。

## 文件

| 路径 | 用途 |
| --- | --- |
| `scripts/build_combinations.py` | job.json → 3-sheet Excel；22 区块 / +版本 / 连字符变体 / 统计摘要 |
| `scripts/read_fitment.py` | xlsx / csv → fitment JSON；表头自动映射、engine 清洗、合并单元格填充 |
| `references/output-format.md` | 输出格式规格：区块表、sheet 结构、配色、去重口径 |
| `references/extraction.md` | 提取规则：干扰项、列名映射、年份写法、engine 清洗 |
