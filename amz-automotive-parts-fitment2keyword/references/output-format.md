# 输出格式规格

生成逻辑在 `scripts/build_combinations.py` 里。改格式时改脚本，不要靠提示词描述。
这份文件用于**核对**脚本产出、以及回答用户关于格式的问题。

完整区块清单**不在这里复制** —— 直接问脚本，它才是唯一事实来源：

```bash
python scripts/build_combinations.py --list-sections
```

## 区块是怎么算出来的

区块 = **纯维度区块** + **每个维度子集配 Keyword 的区块** + **3 个补充信息区块**。

```
区块数 = n + (2^n - 1) + 3        # n = 维度个数
```

| 维度数 n | 维度 | 区块数 |
| --- | --- | --- |
| 4 | Year / Make / Model / Engine | 4 + 15 + 3 = **22** |
| 5 | Year / Make / Model / Cyl / Engine | 5 + 31 + 3 = **39** |

维度在 `DIMENSIONS` 里按固定顺序排列，Keyword 区块用 `itertools.combinations` 逐层展开：

- **1 维**（5 组）：`Year+Keyword` … `Engine+Keyword`
- **2 维**（10 组）：`Year+Make+Keyword` … `Cyl+Engine+Keyword`
- **3 维**（10 组）：`Year+Make+Model+Keyword` … `Model+Cyl+Engine+Keyword`
- **4 维**（5 组）：`Year+Make+Model+Cyl+Keyword` … `Make+Model+Cyl+Engine+Keyword`
- **5 维**（1 组）：`Year+Make+Model+Cyl+Engine+Keyword`

`Data Format` 列的值 = 区块名，token 顺序与 `DIMENSIONS` 一致。
所以 5 维那一条读作 `2021 Hyundai Santa Fe L4 2.5L sensor`（缸型在排量前面）。

**加维度要改 `DIMENSIONS` 一处即可。** 脚本里的断言会按上面的公式校验区块数，
改错立刻报错，不会悄悄少几个区块。

> 历史坑：早期手写的枚举列表漏掉了 `Model+Engine+Keyword`，但区块总数一直写着 22。
> 现在改成公式生成，这类问题不会再出现。

## 依赖关系

| 区块 | 依赖 | 缺失时 |
| --- | --- | --- |
| 前 n 个纯维度区块 + 全部 Keyword 区块 | fitment 数据 | 该维度整列为空则跳过所有含它的区块 |
| `OE+Keyword` | `oe` | 跳过 |
| `Problem+Keyword` | `problems` + `keywords` | 跳过 |
| `Problem+OE` | `problems` + `oe` | 跳过 |

跳过是**正常路径**，不是错误。Sheet 2 的 `Sections` 会如实反映实际产出的区块数。

## Sheet 1「关键词组合结果」

三列，无序号列、无额外汇总列：

| Data Format | 普通版本 | +号限定符版本 |
| --- | --- | --- |
| `Year` | `2009` | `+2009` |
| `Year+Make+Keyword` | `2009 Dodge oil filter` | `+2009 +Dodge +oil +filter` |
| `Model` | `F-150` | `+F-150` |
| `Model` | `F150` ← 黄底变体 | `+F150` ← 黄底变体 |
| `OE+Keyword` | `15400-PLM-A02 oil filter` | `+15400-PLM-A02 +oil +filter` |
| `OE+Keyword` | `15400PLMA02 oil filter` ← 黄底变体 | `+15400PLMA02 +oil +filter` ← 黄底变体 |

**样式**

| 位置 | 值 |
| --- | --- |
| 表头 | 底色 `#4472C4`，白色粗体，水平居中 |
| 数据行 | 白色 / `#DCE6F1` 交替；左对齐；细灰边框 `#BFBFBF` |
| 连字符变体行 | `#FFF2CC`（覆盖交替底色） |
| 冻结 | 首行（`A2`） |
| 列宽 | 自动，上限 100 |
| 页脚 / 序号列 | 无 |

蓝白交替只对**原始行**计数，变体行固定黄色。否则插入变体后会出现两行同色。

## Sheet 2「测试摘要」

两列 `项目 | 值`。前 5 行口径固定：

| 项目 | 值 | 含义 |
| --- | --- | --- |
| `Total Rows` | 499 | 所有区块生成的原始行数（不含变体） |
| `Sections` | 39 | **实际产出非空行**的区块数；缺 Cyl/OE/Problem 时会小于 39 |
| `Original Rows` | 499 | 同 `Total Rows`，保留旧口径 |
| `Hyphen Variants Added` | 0 | 追加的黄色变体行数 |
| `Total Output Rows` | 499 | `Original Rows + Hyphen Variants Added` |

空一行后是诊断信息（适配数据行数、KW/OE/Problem 个数、跳过的区块），方便回溯。

## Sheet 3「原始适配数据」

列 = `DIMENSIONS`，但**整列全空的维度不输出**：

- 有缸型数据 → `Year | Make | Model | Cyl | Engine` 五列
- 没有缸型数据 → 回到 `Year | Make | Model | Engine` 四列

一行一个唯一的维度组合，**年份范围已经展开成逐年行**，可以直接跟原始输入逐行核对。

## 连字符变体规则

任何 token 含 `-` 的行，紧跟一行去掉**全部**连字符的版本（OE 号、车型号一视同仁）：

- `F-150` → `F150`；`+F-150` → `+F150`
- `15400-PLM-A02` → `15400PLMA02`；`+15400-PLM-A02` → `+15400PLMA02`
- 变体行的 `Data Format` **沿用父区块名**（`Model`、`OE+Keyword` …），靠黄色底色区分

注意 `L4`、`V6` 这类缸型**不含连字符**，所以不会产生变体行。整个数据集没有连字符时，
`Hyphen Variants Added` 就是 0，Sheet 1 里一行黄色都不会有 —— 这是正常的。

## 去重

同一区块内按「普通版本」字符串去重；变体行也参与去重，
所以如果输入里同时有 `F-150` 和 `F150`，只会保留先出现的那条。

同一个 `Data Format` 里重复的输入行（例如 OE 号列表里重复写了同一个件号）也在这里被吃掉。

## 文本区块（`--blocks-file`）

给需要把组合粘到别处用的场景。结构：先全部普通版区块，再一段 `# 【+号限定符版本】`，后接同样的区块。
区块标题用中文，如 `## 区块 13：年份 + 缸型 + Keyword`。
