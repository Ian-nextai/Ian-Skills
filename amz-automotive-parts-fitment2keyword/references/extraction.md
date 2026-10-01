# 适配数据提取规则

只在**图片输入**或**手工粘贴的文本**需要时读这份文件。
xlsx / csv 走 `scripts/read_fitment.py`，不用人眼判断。

## 只取五列

| 字段 | 取什么 | 例 |
| --- | --- | --- |
| Year | 年份或年份区间 | `2009`、`2009-2022` |
| Make | 品牌 | `Dodge`、`Ford`、`Mercedes-Benz` |
| Model | 车型 | `Challenger`、`F-150`、`Grand Cherokee` |
| Cyl | 缸型（气缸排列 + 缸数） | `L4`、`V6`、`I6`、`H4` |
| Engine | **只有**排量数值 + 单位 | `5.7L`、`3.5L`、`2.0L` |

`Cyl` 和 `Engine` **必须分开**：`L4` 进 Cyl，`2.5L` 进 Engine。
上图那种 `年份 | 车型 | L4 | 2.5L` 的四列表，第三列是 Cyl、第四列是 Engine。

## 缸型写法

同一个缸型在不同来源里写法不一样，**照抄源数据的写法**，不要自己换：

| 常见写法 | 含义 |
| --- | --- |
| `L4` / `I4` / `l4` | 直列四缸 |
| `L6` / `I6` / `V6` | 直列六缸 / V 型六缸 |
| `V8`、`H4`、`H6`、`W12` | 其余排列 |

⚠️ 无衬线字体里大写 `I` 和小写 `l` 是**同一个竖条**，`L4` 和 `I4` 肉眼几乎分不出来。
拿不准就按源文件的其它位置判断：Hyundai/Kia 系的数据多用 `L4`，北美厂商多用 `I4`/`V6`。
**读错了不会报错，只会让区块 4 / 9 / 13 / 16 / 18 / 20 / 22 / 24 / 26 / 27 / 29 / 30 / 31 / 33 / 34 / 35 / 36 的措词全歪** —— 抽一行跟原图对一下最省事。

## 必须丢弃的干扰项

适配表里常年混着这些，一行都不要带进 Model / Cyl / Engine：

- 马力 / 扭矩：`395 hp`、`410 lb-ft`
- 驱动方式：`4WD`、`AWD`、`FWD`、`RWD`
- 车身类型：`Sedan`、`Coupe`、`Crew Cab`、`Extended Cab`
- 变速箱 / 门数 / 配置包：`Automatic`、`4-Door`、`Sport Package`
- 营销技术名：`Hemi`、`VTEC`、`EcoBoost`、`Turbo`、`TDI`
  —— 这些**不是**缸型，别塞进 `Cyl`，直接丢
- 备注列：`Notes`、`Comment`、`Source`

判断口径：`Cyl` 只放**气缸排列 + 缸数**（`L4`/`V6`），`Engine` 只放**排量**（`5.7L`）。
`5.7L V8 HEMI` 这种挤在一格里的，拆成 `Cyl=V8` + `Engine=5.7L`，`HEMI` 丢掉。

## 保真要求

- Make 里的连字符保留：`Mercedes-Benz`，不要拆成两个词，也不要丢连字符。
- Model 保留厂商写法：`F-150`、`Silverado 1500`、`Grand Cherokee`。
- OE 号保留原始大小写和连字符：`15400-PLM-A02`。
- 去连字符变体由脚本生成，**不要**手工加，否则会和自动变体重复并被去重规则吃掉一条。

## 年份写法

Year 字段原样传给脚本，**不要在对话里手算展开**：

| 输入 | 脚本展开为 |
| --- | --- |
| `2009` | 2009 |
| `2009-2022` | 2009 … 2022（14 年） |
| `2015, 2017-2019` | 2015, 2017, 2018, 2019 |
| `2009/2010`、`2009、2010`、`2009 2010` | 2009, 2010 |
| `2015-present` | ❌ 报错。开放区间无法确定终点，改写成 `2015-2025` 或只留 `2015` |

分隔符支持半角/全角逗号、分号、斜杠、顿号、空格；区间符号支持 `-`、`–`、`—`、`~`。

## 表格列名映射

`read_fitment.py` 自动识别，也会在前 10 行里找表头（适配表常有一两行标题在上面）。

| 字段 | 识别到的列名（不分大小写、忽略空格下划线） |
| --- | --- |
| Year | year, years, model year, my, 年份, 年款, 年度, 车型年份 |
| Make | make, brand, manufacturer, mfr, 品牌, 厂商, 制造商 |
| Model | model, models, car, vehicle, 车型, 车系, 型号 |
| Cyl | cyl, cylinder, cylinder config, engine config, 缸型, 缸数, 气缸数, 发动机配置 |
| Engine | engine, engine size, displacement, liter/litre, 排量, 排气量, 发动机 |

`Model Year` 会归到 Year 而不是 Model（全等匹配优先）；`Engine Config` 归到 Cyl 而不是 Engine。

认不出来时用 `--map` 明确指定：

```bash
python scripts/read_fitment.py table.xlsx --map Year="Model Year" --map Cyl=Config
```

### 合并单元格

Excel 里 Make/Model 常被合并成一个跨多行的单元格，openpyxl 读到的是首行有值、其余为空。
脚本会提示「有 N 处单元格为空但上一行有值」，按提示加 `--forward-fill` 即可：

```bash
python scripts/read_fitment.py table.xlsx --forward-fill Make,Model,Year,Cyl
```

如果表格结构更怪，先另存一份、把合并单元格拆开再喂给脚本，比在提示词里描述更省事。

## engine 列自动清洗

默认开启（`--no-clean-engine` 关闭）。规则是「找出排量 -> 归一成 `X.YL`」，找不到就原样保留并告警：

| 原始值 | 清洗后 |
| --- | --- |
| `5.7L V8 HEMI` | `5.7L` |
| `V6 3.5 Liter` | `3.5L` |
| `2.0L` | `2.0L` |
| `V8` | `V8`（告警：找不到排量 —— 说明这格其实是缸型，该放进 Cyl 列） |

清洗动作会写进 stderr 的 `warning:`，抽查几条确认口径对了再往下走。

**`Cyl` 列不做任何归一化**，源数据怎么写就怎么留。

## 多行输入

多条适配数据（多行）逐行独立生成组合，再在区块内汇总去重。
不要因为「这几行只差一个年份」就合并——年份已经在 Year 字段里表达过了。
