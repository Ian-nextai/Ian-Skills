---
name: ian-search
description: >-
  批量信息检索、网页抓取与表格字段补全工作流。当用户有多条公司 / 供应商 / 参展商 / 商品记录，需要从多个来源补全产品、官网、联系方式、目录、地址或展位等字段时使用。流程覆盖官方 API 与网页抓取、搜索引擎检索、邮箱域名反查、名称与别名匹配、低可信推断、来源标注、逐行合并和单 Sheet Excel 交付。触发场景：批量查资料、批量补产品、供应商信息补全、参展商信息整理、爬展会官网、网页抓取后合并 Excel、给每条记录标注信息来源。
---

# ian-search：批量信息检索与补全

把一份不完整的结构化清单，变成**每一行都有值、且每个值都有来源**的单表。

**核心原则：官方 / 内部数据优先；搜索只补空缺；推断必须标注；没有就留空。**

## 何时使用

- 用户给了一个 Excel / CSV / JSON 清单，要求补全公司、产品、官网、联系人、目录、展位等字段
- 用户要求从网站批量抓取参展商 / 会员 / 商品，并与已有表匹配
- 用户要求“没数据的去网上查”“把产品都补上”
- 结果需要合并成**一个 Sheet**，且保留来源、链接和空白单元格语义

若只查一个简单事实，用普通搜索即可，不必启动本 skill。

## 工作流

- [ ] **Step 1**：确认输入、目标列和交付形态
- [ ] **Step 2**：识别权威来源，先抓官方 / 结构化数据
- [ ] **Step 3**：做名称匹配，合并内部数据与官方数据
- [ ] **Step 4**：只对缺口执行网页 / 搜索补全
- [ ] **Step 5**：清洗、去重、标注来源；低可信值标“推断”
- [ ] **Step 6**：生成单 Sheet Excel 并做统计核验

详细来源顺序见 [references/merge-priority.md](references/merge-priority.md)，各渠道操作要点见 [references/sources.md](references/sources.md)。

---

## Step 1：锁定字段口径

先明确三件事：

1. **主键**：公司名 / 企业统一社会信用代码 / 展商 ID
2. **目标字段**：至少包括“产品”字段，以及用户点名的展位、目录、官网等
3. **空白规则**：缺失就留空，不要写 `N/A` / `无` / `未知`

建议统一字段：

| 字段 | 含义 |
| --- | --- |
| `name` | 展示名称 |
| `product` | 最终产品描述 |
| `product_source` | 产品来源 |
| `product_link` | 产品证据链接 |
| `source_detail` | 补充来源说明 |
| `inferred` | 是否推断 |

每个目标字段都保留“值 / 来源 / 链接”三元组，不要只留最终值。

---

## Step 2：先抓权威来源

优先级：

1. 用户提供的内部 Excel / 数据库
2. 活动官网、协会官网、官方 API
3. 企业官网
4. 搜索引擎结果
5. 名称 / 分类推断

官方 API 常在前端 JS 里暴露公开调用方式。**不要把真实 API Key、Cookie 或登录态写入脚本、README 或示例输出。**

大响应先落地为 `raw.json`，再做清洗；不要反复请求同一份全量数据。

操作细节见 [references/sources.md](references/sources.md)。

---

## Step 3：名称匹配

匹配顺序：

1. 精确名称
2. NFKC + 小写 + 去标点
3. 中英文别名
4. 去公司后缀：`有限公司` / `Co., Ltd.` / `GmbH` 等
5. 导入名称 / 历史名称
6. 人工确认映射表

模糊匹配只用于“找候选”，不要自动覆盖已有的工商数据。品牌名、简称和地名很容易误合并。

---

## Step 4：批量补缺口

联网操作用环境里的 `web-access` skill 执行。

建议把待查记录先标准化为 JSONL：

```json
{"id":"demo-1","fields":{"name":"示例公司","hall":"示例展区","email":""}}
```

生成查询词：

```bash
python scripts/batch_enrich.py \
  --records records.jsonl \
  --results results.jsonl \
  --config config.json \
  --queries-out queries.jsonl \
  --out enriched.jsonl \
  --stats stats.json
```

脚本会根据 `query_cn` / `query_en` 模板生成查询词，并用确定性规则过滤结果。示例配置：

```json
{
  "name_field": "name",
  "product_field": "product",
  "source_field": "product_source",
  "link_field": "product_link",
  "source_label": "网络检索",
  "query_cn": "{name} 产品",
  "query_en": "\"{name}\" products manufacturer",
  "product_terms": ["brake", "filter", "sensor", "刹车", "滤清", "传感器"],
  "generic_terms": ["汽车配件", "auto parts"],
  "min_source_len": 2,
  "require_existing_source": false,
  "max_product_len": 80
}
```

`require_existing_source=true` 时，已有非空产品字段不会被覆盖；否则只在已有来源很短或为空时补。

web-access 结果 JSONL 每行格式：

```json
{"id":"demo-1","results":[{"title":"...","url":"https://example.invalid/","snippet":"...","content":"..."}]}
```

**不要**在公开 Skill 或脚本里保存真实 API Key、Cookie、代理账号或客户数据。示例域名统一用 `example.invalid`。

---

## Step 5：清洗与合并

`scripts/common.py` 提供可复用函数：

- `norm()` / `name_key()`：名称归一化
- `clean_text()`：去 HTML、规范空白
- `accept_product()`：过滤噪音和过泛词
- `extract_product()`：从搜索结果提取产品短语
- `fill_template()`：生成多语言查询词

覆盖规则：

- 官方 / 内部数据 > 官网 > 搜索 > 推断
- 新值必须包含产品词、必须和公司名出现在同一结果上下文
- 地址、电话、注册信息、招聘、认证、导航栏文本一律不写入产品字段
- 产品词去重后用 `、` 连接
- 推断值来源写清楚：`公司名称推断` / `展区推断` / `目录推断`

---

## Step 6：生成单 Sheet Excel

配置文件：

```json
{
  "title": "参展商信息总表",
  "note": "空白单元格表示源数据中没有该项；来源列标注每段数据的出处。",
  "fields": ["name", "product", "product_source", "product_link"],
  "highlight_fields": ["product", "product_source"],
  "freeze": "C5"
}
```

执行：

```bash
python scripts/build_sheet.py \
  --records enriched.jsonl \
  --config sheet_config.json \
  --out final.xlsx
```

脚本会生成标题行、说明行、表头、冻结窗格、筛选和基础样式。缺失值留空。

---

## 交付前核验

必须报告：

1. 总行数、列数、Sheet 数
2. 各来源数量，例如官网 / 内部库 / 搜索 / 推断
3. 空白数：用户要求全部有值时，确认关键字段无空白
4. 变更抽样：5–10 条，显示公司名、补到的值、来源、链接
5. 剩余推断数量及原因
6. 已知失败：验证码、无公开页面、名称歧义等

**不要把“推断”说成“查到”。** 如果某批主要靠分类推断，要明确告知用户建议二次核实。

## 脱敏检查

发布或提交前检查：

- 无 API Key / Token / Cookie / 内网路径 / 代理账号
- 无真实客户名单、手机号、邮箱、联系人
- 示例只用 `示例公司`、`example.invalid`、`127.0.0.1:<端口>`
- 缓存文件、日志和截图里的敏感信息已清除
- `git diff` 与所有新增文件复查一遍

## 文件

| 路径 | 用途 |
| --- | --- |
| `scripts/common.py` | 名称归一化、产品词过滤、查询模板填充、JSON 工具 |
| `scripts/batch_enrich.py` | web-access 结果 → 逐行过滤与回填 |
| `scripts/build_sheet.py` | JSONL → 单 Sheet Excel |
| `references/merge-priority.md` | 来源优先级、匹配与覆盖规则 |
| `references/sources.md` | 官方 API / 搜索引擎 / 官网 / 代理 / 脱敏操作要点 |
