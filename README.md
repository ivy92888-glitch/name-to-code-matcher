# name-to-code-matcher / 名称匹配编码器

> 零依赖的 VBA 宏：把一列「自由文本品名」（可带材质等属性），对照一张或多张可配置的参考表，自动匹配出**标准编码**并写回表格。
> A dependency-free VBA macro that maps a free-text item name (plus an optional attribute like material) to a standardized code by looking it up against configurable reference tables.

> **非官方。** 与 Microsoft（Excel）及任何电子表格厂商无隶属、赞助或背书关系；"workbook" 仅作为表格文件的通用指代。

---

## 痛点：它解决什么问题（中文）

有一类工作反复出现，形状完全一样：**你有一堆乱糟糟的自由文本名称，想给每一个贴上一个「标准编码」**。

典型例子：

- 给商品贴分类编码；
- 把描述映射到科目/类别编码；
- 把用户手填的标签规范到受控词表。

**手工做法的崩溃点**：

- 名称千奇百怪（"蓝色 Widget 支架""工业用 Widget"），人肉去对照词表，慢且标准不一；
- 容易误匹配——比如参考里有 `Widgetron`，结果把含 `Widget` 的名字全错配上；
- 参考表经常不止一张，还要分优先级（先精确表、再宽泛表），手工很难稳定执行；
- 同名不同材质（金属/塑料）要区分，靠肉眼判断属性类别太累。

**这个宏的价值**：内置一套**通用匹配引擎**，把上面这些痛点一次性解决——

1. **词边界感知的子串匹配**：参考名必须出现在非词字符（空格、标点）或字符串边界之间，`Widget` 能匹配 `Blue Widget Stand` 却不会误匹配 `Widgetron`（中英文都适用）；
2. **属性类别过滤**：源行和参考行都带属性（如材质）时，只有同属一大类（都金属/都塑料）才接受；
3. **多级降级**：参考表按顺序尝试，先命中先得，可用「精确表压在宽泛表之上」；
4. **加权择优**：多个编码命中时，用「命中频次 0.3 + 行新鲜度 0.7」的加权分挑最好的。

全程配置驱动，没有任何特定业务被写死。

## What problem it solves (English)

The same task keeps recurring: you have messy free-text names and want each tagged with a canonical code from a reference list. Doing it by hand is slow and inconsistent, and naive substring matching mis-fires (e.g. `Widget` wrongly hits `Widgetron`). This macro ships a reusable engine — boundary-aware matching, attribute-category filtering, multi-tier fallback, and a weighted best-candidate resolver — all configuration-driven, nothing domain-specific hard-coded.

---

## 快速运行（中文）

**环境要求**

- Excel 或任何兼容的电子表格软件（支持 VBA）；
- 一张源表 + 一张或多张**放在同一工作簿内**的参考表。

**五步跑起来**

1. 打开工作簿，按 `Alt + F11` 导入 `src/NameToCode.bas`；
2. 把你的参考表放进**同一个工作簿**（每张表有表头、数据从第 2 行起）；
3. 改模块顶部常量，匹配你的表结构（见下表）；
4. 把工作簿另存为 `.xlsm`；
5. 按 `Alt + F8` 运行宏 **`MatchCodes`**。

**默认列布局**

| 列（默认） | 含义 | 必填 |
|-----------|------|------|
| C（NAME_COL） | 品名（自由文本） | 是 |
| Y（ATTR_COL） | 属性，如材质（可留空关闭） | 否 |
| E（CODE_COL） | 输出：匹配到的编码 | 是 |
| F（SOURCE_COL） | 输出：命中了哪一级（可关闭） | 否 |

`HEADER_ROW = 3`，`DATA_START_ROW = 4`。参考表名用 `-` 可去掉后缀、`/` 可列别名，每个别名独立匹配。

**想立刻看到效果**：打开 `examples/sample-data.xlsx` 直接运行（示例：`Blue Widget Stand` 命中 `Widget Stand` 得 `W-100`，而 `Industrial Widget` 不会误配 `Widgetron`）。

## Quick start (English)

1. Open your workbook, `Alt+F11` → import `src/NameToCode.bas`.
2. Put reference tables in the **same workbook** (header row + data from row 2).
3. Edit the config constants at the top.
4. Save as `.xlsm`, press `Alt+F8`, run **`MatchCodes`**.
5. Try `examples/sample-data.xlsx` first.

---

## 工作原理 / How it works

1. 读取源表（活动表或指定表）的品名与可选属性；
2. 按 `TIERS` 顺序逐层扫描参考表（自底向上），收集精确与边界感知的子串命中，按属性类别过滤，用加权分化解并列；
3. 第一层返回编码即胜出，写回编码列、记录来源层级、按层级给单元格着色；
4. 无匹配写 `N/A`，出错行标红。

**关键常量（唯一真源 `src/NameToCode.bas`）**

| 常量 | 默认 | 含义 |
|------|------|------|
| `NAME_COL` | `"C"` | 品名列 |
| `ATTR_COL` | `"Y"` | 属性列（`""` 关闭） |
| `CODE_COL` | `"E"` | 输出编码列 |
| `TIERS` | `RefPrimary\|A\|B\|C, …` | 有序参考表（`表名\|名列\|编码列\|属性列`） |
| `W_COUNT` / `W_RECENCY` | `0.3` / `0.7` | 频次权重 / 新鲜度权重 |
| `ATTR_FILTER_ON` | `True` | 按属性类别过滤 |

属性分类：若 `CATEGORY_SHEET` 指定的表存在，按 `关键词, 类别` 读取；否则用内置的材质分类（塑料/金属/木/纸/纺织/玻璃/矿物/复合），均可替换为自己的类别。

## 局限 / Limitations

- 参考表必须和宏在同一工作簿（`ThisWorkbook`）；
- 仅做整名相等或边界感知子串匹配，不做模糊/拼音匹配；
- 评分目标（频次+新鲜度）固定，改目标需改 `BestCode`；
- 作者环境未实际执行该宏；正式使用前请在数据副本上测试。

## 适用场景 / Use cases

商品分类打码、描述→科目/类别编码、自由标签规范化到受控词表——任何「自由文本名称 → 标准编码」的映射。

## 法律声明 / Legal & Non-affiliation

本仓库是通用匹配模式的**独立、clean-room 重写**，不含任何第三方的专有数据、内部命名或业务特定逻辑，作者与任何可能启发类似流程的公司无隶属关系、也不代表其立场。MIT 许可证仅覆盖作者本人贡献；使用者需确保用途符合所在地法律与单位政策。

## License

[MIT License](LICENSE) · Copyright (c) 2026 ivy92888-glitch.

---

<details><summary>维护说明（中英双语单文件，怎么避免漂移）</summary>

本文档是**单文件中英双语**：中文为权威详细描述，英文为简洁摘要（不强求逐句对等）。所有参数名（如 `TIERS`、`W_RECENCY`）直接取自源码常量，唯一真源是 `src/NameToCode.bas`。修改代码逻辑时，只要同步更新源码常量名与本文档常量表中的同名项即可，叙述性文字按需增删，不影响一致性。

</details>
