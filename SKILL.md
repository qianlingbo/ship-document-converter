---
name: ship-document-converter
description: 将船舶 IMO Crew List + Port of Call 转换为海事局标准录入格式
category: productivity
tags: [maritime, 单证录入, maritime-declaration]
version: 2.13.0
---

# 单证录入技能

> 将船舶 IMO Crew List + Port of Call 转换为海事局标准录入格式

**版本**: v2.10.0 | **Python**: 3.8+ | **依赖**: `openpyxl`, `xlrd`

### ⚠️ 中国船员姓名必须有中文汉字（不能只写拼音）

这是用户明确提出的标准要求。但 pypinyin 等工具只能转拼音（`ZHANG SAN` → `ZHANGSAN`），无法确定具体汉字。

**正确做法**：
1. 有含中文名的原始证件扫描件 → 从中提取
2. 用户直接提供中文名 → 直接填入
3. 无法获取时 → **告知用户**，让用户提供，不自行猜测

**禁止**：
- 自行用拼音转汉字（如把 `TIAN JINGWEI` 写成 `田景伟`）
- 输出中英混合：`TIAN JINGWEI 张三` / `张三 TIAN JINGWEI`
- 输出纯拼音：`ZHANGSAN`（国籍为 CN 时）

> ⚠️ **实际测试发现的关键坑点**：
> - `port_map.json` 是 **dict** 结构（key=港口代码，value=`"CODE-中文名(英文名)"`），**不是 list-of-dicts**，直接遍历会报错 `AttributeError: str has no .get()`
> - `duty_map.json` 的实际格式是 `{"51-船长": "51-船长", ...}`（代码→代码），不是 `{"MASTER": "51-船长"}`。英文职务缩写必须靠 `ENGLISH_RANK_MAP` 硬编码映射表实现，不得从 duty_map.json 推导。
> - `nationality_map.json` 的 key 是 2位字母码（如 `"CN"`），value 是 `"CN-中国"`。英文国名（如 `"CHINA"`）不在 keys 中，需用 `poc-country-map.json` 或代码中硬编码的 `NAT_FULL_TO_CODE` 反向映射表。
> - Crew List 单行 tuple 长度是 28（含大量 None 列），解析时需用明确列索引，不能用顺序解包。
> - `parse_date()` 时间分配：进港时间 `00:00-11:59`，离港时间 `12:00-23:59`。

## 用户输出偏好

回复风格：**简洁干净**，不装饰花哨符号（ヽ(✿ﾟ▽ﾟ)ノ 之类禁止）。推荐样式：`✌️` / `ok` / `✅` / emoji适度点缀即可。

报告内容优先顺序：船舶信息 → Sheet1人数/职务概览 → Sheet6港口数 → 证件类型说明。不需要逐行列举所有数据。

## 快速开始

```bash
pip install openpyxl xlrd
python3 scripts/单证录入核心.py input/crew_list.xlsx [port_of_call.xlsx] [输出名]
```

## v3 xlsm 导出脚本

已有可复用脚本 `scripts/xlsx_to_v3_xlsm.py`：
```bash
python3 scripts/xlsx_to_v3_xlsm.py 单证录入_COSCO_20260617.xlsx "COSCO SHIPPING WISDOM" 20260617
```

## 目录结构

```
.
├── scripts/
│   ├── 单证录入核心.py           # 核心脚本
│   └── poc录入_MINHUA9.py       # POC专用脚本（参考用）
├── templates/
│   ├── 单证录入标准格式_v2.xlsx  # 旧6-sheet模板（备用）
│   └── 单证录入标准格式_v3.xlsm  # 优先使用！17-sheet，保留VBA宏，Row2=说明行，数据从Row3开始
├── references/
│   ├── nationality_map.json     # 国籍代码 (248条) — 参数A
│   ├── duty_map.json            # 职务代码 (12条) — 参数B
│   ├── port_map.json            # 港口代码 (1956条) — 参数E
│   ├── poc-country-map.json     # POC国家字段专用映射（英文国名→代码-中文名）
│   ├── zyhy-jinqu-layout.md     # 中远海运津渠轮 Crew List 布局笔记
│   ├── zyhy-jinqu-v39-layout.md # 中远海运津渠 V39 布局笔记（数据从第12行起，28列tuple）
│   ├── universe-harmony-layout.md # UNIVERSE HARMONY Crew List 布局笔记（含POC港口映射）
│   ├── tianwangzhixing-poc-layout.md # 天王之星 POC layout（实测列索引/日期格式/港口匹配）
│   ├── fortune-progress-poc-layout.md # FORTUNE PROGRESS POC layout（xlsx格式，yyyy.mm.dd日期，无UNLOCODE列）
│   ├── fortune-progress-crew-layout.md # FORTUNE PROGRESS Crew List layout（.xls姓名双列，15船员）
    - references/universe-harmony-layout.md
│   ├── millie-crew-layout.md    # MILLIE Crew List layout（奇偶行结构，子行col9=登船地点）
- references/millie-poc-layout.md     # MILLIE POC layout（混合日期格式，船长签名行跳过）
- references/millie-crew-layout.md   # MILLIE Crew List layout（奇偶行结构，子行col9=登船地点）
- references/green-munguba-crew-layout.md  # GREEN MUNGUBA Crew List（20人，col10=海员证号，6人登船日期空白）
- references/green-spetiba-crew-layout.md  # GREEN SEPETIBA Crew List 布局（20人中国船员）
- references/green-spetiba-poc-layout.md  # GREEN SEPETIBA POC layout（10港口，PDF格式）
- references/cosco-shipping-wisdom-embark-ports.md  # COSCO SHIPPING WISDOM 登船口岸+POC港口实测
- references/green-salvador-crew-layout.md  # GREEN SALVADOR Crew List（19人，CHINESE nationality）
- references/green-salvador-poc-layout.md  # GREEN SALVADOR POC（10港口，text PDF）
- references/kangshun99-layout.md         # KANGSHUN 99 Crew(.xls)15人+POC(11港口)，含护照号col13确认
- references/v3-template-layout.md       # v3模板(17-sheet xlsm)行偏移规则及脚本
  - GREEN SALVADOR 港口代码已修正：天津→CNTNJ-天津港(Tianjin)，光阳→KRKAN-光阳(Gwangyang/Kwangyang)；TAICANG出发日期=2026/2/05
- references/conversion-rules.md      # 转换规则详解（含所有坑点）
- references/wen-de-poc-layout.md       # WEN DE PORT CALL LIST（10港口，需确认日期）
- references/guanghua-poc-layout.md    # GUANG HUA POC layout（10港口，含4个代码修正）
├── input/                       # 原始文件
└── output/                      # 输出文件
```

---

## ⚠️ 强制规则（Agent 必须严格遵守，不得自由发挥）

### 1. 姓名

| 船员国籍 | 规则 | 示例 |
|----------|------|------|
| **中国（CN）** | **只录中文，禁止出现任何英文/拼音** | 源数据 `张三 ZHANG SAN` → 输出 `张三` |
| 外国 | 全部转大写英文 | 源数据 `nguyen van a` → 输出 `NGUYEN VAN A` |

**中国船员姓名细则：**
- 源数据中英混合（如 `张三 ZHANG SAN`）→ 只保留中文部分 `张三`
- 源数据纯中文（如 `张三`）→ 原样输出 `张三`
- 源数据纯英文/拼音（如 `ZHANG SAN`，但国籍为 CN）→ 原样保留（无法自动生成中文）
- **绝对禁止**：输出 `张三 ZHANG SAN`、`ZHANG SAN 张三`、`Zhang San` 等中英混合格式

### 2. 性别

| 源数据 | 输出 |
|--------|------|
| `M` / `male` / `男` / `1` | `1-男` |
| `F` / `female` / `女` / `2` | `2-女` |
| 空值或无法识别 | 默认 `1-男` |

**强制格式**：`1-男` 或 `2-女`（带数字前缀和连字符）。

### 3. 船员职务

**必须使用映射表**，禁止 Agent 自行翻译。

映射表位于 `references/duty_map.json`（12种标准职务）+ 脚本内置 `ENGLISH_RANK_MAP`。

| 英文缩写 | 标准职务 |
|----------|----------|
| MASTER / CAPT / CAPTAIN | 51-船长 |
| C/O / CHIEF OFFICER / 1/O / FIRST OFFICER | 52-大副 |
| 2/O / SECOND OFFICER | 53-二副 |
| 3/O / THIRD OFFICER | 54-三副 |
| BSN / BOSUN | 55-值班水手 |
| AB / AB1 / A/B / ABLE SEAMAN | 56-高级值班水手 |
| C/E / CHIEF ENGINEER | 61-轮机长 |
| 1/E / 2/E / FIRST ENGINEER / SECOND ENGINEER | 62-大管轮 |
| 3/E / THIRD ENGINEER | 63-二管轮 |
| 4/E / FOURTH ENGINEER | 64-三管轮 |
| FITTER / FTR / COOK / STEWARD / PUMPMAN 等 | 65-值班机工 |
| ETO / E/E / ELECTRICIAN / OIL / OILER 等 | 66-高级值班机工 |
| COMMISSAR | 51-船长（政委） |
| ELECTRO-TECHNICAL OFFICER | 66-高级值班机工 |
| CHIEF MOTORMAN / MOTORMAN | 65-值班机工 |
| CADET CAPTAIN | 55-值班水手 |
| CADET ETO | 66-高级值班机工 |
| DFTR / EFTR / WIPER / C/COOK / MSM | 65-值班机工 |
| **C/CK** | 65-值班机工 |
| **D/C** | 65-值班机工 |
| **E/C** | 65-值班机工 |
| **WPR** | 65-值班机工 |

**职务 fallback 规则**（映射表找不到时）：
1. 识别角色类型（甲板=sailor / 机舱=engineer）
2. sailor 前3人 → `56-高级值班水手`，其余 → `55-值班水手`
3. engineer 前3人 → `66-高级值班机工`，其余 → `65-值班机工`

### 4. 出生地点

| 船员国籍 | 规则 | 示例 |
|----------|------|------|
| **中国（CN）** | **一律填 `中国`** | 源数据 `SHANDONG` / `山东` / `Qingdao` / `青岛市` → 一律输出 `中国` |
| 外国 | 优先保留源数据原文；无源数据时用国籍中文名兜底 | 源数据 `HO CHI MINH` → 输出 `HO CHI MINH`；无源数据、国籍 VN → 输出 `越南` |

**中国船员出生地点细则：**
- 不管源数据给的是省名、市名、拼音、英文、全称、简称 → 全部输出 `中国`
- **绝对禁止**：输出 `山东`、`SHANDONG`、`Qingdao`、`中国山东` 等任何非 `中国` 的值

**外国船员出生地点规则（用户明确纠正，FORTUNE PROGRESS 实测）：**
- 统一填国籍的中文名，例如 `印度尼西亚`、`越南`、`菲律宾`
- 不能填源数据原文（如 `NGANJUK`），也不能留空

### 5. 进港时间 / 离港时间

**强制格式**：`yyyy.MM.dd HH:mm:ss`（**点分隔符，不是斜杠**）

| 字段 | 日期来源 | 时间规则 | 示例 |
|------|----------|----------|------|
| 进港时间 | Port of Call 的 Arrival 日期 | 随机 `00:00:00` ~ `11:59:59` | `2024.01.15 08:23:47` |
| 离港时间 | Port of Call 的 Departure 日期 | 随机 `12:00:00` ~ `23:59:59` | `2024.01.16 18:45:12` |

**禁止格式**：`2024-01-15`、`20240115`、`2024/01/15`、`Jan 15, 2024` 等。

### 6. 停靠港口

**必须使用 `references/port_map.json` 映射表**，匹配策略（三级优先）：

1. **优先精确匹配 UNLOCODE**（去掉空格后与 port_map 的 code 字段比较）
2. **其次 `SPECIAL_PORT_OVERRIDE` 手动映射表**（hardcoded in script）
3. **最后子串匹配 fullname**（`len(port_name) >= 4` 且 `len(code) >= 4` 才做子串，防止短码误匹配长字符串）

- 格式：`代码-中文名(英文名)`，如 `BRSSZ-BRSSZ-桑托斯/圣多斯(Santos)`
- 匹配不到 → fallback 到 UNLOCODE 构建的临时代码（如 `IQBSR-BASRAH`），并**标红该行**（浅红背景 `#FFFFCCCC`）
- **禁止**：Agent 自行编造港口代码或中文名

**⚠️ UNLOCODE 注意事项：**
- PDF/文件中 UNLOCODE 通常带空格（如 `BR SSZ`），需去掉空格后与 `port_map.json` 的 code 字段精确比较
- `CN NGB`（宁波）→ 正确匹配到 `CNNBO`，**不是** `CNNGB`（后者在 port_map 中不存在）
- `IQ BSR`（巴士拉）→ `port_map.json` 中无伊拉克港口，用 UNLOCODE 构建 `IQBSR-BASRAH` 并标红待确认

**⚠️ Port of Call 数据顺序：**
- 港口顺序是**从新到旧**（第1行=最新进港，第最后1行=最旧离港），**不要反转**
- PDF 中日期格式为 `26-03-29`（`%y-%m-%d`），**不是** `%d-%m-%Y`

### 7. 国籍

**必须使用 `references/nationality_map.json` 映射表**。

- 格式：`代码-中文名`，如 `CN-中国`、`VN-越南`
- 匹配不到 → **留空**，并**标红该行数据**（浅红背景 `#FFFFCCCC`）
- ⚠️ 不再默认填 `CN-中国`，避免错误数据混入

### 8. 出生日期

- 格式：`YYYYMMDD`（8位纯数字，无分隔符）
- 示例：`19860328`
- **禁止**：`1986-03-28`、`28/03/1986` 等

### 9. 证件类型 & 证件号码

| 国籍 | 证件类型 | 证件号码来源 |
|------|----------|--------------|
| 中国（CN） | `17-海员证` | **Seaman's Book**（见下表确定列索引） |
| 其他 | `14-普通护照` | Passport（见下表确定列索引） |

**⚠️ 不同文件格式列布局不同，必须打印表头行确认！禁止猜测列索引。**

| 文件格式 | Passport列 | SeamanBook列 | 备注 |
|----------|-----------|--------------|------|
| GREEN SEPETIBA (.xlsx) | **col9** | **col11** | 表头 Row8 含 `14.Passport&Expiry Date` / `15.Seaman book&Expiry Date` |
| 天王之星 (.xls, xlrd) | col10 | col8 | xlrd 0-based |

**验证方法**：读取 Row8 表头行，定位 `Passport` 和 `Seaman book` 所在列。

**⚠️ 禁止手动转录证件号码！** 必须从原始文件直接读取单元格值。今天 session 中 14-20 号共 7 人的海员证号因手动录入全部写错（A→E 开头），此类错误无法事后发现。处理流程：
1. 打开原始文件，逐行打印 crew 数据（包含 col9/col11 值）
2. 确认格式无误后，用脚本直接读取写入
3. 输出后与原始数据抽查比对

---

## 转换规则总览

### Sheet 1: 船上非旅客人员清单（16列）

模板结构：
- Row 1：表头（`*序号`/`姓名`/...）
- Row 2：**说明行**（"必填，此行为说明，不可删除，请从第三行开始填写..."）
- **数据从 Row 3 开始**，公式：`r = crew_no + 2`（Crew #1 → Row3, Crew #2 → Row4...）

> ⚠️ Row 2 是说明行，**不得删除**，数据从 Row 3 开始。v3 模板与 v2 不同（v2 数据从 Row 2 开始）。

| 列 | 字段 | 规则 |
|----|------|------|
| A | 序号 | 从1开始递增 |
| B | 姓名 | 见规则1 |
| C | 性别 | 见规则2 |
| D | 船员职务 | 见规则3 |
| E | 船员国籍 | 见规则7 |
| F | 出生日期 | 见规则8 |
| G | 出生地点 | 见规则4 |
| H | 证件类型 | 见规则9 |
| I | 证件号码 | 中国=海员证号，外国=护照号 |
| J | 是否申请登陆 | 留空 |
| K | 适任证书编号 | 留空 |
| L | 适任证书有效期至 | 留空 |
| M | 证件检查地点 | 留空 |
| N | 登船日期 | `YYYYMMDD` 格式 |
| O | 登船口岸 | 使用 port_map.json 映射 |
| P | 备注 | 留空 |

### Sheet 2: 船上非旅客人员物品清单

模板结构：
- Row 1：表头
- Row 2：**说明行**（"必填，此行为说明，不可删除，请从第三行开始填写..."）
- **数据从 Row 3 开始**，公式：`r = crew_no + 2`

> ⚠️ Row 2 说明行不得删除。删除模板示例行用 `delete_rows(3, 模板示例行数)`，不是清空内容。

### Sheet 3: 海事船岸活动信息（8列）

| 列 | 字段 | 规则 |
|----|------|------|
| A | 序号 | 从1开始递增，数据从 Row 3 开始（`poc_no + 2`） |
| B | 进港时间 | 见规则5 |
| C | 离港时间 | 见规则5 |
| D | 国家/地区名称 | 从港口名提取国家，使用 nationality_map.json + poc-country-map.json |
| E | 船舶保安等级 | 固定 `1-1级` |
| F | 特别或附加的保安设施 | 留空 |
| G | 停靠港口 | 见规则6 |
| H | 港口保安等级 | 固定 `1-1级` |

---

## 输入格式

- **Crew List**: 表头含 `No.` + `Family name` + `Rank`
- **Port of Call**: 表头含 `Voy.` + `Port`

### 仅处理 Port of Call 文件（无 Crew List）

当用户说"仅处理这个文件"且文件只有 Port of Call 时：
- 直接处理 Port of Call 数据，只填写"海事船岸活动信息" sheet
- 船员相关 sheet 留空，无需造假数据

---

## ⚠️ 坑点记录

### 📄 Port of Call `.doc` 文件读取（Word 格式，非 Excel）

**症状**：文件扩展名 `.doc` 误以为可以用 `xlrd` / `openpyxl` 读取，实际是 **Word 文档**（OLE2 compound document），两者都会报错。

**`xlrd` 报错**：`XLRDError: Can't find workbook in OLE2 compound document`
**`openpyxl` 报错**：`InvalidFileException: openpyxl does not support the old .xls file format`

**解法（macOS 内置）**：
```bash
textutil -convert txt -stdout "PORT OF CALL.doc" 2>/dev/null
```
`textutil` 是 macOS 自带工具，无需 pip 安装，可将 Word `.doc` 转为纯文本。

**读取后注意事项**：
- POC 数据以 `\u0007`（制表符）作为列分隔符，`split('\u0007')` 逐列提取
- OCR 日期变体：`16 JAN 2026`、`14FEB 2026`（月份前无空格），需用正则 `r'(\d{1,2})\s*([A-Z]{3})\s*(\d{4})'` 处理
- `.doc` 内文字可能混有乱码行（如 `SECURITY\n  LEVEL`），解析时过滤空行

**示例（MIN HUA 9 实测）**：
```
NO PORT OF CALL COUNTRY SECURITY
  LEVEL DATE OF
ARRIVAL DAT OF
DEPARTURE PURPOSE
01 WEDA INDONESIA 1 09JAN 2026 16 JAN 2026 DIS N LDG
```

### ⚠️ PDF 读取：pdftotext 优先于 pypdfium+OCR

**优先顺序**：
1. `pdftotext -layout <file> -` — 返回结构化纯文本，成功率最高（GREEN MUNGUBA 实测：完美提取表格）
2. `pypdfium2` 渲染 → PNG → Tesseract OCR（仅在步骤1完全失败时尝试）

### RED FILL 颜色格式（openpyxl）

`PatternFill` 颜色必须使用 **ARGB 格式**（8位hex），`FF` = 完全不透明：
```python
RED_FILL = PatternFill(start_color="FFFFCCCC", end_color="FFFFCCCC", fill_type="solid")  # ✅ 正确
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")      # ❌ 缺少Alpha通道 → 透明红色
```
`start_color="FFCCCC"` = 透明色（完全看不见）；`start_color="FFFFCCCC"` = 可见浅红色。

### `random_port_same_country()` bug

**bug**：函数从 port 名取前2字符作为国家码（如 `GO DAU` → `GO` 共和国，`LEAMCHABANG` → `LE` 莱索托），导致随机选到错误国家的港口。

**正确做法**：从 `_raw_country` 字段（原始国家名）用 `normalize_code()` 获取完整国家代码，再取其前2位作为国家前缀筛选 port_map。

### ⚠️ 港口代码必须核对 port_map.json（用户明确要求）

**不得直接信任源文件提供的 UNLOCODE**，必须逐个与 `port_map.json` 核对后填入。

已确认的错误代码（实测 GUANG HUA）：

| 源文件代码 | 正确代码 | 港口 |
|-----------|---------|------|
| AUPWL | AUPWA | 沃尔科特港(PORT WALCOTT) |
| CNBAN | CNBSD | 宝山(Baoshan) |
| CNNGB | CNNBO | 宁波(Ningbo) |
| CNZHA | CNZNG | 湛江港(Zhanjianggang) |

### Port of Call 港口 fallback 速查表

以下港口在 `port_map.json` 中不存在，备用映射已验证可用：

| 原始港口名 | 替代港口 | 替代代码 | 国家 |
|------------|----------|----------|------|
| NAPLES | 那不勒斯 | ITNAP-那不勒斯(NAPLES) | IT |
| MONFALCONE | 蒙法尔科内 | ITMFA-蒙法尔科内(MONFALCONE) | IT |
| GIBRALTAR | 直布罗陀 | GIBGI-直布罗陀(GIBRALTAR) | ES |
| ITAQUI | 伊塔基 | BRITQ-伊塔基(ITAQUI) | BR |
| CHENJIIAGANG | 陈家港 | CNCIG-陈家港(CHENJIAGANG) | CN |
| KENDARI | 乌戎潘当 | IDUPG-乌戎潘当(Ujung Pandang) | ID |
| MORMUGAO | 哈迪亚 | INHDA-哈迪亚(HALDIA) | IN |
| PASIR GUDANG | 巴西拉再也 | MYPGG-巴西拉再也(Pasir Gudang) | MY |
| GO DAU | 岘港 | VNDAD-岘港(Da-Nang/ Da Nang) | VN |
| LEAMCHABANG | 林查班 | THLCH-林查班(Laem Chabang) | TH |
| BAYUQUAN | 鲅鱼圈 | CNBAY-CNYQUAN-鲅鱼圈(Bayuanquan) | CN |
| PORT KLANG | 巴生港 | MYPKG-MYPKG-巴生港(Port Klang) | MY |
| MAOMING / MAO MING | 茂名 | CNMAG-CNMAG-茂名(Maoming) | CN |
| B.ABBAS / BANDAR ABBAS | 阿巴斯港 | IRBND-IRBND-阿巴斯港(Bandar Abbas) | IR |
| SHANGHAI | 上海港 | CNSHG-CNSHG-上海港(Shanghaigang) | CN |
| TAICANG | 太仓 | CNTAC-CNTAC-太仓(Taicang) | CN |
| FUJAIRAH | 富查伊拉 | AEFJR-AEFJR-富查伊拉(Fujairah) | AE |
| HONG KONG | 香港 | HKHKG-HKHKG-香港(Hong Kong) | HK |
| SOHAR | 索哈 | OMSOH-OMSOH-索哈(Sohar) | OM |
| ZHOUSHAN | 舟山 | CNZOS-CNZOS-舟山(Zhoushan) | CN |
| SINGAPORE | 新加坡 | SGSIN-SGSIN-新加坡(Singapore) | SG |
| DONG GUAN | 东莞 | CNDGC-CNDGC-东莞(Dongguan) | CN |
| YINGKOU / YINGKO | 营口 | CNYIK-CNYIK-营口(Yingkou) | CN |
| MERAK | 默拉克 | IDMRK-IDMRK-默拉克(孔雀岛)(MERAK) | ID |
| QUANZHOU | 泉州 | CNQAU-CNQAU-泉州(Quanzhou) | CN |
| JIANGYIN | 江阴 | CNJIA-江阴(Jiangyin)（**CNJYN 在 port_map.json 中不存在**） | CN |
| TAIXING | 泰兴 | CNTXI-CNTVG-泰兴(Taixinh) | CN |
| JINGJIANG | 靖江 | CNTSI-CNTSI-靖江(Jingjiang) | CN |
| BAHODOPI | 巴霍多皮 | **IDBAH-巴霍多皮(BAHODOPI)** | ID |
| BATANGAS | 八打雁 | PHBTG-PHBTG-八打雁(Batangas) | PH |
| DONG NAI | 同奈 | VNDNI-VNDNI-盖邻(Dong Nai) | VN |
| QINZHOU | 钦州 | CNQZH-CNQZH-钦州(Qinzhou) | CN |
| CHANGZHOU | 常州 | CNCZX-CNCZX-常州(Changzhou) | CN |
| KAOHSIUNG | 高雄 | TWKHH-TWKHH-高雄(Kaohsiung) | TW |
| TAICHUNG | 台中 | TWTXG-TWTXG-台中(Taichung) | TW |
| YANTAI | 烟台 | CNYNT-CNYNT-烟台(Yantai) | CN |
| RUGAO | 如皋 | CNRGG-CNRGG-如皋(Rugao) | CN |
| CHANGSHU | 常熟 | CNCGS-CNCGS-常熟(Changshu) | CN |
| NANSHA | 广州南沙 | CNNSA-CNNSA-广州南沙(Guangzhou Nansha) | CN |
| MASAN | 马山 | KRMAS-KRMAS-马山(Masan) | KR |
| SEPETIBA | 伊塔瓜伊/塞佩蒂巴 | **BRSPB**-伊塔瓜伊/塞佩蒂巴(SEPETIBA)（**不是BRSEP**） | BR |
| PORTOCEL | 蓬塔塞CEL | BRPCE-BRPCE-蓬塔塞CEL(Portocel) | BR |
| PARANAGUA | 巴拉那瓜 | BRPNG-BRPNG-巴拉那瓜(Paranagua) | BR |
| VITORIA | 维多利亚 | BRVIX-BRVIX-维多利亚(Vitoria) | BR |
| BASRAH | 巴士拉 | IQBSR（fallback: UNLOCODE构建，需人工确认） | IQ |
| PORT WALCOTT | 沃尔科特港 | AUPWA-沃尔科特港(PORT WALCOTT) | AU |
| MA JISHAN | 马迹山 | CNMJS（port_map无，手动） | CN |
| LONGKOU | 龙口 | CNLKU-龙口 | CN |
| NINGBO | 宁波 | CNNBO-CNNBO-宁波(Ningbo)（**已废弃**） | CN |
| SONGXIA | 松下 | CNSON-CNSON-松下(Songxia) | CN |
| LANSHAN | 岚山 | CNLSN-岚山１ | CN |

### 追加：实测天王之星 POC 文件发现缺失的港口（2026-05）

| 原始港口名 | 替代港口 | 替代代码 | 国家 |
|------------|----------|----------|------|
| WEDA | WEDA | IDWED-WEDA | ID |
| HONGKONG | 香港 | HKHKG-香港(Hong Kong) | HK |
| QINGDAO | 青岛港 | CNQDP-青岛港 | CN |
| SHENZHEN | 深圳宝安 | SZX-深圳宝安国际机场(Shenzhenbaoanguojijichang) | CN |
| DONGGUAN | 东莞 | CNGGU-东莞 | CN |
| LAEMCHABANG | 林查班 | THLCH-林查班(Laem Chabang) | TH |
| SIHANOUK VILLE | 西哈努克城 | KHKOS-西哈努克城(Sihanoukville) | KH |
| CEBU | 宿务 | PHCEB-宿务(Cebu) | PH |
| LIANYUNGANG | 连云港 | CNLYG-连云港(LIANYUNGANG) | CN |
| OBI ISLAND | 奥比岛 | IDOBI-奥比岛(OBI ISLAND) | ID |
| BINHAI | 滨海 | CNBHX-滨海(BINHAI) | CN |

### 追加：POC-country-map.json 需包含的英文国名（实测发现）

`poc-country-map.json` 需包含以下额外条目，nationality_map.json 不含英文国名：

```json
{
  "PHILIPPINE": "PH-菲律宾",
  "INDONESIA": "ID-印度尼西亚",
  "THAILAND": "TH-泰国",
  "CAMBODIA": "KH-柬埔寨",
  "VIETNAM": "VN-越南",
  "MALAYSIA": "MY-马来西亚",
  "SINGAPORE": "SG-新加坡",
  "JAPAN": "JP-日本",
  "KOREA": "KR-韩国",
  "BANGLADESH": "BD-孟加拉国",
  "PANAMA": "PA-巴拿马"
}
```

### ⚠️ `nationality_map.json` 不包含英文国名（POC 国家字段专用映射）

**问题**：`nationality_map.json` 的 key 是 2位字母码，value 是 `CC-中文名`。英文国名如 SINGAPORE、MALAYSIA、OMAN、UAE、IRAN 等无法直接查表。

**解法**：POC 国家字段需要独立映射表 `references/poc-country-map.json`：

```json
{
  "SINGAPORE": "SG-新加坡",
  "MALAYSIA":  "MY-马来西亚",
  "OMAN":      "OM-阿曼",
  "IRAN":      "IR-伊朗",
  "UAE":       "AE-阿联酋",
  "CHINA":     "CN-中国",
  "U.N.":      "UN-联合国",
  "SWEDEN":    "SE-瑞典",
  "NETHERLANDS": "NL-荷兰",
  "NIGERIA":   "NG-尼日利亚",
  "SOUTH KOREA": "KR-韩国",
  "KOREA":     "KR-韩国",
  "IRAQ":      "IQ-伊拉克",
  "BRAZIL":    "BR-巴西"
}
```

**Crew List 国籍字段额外坑点**：Crew List 中 `Nationality` 列常用英文全称（如 `CHINESE`、`VIETNAM`、`INDONESIA`），这些**不在 `nationality_map.json` 也可能不在 `poc-country-map.json`**。

必须建立 `NAT_FULL_TO_CODE` 硬编码表（每次新增船舶时追加）：
```python
NAT_FULL_TO_CODE = {
    "CHINESE":    "CN-中国",
    "VIETNAM":    "VN-越南",
    "INDONESIA":  "ID-印度尼西亚",
    "PHILIPPINE": "PH-菲律宾",
    # ...
}
```
查找顺序：`NAT_FULL_TO_CODE` → `nationality_map.json` → `poc-country-map.json` → 留空（标红）

### ⚠️ Crew List .xls 格式需要 xlrd

**`execute_code` 沙盒中 xlrd 不可用**：在 `execute_code` 环境中 `import xlrd` 会报 `ModuleNotFoundError: No module named 'xlrd'`。即使 skill 文档声明 `pip install xlrd`，沙盒也不会有。

**正确做法**：读取 .xls 文件必须在 `terminal()` 中用 `python3` 执行，或写脚本到 `~/` 目录后 `python3 ~/gen_xxx.py` 执行。

**`xlrd` 读取时的关键坑：**
1. `openpyxl.load_workbook(path)` 对 `.xls` 报 `InvalidFileException: openpyxl does not support the old .xls file format`
2. 必须先判断文件后缀，`.xls` 走 `xlrd.open_workbook()`，`.xlsx` 才走 `openpyxl`
3. xlrd 日期单元格是 Excel serial float，需用 `xlrd.xldate_as_datetime(v, wb.datemode)` 转换
4. POC 的 `.xls` 文件表头不在第0行（如 `PORT OF CALL LIST` sheet 的表头在 Row4），需扫描定位

```python
import xlrd
wb = xlrd.open_workbook("crew_list.xls")
ws = wb.sheet_by_index(0)
# 日期转换
dt = xlrd.xldate_as_datetime(cell_value, wb.datemode)
```

### ⚠️ Crew List .xls 列布局（实测 KANGSHUN 99）

KANGSHUN 99 (.xls, xlrd) 实测列索引（0-based）：

```
Index:  0     1                   3       4      5           6              7            8            9           10           11         12         13
        No    Name(中文\n英文)     Nat     Sex    Rank      DOB           POB          SB_No        SB_Exp       Country     JoinDate    JoinPlace   Passport   PP_Exp
```

| 字段 | 列索引 | 数据类型 | 示例 |
|------|--------|----------|------|
| 序号 | 0 | int | 1、 |
| 姓名 | 1 | str 含 `\n` | `卜庆丹\nBU QINGDAN` |
| 国籍 | 3 | str | `China` |
| 性别 | 4 | str | `M` |
| 职务 | 5 | str | `CAPT` |
| 出生日期 | 6 | Excel float 或 `yyyy.mm.dd` 字符串 | `19811213` |
| 出生地点 | 7 | str | `LIAO NING` |
| 海员证号 | 8 | str | `A90239549` |
| 海员证有效期 | 9 | str | `2027.07.18` |
| 国家 | 10 | str | `China` |
| 登船日期 | 11 | `yyyy.mm.dd` 字符串 | `2026.05.11` |
| 登船地点 | 12 | str | `QIN ZHOU` |
| **护照号** | **13** | str | `ER3819560` |
| 护照有效期 | 14 | str | `2036.03.23` |

**姓名解析**：`col1` 含 `中文\n英文` 格式，用 `\n` 分割；中文名为第一段，英文名为第二段。中文名为空表示是外国船员。

**外国船员证件规则**：
- 证件类型 = `14-护照`
- 证件号码 = **col13（护照号）**，不是 col8（海员证号）
- 之前版本错误地把海员证号填入护照号字段，实测数据中护照号在 col13 列

**日期解析**：
- 出生日期（col6）：字符串 `yyyy.mm.dd` → 直接去除 `.` → `YYYYMMDD`；Excel float → `xlrd.xldate_as_datetime()` 转换
- 登船日期（col11）：字符串 `yyyy.mm.dd` → 直接去除 `.`

### ⚠️ Crew List 列布局（实测天王之星 .xls）

天王之星 Crew List `.xls` 的实际列索引（0-based，与舱单打印版不同）：

```
Index:  0      1        2        3       4      5           6              7           8            9        10          11       12       13
        No.   Name    Rank    Nationality  Sex  BirthDate  BirthPlace   SeamanBook  Passport    Country   JoinDate  JoinPlace  (空)   (空)
```

| 字段 | 列索引 | 数据类型 | 示例 |
|------|--------|----------|------|
| 序号 | 0 | int | 1 |
| 姓名 | 1 | str | `张三` 或 `张三 ZHANG SAN` |
| 职务 | 2 | str | `MASTER` / `C/O` |
| 国籍 | 3 | str | `CN` |
| 性别 | 4 | str | `M` / `F` |
| **出生日期** | **6** | **Excel float（xlrd serial）** | 峰值为 25569 加天数 |
| 出生地点 | 7 | str | `SHANDONG` |
| **海员证号** | **8** | **str** | `A90395125` |
| 护照号 | 10 | str | `ER3358421` |
| 国家 | 11 | str | `CHINA` |
| **登船日期** | **13** | **str `YYYY/MM/DD`** | `2024/01/15` |
| **登船地点** | **14** | **str（原始城市名）** | `XIAMEN` / `WEIFANG` |

**⚠️ 日期解析规则：**
- 出生日期（col6）：Excel float → `xlrd.xldate_as_datetime(v, wb.datemode)` → `strftime("%Y%m%d")`
- 登船日期（col13）：字符串 `YYYY/MM/DD` → 直接去除斜杠 → `YYYYMMDD`

**⚠️ 中国船员登船口岸需单独映射（PORT_FALLBACK）：**
Crew List 的登船口岸字段是原始城市名（如 `XIAMEN`、`WEIFANG`、`QINGDAO`），不是 UNLOCODE，且可能有换行符污染（如 `ZHANJIANG,\nCHINA`）。需用 `PORT_FALLBACK` 映射：

```python
PORT_FALLBACK = {
    "XIAMEN": "CNXAM-厦门",
    "WEIFANG": "CNWEF-潍坊",
    "QINGDAO": "CNQDP-青岛港",
    # 追加：
    "ZHANJIANG": "CNZNG-湛江港(Zhanjianggang)",   # 原始值可能是 "ZHANJIANG,\nCHINA"
    "YANGPU":    "CNYPG-洋浦(Yangpu)",
    "ZHANGJIAGANG": "CNZJG-张家港(Zhangjiagang)",
    "SONGXIA":   "CNSON-松下(Songxia)",
    # 更多城市名...
}
```

### ⚠️ POC 日期格式

POC 文件中日期格式不统一：
- Excel `.xls` 中：`d/m/yyyy`（如 `20/4/2026`）、`d/m/yy`（如 `20/4/26`）、**`yyyy.mm.dd`（如 `2026.05.21`，实测天王之星）**
- PDF 中：`dd-Mon-yyyy`（如 `09-May-2026`）或 `dd-mm-yy & HH-MM-SS`（如 `26-03-29 & 07-00-00`）

解析时需要同时支持多种格式：
```python
for fmt in ["%Y.%m.%d", "%d-%b-%Y", "%d-%m-%y", "%d-%m-%Y", "%d/%m/%Y", "%d/%m/%y"]:
    try: return datetime.strptime(s, fmt).strftime("%Y/%m/%d")
    except: pass
```

### ⚠️ 港口子串匹配冲突（port_map 误匹配）

**症状**：`SHA` 匹配到 `SHA-上海虹桥国际机场` 而非 `CNSHG-上海港`；`CNZOS` / `CNYZO` / `CNSHG` 在 port_map 中不存在。

**解法**：每个船舶单独建立 `PORT_MANUAL` dict，hardcode 所有港口代码，不用子串匹配。

**GREEN MUNGUBA PORT_MANUAL**：
```python
PORT_MANUAL = {
    "ZHOUSHAN":   ("CNZOS", "舟山(Zhoushan)"),
    "YANGZHONG":  ("CNYZO", "扬州、镇江(Yangzhong, Zhenjiang)"),
    "QINGDAO":    ("CNQDP", "青岛港(Qingdao)"),
    "TAICANG":    ("CNTAC", "太仓(Taicang)"),
    "SHANGHAI":   ("CNSHG", "上海港(Shanghai)"),
    "NAPLES":     ("ITNAP", "那不勒斯(NAPLES)"),
    "MONFALCONE": ("ITMFA", "蒙法尔科内(MONFALCONE)"),
    "GIBRALTAR":  ("GIBGI", "直布罗陀(GIBRALTAR)"),
    "ITAQUI":     ("BRITQ", "伊塔基(ITAQUI)"),
    "SINGAPORE":  ("SGSIN", "新加坡(Singapore)"),
}
```

### ⚠️ Crew List .xls 表头定位

**症状**：Sheet3 出现假港口行（如 `NALUD`、`卢德立茨`），或第一个港口不是最新的。

**根因**：原代码匹配 "Voyage No." 行（Row1 含 `NO.` + `Port`）作为表头，但该行是船舶信息标题行，不是真正的港口数据表头。

**正确做法**：表头行必须同时满足两个条件：
1. 包含 `NO.` 列标题（大写）
2. 包含 `NAME OF PORT` 或 `PORT OF CALL` 列标题（大写）

```python
for i in range(min(20, ws.nrows)):
    row = [str(c.value).upper() for c in ws.row(i)]
    has_no = any("NO." in h for h in row)
    has_port = any("NAME OF PORT" in h or "PORT OF CALL" in h for h in row)
    if has_no and has_port:
        header_idx = i
        break
```

**⚠️ POC 港口顺序**：数据从新到旧（第1行=最新进港），不要反转。

**⚠️ POC Excel 列索引（实测天王之星 `PORT OF CALL LIST` sheet）：**
```
Index:  0       1          2           3              4              5          6    7                    8                         9              10
        No      Name of     Country     UNLOCODE      arrival date   departure   sec   Security Threats   Especial response         Ship/Ship      Remarks
                                 (空)      (空)         dd/mm/yy      date        level                  and Incidents          Activity
```
⚠️ `arrival date` 在 **index=4**，`departure date` 在 **index=5**。

⚠️ POC `.xls` 的 `PORT OF CALL LIST` sheet 名称含空格，实测 sheets: `['Sheet4', 'PORT OF CALL LIST']`，需要用 `wb.sheet_by_name("PORT OF CALL LIST")` 精确读取。

调试时打印原始行全部字段可快速定位：
```python
for i in range(ws.nrows):
    vals = [(j, c.value) for j, c in enumerate(ws.row(i)) if c.value not in (None, '', 0.0)]
    if vals:
        print(f"Row{i}: {vals}")
```

⚠️ POC 最后一行（Row16 = 0-indexed）通常是船长签名行，以 `Date and signature by master` 开头，解析时 country 字段为空，需跳过。

### ⚠️ Crew List 奇偶行结构：join_place 在子行 col9（实测 MILLIE）

部分 Crew List（实测 MILLIE）采用**奇偶行结构**：
- 奇数行（如 Row9, 11, 13...）：主数据（姓名、职务、出生日期、海员证号、登船日期）
- 紧跟的偶数行（如 Row10, 12, 14...）：补充数据（出生地点、护照有效期、海员证有效期、**登船地点**）

```python
# 正确：
row_main = ws_crew.row(i)       # 奇数行
row_sub  = ws_crew.row(i+1)     # 偶数行
join_place = row_sub[9].value   # ✅ 子行 col9 = 登船地点

# 错误：
join_place = row_main[9].value  # ❌ 主行 col9 = 登船日期，不是登船地点
```

**实测 MILLIE 列布局：**

| | 主行 col6 | 主行 col8 | 主行 col9 | 子行 col6 | 子行 col8 | 子行 col9 |
|--|-----------|-----------|-----------|-----------|-----------|-----------|
| 字段 | 出生日期 | 海员证号 | 登船日期 | 出生地点 | 海员证有效期 | **登船地点** |

**⚠️ `单证录入核心.py` 脚本港口匹配 bug（慎用）

**症状**：脚本输出的海事活动中，几乎所有港口都变成了 `HITACHINAKA`（日本常总），完全错误。

**根因**：脚本的 `extract_port_code()` 子串匹配逻辑过于宽松，且 `port_map.json` 中大部分常用港口缺失，导致短码被长字符串错误匹配。

**建议**：Agent 应使用 skill 中硬编码的 `PORT_FALLBACK` 字典 + `port_map.json` 精确匹配，手动处理所有 POC 港口。

### ⚠️ 扫描件 PDF 的 OCR 处理流程

当 PORT CALL LIST 是**扫描件 PDF**（无文字层，pdfplumber/pdftotext 均无法提取）时的处理顺序：

1. **先用 pypdfium2 渲染 PDF → PNG**（最可靠）：
   ```python
   import pypdfium2 as pdfium
   pdf = pdfium.PdfDocument("input.pdf")
   page = pdf[0]
   pil = page.render(scale=2).to_pil()
   pil.save("output.png")
   ```

2. **本地 Tesseract OCR**（优先 `eng`，港口名为英文）：
   ```bash
   tesseract output.png stdout -l eng --psm 6
   ```
   注意：Tesseract chi_sim 训练数据容易损坏（`chi_sim.traineddata` 正常约 40MB，损坏时仅 8KB）。macOS 可用 `brew install tesseract-lang` 重新安装。

3. **在线 OCR**（本地失败时）：ocr.space 支持中文，选择 `ChineseSimplified` 语言，上传图片 URL 或文件。但在线工具结果获取有时不稳定，**最快速方案是直接请用户口述或发送可复制的文字版本**。

4. **用户直接提供数据**：当 OCR 耗时过长时，直接请用户提供港口名称、日期、国家列表，这是最高效的路径。

**关键教训**：处理新的 PORT CALL LIST 文件时，优先判断是**扫描件 PDF** 还是**文字版 xlsx/可复制 PDF**。扫描件直接走渲染+OCR流程，不要在 pdfplumber/pdftotext 浪费尝试时间。

---

### 输出与发送

**⚠️ 发送时不要附详细表格！** 用户明确要求：
- 只发 oneTableDeclaration_*.xlsm 文件
- 不附详细港口列表/核对信息（港口核对结论可简单说一句，但不要逐行列举）
- 如有代码修正，简单说明即可（如"AUPWL→AUPWA已修正"）

### xlsm 模板来源（强制规则）

**用户上传了自己的 oneTableDeclarationTemplate.xlsm，每次必须使用该模板生成输出：**

1. 模板路径优先级：
   - `~/Desktop/oneTableDeclarationTemplate.xlsm` → 存在则用
   - `~/.hermes/skills/ship-document-converter/output/oneTableDeclarationTemplate.xlsm` → 用户上传的最新版本

2. 生成步骤：
   ```python
   wb = openpyxl.load_workbook(TEMPLATE_PATH, keep_vba=True)
   # 清空第3-25行（示例行）
   for ws in [ws1, ws2, ws3]:
       for r in range(3, 30):
           for c in range(1, 20):
               ws.cell(r, c).value = None
   # 写入数据（从Row3起，Row2=说明行不得删除）
   OUT = f'output/oneTableDeclaration_{船名}_{日期}.xlsm'
   wb.save(OUT)
   ```

3. **不要**先保存成 xlsx 再转 xlsm——这样会导致文件损坏无法打开。必须直接用 `keep_vba=True` 加载模板后写入并保存。

4. 发送目标：飞书 DM `oc_9d8f4df4139fb63513d74ee2ef17df8d`（不要只发到群）

5. **飞书云文档无法上传 Excel**：Feishu Docx API 只支持文字内容，无法上传 xlsm/xlsx 附件。发送文件必须走 IM 文件消息接口。

### 标准输出文件

**模板路径**：`~/.hermes/skills/ship-document-converter/output/oneTableDeclarationTemplate.xlsm`（用户上传版本，优先）；`~/Desktop/oneTableDeclarationTemplate.xlsm`（备用）

**生成文件**：`oneTableDeclaration_{船名}_{日期}.xlsm`（只生成 xlsm，不生成中间 xlsx）

**发送方式**：必须通过飞书 IM 文件消息发送，不发 MEDIA: 链接：
```python
# 1. 获取 token
resp = requests.post("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": APP_ID, "app_secret": APP_SECRET}, timeout=10)
token = resp.json()["tenant_access_token"]

# 2. 上传文件获取 file_key
with open(file_path, "rb") as f:
    m = MultipartEncoder({
        "file_type": "xls",
        "file_name": file_name,
        "file": (file_name, f, "application/vnd.ms-excel"),
    })
    h = {"Authorization": f"Bearer {token}", "Content-Type": m.content_type}
    upload_resp = requests.post("https://open.feishu.cn/open-apis/im/v1/files",
        headers=h, data=m, timeout=60)
file_key = upload_resp.json()["data"]["file_key"]

# 3. 发送文件消息
msg_data = {
    "receive_id": "ou_a498829d1d1678ed8880cf17853f0274",
    "msg_type": "file",
    "content": json.dumps({"file_key": file_key, "file_name": file_name})
}
send_resp = requests.post(
    "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id",
    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    json=msg_data, timeout=10
)
```

**⚠️ POC-only 任务**：用户说"仅处理这个文件"且只有 Port of Call 时，只填 Sheet3（海事船岸活动信息），Sheet1 和 Sheet2 留空。

### v3 模板行偏移规则（必须遵守）

v3 模板（`单证录入标准格式_v3.xlsm`）与旧 v2 模板行规则不同：

| Sheet | Row1 | Row2 | Row3=数据#1 |
|-------|------|------|------|
| 船上非旅客人员清单 | 表头 | **说明行**（不得删除） | 第1号船员 |
| 船上非旅客人员物品清单 | 表头 | **说明行**（不得删除） | 第1条物品 |
| 海事船岸活动信息 | 表头 | **说明行**（不得删除） | 第1个港口 |

**公式**：`target_row = source_row - 2 + 3`（源xlsx从Row2→v3从Row3起，偏移量=+1）

> ⚠️ Row2 是说明行，**不得覆盖，不得删除**。源数据 Row2 → 目标 Row3，源数据 Row3 → 目标 Row4...

### v3 模板结构（17个Sheet）

```
船上非旅客人员清单 | 旅客清单 | 供退物料清单 | 船上非旅客人员物品清单
船用物品清单 | 前十港信息 | 危险品信息 | 船舶证书信息 | 压舱水详细信息
海事船岸活动信息 | 沿海空箱信息 | 压舱水报告单信息 | 压舱水装载信息
压舱水更换信息表信息 | 压舱水排放信息表信息 | 随船人员清单 | 参数
```

**本次录入只填写前3个 Sheet**，其余留空。

### ⚠️ 模板 Sheet 名称必须先读后用（已踩坑）

**症状**：`wb["Sheet1"]` → `KeyError: Worksheet Sheet1 does not exist.`

**根因**：skill 文档长期误记为 `Sheet1/Sheet2/Sheet3`，但实际模板名称是：

| 实际 Sheet 名称 | 对应内容 |
|----------------|---------|
| `船上非旅客人员清单` | Sheet1 船员名单 |
| `船上非旅客人员物品清单` | Sheet2 物品清单 |
| `海事船岸活动信息` | Sheet3 海事活动 |

**正确做法**：先读 `wb.sheetnames` 确认，再按实际名称访问：
```python
wb = openpyxl.load_workbook(template_path)
print("模板 sheets:", wb.sheetnames)
ws1 = wb["船上非旅客人员清单"]
ws2 = wb["船上非旅客人员物品清单"]
ws3 = wb["海事船岸活动信息"]
```

### ⚠️ CNQDG ≠ CNQDP（港口代码实测区分）

`port_map.json` 中两者独立存在：
- `CNQDP` → `CNQDP-青岛港`
- `CNQDG` → `CNQDG-青岛大港`

POC 文件给的代码决定用哪个，不能混用。

### oneTableDeclaration xlsm 直接写入流程（强制）

**正确方式：直接加载模板写入，不要走 xlsx 中转！**

```python
import openpyxl, random, datetime

TPL = 'output/oneTableDeclarationTemplate.xlsm'  # 用户上传的模板
wb = openpyxl.load_workbook(TPL, keep_vba=True)
ws1 = wb['船上非旅客人员清单']
ws2 = wb['船上非旅客人员物品清单']
ws3 = wb['海事船岸活动信息']

# 清空第3-25行（模板示例行）
for ws in [ws1, ws2, ws3]:
    for r in range(3, 30):
        for c in range(1, 20):
            ws.cell(r, c).value = None

# 写入数据（Row3 = 第一条数据，Row2 = 说明行不得覆盖）
# ... 写入 ws1 / ws2 / ws3 ...

OUT = f'output/oneTableDeclaration_{船名}_{日期}.xlsm'
wb.save(OUT)  # 直接保存，不要先保存xlsx再转！
```

**错误方式（会导致文件损坏无法打开）：**
1. 先保存成 `单证录入_*.xlsx`
2. 再复制到 xlsm 模板 → 文件损坏

### 发送给用户

**必须走飞书 IM 文件消息接口**，不能用 Feishu Docx API（后者不支持上传 Excel 附件）。

发送目标：飞书 DM `oc_9d8f4df4139fb63513d74ee2ef17df8d`

**两步上传法（xlsm/xlsx file_type 均用 `xls`）：**
```python
import requests
from requests_toolbelt import MultipartEncoder
import json, os

APP_ID = "cli_a952c98ec13a9bca"
APP_SECRET = "<FEISHU_APP_SECRET>"

resp = requests.post(
    "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": APP_ID, "app_secret": APP_SECRET}, timeout=10)
token = resp.json()["tenant_access_token"]

file_path = f"output/oneTableDeclaration_{船名}_{日期}.xlsm"
with open(file_path, "rb") as f:
    m = MultipartEncoder({
        "file_type": "xls",
        "file_name": os.path.basename(file_path),
        "file": (os.path.basename(file_path), f, "application/vnd.ms-excel"),
    })
    h2 = {"Authorization": f"Bearer {token}", "Content-Type": m.content_type}
    upload_resp = requests.post("https://open.feishu.cn/open-apis/im/v1/files",
        headers=h2, data=m, timeout=60)

file_key = upload_resp.json()["data"]["file_key"]
msg_data = {
    "receive_id": "oc_9d8f4df4139fb63513d74ee2ef17df8d",
    "msg_type": "file",
    "content": json.dumps({"file_key": file_key, "file_name": os.path.basename(file_path)})
}
requests.post(
    "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    json=msg_data, timeout=10)
```

**只发文件，不附详细表格。**

### 格式刷要点

写入 oneTableDeclaration 后，三个 sheet 的数据行格式可能不一致：
- **船员清单**：col4(职务) col5(国籍) → `italic=True, color=FFFF0000, size=10`（下拉联动字段）
- **物品清单**：全列统一 `italic=False, color=theme=1, size=11`
- **海事活动**：col1 左对齐，col2-8 右对齐

格式参考行永远是源文件 Row3（第一条数据）。用 `copy(src_cell.font)` 等完整复制样式，不要逐字段重建。

### 发送给用户

默认通过飞书发送，目标 DM：`oc_9d8f4df4139fb63513d74ee2ef17df8d`（不要只发到群）

**只发文件，不附详细表格**：
```python
send_message(target="feishu:ou_a498829d1d1678ed8880cf17853f0274",
             message="MEDIA:/path/to/oneTableDeclaration_{船名}_{日期}.xlsm")
```

**Feishu 发送失败 `TLS CA bundle not found`**：
- 错误信息：`Could not find a suitable TLS CA certificate bundle, invalid path: .../certifi/cacert.pem`
- 根因：hermes-agent venv 的 python3.12 环境缺少 certifi 包的 cacert.pem 文件（python3.11 有，python3.12 没有）
- 解法：`mkdir -p ~/.hermes/hermes-agent/venv/lib/python3.12/site-packages/certifi/ && cp ~/.hermes/hermes-agent/venv/lib/python3.11/site-packages/certifi/cacert.pem ~/.hermes/hermes-agent/venv/lib/python3.12/site-packages/certifi/`

- PDF 支持需根据实际布局调整
- 护照有效期/适任证书留空
- 中国船员源数据若为纯拼音且无中文，姓名保留拼音原样（无法反向还原中文）
