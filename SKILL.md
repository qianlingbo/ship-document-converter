---
name: ship-document-converter
description: 将船舶 IMO Crew List + Port of Call 转换为海事局标准录入格式
category: productivity
tags: [maritime, 单证录入, maritime-declaration]
version: 2.5.0
---

# 单证录入技能

> 将船舶 IMO Crew List + Port of Call 转换为海事局标准录入格式

**版本**: v2.5.0 | **Python**: 3.8+ | **依赖**: `openpyxl`, `xlrd`

## 快速开始

```bash
pip install openpyxl xlrd
python3 scripts/单证录入核心.py input/crew_list.xlsx [port_of_call.xlsx] [输出名]
```

## 目录结构

```
.
├── scripts/
│   ├── 单证录入核心.py           # 核心脚本
│   └── poc录入_MINHUA9.py       # POC专用脚本（参考用）
├── templates/
│   └── 单证录入标准格式_v2.xlsx  # 输出模板（6个sheet）
├── references/
│   ├── nationality_map.json     # 国籍代码 (248条) — 参数A
│   ├── duty_map.json            # 职务代码 (12条) — 参数B
│   ├── port_map.json            # 港口代码 (1956条) — 参数E
│   ├── poc-country-map.json     # POC国家字段专用映射（英文国名→代码-中文名）
│   ├── zyhy-jinqu-layout.md    # 中远海运津渠轮 Crew List 布局笔记
│   └── conversion-rules.md      # 转换规则详解（含所有坑点）
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

### 5. 进港时间 / 离港时间

**强制格式**：`yyyy/MM/dd HH:mm:ss`

| 字段 | 日期来源 | 时间规则 | 示例 |
|------|----------|----------|------|
| 进港时间 | Port of Call 的 Arrival 日期 | 随机 `00:00:00` ~ `11:59:59` | `2024/01/15 08:23:47` |
| 离港时间 | Port of Call 的 Departure 日期 | 随机 `12:00:00` ~ `23:59:59` | `2024/01/16 18:45:12` |

**禁止格式**：`2024-01-15`、`20240115`、`2024/1/15`、`Jan 15, 2024` 等。

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

### 9. 证件类型 & 证件号码（重要修正）

| 国籍 | 证件类型 | 证件号码来源 |
|------|----------|--------------|
| 中国（CN） | `17-海员证` | **Seaman's Book（第13列）** |
| 其他 | `14-普通护照` | Passport（第11列） |

**中国船员证件号码取值规则（用户明确纠正）：**
- Crew List 表中，第11列是 Passport（普通护照），第13列是 Seaman's Book（海员证）
- **中国船员必须取第13列（海员证号码）**，不得取第11列
- 示例：`高峰` 的 Seaman's Book = `A90395125`，Passport = `ER3358421` → 应输出 `A90395125`

---

## 转换规则总览

### Sheet 1: 船上非旅客人员清单（16列）

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

每人固定一条记录：物品类型 `0100`，物品名称 `计算机`，数量 `1`，单位 `001`。

### Sheet 3: 海事船岸活动信息（8列）

| 列 | 字段 | 规则 |
|----|------|------|
| A | 序号 | 从1开始递增 |
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

### Port of Call 港口 fallback 速查表

以下港口在 `port_map.json` 中不存在，备用映射已验证可用：

| 原始港口名 | 替代港口 | 替代代码 | 国家 |
|------------|----------|----------|------|
| OPEN SEA | 公海 | THS-公海 | UN |
| CHENJIIAGANG | 陈家港 | CNCIG-陈家港(CHENJIAGANG) | CN |
| KENDARI | 乌戎潘当 | IDUPG-乌戎潘当(Ujung Pandang) | ID |
| MORMUGAO | 哈迪亚 | INHDA-哈迪亚(HALDIA) | IN |
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
| JIANGYIN | 江阴 | CNJYN-CNJYN-江阴(Jiangyin) | CN |
| TAIXING | 泰兴 | CNTXI-CNTVG-泰兴(Taixinh) | CN |
| JINGJIANG | 靖江 | CNTSI-CNTSI-靖江(Jingjiang) | CN |
| BAHODOPI | 巴霍尔迪皮 | IDBDJ-IDBDJ-巴霍尔迪皮(Bahodopi) | ID |
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
| PORTOCEL | 蓬塔塞CEL | BRPCE-BRPCE-蓬塔塞CEL(Portocel) | BR |
| PARANAGUA | 巴拉那瓜 | BRPNG-BRPNG-巴拉那瓜(Paranagua) | BR |
| VITORIA | 维多利亚 | BRVIX-BRVIX-维多利亚(Vitoria) | BR |
| BASRAH | 巴士拉 | IQBSR（fallback: UNLOCODE构建，需人工确认） | IQ |
| NINGBO | 宁波 | CNNBO-CNNBO-宁波(Ningbo)（**非 CNNGB**，后者不存在） | CN |

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

### ⚠️ Crew List .xls 格式需要 xlrd

Crew List 旧版 `.xls` 文件 openpyxl 无法读取，必须用 `xlrd`：

```bash
pip install xlrd
```

```python
import xlrd
xls = xlrd.open_workbook("crew_list.xls")
ws = xls.sheet_by_index(0)
# 日期单元格是 Excel serial float，需转换：
dt = xlrd.xldate_as_datetime(cell_value, xls.datemode)
```

### ⚠️ POC 日期格式

POC 文件中日期格式不统一：
- PDF 中：`dd-Mon-yyyy`（如 `09-May-2026`）或 `dd-mm-yy & HH-MM-SS`（如 `26-03-29 & 07-00-00`）
- Excel 中：`d/m/yyyy`（如 `20/4/2026`）、`d/m/yy`（如 `20/4/26`）

解析时需要同时支持多种格式：
```python
for fmt in ["%d-%b-%Y", "%d-%m-%y", "%d-%m-%Y", "%d/%m/%Y", "%d/%m/%y"]:
    try: return datetime.strptime(s, fmt)
    except: pass
```

### ⚠️ `单证录入核心.py` 脚本港口匹配 bug（慎用）

**症状**：脚本输出的海事活动中，几乎所有港口都变成了 `HITACHINAKA`（日本常总），完全错误。

**根因**：脚本的 `extract_port_code()` 子串匹配逻辑过于宽松，且 `port_map.json` 中大部分常用港口缺失，导致短码被长字符串错误匹配。

**建议**：Agent 应使用 skill 中硬编码的 `PORT_FALLBACK` 字典 + `port_map.json` 精确匹配，手动处理所有 POC 港口。

---

## 已知局限

- PDF 支持需根据实际布局调整
- 护照有效期/适任证书留空
- 中国船员源数据若为纯拼音且无中文，姓名保留拼音原样（无法反向还原中文）
