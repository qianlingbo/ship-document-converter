# 单证录入转换规则详解

> 本文件是 `SKILL.md` 的详细参考文档，包含完整的坑点说明和技术细节。

## 1. XLS vs XLSX 格式差异

输入文件可能是老格式 **Excel 97-2003 (.xls)** 而非 .xlsx，处理方式完全不同：

| 格式 | 库 | 日期处理 |
|------|-----|---------|
| .xlsx | `openpyxl` | 字符串或数字自动识别 |
| .xls | `xlrd` | **所有值都是数字**（日期=Excel序列号；文本=unicode字符串） |

### XLS 日期读取

```python
import xlrd
from datetime import datetime, timedelta

wb = xlrd.open_workbook(file, encoding_override='utf-8')
ws = wb.sheet_by_index(0)
serial = ws.cell_value(row, col)  # e.g. 45789.0
date_str = (datetime(1899, 12, 30) + timedelta(days=int(serial))).strftime("%Y%m%d")
```

> ⚠️ 基准是 `1899-12-30`（不是 1900-01-01）！Excel 的序列号有个历史 bug，多算了 1900 年 2 月 29 日。

### XLS 文本编码

`xlrd` 内部使用 Unicode，需要 `encoding_override='utf-8'` 或 `'cp1252'`。

---

## 2. Excel 列布局

Crew List Excel 有两种变体，列布局完全不同：

### 变体A — 有序号列（标准）

`[0]=No, [1]=Name, [2]=Rank, [3]=Sex, [4]=Nat, [5]=Birth, [6]=SeaBook, [7]=Passport, [8]=JoinDate`

### 变体B — 无序号列（本项目遇到）

`[0]=No, [1]=Name, [2]=Rank, [3]=Sex, [4]=Nat, [5]=Birth(mixed), [6]=SeaBook(mixed), [7]=Passport(mixed), [8]=JoinDate(mixed)`

判断方法：检查 `row[0]` 是否为数字（序号）。变体B 的核心特点是：
- `row[5]` = 出生 `"DD/Mon/YYYY PLACE"`（混合字段）
- `row[6]` = 海员证 `"DD/Mon/YYYY NUMBER"`（混合字段）
- `row[7]` = 护照 `"DD/Mon/YYYY NUMBER"`（混合字段）
- `row[8]` = 登船日期+地点 `"DD/Mon/YYYY PLACE"`（混合字段）

---

## 3. 混合字段解析

原始数据中多个字段是"日期 + 内容"的混合格式，需要用正则分离：

- **出生**: `"28/Nov/1986 SHANDONG"` → 日期=`28/Nov/1986`，地点=`SHANDONG`
- **海员证/护照**: `"28/Jan/2027 A90194049"` → 日期=`28/Jan/2027`，号码=`A90194049`
- **登船**: `"03/Dec/2025 JINZHOU,CHINA"` → 日期=`03/Dec/2025`，地点=`JINZHOU,CHINA`

注意：月份缩写（Jan, Feb, Nov, Dec 等）需要 `%b` 格式符解析。偶尔有畸形格式如 `"05/ Jul/1998"`（日期与月份间有空格），需先 normalize。

---

## 4. 证件号码提取

不要直接用护照/海员证的整段值！原始格式是：

```
"10/Aug/2027 EA9623314"
```

完整值包含日期前缀，证件号码是后半部分。需用 `extract_cert_number()` 分离。

---

## 5. 姓名：中国 vs 外国

- **中国船员**（国籍代码 CN）：姓名只取中文部分，例如 `"LI HAIBIN  李海宾"` → `李海宾`
- **外国船员**（非 CN）：姓名只取英文大写部分，例如 `"NGUYEN VAN MANH"` → `NGUYEN VAN MANH`

---

## 6. 日期格式容错

`normalize_date()` 支持格式：
- `%Y-%m-%d`, `%Y/%m/%d`, `%Y%m%d`
- `%d/%m/%Y`, `%d-%m-%Y`
- **`%d/%b/%Y`**（重要！支持 Jan/Feb/Mar 等月份缩写）
- `%d %b %Y`, `%b %d, %Y`

月份缩写日期前可能有空格，如 `"05/ Jul/1998"`，`parse_mixed_field_date_place()` 会自动清理。

---

## 7. 登船口岸映射

地点输入可能是 `"JINZHOU,CHINA"`、`"TAICANG"`、`"SHANGHAI"` 等，需要与 `port_map.json` 模糊匹配。匹配成功返回格式如 `CNJNZ-锦州(Jinzhou)`。

### 子串匹配陷阱

- `SHA`（上海虹桥机场代码）是 `ZHOUSHAN`（舟山）的子串！导致 `ZHOUSHAN` 错误匹配到 `SHA-上海虹桥国际机场`
- `SHA` 也是 `LANSHAN`（岚山）的子串，同样导致错误匹配
- **根因**：`port_map.json` 中的上海代码是 `SHA/HGH`（机场三字码），与 `ZHOUSHAN` 拼音有字符重叠

### 修复方案

在 `match_port()` 中添加 `SPECIAL_PORT_OVERRIDE` 手动映射表：

```python
SPECIAL_PORT_OVERRIDE = {
    # 注意：LANSHAN 的正确代码是 CNLSN（无空格），不是 CNLSH
    # port_map.json 中对应 "CNLSN-岚山１"（带全角数字１）
    "ZHOUSHAN": "CNZOS-舟山(Zhoushan)",
    "LANSHAN": "CNLSN-岚山１",
    "YANTIAN": "CNYTN-盐田(Yantian)",
    "SHEKOU": "CNSHK-蛇口(Shekou)",
    "XIAMEN": "CNXMN-厦门(Xiamen)",
    "QINGDAO": "CNTAO-青岛(Qingdao)",
    "NINGBO": "CNNBO-CNNBO-宁波(Ningbo)",
    "TIANJIN": "CNTXG-天津(Tianjin)",
    "DALIAN": "CNDLC-大连(Dalian)",
    "HONGKONG": "CNHKG-香港(Hongkong)",
    "BUSAN": "KRPUS-釜山(Busan)",
    "SINGAPORE": "SGSIN-新加坡(Singapore)",
    "TOKYO": "JPTYO-东京(Tokyo)",
    "YOKOHAMA": "JPYOK-横滨(Yokohama)",
}

def match_port(val):
    if not val: return None
    v = str(val).strip().upper().replace(",", " ").replace(".", " ").strip()
    # 去掉国家后缀
    for suffix in ["CHINA", "CN", "PRC", "AUSTRALIA", "AU", "US", "USA", "UK", "GB"]:
        if v.endswith(suffix): v = v[:-len(suffix)].strip()
    if len(v) < 4: return None

    # 0. 特殊手动映射（优先）
    for key, port_code in SPECIAL_PORT_OVERRIDE.items():
        if key in v or v in key: return port_code

    # 1. exact
    for code, full in PORT_MAP.items():
        if v == code.upper() or v == full.upper(): return full
        code_key = code.split("-")[0].upper()
        if v == code_key: return full

    # 2. 子串匹配（长度>=5，避免短码误匹配）
    if len(v) >= 5:
        for code, full in PORT_MAP.items():
            code_key = code.split("-")[0].upper()
            if len(code_key) >= 5 and (code_key in v or v in code_key): return full
        for code, full in PORT_MAP.items():
            full_key = full.split("-")[0].upper()
            if len(full_key) >= 5 and (full_key in v or v in full_key): return full
    return None
```

---

## 8. 职务代码映射（易遗漏）

以下职务代码在 `duty_map.json` 中缺失，必须在代码中硬编码：

| 代码 | 含义 | 映射到 |
|------|------|--------|
| `CARP` | 木匠/Carpenter | `65-值班机工` |
| `AB` | Able Seaman | `56-高级值班水手` |
| `D/CDT`, `CDT` | Deck Cadet | `65-值班机工` |
| `E/CDT`, `E/C` | Engine Cadet | `65-值班机工` |
| `FTR` | Fitter | `65-值班机工` |
| `OLR` | Oiler | `66-高级值班机工` |
| `C/CK`, `CHIEF COOK` | Chief Cook | `65-值班机工` |
| `MTR` | Motorman | `65-值班机工` |

### 职务 fallback 规则

3个 `56-高级值班水手` + 3个 `66-高级值班机工`，其余按角色分配。

---

## 9. 物品清单格式

原始 IMO 船员清单通常不含物品数据。如果源文件没有物品列，固定每人登记一条：

| 列 | 字段 | 值 |
|----|------|-----|
| 1 | 序号 | 船员序号 |
| 2 | 证件类型 | `17-海员证`（中国）或 `14-普通护照`（外国） |
| 3 | 证件号码 | 海员证号（中国）或护照号（外国） |
| 4 | 物品类型 | `0100` |
| 5 | 物品名称 | `计算机` |
| 6 | 物品数量 | `1` |
| 7 | 数量单位 | `001` |
| 8 | 船员携带观赏植物数量 | 空 |
| 9 | 船员携带宠物数量 | 空 |
| 10 | 备注 | 空 |

---

## 10. 港口活动时间

进港时间随机 `00:00-12:00`，离港时间 `12:00-24:00`

## 11. Port of Call 港口 fallback 速查表

以下港口在 `port_map.json` 中不存在（测试于 UNIVERSE HARMONY / IMO 9222546 的 LAST TEN PORT），手动映射已验证可用：

| SEPETIBA | 伊塔瓜伊/塞佩蒂巴 | **BRSPB**-伊塔瓜伊(Itaguai) | BR |
| 原始港口名 | 替代港口 | 替代代码 | 国家 |
|------------|----------|----------|------|
| OPEN SEA | 公海 | THS-公海 | UN |
| CHENJIIAGANG | 天津新港 | CNTXG-天津新港(Tianjingang) | CN |
| KENDARI | 乌戎潘当 | IDUPG-乌戎潘当(Ujung Pandang) | ID |
| MORMUGAO | 哈迪亚 | INHDA-哈迪亚(HALDIA) | IN |
| GO DAU | 岘港 | VNDAD-岘港(Da-Nang/ Da Nang) | VN |
| LEAMCHABANG | 林查班 | THLCH-林查班(Laem Chabang) | TH |
| BINHAI | 滨海 | CNBHI-滨海(BINHAI) | CN |

## 12. RED FILL 颜色格式（openpyxl 坑点）

`PatternFill` 颜色必须使用 **ARGB 格式**（8位hex），`FF` = 完全不透明，`CC` = 半透明浅红：
```python
RED_FILL = PatternFill(start_color="FFFFCCCC", end_color="FFFFCCCC", fill_type="solid")  # ✅ 正确（Alpha=FF 完全不透明）
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")      # ❌ 缺少Alpha通道 → 透明红色
```
实际效果：`start_color="FFCCCC"` = 透明色（完全看不见）；`start_color="FFFFCCCC"` = 可见浅红色。

## 13. `random_port_same_country()` bug

**bug**：函数从 port 名取前2字符作为国家码（如 `GO DAU` → `GO` 共和国，`LEAMCHABANG` → `LE` 莱索托），导致随机选到错误国家的港口。

**正确做法**：从 `_raw_country` 字段（原始国家名）用 `normalize_code()` 获取完整国家代码，再取其前2位作为国家前缀筛选 port_map。

```python
# 错误 ❌
country_prefix = v[:2]  # "GO DAU" → "GO" (加纳共和国)

# 正确 ✅
country_code = normalize_code(country_raw, NATIONALITY_MAP)  # "VIETNAM" → "VN-越南"
country_prefix = country_code[:2]  # "VN"
candidates = [full for code, full in PORT_MAP.items() if code[:2].upper() == country_prefix]
```

## 14. Port of Call 仅处理模式

当用户说"仅处理这个文件"且文件只有 Port of Call 时：
- 调用 `normalize_ports(raw_ports)` 只生成海事活动信息
- 船员 sheet 留空，无需造假数据
- 直接读取 Port of Call Excel/文件，填写"海事船岸活动信息" sheet 即可

## 15. Port of Call PDF 日期格式

POC PDF（如 IMO 标准格式 LAST 10 PORTS OF CALL）中，日期格式为 `dd-Mon-yyyy`（PDF文本层），例如：
- `09-May-2026`
- `26-03-29 & 07-00-00`（日期 + 时间，连字符分隔）

**⚠️ 解析注意事项**：
- PDF 文本中日期是 `dd-Mon-yyyy`（如 `26-Mar-2026`），需用 `%d-%b-%Y` 解析
- 如果是 `26-03-29` 格式（无月份名），可能是 `%y-%m-%d`
- 时间部分用 `&` 分隔：`26-03-29 & 07-00-00`
- **港口顺序**：从新到旧（PDF第1行=最新进港），不要反转

```python
def parse_poc_pdf_date(s):
    """解析 '26-03-29 & 07-00-00' 或 '09-May-2026' 格式"""
    s = s.strip()
    # 分离日期和时间
    date_part = s.split('&')[0].strip()
    time_part = s.split('&')[1].strip() if '&' in s else None

    for fmt in ["%d-%b-%Y", "%d-%m-%y", "%d-%m-%Y"]:
        try:
            dt = datetime.strptime(date_part, fmt)
            break
        except:
            continue

    if time_part:
        for tf in ["%H-%M-%S", "%H-%M-%S"]:
            try:
                t = datetime.strptime(time_part.strip(), tf)
                dt = dt.replace(hour=t.hour, minute=t.minute, second=t.second)
                break
            except:
                pass
    return dt
```
