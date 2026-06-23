# 天王之星 (TIAN WANG ZHI XING) POC Layout

**文件**: `17.OFFICIAL PORT OF CALL LIST.xls`
**船舶**: TIAN WANG ZHI XING | IMO 9464223 | Call Sign BOPI | Voyage 218

## Sheet 结构

sheet names: `['Sheet4', 'PORT OF CALL LIST']`
→ **必须用 `wb.sheet_by_name("PORT OF CALL LIST")`，不能用 index(0)**

## 行结构（PORT OF CALL LIST sheet, 0-indexed）

| Row | 内容 |
|-----|------|
| 0 | `PORT OF CALL LIST` (标题) |
| 1 | Name of Ship / Call Sign / IMO No. |
| 2 | Port of Registry / Voyage No. / Arrival Port |
| 3 | Last Port / Next Port |
| **4** | **表头行** (`No.` `Name of port` `Country` `UNLOCODE` `arrival date` `departure date` `security level` ...) |
| 5-14 | **数据行** (10条记录) |
| 15 | 空行 |
| 16 | `Date and signature by master, authorized agent or officer: LUO JIEJUN` (签名行，跳过) |

## 实测列索引（0-indexed）

```
index:  0       1              2           3          4              5           6    7                              8                              9              10
        No.     Name of port   Country     UNLOCODE   arrival date   departure   sec   Security Threats           Especial response            Ship/Ship      Remarks
                                                                date        level      and Incidents           measure                     Activity
```

- `arrival date` = index **4**
- `departure date` = index **5**
- ⚠️ WEDA（第1条数据）departure date = 空字符串（当天离港未记录）

## 日期格式

实测格式：`yyyy.mm.dd`（如 `2026.05.21`、`2026.04.22`）

⚠️ `xlrd` 读取时是字符串，不是 float（因为源文件就是文本格式 `2026.05.21`）。

## 港口数据（10条，从新到旧）

| # | 港口名 | 国家 | UNLOCODE | 进港 | 离港 |
|----|--------|------|----------|------|------|
| 1 | WEDA | INDONESIA | INWED | 2026.05.21 | *(空)* |
| 2 | CEBU | PHILIPPINE | PHCEB | 2026.05.18 | 2026.05.18 |
| 3 | SHENZHEN | CHINA | CNSHZ | 2026.05.14 | 2026.05.14 |
| 4 | HONGKONG | CHINA | HKHKG | 2026.05.12 | 2026.05.13 |
| 5 | QINGDAO | CHINA | CNTAO | 2026.05.03 | 2026.05.08 |
| 6 | LAEMCHABANG | THAILAND | THLCH | 2026.04.22 | 2026.04.24 |
| 7 | SIHANOUK VILLE | CAMBODIA | KHKOS | 2026.04.20 | 2026.04.21 |
| 8 | DONGGUAN | CHINA | CNDGG | 2026.04.12 | 2026.04.16 |
| 9 | HONGKONG | CHINA | HKHKG | 2026.04.11 | 2026.04.12 |
| 10 | QINGDAO | CHINA | CNTAO | 2026.04.05 | 2026.04.06 |

## 港口匹配注意

- `SHENZHEN` → `SZX` (port_map) ✅
- `HONGKONG` → `HKHKG` (port_map) ✅
- `QINGDAO` → `CNQDP` (port_map, **不是 CNTAO**) ✅
- `DONGGUAN` → `CNGGU` (port_map) ✅
- `CEBU` → `PHCEB` (port_map) ✅
- `WEDA` → `IDWED` (port_map) ✅
- `LAEMCHABANG` → `THLCH` (port_map) ✅
- `SIHANOUK VILLE` → `KHKOS` (port_map) ✅
