# 天王之星 (TIAN WANG ZHI XING) Crew List Layout

**文件**: `1.CREW LIST-天王之星(ARR).xls`
**船舶**: TIAN WANG ZHI XING | IMO 9464223 | Call Sign BOPI

## Sheet 结构

- 单 sheet，sheet name: `CREW LIST`
- 总行数：27 行（含表头），实际船员 20 人

## 实测列索引（0-indexed，xlrd 读取）

| index | 字段名 | 数据类型 | 示例 |
|--------|--------|----------|------|
| 0 | No. | int | `1` |
| 1 | 姓名 | str | `张三` 或 `张三 ZHANG SAN` |
| 2 |  Rank | str | `MASTER` / `C/O` |
| 3 | Nationality | str | `CN` |
| 4 | Sex | str | `M` / `F` |
| **6** | **出生日期** | **Excel float（xlrd serial）** | `高峰` = `25569.0` + `~` |
| 7 | 出生地 | str | `SHANDONG` |
| **8** | **Seaman Book** | **str** | `A90395125` |
| 10 | Passport | str | `ER3358421` |
| 11 | Country | str | `CHINA` |
| **13** | **登船日期** | **str `YYYY/MM/DD`** | `2024/01/15` |
| **14** | **登船地点** | **str（原始城市名）** | `XIAMEN` |

### ⚠️ 关键字段差异（中国船员）

| 字段 | ❌ 错误来源 | ✅ 正确来源 |
|------|------------|------------|
| 证件号码 | col10 (Passport) | **col8 (Seaman Book)** |
| 登船口岸 | col14 原始城市名 | 需映射为 `CNXAM-厦门` 等格式 |

## 日期解析

```python
# 出生日期：Excel float → YYYYMMDD
dt = xlrd.xldate_as_datetime(cell_value, wb.datemode)
birth_date = dt.strftime("%Y%m%d")

# 登船日期：字符串 YYYY/MM/DD → YYYYMMDD
join_date = cell_value.replace("/", "")  # "2024/01/15" → "20240115"
```

## 登船口岸 PORT_FALLBACK（中国船员）

| 原始值（col14） | 映射结果 |
|-----------------|----------|
| `XIAMEN` | `CNXAM-厦门` |
| `WEIFANG` | `CNWEF-潍坊` |
| `QINGDAO` | `CNQDP-青岛港` |

## 船员职务映射（实测）

| 原始值 | 标准职务 |
|--------|----------|
| `MASTER` | `51-船长` |
| `C/O` | `52-大副` |
| `2/O` | `53-二副` |
| `C/E` | `61-轮机长` |
| `1/E` | `62-大管轮` |
| `2/E` | `63-二管轮` |
| `3/E` | `64-三管轮` |
| `BSN` | `55-值班水手` |
| `AB` | `56-高级值班水手` |
| `PUMPMAN` | `65-值班机工` |
| `MOTORMAN` | `65-值班机工` |
| `COOK` | `65-值班机工` |
| `STEWARD` | `65-值班机工` |
