# MILLIE Last Ten Calling Ports Layout

**文件**: `Last Ten Calling Ports.xls`
**船舶**: MV MILLIE | IMO 9492103

## Sheet 结构

- Sheet name: `Last 10 Port Of Call`
- 表头行：Row5（含 `NO.` + `Port Name` 列标题）

## 实测列索引（0-indexed，xlrd 读取）

| index | 字段名 | 数据类型 | 示例 |
|--------|--------|----------|------|
| 0 | NO. | float | `1.0` |
| 1 | Date of Arrival | Excel float 或 str | `46147.0` 或 `04/05/2026(0318)GMT+8` |
| 2 | Date of Departure | Excel float 或 str | `46160.0` 或 `05/05/2026(1730)GMT+8` |
| 3 | Port Name | str | `JINGTANG` |
| 4 | Country | str | `CHINA` |
| 5 | Unlocode | str | `CNTGS` |
| 6 | SECURITY LEVEL OF Port | float | `1.0` |
| 7 | SECURITY LEVEL OF Ship | float | `1.0` |
| 8 | PURPOSE | str | `IRON ORE UNLOADING` |

## 日期格式（实测两种混合）

```python
def parse_poc_date(val, dm):
    # 格式1：Excel float → xlrd.xldate_as_datetime
    if isinstance(val, float):
        return xldate_to_yyyy_mm_dd(val, dm)
    # 格式2：字符串 "04/05/2026(0318)GMT+8"
    s = str(val).strip()
    for fmt in ["%d/%m/%Y", "%d/%m/%y"]:
        try:
            return datetime.datetime.strptime(s[:10], fmt).strftime("%Y/%m/%d")
        except:
            pass
    return None
```

## 港口映射（POC_PORT_FALLBACK）

| 原始港口名 | UNLOCODE | 替代代码 | 替代名称 |
|------------|----------|----------|----------|
| JINGTANG | CNTGS | CNTGS | 京唐(Jingtang) |
| CAOFEIDIAN | CNCFD | CNCFD | 曹妃甸(Caofedian) |
| ILHA GUAIBA | BRSPB | BRSPB | 伊尔哈瓜iba(Ilha Guaita) |
| SAO FRANCISCO DO SUL | BRSFS | BRSFS | 南圣弗朗西斯科(Sao Francisco do Sul) |
| SINGAPORE | SGSIN | SGSIN | 新加坡(Singapore) |
| HUANGHUA | CNHUH | CNHUH | 黄骅(Huanghua) |
| FREETOWN PORT | SLFNA | SLFNA | 弗里敦(Freetown) |
| LANSHAN | CNLSN | CNLSN | 岚山１ |

## 国家字段（POC-country-map 需补充）

POC 中 `FREETOWN PORT` 对应国家为 `SIERRA LEONE`，需确认 `poc-country-map.json` 包含：

```json
{ "SIERRA LEONE": "SL-塞拉利昂" }
```

## 船长签名行

Row20 含有 `CAPT. WANG,DAOHUI`，Row21 含有 `MASTER OF MV. MILLIE` — 解析时跳过以 `CAPTAIN` / `MASTER` 开头的行。
