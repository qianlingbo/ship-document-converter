# GENIUS POC Layout（实测）

**船舶**：M.V. GENIUS | 船旗：CYPRUS | Call Sign: 5BHP6

## POC 数据（10港口，从新到旧）

| # | 港口 | 日期（到达-离开） | 目的 | 代码 | 国家 |
|---|------|------------------|------|------|------|
| 1 | HAY POINT | 26-06-12 ~ 26-06-13 | LOADING CARGO | AUHAY | AU |
| 2 | BAHODOPI | 26-05-31 ~ 26-06-02 | CARGO | IDBAH | ID |
| 3 | TELUK RUBIAH | 26-05-13 ~ 26-05-21 | LOADING CARGO | MYTLR（临时代码） | MY |
| 4 | ZHUHAI | 26-05-02 ~ 26-05-07 | CARGO | CNZUH | CN |
| 5 | ABBOT POINT | 26-04-14 ~ 26-04-16 | LOADING CARGO | AUABP | AU |
| 6 | HUANGHUA | 26-03-20 ~ 26-03-27 | CARGO | CNHUA | CN |
| 7 | SINGAPORE | 26-03-05 ~ 26-03-05 | BUNKERING | SGSIN | SG |
| 8 | BOFFA | 26-01-18 ~ 26-01-28 | LOADING CARGO | GNBOF（临时代码） | GN |
| 9 | SINGAPORE | 25-12-19 ~ 25-12-20 | BUNKERING | SGSIN | SG |
| 10 | ZHOUSHAN | 25-12-09 ~ 25-12-10 | CARGO | CNZOS | CN |

## 港口代码确认

- `AUHAY` → `port_map.json` 中存在：`AUHAY-海角(HAY POINT)`
- `CNZUH` → `port_map.json` 中存在：`CNZUH-珠海(Zhuhai)`
- `CNHUA` → `port_map.json` 中存在：`CNHUA-黄埔(Huangpuwaimaocangku)`
- `AUABP` → `port_map.json` 中存在：`AUABP-阿博特波特`
- `SGSIN` → `port_map.json` 中存在
- `IDBAH` → `port_map.json` 中存在
- `CNZOS` → `port_map.json` 中存在

## 临时代码（需人工确认）

- **TELUK RUBIAH**：`MYTLR` — 马来西亚港口，port_map.json 中无，UNLOCODE 构建为 MYTLR（待确认）
- **BOFFA**：`GNBOF` — 几内亚港口，port_map.json 中无，GN 只有 GNKAM/GNCKY，BOFFA 不在表中，用 GNBOF 临时代码（待确认）
