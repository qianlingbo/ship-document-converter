# FORTUNE PROGRESS PORT OF CALL Layout

**船舶**: FORTUNE PROGRESS | 国籍: PANAMA
**文件**: `PORT OF CALL-ZHOUSHAN.xlsx`

## Sheet 结构

单 sheet: `PORT OF CALL`

## 行结构（0-indexed）

| Row | 内容 |
|-----|------|
| 0 | `LIST OF PORTS OF CALL`（标题） |
| 1 | `Name of ship: FORTUNE PROGRESS` + `Port of Destination:` + `Date of arrived:` |
| 2 | `Nationality of ship: PANAMA` + `Last port of call: BINHAI/CN` |
| 3 | **表头行** (`No` `PORT OF CALL` `COUNTRY` `ARRIVAL` `DEPARTURE` `SHIP SECURITY` `PORT SECURITY`) |
| 4-13 | **数据行** (10条记录，seq 1-10) |
| 14-20 | 空行（seq 11-16均为空） |

## 列结构（0-indexed）

```
index:  0      1              2           3                  4                  5              6              7
        No     PORT OF CALL    COUNTRY     ARRIVAL            DEPARTURE          SHIP SECURITY   PORT SECURITY
```

## 特点

- **日期格式**: `yyyy.mm.dd HH:MM:SS`（如 `2026.03.19 13:54:00`）或 `yyyy.mm.dd`（无时间，如 `2026.05.28`）
- **港口名**: 全大写英文，如 `HUANGPU`, `ZHOUSHAN`, `WEDA`, `OBI ISLAND`
- **国家**: 全大写英文，如 `CHINA`, `INDONESIA`
- **港口匹配**: HUANGPU→CNHUA, HONGKONG→HKHKG, LIANYUNGANG→CNLYG, ZHOUSHAN→CNZOS, WEIFANG→CNWEF, WEDA→IDWED, OBI ISLAND→IDOBI
- **BINHAI**: 不在 port_map.json，用 `CNBHI-滨海(BINHAI)` 临时代码，需标红

## 港口数据（10条，从新到旧）

| # | 港口名 | 国家 | 进港 | 离港 |
|---|--------|------|------|------|
| 1 | HUANGPU | CHINA | 2026.03.19 13:54 | 2026.03.21 11:06 |
| 2 | HONGKONG | CHINA | 2026.03.21 21:00 | 2026.03.22 01:18 |
| 3 | LIANYUNGANG | CHINA | 2026.03.27 08:24 | 2026.03.28 20:24 |
| 4 | ZHOUSHAN | CHINA | 2026.03.30 15:18 | 2026.04.09 13:36 |
| 5 | WEIFANG | CHINA | 2026.04.12 15:18 | 2026.04.16 08:18 |
| 6 | ZHOUSHAN | CHINA | 2026.04.19 08:12 | 2026.04.20 11:18 |
| 7 | WEDA | INDONESIA | 2026.04.29 11:30 | 2026.05.15 06:24 |
| 8 | OBI ISLAND | INDONESIA | 2026.05.16 06:06 | 2026.05.20 05:42 |
| 9 | ZHOUSHAN | CHINA | 2026.05.28 | 2026.05.29 |
| 10 | BINHAI | CHINA | 2026.05.31 | 2026.06.01 |
