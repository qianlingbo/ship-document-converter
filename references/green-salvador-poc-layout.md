# GREEN SALVADOR PORT OF CALL Layout

**船舶**: M.V. GREEN SALVADOR | 国籍: LIBERIA | IMO: 9976460
**Gross Tons**: 49122 MT
**文件**: `7. LAST 10 Port of Call.pdf`（文字版 PDF，pdftotext -layout 完美提取）
**ARRIVAL PORT**: ZHOU SHAN, CHINA | DATE: 17-Jun-2026

## Sheet 结构

单 sheet：`海事船岸活动信息`

## 港口数据（10条，从新到旧）

| # | 港口名 | 国家 | 进港 | 离港 |
|---|--------|------|------|------|
| 1 | SHANGHAI | CN-中国 | 2026/1/26 | 2026/1/30 |
| 2 | ZHANG JIA GANG | CN-中国 | 2026/1/31 | 2026/2/30（无效日期，修正为2026/2/28） |
| 3 | TAICANG | CN-中国 | **2026/2/05**（注意：旧版PDF写2026/1/05，新版写2026/2/05，以最新文件为准） | 2026/2/09 |
| 4 | SINGAPORE | SG-新加坡 | 2026/02/16 | 2026/2/17 |
| 5 | VITORIA | BR-巴西 | 2026/3/18 | 2026/3/30 |
| 6 | SANTOS | BR-巴西 | 2026/3/31 | 2026/4/10 |
| 7 | SEPETIBA | BR-巴西 | 2026/04/11 | 2026/4/16 |
| 8 | QINGDAO | CN-中国 | 2026/05/30 | 2026/06/05 |
| 9 | TIAN JIN | CN-中国 | 2026/06/07 | 2026/06/11 |
| 10 | KWANG YANG | KR-韩国 | 2026/06/13 | 2026/06/15 |

## 正确港口代码

| 港口名 | 正确代码 | 错误代码（曾用） |
|--------|----------|----------------|
| SHANGHAI | CNSHG-上海港(Shanghai) | - |
| ZHANG JIA GANG | CNZJG-张家港(Zhangjiagang) | - |
| TAICANG | CNTAC-太仓(Taicang) | - |
| SINGAPORE | SGSIN-新加坡(Singapore) | - |
| VITORIA | BRVIX-维多利亚(Vitoria) | - |
| SANTOS | BRSSZ-桑托斯/圣多斯(Santos) | - |
| SEPETIBA | BRSPB-塞佩蒂巴(Sepetiba) | - |
| QINGDAO | CNQDP-青岛港(Qingdao) | - |
| **TIAN JIN** | **CNTNJ-天津港(Tianjin)** | ~~CNTJN~~ |
| **KWANG YANG** | **KRKAN-光阳(Gwangyang/Kwangyang)** | ~~KRKwangyang~~ |

**教训**：
- 天津的正确 UNLOCODE 是 `CNTNJ`，不是 `CNTJN`
- 光阳的正确 UNLOCODE 是 `KRKAN`（port_map.json 中 KRKAN），不是自行拼接的 `KRKwangyang`

## POC 时间格式

标准格式：`YYYY.MM.DD HH:MM:SS`（点分隔符）
- 进港时间：`00:00:00` ~ `11:59:59`（随机）
- 离港时间：`12:00:00` ~ `23:59:59`（随机）
