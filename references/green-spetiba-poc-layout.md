# GREEN SEPETIBA Port of Call Layout

**文件**: `07 Last 10 Ports of Call List.pdf`
**格式**: PDF（pdfminer 提取文本）

## 文本结构

```
PORTS OF CALL LIST
1.Name of ship: GREEN SEPETIBA
2.Port of arrival: ZHOUSHAN
3.Date of arrival: 2026/6/2
4.Nationality of ship: LIBERIA
5.Port arrived from: TAICANG
7. Next Port of Call: NANSHA,CHINA

6.No. | 7. Name & Country of Port | 8.Date of Arrival | 9.Date of Departure | 10.Purpose of Calling | 11.Security Level

1 | TAICANG / CHINA | 30/05/2026 | 01/06/2026 | LOADING | 1
2 | SHANGHAI / CHINA | 29/05/2026 | 30/05/2026 | LOADING | 1
3 | QINGDAO / CHINA | 22/05/2026 | 28/05/2026 | UNLOADING/LOADING | 1
4 | SHANGHAI / CHINA | 18/05/2026 | 20/05/2026 | UNLOADING | 1
5 | QUANZHOU / CHINA | 14/05/2026 | 16/05/2026 | UNLOADING | 1
6 | SINGAPORE / SINGAPORE | 07/05/2026 | 07/05/2026 | BUNKERING | 1
7 | SEPETIBA / BRAZIL | 28/03/2026 | 03/04/2026 | LOADING | 1
8 | SANTOS / BRAZIL | 26/03/2026 | 27/03/2026 | LOADING | 1
9 | PORTOCEL / BRAZIL | 20/03/2026 | 24/03/2026 | LOADING | 1
10 | SEPETIBA / BRAZIL | 16/03/2026 | 19/03/2026 | UNLOADING | 1
```

## 日期格式
`dd/mm/yyyy`（斜杠分隔）

## 港口数据（从新到旧，第1行=最新进港）

| # | 港口 | 国家 | 进港 | 离港 |
|---|------|------|------|------|
| 1 | TAICANG | CHINA | 30/05/2026 | 01/06/2026 |
| 2 | SHANGHAI | CHINA | 29/05/2026 | 30/05/2026 |
| 3 | QINGDAO | CHINA | 22/05/2026 | 28/05/2026 |
| 4 | SHANGHAI | CHINA | 18/05/2026 | 20/05/2026 |
| 5 | QUANZHOU | CHINA | 14/05/2026 | 16/05/2026 |
| 6 | SINGAPORE | SINGAPORE | 07/05/2026 | 07/05/2026 |
| 7 | SEPETIBA | BRAZIL | 28/03/2026 | 03/04/2026 |
| 8 | SANTOS | BRAZIL | 26/03/2026 | 27/03/2026 |
| 9 | PORTOCEL | BRAZIL | 20/03/2026 | 24/03/2026 |
| 10 | SEPETIBA | BRAZIL | 16/03/2026 | 19/03/2026 |

## 港口代码映射（SPECIAL）

| 港口名 | 代码 | 中文名 |
|--------|------|--------|
| TAICANG | CNTAC | 太仓 |
| SHANGHAI | CNSHG | 上海港 |
| QINGDAO | CNQDP | 青岛港 |
| QUANZHOU | CNQAU | 泉州 |
| SINGAPORE | SGSIN | 新加坡 |
| SEPETIBA | BRSPT | 塞佩提巴 |
| SANTOS | BRSSZ | 桑托斯/圣多斯 |
| PORTOCEL | BRPCE | 蓬塔塞CEL |
