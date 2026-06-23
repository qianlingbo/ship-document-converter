# GUANG HUA — Port of Call Layout

**船舶**: M/V GUANG HUA (IMO 9659696, Call Sign: VRMM8, Flag: Hong Kong)
**文件**: `4.GUANGHUA PORTS OF CALL LIST (1).XLS`
**格式**: .xls (xlrd 读取)
**港口数**: 10个

## Sheet: Sheet1

Row0: `PORTS OF CALL LIST` (标题)
Row1: `M/V  GUANG HUA`
Row2: `PORT: ZHOU SHAN | Call Sign: VRMM8 | IMO NO. 9659696 | DATE: <excel_date>`
Row3: `VOY: 34 | Port of Registry: Hong Kong`
Row4: 表头 — `NO. | HARBOR,COUNTRY | DATE OF ARRIVAL | DATE OF DEPARTURE | REASON OF CALL | SECURITY LEVEL | UNCODE`
Row5-Row14: 数据行（从新到旧）
Row16: `MASTER:CHEN ZHE` (船长签名区，跳过)
Row17-Row20: 空行

## 数据（Row5-Row14）

| # | HARBOR,COUNTRY | ARRIVAL | DEPARTURE | REASON | SEC | UNCODE |
|---|----------------|---------|-----------|--------|-----|--------|
| 1 | LIAN YUNGANG/CHINA | 16/06/2026 | 19/06/2026 | UNLOAD IRON ORE | 1 | CNLYG |
| 2 | PORT WALCOTT/AUS | 27/05/2026 | 31/05/2026 | LOAD IRON ORE | 1 | AUPWL |
| 3 | BAOSHAN/CHINA | 10/05/2026 | 13/05/2026 | UNLOAD IRON ORE | 1 | CNBAN |
| 4 | MA JISHAN/CHINA | 06/05/2026 | 09/05/2026 | UNLOAD IRON ORE | 1 | CNMJS |
| 5 | PORT HEDLAND/AUS | 16/04/2026 | 21/04/2026 | LOAD IRON ORE | 1 | AUPHE |
| 6 | LONGKOU,CHINA | 28/03/2026 | 01/04/2026 | UNLOAD IRON ORE | 1 | CNLKU |
| 7 | NINGBO/CHINA | 21/03/2026 | 25/03/2026 | UNLOAD IRON ORE | 1 | CNNGB |
| 8 | PORT HEDLAND/AUS | 28/02/2026 | 06/03/2026 | LOAD IRON ORE | 1 | AUPHE |
| 9 | ZHUHAI/CHINA | 15/02/2026 | 18/02/2026 | UNLOAD IRON ORE | 1 | CNZUH |
| 10 | ZHANJIANG/CHINA | 04/02/2026 | 13/02/2026 | UNLOAD IRON ORE | 1 | CNZHA |

## 港口代码核对结果

| 源文件代码 | port_map核对 | 正确代码 | 正确港口名 |
|-----------|-------------|---------|-----------|
| CNLYG | ✓ 存在 | CNLYG | CNLYG-连云港(Lianyungang) |
| AUPWL | ✗ 错误 | AUPWA | AUPWA-沃尔科特港(PORT WALCOTT) |
| CNBAN | ✗ 错误 | CNBSD | CNBSD-宝山(Baoshan) |
| CNMJS | ✗ 不存在 | — | 手动 CNMJS-马迹山 |
| AUPHE | ✓ 存在 | AUPHE | AUPHE-黑德兰港(Port Hedland) |
| CNLKU | ✓ 存在 | CNLKU | CNLKU-龙口 |
| CNNGB | ✗ 错误 | CNNBO | CNNBO-宁波(Ningbo) |
| CNZUH | ✓ 存在 | CNZUH | CNZUH-珠海(Zhuhai) |
| CNZHA | ✗ 错误 | CNZNG | CNZNG-湛江港(Zhanjianggang) |

## 国家代码映射

- `CHINA` → `CN-中国`
- `AUS` → `AU-澳大利亚`
- `LONGKOU,CHINA` 用 `,` 分割（不是 `/`）

## 日期格式

- `dd/mm/yyyy`（如 `16/06/2026`）
