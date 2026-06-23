# COSCO SHIPPING WISDOM 登船口岸映射

**船舶**：COSCO SHIPPING WISDOM | **日期**：2026-06-17

## 实测登船口岸城市 → port_map 映射

| 原始值（去空格后） | port_map key | 映射结果 |
|-----------------|-------------|---------|
| ZHANGJIAGANG | CNZJG | CNZJG-张家港(Zhangjiagang) |
| TAICANG | CNTAC | CNTAC-太仓(Taicang) |
| QINGDAO | CNQDP | CNQDP-青岛港 |
| CHANGSHU | CNCGS | CNCGS-常熟(Changshu) |
| NANSHA | CNNSA | CNNSA-南沙(Nansha) |

**注意**：`NANSHA` 在 POC 中对应 `CNNSA-南沙(Nansha)`，embark_port_map 也应映射到 CNNSA。

## 本次 POC 港口代码（10个）

| 港口 | port_map key | 映射结果 | 国家 |
|------|-------------|---------|------|
| TAICANG | CNTAC | CNTAC-太仓(Taicang) | CN |
| QINGDAO | CNQDG | CNQDG-青岛大港 | CN |
| LIANYUNGANG | CNLYG | CNLYG-连云港(Lianyungang) | CN |
| SHANGHAI | CNSHG | CNSHG-上海港(Shanghaigang) | CN |
| NANSHA | CNNSA | CNNSA-南沙(Nansha) | CN |
| JAKARTA | IDJKT | IDJKT-雅加达(Jakarta) | ID |
| SANTOS | BRSSZ | BRSSZ-桑托斯/圣多斯(Santos) | BR |
| PORTOCEL | BRPCE | BRPCE-蓬塔塞CEL(Portocel) | BR |
| RIO DE JANEIRO | BRRIO | BRRIO-里约热内卢(Rio de Janeiro) | BR |

**注意**：BRPCE（蓬塔塞CEL）在 port_map.json 中不存在，用 UNLOCODE 构建的临时代码 `BRPCE-蓬塔塞CEL(Portocel)`。

## 关键教训

- `ZHANGJIAGANG` → 映射到 `CNZJG-张家港`，不要留空或用其他代码
- NANSHA 既是 POC 港口（CNNSA-南沙）也是登船口岸（CNNSA-南沙）
- BRPCE 是唯一在 port_map 中不存在的港口，使用 fallback 代码
