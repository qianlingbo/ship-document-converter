# 中远海运津渠 V39 — IMO Crew List 布局笔记

## 文件信息
- 船名：ZHONG YUAN HAI YUN JIN QU（中远海运津渠）
- IMO：9846495 | 呼号：BOPN
- 格式：CHINA IMO CREW LIST（新版 FAL 格式）
- 船员数：23人 | 国籍：全部 CHINA（中国）
- 表头行：第1行（IMO标题）、第4行（Arrival/Departure）、第5行（船舶信息标签）、第9行（字段标签）、第11行（列名）
- **数据从第12行开始**

## 列索引（0-based tuple index）

> ⚠️ Crew List 单行 tuple 长度是 28（含大量 None 列），解析时必须用明确列索引，不能用顺序解包。

| 0-based index | 内容 | 说明 |
|---------------|------|------|
| 3 | No. | 序号（int） |
| 4 | Family name | 英文姓 |
| 5 | Given names | **中文名**（中国船员用此列） |
| 6 | Sex | M/F |
| 7 | Rank or rating | 职务英文全称 |
| 8 | Nationality | CHINA（字符串） |
| 9 | Date of birth | datetime 对象 |
| 10 | Place of birth | 出生地（省英文，如 JIANGSU） |
| 11 | Passport No. | 护照号 |
| 12 | Passport 有效期 | datetime |
| 13 | Seaman's Book No. | **海员证号（中国船员用此列）** |
| 14 | Seaman's Book 有效期 | datetime |
| 15 | Embarkation date | 登船日期 datetime |
| 16 | Embarkation place | 登船港口（英文，如 TIANJIN / NANSHA / CHANGSHU） |

## 职务映射（ENGLISH_RANK_MAP 中已包含）

本文件出现的职务：
- `Master` → 51-船长
- `Commissar` → 51-船长（政委）
- `Chief Officer` → 52-大副
- `Second Officer` → 53-二副
- `Third Officer` → 54-三副
- `Bosun` → 55-值班水手
- `Carpenter` → 55-值班水手
- `Able Bodied` → 56-高级值班水手
- `Chief Engineer` → 61-轮机长
- `Second Engineer` → 62-大管轮
- `Third Engineer` → 63-二管轮
- `Forth Engineer` → 64-三管轮
- `Electro-technical Officer` → 66-高级值班机工
- `Chief Motorman` → 65-值班机工
- `Motorman` → 65-值班机工
- `Chief Cook` → 65-值班机工
- `Steward` → 65-值班机工
- `Cadet Captain` → 55-值班水手
- `Cadet ETO` → 66-高级值班机工

## 读取代码模板

```python
import openpyxl
wb = openpyxl.load_workbook('file.xlsx')
ws = wb.active
src_rows = list(ws.iter_rows(values_only=True))
crew_rows = [r for r in src_rows[11:] if r[3] is not None and isinstance(r[3], int)]

for cr in crew_rows:
    no      = cr[3]   # 序号
    en_name = cr[4]   # 英文姓
    cn_name = cr[5]   # 中文名
    sex     = cr[6]
    rank    = cr[7]
    nat     = cr[8]
    dob     = cr[9]    # datetime
    birth_p = cr[10]   # 省英文
    passport= cr[11]
    sb      = cr[13]   # 海员证号
    embark_d= cr[15]   # datetime
    embark_p= cr[16]   # 港口英文
```

## 登船口岸原始值（V39）

| 船员序号 | 登船日期 | 登船口岸（原始） | 备注 |
|----------|----------|------------------|------|
| 1 | 2026-04-17 | TIANJIN | |
| 2-15 | 2025-12-14 | NANSHA | 南沙 |
| 16 | 2026-05-09 | RUGAO | 如皋 |
| 17 | 2025-12-14 | NANSHA | |
| 22 | 2026-04-17 | TIAN JIN | 天津（空格变体） |
| 23 | 2026-04-26 | CHANGSHU | 常熟 |

> 注意：`TIAN JIN`（中间有空格）不是标准 UNLOCODE，port_map 匹配不到，需做 normalize 处理。
