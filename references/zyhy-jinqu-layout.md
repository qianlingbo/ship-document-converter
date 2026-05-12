# ZHONG YUAN HAI YUN JIN QU — Crew List 布局笔记

## 文件格式
- 新版 IMO FAL 格式，表头行分散在多行
- 数据从 **第12行** 开始（index 12，Excel row 12）

## 列索引（0-based）

| 列号 | 内容 | 说明 |
|------|------|------|
| 3 | No. | 序号 |
| 4 | Family name | 英文姓 |
| 5 | Given names | 中文名 |
| 6 | Sex | M/F |
| 7 | Rank or rating | 职务英文 |
| 8 | Nationality | CHINA |
| 9 | Date of birth | datetime 对象 |
| 10 | Place of birth | 出生地（省/英文） |
| 11 | Passport No. | 护照号（外国人用） |
| 12 | Passport 有效期 | datetime |
| 13 | Seaman's Book No. | **海员证号（中国船员用）** |
| 14 | Seaman's Book 有效期 | datetime |
| 15 | Embarkation date | 登船日期 datetime |
| 16 | Embarkation place | 登船港口（英文） |

## 职务映射补充
- `COMMISSAR` → 51-船长（政委）
- `ELECTRO-TECHNICAL OFFICER` → 66-高级值班机工
- `CHIEF MOTORMAN` / `MOTORMAN` → 65-值班机工
- `CADET CAPTAIN` → 55-值班水手
- `CADET ETO` → 66-高级值班机工

## POC 格式
- 标题行在 row 6-7
- 数据从 row 7 开始
- 列：No.(0), PORT(1), COUNTRY(2), UNLOCODE(3), ARR(4), DEP(5), LEVEL(6), OPERATION(7)
- ARR/DEP 列可能包含 `ETA yyyy/m/d` 或 `ETD yyyy/m/d` 前缀，需剥离
