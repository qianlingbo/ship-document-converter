# UNIVERSE HARMONY Crew List — 布局笔记

**船舶**：UNIVERSE HARMONY | **IMO**：9222546 | **船旗**：PANAMA

## Crew List 表头（第6行）

```
Index:  0    1                    2        3           4           5      6         7                    8         9                    10                  11                  12                  13                  14
        No   Family name          Chinese  Rank        Sex         Natl   Birth                    Seaman's Book              Passport             Passport           Sign on date & place
                                                         name        (P.R.C) Date&Place            Number    Exp               Number  Exp          
```

**关键列映射**：
- `row[0]` → 序号
- `row[1]` → Family name（英文姓）
- `row[3]` → **Chinese name（中文名）** ← 姓名用这个
- `row[4]` → Rank/职务缩写
- `row[5]` → Sex（M/F）
- `row[6]` → Nationality（P.R.C）
- `row[7]` → 出生日期（datetime对象）
- `row[8]` → 出生地点（英文省名，如 HEILONGJIANG / FUJIAN）
- `row[9]` → **Seaman's Book 号码** ← 中国船员证件号用这个
- `row[11]` → Passport 号码（外国船员用这个）
- `row[13]` → 登船日期（datetime对象）
- `row[14]` → 登船地点/口岸

## 职务代码映射（与skill一致）

| 缩写 | 标准职务 |
|------|----------|
| MSTR | 51-船长 |
| C/O | 52-大副 |
| 2/O | 53-二副 |
| 3/O | 54-三副 |
| C/E | 61-轮机长 |
| 2/E | 62-大管轮 |
| 3/E | 63-二管轮 |
| 4/E | 64-三管轮 |
| BSN | 55-值班水手 |
| AB | 56-高级值班水手 |
| OS | 55-值班水手 |
| FTR | 65-值班机工 |
| OLR | 66-高级值班机工 |
| E/E | 66-高级值班机工 |
| C/CK | 65-值班机工 |

## 登船地点（sign_on_port）映射

| 原始值 | 映射结果 |
|--------|----------|
| CAMPHA | VNCMP → 越南岘港 |
| GODAU | VNDAD → 越南岘港 |
| CHENJIAGANG | CNCIG → 陈家港 |
| DONGJIAKOU | CNDJK → 董家口 |
| SHANGHAI | CNSGH → 上海港 |
| JINGJIANG | CNJJG → 靖江 |

## POC 港口（Port of Call）— UNIVERSE HARMONY 实测

POC Excel 列结构（LAST TEN PORT sheet）：
```
Index:  0       1          2           3              4              5          6    7         8
        No      Port Name  Country     Arrival Date   Depature Date  Port(Sec)  Ship Purpose   Cargo
```
⚠️ `Depature Date` 在 **index=4**，index=5 是 Port security level（固定为1）。

POC 港口原始值 → PORT_FALLBACK 映射：

| 原始港口名 | 替代代码 | 说明 |
|------------|----------|------|
| DONGJIAKOU | CNDJK | 董家口 |
| CHENJIIAGANG | CNCIG | 陈家港（原文 CHENJIIAGANG 拼写）|
| ZHOUSHAN | CNZOS | 舟山 |
| KENDARI | IDUPG | 乌戎潘当（印尼）|
| MORMUGAO | INHDA | 哈迪亚（印度，原拼写 MORMUGAO）|
| PASIR GUDANG | MYPGG | 巴西拉再也（马来西亚）|
| GO DAU | VNDAD | 岘港（越南）|
| BANGKOK | THBKK | 曼谷 |
| LEAMCHABANG | THLCH | 林查班（泰国）|
| SHANGHAI | CNSGH | 上海港 |

## 日期格式说明

- 出生日期 `row[7]`：直接是 `datetime.datetime` 对象，无需解析字符串
- 登船日期 `row[13]`：同上，直接 `datetime` 对象
- POC 到达/离港日期：字符串格式 `dd/mm/yyyy`（如 `13/05/2026`）
