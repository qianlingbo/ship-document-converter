# GREEN SALVADOR Crew List Layout

**船舶**: M.V. GREEN SALVADOR | 国籍: LIBERIA | IMO: 9976460
**文件**: `1. NEW IMO Crew list - .xls`（xlrd 读取，Sheet: CREW LIST-IMO）

## Sheet 结构

- 总行数：34行，数据行 Row10~Row28（船员19人）
- 表头行：Row9（0-indexed: Row8）

## 列布局（xlrd 0-based）

```
Index:  0     1        3         9        10       11       12         13       14           15            16            17         18           19
        No.   Name     Rank      Nationality Sex    BirthDate BirthPlace Passport    P/P Expiry  SeamanBook    S-B-Expiry   Date         Place
```

| 字段 | 列索引 | 类型 | 备注 |
|------|--------|------|------|
| 序号 | 1 | float | 1~19 |
| 姓名 | 3 | str | 英文拼音全大写 |
| 职务 | 9 | str | CAPT/C/O/2/O/3/O/C/E/2/E/3/E/AB/OS/FITTER/COOK等 |
| 国籍 | 10 | str | CHINESE/VIETNAM/INDONESIA |
| 性别 | 11 | str | M/F |
| 出生日期 | 12 | xlrd float | → `xlrd.xldate_as_datetime(v, wb.datemode).strftime("%Y%m%d")` |
| 出生地点 | 13 | str | 省/城市英文 |
| **护照号** | **14** | str | 外国船员用此列 |
| 护照有效期 | 15 | xlrd float | |
| **海员证号** | **16** | str | 中国船员用此列（col16，非col14！） |
| 海员证有效期 | 17 | xlrd float | |
| 登船日期 | 18 | xlrd float | → `xlrd.xldate_as_datetime(v, wb.datemode).strftime("%Y%m%d")` |
| 登船地点 | 19 | str | SHANGHAI/ZHUHAI/QINGDAO等原始城市名 |

## 船员数据（19人）

| # | 姓名 | 职务 | 国籍 | 出生日期 | 证件类型 | 证件号码 | 登船日期 | 登船地点 |
|---|------|------|------|----------|----------|----------|----------|----------|
| 1 | TIAN JINGWEI | 51-船长 | CN-中国 | 19880902 | 17-海员证 | A90401081 | 20260127 | SHANGHAI |
| 2 | XIA YUAN | 52-大副 | CN-中国 | 19930221 | 17-海员证 | A90227471 | 20260119 | ZHUHAI |
| 3 | PHAM NGUYEN DUC HUY | 53-二副 | VN-越南 | 19960909 | 14-普通护照 | C5870472 | 20260603 | QINGDAO |
| 4 | LISTON RUDIANTO SITANGGANG | 54-三副 | ID-印度尼西亚 | 19891007 | 14-普通护照 | X8741075 | 20260603 | QINGDAO |
| 5 | ZHANG BIN | 61-轮机长 | CN-中国 | 19891101 | 17-海员证 | A90412528 | 20260127 | SHANGHAI |
| 6 | CHEN XIN | 62-大管轮 | CN-中国 | 19881211 | 17-海员证 | A90511346 | 20260119 | ZHUHAI |
| 7 | MOCHAMMAD HARY PRASETIA | 63-二管轮 | ID-印度尼西亚 | 19970609 | 14-普通护照 | E0791490 | 20260127 | SHANGHAI |
| 8 | ZHANG LIHUA | 66-高级值班机工 | CN-中国 | 19860601 | 17-海员证 | A90374794 | 20260603 | QINGDAO |
| 9 | WANG JIE | 56-高级值班水手 | CN-中国 | 19871206 | 17-海员证 | A90487722 | 20260611 | TIAN JIN |
| 10 | DINH TRUONG GIANG | 56-高级值班水手 | VN-越南 | 19920511 | 14-普通护照 | C4509084 | 20260603 | QINGDAO |
| 11 | REZA DWIYANTORO | 56-高级值班水手 | ID-印度尼西亚 | 19990725 | 14-普通护照 | X5196916 | 20260127 | SHANGHAI |
| 12 | M. ILZAM NUZULI | 56-高级值班水手 | ID-印度尼西亚 | 19991227 | 14-普通护照 | C9186799 | 20260127 | SHANGHAI |
| 13 | MUHAMMAD AQSAH | 56-高级值班水手 | ID-印度尼西亚 | 20021227 | 14-普通护照 | X2287127 | 20260603 | QINGDAO |
| 14 | REZQHI ASYSYAWALIH | 56-高级值班水手 | ID-印度尼西亚 | 20011228 | 14-普通护照 | X6114494 | 20260603 | QINGDAO |
| 15 | RAHMAD RIZKY | 65-值班机工 | ID-印度尼西亚 | 19960420 | 14-普通护照 | E3112529 | 20260217 | SINGAPORE |
| 16 | M SIGIT HARIANTO | 65-值班机工 | ID-印度尼西亚 | 19920826 | 14-普通护照 | E2408421 | 20260603 | QINGDAO |
| 17 | SEGER WAHYUDI | 65-值班机工 | ID-印度尼西亚 | 19961104 | 14-普通护照 | X8590792 | 20260603 | QINGDAO |
| 18 | RAZOKY FEMAL PANE | 65-值班机工 | ID-印度尼西亚 | 20050111 | 14-普通护照 | X6187730 | 20260603 | QINGDAO |
| 19 | CAO KHAC THIEN | 65-值班机工 | VN-越南 | 19830810 | 14-普通护照 | P03229797 | 20260603 | QINGDAO |

## 关键教训

**中国船员姓名**：
- 源数据只有英文拼音（TIAN JINGWEI），无法自动还原汉字
- 需要用户提供正确中文名，或提供含中文的证件扫描件
- 禁止输出中英混合：`TIAN JINGWEI` 可以保留英文拼音，但用户要求中文名时必须用户提供

**出生地点规则**：
- 中国船员 → `中国`
- 外国船员 → 国籍中文名（如 `越南`、`印度尼西亚`），不是源数据原文

**证件号码来源**：
- 中国船员 → SeamanBook（col16）
- 外国船员 → Passport（col14）
- **注意**：不能用错列，之前误用 col14（护照）给中国船员，导致证件号全错
