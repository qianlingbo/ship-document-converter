# FORTUNE PROGRESS Crew List Layout

**文件**: `IMO CREW LIST-PASSPORT&SEAMAN'S BOOK-ARR.xls`
**船舶**: FORTUNE PROGRESS | 船旗 PANAMA | ARR=抵港报告
**日期**: 2026/05/28 舟山

## Sheet 结构

- 单 sheet，name: `Sheet1`
- 总行数：25行，船员 15人（Row8 ~ Row22）

## 实测列索引（0-indexed，xlrd 读取）

```
Index:  0      1              2                  3       4         5         6          7                      8                      9         10                  11
        No.    Family name    Given names        Rank    Nationality BirthDate BirthPlace SeamanBook/Passport   Expiry(SB/PP)     SignOn    SignOnPlace          Phone
```

| 字段 | 列索引 | 数据类型 | 示例 |
|------|--------|----------|------|
| 序号 | 0 | float | `1.0` |
| 姓（合并列） | 1 | str | `WANG 王 `（含空格+中文） |
| 名（合并列） | 2 | str | `XIONGWEI 雄伟`（含拼音+中文） |
| 职务 | 3 | str | `CAPT` / `C/O` / `AB` / `OLR` |
| 国籍 | 4 | str | `CHINESE` / `INDONESIA` |
| 出生日期 | 5 | Excel float（xlrd serial） | `27191.0` → `19740611` |
| 出生地点 | 6 | str | `ZHEJIANG` / `JIANGXI` |
| 证件号 | 7 | str（斜杠分隔） | `A90606061/EJ3982685` = 海员证/护照 |
| 有效期 | 8 | str | `22.APR.2031/16.SEP.2030` |
| 签船日期 | 9 | Excel float（xlrd serial） | `45995.0` → `20251204` |
| 签船地点 | 10 | str | `HAIPHONG,VIETNAM` / `JIANGYIN,CHINA` |
| 电话 | 11 | float | `13567657811.0` |

## 关键特征

### 姓名列合并（不同于其他变体）

其他 Crew List 姓名在单列（如 `张三 ZHANG SAN`），本文件姓/名分占两列：
- col1 = `WANG 王 `（英文姓 + 空格 + 中文）
- col2 = `XIONGWEI 雄伟`（英文名 + 空格 + 中文）

**姓名提取逻辑**：
```python
name1 = str(ws.cell(r, 1).value or "").strip()  # "WANG 王 "
name2 = str(ws.cell(r, 2).value or "").strip()  # "XIONGWEI 雄伟"
cn_chars = re.findall(r'[\u4e00-\u9fff]', name1 + name2)  # 提取所有中文
en_parts = re.sub(r'[\u4e00-\u9fff]', ' ', name1 + name2).split()  # 提取非中文部分
full_cn = "".join(cn_chars)  # 中文合并
full_en = " ".join(en_parts).strip().upper()  # 英文部分大写
full_name = full_cn if full_cn else full_en  # 有中文用中文，否则用英文
```

### 证件号合并格式

`A90606061/EJ3982685` → 斜杠前=海员证号，斜杠后=护照号

### 外国船员出生地点

用户纠正规则：外国船员出生地点 = 国籍中文名（如 `印度尼西亚`），不是源数据原文。

### 职务映射（实测）

| 原始值 | 标准职务 |
|--------|----------|
| `CAPT` | `51-船长` |
| `C/O` | `52-大副` |
| `2/O` | `53-二副` |
| `C/E` | `61-轮机长` |
| `2/E` | `62-大管轮` |
| `4/E` | `64-三管轮` |
| `AB` | `56-高级值班水手` |
| `ASSISTANT OFFICER` | `54-三副` |
| `OLR` | `65-值班机工` |
| `C/COOK` | `65-值班机工` |

### 国籍映射

| 源数据 | nat_map key | 输出 |
|--------|-------------|------|
| `CHINESE` | 不在keys中 | `CN-中国` |
| `INDONESIA` | 不在keys中 | `ID-印度尼西亚` |

需要额外 NAT_FULL 映射处理英文国名。

### 船舶信息（Row4 / Row6）

- Row4: `FORTUNE PROGRESS` / `ZHOUSHAN/CN`（船名 + 抵港港口）
- Row6: `PANAMA` / `BINHAI/CN`（船旗 + 来自港口）

### 船长签名行

- Row23: `16. Date and signature by master...`（空行）
- Row24: `MASTER:` / `WANG XIONGWEI`（签名）
