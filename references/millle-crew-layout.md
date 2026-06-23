# MILLIE Crew List Layout

**文件**: `MILLIE IMO Crew List -.xls`
**船舶**: MV MILLIE | IMO 9492103 | 船旗：利比里亚 | 出发港：京唐

## 实测列索引（0-indexed，xlrd 读取）

主数据行（奇数行，如 Row9, 11, 13...）：

| index | 字段名 | 数据类型 | 示例 |
|--------|--------|----------|------|
| 1 | 序号 | float/int | `1.0` |
| 2 | 姓名 | str | `WANG,DAOHUI  王道慧` 或 `HTUN HTUN WIN` |
| 3 | 性别 | str | `M` |
| 4 | 职务 | str | `MASTER` / `C/O` |
| 5 | 国籍（英文全称）| str | `CHINESE` / `MYANMAR` |
| 6 | 出生日期 | Excel float | `27105.0` |
| 7 | 护照号/有效期 | str | `ER2923687` |
| 8 | 海员证号/有效期 | str | `A90264092` |
| 9 | 登船日期 | Excel float | `46146.0` |

子数据行（偶数行，紧跟主数据行，如 Row10, 12, 14...）：

| index | 字段名 | 数据类型 | 示例 |
|--------|--------|----------|------|
| 6 | 出生地点 | str | `JIANGSU` / `MAWLAMYINE` |
| 7 | 护照有效期 | Excel float | `49700.0` |
| 8 | 海员证有效期 | Excel float | `46653.0` |
| 9 | **登船地点** | str | `CAOFEIDIAN` / `SINGAPORE` |

## ⚠️ 关键解析：join_place 在子行 col9

```python
# 正确：
row_main = ws_crew.row(i)       # 奇数行
row_sub  = ws_crew.row(i+1)     # 偶数行
join_place = row_sub[9].value   # ✅ 子行 col9

# 错误（曾犯过）：
join_place = row_main[9].value  # ❌ 这是登船日期，不是登船地点
```

## 日期解析

```python
dm = wb_crew.datemode
birth_date = xldate_to_yyyymmdd(row_main[6].value, dm)   # 出生日期
join_date  = xldate_to_yyyymmdd(row_main[9].value, dm)   # 登船日期
```

## Crew List 国籍值（英文全称 → 代码映射）

```python
NAT_FULL_TO_CODE = {
    "CHINESE": "CN",
    "MYANMAR": "MM",
    "SIERRA LEONE": "SL",
    "LIBERIA": "LR",
}
```

## 新增职务类型（MILLIE 实测）

| 原始值 | 标准职务 |
|--------|----------|
| `DFTR` | `65-值班机工` |
| `EFTR` | `65-值班机工` |
| `OLR` | `65-值班机工` |
| `WIPER` | `65-值班机工` |
| `C/COOK` | `65-值班机工` |
| `MSM` | `65-值班机工` |
| `4/E` | `64-三管轮` |
| `ETO` | `66-高级值班机工` |

## 登船口岸映射（PORT_FALLBACK_JOIN）

| 原始值（子行col9）| 映射结果 |
|-----------------|----------|
| `CAOFEIDIAN` | `CNCFD-曹妃甸(Caofedian)` |
| `JINGTANG` | `CNJTG-京唐(Jingtang)` |
| `SINGAPORE` | `SGSIN-新加坡(Singapore)` |
| `EL DEKHEILA` | `EGDK-埃尔 Dekheila(El Dekheila)` |
| `ZHOUSHAN` | `CNZOS-舟山(Zhoushan)` |
| `SHANGHAI` | `CNSHG-上海(Shanghai)` |
