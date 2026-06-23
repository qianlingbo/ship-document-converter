# WEN DE Crew List Layout（23人，IMO 9642772）

## 文件信息
- 文件名：`doc_d3e2822e6a9e_1.IMO CREW LIST (带中文名) - (ARR ZHOUSAN).xlsx`
- 船舶：WEN DE / IMO 9642772
- 到达港：ZHOUSHAN，到达日期：2026.06.06
- 船员：23人，中国籍

## 列布局（Row 7 = header，Row 9+ = 数据）

| col | 字段 | 示例 |
|-----|------|------|
| 2 | 序号 | 1, 2, ... |
| 3 | 姓名（含中英文换行） | `朱勇\nZHU YONG` |
| 4 | 职务缩写 | `CAPTAIN`, `C/O`, `2/O`, `3/O`, `C/E`, `2/E`, `3/E`, `4/E`, `ETO`, `BOSUN`, `A.B`, `FITTER`, `OILER`, `C/COOK`, `STEWARD`, `CADET` |
| 5 | 性别 | `M` |
| 6 | 国籍 | `China` |
| 7 | 出生日期（datetime 或 `DD/MM/YYYY` 字符串） | `1978-03-18 00:00:00` 或 `25/02/1986` |
| 8 | 出生地点 | `JIANGSU `（带尾部空格） |
| 9 | 海员证号 | `A90435582` |
| 10 | 海员证有效期 | `2029-07-05 00:00:00` |
| 11 | 护照号 | `EC2519870` |
| 12 | 护照有效期 | `2028-01-31 00:00:00` |
| 13 | 登船日期（datetime 或 `DD/MM/YYYY` 字符串） | `2025-08-16 00:00:00` 或 `05/02/2026` |
| 14 | 登船地点 | `SONGXIA,  CHINA`（带换行和逗号） |

## 职务代码映射（ENGLISH_RANK_MAP — 已验证正确）

```python
{
    "CAPTAIN": "51-船长",
    "C/O": "52-大副",
    "2/O": "53-二副",
    "3/O": "54-三副",
    "4/E": "64-三管轮",
    "BOSUN": "55-值班水手",
    "A.B": "56-高级值班水手",    # 注意：不是55-值班水手（实测 WEN DE 轮 A.B 输出56）
    "C/E": "61-轮机长",
    "2/E": "62-大管轮",
    "3/E": "63-二管轮",
    "FITTER": "65-值班机工",
    "OILER": "66-高级值班机工",  # 注意：不是65-值班机工（实测 WEN DE 轮 OILER 输出66）
    "C/COOK": "65-值班机工",     # 注意：不是66-高级值班机工（实测 WEN DE 轮 C/COOK 输出65）
    "STEWARD": "65-值班机工",    # 注意：不是66-高级值班机工（实测 WEN DE 轮 STEWARD 输出65）
    "CARPENTER": "55-值班水手",
    "ETO": "66-高级值班机工",
    "CADET": "55-值班水手",
}
```

> ⚠️ 此文件的旧版本曾有错误映射（A.B→55, OILER→65, C/COOK→66, STEWARD→66），已修正。

## 登船口岸映射（PORT_FALLBACK）

```python
{
    "SONGXIA":   "CNSON-松下(Songxia)",
    "ZHANJIANG": "CNZNG-湛江港(Zhanjianggang)",
    "YANGPU":    "CNYPG-洋浦(Yangpu)",
    "ZHANGJIAGANG": "CNZJG-张家港(Zhangjiagang)",
}
```

## 日期解析

- datetime 对象：`dt.strftime("%Y%m%d")`
- `DD/MM/YYYY` 字符串：`parts[2]+parts[1]+parts[0]` → `YYYYMMDD`

## 姓名处理

中国船员姓名只保留中文，去掉英文和空格：
```python
import re
chinese_only = re.sub(r'[A-Za-z\s]', '', name).strip()
```
