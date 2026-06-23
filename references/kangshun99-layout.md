# KANGSHUN 99 单证录入实测记录

## 基本信息
- 船名：KANGSHUN 99
- 船旗：PANAMA（巴拿马）
- 泊位：ZHOUSHAN（舟山）
- 船员数：15人
- 处理日期：20260620

## Crew List 列布局（.xls，xlrd，0-based）

```
Index:  0     1                      3       4      5           6              7            8            9           10           11         12         13         14
        No    Name(中文\n英文)     Nat     Sex    Rank      DOB           POB          SB_No        SB_Exp       Country     JoinDate    JoinPlace   Passport   PP_Exp
```

| 字段 | 列索引 | 示例 |
|------|--------|------|
| 序号 | 0 | `1、` |
| 姓名（中文+英文） | 1 | `卜庆丹\nBU QINGDAN` / `NGUYEN THANH SON` |
| 国籍 | 3 | `China` / `Vietnam` / `Myanmar` / `Indonesia` |
| 性别 | 4 | `M` |
| 职务 | 5 | `CAPT` / `C/O` / `AB3` / `BSN` / `OS` / `CE` / `OLR1` |
| 出生日期 | 6 | 字符串 `1981.12.13` 或 Excel float |
| 出生地点 | 7 | `LIAO NING` / `THAI BINH` |
| 海员证号 | 8 | `A90239549` |
| 海员证有效期 | 9 | `2027.07.18` / `LONG TIME` |
| 国家 | 10 | `China` |
| 登船日期 | 11 | `2026.05.11` |
| 登船地点 | 12 | `QIN ZHOU` / `WU HAN` / `YANG ZHOU` |
| 护照号 | 13 | `C9964251` / `ER3819560` |
| 护照有效期 | 14 | `2036.03.23` |

**职务映射（ENGLISH_RANK_MAP 追加条目）**：
```python
ENGLISH_RANK_MAP = {
    "CAPT":"51-船长","C/O":"52-大副","2/O":"53-二副","3/O":"54-三副","4/O":"54-三副",
    "BSN":"55-值班水手","AB1":"56-高级值班水手","AB2":"56-高级值班水手","AB3":"56-高级值班水手",
    "OS":"55-值班水手",
    "CE":"61-轮机长","2/E":"62-大管轮","3/E":"63-二管轮","4/E":"64-三管轮",
    "WIPER":"65-值班机工","OLR1":"65-值班机工","OLR2":"65-值班机工","OLR3":"65-值班机工",
    "C/COOK":"65-值班机工",
}
```

**国籍映射（NAT_MAP 追加条目）**：
```python
NAT_MAP = {
    "China":"CN-中国","CHINA":"CN-中国",
    "Vietnam":"VN-越南","VIETNAM":"VN-越南",
    "Myanmar":"MM-缅甸","MYANMAR":"MM-缅甸",
    "Indonesia":"ID-印度尼西亚","INDONESIA":"ID-印度尼西亚",
}
```

## POC 列布局（.xls，xlrd）

Sheet名称：`10-PORT CALL`

| 字段 | 列索引 |
|------|--------|
| No | 0 |
| 港口名 | 1（含 `\n` 污染） |
| 国家 | 2 |
| 操作类型 | 3 |
| 进港时间（Excel float） | 4 |
| 离港时间（Excel float） | 6（第7列） |

**xlrd 日期转换**：
```python
from xlrd import xldate_as_datetime
dt = xldate_as_datetime(excel_float, wb.datemode)
dt.strftime("%Y.%m.%d %H:%M")
```

## 港口映射（PORT_MANUAL）

```python
PORT_MANUAL = {
    "WUHAN":     "CNWHG-武汉港(Wuhangang)",
    "ZHOUSHAN":  "CNZOS-舟山(Zhoushan)",
    "BAHODOPI":  "IDBHP-巴霍多皮(Bahodopi)",
    "GRESIK":    "IDGRK-格雷西(Gresik)",
    "HONGKONG":  "HKHKG-香港(Hong Kong)",
    "QIN ZHOU":  "CNQZU-钦州港(Qinzhou)",
    "MAO MING":  "CNMGM-茂名(Maoming)",
    "MOROWALI":  "IDMOW-莫罗瓦利(Morowali)",
    "KENDARI":   "IDKEN-肯达里(Kendari)",
    "CHANGZHOU": "CNCZX-常州(Changzhou)",
    "YANGZHOU":  "CNYZU-扬州港(Yangzhou)",
}
```

COUNTRY_MAP：
```python
COUNTRY_MAP = {
    "WUHAN":"CN-中国","ZHOUSHAN":"CN-中国","QIN ZHOU":"CN-中国",
    "MAO MING":"CN-中国","CHANGZHOU":"CN-中国","YANGZHOU":"CN-中国",
    "BAHODOPI":"ID-印度尼西亚","GRESIK":"ID-印度尼西亚","MOROWALI":"ID-印度尼西亚","KENDARI":"ID-印度尼西亚",
    "HONGKONG":"HK-香港",
}
```

## 登船口岸映射

```python
EMBARK_PORT_MAP = {
    "WU HAN":"CNWHG-武汉港(Wuhangang)","WUHAN":"CNWHG-武汉港(Wuhangang)",
    "ZHOUSHAN":"CNZOS-舟山(Zhoushan)",
    "QIN ZHOU":"CNQZU-钦州港(Qinzhou)","QINZHOU":"CNQZU-钦州港(Qinzhou)",
    "YANG ZHOU":"CNYZU-扬州港(Yangzhou)","YANGZHOU":"CNYZU-扬州港(Yangzhou)",
}
```

## 船员国籍分布（15人）

| 国籍 | 人数 | 证件类型 | 证件号码来源 |
|------|------|----------|--------------|
| CN-中国 | 8 | 17-海员证 | col8（海员证号） |
| VN-越南 | 2 | 14-护照 | col13（护照号） |
| MM-缅甸 | 4 | 14-护照 | col13（护照号） |
| ID-印度尼西亚 | 1 | 14-护照 | col13（护照号） |

**关键教训**：外国船员的护照号在 col13（xlrd 0-based），不是海员证号 col8。之前版本全部填错。
