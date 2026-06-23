# GREEN MUNGUBA — Ports of Call Layout

**船名**: GREEN MUNGUBA | **IMO**: 9976496 | **Flag**: LIBERIA | **Port of Registry**: MONROVIA

## PDF 信息

- 文件：`Ports of call List.pdf`（1页）
- 文字提取：`pdftotext -layout` 完美提取，无需 OCR

## 港口数据（从新到旧，第1行=最新进港）

| # | 港口 | 进港 | 离港 | 国家 | 港口代码 |
|---|------|------|------|------|----------|
| 1 | NAPLES, ITALY | 13-Feb-2026 | 13-Feb-2026 | IT-意大利 | ITNAP-那不勒斯(NAPLES) |
| 2 | MONFALCONE, ITALY | 16-Feb-2026 | 18-Feb-2026 | IT-意大利 | ITMFA-蒙法尔科内(MONFALCONE) |
| 3 | GIBRALTAR, SPAIN | 24/02/2026 | 24/02/2026 | ES-西班牙 | GIBGI-直布罗陀(GIBRALTAR) |
| 4 | ITAQUI, BRAZIL | 6-Mar-2026 | 8-Apr-2026 | BR-巴西 | BRITQ-伊塔基(ITAQUI) |
| 5 | SINGAPORE, SINGAPORE | 14-May-2026 | 14-May-2026 | SG-新加坡 | SGSIN-新加坡(Singapore) |
| 6 | ZHOUSHAN, CHINA | 28-May-2026 | 29-May-2026 | CN-中国 | CNZOS-舟山(Zhoushan) |
| 7 | YANGZHONG, CHINA | 31-May-2026 | 5-Jun-2026 | CN-中国 | CNYZO-扬州、镇江(Yangzhong, Zhenjiang) |
| 8 | QINGDAO, CHINA | 7-Jun-2026 | 9-Jun-2026 | CN-中国 | CNQDP-青岛港(Qingdao) |
| 9 | TAICANG, CHINA | 11-Jun-2026 | 12-jun-2026 | CN-中国 | CNTAC-太仓(Taicang) |
| 10 | SHANGHAI, CHINA | 13-jun-2026 | 13-jun-2026 | CN-中国 | CNSHG-上海港(Shanghai) |

## PORT_MANUAL（实测，不需要子串匹配）

```python
PORT_MANUAL = {
    "ZHOUSHAN":   ("CNZOS", "舟山(Zhoushan)"),
    "YANGZHONG":  ("CNYZO", "扬州、镇江(Yangzhong, Zhenjiang)"),
    "QINGDAO":    ("CNQDP", "青岛港(Qingdao)"),
    "TAICANG":    ("CNTAC", "太仓(Taicang)"),
    "SHANGHAI":   ("CNSHG", "上海港(Shanghai)"),
    "NAPLES":     ("ITNAP", "那不勒斯(NAPLES)"),
    "MONFALCONE": ("ITMFA", "蒙法尔科内(MONFALCONE)"),
    "GIBRALTAR":  ("GIBGI", "直布罗陀(GIBRALTAR)"),
    "ITAQUI":     ("BRITQ", "伊塔基(ITAQUI)"),
    "SINGAPORE":  ("SGSIN", "新加坡(Singapore)"),
}
```

## 日期格式说明

- PDF 中存在混合格式：`13-Feb-2026`（`dd-Mmm-yyyy`）和 `24/02/2026`（`dd/mm/yyyy`）
- `pdftotext -layout` 均可正确提取
- POC 顺序是**从新到旧**，不要反转

## POC_NATIONALITY_MAP

```python
POC_NATIONALITY_MAP = {
    "ITALY": "IT-意大利",
    "SPAIN": "ES-西班牙",
    "BRAZIL": "BR-巴西",
    "SINGAPORE": "SG-新加坡",
    "CHINA": "CN-中国",
    "MONROVIA": "LR-利比里亚",   # 船籍港，用于当前港口
}
```
