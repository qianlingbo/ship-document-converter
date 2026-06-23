# WEN DE PORT CALL LIST（10港口）

## 文件信息
- 格式：扫描件 PDF（pypdfium2 渲染 + OCR）
- 港口数量：10个
- 到达港：ZHOUSHAN（最新港口）

## 原始数据（用户口述/OCR整理）

| # | PORT | COUNTRY | ARR | DEP |
|---|------|---------|-----|-----|
| 1 | ZHANGJIAGANG | CHINA | 2026/6/5 | 2026/5/30 |
| 2 | ZHOUSHAN | CHINA | 2026/5/15 | 2026/5/24 |
| 3 | SINGAPORE | SINGAPORE | 2026/2/10 | 2026/3/25 |
| 4 | PARANAGUA | BRAZIL | 2026/1/26 | 2026/2/1 |
| 5 | ZHANGZHOU | CHINA | 2026/5/16 | 2026/5/29 |
| 6 | ZHANJIANG | CHINA | 2026/2/14 | 2026/4/5 |
| 7 | SINGAPORE | SINGAPORE | 2026/1/26 | 2026/2/7 |
| 8 | SINGAPORE | SINGAPORE | 2026/1/7 | 2026/1/11 |
| 9 | PORT LINCOLN | AUSTRALIA | 2025/12/26 | 2026/1/7 |
| 10 | WALLAROO | AUSTRALIA | 2025/12/7 | 2025/12/2 |

## ⚠️ 日期顺序存疑（需用户确认）

以下港口 ARR > DEP（到达日晚于离港日），可能是 ARR/DEP 填反了：
- WALLAROO: ARR=2025/12/7, DEP=2025/12/2
- ZHANJIANG: ARR=2026/2/14, DEP=2026/4/5
- PARANAGUA: ARR=2026/1/26, DEP=2026/2/1
- SINGAPORE: ARR=2026/1/26, DEP=2026/2/7（第三个SGP）

需用户确认后再录入。

## 港口代码映射

```python
{
    "ZHANGJIAGANG": "CNZJG-张家港(Zhangjiagang)",
    "ZHOUSHAN":     "CNZOS-舟山(Zhoushan)",
    "SINGAPORE":    "SGSIN-新加坡(Singapore)",
    "PARANAGUA":    "BRPNG-巴拉那瓜(Paranagua)",  # port_map.json 无PARANAGUA，用 fallback
    "ZHANGZHOU":    "CNZZU-漳州(Zhangzhou)",
    "ZHANJIANG":    "CNZNG-湛江港(Zhanjianggang)",
    "PORT LINCOLN": "AUPOL-澳大利亚林肯港(Port Lincoln)",
    "WALLAROO":     "AUWAL-澳大利亚瓦拉鲁(Wallaroo)",
}
```
