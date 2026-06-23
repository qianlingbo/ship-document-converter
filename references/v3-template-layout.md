# v3 模板结构说明（oneTableDeclaration 17-sheet xlsm）

## 基本信息

- 文件：`单证录入标准格式_v3.xlsm`
- 宏：保留（keep_vba=True）
- 核心 Sheet：3个（船员/物品/海事），其余14个留空

## Sheet 结构

| # | Sheet名称 | 数据行规则 | 关键列 |
|---|---------|---------|-------|
| 1 | 船上非旅客人员清单 | Row3=crew#1 | B=姓名, D=职务, E=国籍, H=证件类型, I=证件号码 |
| 2 | 旅客清单 | Row3=passenger#1 | - |
| 3 | 供退物料清单 | Row3=item#1 | - |
| 4 | 船上非旅客人员物品清单 | Row3=item#1 | B=证件类型, C=证件号码, D=物品类型 |
| 5 | 船用物品清单 | Row3=item#1 | - |
| 6 | 前十港信息 | Row3=port#1 | - |
| 7 | 危险品信息 | Row3=item#1 | - |
| 8 | 船舶证书信息 | Row3=cert#1 | - |
| 9 | 压舱水详细信息 | Row3=item#1 | - |
| 10 | **海事船岸活动信息** | Row3=port#1 | B=进港时间, C=离港时间, D=国家, G=港口 |
| 11 | 沿海空箱信息 | Row3=item#1 | - |
| 12 | 压舱水报告单信息 | Row3=item#1 | - |
| 13 | 压舱水装载信息 | Row3=item#1 | - |
| 14 | 压舱水更换信息表信息 | Row3=item#1 | - |
| 15 | 压舱水排放信息表信息 | Row3=item#1 | - |
| 16 | 随船人员清单 | Row3=person#1 | - |
| 17 | 参数 | 下拉选项数据 | A=船员国籍, B=船员职务, C=证件类型, D=性别, E=装载港 |

## v2 vs v3 区别

| | v2模板 | v3模板 |
|-|--------|--------|
| Sheet数 | 6 | 17 |
| 数据起始行 | Row2（无说明行） | Row3（Row2=说明行） |
| 格式 | xlsx | xlsm（保留VBA宏） |
| 宏支持 | 无 | 有（下拉联动） |

## 复制写入脚本

```python
import openpyxl, shutil
from pathlib import Path
from copy import copy

tpl = Path("~/.hermes/skills/ship-document-converter/templates/单证录入标准格式_v3.xlsm")
src = Path("output/单证录入_xxx.xlsx")
out = src.parent / f"oneTableDeclaration_{船名}_{日期}.xlsm"

shutil.copy(tpl, out)
wb_src = openpyxl.load_workbook(src, data_only=True)
wb_dst = openpyxl.load_workbook(out, keep_vba=True)

SRC_START = 2   # v2 xlsx 数据从Row2起
DST_START = 3   # v3 xlsm 数据从Row3起（偏移+1）

for sn in ["船上非旅客人员清单", "船上非旅客人员物品清单", "海事船岸活动信息"]:
    ws_src = wb_src[sn]
    ws_dst = wb_dst[sn]
    for r in range(SRC_START, ws_src.max_row + 1):
        vals = [ws_src.cell(r, c).value for c in range(1, (ws_src.max_column or 16) + 1)]
        if not any(v is not None for v in vals):
            continue
        target_r = r - SRC_START + DST_START
        for ci, val in enumerate(vals):
            src_c = ws_src.cell(r, ci + 1)
            dst_c = ws_dst.cell(target_r, ci + 1)
            dst_c.value = val
            if src_c.has_style:
                try:
                    dst_c.font = copy(src_c.font)
                    dst_c.fill = copy(src_c.fill)
                    dst_c.border = copy(src_c.border)
                    dst_c.alignment = copy(src_c.alignment)
                    dst_c.number_format = src_c.number_format
                except: pass
    wb_dst.save(out)
```
