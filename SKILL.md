# 单证录入技能

> 将船舶 IMO Crew List + Port of Call 转换为海事局标准录入格式

**版本**: v2.0.0 | **Python**: 3.8+ | **依赖**: `openpyxl`

## 快速开始

```bash
pip install openpyxl
python3 scripts/单证录入核心.py input/crew_list.xlsx [port_of_call.xlsx] [输出名]
```

## 目录结构

```
.
├── scripts/单证录入核心.py      # 核心脚本
├── templates/单证录入标准格式_v2.xlsx  # 输出模板
├── references/                  # 参数映射
│   ├── nationality_map.json     # 国籍代码 (248条)
│   ├── duty_map.json            # 职务代码 (12条)
│   └── port_map.json            # 港口代码 (1956条)
├── input/                       # 原始文件
└── output/                      # 输出文件
```

## 转换规则

### Sheet 1: 船员名单

| 字段 | 规则 |
|------|------|
| 姓名 | 中国船员=中文，外国船员=大写 |
| 船员职务 | 英文缩写自动映射，找不到按规则分配 |
| 国籍 | `CN-中国` / `VN-越南` 等格式 |
| 出生日期 | `YYYYMMDD` 格式 |
| 证件类型 | 中国=`17-海员证`，外国=`14-普通护照` |

**职务 fallback 规则**: 3个`56-高级值班水手` + 3个`66-高级值班机工`，其余按角色分配

### Sheet 2: 物品清单
固定每人一台计算机

### Sheet 3: 港口活动
进港时间随机 `00:00-12:00`，离港时间 `12:00-24:00`

## 输入格式

- **Crew List**: 表头含 `No.` + `Family name` + `Rank`
- **Port of Call**: 表头含 `Voy.` + `Port`

## 已知局限

- PDF 支持需根据实际布局调整
- 护照有效期/适任证书留空
