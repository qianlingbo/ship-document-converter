# 单证录入工作区

> 船舶单证录入自动化工具 —— 将 IMO Crew List + Port of Call 原始 Excel 文件，一键转换为海事局标准录入格式。

## 功能

- 🤖 **全自动转换**：原始文件 → 标准格式，无需手动填表
- 🌐 **多语言支持**：中英文船员名、中英文港口名自动识别
- 📋 **三表合一**：船员名单 + 物品清单 + 港口活动，一键生成
- 🔢 **标准编码**：使用海事局参数字段表（参数A/B/E），保证数据合规

## 环境要求

```bash
# Python 3.8+
python3 --version

# 安装依赖
pip install openpyxl
```

## 使用方法

### 方式一：命令行

```bash
cd ~/单证录入工作区

# 船员名单 + 港口
python3 scripts/单证录入核心.py input/crew_list.xlsx input/port_of_call.xlsx 2025航次报告

# 仅船员名单
python3 scripts/单证录入核心.py input/crew_list.xlsx

# 指定输出文件名（默认：单证录入_YYYYMMDD_HHMMSS.xlsx）
python3 scripts/单证录入核心.py crew.xlsx port.xlsx 我的航次
```

### 方式二：拖入 Hermes 对话

直接将 `IMO CREW LIST.xlsx` 和 `PORT OF CALL LIST.xlsx` 拖入飞书/Hermes 对话，AI 自动处理并推送结果文件。

## 输入文件示例

### IMO CREW LIST.xlsx

| No. | Family Name | Rank | Nationality | ... |
|-----|-------------|------|-------------|-----|
| 1 | NGUYEN VAN A | Master | VIETNAM | ... |
| 2 | 张三 | C/O | CHINA | ... |

### PORT OF CALL LIST.xlsx

| Voy. | Port | Arrival | Departure | ... |
|------|------|---------|-----------|-----|
| 1 | HITACHINAKA, JAPAN | 2024-01-15 | 2024-01-16 | ... |

## 输出文件

`output/2025航次报告.xlsx`，含三个 Sheet：

1. **船员名单**（船上非旅客人员清单）：16列，符合海事局格式
2. **物品清单**（船上非旅客人员物品清单）：每人1台计算机
3. **港口活动**（海事船岸活动信息）：含进离港时间、保安等级

## 目录结构

```
单证录入工作区/
├── SKILL.md                      # 技能详细文档（含所有规则）
├── README.md                     # 本文件
├── scripts/
│   └── 单证录入核心.py            # 核心脚本（单文件，800行）
├── templates/
│   └── 单证录入标准格式.xlsx       # 标准输出模板
├── references/                   # 参数映射（海事局参数字段表）
│   ├── nationality_map.json       # 国籍代码（248条）
│   ├── duty_map.json             # 职务代码（12条）
│   └── port_map.json             # 港口代码（1956条）
├── input/                        # 放原始文件
└── output/                       # 生成的文件
```

## 工作原理

```
原始文件 → 智能解析（自动找表头/列索引）
        → 标准化映射（国籍/职务/港口/日期）
        → 规则补全（fallback职务/国家提取）
        → Excel写入（3个Sheet）
        → 输出文件
```

## 技术栈

- **Python 3.8+**
- **openpyxl**：Excel 读写
- **pdfplumber**（可选）：PDF 支持

## 许可证 & 作者

MIT License  
Built with [Hermes](https://github.com/) + Claude MiniMax-M2.5
