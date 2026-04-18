---
name: ship-document-converter
description: 将船舶 IMO Crew List Excel 文件转换为海事局单证录入标准格式（xlsx）。使用时机：用户上传 crew list Excel、提到"单证录入"、"船员名单转换"、"IMO Crew List"、"生成海事局文件"。支持 .xls / .xlsx，输入原始表格自动输出标准三表（船员清单+物品清单+港口活动）。
license: MIT
compatibility: python 3.8+
metadata:
  author: Qian Lingbo
  version: 3.0.0
  language: zh
  tags: [maritime, excel, document-conversion, 海事, 单证录入]
---

# 单证录入技能

> 将船舶 IMO Crew List 转换为海事局标准录入格式（xlsx）

**版本**: v3.0.0 | **Python**: 3.8+ | **依赖**: `openpyxl`, `xlrd`

## 目录结构

```
ship-document-converter/
├── SKILL.md                      # 本文件
├── scripts/
│   └── 单证录入核心.py            # 核心脚本（885行）
├── references/
│   ├── nationality_map.json      # 国籍代码 (248条)
│   ├── duty_map.json            # 职务代码 (12条)
│   ├── port_map.json            # 港口代码 (1956条)
│   └── conversion-rules.md      # 完整转换规则与坑点详解
├── templates/
│   └── 单证录入标准格式_v2.xlsx   # 输出模板（含6个sheet）
├── input/                        # 原始文件
└── output/                        # 生成的文件
```

## 快速开始

### 命令行用法

```bash
# 安装依赖
pip install openpyxl xlrd

# 运行转换
python3 scripts/单证录入核心.py input/crew_list.xlsx [port_of_call.xlsx] [输出名]
```

### 输入文件

- **Crew List**: 表头含 `No.` + `Family name` + `Rank`
- 自动检测两种列布局（有序号列 / 无序号列）
- 支持 .xls（Excel 97-2003）和 .xlsx 格式

### 输出文件

`output/单证录入_YYYYMMDD_HHMMSS.xlsx`，含三个 Sheet：

| Sheet | 内容 |
|-------|------|
| 船员名单 | 船上非旅客人员清单，16列，符合海事局格式 |
| 物品清单 | 每人 1 台计算机，固定格式 |
| 港口活动 | 进离港时间、保安等级 |

---

## ⚠️ 关键规则

### 必须基于模板写入

**绝对不要**用 `openpyxl.Workbook()` 创建空白文件！必须先复制模板：

```python
import shutil, openpyxl

# ✅ 正确
shutil.copy("templates/单证录入标准格式_v2.xlsx", "output/单证录入_xxx.xlsx")
wb = openpyxl.load_workbook("output/单证录入_xxx.xlsx")

# ❌ 错误（会丢失所有sheet和格式）
wb = openpyxl.Workbook()
```

### 船员名单转换规则

| 字段 | 规则 |
|------|------|
| 姓名 | 中国船员=中文；外国船员=大写英文 |
| 船员职务 | 英文缩写自动映射（见 `references/conversion-rules.md`） |
| 国籍 | `CN-中国` / `VN-越南` 格式 |
| 出生日期 | `YYYYMMDD` 格式 |
| 证件类型 | 中国=`17-海员证`，外国=`14-普通护照` |

### 物品清单

原始 IMO 船员清单通常不含物品数据，每人固定登记一条：

| 列 | 字段 | 值 |
|----|------|-----|
| 1 | 序号 | 船员序号 |
| 2 | 证件类型 | `17-海员证`（中国）或 `14-普通护照`（外国） |
| 3 | 证件号码 | 海员证号或护照号 |
| 4 | 物品类型 | `0100` |
| 5 | 物品名称 | `计算机` |
| 6 | 物品数量 | `1` |
| 7 | 数量单位 | `001` |

---

## Feishu 文件发送

不要用 `send_message_tool`（非 Telegram 平台会丢弃附件）。正确方式：

```python
import requests, os

APP_ID = "cli_a952c98ec13a9bca"
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "***")
CHAT_ID = "oc_9d8f4df4139fb63513d74ee2ef17df8d"

# 1. 获取 token
resp = requests.post("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
    json={"app_id": APP_ID, "app_secret": APP_SECRET},
    proxies={"http": "socks5h://localhost:7897", "https": "socks5h://localhost:7897"})
token = resp.json()["tenant_access_token"]

# 2. 上传文件
with open(output_file, "rb") as f:
    files = {"file": (filename, f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    data = {"file_name": filename, "file_size": str(os.path.getsize(output_file)), "file_type": "xlsx"}
    resp = requests.post("https://open.feishu.cn/open-apis/im/v1/files",
        headers={"Authorization": f"Bearer {token}"}, data=data, files=files,
        proxies={"http": "socks5h://localhost:7897", "https": "socks5h://localhost:7897"})
file_key = resp.json()["data"]["file_key"]

# 3. 发送消息
requests.post(
    "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
    json={"receive_id": CHAT_ID, "msg_type": "file", "content": f'{{"file_key": "{file_key}"}}'},
    proxies={"http": "socks5h://localhost:7897", "https": "socks5h://localhost:7897"})
```

---

## 常见问题速查

| 问题 | 原因 | 解决 |
|------|------|------|
| 日期变成数字（如 45789） | 输入文件是 .xls 格式 | 使用 xlrd 读取，Excel 序列号→日期用 `datetime(1899,12,30)+timedelta(days=int(serial))` |
| 证件号码含日期前缀 | 直接取了整段值 | 用 `extract_cert_number()` 分离，格式 `"DD/Mon/YYYY NUMBER"` |
| 舟山匹配到上海 | `SHA` 是 `ZHOUSHAN` 子串 | 使用 `SPECIAL_PORT_OVERRIDE` 手动映射表 |
| 职务 `AB` / `CDT` 未识别 | 不在 duty_map.json | 在代码中硬编码 fallback 映射 |

详细规则、混合字段解析、职务代码表见 `references/conversion-rules.md`。

---

## 示例对话

**用户**: "帮我把这个 IMO crew list 转成海事局格式"  
**操作**: 确认文件格式 → 运行 `scripts/单证录入核心.py` → 输出文件到 `output/` → 发送至飞书

**用户**: "上传了船员名单，生成单证录入"  
**操作**: 读取 input/ 目录最新文件 → 检测 .xls 或 .xlsx → 执行转换 → 返回结果

---

## 已知局限

- PDF 输入需要根据实际布局调整（当前版本未实现）
- 护照有效期、适任证书字段留空
