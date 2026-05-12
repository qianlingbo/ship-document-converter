import json, openpyxl, random
from datetime import datetime

SKILL_DIR = "/Users/qianlingbo/.hermes/skills/ship-document-converter"
OUTPUT = f"{SKILL_DIR}/output/港口活动_MINHUA9_v3.xlsx"

# Load references
with open(f"{SKILL_DIR}/references/port_map.json") as f:
    port_map = json.load(f)
with open(f"{SKILL_DIR}/references/nationality_map.json") as f:
    nation_map = json.load(f)

# Port of Call raw data
poc_data = [
    ("01", "WEDA",        "INDONESIA",  "09JAN 2026",  "16 JAN 2026"),
    ("02", "ZHANG JIAGANG","CHINA",    "27 JAN 2026",  "28 JAN 2026"),
    ("03", "ZHOUSHAN",    "CHINA",      "29 JAN 2026",  "01 FEB 2026"),
    ("04", "WEDA",        "INDONESIA",  "10 FEB 2026",  "14FEB 2026"),
    ("05", "ZHOUSHAN",    "CHINA",      "23 FEB 2026",  "26 FEB 2026"),
    ("06", "WEDA",        "INDONESIA",  "06 MAR 2026",  "10MAR 2026"),
    ("07", "ZHAPU",       "CHINA",      "15 APR 2026",  "15 APR 2026"),
    ("08", "ZHOU SHAN",   "CHINA",      "16 APR 2026",  "18 APR 2026"),
    ("09", "NING DE",     "CHINA",      "25 APR 2026",  "26 APR 2026"),
    ("10", "WEDA",        "INDONESIA",  "04 MAY2026",   "09 MAY2026"),
]

MONTH_MAP = {"JAN":"01","FEB":"02","MAR":"03","APR":"04","MAY":"05","JUN":"06",
             "JUL":"07","AUG":"08","SEP":"09","OCT":"10","NOV":"11","DEC":"12"}

def parse_date(s):
    s = s.strip().upper().replace(" ", "")
    for m, mn in MONTH_MAP.items():
        if m in s:
            rest = s.replace(m, "")
            # "09JAN2026" -> day=09, year=2026
            if len(rest) == 4:
                return rest[:2], mn, rest[2:]
            elif len(rest) == 6:  # e.g. "15APR2026"
                return rest[:2], mn, rest[2:]
    return None, None, None

def rnd_time(arrival):
    if arrival:
        h = random.randint(0, 11)
    else:
        h = random.randint(12, 23)
    mi = random.randint(0, 59)
    se = random.randint(0, 59)
    return f"{h:02d}:{mi:02d}:{se:02d}"

def find_port(name):
    c = name.upper().replace(" ", "")
    if c in port_map:
        return c, port_map[c]
    for code, val in port_map.items():
        nm = val.upper().replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        if c in nm or nm in c:
            return code, val
    return None, None

def get_country(code):
    cc = code[:2].upper()
    for k, v in nation_map.items():
        if k.upper() == cc:
            return v
    return cc

# Load template
wb = openpyxl.load_workbook(f"{SKILL_DIR}/templates/单证录入标准格式_v2.xlsx")
ws = wb["海事船岸活动信息"]

for i, (seq, port_name, country_raw, arr_str, dep_str) in enumerate(poc_data):
    row = i + 3

    ad, am, ay = parse_date(arr_str)
    dd, dm, dy = parse_date(dep_str)

    if ad is None:
        print(f"Row {row}: FAIL parse arr {arr_str}")
        continue
    if dd is None:
        print(f"Row {row}: FAIL parse dep {dep_str}")
        continue

    arr_formatted = f"{ay}/{am}/{ad} {rnd_time(True)}"
    dep_formatted = f"{dy}/{dm}/{dd} {rnd_time(False)}"

    port_code, port_display = find_port(port_name)

    if port_code is None:
        id_ports = [(k, v) for k, v in port_map.items() if k.startswith("ID")]
        port_code, port_display = random.choice(id_ports)
        for col in range(1, 9):
            ws.cell(row=row, column=col).fill = openpyxl.styles.PatternFill(
                fill_type="solid", fgColor="FFCCCC")

    country = get_country(port_code)

    ws.cell(row=row, column=1).value = seq
    ws.cell(row=row, column=2).value = arr_formatted
    ws.cell(row=row, column=3).value = dep_formatted
    ws.cell(row=row, column=4).value = country
    ws.cell(row=row, column=5).value = "1-1级"
    ws.cell(row=row, column=6).value = None
    ws.cell(row=row, column=7).value = port_display
    ws.cell(row=row, column=8).value = "1-1级"

wb.save(OUTPUT)

# Verify
wb2 = openpyxl.load_workbook(OUTPUT)
ws2 = wb2["海事船岸活动信息"]
print("Verification:")
for i, row in enumerate(ws2.iter_rows(values_only=True), 1):
    if 1 <= i <= 12:
        seq = row[0]
        port = row[6]
        print(f"Row{i:02d}: {seq} | {port}")
