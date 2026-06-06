import json
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE_DIR, "products_full.json")
EXCEL_PATH = os.path.join(BASE_DIR, "mltzao_unique_combos.xlsx")

with open(JSON_PATH, "r", encoding="utf-8") as f:
    products = json.load(f)

print(f"Loaded {len(products)} products")

# ---- GPU core extraction ----
def extract_gpu_core(gpu_name):
    """Extract GPU core from full GPU name.
    E.g., '华硕 DUAL-RTX3050-O6G' → 'RTX 3050'
          '技嘉 RTX 5060 WINDFORCE OC 8G' → 'RTX 5060'
          '盈通 RTX5060Ti 8G 大地之神' → 'RTX 5060 Ti'
          '七彩虹 iGame RTX 5090 D v2 Vulcan W OC 24GB' → 'RTX 5090'
          '华硕 ATS-RX7650GRE-O8G' → 'RX 7650 GRE'
    """
    if not gpu_name:
        return "未知"

    name = gpu_name.upper()

    # NVIDIA patterns
    # RTX 5090 D / RTX5090D
    m = re.search(r'RTX\s*5090\s*D?', name)
    if m:
        return "RTX 5090"

    m = re.search(r'RTX\s*5080', name)
    if m:
        return "RTX 5080"

    m = re.search(r'RTX\s*5070\s*TI', name)
    if m:
        return "RTX 5070 Ti"

    m = re.search(r'RTX\s*5070', name)
    if m:
        return "RTX 5070"

    m = re.search(r'RTX\s*5060\s*TI', name)
    if m:
        return "RTX 5060 Ti"

    m = re.search(r'RTX\s*5060', name)
    if m:
        return "RTX 5060"

    m = re.search(r'RTX\s*5050', name)
    if m:
        return "RTX 5050"

    m = re.search(r'RTX\s*4090', name)
    if m:
        return "RTX 4090"

    m = re.search(r'RTX\s*4080\s*S?', name)
    if m:
        return "RTX 4080 SUPER"

    m = re.search(r'RTX\s*4070\s*TI\s*S', name)
    if m:
        return "RTX 4070 Ti SUPER"

    m = re.search(r'RTX\s*4070\s*S', name)
    if m:
        return "RTX 4070 SUPER"

    m = re.search(r'RTX\s*4070', name)
    if m:
        return "RTX 4070"

    m = re.search(r'RTX\s*4060\s*TI', name)
    if m:
        return "RTX 4060 Ti"

    m = re.search(r'RTX\s*4060', name)
    if m:
        return "RTX 4060"

    m = re.search(r'RTX\s*3060', name)
    if m:
        return "RTX 3060"

    m = re.search(r'RTX\s*3050', name)
    if m:
        return "RTX 3050"

    # AMD patterns
    m = re.search(r'RX\s*9070\s*XT', name)
    if m:
        return "RX 9070 XT"

    m = re.search(r'RX\s*9070', name)
    if m:
        return "RX 9070"

    m = re.search(r'RX\s*7900\s*XTX', name)
    if m:
        return "RX 7900 XTX"

    m = re.search(r'RX\s*7900\s*XT', name)
    if m:
        return "RX 7900 XT"

    m = re.search(r'RX\s*7800\s*XT', name)
    if m:
        return "RX 7800 XT"

    m = re.search(r'RX\s*7700\s*XT', name)
    if m:
        return "RX 7700 XT"

    m = re.search(r'RX\s*7650\s*GRE', name)
    if m:
        return "RX 7650 GRE"

    m = re.search(r'RX\s*7600', name)
    if m:
        return "RX 7600"

    m = re.search(r'RX\s*6750', name)
    if m:
        return "RX 6750 GRE"

    m = re.search(r'RX\s*6650', name)
    if m:
        return "RX 6650 XT"

    m = re.search(r'RX\s*6600', name)
    if m:
        return "RX 6600"

    return gpu_name  # fallback


# ---- RAM spec extraction ----
def extract_ram_spec(ram_name):
    """Extract RAM capacity + type + frequency.
    E.g., 'GEIL金邦 DDR4 16G 3200 内存条' → '16G DDR4 3200'
          '芝奇 烈焰枪 32GB（16G*2）DDR5 6000 黑色（C28/EXPO）' → '32G DDR5 6000'
          '佰维（BIWIN）时空行者DW100 48GB（24G*2）DDR5 6000' → '48G DDR5 6000'
    """
    if not ram_name:
        return "未知"

    # Extract capacity
    cap = None
    # Match patterns like "32GB（16G*2）" or "48GB（24G*2）" - total capacity first
    m = re.search(r'(\d+)\s*GB\s*（', ram_name)
    if m:
        cap = int(m.group(1))
    else:
        # Match "16G" or "32GB"
        m = re.search(r'(\d+)\s*G(?:B)?', ram_name)
        if m:
            cap = int(m.group(1))

    # Extract frequency first (needed for DDR type inference)
    freq = ""
    # Try to find frequency after DDR4/DDR5
    m = re.search(r'DDR[45]\s*(\d{4})', ram_name.upper())
    if m:
        freq = m.group(1)
    else:
        # Common frequencies: 2400, 2666, 3000, 3200, 3600, 4800, 5000, 5200, 5600, 6000, 6400, 7200, 8000
        m = re.search(r'(2[1-9]\d{2}|[3-9]\d{3})\s*(?:MHz)?', ram_name)
        if m:
            freq = m.group(1)

    # Extract DDR type
    ddr = ""
    m = re.search(r'(DDR[45])', ram_name.upper())
    if m:
        ddr = m.group(1)
    elif freq and int(freq) >= 4800:
        ddr = "DDR5"
    elif freq and int(freq) <= 4000:
        ddr = "DDR4"

    if cap and ddr and freq:
        return f"{cap}G {ddr} {freq}"
    elif cap and ddr:
        return f"{cap}G {ddr}"
    return ram_name


# ---- Process products ----
groups = {}
for p in products:
    specs = p.get("specs", {})
    cpu = specs.get("CPU", "").strip()
    gpu_full = specs.get("显卡", "").strip()
    ram = specs.get("内存", "").strip()

    if not cpu or not gpu_full:
        continue

    gpu_core = extract_gpu_core(gpu_full)
    ram_spec = extract_ram_spec(ram)

    key = (cpu, gpu_core, ram_spec)

    if key not in groups or p.get("price", 999999) < groups[key].get("price", 999999):
        groups[key] = p

print(f"Unique combos: {len(groups)}")

# Sort by price
results = sorted(groups.values(), key=lambda x: x.get("price", 0))

# ---- Generate Excel ----
wb = Workbook()
ws = wb.active
ws.title = "整机配置组合"

headers = [
    "序号", "配置名称", "编号", "价格(元)", "鲁大师跑分",
    "CPU型号", "GPU核心", "内存规格",
    "机箱", "CPU", "显卡(完整)", "主板", "内存(完整)", "硬盘", "CPU散热", "电源",
    "发货时效"
]

header_font = Font(bold=True, color="FFFFFF", size=11)
header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
    cell.border = thin_border

for idx, p in enumerate(results, 1):
    specs = p.get("specs", {})
    cpu = specs.get("CPU", "").strip()
    gpu_full = specs.get("显卡", "").strip()
    ram = specs.get("内存", "").strip()
    gpu_core = extract_gpu_core(gpu_full)
    ram_spec = extract_ram_spec(ram)

    row_data = [
        idx,
        p.get("name", ""),
        p.get("computerNo", ""),
        p.get("price", 0),
        p.get("score", 0),
        cpu,
        gpu_core,
        ram_spec,
        specs.get("机箱", ""),
        cpu,
        gpu_full,
        specs.get("主板", ""),
        ram,
        specs.get("硬盘", ""),
        specs.get("CPU散热", specs.get("散热器", "")),
        specs.get("电源", ""),
        p.get("deliveryDate", ""),
    ]
    for col, value in enumerate(row_data, 1):
        cell = ws.cell(row=idx + 1, column=col, value=value)
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if col == 4 and isinstance(value, (int, float)) and value > 0:
            cell.number_format = '#,##0'
        if col == 5 and isinstance(value, (int, float)) and value > 0:
            cell.number_format = '#,##0'

col_widths = [6, 40, 14, 12, 12, 28, 18, 20, 35, 28, 35, 28, 35, 28, 30, 40, 14]
for i, width in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = width

ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(results) + 1}"

wb.save(EXCEL_PATH)
print(f"Excel saved: {EXCEL_PATH}")
print(f"Total unique combos: {len(results)}")
