"""
NC / CIC Site Code Cross-Analysis  (v2 – aligned with Review logic)
───────────────────────────────────────────────────────────────────
Objective
  Find items present in CIC that also have a 'Next PO' in NC.
  Determine whether the NC PO is truly needed by checking:
    1. NC Supply (excl. this PO) = On Hand + (On Order − Next PO Qty)
    2. NC Shortfall = max(0, Open Sales − NC Supply)
    3. CIC Available = CIC On Hand − CIC Open Sales   (surplus)
  Then compare Shortfall vs. CIC Available to classify:
    • PO NOT Needed       – CIC covers or NC already has enough
    • PO Partially Needed – CIC covers some of the shortfall
    • PO Needed           – CIC cannot help at all

Output: data/NC_CIC_PO_Analysis.xlsx
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

# ── 1. Load workbook ────────────────────────────────────────────────
INPUT_FILE  = os.path.join(os.path.dirname(__file__), "..", "data", "Forecast_Analysis.xlsx")
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "..", "data", "NC_CIC_PO_Analysis.xlsx")

wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
ws = wb["Sheet"]

# ── 2. Parse NC and CIC rows ───────────────────────────────────────
nc_items = {}
cic_items = {}

for row in ws.iter_rows(min_row=2, values_only=False):
    site_code = row[0].value       # A: Site Code
    item_code = str(row[2].value)  # C: Item Code
    data = {
        "item_desc":              row[3].value,          # D
        "supplier":               row[4].value,          # E
        "safety_stock":           row[15].value or 0,    # P
        "on_hand":                row[16].value or 0,    # Q
        "on_order":               row[17].value or 0,    # R
        "all_open_sales":         row[18].value or 0,    # S
        "current_mo_open_sales":  row[19].value or 0,    # T
        "mo1_open_sales":         row[20].value or 0,    # U
        "mo2_open_sales":         row[21].value or 0,    # V
        "open_estimates":         row[22].value or 0,    # W
        "next_po_date":           row[24].value,         # Y
        "next_po_qty":            row[25].value or 0,    # Z
    }
    if site_code == "NC":
        nc_items[item_code] = data
    elif site_code == "CIC":
        cic_items[item_code] = data

# ── 3. Identify common items with Next PO in NC ────────────────────
common_items = set(nc_items) & set(cic_items)

results = []
for item_code in sorted(common_items):
    nc  = nc_items[item_code]
    cic = cic_items[item_code]

    # Skip if NC has no Next PO
    if nc["next_po_date"] is None and nc["next_po_qty"] == 0:
        continue

    # ── Key calculation (Python logic for status) ─────────────
    nc_supply_excl_po = nc["on_hand"] + (nc["on_order"] - nc["next_po_qty"])
    nc_shortfall      = max(0, nc["all_open_sales"] - nc_supply_excl_po)
    cic_available     = cic["on_hand"] - cic["all_open_sales"]  # surplus

    # Excel Formula Strings for the report
    # Columns:
    # D: NC On Hand, E: NC On Order, F: NC Next PO Qty
    # H: NC Open Sales, I: NC Supply (excl PO), J: NC Shortfall
    # K: CIC On Hand, M: CIC Open Sales, N: CIC Available
    
    # These will be dynamically adjusted with the correct row index during write_sheet
    formula_nc_supply = '=D{row} + (E{row} - F{row})'
    formula_nc_shortfall = '=MAX(0, H{row} - I{row})'
    formula_cic_available = '=K{row} - M{row}'

    # ── Classify ────────────────────────────────────────────────
    if nc_shortfall == 0:
        status = "PO NOT Needed"
        reason = "NC has enough on hand + on order (excl. this PO) to cover open sales"
    elif cic_available >= nc_shortfall:
        status = "PO NOT Needed"
        reason = f"CIC surplus of {cic_available} can cover NC shortfall of {nc_shortfall}"
    elif cic_available > 0:
        still_need = nc_shortfall - cic_available
        status = "PO Partially Needed"
        reason = f"CIC covers {cic_available} of {nc_shortfall} shortfall; still need {still_need}"
    else:
        status = "PO Needed"
        reason = "CIC has no available surplus to transfer"

    results.append({
        "item_code":          item_code,
        "item_desc":          nc["item_desc"],
        "supplier":           nc["supplier"],
        # NC columns
        "nc_on_hand":         nc["on_hand"],
        "nc_on_order":        nc["on_order"],
        "nc_open_sales":      nc["all_open_sales"],
        "nc_safety_stock":    nc["safety_stock"],
        "nc_supply_excl_po_val": nc_supply_excl_po, # Keep val for python checks
        "nc_supply_excl_po":  formula_nc_supply,
        "nc_shortfall":       formula_nc_shortfall,
        "nc_next_po_date":    nc["next_po_date"],
        "nc_next_po_qty":     nc["next_po_qty"],
        # CIC columns
        "cic_on_hand":        cic["on_hand"],
        "cic_on_order":       cic["on_order"],
        "cic_open_sales":     cic["all_open_sales"],
        "cic_available":      formula_cic_available,
        # Verdict
        "status":             status,
        "reason":             reason,
    })

# ── 4. Summary stats ───────────────────────────────────────────────
total         = len(results)
not_needed    = sum(1 for r in results if r["status"] == "PO NOT Needed")
partial       = sum(1 for r in results if r["status"] == "PO Partially Needed")
still_needed  = sum(1 for r in results if r["status"] == "PO Needed")

print("=" * 65)
print("  NC / CIC  Next-PO Cross-Analysis  (v2)")
print("=" * 65)
print(f"  Total NC items:                   {len(nc_items):>6}")
print(f"  Total CIC items:                  {len(cic_items):>6}")
print(f"  Items in both NC & CIC:           {len(common_items):>6}")
print(f"  … with Next PO in NC:             {total:>6}")
print(f"  ✅ PO NOT Needed:                   {not_needed:>4}")
print(f"  🟡 PO Partially Needed:             {partial:>4}")
print(f"  ⚠️  PO Needed:                       {still_needed:>4}")
print("=" * 65)

# ── 5. Write to Excel ──────────────────────────────────────────────
out_wb = openpyxl.Workbook()

HEADER_FONT  = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
HEADER_FILL  = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
GREEN_FILL   = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
YELLOW_FILL  = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RED_FILL     = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREEN_FONT   = Font(name="Calibri", color="006100")
YELLOW_FONT  = Font(name="Calibri", color="9C6500")
RED_FONT     = Font(name="Calibri", color="9C0006")
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

HEADERS = [
    "Item Code", "Item Description", "Supplier",
    "NC On Hand", "NC On Order", "NC Next PO Qty", "NC Next PO Date",
    "NC Open Sales", "NC Supply (excl PO)", "NC Shortfall",
    "CIC On Hand", "CIC On Order", "CIC Open Sales", "CIC Available",
    "PO Status", "Reason",
]

COLUMN_WIDTHS = {
    1: 12, 2: 38, 3: 30,
    4: 13, 5: 13, 6: 15, 7: 16,
    8: 14, 9: 19, 10: 14,
    11: 13, 12: 13, 13: 14, 14: 14,
    15: 22, 16: 55,
}

STATUS_STYLES = {
    "PO NOT Needed":       (GREEN_FILL,  GREEN_FONT),
    "PO Partially Needed": (YELLOW_FILL, YELLOW_FONT),
    "PO Needed":           (RED_FILL,    RED_FONT),
}

def write_sheet(sheet, title, items):
    sheet.title = title
    for col_idx, header in enumerate(HEADERS, 1):
        cell = sheet.cell(row=1, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER
    for col_idx, w in COLUMN_WIDTHS.items():
        sheet.column_dimensions[get_column_letter(col_idx)].width = w
    sheet.freeze_panes = "A2"

    for row_idx, r in enumerate(items, 2):
        vals = [
            r["item_code"], r["item_desc"], r["supplier"],
            r["nc_on_hand"], r["nc_on_order"], r["nc_next_po_qty"],
            r["nc_next_po_date"], r["nc_open_sales"],
            r["nc_supply_excl_po"].format(row=row_idx) if isinstance(r["nc_supply_excl_po"], str) else r["nc_supply_excl_po"], 
            r["nc_shortfall"].format(row=row_idx) if isinstance(r["nc_shortfall"], str) else r["nc_shortfall"],
            r["cic_on_hand"], r["cic_on_order"], r["cic_open_sales"],
            r["cic_available"].format(row=row_idx) if isinstance(r["cic_available"], str) else r["cic_available"],
            r["status"], r["reason"],
        ]
        for col_idx, val in enumerate(vals, 1):
            cell = sheet.cell(row=row_idx, column=col_idx, value=val)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
        # Colour status
        fill, font = STATUS_STYLES.get(r["status"], (None, None))
        if fill:
            sheet.cell(row=row_idx, column=15).fill = fill
            sheet.cell(row=row_idx, column=15).font = font

# ─── All items in one detail sheet ──────────────────────────────────
ws1 = out_wb.active
write_sheet(ws1, "PO Review Detail", results)

# ─── Action sheet: items to cancel ──────────────────────────────────
ws2 = out_wb.create_sheet("Action – Cancel POs")
cancel_headers = ["Item Code", "Item Description", "NC Next PO Qty", "NC Next PO Date", "Reason"]
for col_idx, h in enumerate(cancel_headers, 1):
    cell = ws2.cell(row=1, column=col_idx, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = THIN_BORDER
ws2.column_dimensions["A"].width = 12
ws2.column_dimensions["B"].width = 38
ws2.column_dimensions["C"].width = 15
ws2.column_dimensions["D"].width = 16
ws2.column_dimensions["E"].width = 55
ws2.freeze_panes = "A2"
cancel_items = [r for r in results if r["status"] == "PO NOT Needed"]
for row_idx, r in enumerate(cancel_items, 2):
    for col_idx, val in enumerate([
        r["item_code"], r["item_desc"], r["nc_next_po_qty"],
        r["nc_next_po_date"], r["reason"]
    ], 1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=val)
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")

# ─── Summary sheet ──────────────────────────────────────────────────
ws3 = out_wb.create_sheet("Summary")
summary_data = [
    ("NC vs CIC – Next PO Review", ""),
    ("", ""),
    ("Metric", "Count"),
    ("Total NC items with PO that also exist in CIC", total),
    ("PO NOT Needed (can cancel/remove)", not_needed),
    ("PO Partially Needed (reduce qty)", partial),
    ("PO Needed (CIC cannot help)", still_needed),
]
for row_idx, (label, value) in enumerate(summary_data, 1):
    c1 = ws3.cell(row=row_idx, column=1, value=label)
    c2 = ws3.cell(row=row_idx, column=2, value=value if value != "" else None)
    c1.border = THIN_BORDER
    c2.border = THIN_BORDER
    if row_idx == 1:
        c1.font = Font(name="Calibri", bold=True, size=14)
    elif row_idx == 3:
        c1.font = HEADER_FONT; c1.fill = HEADER_FILL
        c2.font = HEADER_FONT; c2.fill = HEADER_FILL
    c1.alignment = Alignment(horizontal="left")
    c2.alignment = Alignment(horizontal="center")
ws3.column_dimensions["A"].width = 48
ws3.column_dimensions["B"].width = 14

out_wb.save(OUTPUT_FILE)
print(f"\n📄 Report saved → {os.path.abspath(OUTPUT_FILE)}")

# ── 6. Console detail ──────────────────────────────────────────────
for label, status_val in [
    ("✅ PO NOT Needed",       "PO NOT Needed"),
    ("🟡 PO Partially Needed", "PO Partially Needed"),
    ("⚠️  PO Needed",           "PO Needed"),
]:
    group = [r for r in results if r["status"] == status_val]
    if not group:
        continue
    print(f"\n{label} ({len(group)} items)")
    print("-" * 130)
    for r in group:
        print(f"  {r['item_code']:<10} {(r['item_desc'] or '')[:34]:<36} {r['reason']}")
