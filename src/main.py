"""
NC / CIC Site Code Cross-Analysis  (v3 – Transfer Order aware)
───────────────────────────────────────────────────────────────────
Objective
  Find items present in CIC that also have a 'Next PO' in NC.
  Determine whether the NC PO is truly needed by checking:
    1. NC Supply (excl. this PO) = On Hand + (On Order − Next PO Qty)
       ⚠ If the "Next PO" is actually a Transfer Order (NC-TO),
         do NOT subtract the qty from On Order.
    2. NC Shortfall = max(0, Open Sales − NC Supply)
    3. CIC Available = CIC On Hand − CIC Open Sales   (surplus)
  Then compare Shortfall vs. CIC Available to classify:
    • PO NOT Needed       – CIC covers or NC already has enough
    • PO Partially Needed – CIC covers some of the shortfall
    • PO Needed           – CIC cannot help at all

  The "Next PO" in Forecast_Analysis may be a Transfer Order from
  CIC → NC. We cross-reference with Open Supply Orders.xlsx to
  identify these (Order # starting with NC-TO).

Output: data/NC_CIC_PO_Analysis.xlsx
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

# ── 1. Load workbooks ───────────────────────────────────────────────
INPUT_FILE        = os.path.join(os.path.dirname(__file__), "..", "data", "Forecast Analysis.xlsx")
SUPPLY_ORDERS_FILE = os.path.join(os.path.dirname(__file__), "..", "data", "Open Supply Orders.xlsx")
OUTPUT_FILE       = os.path.join(os.path.dirname(__file__), "..", "data", "NC_CIC_PO_Analysis.xlsx")

wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
ws = wb["Sheet"]

# ── 1b. Load Open Supply Orders ────────────────────────────────────
wb_supply = openpyxl.load_workbook(SUPPLY_ORDERS_FILE, data_only=True)
ws_supply = wb_supply[wb_supply.sheetnames[0]]

# Build lookup: item_code → { to_orders: [...], po_orders: [...] }
supply_orders = {}
for row in ws_supply.iter_rows(min_row=2, values_only=True):
    site      = row[3]          # D: Site
    if site != "NC":
        continue
    item_code = str(row[2])     # C: Item
    order_num = str(row[4] or "")  # E: Order #
    order_date = row[9]         # J: Order Date
    release_qty = row[10] or 0  # K: Release Qty
    is_to = order_num.startswith("NC-TO")

    if item_code not in supply_orders:
        supply_orders[item_code] = {"to_orders": [], "po_orders": []}

    order_info = {
        "order_num": order_num,
        "order_date": order_date,
        "release_qty": release_qty,
    }
    if is_to:
        supply_orders[item_code]["to_orders"].append(order_info)
    else:
        supply_orders[item_code]["po_orders"].append(order_info)

wb_supply.close()


def _normalise_date(d):
    """Return a date object regardless of whether d is datetime, date, or 'YYYY/MM/DD' string."""
    if d is None:
        return None
    if isinstance(d, datetime):
        return d.date()
    if hasattr(d, "date"):  # date-like
        return d
    if isinstance(d, str):
        for fmt in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                return datetime.strptime(d, fmt).date()
            except ValueError:
                continue
    return None


def _match_order(item_code, next_po_date, next_po_qty):
    """
    Try to match the Forecast_Analysis Next PO against Open Supply Orders.
    Returns (order_type, verified, matched_order_num).
    """
    if item_code not in supply_orders:
        return ("PO", False, "")

    orders = supply_orders[item_code]
    fc_date = _normalise_date(next_po_date)

    # Try Transfer Orders first
    for o in orders["to_orders"]:
        o_date = _normalise_date(o["order_date"])
        if o_date == fc_date and o["release_qty"] == next_po_qty:
            return ("Transfer Order", True, o["order_num"])

    # Try PO orders
    for o in orders["po_orders"]:
        o_date = _normalise_date(o["order_date"])
        if o_date == fc_date and o["release_qty"] == next_po_qty:
            return ("PO", True, o["order_num"])

    # No exact match – check if item has ANY transfer orders at all
    # and the next PO qty matches a TO qty (date may differ slightly)
    for o in orders["to_orders"]:
        if o["release_qty"] == next_po_qty:
            return ("Transfer Order", False, o["order_num"])

    # Default: assume PO (conservative)
    return ("PO", False, "")


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

wb.close()

# ── 3. Identify common items with Next PO in NC ────────────────────
common_items = set(nc_items) & set(cic_items)

results = []
for item_code in sorted(common_items):
    nc  = nc_items[item_code]
    cic = cic_items[item_code]

    # Skip if NC has no Next PO
    if nc["next_po_date"] is None and nc["next_po_qty"] == 0:
        continue

    # ── Determine order type via Open Supply Orders ───────────
    order_type, verified, matched_order = _match_order(
        item_code, nc["next_po_date"], nc["next_po_qty"]
    )

    # ── Key calculation ───────────────────────────────────────
    # If the "Next PO" is a Transfer Order, do NOT subtract from On Order
    if order_type == "Transfer Order":
        nc_supply_excl_po = nc["on_hand"] + nc["on_order"]
    else:
        nc_supply_excl_po = nc["on_hand"] + (nc["on_order"] - nc["next_po_qty"])

    nc_shortfall      = max(0, nc["all_open_sales"] - nc_supply_excl_po)
    cic_available     = cic["on_hand"] - cic["all_open_sales"]  # surplus

    # Excel Formula Strings for the report
    # Output columns (1-indexed):
    #  A: Item Code                B: Item Description    C: Supplier
    #  D: NC On Hand               E: NC On Order         F: NC Next PO/TO Qty
    #  G: NC Next PO/TO Date       H: Order Type          I: Verified
    #  J: Matched Order #          K: NC Open Sales       L: NC Supply (excl PO)
    #  M: NC Shortfall             N: CIC On Hand         O: CIC On Order
    #  P: CIC Open Sales           Q: CIC Available       R: PO Status
    #  S: Reason

    if order_type == "Transfer Order":
        formula_nc_supply = '=D{row} + E{row}'
    else:
        formula_nc_supply = '=D{row} + (E{row} - F{row})'

    formula_nc_shortfall = '=MAX(0, K{row} - L{row})'
    formula_cic_available = '=N{row} - P{row}'

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
        "nc_next_po_qty":     nc["next_po_qty"],
        "nc_next_po_date":    nc["next_po_date"],
        # New columns
        "order_type":         order_type,
        "verified":           "Yes" if verified else "No",
        "matched_order":      matched_order,
        # NC continued
        "nc_open_sales":      nc["all_open_sales"],
        "nc_safety_stock":    nc["safety_stock"],
        "nc_supply_excl_po_val": nc_supply_excl_po,  # Keep val for python checks
        "nc_supply_excl_po":  formula_nc_supply,
        "nc_shortfall":       formula_nc_shortfall,
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
to_count      = sum(1 for r in results if r["order_type"] == "Transfer Order")
po_count      = sum(1 for r in results if r["order_type"] == "PO")

print("=" * 65)
print("  NC / CIC  Next-PO Cross-Analysis  (v3 – TO aware)")
print("=" * 65)
print(f"  Total NC items:                   {len(nc_items):>6}")
print(f"  Total CIC items:                  {len(cic_items):>6}")
print(f"  Items in both NC & CIC:           {len(common_items):>6}")
print(f"  … with Next PO/TO in NC:          {total:>6}")
print(f"    ├─ Purchase Orders:             {po_count:>6}")
print(f"    └─ Transfer Orders:             {to_count:>6}")
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
BLUE_FILL    = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
GREEN_FONT   = Font(name="Calibri", color="006100")
YELLOW_FONT  = Font(name="Calibri", color="9C6500")
RED_FONT     = Font(name="Calibri", color="9C0006")
BLUE_FONT    = Font(name="Calibri", color="1F4E79")
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

HEADERS = [
    "Item Code", "Item Description", "Supplier",
    "NC On Hand", "NC On Order", "NC Next PO/TO Qty", "NC Next PO/TO Date",
    "Order Type", "Verified", "Matched Order #",
    "NC Open Sales", "NC Supply (excl PO)", "NC Shortfall",
    "CIC On Hand", "CIC On Order", "CIC Open Sales", "CIC Available",
    "PO Status", "Reason",
]

COLUMN_WIDTHS = {
    1: 12, 2: 38, 3: 30,
    4: 13, 5: 13, 6: 17, 7: 18,
    8: 16, 9: 10, 10: 18,
    11: 14, 12: 19, 13: 14,
    14: 13, 15: 13, 16: 14, 17: 14,
    18: 22, 19: 55,
}

STATUS_STYLES = {
    "PO NOT Needed":       (GREEN_FILL,  GREEN_FONT),
    "PO Partially Needed": (YELLOW_FILL, YELLOW_FONT),
    "PO Needed":           (RED_FILL,    RED_FONT),
}

ORDER_TYPE_STYLES = {
    "Transfer Order": (BLUE_FILL, BLUE_FONT),
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
            r["nc_next_po_date"],
            r["order_type"], r["verified"], r["matched_order"],
            r["nc_open_sales"],
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

        # Colour status (col 18)
        fill, font = STATUS_STYLES.get(r["status"], (None, None))
        if fill:
            sheet.cell(row=row_idx, column=18).fill = fill
            sheet.cell(row=row_idx, column=18).font = font

        # Colour order type (col 8)
        ot_fill, ot_font = ORDER_TYPE_STYLES.get(r["order_type"], (None, None))
        if ot_fill:
            sheet.cell(row=row_idx, column=8).fill = ot_fill
            sheet.cell(row=row_idx, column=8).font = ot_font

# ─── All items in one detail sheet ──────────────────────────────────
ws1 = out_wb.active
write_sheet(ws1, "PO Review Detail", results)

# ─── Action sheet: items to cancel ──────────────────────────────────
ws2 = out_wb.create_sheet("Action – Cancel POs")
cancel_headers = ["Item Code", "Item Description", "NC Next PO/TO Qty",
                  "NC Next PO/TO Date", "Order Type", "Reason"]
for col_idx, h in enumerate(cancel_headers, 1):
    cell = ws2.cell(row=1, column=col_idx, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = THIN_BORDER
ws2.column_dimensions["A"].width = 12
ws2.column_dimensions["B"].width = 38
ws2.column_dimensions["C"].width = 17
ws2.column_dimensions["D"].width = 18
ws2.column_dimensions["E"].width = 16
ws2.column_dimensions["F"].width = 55
ws2.freeze_panes = "A2"
cancel_items = [r for r in results if r["status"] == "PO NOT Needed"]
for row_idx, r in enumerate(cancel_items, 2):
    for col_idx, val in enumerate([
        r["item_code"], r["item_desc"], r["nc_next_po_qty"],
        r["nc_next_po_date"], r["order_type"], r["reason"]
    ], 1):
        cell = ws2.cell(row=row_idx, column=col_idx, value=val)
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center")

# ─── Summary sheet ──────────────────────────────────────────────────
ws3 = out_wb.create_sheet("Summary")
summary_data = [
    ("NC vs CIC – Next PO/TO Review", ""),
    ("", ""),
    ("Metric", "Count"),
    ("Total NC items with PO/TO that also exist in CIC", total),
    ("  └─ Purchase Orders", po_count),
    ("  └─ Transfer Orders", to_count),
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
    print("-" * 140)
    for r in group:
        type_tag = " [TO]" if r["order_type"] == "Transfer Order" else ""
        print(f"  {r['item_code']:<10} {(r['item_desc'] or '')[:34]:<36} {r['reason']}{type_tag}")
