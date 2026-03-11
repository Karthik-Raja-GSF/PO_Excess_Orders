# NC & CIC Purchase Order / Transfer Order Analysis Flow

## Objective

The goal of this analysis is to identify items that exist in both the `NC` and `CIC` Site Codes and determine whether a scheduled "Next PO" (or Transfer Order) for the `NC` site is truly necessary, partially necessary, or unnecessary. We determine this by checking if the `CIC` site has enough surplus inventory to transfer over and cover `NC`'s shortfall.

**Important:** The "Next PO" field in the Forecast Analysis may actually represent a **Transfer Order** (TO) from CIC → NC, not a true Purchase Order. The script cross-references against `Open Supply Orders.xlsx` to identify these.

## Data Sources

- **Input Files:**
  - `Forecast Analysis.xlsx` – Main forecast data with on-hand, on-order, open sales, and next PO info
  - `Open Supply Orders.xlsx` – All open supply orders; used to identify Transfer Orders (`NC-TO*` in Order #) and validate PO qty/date
- **Key Fields Extracted:**
  - Site Code (`NC` or `CIC`)
  - Item Code & Description
  - On Hand Quantity, On Order Quantity, All Open Sales
  - Next PO Date & Next PO Quantity
  - Order # (from Open Supply Orders, to distinguish PO vs TO)

## Logical Flow

### 1. Load & Cross-Reference Open Supply Orders

The script loads `Open Supply Orders.xlsx` and groups NC orders by item:

- **Transfer Orders** – Order # starts with `NC-TO`
- **Purchase Orders** – All other orders

For each NC item with a "Next PO" in the Forecast Analysis, the script matches the PO Date and Qty against the Open Supply Orders to determine whether it's a PO or Transfer Order.

### 2. Identify Target Items

The script scans all rows in the Forecast Analysis and isolates items that meet two conditions:

1. The item is stocked in **both** `NC` and `CIC`.
2. The item has a **Next PO Date** and/or **Next PO Quantity** scheduled for the `NC` site.

### 3. Calculate NC Shortfall

**For Purchase Orders:**
`NC Supply (excluding PO) = NC On Hand + (NC On Order - NC Next PO Qty)`

Because `NC On Order` already includes the `Next PO Qty`, we subtract it to see supply _without_ that PO.

**For Transfer Orders:**
`NC Supply (excluding PO) = NC On Hand + NC On Order`

Transfer Orders are **not** subtracted from On Order, because they represent incoming transfers from CIC, not external POs to be potentially cancelled.

- **NC Shortfall** = `max(0, NC Open Sales - NC Supply (excluding PO))`

### 4. Calculate CIC Available Surplus

- **CIC Available** = `CIC On Hand` - `CIC Open Sales`

### 5. Determine PO Status

1. 🟢 **PO NOT Needed** – `NC Shortfall` is 0 **OR** `CIC Available` ≥ `NC Shortfall`
2. 🟡 **PO Partially Needed** – `CIC Available` > 0 but less than `NC Shortfall`
3. 🔴 **PO Needed** – `CIC Available` ≤ 0

## Output

The script generates `NC_CIC_PO_Analysis.xlsx` containing:

1. **PO Review Detail**: Full breakdown including Order Type (PO/Transfer Order), Verified status, and Matched Order #.
2. **Action – Cancel POs**: Orders that should be cancelled entirely.
3. **Summary**: High-level counts including PO vs Transfer Order breakdown.
