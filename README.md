# NC & CIC Purchase Order (PO) Analysis Flow

## Objective

The goal of this analysis is to identify items that exist in both the `NC` and `CIC` Site Codes and determine whether a scheduled "Next PO" for the `NC` site is truly necessary, partially necessary, or unnecessary. We determine this by checking if the `CIC` site has enough surplus inventory to transfer over and cover `NC`'s shortfall..

## Data Source

- **Input File:** `Forecast_Analysis.xlsx`, `Open Supply Orders.xlsx`
- **Key Fields Extracted:**
  - Site Code (`NC` or `CIC`)
  - Item Code & Description
  - On Hand Quantity
  - On Order Quantity
  - All Open Sales
  - Next PO Date
  - Next PO Quantity

## Logical Flow

### 1. Identify Target Items

The script scans all rows and isolates items that meet two conditions:

1. The item is stocked in **both** `NC` and `CIC`.
2. The item has a **Next PO Date** and/or **Next PO Quantity** scheduled for the `NC` site.

### 2. Calculate NC Shortfall

To figure out if `NC` actually needs the scheduled PO, we calculate what their supply looks like _without_ that specific PO.

**The Formula: `NC Supply (excluding PO) = NC On Hand + (NC On Order - NC Next PO Qty)`**

Here is what each part means:

- **NC On Hand**: The physical inventory currently sitting on the shelves in the NC warehouse.
- **NC On Order**: The _total_ amount of product ordered from suppliers that hasn't arrived yet. **This number includes the Next PO.**
- **NC Next PO Qty**: The specific amount scheduled to come in on the _very next_ PO (the one we are deciding whether to cancel or keep).

Because `NC On Order` already includes the `Next PO Qty`, we have to subtract it back out to see what the future supply looks like _if we cancel it_.

_Example: If we have 100 on Hand, 150 total on Order, and the Next PO is for 50. If we cancel the Next PO, our true supply is 100 + (150 - 50) = 200._

- **NC Shortfall** = `NC Open Sales` - `NC Supply (excluding PO)`
  - _If this number is ≤ 0, NC already has enough supply and the PO is unnecessary regardless of CIC._

### 3. Calculate CIC Available Surplus

We calculate how many extra items `CIC` has that could potentially be transferred to `NC`.

- **CIC Available** = `CIC On Hand` - `CIC Open Sales`

### 4. Determine PO Status

We compare the `NC Shortfall` against the `CIC Available` surplus to assign one of three statuses:

1. 🟢 **PO NOT Needed**
   - _Condition:_ `NC Shortfall` is 0 (NC already has enough) **OR** `CIC Available` ≥ `NC Shortfall`
   - _Action:_ The scheduled NC PO can be fully cancelled.

2. 🟡 **PO Partially Needed**
   - _Condition:_ `CIC Available` > 0, but it is less than the `NC Shortfall`.
   - _Action:_ `CIC` can supply _some_ of the need. The NC PO quantity should be reduced by the `CIC Available` amount.

3. 🔴 **PO Needed**
   - _Condition:_ `CIC Available` ≤ 0
   - _Action:_ `CIC` has no surplus to give. The NC PO must remain as scheduled.

## Output

The script generates `NC_CIC_PO_Analysis.xlsx` containing:

1. **PO Review Detail**: A comprehensive breakdown of all cross-referenced items and their calculated metrics.
2. **Action – Cancel POs**: A streamlined list of only the orders that should be cancelled entirely.
3. **Summary**: High-level counts of items falling into the Not Needed, Partially Needed, and Needed categories.
