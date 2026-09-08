# Retail Sales Data Cleaning & Analysis (Excel)

Cleaning and validating a messy retail transactions dataset using formula-driven
Excel logic (no manual overwrites), then summarizing category-level revenue insights.

## Business Problem

Retail POS exports are rarely clean: missing item names, missing prices, missing
quantities, totals that don't match `Price × Quantity`, inconsistent category casing,
stray whitespace, and duplicate transaction rows. Before this data can be trusted for
reporting or a dashboard, it needs a **repeatable, auditable cleaning process** — not
one-off manual edits that can't be re-run when new data arrives.

## Dataset

- **615 rows** of retail transactions across 8 categories (Beverages, Food, Milk
  Products, Patisserie, Butchers, Electric Accessories, Computers, Furniture)
- Columns: Transaction ID, Date, Category, Item, Price Per Unit, Quantity, Total
  Spent, Payment Method
- Deliberately messy: missing values across Item / Price / Quantity / Total Spent,
  mismatched totals, inconsistent text casing, extra whitespace, and duplicate rows

## Approach

All cleaning logic lives in **live Excel formulas**, so the workbook re-validates
itself automatically if the raw data changes — nothing is hardcoded.

1. **Reference Table** — built from the rows that already have both a valid Item and
   Price, giving an Item → Category → Price lookup table.
2. **Missing Item** → reverse-looked-up from the raw Price against the Reference
   Table (`INDEX`/`MATCH`).
3. **Missing Price** → looked up from the Item against the Reference Table.
4. **Missing Quantity / Total Spent** → derived from whichever two of
   `Price, Quantity, Total` are known (`Total = Price × Quantity`).
5. **Total validation** — every row is flagged `OK`, `Mismatch`, or `Incomplete` by
   comparing `Price × Quantity` against the reported Total Spent.
6. **Text standardization** — category casing normalized, item names trimmed of
   stray whitespace.
7. **Item Fix Note** — every row is labeled with exactly what happened to it (OK /
   Filled from Price lookup / Trimmed whitespace).

## Results

| Metric | Before | After |
|---|---|---|
| Missing Item | 68 | Recovered via Price lookup |
| Missing Price Per Unit | 47 | 15 remaining |
| Missing Quantity | 41 | 15 remaining |
| Missing Total Spent | 38 | 15 remaining |
| Total Spent mismatches flagged | not checked | Caught and labeled |

## Key Insight

**Furniture** is the top revenue-generating category, followed by **Computers**,
while **Patisserie** contributes the least — useful for a business deciding where to
focus inventory or promotions.

## Files in this repo

| File | Description |
|---|---|
| `Retail_Sales_Cleaning_Project.xlsx` | Full workbook: Raw Data, Reference Table, Cleaned Data |
| `dirty_retail_sales.csv` | Raw, uncleaned source data |
| `cleaned_retail_sales.csv` | Cleaned output as a static CSV |

## Tools Used

Excel (formulas: `INDEX`, `MATCH`, `IFERROR`, `PROPER`, `TRIM`, conditional
formatting)

## What I'd Do Next

- Rebuild the Cleaned Data sheet as a Power Query pipeline for one-click refresh on
  new data drops
- Extend the analysis into a full Power BI/Tableau dashboard (see my next project)
