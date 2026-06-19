# Part 1 — Business Data Cleaning, Validation & Excel Reporting

## Problem Summary
A retail company exports order-level sales data from multiple internal systems. The raw dataset (`raw_orders.xlsx`) contains 932 records with issues including inconsistent date formats, duplicate records, missing values, invalid discounts, sales/profit mismatches, and order status inconsistencies. This project cleans and validates the data, applies business rules, creates calculated columns, and generates summary pivot reports.

---

## Dataset Description
| Property | Value |
|----------|-------|
| File | raw_orders.xlsx |
| Sheet | raw_orders |
| Raw Rows | 932 |
| Columns | 21 |
| Key Fields | order_id, order_date, ship_date, customer_name, segment, region, category, sub_category, ship_mode, quantity, unit_price, discount, sales, cost, profit, payment_status, order_status |

---

## Tools Used
- Microsoft Excel
- Excel Formulas
- Data Validation
- Pivot Tables

---

## Cleaning Steps Performed
1. **Text Standardization** — stripped whitespace, normalized capitalization , collapsed multiple spaces across all text columns.
2. **Date Parsing** — handled 8+ mixed date formats; standardized to a consistent Excel date format.
3. **Duplicate Detection** — identified 20 exact duplicate rows (removed) and 12 conflicting duplicate order_ids (flagged).
4. **Missing Values** — filled missing `region` and `ship_mode` with "Unknown"; treated missing `discount` as 0.
5. **Discount Validation** — flagged 15 negative discounts as invalid; set to 0 for calculation purposes.
6. **Calculated Columns** — added `cleaned_discount`, `calculated_sales`, `calculated_profit`, `profit_margin`, `shipping_delay_days`, `order_month`, `order_year`, `data_quality_flag`.
7. **Business Rule Application** — excluded Cancelled/Failed/Refunded records from completed-sales summaries; flagged invalid shipping date sequences.

---

## Business Rules Applied
| Rule | Action |
|------|--------|
| Missing region | Filled as "Unknown", flagged |
| Missing ship_mode | Filled as "Unknown", flagged |
| Missing discount | Treated as 0 if other sales fields valid |
| Negative discount | Flagged invalid; set to 0 for calculations |
| Cancelled orders | Excluded from completed-sales summaries |
| Failed payments | Excluded from completed-sales summaries |
| Refunded orders | Separately summarized in pivot report |
| Ship date before order date | Flagged as invalid record |
| Conflicting duplicate order_ids | Retained; flagged for review |

---

## Summary of Data Quality Issues Found
| Issue | Count |
|-------|-------|
| Exact duplicate rows removed | 20 |
| Conflicting duplicate order_ids flagged | 12 |
| Ship date before order date | 21 |
| Missing discount (treated as 0) | 26 |
| Negative discount (flagged invalid) | 15 |
| Missing ship_mode (filled Unknown) | 21 |
| Missing region (filled Unknown) | 25 |
| Sales calculation mismatch | 56 |
| Profit calculation mismatch | 56 |
| Clean records = 703
| Invalid records = 209 

---

## Summary of Final Pivot Reports (`pivot_summary.xlsx`)
| Sheet | Description |
|-------|-------------|
| Sales by Region | Total sales & profit by region, sorted by sales (completed+paid only) |
| Sales by Category | Breakdown by category & sub-category, sorted |
| Orders by Ship Mode | Order count and total sales by shipping method |
| Margin by Segment | Profit margin % by customer segment |
| Issues by Region | Count of cancelled/returned/failed/refunded orders by region |
| Monthly Sales Trend | Month-by-month sales trend for completed orders |

---

## Key Business Insights
1. **North and South regions** contribute the highest sales volumes.
2. **Technology** category generates the highest sales; **Furniture** has the widest sub-category spread.
3. **Standard Class** is the most popular shipping method by volume.
4. **Corporate segment** tends to have higher average order values.
5. **56 orders** have reported sales that don't match the calculated value — suggesting a source system calculation issue that needs investigation.
6. **21 orders** were shipped before their recorded order date — likely data entry errors in the source systems.
7. **Refunded orders** are concentrated in a few regions — may indicate product quality or logistics issues in those areas.
8. **Cancelled + Failed payment orders = 214 records** — these must be excluded from revenue reporting to avoid inflating sales figures.

---

## Assumptions and Limitations
- Ambiguous dates were interpreted using a consistent DD-MM-YYYY approach.
- Conflicting duplicate order_ids (12 records) were flagged but not deleted — business decision required.
- Calculated profit = calculated_sales − cost (not using reported `profit` column); mismatches flagged.
- No external reference data available to validate unit prices.

---

## Screenshots Included
| File | Description |
|------|-------------|
| `screenshots/raw_data_preview.png` | Raw dataset before cleaning |
| `screenshots/cleaned_data_preview.png` | Cleaned dataset with all calculated columns |
| `screenshots/pivot_summary_1.png` | Sales & Profit by Region pivot |
| `screenshots/pivot_summary_2.png` | Monthly Sales Trend pivot |
