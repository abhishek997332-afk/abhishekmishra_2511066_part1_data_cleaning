# Cleaning Log — Part 1: Business Data Cleaning & Validation

## Dataset Overview
- **File:** raw_orders.xlsx (sheet: raw_orders)
- **Raw Records:** 932 rows, 21 columns
- **Cleaned Records:** 912 rows (after removing 20 exact duplicate rows)
- **Tool Used:** Microsoft Excel (Data Cleaning, Validation, Formulas, and Pivot Tables)

---

## Issues Found

### 1. Text / Formatting Issues
| Field | Issues Observed |
|-------|----------------|
| `segment` | Trailing spaces (e.g., "Corporate "), mixed case |
| `ship_mode` | ALL CAPS entries (e.g., "STANDARD CLASS", "FIRST CLASS") |
| `region` | Leading/trailing spaces (e.g., "  North ") |
| `customer_name` | Trailing spaces (e.g., "Ananya Rao ") |
| `payment_status` | Trailing spaces (e.g., "Paid ") |
| `order_status` | Trailing spaces (e.g., "Completed ") |

### 2. Date Issues
| Issue | Count |
|-------|-------|
| Ship date before order date | 21 |
| Mixed date formats (DD MMM YYYY, MM/DD/YYYY, YYYY-MM-DD, DD-MM-YYYY) | All rows |

### 3. Duplicate Records
| Type | Count |
|------|-------|
| Exact duplicate rows | 20 |
| Duplicate order_ids (with conflicting data) | 12 |

### 4. Missing Values
| Field | Count | Handling |
|-------|-------|---------|
| `discount` | 26 | Treated as 0 (other fields valid) |
| `ship_mode` | 21 | Filled as "Unknown" |
| `region` | 25 | Filled as "Unknown" |

### 5. Invalid Discounts
| Issue | Count |
|-------|-------|
| Negative discount | 15 |
| Discount > 90% | 0 |

### 6. Sales / Profit Mismatches
| Metric | Count |
|--------|-------|
| Reported sales ≠ calculated_sales (>1 INR diff) | 56 |
| Reported profit ≠ calculated_profit (>1 INR diff) | 56 |

---

## Cleaning Actions Performed

### Text Cleaning
Removed extra spaces using Excel functions.
Standardized text formatting.
Removed duplicate records using Excel's Remove Duplicates feature.

### Date Cleaning
- Standardized date formats using Excel date conversion and formatting functions.
- Converted mixed date formats into a consistent date format for analysis.

### Duplicate Handling
- Exact duplicate rows: Removed 20 duplicate records using Excel's Remove Duplicates feature while retaining the first occurrence.

### Missing Values
-  ship_mode: 21 missing values filled as "Unknown".
- region: 25 missing values filled as "Unknown".
- discount: 26 missing values treated as 0 for calculation purposes.

### Discount Validation
- 15 negative discount rows → flagged as invalid in `discount_flag`; `cleaned_discount` set to 0 for calculation purposes.
- No discounts exceeded 90%.

---

## Business Rules Applied

| Rule Area | Action Taken |
|-----------|-------------|
| Missing region | Filled as "Unknown"; flagged in `region_flag` column |
| Missing ship_mode | Filled as "Unknown"; flagged in `ship_mode_flag` column |
| Missing discount | Treated as 0 since other fields were valid |
| Negative discount | Flagged invalid; cleaned_discount set to 0 |
| Discount > 90% | None found in this dataset |
| Cancelled orders | Excluded from completed-sales pivot summaries |
| Failed payments | Excluded from completed-sales pivot summaries |
| Refunded orders | Included in a separate refunded summary pivot |
| Ship date before order date | Flagged as invalid shipping record; not used for delay calculation |
| Conflicting duplicate order_ids | Retained all records; flagged for human review |

---

## Calculated Columns Added

| Column | Formula / Logic |
|--------|----------------|
| `cleaned_discount` | Missing → 0; Negative → 0; otherwise same as discount |
| `calculated_sales` | `quantity × unit_price × (1 − cleaned_discount)` |
| `calculated_profit` | `calculated_sales − cost` |
| `profit_margin` | `calculated_profit / calculated_sales` |
| `shipping_delay_days` | `ship_date − order_date` (days); blank if either date invalid or ship < order |
| `order_month` | Month number extracted from order_date |
| `order_year` | Year extracted from order_date |
| `data_quality_flag` | "Clean", "Invalid" (see below) |

### data_quality_flag Logic
- **Clean:** No issues at all
- **Invalid:** Has critical issues (ship date before order date, missing/invalid order date, negative discount)

---

## Records Removed
- 20 exact duplicate rows removed (kept first occurrence).

## Records Flagged
- 209 records flagged "Invalid"
- 703 records flagged "Clean"

---

## Assumptions Made
1. Mixed date formats were interpreted using a consistent DD-MM-YYYY approach where dates were ambiguous.
2. Negative discounts were zeroed out for calculations but flagged; the original discount value is preserved in the `discount` column.
3. "Completed" order_status AND "Paid" payment_status both required to include a record in completed-sales summaries.
4. Profit was recalculated as `calculated_sales − cost`; any deviation from the reported `profit` column is flagged as a mismatch.
5. Records with conflicting duplicate order_ids were retained (not deleted) as the correct version cannot be determined without business context.

---

## Limitations
- Date ambiguity: Some dates (e.g., 06/07/2024) could be June 7 or July 6. 
- Conflicting duplicate order_ids (12 records) were flagged but not resolved — a business decision is needed.
- `calculated_sales` may differ from reported `sales` due to rounding in source system; threshold of 1 INR used for mismatch flagging.
- No external price list available to independently verify `unit_price` correctness.
