# Cafe Sales Data Cleaning: From Untrustworthy Export to Audited Dataset

> DecodeLabs Data Analytics, Project 1 (Data Cleaning & Preparation) | Tool: Microsoft Excel

## The scenario

A small cafe owner wants to answer two simple questions: **which items sell best**, and **do customers buy more in-store or as takeaway?**

The owner exports a year of transactions from the till system and finds the file can't be trusted. Cells contain `ERROR` and `UNKNOWN` instead of values, payment method and location are blank on thousands of rows, and some prices and totals are missing. Any report built on this export would be quietly wrong.

**My job:** turn the raw export into a dataset the owner can rely on, and prove it is reliable.

## The data

- Source: *Cafe Sales - Dirty Data for Cleaning Training* on Kaggle (a practice dataset for a fictional cafe, so the business is made up but the data problems are realistic).
- 10,000 transactions, 8 columns: Transaction ID, Item, Quantity, Price Per Unit, Total Spent, Payment Method, Location, Transaction Date.
- 10,082 missing cells in the raw file (placeholders such as `ERROR`/`UNKNOWN`, plus true blanks).

| Column | Missing in raw data |
|---|---|
| Location | 39.6% |
| Payment Method | 31.8% |
| Item | 9.7% |
| Price Per Unit | 5.3% |
| Total Spent | 5.0% |
| Quantity | 4.8% |
| Transaction Date | 4.6% |
| Transaction ID | 0% |

## Approach

I built a formula-driven Excel workbook so every step is visible and repeatable, then checked it with an audit sheet.

1. **Identify:** profile every column for `ERROR`, `UNKNOWN` and blank cells.
2. **Handle missing values (strategic imputation, no deleting):**
   - Derive values from the same row first: Total / Quantity, Total / Price, Quantity x Price.
   - Use the item's reference price when needed (every item has exactly one price in this data, which I verified).
   - Use the median quantity only as a last resort.
   - Infer a missing item from its price only where that price belongs to exactly one item.
3. **Remove duplicates:** check the Transaction ID column and keep the first occurrence.
4. **Standardise formats:** ISO 8601 dates (`YYYY-MM-DD`), trimmed Proper Case text, prices and totals to 2 decimals, whole-number quantities.
5. **Audit and document:** a live audit sheet and a numbered change log.

## Results

| Check | Result |
|---|---|
| Rows in / rows out | 10,000 / 10,000 (nothing deleted) |
| Duplicate Transaction IDs | **0** |
| Incorrectly formatted dates | **0** |
| Rows where Total does not equal Quantity x Price | 0 |
| Missing items recovered from price | 489 |
| Missing prices filled | 527 |
| Missing quantities filled | 479 |
| Missing totals derived | 499 |

## Decisions and trade-offs

- **"Unknown" instead of the most common value** for Payment Method and Location. With 32% and 40% missing, filling with the mode would have made one category look far bigger than the data supports.
- **Dates were not guessed.** 460 dates were missing in the raw file and are left blank, flagged for review. They are excluded from time-based analysis only.
- **6 rows could not be fully repaired** because too many of their fields were missing. They are kept, flagged, and listed by ID in the change log.
- **Quantity and Price are trusted over Total.** In this dataset there were no conflicts, but the workbook would recompute Total if they disagreed.

## Verifying the checker, not just the data

A PASS is only meaningful if the check can fail. I tested the audit by deliberately breaking it (for example, pasting formulas instead of values into the final sheet) and confirming it reported FAIL. I also fixed one check that was silently passing while its underlying formula returned an error.

## Files

| File | What it is |
|---|---|
| `cafe_sales_cleaning_workbook.xlsx` | The full workbook: raw data, cleaning formulas, final dataset, audit, change log |
| `cafe_sales_clean.csv` | The cleaned dataset |
| `cafe_sales_change_log.pdf` | Change log (17 numbered changes) with the verification results |

The raw dataset is not re-uploaded here. Download it from Kaggle (link below) and paste it into the `Raw_Data` sheet to reproduce the results.

## How to reproduce

1. Download the raw file from Kaggle: [add link].
2. Open the workbook and paste the data into `Raw_Data` starting at A1.
3. Review `Lookups`, then check `Clean_Data` and `Audit`.
4. Copy the kept rows into `Final_Clean` as **values only**; the `Audit` sheet should read PASS.
5. Export `Change_Log` to PDF.

## Skills demonstrated

Data cleaning and preparation, missing-value imputation strategy, duplicate detection, data-format standardisation, Excel formulas (`SUMPRODUCT`, `COUNTIFS`, `INDEX/MATCH`, `TEXT/DATEVALUE`), audit design, and documentation.

## Limitations

This is a practice dataset, so it shows the method rather than a real business outcome. The next step would be using the cleaned data to answer the owner's two questions (best sellers, in-store vs takeaway).
