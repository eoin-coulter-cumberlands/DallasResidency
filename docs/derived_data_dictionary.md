# Derived Data Dictionary & Cleaning Decisions

This document records every field the data-engineering pipeline adds to the raw
UCI Online Retail dataset, and the justification for each cleaning decision. It
supplements the original `E-COMMERCE DATA DICTIONARY` for the 8 source columns
(`InvoiceNo`, `StockCode`, `Description`, `Quantity`, `InvoiceDate`,
`UnitPrice`, `CustomerID`, `Country`).

Pipeline: [`src/data_engineering.py`](../src/data_engineering.py)
Helpers:  [`src/analysis_helpers.py`](../src/analysis_helpers.py)

---

## 1. Processed datasets

| File | Grain | Used by | Description |
|------|-------|---------|-------------|
| `transactions_clean.parquet` | one line item | all | Master: every cleaned row + all flags below. Nothing dropped except duplicates and future-dated rows. |
| `sales.parquet` | one line item | Analyses 1–7, 9, 10 | Canonical revenue rows: `Quantity > 0` AND `UnitPrice > 0` AND not cancelled AND a real product. |
| `returns.parquet` | one line item | Analysis 8 | All returns/cancellations (`IsReturn == True`). |
| `products.parquet` | one product (StockCode) | Analyses 2, 7 | Per-product revenue, quantity, orders, avg price, customers. |
| `customers.parquet` | one customer (CustomerID) | Analyses 3, 9 | Per-customer RFM building blocks, CLV, cohort month. Named customers only. |

Parquet is the canonical format (compact, preserves dtypes and boolean flags).
A CSV export of the master is available with `python src/data_engineering.py --csv`
for the zip submission (~100 MB, git-ignored).

---

## 2. Derived line-item fields (master & sales)

| Field | Type | Source / formula | Purpose |
|-------|------|------------------|---------|
| `TotalPrice` | float | `Quantity × UnitPrice` | Line-item revenue. Negative for returns. |
| `Year` | int | from `InvoiceDate` | Time analysis |
| `Month` | int (1–12) | from `InvoiceDate` | Seasonality |
| `MonthName` | str | from `InvoiceDate` | Labels |
| `Day` | int | from `InvoiceDate` | Time analysis |
| `DayOfWeek` | str | from `InvoiceDate` | Day-of-week patterns (Analysis 5) |
| `DayOfWeekNum` | int (0=Mon) | from `InvoiceDate` | Sorting |
| `Hour` | int (0–23) | from `InvoiceDate` | Hour-of-day patterns (Analysis 5) |
| `Date` | datetime (midnight) | from `InvoiceDate` | Daily grouping |
| `YearMonth` | str `YYYY-MM` | from `InvoiceDate` | Monthly trend/cohort keys |
| `WeekOfYear` | int | ISO week | Weekly trend |
| `IsWeekend` | bool | `DayOfWeekNum >= 5` | Weekend vs weekday |
| `Season` | str | month → Winter/Spring/Summer/Autumn | Seasonal analysis |

### Flag columns (master only — let analysts filter as needed)

| Flag | Definition | Why |
|------|-----------|-----|
| `IsGuest` | `CustomerID` is null | 25% of rows lack an ID; flag instead of drop so they still count toward revenue/product totals. |
| `IsCancelled` | `InvoiceNo` starts with `C` | Identifies cancellations per data dictionary. |
| `IsReturn` | `IsCancelled` OR `Quantity < 0` | Routes rows to the returns dataset (Analysis 8). |
| `IsProduct` | StockCode is not a fee/adjustment code | Excludes postage, fees, etc. from product analysis. |
| `IsZeroOrNegPrice` | `UnitPrice <= 0` | Excluded from `sales`; flagged for inspection. |
| `IsQtyOutlier` | `Quantity` outside [−1000, 10000] | Data-dictionary bounds; flagged not dropped. |
| `IsPriceOutlier` | `UnitPrice` outside [0.01, 1000] | Data-dictionary bounds; flagged not dropped. |
| `IsMissingDescription` | description null or `"?"` | Data-quality signal. |

---

## 3. Cleaning decisions (and justifications)

1. **Duplicate rows — dropped (5,268).** Exact duplicates across all 8 columns
   are almost certainly double-logged lines and would inflate revenue/quantity.

2. **Missing `CustomerID` (≈25%) — kept and flagged `IsGuest`.** Dropping would
   lose a quarter of revenue and distort product/geographic/time analyses.
   Customer-level helpers (RFM, CLV, cohort) exclude guests automatically, since
   those metrics are meaningless without an identifier. *(Dictionary Option 2.)*

3. **Returns / cancellations — split into a separate dataset.** Kept in the
   master (flagged) and surfaced as `returns.parquet`. The `sales` dataset
   excludes them so revenue is not understated/distorted. *(Dictionary Option 2.)*

4. **Zero / negative prices (2,512) — excluded from `sales`, flagged in master.**
   These are gifts, samples, or bad-debt adjustments (one price is negative GBP 11,062) and
   would skew pricing analysis.

5. **Non-product line items — flagged `IsProduct = False`.** Codes such as
   `POST`, `DOT`, `C2`, `M`, `D`, `S`, `B`, `BANK CHARGES`, `AMAZONFEE`, `CRUK`,
   `PADS`, and `gift_*` are fees/adjustments, not sellable goods, so they are
   removed from product/pricing analysis but retained in the master.

6. **Description inconsistencies — StockCode is the product key.** One canonical
   description per StockCode (the most frequent value) avoids splitting a product
   across spelling/capitalisation variants.

7. **Outliers — flagged, not dropped.** Extreme quantities/prices may be genuine
   wholesale orders. Flags let each analysis decide; nothing is silently removed.

8. **Future-dated rows — clipped to the real data range** (ends 2011-12-09) as a
   defensive step.

---

## 4. Customer table fields (RFM / CLV / cohort)

| Field | Meaning |
|-------|---------|
| `FirstPurchase` / `LastPurchase` | first/last order datetime |
| `NumOrders` (`Frequency`) | distinct invoices |
| `TotalSpend` (`Monetary`) | lifetime revenue |
| `TotalItems` | lifetime units |
| `Recency` | days from last purchase to snapshot (day after final order) |
| `AvgOrderValue` | `TotalSpend / NumOrders` |
| `CohortMonth` | `YYYY-MM` of first purchase |
| `IsRepeat` | more than one order |

> **Note (Analysis 5):** the dataset contains **no Saturday transactions** — a
> genuine characteristic of this source, not a pipeline artifact.
