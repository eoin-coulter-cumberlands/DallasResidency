"""
E-Commerce Business Analytics
Data Engineering Pipeline
=========================

Owner: Data Engineer

This module is the single source of truth for loading, cleaning, and feature
engineering of the UCI Online Retail dataset (Dec 2010 - Dec 2011). It produces
clean, analysis-ready datasets that every one of the 10 required analyses can
load directly, so the analysts never have to re-clean the raw file.

Run as a script to (re)build everything in data/processed/:

    python src/data_engineering.py

Or import the pieces you need:

    from src.data_engineering import load_clean, load_sales, load_returns
    sales = load_sales()          # canonical revenue-generating transactions

Cleaning decisions (justified in docs/derived_data_dictionary.md):
  1. Drop exact duplicate rows.
  2. Parse InvoiceDate; clip to the real data range (drops future-dated rows).
  3. Derive TotalPrice and a full set of time fields.
  4. Flag returns/cancellations (IsReturn, IsCancelled) and keep them; ship a
     separate returns dataset for Analysis 8.
  5. Flag missing CustomerID as IsGuest; keep the rows for revenue/product
     analysis but exclude them from customer/cohort helpers.
  6. Flag non-product line items (POST, M, D, BANK CHARGES, ...) as IsProduct=False.
  7. Use StockCode as the product key; derive one canonical description per code.
  8. Flag outliers using the data-dictionary bounds; do NOT drop them from the
     master so analysts can decide.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
ROOT = Path(__file__).resolve().parents[1]
RAW_XLSX = ROOT / "data" / "raw" / "Online Retail.xlsx"
PROCESSED = ROOT / "data" / "processed"

# Canonical processed-file names (analysts load these by name).
F_MASTER = PROCESSED / "transactions_clean.parquet"
F_SALES = PROCESSED / "sales.parquet"
F_RETURNS = PROCESSED / "returns.parquet"
F_PRODUCTS = PROCESSED / "products.parquet"
F_CUSTOMERS = PROCESSED / "customers.parquet"

# --------------------------------------------------------------------------- #
# Reference values
# --------------------------------------------------------------------------- #
# Non-product line items. These StockCodes (or code prefixes) are fees,
# adjustments, postage, samples, gift vouchers, etc. - not sellable products.
# Excluded from product/pricing analysis but kept in the master for completeness.
NON_PRODUCT_CODES = {
    "POST",          # postage
    "DOT",           # dotcom postage
    "C2",            # carriage
    "M", "m",        # manual adjustment
    "D",             # discount
    "S",             # samples
    "B",             # adjust bad debt
    "BANK CHARGES",  # bank fees
    "AMAZONFEE",     # amazon fee
    "CRUK",          # charity donation
    "PADS",          # pads to match all inventory
}
NON_PRODUCT_PREFIXES = ("gift_",)  # gift voucher top-ups (gift_0001_10, ...)

# Outlier bounds from the data dictionary (Analysis Recommendations section).
QTY_MIN, QTY_MAX = -1000, 10000
PRICE_MIN, PRICE_MAX = 0.01, 1000.0

# UK English day order for consistent sorting across analyses.
DOW_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday",
             "Friday", "Saturday", "Sunday"]


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _season(month: int) -> str:
    """Map a month number to a meteorological season (Northern Hemisphere)."""
    if month in (12, 1, 2):
        return "Winter"
    if month in (3, 4, 5):
        return "Spring"
    if month in (6, 7, 8):
        return "Summer"
    return "Autumn"


def _is_non_product(code: str) -> bool:
    code = str(code)
    if code in NON_PRODUCT_CODES:
        return True
    return any(code.startswith(p) for p in NON_PRODUCT_PREFIXES)


# --------------------------------------------------------------------------- #
# Load + clean
# --------------------------------------------------------------------------- #
def load_raw() -> pd.DataFrame:
    """Read the raw Excel file exactly as delivered."""
    if not RAW_XLSX.exists():
        raise FileNotFoundError(
            f"Raw dataset not found at {RAW_XLSX}. "
            "Download 'Online Retail.xlsx' from the UCI repository and place it "
            "in data/raw/."
        )
    return pd.read_excel(RAW_XLSX, dtype={"InvoiceNo": str, "StockCode": str})


def clean(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the raw frame and add all derived/flag columns.

    Returns the *master* dataset: every row is kept, but each row is now
    labelled with flags so any analysis can filter to exactly what it needs.
    """
    df = df.copy()

    # 1. Drop exact duplicate rows (5,268 in the raw file).
    n_before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    n_dupes = n_before - len(df)

    # 2. Normalise text columns.
    df["InvoiceNo"] = df["InvoiceNo"].astype(str).str.strip()
    df["StockCode"] = df["StockCode"].astype(str).str.strip()
    df["Description"] = df["Description"].astype("string").str.strip()
    df["Country"] = df["Country"].astype(str).str.strip()

    # 3. Parse dates and clip to the real data range (defends against any
    #    future-dated rows noted in the dictionary).
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])
    last_valid = df["InvoiceDate"].max().normalize() + pd.Timedelta(days=1)
    df = df[df["InvoiceDate"] < last_valid].reset_index(drop=True)

    # 4. Derived monetary field.
    df["TotalPrice"] = df["Quantity"] * df["UnitPrice"]

    # 5. Time features for the time-based / seasonal analyses.
    dt = df["InvoiceDate"].dt
    df["Year"] = dt.year
    df["Month"] = dt.month
    df["MonthName"] = dt.month_name()
    df["Day"] = dt.day
    df["DayOfWeek"] = dt.day_name()
    df["DayOfWeekNum"] = dt.dayofweek            # 0 = Monday
    df["Hour"] = dt.hour
    df["Date"] = dt.normalize()
    df["YearMonth"] = dt.to_period("M").astype(str)   # e.g. "2011-03"
    df["WeekOfYear"] = dt.isocalendar().week.astype(int)
    df["IsWeekend"] = df["DayOfWeekNum"] >= 5
    df["Season"] = df["Month"].map(_season)

    # 6. Customer flags. Keep CustomerID as a nullable integer.
    df["CustomerID"] = df["CustomerID"].astype("Int64")
    df["IsGuest"] = df["CustomerID"].isna()

    # 7. Return / cancellation flags.
    df["IsCancelled"] = df["InvoiceNo"].str.upper().str.startswith("C")
    df["IsReturn"] = df["IsCancelled"] | (df["Quantity"] < 0)

    # 8. Product vs non-product line items.
    df["IsProduct"] = ~df["StockCode"].map(_is_non_product)

    # 9. Data-quality flags (price/quantity sanity).
    df["IsZeroOrNegPrice"] = df["UnitPrice"] <= 0
    df["IsQtyOutlier"] = ~df["Quantity"].between(QTY_MIN, QTY_MAX)
    df["IsPriceOutlier"] = ~df["UnitPrice"].between(PRICE_MIN, PRICE_MAX)
    df["IsMissingDescription"] = df["Description"].isna() | (df["Description"] == "?")

    df.attrs["n_duplicates_dropped"] = n_dupes
    return df


# --------------------------------------------------------------------------- #
# Dataset builders
# --------------------------------------------------------------------------- #
def build_canonical_descriptions(master: pd.DataFrame) -> pd.Series:
    """One canonical description per StockCode (the most frequent non-null one)."""
    named = master[master["Description"].notna() & (master["Description"] != "?")]
    canon = (
        named.groupby("StockCode")["Description"]
        .agg(lambda s: s.mode().iat[0] if not s.mode().empty else s.iloc[0])
    )
    return canon


def build_sales(master: pd.DataFrame) -> pd.DataFrame:
    """Canonical revenue-generating transactions.

    Definition (data dictionary): Quantity > 0 AND UnitPrice > 0 AND not a
    cancellation. This is what Analyses 1-7, 9, 10 should treat as "a sale".
    Non-product line items (postage, fees) are excluded so product/revenue
    figures reflect actual goods.
    """
    sales = master[
        (master["Quantity"] > 0)
        & (master["UnitPrice"] > 0)
        & (~master["IsCancelled"])
        & (master["IsProduct"])
    ].copy()
    return sales


def build_returns(master: pd.DataFrame) -> pd.DataFrame:
    """All returns / cancellations, for Analysis 8."""
    return master[master["IsReturn"]].copy()


def build_products(sales: pd.DataFrame, canon: pd.Series) -> pd.DataFrame:
    """Per-product summary keyed on StockCode (Analyses 2 & 7)."""
    g = sales.groupby("StockCode")
    products = pd.DataFrame({
        "TotalQuantity": g["Quantity"].sum(),
        "TotalRevenue": g["TotalPrice"].sum(),
        "NumOrders": g["InvoiceNo"].nunique(),
        "AvgUnitPrice": g["UnitPrice"].mean(),
        "NumCustomers": g["CustomerID"].nunique(),
    })
    # Map canonical description onto the products that actually sold.
    products.insert(0, "Description", products.index.map(canon))
    products = products.sort_values("TotalRevenue", ascending=False)
    products["RevenueRank"] = products["TotalRevenue"].rank(ascending=False).astype(int)
    return products.reset_index()


def build_customers(sales: pd.DataFrame) -> pd.DataFrame:
    """Per-customer summary for CLV / segmentation / cohort (Analyses 3 & 9).

    Guests (missing CustomerID) are excluded - customer-level metrics are
    meaningless without an identifier.
    """
    named = sales[~sales["IsGuest"]].copy()
    snapshot = named["InvoiceDate"].max() + pd.Timedelta(days=1)
    g = named.groupby("CustomerID")
    customers = pd.DataFrame({
        "FirstPurchase": g["InvoiceDate"].min(),
        "LastPurchase": g["InvoiceDate"].max(),
        "NumOrders": g["InvoiceNo"].nunique(),
        "TotalSpend": g["TotalPrice"].sum(),
        "TotalItems": g["Quantity"].sum(),
        "Country": g["Country"].agg(lambda s: s.mode().iat[0]),
    })
    # RFM building blocks.
    customers["Recency"] = (snapshot - customers["LastPurchase"]).dt.days
    customers["Frequency"] = customers["NumOrders"]
    customers["Monetary"] = customers["TotalSpend"]
    customers["AvgOrderValue"] = customers["TotalSpend"] / customers["NumOrders"]
    customers["CohortMonth"] = customers["FirstPurchase"].dt.to_period("M").astype(str)
    customers["IsRepeat"] = customers["NumOrders"] > 1
    return customers.reset_index()


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def build_all(save: bool = True, write_csv: bool = False) -> dict[str, pd.DataFrame]:
    """Run the full pipeline and (optionally) persist every dataset.

    Parquet is the canonical processed format (compact, keeps dtypes/flags).
    Set write_csv=True to also export the cleaned master as CSV - useful for the
    zip submission, but it is ~100MB and is git-ignored, so it is off by default.
    """
    raw = load_raw()
    master = clean(raw)
    canon = build_canonical_descriptions(master)
    sales = build_sales(master)
    returns = build_returns(master)
    products = build_products(sales, canon)
    customers = build_customers(sales)

    datasets = {
        "master": master,
        "sales": sales,
        "returns": returns,
        "products": products,
        "customers": customers,
    }

    if save:
        PROCESSED.mkdir(parents=True, exist_ok=True)
        master.to_parquet(F_MASTER, index=False)
        sales.to_parquet(F_SALES, index=False)
        returns.to_parquet(F_RETURNS, index=False)
        products.to_parquet(F_PRODUCTS, index=False)
        customers.to_parquet(F_CUSTOMERS, index=False)
        if write_csv:
            # Cleaned master as CSV (rubric: "cleaned dataset"), in the
            # dictionary-specified encoding so the starter script can read it.
            master.to_csv(PROCESSED / "transactions_clean.csv",
                          index=False, encoding="ISO-8859-1")

    return datasets


# --------------------------------------------------------------------------- #
# Convenience loaders (used by analysis_helpers and the analysts)
# --------------------------------------------------------------------------- #
def _load_or_build(path: Path, key: str) -> pd.DataFrame:
    if path.exists():
        return pd.read_parquet(path)
    return build_all(save=True)[key]


def load_clean() -> pd.DataFrame:
    return _load_or_build(F_MASTER, "master")


def load_sales() -> pd.DataFrame:
    return _load_or_build(F_SALES, "sales")


def load_returns() -> pd.DataFrame:
    return _load_or_build(F_RETURNS, "returns")


def load_products() -> pd.DataFrame:
    return _load_or_build(F_PRODUCTS, "products")


def load_customers() -> pd.DataFrame:
    return _load_or_build(F_CUSTOMERS, "customers")


# --------------------------------------------------------------------------- #
# Data-quality report
# --------------------------------------------------------------------------- #
def quality_report(datasets: dict[str, pd.DataFrame]) -> str:
    master = datasets["master"]
    sales = datasets["sales"]
    returns = datasets["returns"]
    lines = []
    add = lines.append
    add("=" * 70)
    add(" " * 22 + "DATA QUALITY REPORT")
    add("=" * 70)
    add(f"Duplicate rows dropped       : {master.attrs.get('n_duplicates_dropped', 0):,}")
    add(f"Master rows (cleaned)        : {len(master):,}")
    add(f"  - flagged as guest         : {int(master['IsGuest'].sum()):,} "
        f"({master['IsGuest'].mean()*100:.1f}%)")
    add(f"  - flagged as return        : {int(master['IsReturn'].sum()):,}")
    add(f"  - flagged non-product      : {int((~master['IsProduct']).sum()):,}")
    add(f"  - zero/neg price           : {int(master['IsZeroOrNegPrice'].sum()):,}")
    add(f"  - quantity outliers        : {int(master['IsQtyOutlier'].sum()):,}")
    add(f"  - price outliers           : {int(master['IsPriceOutlier'].sum()):,}")
    add(f"  - missing description      : {int(master['IsMissingDescription'].sum()):,}")
    add("-" * 70)
    add(f"SALES rows (clean revenue)   : {len(sales):,}")
    add(f"  total revenue              : GBP {sales['TotalPrice'].sum():,.2f}")
    add(f"  unique invoices            : {sales['InvoiceNo'].nunique():,}")
    add(f"  unique customers (named)   : {sales['CustomerID'].nunique():,}")
    add(f"  unique products            : {sales['StockCode'].nunique():,}")
    add(f"  date range                 : {sales['InvoiceDate'].min()} -> {sales['InvoiceDate'].max()}")
    add("-" * 70)
    add(f"RETURNS rows                 : {len(returns):,}")
    add(f"  financial impact           : GBP {returns['TotalPrice'].sum():,.2f}")
    add("=" * 70)
    return "\n".join(lines)


if __name__ == "__main__":
    import sys

    write_csv = "--csv" in sys.argv
    print("Building clean datasets from raw Online Retail.xlsx ...")
    data = build_all(save=True, write_csv=write_csv)
    print(quality_report(data))
    print(f"\nProcessed files written to: {PROCESSED}")
    for f in sorted(PROCESSED.glob("*")):
        print(f"  - {f.name}")
