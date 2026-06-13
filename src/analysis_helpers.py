"""
E-Commerce Business Analytics
Analysis Helpers
================

Owner: Data Engineer (consumed by all analysts)

These functions return clean, ready-to-plot tables for each of the 10 required
analyses. They sit on top of the datasets produced by `data_engineering.py`, so
no analyst has to re-clean the raw data or re-decide how to treat returns,
guests, or outliers.

Typical use inside a notebook / analysis script:

    from src import analysis_helpers as ah

    monthly = ah.monthly_revenue()          # Analysis 1
    top_rev = ah.top_products_by_revenue()  # Analysis 2
    rfm     = ah.rfm_segments()             # Analysis 3
    ...

Every function loads its inputs lazily via the cached parquet files, so they are
cheap to call repeatedly and independently.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from .data_engineering import (
    DOW_ORDER,
    load_customers,
    load_products,
    load_returns,
    load_sales,
)


# --------------------------------------------------------------------------- #
# Analysis 1: Sales Overview & Trends
# --------------------------------------------------------------------------- #
def headline_metrics() -> dict:
    """Top-line KPIs for the executive summary / dashboard header."""
    sales = load_sales()
    order_value = sales.groupby("InvoiceNo")["TotalPrice"].sum()
    return {
        "total_revenue": float(sales["TotalPrice"].sum()),
        "total_orders": int(sales["InvoiceNo"].nunique()),
        "total_customers": int(sales["CustomerID"].nunique()),
        "total_products": int(sales["StockCode"].nunique()),
        "countries": int(sales["Country"].nunique()),
        "avg_order_value": float(order_value.mean()),
        "median_order_value": float(order_value.median()),
        "date_start": sales["InvoiceDate"].min(),
        "date_end": sales["InvoiceDate"].max(),
    }


def monthly_revenue() -> pd.DataFrame:
    """Revenue, orders and growth rate per calendar month."""
    sales = load_sales()
    m = (
        sales.groupby("YearMonth")
        .agg(Revenue=("TotalPrice", "sum"),
             Orders=("InvoiceNo", "nunique"),
             Units=("Quantity", "sum"))
        .reset_index()
        .sort_values("YearMonth")
    )
    m["MoM_Growth"] = m["Revenue"].pct_change()
    return m


def weekly_revenue() -> pd.DataFrame:
    """Revenue per ISO week (for finer-grained trend lines)."""
    sales = load_sales()
    w = (
        sales.set_index("InvoiceDate")
        .resample("W")["TotalPrice"]
        .sum()
        .rename("Revenue")
        .reset_index()
    )
    return w


def seasonal_revenue() -> pd.DataFrame:
    """Revenue grouped by meteorological season."""
    sales = load_sales()
    order = ["Winter", "Spring", "Summer", "Autumn"]
    s = sales.groupby("Season")["TotalPrice"].sum().reindex(order).reset_index()
    s.columns = ["Season", "Revenue"]
    return s


# --------------------------------------------------------------------------- #
# Analysis 2: Product Performance
# --------------------------------------------------------------------------- #
def top_products_by_revenue(n: int = 10) -> pd.DataFrame:
    return load_products().nlargest(n, "TotalRevenue").reset_index(drop=True)


def top_products_by_quantity(n: int = 10) -> pd.DataFrame:
    return load_products().nlargest(n, "TotalQuantity").reset_index(drop=True)


def underperforming_products(n: int = 10, min_orders: int = 5) -> pd.DataFrame:
    """Lowest-revenue products that still sold often enough to be meaningful."""
    p = load_products()
    p = p[p["NumOrders"] >= min_orders]
    return p.nsmallest(n, "TotalRevenue").reset_index(drop=True)


def price_vs_volume() -> pd.DataFrame:
    """One row per product: avg price vs units sold (for a scatter/correlation)."""
    p = load_products()
    return p[["StockCode", "Description", "AvgUnitPrice",
              "TotalQuantity", "TotalRevenue"]].copy()


# --------------------------------------------------------------------------- #
# Analysis 3: Customer Behaviour & Segmentation
# --------------------------------------------------------------------------- #
def customer_frequency() -> pd.DataFrame:
    """Distribution of orders-per-customer plus repeat/one-time split."""
    c = load_customers()
    return c[["CustomerID", "NumOrders", "TotalSpend",
              "AvgOrderValue", "IsRepeat"]].copy()


def repeat_vs_onetime() -> pd.DataFrame:
    c = load_customers()
    out = (
        c["IsRepeat"].map({True: "Repeat", False: "One-time"})
        .value_counts()
        .rename_axis("Segment")
        .reset_index(name="Customers")
    )
    return out


def top_customers_by_clv(n: int = 10) -> pd.DataFrame:
    return load_customers().nlargest(n, "TotalSpend").reset_index(drop=True)


def rfm_segments() -> pd.DataFrame:
    """RFM scoring + simple named segments for Analysis 3.

    Each of Recency, Frequency, Monetary is scored 1-4 by quartile; the segment
    label is a coarse business grouping derived from the combined score.
    """
    c = load_customers().copy()

    # Recency: lower is better -> reverse the labels.
    c["R"] = pd.qcut(c["Recency"], 4, labels=[4, 3, 2, 1]).astype(int)
    c["F"] = pd.qcut(c["Frequency"].rank(method="first"), 4,
                     labels=[1, 2, 3, 4]).astype(int)
    c["M"] = pd.qcut(c["Monetary"].rank(method="first"), 4,
                     labels=[1, 2, 3, 4]).astype(int)
    c["RFM_Score"] = c["R"] + c["F"] + c["M"]

    def label(row) -> str:
        if row["R"] >= 3 and row["F"] >= 3 and row["M"] >= 3:
            return "Champions"
        if row["F"] >= 3 and row["M"] >= 3:
            return "Loyal"
        if row["R"] >= 3:
            return "Recent / Promising"
        if row["R"] <= 2 and row["F"] <= 2:
            return "At Risk / Lapsed"
        return "Needs Attention"

    c["Segment"] = c.apply(label, axis=1)
    return c


# --------------------------------------------------------------------------- #
# Analysis 4: Geographic Analysis
# --------------------------------------------------------------------------- #
def revenue_by_country() -> pd.DataFrame:
    sales = load_sales()
    g = (
        sales.groupby("Country")
        .agg(Revenue=("TotalPrice", "sum"),
             Orders=("InvoiceNo", "nunique"),
             Customers=("CustomerID", "nunique"))
        .sort_values("Revenue", ascending=False)
        .reset_index()
    )
    g["RevenueShare"] = g["Revenue"] / g["Revenue"].sum()
    return g


def domestic_vs_international() -> pd.DataFrame:
    sales = load_sales()
    region = np.where(sales["Country"] == "United Kingdom",
                      "Domestic (UK)", "International")
    out = (
        sales.assign(Region=region)
        .groupby("Region")
        .agg(Revenue=("TotalPrice", "sum"),
             Orders=("InvoiceNo", "nunique"))
        .reset_index()
    )
    return out


# --------------------------------------------------------------------------- #
# Analysis 5: Time-Based Patterns
# --------------------------------------------------------------------------- #
def revenue_by_day_of_week() -> pd.DataFrame:
    sales = load_sales()
    d = sales.groupby("DayOfWeek")["TotalPrice"].sum().reindex(DOW_ORDER)
    return d.rename("Revenue").reset_index()


def revenue_by_hour() -> pd.DataFrame:
    sales = load_sales()
    return sales.groupby("Hour")["TotalPrice"].sum().rename("Revenue").reset_index()


def revenue_heatmap_day_hour() -> pd.DataFrame:
    """Day-of-week x hour matrix of revenue (for a heatmap)."""
    sales = load_sales()
    pivot = sales.pivot_table(index="DayOfWeek", columns="Hour",
                              values="TotalPrice", aggfunc="sum",
                              fill_value=0).reindex(DOW_ORDER)
    return pivot


def weekend_vs_weekday() -> pd.DataFrame:
    sales = load_sales()
    out = (
        sales.assign(Bucket=np.where(sales["IsWeekend"], "Weekend", "Weekday"))
        .groupby("Bucket")["TotalPrice"].sum()
        .rename("Revenue").reset_index()
    )
    return out


# --------------------------------------------------------------------------- #
# Analysis 6: Basket Analysis
# --------------------------------------------------------------------------- #
def basket_summary() -> pd.DataFrame:
    """One row per order: item count, units, and value."""
    sales = load_sales()
    b = (
        sales.groupby("InvoiceNo")
        .agg(Items=("StockCode", "nunique"),
             Units=("Quantity", "sum"),
             Value=("TotalPrice", "sum"),
             Country=("Country", "first"))
        .reset_index()
    )
    return b


def basket_metrics() -> dict:
    b = basket_summary()
    return {
        "avg_items": float(b["Items"].mean()),
        "avg_units": float(b["Units"].mean()),
        "avg_value": float(b["Value"].mean()),
        "median_value": float(b["Value"].median()),
    }


# --------------------------------------------------------------------------- #
# Analysis 7: Pricing Analysis
# --------------------------------------------------------------------------- #
def revenue_by_price_band() -> pd.DataFrame:
    sales = load_sales()
    bins = [0, 1, 2, 5, 10, 20, 50, 100, np.inf]
    labels = ["<£1", "£1-2", "£2-5", "£5-10", "£10-20",
              "£20-50", "£50-100", "£100+"]
    band = pd.cut(sales["UnitPrice"], bins=bins, labels=labels, right=False)
    out = (
        sales.assign(PriceBand=band)
        .groupby("PriceBand", observed=True)
        .agg(Revenue=("TotalPrice", "sum"),
             Units=("Quantity", "sum"),
             Lines=("InvoiceNo", "count"))
        .reset_index()
    )
    return out


def price_points() -> pd.DataFrame:
    """Unit-price distribution (trimmed to the dictionary's sane range)."""
    sales = load_sales()
    return sales.loc[sales["UnitPrice"].between(0.01, 1000),
                     ["StockCode", "UnitPrice"]].copy()


# --------------------------------------------------------------------------- #
# Analysis 8: Returns / Cancellations
# --------------------------------------------------------------------------- #
def returns_summary() -> dict:
    returns = load_returns()
    sales = load_sales()
    return {
        "return_lines": int(len(returns)),
        "return_invoices": int(returns["InvoiceNo"].nunique()),
        "units_returned": int(abs(returns["Quantity"].sum())),
        "financial_impact": float(returns["TotalPrice"].sum()),
        "return_rate_vs_sales": float(len(returns) / (len(sales) + len(returns))),
    }


def top_returned_products(n: int = 10) -> pd.DataFrame:
    returns = load_returns()
    g = (
        returns.groupby("StockCode")
        .agg(Description=("Description", "first"),
             UnitsReturned=("Quantity", lambda s: int(abs(s.sum()))),
             ValueReturned=("TotalPrice", "sum"))
        .sort_values("UnitsReturned", ascending=False)
        .head(n)
        .reset_index()
    )
    return g


def returns_over_time() -> pd.DataFrame:
    returns = load_returns()
    r = (
        returns.groupby("YearMonth")
        .agg(ReturnValue=("TotalPrice", "sum"),
             ReturnLines=("InvoiceNo", "count"))
        .reset_index()
        .sort_values("YearMonth")
    )
    return r


def returns_by_country() -> pd.DataFrame:
    returns = load_returns()
    return (
        returns.groupby("Country")["TotalPrice"].sum()
        .sort_values()
        .rename("ReturnValue")
        .reset_index()
    )


# --------------------------------------------------------------------------- #
# Analysis 9: Cohort Analysis
# --------------------------------------------------------------------------- #
def cohort_retention() -> pd.DataFrame:
    """Monthly acquisition cohorts x retention by months-since-first-purchase.

    Returns a matrix: rows = cohort month, columns = period index (0,1,2,...),
    values = number of active customers from that cohort in that period.
    """
    sales = load_sales()
    s = sales[~sales["IsGuest"]].copy()

    s["OrderMonth"] = s["InvoiceDate"].dt.to_period("M")
    cohort = s.groupby("CustomerID")["OrderMonth"].min().rename("CohortMonth")
    s = s.join(cohort, on="CustomerID")

    s["PeriodIndex"] = (
        (s["OrderMonth"].dt.year - s["CohortMonth"].dt.year) * 12
        + (s["OrderMonth"].dt.month - s["CohortMonth"].dt.month)
    )

    counts = (
        s.groupby(["CohortMonth", "PeriodIndex"])["CustomerID"]
        .nunique()
        .reset_index()
    )
    matrix = counts.pivot(index="CohortMonth", columns="PeriodIndex",
                          values="CustomerID")
    matrix.index = matrix.index.astype(str)
    return matrix


def cohort_retention_rate() -> pd.DataFrame:
    """Cohort matrix expressed as a retention percentage of the cohort size."""
    matrix = cohort_retention()
    return matrix.divide(matrix[0], axis=0)


# --------------------------------------------------------------------------- #
# Analysis 10: Forecasting & Predictions
# --------------------------------------------------------------------------- #
def revenue_forecast(periods: int = 3) -> pd.DataFrame:
    """Simple linear-trend revenue forecast for the next `periods` months.

    Uses ordinary least squares on monthly revenue. Kept deliberately light so
    the Visualization Specialist (Analysis 10 lead) can swap in a richer model
    later; this provides a defensible baseline + trend line.
    """
    m = monthly_revenue().copy()
    # The final month is partial (data ends 2011-12-09) - exclude it from the fit.
    fit = m.iloc[:-1] if len(m) > 1 else m
    x = np.arange(len(fit))
    y = fit["Revenue"].to_numpy(dtype=float)
    slope, intercept = np.polyfit(x, y, 1)

    future_x = np.arange(len(fit), len(fit) + periods)
    last_period = pd.Period(fit["YearMonth"].iloc[-1], freq="M")
    future_months = [(last_period + i + 1).strftime("%Y-%m") for i in range(periods)]

    forecast = pd.DataFrame({
        "YearMonth": future_months,
        "Revenue": slope * future_x + intercept,
        "Type": "Forecast",
    })
    actual = fit[["YearMonth", "Revenue"]].assign(Type="Actual")
    return pd.concat([actual, forecast], ignore_index=True)


# --------------------------------------------------------------------------- #
# Smoke test: confirm every helper returns something usable.
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    checks = {
        "headline_metrics": headline_metrics,
        "monthly_revenue": monthly_revenue,
        "weekly_revenue": weekly_revenue,
        "seasonal_revenue": seasonal_revenue,
        "top_products_by_revenue": top_products_by_revenue,
        "top_products_by_quantity": top_products_by_quantity,
        "underperforming_products": underperforming_products,
        "price_vs_volume": price_vs_volume,
        "customer_frequency": customer_frequency,
        "repeat_vs_onetime": repeat_vs_onetime,
        "top_customers_by_clv": top_customers_by_clv,
        "rfm_segments": rfm_segments,
        "revenue_by_country": revenue_by_country,
        "domestic_vs_international": domestic_vs_international,
        "revenue_by_day_of_week": revenue_by_day_of_week,
        "revenue_by_hour": revenue_by_hour,
        "revenue_heatmap_day_hour": revenue_heatmap_day_hour,
        "weekend_vs_weekday": weekend_vs_weekday,
        "basket_summary": basket_summary,
        "basket_metrics": basket_metrics,
        "revenue_by_price_band": revenue_by_price_band,
        "price_points": price_points,
        "returns_summary": returns_summary,
        "top_returned_products": top_returned_products,
        "returns_over_time": returns_over_time,
        "returns_by_country": returns_by_country,
        "cohort_retention": cohort_retention,
        "cohort_retention_rate": cohort_retention_rate,
        "revenue_forecast": revenue_forecast,
    }
    print(f"Running smoke test on {len(checks)} helpers ...\n")
    for name, fn in checks.items():
        out = fn()
        if isinstance(out, dict):
            shape = f"dict({len(out)} keys)"
        else:
            shape = f"{out.shape[0]} rows x {out.shape[1]} cols"
        print(f"  [ok] {name:30s} -> {shape}")
    print("\nAll helpers returned data successfully.")
