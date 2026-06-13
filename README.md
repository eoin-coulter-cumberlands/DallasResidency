# E-Commerce Business Analytics — Project 2

Analysis of one year (Dec 2010 – Dec 2011) of UCI *Online Retail* transaction
data (~540K rows) from a UK-based wholesale gift company, to surface trends,
customer behaviour, and actionable business insights.

## Repository layout

```
DallasResidency/
├── data/
│   ├── raw/                     # Online Retail.xlsx (source dataset)
│   └── processed/               # generated clean datasets (git-ignored)
├── src/
│   ├── data_engineering.py      # load → clean → feature-engineer → split
│   └── analysis_helpers.py      # ready-to-plot tables for all 10 analyses
├── docs/
│   ├── derived_data_dictionary.md  # new fields + cleaning justifications
│   └── starter_code_python.py      # original analysis template
├── requirements.txt
└── README.md
```

## Setup

```bash
# 1. (optional) create a virtual environment
python3 -m venv .venv && source .venv/bin/activate

# 2. install dependencies
pip install -r requirements.txt

# 3. build the clean datasets from the raw Excel file
python src/data_engineering.py
```

This reads `data/raw/Online Retail.xlsx` and writes the processed datasets to
`data/processed/`, then prints a data-quality report. Add `--csv` to also export
the cleaned master as CSV for the zip submission.

## Data engineering layer (this is the foundation for every analysis)

The pipeline does all loading/cleaning/feature-engineering **once** and produces
clean datasets so no analyst has to touch the raw file:

| Dataset | What it is |
|---------|-----------|
| `sales.parquet` | Canonical revenue rows (positive qty & price, not cancelled, real products) |
| `returns.parquet` | All returns/cancellations (Analysis 8) |
| `products.parquet` | Per-product revenue/quantity/orders summary |
| `customers.parquet` | Per-customer RFM / CLV / cohort building blocks |
| `transactions_clean.parquet` | Master: every row + flags (`IsGuest`, `IsReturn`, `IsProduct`, outlier flags, time fields) |

Cleaning decisions and every derived field are documented in
[`docs/derived_data_dictionary.md`](docs/derived_data_dictionary.md).

### Headline numbers (after cleaning)

- **536,641** clean transaction rows (5,268 duplicates removed)
- **£10.2M** total clean revenue across **19,773** orders
- **4,334** identified customers · **3,907** products · **38** countries
- **10,587** return/cancellation lines (−£894K impact)

## Using the helpers in an analysis

Each of the 10 required analyses has ready-to-use helper functions that return
tidy, plot-ready DataFrames:

```python
from src import analysis_helpers as ah

ah.monthly_revenue()           # Analysis 1: trend + MoM growth
ah.top_products_by_revenue()   # Analysis 2
ah.rfm_segments()              # Analysis 3: RFM scores + segments
ah.revenue_by_country()        # Analysis 4
ah.revenue_heatmap_day_hour()  # Analysis 5: day×hour matrix
ah.basket_summary()            # Analysis 6
ah.revenue_by_price_band()     # Analysis 7
ah.top_returned_products()     # Analysis 8
ah.cohort_retention_rate()     # Analysis 9
ah.revenue_forecast(periods=3) # Analysis 10
```

Run `python -m src.analysis_helpers` to smoke-test all helpers at once.

## Analysis ownership

| Role | Analyses |
|------|----------|
| Data Engineer | Data pipeline (all) + lead **1, 8** |
| Analyst 1 | Customer & cohort — **3, 9** |
| Analyst 2 | Product, basket, pricing — **2, 6, 7** |
| Visualization Specialist | Geographic, time, forecasting — **4, 5, 10** |
| Communications Lead | Report, slides, executive summary |

## Dataset citation

Daqing Chen, Sai Liang Sain, and Kun Guo (2012). *Data mining for the online
retail industry: A case study of RFM model-based customer segmentation using
data mining.* Journal of Database Marketing & Customer Strategy Management,
19(3), 197–208. doi:10.1057/dbm.2012.17
