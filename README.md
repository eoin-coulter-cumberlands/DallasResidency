# E-Commerce Business Analytics - Project 2

This project uses the UCI Online Retail dataset from a UK-based wholesale
company to analyze one year of e-commerce transactions, identify revenue
drivers, understand customer behavior, evaluate return risk, and recommend
actions for sales growth.

## Repository Layout

```text
data/
  raw/                     Online Retail.xlsx source dataset
  processed/               Cleaned parquet/CSV datasets
src/
  data_engineering.py      Data loading, cleaning, feature engineering, exports
  analysis_helpers.py      Analysis tables for the 10 required analyses
  generate_deliverables.py Figures, tables, dashboard, report, deck, manifest
docs/
  derived_data_dictionary.md
deliverables/
  dashboard/               Interactive HTML dashboard
  figures/                 PNG visualizations
  presentation/            Final PowerPoint deck
  report/                  Final Word technical report
  tables/                  CSV output tables
requirements.txt
SUBMISSION_CHECKLIST.md
```

## Setup

```bash
python -m venv .venv
pip install -r requirements.txt
```

## Run the Project

Build or refresh the cleaned datasets:

```bash
python src/data_engineering.py
```

Generate project outputs from the cleaned datasets:

```bash
python src/generate_deliverables.py
```

The data pipeline reads `data/raw/Online Retail.xlsx`, writes cleaned datasets to
`data/processed/`, and prints a data-quality summary. Add `--csv` when running
`src/data_engineering.py` to export a cleaned master CSV.

## Code Scope

The code files are limited to the functions required for this project:

- imports and setup
- data loading
- data cleaning and feature engineering
- sales, product, customer, geographic, time, basket, pricing, returns, cohort,
  and forecasting analyses
- visualizations
- exports for tables, figures, dashboard, report, presentation, and manifest
- concise printed summaries for reproducible command-line runs

## Required Analyses

1. Sales overview and trends
2. Product performance analysis
3. Customer behavior and segmentation
4. Geographic sales analysis
5. Time-based patterns
6. Basket analysis
7. Pricing analysis
8. Returns/cancellations analysis
9. Cohort analysis
10. Sales forecasting and predictions

## Final Deliverables

- `deliverables/dashboard/ecommerce_dashboard.html`
- `deliverables/presentation/ecommerce_business_analytics_deck.pptx`
- `deliverables/report/ecommerce_technical_report.docx`
- `deliverables/figures/*.png`
- `deliverables/tables/*.csv`
- `deliverables/deliverables_manifest.json`

## Dataset Citation

Daqing Chen, Sai Liang Sain, and Kun Guo (2012). *Data mining for the online
retail industry: A case study of RFM model-based customer segmentation using
data mining.* Journal of Database Marketing & Customer Strategy Management,
19(3), 197-208. doi:10.1057/dbm.2012.17
