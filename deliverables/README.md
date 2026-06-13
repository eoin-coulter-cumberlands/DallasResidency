# Deliverables

This folder contains the final project submission assets generated from the
cleaned e-commerce dataset.

## Primary submission files

- `dashboard/ecommerce_dashboard.html`
- `report/ecommerce_technical_report.docx`
- `presentation/ecommerce_business_analytics_deck.pptx`
- `executive_summary.md`

## Supporting assets

- `figures/` stores chart PNGs used in the dashboard, report, and slides.
- `tables/` stores CSV extracts for the main analysis outputs.
- `deliverables_manifest.json` stores headline metrics and artifact locations.

## Regeneration

From the repository root:

```bash
python3 src/data_engineering.py --csv
python3 src/generate_deliverables.py
```
