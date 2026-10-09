# Qualcomm Financial Revenue Forecast & Earnings Variance Analysis
*Forecasting Qualcomm's shift from phones to cars & AI*

An FP&A-style project analyzing Qualcomm (NASDAQ: QCOM) segment performance: guidance vs. actual variance, driver-based rolling forecast, and scenario analysis.

**Core question:** Can QCT Automotive, IoT, and Data Center growth offset handset decline through FY27?

## Scope
- **Segments:** QCT (Handsets, Automotive, IoT), QTL (Licensing)
- **Period:** FY2024 Q1 – FY2026 Q4 actuals, FY2027 forecast
- **Granularity:** Quarterly (matches public reporting)

## Deliverables
| Component | Tool | Location |
|---|---|---|
| Driver-based forecast + variance model | Excel | `model/` |
| CFO-style dashboard | Tableau Public | [link TBD] |
| Recommendation memo | PDF | `docs/` |

## Model Structure
1. **Assumptions** – all inputs (blue), scenario toggles
2. **Actuals** – segment revenue and EBT by quarter
3. **Guidance** – company guidance ranges and midpoints
4. **Variance** – actual vs. guidance midpoint, $ and %
5. **Drivers** – same-quarter YoY growth by stream (seasonality-preserving), EBT = revenue × margin
6. **Forecast** – rolling 4–6 quarter projection, base/bull/bear
7. **Checks** – totals tie to reported consolidated revenue

## Repository Layout
```
qualcomm-revenue-forecast/
├── data/
│   ├── raw/              # Earnings releases, 10-Q/10-K excerpts
│   └── processed/        # Clean CSVs feeding Excel + Tableau
├── model/                # Excel forecast & variance model
├── tableau/              # Tableau workbook (.twbx)
├── docs/                 # Memo, assumptions log, source log
├── scripts/              # Optional data-pull helpers
└── images/               # Dashboard screenshots
```

## Data Sources
All data comes from Qualcomm's public SEC filings (10-Q, 10-K) and investor relations earnings releases. Full log in `docs/sources.csv`.

## Disclaimer
This is an independent portfolio project built solely from public information. It is not affiliated with or endorsed by Qualcomm. It contains no internal or confidential data, and it is not investment advice.
