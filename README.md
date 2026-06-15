# PenBev CCPB — Business Intelligence Platform

Multi-dimensional business intelligence dashboards for **PenBev**, covering consumer channel performance, product health, and business KPI tracking — built with HTML5 / JavaScript and fully self-contained (no server required).

---

## Overview

| Dashboard | Focus |
|-----------|-------|
| Business Intelligence | Revenue, channel mix, trade spend, ROI |
| Health Dashboard | Product performance health metrics |
| Product Master | Full SKU catalog with pricing & grading |

---

## Dashboards

### `penbev_ccpb_business_dashboard.html`
Core BI dashboard covering:
- Channel revenue breakdown (mass retail, trade, e-commerce)
- Trade spend vs ROI analysis
- Period-over-period KPI tracking
- Outlet distribution and penetration

### `penbev_ccpb_health_dashboard.html`
Product health monitoring:
- Weighted distribution scores
- Off-take velocity by category
- Promo vs base split
- Category health index

### `penbev_ccpb_product_master.html`
Full product catalog:
- SKU master with channel pricing
- Pack size and format hierarchy
- Product grading (A/B/C/D)
- Listing status by retailer

### `penbev_dashboard.html`
Executive summary dashboard with consolidated KPIs across all functions.

---

## Data Files

| File | Description |
|------|-------------|
| `data/penbev_product_master.csv` | SKU-level product master data |
| `data/PenBev_Sample_Data.xlsx` | Sample dataset for analysis |
| `data/PenBev.xlsx` | Full data workbook with pivot tables |

---

## Repository Structure

```
PenBev-Analytics/
├── dashboards/
│   ├── penbev_ccpb_business_dashboard.html
│   ├── penbev_ccpb_health_dashboard.html
│   ├── penbev_ccpb_product_master.html
│   └── penbev_dashboard.html
└── data/
    ├── penbev_product_master.csv
    ├── PenBev_Sample_Data.xlsx
    └── PenBev.xlsx
```

---

## Running the Dashboards

All dashboards are self-contained HTML files. Open any file directly in a browser:

```
# Windows
start dashboards/penbev_ccpb_business_dashboard.html
```

No server, no build step, no dependencies needed.

---

## Tech Stack

| Technology | Usage |
|------------|-------|
| HTML5 / CSS3 | Dashboard layout and styling |
| JavaScript (Vanilla) | Interactive charts and filtering |
| Chart.js | Data visualisation |
| Excel | Pivot analysis and data workbooks |

---

## Author

**Anthony Apollis** — Data Engineer & Analytics Specialist  
[GitHub](https://github.com/anthonyapollis) · [Portfolio](https://anthonyapollis.github.io)
