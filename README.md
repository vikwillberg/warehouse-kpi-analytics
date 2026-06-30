# Warehouse KPI & Analytics

A portfolio of production data-analytics tools I built to automate **warehouse operations
reporting** — turning raw warehouse-management-system (WMS) exports into weekly/daily KPI
reports, operator performance scoring, FIFO compliance analysis, inventory rundowns, and
live operational dashboards.

Each tool is a self-contained Python (or PowerShell) pipeline that reads raw CSV/Excel
exports, transforms and aggregates them with **pandas/numpy**, and produces interactive
**Plotly/matplotlib** HTML reports and Power BI–ready datasets. Several run on a weekly
cadence and feed Power BI dashboards via SharePoint-synced folders.

![Daily Shipping KPI dashboard](in-daily-shipping-kpi/screenshots/daily-shipping-dashboard.png)

<sub>Example output — the Daily Shipping KPI dashboard. Each project folder includes a
redacted screenshot of its report.</sub>

> **Note on data & privacy.** This is a **code + documentation** repository only. All
> proprietary operational data, customer identifiers, employee names, internal file paths,
> credentials, and the employer's name have been removed or genericized. No raw data or
> `.pbix` files are included. The report **screenshots** in this repo are previews only —
> operational volumes are shown as generated, while customer/supplier names, order/part
> identifiers, and employee names have been **redacted or replaced with placeholders**
> (`Site 1`, `CUST1`, `Operator N`, `Example Logistics`). The code is shared to demonstrate
> engineering approach, not to reproduce any real dataset.

---

## What this demonstrates

- **End-to-end data pipelines** — ingest raw WMS exports → clean/normalize → aggregate →
  report, all in maintainable single-file or orchestrator/engine designs.
- **Operational analytics** — operator efficiency scoring (data-driven benchmarks, sigmoid/
  logistic curves, Bayesian shrinkage), FIFO compliance, shortage and capacity analysis.
- **Reporting & visualization** — self-contained interactive HTML reports (Plotly + custom
  JS/CSS), base64-embedded matplotlib charts, Excel outputs, and email-ready snapshots.
- **Power BI integration** — scripts emit clean CSV models consumed by Power BI; one project
  keeps an embedded Power BI tile alive on a wall-mounted TV via a scheduled-task failsafe.
- **Automation & robustness** — date-range auto-calculation, shift/overnight handling,
  graceful degradation when optional dependencies (Plotly, scikit-learn) are missing,
  and Windows Task Scheduler deployment.

## Tech stack

`Python 3.13` · `pandas` · `numpy` · `matplotlib` · `Plotly` · `scikit-learn` ·
`openpyxl` · `Power BI` · `PowerShell` · `Windows Task Scheduler`

---

## Projects

| Project | What it does | Key tech |
|---|---|---|
| [Inbound Sorting KPI](in-sorting-kpi/) | Operator sorting/putaway performance pipeline with data-driven efficiency scoring and optional ML volume forecasting | pandas, plotly, scikit-learn |
| [Daily Shipping KPI](in-daily-shipping-kpi/) | Daily shipment-status reporting (Allocated→Picked→Shipped) with HTML + Excel + PNG snapshot outputs | pandas, plotly, openpyxl |
| [Monthly KPIs](in-monthly-kpis/) | Three monthly report generators — Picking, Shipping, and Shortage — as interactive Plotly HTML | pandas, plotly |
| [Inbound FIFO Pick Compliance](in-fifo-pick/) | Weekly FIFO (first-in-first-out) pick-compliance analysis from arrival vs. pick timing | pandas, numpy |
| [Multi-Warehouse FIFO Inventory Rundown](fifo-inventory-rundown/) | Single multi-site HTML report for arrival-based FIFO inventory rundown across warehouses | pandas |
| [Inbound Capacity Calculation](in-capacity-calculation/) | Inbound receiving/capacity planning calculations | pandas, numpy |
| [DS Sorting KPI](ds-sorting-kpi/) | Weekly sorting & putaway KPI report with embedded chart + Power BI CSV exports | pandas, matplotlib |
| [IP Picking KPI](ip-picking-kpi/) | Weekly picking KPI with an orchestrator + engine design (subprocess + JSON config), SharePoint sync | pandas, openpyxl |
| [IP Sorting KPI](ip-sorting-kpi/) | Weekly IP sorting KPI report and Power BI dataset export | pandas |
| [Picking TV Dashboard + Token-Refresh Failsafe](in-picking-dashboard/) | Keeps an embedded Power BI tile alive on a wall TV via a PowerPoint restart scheduled task | PowerShell, Task Scheduler |
| [Other tools](other/) | Smaller utilities: DS/IP FIFO analyzers, DS aged-inventory, and an inbound next-week plan & forecast report | pandas |

Each project folder has its own `README.md` documenting how it runs, its data flow,
configuration, and outputs.

---

## Repository layout

```
warehouse-kpi-analytics/
├── in-sorting-kpi/            # operator efficiency scoring + forecasting
├── in-daily-shipping-kpi/     # daily shipment status report
├── in-monthly-kpis/           # monthly picking / shipping / shortage reports
├── in-fifo-pick/              # FIFO pick-compliance analyzer
├── fifo-inventory-rundown/    # multi-warehouse FIFO inventory rundown
├── in-capacity-calculation/   # inbound capacity planning
├── ds-sorting-kpi/            # DS sorting & putaway KPI
├── ip-picking-kpi/            # IP picking (orchestrator + engine)
├── ip-sorting-kpi/            # IP sorting KPI
├── in-picking-dashboard/      # TV dashboard token-refresh failsafe
├── other/                     # smaller auxiliary tools
├── requirements.txt
└── .gitignore                 # blocks data/binary files from ever being committed
```

## Running

```bash
pip install -r requirements.txt
```

The scripts were written to run on Windows against local/SharePoint-synced folders, so
file paths in the code are **relative placeholders** (e.g. `./Data/`, `./bi_data/`). To run
a project against your own data, point those paths at your files and drop the expected CSV
exports into the project's data folder. See each project's `README.md` for the exact inputs.

---

## About

Built by **Viktor Berg** — data analyst / developer focused on warehouse and logistics
operations analytics. These tools were developed and run in a real 3PL warehouse environment
to replace manual reporting with automated, repeatable pipelines.
