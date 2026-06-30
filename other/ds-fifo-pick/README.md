# DS FIFO Pick Analyzer


## Project Overview

FIFO compliance analysis for the **DS warehouse** (SITE3 / CUST1). Reads picking events plus a stack of inventory snapshots, decodes vanning (manufacturing) dates from module-number prefixes, and emits a Power BI Excel workbook plus a `.pbix` for review. Sister project to `Other/IP FIFO/` (same algorithm, different warehouse).

This is the older / pre-v2 form of the FIFO Pick analyzer. The actively maintained version for SITE1 lives at `KPI/FIFO Pick/indiana_fifo_analyzer v2.0.py`.

## Running the script

```
python ds_fifo_analyzer.py
```

No CLI flags — paths are hardcoded inside `FIFOAnalyzer.__init__`. The `output_folder` default points at `…\KPI\DS FIFO Pick\Power BI Data\` (a path that **does not exist** at the parent KPI level — see Known issues).

Dependencies: `pandas` (heapq, defaultdict, datetime are stdlib). No `requirements.txt`.

## Folder structure

```
Other/DS FIFO Pick/
  ds_fifo_analyzer.py
  Data/
    Picking_MODULE_LIST.csv         # picking events (PROCESS TIME, PRODUCT CODE, MODULE NO, ...)
    MODULE_LOC Files/
      MODULE_LOC_SITE3_CUST1_YYYYMMDD_HHMMSS.csv   # daily inventory snapshots (one or more)
  Power BI Data/
    FIFO_PowerBI_Export.xlsx        # main consumable workbook
  Output/
    DS_FIFO_KPI_Data_YYYYMMDD.xlsx  # archived weekly snapshots
    DS FIFO Report.pbix
    DS_FIFO_Pick.pbix
```

## Data sources (`Data/`)

| File | Provides |
|---|---|
| `Picking_MODULE_LIST.csv` | Pick-scan events. Columns include `PROCESS TIME`, `PRODUCT CODE`, `MODULE NO`, plus operator/dock/shift fields used for grouping. |
| `MODULE_LOC Files/MODULE_LOC_SITE3_CUST1_*.csv` | Periodic WMS inventory dumps (one per `~`8 hours) for warehouse `SITE3` (CUST1). Loaded as a stack and deduplicated on `MODULE#` (keeps first = latest snapshot listed). |

## Business logic / KPIs

### Vanning-date decoding (`extract_vanning_date`)

Decodes manufacturing date from the module-number prefix using **Sachiko's PQ_VanningDate lookup tables** (replicated in code; do not change without updating both warehouses):

- **DateTable1 (Japan)** — prefixes `KJ699/KJ900/KJ999/22200/26700/27100/2Z400/2G400/WN000`, plus single-letter `S`/`K` defaults. 6th char = month letter (`M..X` → 1..12 or `A..L` → 1..12).
- **DateTable2 (Mexico/China/Canada)** — `KJ550..KJ698`, `KJ540`, `KA079..KA277`, `KJ912`, plus 4-char `KJ50..KJ69`. 6th–7th chars = month digits.
- **DateTable3 (Thailand)** — `KJ911`, `KA085..KA331`, `KA118`. 6th char = year digit (`1=2021`, `2=2022`, …); has a baked-in 2-month offset.
- **No DT (US/IS)** — `KJ0..KJ4`, `KJ7..KJ8` → no vanning date returned.
- **AAA modules** — when `'AAA' in module_str` and ETA is present: DT2 → first of ETA's month/year; DT1/DT3 → ETA minus 2 months.

### FIFO compliance (`check_fifo`)

For each pick: identify the oldest still-available module of the same `PRODUCT` (min-heap per part keyed on `VANNING DATE`); flag the pick `fifo_compliant=1` iff the picked module's vanning date equals the part's oldest available vanning date. Modules already picked are removed from availability.

`export_oldest_location_file` separately produces a per-part lookup of the single oldest available module's `LOCATION` for warehouse-floor coaching.

## Outputs

| File | Destination | Purpose |
|---|---|---|
| `FIFO_PowerBI_Export.xlsx` | `Power BI Data/` | Active workbook consumed by the `.pbix` files |
| `DS_FIFO_KPI_Data_YYYYMMDD.xlsx` | `Output/` | Dated weekly archive |
| `DS FIFO Report.pbix`, `DS_FIFO_Pick.pbix` | `Output/` | Power BI reports — TODO: confirm which is the live one |

## Schedules / triggers

Manual weekly run by the operator. No Power Automate, no scheduled task.

## Configuration

All hardcoded inside `ds_fifo_analyzer.py`:

```python
output_folder = './DS FIFO Pick/Power BI Data/'
```

The path uses a top-level `KPI/DS FIFO Pick/` that doesn't exist on disk — the actual file lives under `KPI/Other/DS FIFO Pick/`. Either the script's default has drifted from the folder structure, or the operator passes a different path when instantiating `FIFOAnalyzer`. Verify before assuming the default works as-is.

## Dependencies

`pandas` (uses `pandas.to_datetime`, groupby, heapq integration). Stdlib: `datetime`, `heapq`, `collections.defaultdict`, `pathlib`. Install with `pip install pandas openpyxl`.

## Known issues / gotchas

- **Hardcoded `output_folder` references a non-existent `KPI/DS FIFO Pick/` (sibling to `Other/`)**. Either fix the default in the script or pass an explicit folder when constructing `FIFOAnalyzer`.
- **Two `.pbix` files in `Output/`** with similar names — confirm with the user which is current.
- Vanning-date decoder is duplicated across `Other/DS FIFO Pick/`, `Other/IP FIFO/`, and `KPI/FIFO Pick/` (newer v2.0 in SITE1). Changes to the prefix tables must be propagated to all three.
- `desktop.ini` is a Windows artifact — ignore.
- This project sits inside `Other/`, suggesting it is a legacy/inactive variant. The SITE1 equivalent (`KPI/FIFO Pick/indiana_fifo_analyzer v2.0.py`) is the v2 reference implementation; if changing FIFO logic, port from there.