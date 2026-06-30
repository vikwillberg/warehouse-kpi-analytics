# IP FIFO Analyzer


## Project Overview

FIFO compliance analysis for the **IP warehouse** (SITE2 / CUST1). Reads picking events plus a stack of inventory snapshots, decodes vanning (manufacturing) dates from module-number prefixes, and emits a Power BI Excel workbook for downstream reports. Sister project to `Other/DS FIFO Pick/` (same algorithm, different warehouse).

This is a legacy / pre-v2 form of the FIFO analyzer. The actively maintained reference for SITE1 is `KPI/FIFO Pick/indiana_fifo_analyzer v2.0.py`.

## Running the script

```
python ip_fifo_analyzer.py
```

No CLI flags. The `output_folder` default points at `…\KPI\IP FIFO\Power BI Data\` (a top-level folder that **does not exist** — see Known issues).

Dependencies: `pandas`. No `requirements.txt`.

## Folder structure

```
Other/IP FIFO/
  ip_fifo_analyzer.py
  Data/
    Picking_MODULE_LIST.csv          # picking events
    MODULE_LOC Files/
      MODULE_LOC_SITE2_CUST1_YYYYMMDD_HHMMSS.csv   # daily inventory snapshots (SITE2 / CUST1 = IP warehouse)
  Power BI Data/
    FIFO_PowerBI_Export.xlsx          # consumable workbook
  Output/
    IP_FIFO_KPI_Data_YYYYMMDD.xlsx    # dated archives
```

## Data sources (`Data/`)

| File | Provides |
|---|---|
| `Picking_MODULE_LIST.csv` | Pick-scan events: `PROCESS TIME`, `PRODUCT CODE`, `MODULE NO`, plus operator/dock/shift fields. |
| `MODULE_LOC Files/MODULE_LOC_SITE2_CUST1_*.csv` | WMS inventory dumps for warehouse `SITE2` (CUST1, the IP site). Loaded as a stack, deduplicated on `MODULE#` (first wins). |

## Business logic

### Vanning-date decoding (`extract_vanning_date`)

Identical to the DS variant — decodes the manufacturing date from the module-number prefix using **Sachiko's PQ_VanningDate** lookup tables:

- **DateTable1 (Japan)** — `KJ699/KJ900/KJ999/22200/26700/27100/2Z400/2G400/WN000` (5-char) plus single-letter `S`/`K` default. Month from 6th char letter.
- **DateTable2 (Mexico/China/Canada)** — 5-char `KJ550..KJ698, KJ540, KA079..KA277, KJ912`; 4-char `KJ50..KJ69`. Month from 6th–7th chars.
- **DateTable3 (Thailand)** — `KJ911, KA085..KA331, KA118`. Year from 6th char (1→2021, 2→2022 …); 2-month offset.
- **No DT (US/IS)** — `KJ0..KJ4, KJ7..KJ8` → returns `None`.
- **AAA modules** — DT2 → first of ETA month/year; DT1/DT3 → ETA minus 2 months.

### FIFO compliance (`check_fifo`)

Min-heap per `PRODUCT` keyed on `VANNING DATE`. A pick is `fifo_compliant=1` iff the picked module's vanning date equals the oldest still-available module's vanning date for that part. Picked modules are removed from availability.

## Outputs

| File | Destination | Purpose |
|---|---|---|
| `FIFO_PowerBI_Export.xlsx` | `Power BI Data/` | Power BI consumable workbook |
| `IP_FIFO_KPI_Data_YYYYMMDD.xlsx` | `Output/` | Dated weekly archives |

No `.pbix` lives in this folder — the consuming Power BI report likely lives elsewhere (sibling reports / SharePoint). TODO: confirm the consumer.

## Schedules / triggers

Manual weekly run. No Power Automate, no scheduled task.

## Configuration

All hardcoded inside `ip_fifo_analyzer.py`:

```python
output_folder = './IP FIFO/Power BI Data/'
```

References `KPI/IP FIFO/` at the top level (sibling to `Other/`); the actual location is `KPI/Other/IP FIFO/`. Same drift as the DS variant.

## Dependencies

`pandas`. Stdlib: `datetime`, `heapq`, `collections.defaultdict`, `pathlib`. Install with `pip install pandas openpyxl`.

## Known issues / gotchas

- **Hardcoded `output_folder` path uses `KPI/IP FIFO/` (top-level), but the script lives under `KPI/Other/IP FIFO/`**. The default path won't write to this folder unless overridden.
- Vanning-date decoder is duplicated across `Other/IP FIFO/`, `Other/DS FIFO Pick/`, and `KPI/FIFO Pick/`. Keep prefix tables in sync.
- The SITE1 v2 sibling (`KPI/FIFO Pick/indiana_fifo_analyzer v2.0.py`) is the modern reference; port logic *from* there *to* here, not the other way around.
- `desktop.ini` is a Windows artifact — ignore.