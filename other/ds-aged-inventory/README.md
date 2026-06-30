# DS Aged Inventory


## Project Overview

Tracks **aged inventory** at the DS warehouse (SITE3 / CUST1) for the OEM client's "zero aged stock" initiative. Modules with a vanning (manufacturing) date on or before a configured cutoff (default `2025-03-01`) are flagged. The Python report under `New/` is the modern automated form; the loose `.xlsx`/`.xlsm` files at this folder's root are the legacy spreadsheet workflow this script is replacing.

the OEM client's stated target is to reach **zero aged inventory by 2026-03-31** (`TARGET_ZERO_DATE` in the script — informational only, not used in calculations).

## Folder structure

```
Other/DS Aged Inventory/
  DS Aged Invetory list 20260420.xlsx     # weekly count + detail (legacy spreadsheet output)
  DS_Vanning Date.xlsm                    # legacy macro-enabled vanning-date decoder
  Inventory Data DS work 20260316.xlsx    # legacy work file
  New/
    ds_aged_inventory.py                  # Python replacement (active)
    502.csv                               # current inventory snapshot
    201P.csv                              # planned outbound orders
    201S.csv                              # shipped/in-progress orders
    DS_Aged_Inventory_YYYYMMDD.xlsx       # generated report
```

The legacy spreadsheet name is misspelled "Invetory" — preserved as-is in the filename.

## Running the Python report

```
cd "Other/DS Aged Inventory/New"
python ds_aged_inventory.py
```

The script reads `502.csv`, `201P.csv`, and `201S.csv` from its own directory (paths are derived from `__file__`) and writes a date-stamped `DS_Aged_Inventory_YYYYMMDD.xlsx` next to it.

Dependencies: `pandas`, `numpy`, `python-dateutil`, `openpyxl`. No `requirements.txt`.

## Data sources (`New/`)

| File | Provides | Notes |
|---|---|---|
| `502.csv` | On-hand inventory with `MODULE#`, `LOCATION`, `PRODUCT`, `QUANTITY`, `ETA` | Used as the universe of modules being checked for age |
| `201P.csv` | Planned outbound orders | Cross-referenced to flag aged stock that already has a planned ship |
| `201S.csv` | Shipped / in-progress orders | Same purpose — distinguishes truly-stuck from soon-to-ship |

## Business logic

### Aged-cutoff flag

A module is **aged** iff its decoded vanning date ≤ `AGED_CUTOFF_DATE` (default `2025-03-01`). The cutoff is a configuration constant near the top of the script — adjust per the OEM client's review cadence.

### Vanning-date decoding (`ORIGIN_MAP`)

The script uses an explicit `ORIGIN_MAP` keyed on module-number prefix (longest match wins) to derive **(origin_name, date_table)**:

- **DateTable 1 (Japan-style)** — month from a single letter at position 6.
- **DateTable 2 (China/Canada/Mexico-style)** — month from digits at positions 6–7.
- **DateTable 3 (Thailand-style)** — year from a single digit, with a 2-month offset applied via the ETA.

Origin names exposed in this map go beyond the `KPI/FIFO Pick` family (e.g. `MMVO`, `JS`, `3RD(THAI)`, `3RD(CHINA)`) — the friendly origin label is part of the deliverable.

### Output sheet structure

The generated `DS_Aged_Inventory_YYYYMMDD.xlsx` contains formatted summary + detail sheets (count of aged modules, list of every aged module with location/product/quantity/age-in-days, and origin breakdown). Formatting is done via `openpyxl` styles in-script; the OEM client's reviewers consume this directly.

## Schedules / triggers

Manual run by the operator (typically weekly, sometimes after a major shipment). No Power Automate, no scheduled task.

## Configuration

Edit constants at the top of `ds_aged_inventory.py`:

| Constant | Default | Purpose |
|---|---|---|
| `AGED_CUTOFF_DATE` | `"2025-03-01"` | Modules with vanning date ≤ this are aged |
| `TARGET_ZERO_DATE` | `"2026-03-31"` | Informational only — the OEM client's zero-aged target |
| `FILE_502`, `FILE_201P`, `FILE_201S` | derived from `__file__` | Input paths |
| `OUTPUT_FILE` | `DS_Aged_Inventory_<TODAY>.xlsx` | Date-stamped output path |

## Dependencies

`pandas`, `numpy`, `python-dateutil`, `openpyxl`. Tested on Python 3.13.

## Known issues / gotchas

- **Filename misspelling** ("Invetory") is intentional — matches what the OEM client already files this report under.
- Legacy `.xlsm` (`DS_Vanning Date.xlsm`) at the root holds the original macro-driven vanning-date decoder. Don't edit it; the Python script is now authoritative.
- The `ORIGIN_MAP` is **independent of** the FIFO-Pick scripts' decoder — they each maintain their own copy. Cross-check when prefix tables change.
- `desktop.ini` is a Windows artifact — ignore.
- TODO: confirm the report is delivered to the OEM client by email or by drop-folder, and on what cadence — currently undocumented in this folder.