"""
IP Sorting KPI Automation
Combines 502 data, runs KPI calculations, validates data, opens folders for upload.

Usage:
  python run_kpi.py                        # Auto-detect last week + Saturday
  python run_kpi.py --week-of 2026-03-30   # Target specific week (Monday date)
  python run_kpi.py --saturday             # Force-include Saturday even without data file
"""

import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import numpy as np

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(r'./IP Sorting')
DATA_DIR = BASE_DIR / 'Data'
DIR_502  = DATA_DIR / '502'
WORK_DIR = BASE_DIR / 'Work3'
HTML_DIR = BASE_DIR / 'Html'
RTI_PATH = DATA_DIR / 'RTI.xlsx'
DUMMY_502 = DIR_502 / '502.csv'
KPI_SCRIPT = BASE_DIR / 'Old' / 'IP_Sorting_KPI_20260315ver8.py'

# SharePoint BI Data folder synced locally via OneDrive
SHAREPOINT_LOCAL = Path(
    r'./'
    r'Shared BI/bi_reports/IP_Sorting/BI Data'
)

UPLOAD_FILES = [
    'df_deli2.csv',
    'df_deli5.csv',
    'df_deli_RTI_progress1.csv',
    'df_deli_RTI_progress_list.csv',
    'df_deli_carton_progress1.csv',
    'df_deli_carton_progress_list.csv',
    'df_deli_D0_progress1.csv',
    'df_deli_D0_progress_list.csv',
    'df_sort5.csv',
    'df_sort_progress1.csv',
    'df_sort_RTI_progress1.csv',
    'df_sort_RTI_progress_list.csv',
]


# ── Date Calculation ─────────────────────────────────────────────────────────

def find_last_monday():
    """Return the Monday of the previous completed work week."""
    today = date.today()
    current_monday = today - timedelta(days=today.weekday())
    return current_monday - timedelta(days=7)  # Always go back one full week


def calc_target_dates(monday, include_saturday=False):
    """Calculate all 6 date variables used by the main KPI script."""
    workdays = [monday + timedelta(days=i) for i in range(5)]
    if include_saturday:
        workdays.append(monday + timedelta(days=5))

    day_after_last = workdays[-1] + timedelta(days=1)

    return {
        'target_date':      [d.strftime('%Y%m%d') for d in workdays],
        'target_date2':     [d.strftime('%Y%m%d') for d in workdays] + ['0000'],
        'target_date_from': int(monday.strftime('%Y%m%d') + '0501'),
        'target_date_to':   int(day_after_last.strftime('%Y%m%d') + '0808'),
        'html_date':        monday.strftime('%m/%d/%Y'),
        'html_file':        monday.strftime('%Y%m%d') + '.html',
    }


def check_saturday_file(monday):
    """Return True if a MODULE_LOC file exists for Saturday of this week."""
    saturday = monday + timedelta(days=5)
    pattern = str(DIR_502 / f"MODULE_LOC_SITE2_CUST1_{saturday.strftime('%Y%m%d')}_*.csv")
    return len(glob.glob(pattern)) > 0, saturday


# ── 502 Combiner ─────────────────────────────────────────────────────────────

def run_502_combiner():
    """Combine all 502 CSVs, deduplicate, merge with RTI. Outputs Work3/df_502.csv."""
    print('\n--- Running 502 Combiner ---')

    csv_files = glob.glob(str(DIR_502 / '*.csv'))
    print(f'  Found {len(csv_files)} CSV files in Data/502/')

    data_list = []
    for f in csv_files:
        data_list.append(pd.read_csv(f))

    df = pd.concat(data_list, axis=0, sort=True)
    print(f'  Total rows after concat: {df.shape[0]:,}')

    df = df[['MODULE#', 'COMM PRODUCT', 'ARRIVAL DATE']]
    df = df.rename(columns={'MODULE#': 'MODULE'})
    df = df.drop_duplicates()

    # Clean COMM PRODUCT: remove spaces and hyphens
    df['COMM PRODUCT'] = df['COMM PRODUCT'].astype(str).str.replace(' ', '', regex=False).str.replace('-', '', regex=False)
    df['COMM PRODUCT'] = df['COMM PRODUCT'].replace('nan', pd.NA)

    print(f'  Rows after dedup: {df.shape[0]:,}')

    # Merge with RTI lookup
    df_rti = pd.read_excel(str(RTI_PATH), sheet_name='RTI')
    df = pd.merge(df, df_rti, on='COMM PRODUCT', how='left')

    output = WORK_DIR / 'df_502.csv'
    df.to_csv(str(output))
    print(f'  Saved {output.name}')


# ── Main KPI Script ──────────────────────────────────────────────────────────

def run_kpi_main(dates):
    """Run the main KPI script with date variables replaced."""
    print('\n--- Running Main KPI Script ---')

    source = KPI_SCRIPT.read_text(encoding='utf-8')
    lines = source.split('\n')
    new_lines = []

    # Patterns to match the 6 date variable lines (lines 18-25 of the original)
    replacements = [
        (re.compile(r"^TargetDate = \["),
         f"TargetDate = {dates['target_date']}"),

        (re.compile(r"^TargetDate2 = \["),
         f"TargetDate2 = {dates['target_date2']}"),

        (re.compile(r"^TargetDateFrom = \d"),
         f"TargetDateFrom = {dates['target_date_from']}"),

        (re.compile(r"^TargetDateTo\s+= \d"),
         f"TargetDateTo   = {dates['target_date_to']}"),

        (re.compile(r"^sHtmlDate = '"),
         f"sHtmlDate = '{dates['html_date']}'"),

        (re.compile(r"^sHtmlFile = '"),
         f"sHtmlFile = '{dates['html_file']}'"),
    ]

    for line in lines:
        stripped = line.strip()
        matched = False
        for pattern, replacement in replacements:
            if pattern.match(stripped):
                new_lines.append(replacement)
                matched = True
                break
        if not matched:
            new_lines.append(line)

    modified = '\n'.join(new_lines)

    # Write temp file and run as subprocess (clean Python state, no exec issues)
    temp_path = BASE_DIR / '_temp_kpi_run.py'
    temp_path.write_text(modified, encoding='utf-8')

    try:
        subprocess.run([sys.executable, str(temp_path)], check=True)
    finally:
        temp_path.unlink(missing_ok=True)

    print('  Main KPI script completed.')


# ── Validation ───────────────────────────────────────────────────────────────

def auto_fix_missing_modules():
    """Check df_no_RTI_data1.csv. If it has data, add missing modules to 502.csv.
    Returns True if fixes were made (caller should re-run)."""
    no_rti_path = WORK_DIR / 'df_no_RTI_data1.csv'
    df = pd.read_csv(str(no_rti_path), index_col=0)

    if df.shape[0] == 0:
        print('  df_no_RTI_data1.csv is empty -- no missing modules.')
        return False

    print(f'\n  WARNING: {df.shape[0]} module(s) in Replenishment data missing from 502 files.')
    print('  Auto-adding to Data/502/502.csv ...')

    new_rows = []
    for _, row in df.iterrows():
        # Use PRODUCT CODE if available, fall back to COMM PRODUCT
        product = row.get('PRODUCT CODE', '')
        if pd.isna(product) or str(product).strip() == '':
            product = row.get('COMM PRODUCT', '')

        new_rows.append({
            'LOCATION':      row.get('LOCATION', ''),
            'MODULE#':       row.get('MODULE', ''),
            'PRODUCT':       product,
            'COMM PRODUCT':  row.get('COMM PRODUCT', ''),
            'PILOT': '', 'QUANTITY': '', 'ETA': '', 'DAMAGE': '',
            'CONTAINER': '', 'UNITLOAD#': '', 'ARRIVAL DATE': '',
            'ORDER NO': '', 'ORDER LINE NO': '',
        })

    df_new = pd.DataFrame(new_rows)

    # Append to existing 502.csv, deduplicate
    df_existing = pd.read_csv(str(DUMMY_502))
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined = df_combined.drop_duplicates(subset=['MODULE#', 'COMM PRODUCT'])
    df_combined.to_csv(str(DUMMY_502), index=False)

    added = df_combined.shape[0] - df_existing.shape[0]
    print(f'  Added {added} new entries to 502.csv (total: {df_combined.shape[0]:,} rows)')
    return added > 0


def check_blank_rti():
    """Check df_502.csv for blank RTI values.
    If found, prompt for Packing Code per product and auto-update RTI.xlsx.
    Returns True if updates were made (caller should re-run)."""
    df = pd.read_csv(str(WORK_DIR / 'df_502.csv'), index_col=0)

    blank_mask = df['RTI'].isna()
    if not blank_mask.any():
        print('  No blank RTI values -- all good.')
        return False

    missing = df.loc[blank_mask, ['COMM PRODUCT']].drop_duplicates()
    missing = missing.dropna(subset=['COMM PRODUCT'])

    if missing.empty:
        print('  No blank RTI values -- all good.')
        return False

    # Load RTI_Code lookup: Packing Codes in this list are RTI=1, all others RTI=0
    df_rti_code = pd.read_excel(str(RTI_PATH), sheet_name='RTI_Code')
    rti_codes = set(df_rti_code['Type'].astype(str).str.upper())

    print(f'\n{"=" * 60}')
    print(f'  {len(missing)} COMM PRODUCT(s) need RTI mapping.')
    print(f'  RTI packing codes (=1): {", ".join(sorted(rti_codes))}')
    print(f'  Any other code = Carton (RTI=0)')
    print(f'  Dock options: S3 (default), D0, C0, S0, S1')
    print()

    new_entries = []
    for _, row in missing.iterrows():
        comm = row['COMM PRODUCT']
        packing = input(f'    {comm}  Packing Code: ').strip().upper()
        if not packing:
            print('      Skipped.')
            continue

        rti_val = 1 if packing in rti_codes else 0
        rti_label = 'RTI' if rti_val == 1 else 'Carton'

        dock = input(f'      Dock [S3]: ').strip().upper() or 'S3'

        new_entries.append({
            'COMM PRODUCT': comm,
            'Packing Code': packing,
            'RTI': rti_val,
            'Dock': dock,
        })
        print(f'      -> {rti_label} (RTI={rti_val}), Dock={dock}')

    if not new_entries:
        print('\n  No entries added. Please provide at least one mapping.')
        return False

    # Read all existing sheets
    df_rti = pd.read_excel(str(RTI_PATH), sheet_name='RTI')
    df_dock = pd.read_excel(str(RTI_PATH), sheet_name='DockCode')
    df_rti_code = pd.read_excel(str(RTI_PATH), sheet_name='RTI_Code')

    # Append to RTI sheet
    df_new = pd.DataFrame(new_entries)
    df_rti = pd.concat([df_rti, df_new], ignore_index=True)

    # Append to DockCode sheet
    df_dock_new = df_new[['COMM PRODUCT', 'Dock']].rename(
        columns={'COMM PRODUCT': 'PRODUCT CODE', 'Dock': 'DOCK'}
    )
    df_dock = pd.concat([df_dock, df_dock_new], ignore_index=True)

    # Write back all sheets
    with pd.ExcelWriter(str(RTI_PATH), engine='openpyxl', mode='w') as writer:
        df_rti.to_excel(writer, sheet_name='RTI', index=False)
        df_dock.to_excel(writer, sheet_name='DockCode', index=False)
        df_rti_code.to_excel(writer, sheet_name='RTI_Code', index=False)

    print(f'\n  Updated RTI.xlsx with {len(new_entries)} new entries.')
    return True


# ── Copy to SharePoint ───────────────────────────────────────────────────────

def copy_to_sharepoint():
    """Copy output files from Work3 to the SharePoint-synced BI Data folder."""
    print('\n--- Copying to SharePoint (BI Data) ---')

    if not SHAREPOINT_LOCAL.exists():
        print(f'  ERROR: SharePoint sync folder not found at:')
        print(f'    {SHAREPOINT_LOCAL}')
        print(f'  Please sync the Teams folder via OneDrive first.')
        print(f'  (In Teams > Files tab > click "Sync")')
        return

    copied = 0
    for fname in UPLOAD_FILES:
        src = WORK_DIR / fname
        dst = SHAREPOINT_LOCAL / fname
        if src.exists():
            shutil.copy2(str(src), str(dst))
            size_kb = src.stat().st_size / 1024
            print(f'    {fname:45s} ({size_kb:,.1f} KB)')
            copied += 1
        else:
            print(f'    {fname:45s} ** MISSING - skipped **')

    print(f'\n  Copied {copied}/{len(UPLOAD_FILES)} files to SharePoint BI Data.')
    print(f'  OneDrive will sync them automatically.')


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()

    # Determine Monday of target week
    if args.week_of:
        monday = date.fromisoformat(args.week_of)
        if monday.weekday() != 0:
            print(f'  Note: {args.week_of} is not a Monday. Adjusting to Monday of that week.')
            monday -= timedelta(days=monday.weekday())
    else:
        monday = find_last_monday()

    # Saturday is opt-in only (~95% of weeks are Mon-Fri).
    # If Saturday data is present but --saturday wasn't passed, warn so it's not missed silently.
    sat_exists, saturday = check_saturday_file(monday)
    include_saturday = args.saturday
    if sat_exists and not args.saturday:
        print(f'  NOTE: Saturday 502 file found for {saturday} but --saturday flag not set.')
        print(f'        Saturday will be EXCLUDED. Re-run with --saturday to include it.')
    elif args.saturday and not sat_exists:
        print(f'  WARNING: --saturday set but no 502 file found for {saturday}.')

    # Calculate dates
    dates = calc_target_dates(monday, include_saturday)

    # Print summary
    print(f'\n{"=" * 60}')
    print(f'  IP Sorting KPI -- Week of {dates["html_date"]}')
    print(f'{"=" * 60}')
    print(f'  Workdays:  {", ".join(dates["target_date"])}')
    print(f'  From:      {dates["target_date_from"]}')
    print(f'  To:        {dates["target_date_to"]}')
    print(f'  HTML file: {dates["html_file"]}')

    resp = input('\n  Proceed? [Y/n]: ')
    if resp.strip().lower() == 'n':
        print('  Aborted.')
        return

    # Validation loop
    MAX_ITER = 5
    for iteration in range(MAX_ITER):
        print(f'\n{"=" * 60}')
        print(f'  Iteration {iteration + 1}')
        print(f'{"=" * 60}')

        run_502_combiner()
        run_kpi_main(dates)

        if auto_fix_missing_modules():
            print('  Re-running with updated 502.csv ...')
            continue

        if check_blank_rti():
            print('  Re-running with updated RTI.xlsx ...')
            continue

        print(f'\n{"=" * 60}')
        print('  All validations passed!')
        print(f'{"=" * 60}')
        break
    else:
        print(f'\n  ERROR: Still have data issues after {MAX_ITER} iterations.')
        print('  Please investigate manually.')
        return

    # Copy to SharePoint
    copy_to_sharepoint()
    print('\n  Done!')


def parse_args():
    parser = argparse.ArgumentParser(description='IP Sorting KPI Automation')
    parser.add_argument('--saturday', action='store_true',
                        help='Include Saturday as a workday')
    parser.add_argument('--week-of', type=str, default=None,
                        help='Monday of target week (YYYY-MM-DD). Default: last week.')
    return parser.parse_args()


if __name__ == '__main__':
    main()
