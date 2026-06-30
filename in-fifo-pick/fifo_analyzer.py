import pandas as pd
import datetime as dt
import shutil
from pathlib import Path

PICK_INTERVAL_MIN = 15

class FIFOAnalyzer:

    # KA prefixes whose barcode does NOT encode a vanning date. The encoder
    # returns None for these (see datetable2_prefixes), so we fall back to ETA
    # and then shift backward by ETA_OFFSET_DAYS to approximate the real S DATE.
    # Calibration: warehouse-label sampling on 2026-05-11 showed ETA runs
    # ~5-72 days after the printed S DATE (median ~50). 50 minimizes mean
    # absolute error across the sample.
    ETA_OFFSET_PREFIXES = (
        'KA120', 'KA216', 'KA246', 'KA267', 'KA357', 'KA359', 'KA361', 'KA363',
        'KA365', 'KA367', 'KA369', 'KA371', 'KA373', 'KA374',
    )
    ETA_OFFSET_DAYS = 50

    def __init__(self, data_folder,
                 output_folder='./IN FIFO Pick/Power BI Data/',
                 teams_folder='./Shared/General/Data Analysis/FIFO KPI/Power BI Data/'):
        self.data_folder = Path(data_folder)
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(parents=True, exist_ok=True)
        self.teams_folder = Path(teams_folder) if teams_folder else None
        if self.teams_folder:
            self.teams_folder.mkdir(parents=True, exist_ok=True)
    
    def clean_part_number(self, part_str):
        if pd.isna(part_str):
            return ''
        return str(part_str).replace(' ', '').replace('-', '')

    def _needs_eta_offset(self, module):
        """True for KA prefixes whose barcode lacks a vanning encoding —
        these depend on ETA fallback and need the ETA→S DATE shift applied."""
        if not isinstance(module, str):
            return False
        return any(module.startswith(p) for p in self.ETA_OFFSET_PREFIXES)

    def extract_vanning_date(self, module_str):
        """
        Extract vanning/manufacturing date from module number using 502 encoding tables
        MATCHES THE EXCEL FORMULA LOGIC
        
        DateTable1 (Japan): 6th digit = month (M-X, A-L)
        DateTable2 (China/Canada/Mexico): 6th & 7th digits = month (01-12)  
        DateTable3 (Thailand): 6th digit = year (1=2021, 2=2022, etc.) - with 2-month offset
        """
        try:
            if pd.isna(module_str):
                return None
            module_str = str(module_str).strip().upper()
            
            if len(module_str) < 6:
                return None
            
            # MNA encoding: positions 4-6 = "MNA" (covers V1, VA, V2, VB, VS,
            # A1, A2, T1, TA, CC, AS, VC, N2 prefixed modules)
            # Position 7: year digit (2020 + digit)
            # Position 8: month letter (A-N, skipping I and L)
            # Positions 9-10: day of month
            if len(module_str) >= 11 and module_str[4:7] == 'MNA':
                mna_month_map = {
                    'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6,
                    'G': 7, 'H': 8, 'J': 9, 'K': 10, 'M': 11, 'N': 12
                }
                year_char = module_str[7]
                month_char = module_str[8]
                day_str = module_str[9:11]
                
                if year_char.isdigit() and month_char in mna_month_map and day_str.isdigit():
                    year = 2020 + int(year_char)
                    month = mna_month_map[month_char]
                    day = int(day_str)
                    if 1 <= day <= 31:
                        try:
                            return dt.date(year, month, day)
                        except ValueError:
                            return dt.date(year, month, 1)
                return None
            
            # Get current date for year logic
            today = dt.date.today()
            current_year = today.year
            current_month = today.month
            
            # Define prefix to DateTable mappings
            datetable1_prefixes = ['KJ699', 'KJ900', 'KJ999', '22200', '26700', '27100', '2Z400', '2G400', 'WN000']
            datetable2_prefixes = [
                'KJ550', 'KJ552', 'KJ563', 'KJ579', 'KJ598', 'KJ617', 'KJ621', 'KJ623', 'KJ624', 'KJ626',
                'KJ646', 'KJ656', 'KJ694', 'KJ698', 'KJ540', 'KA079', 'KA125', 'KA184', 'KA255', 'KA277',
                'KJ912',  # China
                # KA prefixes whose module numbers don't encode a vanning date in
                # positions 5-7 (the digits there are batch codes, not month). Listed
                # under DT2 so the decoder returns None for them (month digits never
                # parse as 01-12), which routes them to the ETA/ARRIVAL fallback.
                'KA120', 'KA216', 'KA246', 'KA267', 'KA357', 'KA359', 'KA361', 'KA363',
                'KA365', 'KA367', 'KA369', 'KA371', 'KA373', 'KA374',
            ]
            datetable3_prefixes = ['KJ911', 'KA085', 'KA152', 'KA158', 'KA199', 'KA224', 'KA261', 'KA296', 'KA331', 'KA118']
            
            # Determine which DateTable to use
            date_table = None
            origin = None
            
            # Check exact prefix matches
            for prefix in datetable1_prefixes:
                if module_str.startswith(prefix):
                    date_table = 1
                    origin = 'Japan'
                    break
            
            if date_table is None:
                for prefix in datetable2_prefixes:
                    if module_str.startswith(prefix):
                        date_table = 2
                        origin = 'Mexico/China/Canada'
                        break
            
            if date_table is None:
                for prefix in datetable3_prefixes:
                    if module_str.startswith(prefix):
                        date_table = 3
                        origin = 'Thailand'
                        break
            
            # Check pattern-based matches for DateTable2
            if date_table is None:
                if module_str.startswith('KJ') and len(module_str) >= 4:
                    code_num = module_str[2:4]
                    if code_num.isdigit():
                        num = int(code_num)
                        # KJ50-KJ69 = Canada/Mexico (DateTable2)
                        if 50 <= num <= 69:
                            date_table = 2
                            origin = 'Mexico/Canada'
            
            # Single letter prefixes
            if date_table is None:
                if module_str.startswith('S') and len(module_str) >= 1:
                    date_table = 1
                    origin = 'Japan'
                elif module_str.startswith('K') and not module_str.startswith('KJ') and not module_str.startswith('KA'):
                    date_table = 1
                    origin = 'Japan'
            
            # If still no match, default based on prefix
            if date_table is None:
                if module_str.startswith('KJ'):
                    date_table = 2
                    origin = 'Unknown'
                elif module_str.startswith('KA'):
                    date_table = 3
                    origin = 'Thailand'
                else:
                    return None
            
            # Extract date based on DateTable
            if date_table == 1:
                # DateTable1: 6th digit = month (M-X, A-L)
                if len(module_str) < 6:
                    return None
                
                month_char = module_str[5].upper()
                
                # M-X mapping (Jan-Dec)
                month_map_mx = {
                    'M': 1, 'N': 2, 'O': 3, 'P': 4, 'Q': 5, 'R': 6,
                    'S': 7, 'T': 8, 'U': 9, 'V': 10, 'W': 11, 'X': 12
                }
                # A-L mapping (Jan-Dec)
                month_map_al = {
                    'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'F': 6,
                    'G': 7, 'H': 8, 'I': 9, 'J': 10, 'K': 11, 'L': 12
                }
                
                month = month_map_mx.get(month_char) or month_map_al.get(month_char)
                
                if month:
                    # Year logic: if extracted month > current month, it's from last year
                    if month > current_month:
                        year = current_year - 1
                    else:
                        year = current_year
                    
                    return dt.date(year, month, 1)  # Use day 1 like Excel formula
                return None
            
            elif date_table == 2:
                # DateTable2: 6th & 7th digits = month (01-12)
                if len(module_str) < 7:
                    return None
                
                month_str = module_str[5:7]
                
                if month_str.isdigit():
                    month = int(month_str)
                    if 1 <= month <= 12:
                        # Year logic: if extracted month > current month, it's from last year
                        if month > current_month:
                            year = current_year - 1
                        else:
                            year = current_year
                        
                        return dt.date(year, month, 1)  # Use day 1 like Excel formula
                return None
            
            elif date_table == 3:
                # DateTable3: 6th digit = year (1=2021, 2=2022, etc.)
                # CRITICAL: Thailand dates need -2 month offset (EDATE in Excel)
                if len(module_str) < 8:
                    return None

                year_char = module_str[5]
                month_str = module_str[6:8]

                # Strict: both year digit and month digits must be valid 1-12.
                # No mid-year default — if the digits aren't a real month, return
                # None so the ETA/ARRIVAL fallback can supply the date.
                if not (year_char.isdigit() and month_str.isdigit()):
                    return None
                month = int(month_str)
                if not (1 <= month <= 12):
                    return None

                year = 2020 + int(year_char)  # 1=2021, 2=2022, etc.

                # Apply -2 month offset for Thailand (EDATE logic)
                if month <= 2:
                    return dt.date(year - 1, month + 10, 1)
                return dt.date(year, month - 2, 1)
            
            return None
            
        except Exception as e:
            print(f"Warning: failed to parse vanning date for '{module_str}': {e}")
            return None

    def find_picking_file(self):
        """
        Find the most recent picking file.
        The picking file has a dynamic name like: Picking_MODULE_LIST_SITE1_CUST1_YYYYMMDD_HHMMSS.csv
        """
        # Look for files matching the pattern
        pattern = 'Picking_MODULE_LIST_*.csv'
        picking_files = list(self.data_folder.glob(pattern))
        
        if not picking_files:
            # Fallback to the original static name
            fallback = self.data_folder / 'Picking_MODULE_LIST.csv'
            if fallback.exists():
                return fallback
            raise FileNotFoundError(
                f"No picking file found matching '{pattern}' or 'Picking_MODULE_LIST.csv' "
                f"in {self.data_folder}"
            )
        
        # Return the most recently modified file
        most_recent = max(picking_files, key=lambda p: p.stat().st_mtime)
        return most_recent

    def load_combined_inventory(self):
        """
        Load and combine ALL 502 inventory files from the 502/ subfolder.
        Deduplicates by MODULE#, keeping the earliest record so that modules
        picked mid-week (and removed from later snapshots) are still captured,
        while new arrivals from later snapshots are also included.
        Falls back to single MODULE_LOC.csv if no 502 folder exists.
        """
        inv_folder = self.data_folder / '502'
        inv_files = sorted(inv_folder.glob('MODULE_LOC_SITE1_CUST1_*.csv')) if inv_folder.exists() else []
        
        if inv_files:
            frames = []
            for f in inv_files:
                df = pd.read_csv(f)
                frames.append(df)
            df_inv = pd.concat(frames, ignore_index=True)
            df_inv = df_inv.drop_duplicates(subset=['MODULE#'], keep='first')
            print(f"Loaded {len(inv_files)} snapshots -> {len(df_inv):,} unique modules")
        else:
            # Fallback to single file
            inv_file = self.data_folder / 'MODULE_LOC.csv'
            print(f"No 502/ folder found, using {inv_file.name}")
            df_inv = pd.read_csv(inv_file)

        return df_inv

    def load_and_merge_data(self):
        """
        Load picking and inventory
        Extract VANNING DATES from module numbers, with ETA fallback.
        Also builds self.module_arrival_lookup (module -> raw ARRIVAL DATE)
        for the parallel arrival-date FIFO calc.
        """
        # Load picking data
        pick_file = self.find_picking_file()
        df_pick = pd.read_csv(pick_file)
        df_pick['PROCESS TIME'] = pd.to_numeric(df_pick['PROCESS TIME'], errors='coerce')
        invalid_time = df_pick['PROCESS TIME'].isna().sum()
        if invalid_time:
            print(f"Warning: dropping {invalid_time:,} picks with missing/invalid PROCESS TIME")
            df_pick = df_pick.dropna(subset=['PROCESS TIME']).copy()
        df_pick['PROCESS TIME'] = df_pick['PROCESS TIME'].astype('int64')

        # Clean part numbers
        df_pick['PRODUCT CODE'] = df_pick['PRODUCT CODE'].apply(self.clean_part_number)

        # Load operator mapping
        try:
            operator_file = Path('./config/scanner_barcode.xlsx')
            df_operators = pd.read_excel(operator_file, usecols='C:D', dtype=str)

            if len(df_operators.columns) >= 2:
                df_operators = df_operators.iloc[:, :2].copy()
                df_operators.columns = ['Code', 'Assigned to']

            def normalize_code(series):
                return (
                    series.fillna('')
                    .astype(str)
                    .str.replace(r'\s+', '', regex=True)
                    .str.upper()
                    .str.replace(r'\.0$', '', regex=True)
                )

            df_operators['Code'] = normalize_code(df_operators['Code'])
            df_operators['Assigned to'] = df_operators['Assigned to'].fillna('').astype(str).str.strip()
            df_operators = df_operators[df_operators['Code'].ne('') & df_operators['Code'].ne('NAN')]

            df_pick['OPCD'] = normalize_code(df_pick['OPCD'])
            operator_map = dict(zip(df_operators['Code'], df_operators['Assigned to']))
            df_pick['Operator_Name'] = df_pick['OPCD'].map(operator_map)
            df_pick['Operator_Name'] = df_pick['Operator_Name'].replace('', pd.NA).fillna('Unknown')
        except Exception as e:
            print(f"Warning: Could not load operator mapping: {e}")
            df_pick['Operator_Name'] = 'Unknown'

        df_pick = df_pick.sort_values('PROCESS TIME').reset_index(drop=True)
        df_pick = df_pick.drop_duplicates(subset=['MODULE NO'], keep='first').reset_index(drop=True)

        # Load combined inventory from all 502 snapshots
        df_inv = self.load_combined_inventory()

        # Filter and clean
        df_inv = df_inv[df_inv['DAMAGE'] != 'Y'].copy()
        df_inv = df_inv[~df_inv['LOCATION'].str.contains('-HLD-', na=False)].copy()
        df_inv = df_inv[~df_inv['LOCATION'].str.contains('EC-QPC', na=False)].copy()
        df_inv['PRODUCT'] = df_inv['PRODUCT'].apply(self.clean_part_number)

        # Extract VANNING DATES: module number encoding first, ETA fallback
        df_inv['VANNING DATE'] = df_inv['MODULE#'].apply(self.extract_vanning_date)

        # Capture raw ARRIVAL DATE (independent of the vanning fallback chain)
        # for the parallel arrival-date FIFO calculation.
        if 'ARRIVAL DATE' in df_inv.columns:
            df_inv['ARRIVAL_RAW'] = pd.to_datetime(df_inv['ARRIVAL DATE'], errors='coerce').dt.date
        else:
            df_inv['ARRIVAL_RAW'] = pd.NaT
        self.module_arrival_lookup = {
            mod: d for mod, d in zip(df_inv['MODULE#'], df_inv['ARRIVAL_RAW'])
            if pd.notna(d)
        }

        # Fallback chain for modules where the encoder returned None:
        #   1) ETA column   2) ARRIVAL DATE column
        for col in ('ETA', 'ARRIVAL DATE'):
            if col not in df_inv.columns:
                continue
            missing_mask = df_inv['VANNING DATE'].isna()
            if not missing_mask.any():
                break
            fallback_dates = pd.to_datetime(df_inv.loc[missing_mask, col], errors='coerce').dt.date
            valid = fallback_dates.notna()
            target_mask = missing_mask & valid.reindex(df_inv.index, fill_value=False)
            df_inv.loc[target_mask, 'VANNING DATE'] = fallback_dates[valid]

        # KA-catchall prefixes encode no vanning date; their VANNING DATE came
        # from ETA/ARRIVAL, which lags real S DATE by ~50 days. Shift back so
        # the date stored matches what's printed on the warehouse label.
        ka_mask = df_inv['MODULE#'].apply(self._needs_eta_offset)
        shift_mask = ka_mask & df_inv['VANNING DATE'].notna()
        if shift_mask.any():
            offset = dt.timedelta(days=self.ETA_OFFSET_DAYS)
            df_inv.loc[shift_mask, 'VANNING DATE'] = df_inv.loc[shift_mask, 'VANNING DATE'].apply(
                lambda d: d - offset
            )

        vanning_dated = df_inv['VANNING DATE'].notna().sum()
        print(f"Inventory: {len(df_inv):,} modules, {vanning_dated:,} dated ({vanning_dated/len(df_inv)*100:.1f}%)")

        # Create lookups for vanning dates and ETA/ARRIVAL fallback.
        # inv_date_lookup is keyed by module and prefers ETA, then ARRIVAL DATE,
        # so picked modules that the encoder can't date still get a real date.
        inv_vanning_lookup = dict(zip(df_inv['MODULE#'], df_inv['VANNING DATE']))
        inv_date_lookup = {}
        offset = dt.timedelta(days=self.ETA_OFFSET_DAYS)
        for col in ('ETA', 'ARRIVAL DATE'):
            if col not in df_inv.columns:
                continue
            for mod, raw in zip(df_inv['MODULE#'], df_inv[col]):
                if mod in inv_date_lookup or pd.isna(raw):
                    continue
                try:
                    parsed = pd.to_datetime(raw).date()
                    if self._needs_eta_offset(mod):
                        parsed = parsed - offset
                    inv_date_lookup[mod] = parsed
                except (ValueError, TypeError):
                    pass

        # Add picked modules to inventory with VANNING DATES
        picked_modules_data = []
        for _, row in df_pick.iterrows():
            module = row['MODULE NO']
            vanning_date = inv_vanning_lookup.get(module)
            if vanning_date is None or pd.isna(vanning_date):
                vanning_date = self.extract_vanning_date(module)
                if vanning_date is None:
                    vanning_date = inv_date_lookup.get(module)
            picked_modules_data.append({
                'MODULE#': module,
                'PRODUCT': row['PRODUCT CODE'],
                'VANNING DATE': vanning_date,
                'LOCATION': row.get('PICK LOCATION', ''),
                'DAMAGE': 'N'
            })

        df_picked_modules = pd.DataFrame(picked_modules_data)

        # Merge: combined inventory + picked modules
        df_inv_combined = pd.concat([df_inv[['MODULE#', 'PRODUCT', 'VANNING DATE', 'LOCATION', 'DAMAGE']],
                                    df_picked_modules], ignore_index=True)
        df_inv_combined = df_inv_combined.drop_duplicates(subset=['MODULE#'], keep='first')
        df_inv_combined['tPickTime'] = pd.Series(pd.NA, index=df_inv_combined.index, dtype='Int64')

        print(f"Picks: {len(df_pick):,} unique scans")
        return df_pick, df_inv_combined
    
    def _enrich_snapshot_dates(self, df):
        """Apply encoder + ETA/ARRIVAL fallback + KA offset to a raw 502 snapshot.
        Returns the dataframe filtered to non-damaged, valid-location rows with
        PRODUCT cleaned, VANNING DATE populated, and ARRIVAL_PARSED populated
        with the raw 'ARRIVAL DATE' column parsed to date (no fallbacks).
        Does NOT filter to dated rows — callers need to see undated modules
        to distinguish 'part absent' from 'part present but undated'."""
        df = df[df['DAMAGE'] != 'Y']
        df = df[~df['LOCATION'].str.contains('-HLD-', na=False)]
        df = df[~df['LOCATION'].str.contains('EC-QPC', na=False)]
        df = df.copy()
        df['PRODUCT'] = df['PRODUCT'].apply(self.clean_part_number)
        df['VANNING DATE'] = df['MODULE#'].apply(self.extract_vanning_date)

        # Parse raw ARRIVAL DATE separately — used for arrival-date FIFO,
        # independent of the vanning fallback chain below.
        if 'ARRIVAL DATE' in df.columns:
            df['ARRIVAL_PARSED'] = pd.to_datetime(df['ARRIVAL DATE'], errors='coerce').dt.date
        else:
            df['ARRIVAL_PARSED'] = pd.NaT

        for col in ('ETA', 'ARRIVAL DATE'):
            if col not in df.columns:
                continue
            missing = df['VANNING DATE'].isna()
            if not missing.any():
                break
            vals = pd.to_datetime(df.loc[missing, col], errors='coerce').dt.date
            ok = vals.notna()
            target = missing & ok.reindex(df.index, fill_value=False)
            df.loc[target, 'VANNING DATE'] = vals[ok]

        ka_mask = df['MODULE#'].apply(self._needs_eta_offset)
        shift = ka_mask & df['VANNING DATE'].notna()
        if shift.any():
            off = dt.timedelta(days=self.ETA_OFFSET_DAYS)
            df.loc[shift, 'VANNING DATE'] = df.loc[shift, 'VANNING DATE'].apply(lambda d: d - off)

        return df

    def _build_eod_floor_per_day(self, pick_days):
        """For each unique pick day, return per-part 'oldest still-on-shelf at
        end of day' lookups by BOTH vanning date and arrival date, plus the
        set of parts present in the EOD snapshot at all.

        Uses the earliest 502/MODULE_LOC_*.csv snapshot on (pick_day + 1) as the
        EOD proxy. If no snapshot at-or-after that target exists (e.g., for the
        latest day), uses the latest available snapshot.

        The parts-present set is needed to distinguish 'part absent from EOD'
        (truly trivially FIFO) from 'part present but undated' (unscorable).

        Returns:
            vanning_floor: dict {(pick_day, part_clean): date}
            arrival_floor: dict {(pick_day, part_clean): date}
            parts_present: dict {pick_day: set(part_clean)}
            approximate_days: set of pick_days where no day+1 snapshot existed
        """
        import re
        inv_folder = self.data_folder / '502'
        vanning_floor = {}
        arrival_floor = {}
        parts_present = {}
        approximate_days = set()
        if not inv_folder.exists():
            return vanning_floor, arrival_floor, parts_present, set(pick_days)

        pattern = re.compile(r'MODULE_LOC_SITE1_CUST1_(\d{8})_(\d{6})\.csv')
        snaps_by_date = {}  # snap_date -> (time_str, path)
        for f in sorted(inv_folder.glob('MODULE_LOC_SITE1_CUST1_*.csv')):
            m = pattern.search(f.name)
            if not m:
                continue
            date_str, time_str = m.groups()
            d = dt.datetime.strptime(date_str, '%Y%m%d').date()
            if d not in snaps_by_date or time_str < snaps_by_date[d][0]:
                snaps_by_date[d] = (time_str, f)

        if not snaps_by_date:
            return vanning_floor, arrival_floor, parts_present, set(pick_days)

        sorted_snap_dates = sorted(snaps_by_date)
        latest_snap_date = sorted_snap_dates[-1]

        # Map each pick day → snapshot date that approximates its EOD
        pick_to_snap = {}
        for pday in set(pick_days):
            target = pday + dt.timedelta(days=1)
            chosen = next((sd for sd in sorted_snap_dates if sd >= target), None)
            if chosen is None:
                chosen = latest_snap_date
                approximate_days.add(pday)
            pick_to_snap[pday] = chosen

        # Load each unique snapshot once → compute per-part oldest vanning AND
        # arrival, plus the set of parts that have ANY row in the snapshot.
        snap_floor_v = {}     # snap_date -> {part: vanning date}
        snap_floor_a = {}     # snap_date -> {part: arrival date}
        snap_present = {}     # snap_date -> set(part)
        for snap_date in set(pick_to_snap.values()):
            df = pd.read_csv(snaps_by_date[snap_date][1])
            df = self._enrich_snapshot_dates(df)
            if df.empty:
                snap_floor_v[snap_date] = {}
                snap_floor_a[snap_date] = {}
                snap_present[snap_date] = set()
                continue
            snap_present[snap_date] = set(df['PRODUCT'].dropna().unique())
            snap_floor_v[snap_date] = df.groupby('PRODUCT')['VANNING DATE'].min().to_dict()
            snap_floor_a[snap_date] = df.groupby('PRODUCT')['ARRIVAL_PARSED'].min().to_dict()

        # Compose final lookups
        for pday, snap_date in pick_to_snap.items():
            parts_present[pday] = snap_present[snap_date]
            for part, d in snap_floor_v[snap_date].items():
                if pd.notna(d):
                    vanning_floor[(pday, part)] = d
            for part, d in snap_floor_a[snap_date].items():
                if pd.notna(d):
                    arrival_floor[(pday, part)] = d

        return vanning_floor, arrival_floor, parts_present, approximate_days

    def check_fifo(self, df_pick, df_inv):
        """End-of-Day FIFO compliance with 5-day grace.

        Computes two parallel flags per pick:
          - fifo_compliant (vanning date — primary, exported to Power BI)
          - fifo_arrival_compliant (raw ARRIVAL DATE — diagnostic only)

        A pick of part P on operation-day D is FIFO iff the picked module's
        date is within 5 days of the oldest date of part P STILL ON THE SHELF
        at end of day D. Same-day picks that include both a newer and the
        oldest module both pass — by the time the next-morning snapshot is
        taken, the oldest is gone and the newer pick is compared only against
        whatever remains.
        """
        # Operation-date (shift-adjusted, matching what gets exported as OperationDate)
        op_dt = pd.to_datetime(df_pick['PROCESS TIME'].astype(str).str.zfill(12), format='%Y%m%d%H%M')
        op_date_only = op_dt.apply(
            lambda d: (d - pd.Timedelta(days=1)).date() if d.hour < 6 else d.date()
        )
        pick_days = op_date_only.tolist()

        vanning_floor, arrival_floor, parts_present, approximate_days = \
            self._build_eod_floor_per_day(pick_days)
        if approximate_days:
            print(f"  ({len(approximate_days)} pick-day(s) had no next-morning snapshot; latest used)")

        module_to_vanning = dict(zip(df_inv['MODULE#'], df_inv['VANNING DATE']))
        module_to_arrival = getattr(self, 'module_arrival_lookup', {})

        total = len(df_pick)
        oldest_out = [None] * total
        picked_out = [None] * total
        fifo_flags = [pd.NA] * total
        fifo_arr_flags = [pd.NA] * total

        columns = df_pick.columns
        idx_part = columns.get_loc('PRODUCT CODE')
        idx_module = columns.get_loc('MODULE NO')

        for idx, row in enumerate(df_pick.itertuples(index=False, name=None)):
            part = row[idx_part]
            module = row[idx_module]
            pday = pick_days[idx]
            present_set = parts_present.get(pday, set())
            part_present = part in present_set

            # Vanning-date compliance
            pdate = module_to_vanning.get(module)
            if not pd.isna(pdate):
                picked_out[idx] = pdate
                floor = vanning_floor.get((pday, part))
                if floor is not None:
                    oldest_out[idx] = floor
                    fifo_flags[idx] = 1 if (pdate - floor).days <= 5 else 0
                elif not part_present:
                    # Part fully cleared off the shelf by EOD → trivially FIFO
                    oldest_out[idx] = pdate
                    fifo_flags[idx] = 1
                # else: part present but no dated module → leave as NA (unscorable)

            # Arrival-date compliance (parallel, same presence rule)
            adate = module_to_arrival.get(module)
            if not (adate is None or pd.isna(adate)):
                a_floor = arrival_floor.get((pday, part))
                if a_floor is not None:
                    fifo_arr_flags[idx] = 1 if (adate - a_floor).days <= 5 else 0
                elif not part_present:
                    fifo_arr_flags[idx] = 1
                # else: part present but no arrival-dated module → leave as NA

        df_pick['oldest_vanning_date'] = oldest_out
        df_pick['picked_vanning_date'] = picked_out
        df_pick['fifo_compliant'] = pd.array(fifo_flags, dtype='Int64')
        df_pick['fifo_arrival_compliant'] = pd.array(fifo_arr_flags, dtype='Int64')
        # tPickTime is a legacy column; not used by exports under the new logic
        df_inv['tPickTime'] = pd.array([pd.NA] * len(df_inv), dtype='Int64')

        return df_pick
    
    def export_inventory_excel(self):
        """
        Export a simple inventory Excel file with columns:
        LOCATION, MODULE#, PRODUCT, S DATE, ETA, ARRIVAL DATE, QUANTITY, CONTAINER, UNITLOAD#
        S DATE = vanning date extracted from module number encoding
        """

        # Use only the latest 502 inventory snapshot
        inv_folder = self.data_folder / '502'
        inv_files = sorted(inv_folder.glob('MODULE_LOC_SITE1_CUST1_*.csv'))
        if not inv_files:
            print("No inventory files found in 502/ folder")
            return None
        latest_file = inv_files[-1]
        df = pd.read_csv(latest_file)

        df['S DATE'] = df['MODULE#'].apply(self.extract_vanning_date)

        cols = ['LOCATION', 'MODULE#', 'PRODUCT', 'S DATE', 'ETA', 'ARRIVAL DATE', 'QUANTITY', 'CONTAINER', 'UNITLOAD#']
        cols = [c for c in cols if c in df.columns]
        df_export = df[cols]

        output_file = self.output_folder / 'Current_Inventory_Overview.xlsx'
        df_export.to_excel(output_file, index=False)

        return output_file

    def export_oldest_location_file(self, df_inv):
        """
        Export the oldest location by part for Power BI lookup table.
        This helps warehouse operators find where the oldest module for each part is located.
        """
        # Filter to modules with valid vanning dates only
        df_with_dates = df_inv[df_inv['VANNING DATE'].notna()].copy()

        if len(df_with_dates) == 0:
            print("Warning: No modules with valid vanning dates found")
            oldest_export = pd.DataFrame(columns=['PartNumber', 'OldestLocation', 'OldestVanningDate', 'OldestModuleNumber'])
        else:
            oldest_by_part = df_with_dates.loc[
                df_with_dates.groupby('PRODUCT')['VANNING DATE'].idxmin()
            ].copy()
            oldest_export = pd.DataFrame({
                'PartNumber': oldest_by_part['PRODUCT'],
                'OldestLocation': oldest_by_part['LOCATION'],
                'OldestVanningDate': oldest_by_part['VANNING DATE'],
                'OldestModuleNumber': oldest_by_part['MODULE#']
            })
            oldest_export = oldest_export.sort_values('PartNumber').reset_index(drop=True)

        oldest_file = self.output_folder / 'PowerBI_OldestLocationByPart.csv'
        oldest_export.to_csv(oldest_file, index=False)
        return oldest_file

    def get_oldest_location_lookup(self, df_inv):
        """Create lookup dictionary for oldest location by part"""
        df_with_dates = df_inv[df_inv['VANNING DATE'].notna()].copy()
        if len(df_with_dates) == 0:
            return {}, {}
        oldest_by_part = df_with_dates.loc[
            df_with_dates.groupby('PRODUCT')['VANNING DATE'].idxmin()
        ].copy()
        location_lookup = dict(zip(oldest_by_part['PRODUCT'], oldest_by_part['LOCATION']))
        module_lookup = dict(zip(oldest_by_part['PRODUCT'], oldest_by_part['MODULE#']))
        return location_lookup, module_lookup

    def generate_report(self, df_pick):
        all_picks = len(df_pick)

        # Vanning-date rate (primary)
        dated_v = int(df_pick['fifo_compliant'].notna().sum())
        fifo_v = int(df_pick['fifo_compliant'].sum())
        rate_v = (fifo_v / dated_v * 100) if dated_v > 0 else 0

        # Arrival-date rate (diagnostic)
        dated_a = int(df_pick['fifo_arrival_compliant'].notna().sum())
        fifo_a = int(df_pick['fifo_arrival_compliant'].sum())
        rate_a = (fifo_a / dated_a * 100) if dated_a > 0 else 0

        print("\n" + "=" * 60)
        print("FIFO COMPLIANCE SUMMARY (End-of-Day, 5-day grace)")
        print("=" * 60)
        print(f"  Vanning date : {fifo_v:>6,}/{dated_v:>6,} = {rate_v:5.1f}%   ({all_picks - dated_v:,} excluded)")
        print(f"  Arrival date : {fifo_a:>6,}/{dated_a:>6,} = {rate_a:5.1f}%   ({all_picks - dated_a:,} excluded)")
        print("=" * 60)

        # Week-by-week (ISO week, Monday-start) — both rates side by side
        op_dates = pd.to_datetime(df_pick['PROCESS TIME'].astype(str).str.zfill(12), format='%Y%m%d%H%M').dt.date
        op_dates = pd.to_datetime(op_dates)
        week_start = (op_dates - pd.to_timedelta(op_dates.dt.weekday, unit='d')).dt.date
        df_pick['_week_start'] = week_start

        by_week = pd.DataFrame({
            'FIFO_V': df_pick.groupby('_week_start')['fifo_compliant'].sum(),
            'Total_V': df_pick.groupby('_week_start')['fifo_compliant'].count(),
            'FIFO_A': df_pick.groupby('_week_start')['fifo_arrival_compliant'].sum(),
            'Total_A': df_pick.groupby('_week_start')['fifo_arrival_compliant'].count(),
        })
        by_week['Van_%'] = (by_week['FIFO_V'] / by_week['Total_V'].replace(0, pd.NA) * 100).round(1)
        by_week['Arr_%'] = (by_week['FIFO_A'] / by_week['Total_A'].replace(0, pd.NA) * 100).round(1)

        print("\nBy Week (Mon-start):")
        print(f"  {'Week':<12} {'Van_picks':>10} {'Van_%':>7} {'Arr_picks':>10} {'Arr_%':>7}")
        for week, row in by_week.iterrows():
            van_pct = f"{row['Van_%']:.1f}" if pd.notna(row['Van_%']) else '   - '
            arr_pct = f"{row['Arr_%']:.1f}" if pd.notna(row['Arr_%']) else '   - '
            print(f"  {str(week):<12} {int(row['Total_V']):>10,} {van_pct:>7} "
                  f"{int(row['Total_A']):>10,} {arr_pct:>7}")
        df_pick.drop(columns=['_week_start'], inplace=True)

        # Export (clean up prior runs so the folder doesn't accumulate one file per run)
        for old in self.output_folder.glob('FIFO_Results_Vanning_*.csv'):
            try:
                old.unlink()
            except OSError as e:
                print(f"Warning: could not remove old results file {old.name}: {e}")

        timestamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = self.output_folder / f'FIFO_Results_Vanning_{timestamp}.csv'
        df_pick.to_csv(output_file, index=False)

        return rate_v
    
    def export_for_powerbi(self, df_pick, df_inv):
        """Export vanning date FIFO for Power BI"""
        
        # Get oldest location lookup
        location_lookup, module_lookup = self.get_oldest_location_lookup(df_inv)
        
        # Parse dates and shifts
        df_pick['Operation_Date'] = pd.to_datetime(df_pick['PROCESS TIME'].astype(str).str.zfill(12), format='%Y%m%d%H%M')
        df_pick['Hour_of_Day'] = df_pick['Operation_Date'].dt.hour
        # Shift 1 runs 06:00–16:30; anything outside that window is Shift 2.
        # Shift 2 picks before 06:00 belong to the *previous* operation date.
        df_pick['Operation_Date_Only'] = df_pick['Operation_Date'].apply(
            lambda dt: (dt - pd.Timedelta(days=1)).date() if dt.hour < 6 else dt.date()
        )

        def get_shift(dt_val):
            hour = dt_val.hour
            if 6 <= hour < 16 or (hour == 16 and dt_val.minute < 30):
                return 1
            return 2
        
        df_pick['Shift'] = df_pick['Operation_Date'].apply(get_shift)

        # Assign SessionId so PickDetails and OperatorTimeline share a bridge key
        self._assign_session_ids(df_pick)

        # 1. DETAILED PICK-LEVEL DATA
        detail_export = df_pick[[
            'PROCESS TIME', 'Operation_Date', 'Operation_Date_Only', 'Shift', 'Hour_of_Day',
            'OPCD', 'Operator_Name',
            'PRODUCT CODE', 'MODULE NO',
            'oldest_vanning_date', 'picked_vanning_date', 'fifo_compliant',
            'SessionId'
        ]].copy()

        detail_export = detail_export.rename(columns={
            'PROCESS TIME': 'ProcessTime',
            'Operation_Date': 'OperationDateTime',
            'Operation_Date_Only': 'OperationDate',
            'Shift': 'Shift',
            'Hour_of_Day': 'HourOfDay',
            'OPCD': 'OperatorCode',
            'Operator_Name': 'OperatorName',
            'PRODUCT CODE': 'PartNumber',
            'MODULE NO': 'ModuleNumber',
            'oldest_vanning_date': 'OldestAvailableDate',
            'picked_vanning_date': 'PickedModuleDate',
            'fifo_compliant': 'FIFO_Pick'
        })
        
        detail_file = self.output_folder / 'PowerBI_PickDetails.csv'
        detail_export.to_csv(detail_file, index=False)

        # 2. DAILY SUMMARY
        # sum/count on a nullable Int64 column skip NA, so Total_Picks here
        # counts only picks with a resolvable vanning date (the "compliance
        # denominator"). Column schema preserved for Power BI compatibility.
        daily_summary = df_pick.groupby('Operation_Date_Only').agg(
            FIFO_Picks=('fifo_compliant', 'sum'),
            Total_Picks=('fifo_compliant', 'count'),
        ).reset_index().rename(columns={'Operation_Date_Only': 'Date'})
        daily_summary['Non_FIFO_Picks'] = daily_summary['Total_Picks'] - daily_summary['FIFO_Picks']
        daily_summary['FIFO_Ratio'] = (daily_summary['FIFO_Picks'] / daily_summary['Total_Picks'].replace(0, pd.NA) * 100).round(1)

        daily_file = self.output_folder / 'PowerBI_DailySummary.csv'
        daily_summary.to_csv(daily_file, index=False)

        # 3. SHIFT SUMMARY
        shift_summary = df_pick.groupby(['Operation_Date_Only', 'Shift']).agg(
            FIFO_Picks=('fifo_compliant', 'sum'),
            Total_Picks=('fifo_compliant', 'count'),
        ).reset_index().rename(columns={'Operation_Date_Only': 'Date'})
        shift_summary['Non_FIFO_Picks'] = shift_summary['Total_Picks'] - shift_summary['FIFO_Picks']
        shift_summary['FIFO_Ratio'] = (shift_summary['FIFO_Picks'] / shift_summary['Total_Picks'].replace(0, pd.NA) * 100).round(1)

        shift_file = self.output_folder / 'PowerBI_ShiftSummary.csv'
        shift_summary.to_csv(shift_file, index=False)

        # 4. PART-LEVEL SUMMARY (WITH OLDEST LOCATION!)
        part_summary = df_pick.groupby('PRODUCT CODE').agg(
            FIFO_Picks=('fifo_compliant', 'sum'),
            Total_Picks=('fifo_compliant', 'count'),
        ).reset_index().rename(columns={'PRODUCT CODE': 'PartNumber'})
        part_summary['FIFO_Ratio'] = (part_summary['FIFO_Picks'] / part_summary['Total_Picks'].replace(0, pd.NA) * 100).round(1)
        
        # ADD OLDEST LOCATION COLUMNS
        part_summary['First_OldestLocation'] = part_summary['PartNumber'].map(location_lookup)
        part_summary['First_OldestModuleNumber'] = part_summary['PartNumber'].map(module_lookup)
        missing_module_mask = part_summary['First_OldestModuleNumber'].isna()
        if missing_module_mask.any():
            part_summary.loc[missing_module_mask, 'First_OldestModuleNumber'] = (
                'UNKNOWN_' + part_summary.loc[missing_module_mask, 'PartNumber'].astype(str)
            )
        
        part_summary = part_summary.sort_values('Total_Picks', ascending=False)
        
        part_file = self.output_folder / 'PowerBI_PartSummary.csv'
        part_summary.to_csv(part_file, index=False)

        # 5. OPERATOR-LEVEL SUMMARY
        operator_summary = df_pick.groupby(['Operator_Name', 'OPCD']).agg(
            FIFO_Picks=('fifo_compliant', 'sum'),
            Total_Picks=('fifo_compliant', 'count'),
        ).reset_index().rename(columns={'Operator_Name': 'OperatorName', 'OPCD': 'OperatorCode'})
        operator_summary['Non_FIFO_Picks'] = operator_summary['Total_Picks'] - operator_summary['FIFO_Picks']
        operator_summary['FIFO_Ratio'] = (operator_summary['FIFO_Picks'] / operator_summary['Total_Picks'].replace(0, pd.NA) * 100).round(1)
        operator_summary = operator_summary.sort_values('Total_Picks', ascending=False)

        operator_file = self.output_folder / 'PowerBI_OperatorSummary.csv'
        operator_summary.to_csv(operator_file, index=False)

        # 6. HOURLY SUMMARY
        hourly_summary = df_pick.groupby('Hour_of_Day').agg(
            FIFO_Picks=('fifo_compliant', 'sum'),
            Total_Picks=('fifo_compliant', 'count'),
        ).reset_index().rename(columns={'Hour_of_Day': 'Hour'})
        hourly_summary['Non_FIFO_Picks'] = hourly_summary['Total_Picks'] - hourly_summary['FIFO_Picks']
        hourly_summary['FIFO_Ratio'] = (hourly_summary['FIFO_Picks'] / hourly_summary['Total_Picks'].replace(0, pd.NA) * 100).round(1)

        hourly_file = self.output_folder / 'PowerBI_HourlySummary.csv'
        hourly_summary.to_csv(hourly_file, index=False)

        print(f"Power BI files exported -> {self.output_folder}")

    def _assign_session_ids(self, df_pick):
        """
        Compute a SessionId for each pick row in place. A new session starts
        when operator, shift, work-date, or FIFO state changes, or when the
        idle gap to the previous pick is >= PICK_INTERVAL_MIN. Sessions are
        the unit of the Operator Timeline visual. Writing SessionId back onto
        df_pick lets PickDetails and OperatorTimeline share a bridge key.
        Requires Operation_Date, Shift, Operation_Date_Only, fifo_compliant
        already set.
        """
        shift_label = df_pick['Shift'].map({1: 'shift1', 2: 'shift2'})
        order = df_pick.assign(_ShiftLabel=shift_label).sort_values(
            ['OPCD', 'Shift', 'Operation_Date'], kind='mergesort'
        ).index

        op = df_pick.loc[order, 'OPCD']
        sl = shift_label.loc[order]
        od = df_pick.loc[order, 'Operation_Date']
        odo = df_pick.loc[order, 'Operation_Date_Only']
        fc = df_pick.loc[order, 'fifo_compliant']

        gap_min = (od - od.shift()).dt.total_seconds() / 60
        new_session = (
            (gap_min >= PICK_INTERVAL_MIN)
            | (op != op.shift())
            | (sl != sl.shift())
            | (odo != odo.shift())
            | (fc != fc.shift())
        )
        new_session.iloc[0] = True
        session_ids = new_session.cumsum()

        df_pick['SessionId'] = session_ids.reindex(df_pick.index)

    def export_operator_timeline(self, df_pick):
        """
        Aggregate the per-pick SessionId (already on df_pick from
        _assign_session_ids) into one row per session for the Gantt visual.
        Every session is purely FIFO or purely Non-FIFO by construction, so
        FIFO_Status can drive bar color directly.
        """
        df = df_pick.dropna(subset=['OPCD', 'Operation_Date', 'SessionId']).copy()
        df['ShiftLabel'] = df['Shift'].map({1: 'shift1', 2: 'shift2'})

        grouped = df.groupby('SessionId', sort=False).agg(
            OP_ID=('OPCD', 'first'),
            OperatorName=('Operator_Name', 'first'),
            Shift1=('ShiftLabel', 'first'),
            OperationDate=('Operation_Date_Only', 'first'),
            tDate1=('Operation_Date', 'min'),
            tDate2=('Operation_Date', 'max'),
            pickCount=('OPCD', 'size'),
            FIFO_Compliant=('fifo_compliant', 'first'),
        ).reset_index()

        grouped['Timeline_Y'] = grouped['Shift1'] + '_' + grouped['OP_ID'].astype(str)
        grouped['FIFO_Status'] = grouped['FIFO_Compliant'].map({1: 'FIFO', 0: 'Non-FIFO'})
        grouped['tspend'] = (grouped['tDate2'] - grouped['tDate1']).dt.total_seconds() / 60
        grouped['ispeed'] = 0.0
        nonzero = (grouped['pickCount'] > 1) & (grouped['tspend'] > 0)
        grouped.loc[nonzero, 'ispeed'] = (
            grouped.loc[nonzero, 'tspend'] / grouped.loc[nonzero, 'pickCount'] * 60
        ).round(1)

        zero_width = grouped['tDate1'] == grouped['tDate2']
        grouped['tDate2_Display'] = grouped['tDate2']
        grouped.loc[zero_width, 'tDate2_Display'] = (
            grouped.loc[zero_width, 'tDate1'] + pd.Timedelta(seconds=30)
        )

        out = grouped[[
            'SessionId', 'OP_ID', 'OperatorName', 'Shift1', 'Timeline_Y',
            'OperationDate', 'tDate1', 'tDate2', 'tDate2_Display',
            'pickCount', 'FIFO_Status', 'tspend', 'ispeed',
        ]]

        timeline_file = self.output_folder / 'PowerBI_OperatorTimeline.csv'
        out.to_csv(timeline_file, index=False)

    def _sync_to_teams(self):
        """Copy all output files to the Teams shared drive folder."""
        if self.teams_folder is None:
            return
        copied = 0
        for src in self.output_folder.iterdir():
            if src.is_file():
                dst = self.teams_folder / src.name
                shutil.copy2(src, dst)
                copied += 1
        print(f"Synced {copied} file(s) to Teams")

def main():
    analyzer = FIFOAnalyzer(
        data_folder='./IN FIFO Pick/',
        output_folder='./IN FIFO Pick/Power BI Data/'
    )
    
    # Export simple inventory Excel overview
    analyzer.export_inventory_excel()

    # Load and merge data
    df_pick, df_inv = analyzer.load_and_merge_data()
    
    # Export oldest location lookup (for Power BI part lookup feature)
    analyzer.export_oldest_location_file(df_inv)
    
    # Calculate FIFO using vanning dates
    df_pick = analyzer.check_fifo(df_pick, df_inv)
    
    # Generate console report
    fifo_rate = analyzer.generate_report(df_pick)
    
    # Export for Power BI (now includes oldest location in PartSummary!)
    analyzer.export_for_powerbi(df_pick, df_inv)

    # Export operator session timeline for Gantt visual
    analyzer.export_operator_timeline(df_pick)

    # Mirror all output files to Teams shared drive
    analyzer._sync_to_teams()

if __name__ == "__main__":
    main()