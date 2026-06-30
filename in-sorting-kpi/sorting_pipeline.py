import pandas as pd
import numpy as np
import datetime as dt
from pathlib import Path
import json
import warnings
warnings.filterwarnings('ignore')

# Visualization libraries
try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False
    print("⚠️  Plotly not available. Install: pip install plotly")

# ML libraries
try:
    from sklearn.ensemble import RandomForestRegressor
    ML_AVAILABLE = True
except ImportError:
    ML_AVAILABLE = False
    print("⚠️  ML libraries not available. Install: pip install scikit-learn")

# ============================================================================
# CONFIGURATION
# ============================================================================
class Config:
    # Paths
    DATA_PATH = Path('./IN Sorting KPI/Data/')
    OUTPUT_PATH = Path('./IN Sorting KPI/Output')
    
    # File names
    FILE_810_SORTING = 'SORTING_MODULE_LIST.csv'
    FILE_PUTAWAY_CONTAINER = 'PUTAWAY_CONTAINER_LIST.csv'
    FILE_PUTAWAY_UNIT = 'PUTAWAY_UNIT_LIST.csv'
    FILE_102_MODULE = '102.csv'
    FILE_502_RECEIVING = '502.csv'
    FILE_OPERATOR_NAMES = './config/scanner_barcode.xlsx'
    
    # Output files
    OUTPUT_DAILY_METRICS = 'daily_metrics.xlsx'
    OUTPUT_RAW_DATA = 'processed_data.xlsx'
    OUTPUT_EMAIL_REPORT = 'Weekly Sorting and Putaway KPI.html'
    OUTPUT_PLOTLY_DASHBOARD = 'interactive_dashboard.html'
    OUTPUT_CONFIG = 'kpi_config.json'
    
    # Settings
    SORT_INTERVAL_MINUTES = 20
    WEEKLY_REPORT_DAYS = 5  # Last 5 business days

    # Shift definitions
    SHIFT1_START = dt.time(6, 0)
    SHIFT1_END = dt.time(16, 30)
    
    # Speed filtering
    MIN_SPEED_SECONDS = 5
    MAX_SPEED_SECONDS = 300
    
    # Benchmarking
    BENCHMARK_PERCENTILE = 0.1
    PUTAWAY_BENCHMARK_ADJUSTMENT = 1
    
    # Per-operation efficiency weights (speed + throughput must sum to 1.0).
    # Used as exponents in the weighted geometric mean inside
    # ``calculate_efficiency`` — higher SPEED_WEIGHT makes speed the
    # dominant factor (slow operators can't make up for it with volume).
    SORT_SPEED_WEIGHT = 0.20           # Sorting: speed-leaning
    SORT_THROUGHPUT_WEIGHT = 0.10
    PUTAWAY_SPEED_WEIGHT = 0.20        # Putaway: speed-leaning
    PUTAWAY_THROUGHPUT_WEIGHT = 0.10
    # Putaway difficulty bonus (0-15) – additive credit for physical demands.
    # Kept small so it can't push a peer-median operator into the high band on
    # its own.
    PUTAWAY_DIFFICULTY_BONUS = 3
    # Throughput curve steepness – higher = sharper rise toward the goal
    THROUGHPUT_CURVE_K = 1.0
    # Efficiency anchor: score given to an operator who matches peer median on
    # both axes. With anchor=60, peer-median performance lands at 60% — high
    # 80s/90s now genuinely require above-peer work on BOTH speed and volume,
    # not just being average.
    EFFICIENCY_MEDIAN_ANCHOR = 60
    # Per-component cap. Holds even outsized days to 100/component so a single
    # huge axis can't single-handedly lift the geometric mean above 100.
    EFFICIENCY_SCORE_CAP = 100
    # Slopes of the log curves applied to speed and volume:
    #     score = anchor × (1 + slope × ln(actual / peer_median))
    # 0.5 means a 2× peer-median axis adds ~anchor×0.35 ≈ +21 points; 3× adds
    # ~anchor×0.55 ≈ +33. Reaching 100/component now needs roughly e^(1/slope)
    # = ~2.7× the peer median on that axis, so high scores reflect genuine
    # outperformance rather than average days.
    SPEED_LOG_SLOPE = 0.5
    VOLUME_LOG_SLOPE = 0.5
    # Fallback daily-volume normalisation base used only when peer-median is
    # unavailable (very small datasets, <3 operator-days). Not a target —
    # purely a numerical scale for the throughput sigmoid.
    FALLBACK_DAILY_VOLUME = 50
    # Fallback speed percentiles used when dataset is too small (<5 records)
    FALLBACK_SORT_SPEED_P25 = 60
    FALLBACK_SORT_SPEED_P75 = 110
    FALLBACK_PUTAWAY_SPEED_P25 = 22
    FALLBACK_PUTAWAY_SPEED_P75 = 55

    # ---- Adaptive accuracy settings ----
    # Normalise throughput by hours worked vs peer-median hourly rate.
    # Makes part-shift operators comparable to full-shift peers.
    USE_HOURS_NORMALISED_THROUGHPUT = True
    # Bayesian shrinkage toward team-median score, weight = n/(n+k).
    # Small-sample days get pulled toward the median; 0 disables shrinkage.
    SHRINKAGE_K = 15
    # Use median of session speeds (not mean) when rolling up daily.
    # Median is robust to session start-up / wind-down outliers.
    USE_MEDIAN_SPEED = True
    # Minimum p75-p25 spread (seconds) before the speed sigmoid is used.
    # Prevents a converged team from producing cliff-edge scoring.
    MIN_SPEED_SPREAD_SORT = 10
    MIN_SPEED_SPREAD_PUTAWAY = 5
    # Fallback peer hourly rates (modules/hour) used when dataset too small.
    FALLBACK_SORT_HOURLY_RATE = 30.0
    FALLBACK_PUTAWAY_HOURLY_RATE = 25.0

    @classmethod
    def load_config(cls):
        """Load configuration from JSON file if exists"""
        config_file = cls.OUTPUT_PATH / cls.OUTPUT_CONFIG
        if config_file.exists():
            import json
            with open(config_file, 'r') as f:
                config_data = json.load(f)
                for key, value in config_data.items():
                    if hasattr(cls, key):
                        setattr(cls, key, value)
            print(f"   ✅ Loaded configuration from {config_file}")
    
    @classmethod
    def save_config(cls):
        """Save current configuration to JSON file"""
        import json
        config_data = {
            'SORT_INTERVAL_MINUTES': cls.SORT_INTERVAL_MINUTES,
            'WEEKLY_REPORT_DAYS': cls.WEEKLY_REPORT_DAYS,
            'MIN_SPEED_SECONDS': cls.MIN_SPEED_SECONDS,
            'MAX_SPEED_SECONDS': cls.MAX_SPEED_SECONDS,
            'BENCHMARK_PERCENTILE': cls.BENCHMARK_PERCENTILE,
            'SORT_SPEED_WEIGHT': cls.SORT_SPEED_WEIGHT,
            'SORT_THROUGHPUT_WEIGHT': cls.SORT_THROUGHPUT_WEIGHT,
            'PUTAWAY_SPEED_WEIGHT': cls.PUTAWAY_SPEED_WEIGHT,
            'PUTAWAY_THROUGHPUT_WEIGHT': cls.PUTAWAY_THROUGHPUT_WEIGHT,
            'PUTAWAY_DIFFICULTY_BONUS': cls.PUTAWAY_DIFFICULTY_BONUS,
            'THROUGHPUT_CURVE_K': cls.THROUGHPUT_CURVE_K,
        }
        config_file = cls.OUTPUT_PATH / cls.OUTPUT_CONFIG
        with open(config_file, 'w') as f:
            json.dump(config_data, f, indent=2)
        print(f"   ✅ Saved configuration to {config_file}")

def configure_system_interactive():
    """Interactive configuration panel"""
    print("\n" + "=" * 70)
    print("⚙️  KPI SYSTEM CONFIGURATION PANEL")
    print("=" * 70)
    
    print(f"\n📋 Current Settings:")
    print(f"   1. Sort Interval (minutes):     {Config.SORT_INTERVAL_MINUTES}")
    print(f"   2. Report Days (business days):  {Config.WEEKLY_REPORT_DAYS}")
    print(f"   3. Min Speed Filter (seconds):   {Config.MIN_SPEED_SECONDS}")
    print(f"   4. Max Speed Filter (seconds):   {Config.MAX_SPEED_SECONDS}")
    print(f"   5. Benchmark Percentile:         {Config.BENCHMARK_PERCENTILE:.0%} (top performers)")
    print(f"   6. Shift 1 Start Time:           {Config.SHIFT1_START}")
    print(f"   7. Shift 1 End Time:             {Config.SHIFT1_END}")

    print(f"\n🔧 Options:")
    print(f"   [1-7] Change a setting")
    print(f"   [S]   Save current configuration")
    print(f"   [R]   Reset to defaults")
    print(f"   [Q]   Quit configuration")
    
    while True:
        choice = input("\nEnter your choice: ").strip().upper()
        
        if choice == 'Q':
            print("   Exiting configuration...")
            break
        elif choice == 'S':
            Config.save_config()
            print("   ✅ Configuration saved!")
            break
        elif choice == 'R':
            Config.SORT_INTERVAL_MINUTES = 20
            Config.WEEKLY_REPORT_DAYS = 5
            Config.MIN_SPEED_SECONDS = 5
            Config.MAX_SPEED_SECONDS = 300
            Config.BENCHMARK_PERCENTILE = 0.15
            Config.SHIFT1_START = dt.time(6, 0)
            Config.SHIFT1_END = dt.time(16, 30)
            print("   ✅ Reset to default values!")
            configure_system_interactive()
            break
        elif choice == '1':
            try:
                val = int(input("   Enter sort interval in minutes (default 20): "))
                Config.SORT_INTERVAL_MINUTES = val
                print(f"   ✅ Sort interval set to {val} minutes")
            except ValueError:
                print("   ❌ Invalid input")
        elif choice == '2':
            try:
                val = int(input("   Enter number of business days for report (default 5): "))
                Config.WEEKLY_REPORT_DAYS = val
                print(f"   ✅ Report days set to {val}")
            except ValueError:
                print("   ❌ Invalid input")
        elif choice == '3':
            try:
                val = int(input("   Enter minimum speed in seconds (default 5): "))
                Config.MIN_SPEED_SECONDS = val
                print(f"   ✅ Min speed set to {val} seconds")
            except ValueError:
                print("   ❌ Invalid input")
        elif choice == '4':
            try:
                val = int(input("   Enter maximum speed in seconds (default 300): "))
                Config.MAX_SPEED_SECONDS = val
                print(f"   ✅ Max speed set to {val} seconds")
            except ValueError:
                print("   ❌ Invalid input")
        elif choice == '5':
            try:
                val = float(input("   Enter benchmark percentile (0.10 = top 10%, default 0.15): "))
                Config.BENCHMARK_PERCENTILE = val
                print(f"   ✅ Benchmark percentile set to {val:.0%}")
            except ValueError:
                print("   ❌ Invalid input")
        elif choice == '6':
            try:
                time_str = input("   Enter Shift 1 start time (HH:MM format, default 06:00): ")
                hour, minute = map(int, time_str.split(':'))
                Config.SHIFT1_START = dt.time(hour, minute)
                print(f"   ✅ Shift 1 start time set to {Config.SHIFT1_START}")
            except ValueError:
                print("   ❌ Invalid input")
        elif choice == '7':
            try:
                time_str = input("   Enter Shift 1 end time (HH:MM format, default 16:30): ")
                hour, minute = map(int, time_str.split(':'))
                Config.SHIFT1_END = dt.time(hour, minute)
                print(f"   ✅ Shift 1 end time set to {Config.SHIFT1_END}")
            except ValueError:
                print("   ❌ Invalid input")
        else:
            print("   ❌ Invalid choice")

# ============================================================================
# DATE RANGE SELECTION
# ============================================================================

def select_date_range():
    """Interactive date range selection menu"""
    today = pd.Timestamp.now().date()
    today_weekday = pd.Timestamp.now().weekday()  # 0=Monday, 6=Sunday
    
    print("\n" + "=" * 70)
    print("SELECT REPORT DATE RANGE")
    print("=" * 70)
    print(f"   Today: {today} ({pd.Timestamp.now().strftime('%A')})")
    print()
    
    # Option 1: Yesterday (or last Friday if today is Monday)
    if today_weekday == 0:
        yesterday = today - pd.Timedelta(days=3)
    else:
        yesterday = today - pd.Timedelta(days=1)
    print(f"   [1] Yesterday          -> {yesterday} ({pd.Timestamp(yesterday).strftime('%A')})")
    
    # Option 2: This week (Monday to current day or yesterday)
    this_monday = today - pd.Timedelta(days=today_weekday)
    if today_weekday == 0:
        this_week_end = this_monday
    else:
        this_week_end = today - pd.Timedelta(days=1) if today_weekday <= 4 else today - pd.Timedelta(days=today_weekday - 4)
    print(f"   [2] This week          -> {this_monday} to {this_week_end}")
    
    # Option 3: Last week (last full business week Mon-Fri) - always the previous week
    this_week_monday = today - pd.Timedelta(days=today_weekday)
    lw_monday = this_week_monday - pd.Timedelta(days=7)
    lw_friday = lw_monday + pd.Timedelta(days=4)
    print(f"   [3] Last week          -> {lw_monday} to {lw_friday}")
    
    # Option 4: Last 30 days
    thirty_days_ago = today - pd.Timedelta(days=30)
    print(f"   [4] Last 30 days       -> {thirty_days_ago} to {today}")
    
    # Option 5: Last 7 days
    seven_days_ago = today - pd.Timedelta(days=7)
    print(f"   [5] Last 7 days        -> {seven_days_ago} to {today}")
    
    # Option 6: All available data
    print(f"   [6] All available data  -> All dates in dataset")
    
    # Option 7: Custom date range
    print(f"   [7] Custom date range   -> Enter specific dates")
    
    # Option 8: Interactive Web Report
    print(f"   [8] Interactive Report   -> All-in-One HTML with Date Picker")
    
    print()
    
    while True:
        choice = input("   Enter your choice (1-8): ").strip()
        
        if choice == '1':
            print(f"\n   Selected: Yesterday ({yesterday})")
            return yesterday, yesterday, "yesterday"
        elif choice == '2':
            print(f"\n   Selected: This week ({this_monday} to {this_week_end})")
            return this_monday, this_week_end, "this_week"
        elif choice == '3':
            print(f"\n   Selected: Last week ({lw_monday} to {lw_friday})")
            return lw_monday, lw_friday, "last_week"
        elif choice == '4':
            print(f"\n   Selected: Last 30 days ({thirty_days_ago} to {today})")
            return thirty_days_ago, today, "last_30_days"
        elif choice == '5':
            print(f"\n   Selected: Last 7 days ({seven_days_ago} to {today})")
            return seven_days_ago, today, "last_7_days"
        elif choice == '6':
            print(f"\n   Selected: All available data")
            return None, None, "all_data"
        elif choice == '7':
            print("\n   [CUSTOM DATE RANGE]")
            while True:
                start_input = input("   Enter start date (YYYY-MM-DD): ").strip()
                try:
                    custom_start = pd.to_datetime(start_input).date()
                    break
                except ValueError:
                    print("   ❌ Invalid date format. Please use YYYY-MM-DD.")
            
            while True:
                end_input = input("   Enter end date (YYYY-MM-DD) [Press Enter for today]: ").strip()
                if not end_input:
                    custom_end = today
                    break
                try:
                    custom_end = pd.to_datetime(end_input).date()
                    break
                except ValueError:
                    print("   ❌ Invalid date format. Please use YYYY-MM-DD.")
            
            if custom_start > custom_end:
                print("   ⚠️ Start date is after end date. Swapping them.")
                custom_start, custom_end = custom_end, custom_start
                
            print(f"\n   Selected: Custom Range ({custom_start} to {custom_end})")
            # Format the filename string
            start_str = custom_start.strftime("%Y%m%d")
            end_str = custom_end.strftime("%Y%m%d")
            return custom_start, custom_end, f"custom_{start_str}_to_{end_str}"
        elif choice == '8':
            print(f"\n   Selected: Interactive Web Report (all data with date picker)")
            return None, None, "interactive"
        else:
            print("   Invalid choice. Please enter 1-8.")

# ============================================================================
# OPERATOR NAME LOOKUP
# ============================================================================

def load_operator_names():
    try:
        print("\n[LOADING OPERATOR NAMES]")
        print("-" * 70)
        
        operator_file = Config.DATA_PATH / Config.FILE_OPERATOR_NAMES
        
        if not operator_file.exists():
            print(f"   ⚠️  File not found: {Config.FILE_OPERATOR_NAMES}")
            print("   Continuing without operator name mapping...")
            return pd.DataFrame()
        
        xl_file = pd.ExcelFile(operator_file)
        print(f"   Sheets available: {xl_file.sheet_names}")
        
        df = pd.read_excel(operator_file, sheet_name=0)
        
        print(f"   Loaded: {len(df)} rows")
        print(f"   Columns: {list(df.columns)}")
        
        opcd_col = None
        name_col = None
        
        for col in df.columns:
            col_upper = str(col).upper().strip()
            if 'OPCD' in col_upper or 'CODE' in col_upper or 'OP' in col_upper:
                opcd_col = col
            if 'ASSIGNED' in col_upper:
                if name_col is None:
                    name_col = col
        
        if opcd_col and name_col:
            print(f"   Found: OPCD='{opcd_col}', Name='{name_col}'")
            
            operator_lookup = df[[opcd_col, name_col]].dropna()
            operator_lookup.columns = ['OPCD', 'Operator_Name']
            
            operator_lookup['OPCD'] = operator_lookup['OPCD'].astype(str).str.strip().str.upper()
            operator_lookup['Operator_Name'] = operator_lookup['Operator_Name'].astype(str).str.strip()
            
            operator_lookup = operator_lookup.drop_duplicates(subset='OPCD', keep='first')
            
            print(f"   Mapped {len(operator_lookup)} operator codes to names")
            print(f"   Sample: {operator_lookup.head(3).to_dict('records')}")
            
            return operator_lookup
        else:
            print(f"   ⚠️  Could not find OPCD and Name columns")
            print(f"   Available columns: {list(df.columns)}")
            return pd.DataFrame()
            
    except Exception as e:
        print(f"   ❌ Error loading operator names: {e}")
        return pd.DataFrame()

def map_operator_names(df, operator_lookup):
    """Map OPCD codes to actual operator names"""
    if len(operator_lookup) == 0:
        return df
    
    df = df.merge(
        operator_lookup, 
        left_on='Operator', 
        right_on='OPCD', 
        how='left'
    )
    
    df['Operator_Full'] = df.apply(
        lambda x: f"{x['Operator']}_{x['Operator_Name']}" if pd.notna(x.get('Operator_Name')) else x.get('Operator_Full', x['Operator']),
        axis=1
    )
    
    mapped_count = df['Operator_Name'].notna().sum()
    total_count = len(df)
    
    if mapped_count > 0:
        print(f"   Mapped {mapped_count:,}/{total_count:,} records ({mapped_count/total_count*100:.1f}%) to operator names")
    
    return df

# ============================================================================
# DATA PROCESSING FUNCTIONS
# ============================================================================

def parse_timestamp(timestamp_str):
    """Convert AS400 timestamp to datetime"""
    try:
        ts_str = str(timestamp_str).strip()
        
        if 'E+' in ts_str or 'e+' in ts_str:
            num = float(ts_str)
            ts_str = f"{num:.0f}"
        
        if len(ts_str) >= 12 and ts_str[:12].isdigit():
            return pd.to_datetime(ts_str[:12], format='%Y%m%d%H%M')
        
        return pd.NaT
    except:
        return pd.NaT

def determine_shift(time_obj):
    """Determine shift based on time"""
    if pd.isna(time_obj):
        return 'Unknown'
    
    t = time_obj.time()
    if Config.SHIFT1_START <= t < Config.SHIFT1_END:
        return 'Shift1'
    else:
        return 'Shift2'

def get_operation_date(timestamp):
    if pd.isna(timestamp):
        return pd.NaT
    
    if timestamp.time() <= dt.time(5, 30):
        return (timestamp - pd.Timedelta(days=1)).date()
    return timestamp.date()

def _normalize_location_type(raw):
    """Map a UNIT / LOCATION-prefix value to one of {'M1', 'S1', 'Other'}.

    M1 = small boxes carried by hand (fast).
    S1 = pallets handled by forklift (slow, ~1 module/pallet).
    Anything else (H1/EC/OV/XX/blank/missing) → 'Other'.
    """
    if raw is None:
        return 'Other'
    try:
        if pd.isna(raw):
            return 'Other'
    except Exception:
        pass
    s = str(raw).strip().upper()
    if not s or s == 'NAN':
        return 'Other'
    # Take the prefix up to first '-' (handles putaway "S1-C07-070-00" style)
    prefix = s.split('-', 1)[0][:2]
    if prefix == 'M1':
        return 'M1'
    if prefix == 'S1':
        return 'S1'
    return 'Other'


def classify_mix_solid(df_102):
    """Process 102 file to classify modules as Mix/Solid and tag location type."""
    df = df_102.copy()

    print(f"   102.csv: {len(df):,} rows")
    print(f"   Columns: {list(df.columns)}")

    # ---- Mix/Solid classification (CL column) ----
    cl_col = None
    for col in df.columns:
        col_upper = str(col).upper().strip()
        if col_upper in ['CL', 'CLASS', 'TYPE', 'CLASSIFICATION', 'MODULE_TYPE', 'MOD_TYPE']:
            cl_col = col
            break
        if 'CL' in col_upper or 'CLASS' in col_upper or 'TYPE' in col_upper:
            cl_col = col
            break

    if cl_col:
        print(f"   Found classification column: '{cl_col}'")
        df['Module_Type'] = df[cl_col].map({'SO': 'Solid', 'MO': 'Mix'}).fillna('Mix')
        df['Mix_Solid'] = df[cl_col]
    else:
        print(f"   ⚠️  No classification column found - treating all as Mix")
        df['Module_Type'] = 'Mix'
        df['Mix_Solid'] = 'MO'

    # ---- Location classification (UNIT column → M1/S1/Other) ----
    unit_col = None
    for col in df.columns:
        if str(col).upper().strip() == 'UNIT':
            unit_col = col
            break
    if unit_col:
        print(f"   Found location/unit column in 102: '{unit_col}'")
        df['Location_Type'] = df[unit_col].apply(_normalize_location_type)
    else:
        print(f"   ⚠️  No UNIT column in 102 - location type will fall back to source files")
        df['Location_Type'] = 'Other'

    result = df[['MODULE#', 'Mix_Solid', 'Module_Type', 'Location_Type']].copy()
    result.columns = ['Module', 'Mix_Solid', 'Module_Type', 'Location_Type_102']

    print(f"   - Mix: {(result['Module_Type'] == 'Mix').sum():,}")
    print(f"   - Solid: {(result['Module_Type'] == 'Solid').sum():,}")
    loc_counts = result['Location_Type_102'].value_counts().to_dict()
    print(f"   - Location types (from 102): {loc_counts}")

    return result

def process_sorting_data(df_810, df_102, operator_lookup):
    """Process 810 sorting data"""
    print("\n[2/7] PROCESSING SORTING DATA")
    print("-" * 70)
    
    df = df_810.copy()                         
    print(f"   Raw records: {len(df):,}")
    
    if 'STEP' in df.columns:                
        step_values = df['STEP'].astype(str).unique()
        print(f"   STEP values found: {step_values}")
        
        before_step_filter = len(df)
        df = df[df['STEP'].astype(str).str.strip() == '2']
        print(f"   After STEP='2' filter: {len(df):,} (removed {before_step_filter - len(df):,})")
    
    df['Start_Time'] = df['START TIME'].apply(parse_timestamp)
    df['End_Time'] = df['END TIME'].apply(parse_timestamp)
    
    valid_count = df['Start_Time'].notna().sum()
    print(f"   Valid timestamps: {valid_count:,} ({valid_count/len(df)*100:.1f}%)")
    
    initial_count = len(df)
    df = df.dropna(subset=['Start_Time', 'End_Time'])
    print(f"   After dropping invalid timestamps: {len(df):,} (lost {initial_count - len(df):,})")
    
    if len(df) == 0:
        return pd.DataFrame()
    
    df['Operation_Date'] = df['Start_Time'].apply(get_operation_date)
    df['Shift'] = df['Start_Time'].apply(determine_shift)
    
    print(f"   📅 Date range (before filtering): {df['Operation_Date'].min()} to {df['Operation_Date'].max()}")
    
    pre_merge_count = len(df)
    df = df.merge(df_102, left_on='MODULE', right_on='Module', how='left')
    print(f"   After merge: {len(df):,} (lost {pre_merge_count - len(df):,})")

    # ---- Location type (M1 hand-carry / S1 forklift / Other) ----
    # Prefer the 102 master classification; fall back to the raw UNIT column
    # in the sorting file, then to 'Other'.
    unit_src_col = None
    for col in ['UNIT', 'Unit', 'unit']:
        if col in df.columns:
            unit_src_col = col
            break
    raw_loc = df[unit_src_col].apply(_normalize_location_type) if unit_src_col else pd.Series('Other', index=df.index)
    if 'Location_Type_102' in df.columns:
        df['Location_Type'] = df['Location_Type_102'].where(
            df['Location_Type_102'].notna() & (df['Location_Type_102'] != 'Other'),
            raw_loc,
        )
    else:
        df['Location_Type'] = raw_loc
    df['Location_Type'] = df['Location_Type'].fillna('Other')
    loc_counts = df['Location_Type'].value_counts().to_dict()
    print(f"   Sorting location mix: {loc_counts}")

    # Count by type before filtering
    mix_count = len(df[df['Module_Type'] == 'Mix'])
    solid_count = len(df[df['Module_Type'] == 'Solid'])
    unknown_count = df['Module_Type'].isna().sum()
    print(f"   Module types: Mix={mix_count:,}, Solid={solid_count:,}, Unknown={unknown_count:,}")
    
    # We NO LONGER filter out Solid or Unknown modules for sorting
    # because the user wants to see all sorting activity, even for recent
    # modules that might not be fully classified in 102.csv yet.
    df_mix = df.copy()
    print(f"   Keeping all {len(df_mix):,} sorting records (Mix + Solid + Unknown)")
    
    if len(df_mix) == 0:
        return pd.DataFrame()
    
    before_dedup = len(df_mix)
    df_mix = df_mix.sort_values(['MODULE', 'Start_Time'])
    df_mix = df_mix.drop_duplicates(subset=['Operation_Date', 'MODULE'], keep='first')
    print(f"   🔧 FIX: Deduplicated {before_dedup - len(df_mix):,} duplicate module records")
    
    df_mix['Operator'] = df_mix['OPCD'].fillna('Unknown').astype(str).str.strip().str.upper()
    
    if 'Op 1st Name' in df_mix.columns and 'Op 2nd Name' in df_mix.columns:
        df_mix['Operator_Full'] = (
            df_mix['OPCD'].fillna('') + '_' + 
            df_mix['Op 1st Name'].fillna('') + '_' + 
            df_mix['Op 2nd Name'].fillna('')
        ).str.strip('_')
    else:
        df_mix['Operator_Full'] = df_mix['Operator']
    
    df_mix = map_operator_names(df_mix, operator_lookup)
    
    df_mix['Week'] = df_mix['Operation_Date'].apply(lambda x: x.isocalendar()[1] if pd.notna(x) else None)
    df_mix['Year_Month'] = df_mix['Operation_Date'].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else None)
    
    print(f"   Unique operators: {df_mix['Operator'].nunique()}")
    print(f"   📅 Final date range (Mix only): {df_mix['Operation_Date'].min()} to {df_mix['Operation_Date'].max()}")
    
    return df_mix

def calculate_sorting_speed_metrics(df):
    """Calculate speed metrics"""
    if len(df) == 0:
        return pd.DataFrame()
    
    print("\n[3/7] CALCULATING SORTING SPEEDS")
    print("-" * 70)
    
    df = df.copy()
    if 'Location_Type' not in df.columns:
        df['Location_Type'] = 'Other'
    df = df.sort_values(['Operation_Date', 'Operator', 'Location_Type', 'Start_Time'])

    print(f"   Input records: {len(df):,}")
    print(f"   Input date range: {df['Operation_Date'].min()} to {df['Operation_Date'].max()}")

    # Sessions are split by Location_Type so that fast M1 hand-carry scans and
    # slow S1 forklift moves are never averaged together inside one session.
    df['Time_Gap'] = df.groupby(['Operation_Date', 'Operator', 'Location_Type'])['Start_Time'].diff().dt.total_seconds() / 60
    df['New_Session'] = (df['Time_Gap'] > Config.SORT_INTERVAL_MINUTES) | df['Time_Gap'].isna()
    df['Session_ID'] = df.groupby(['Operation_Date', 'Operator', 'Location_Type'])['New_Session'].cumsum()

    df['Session_Key'] = (df['Operation_Date'].astype(str) + '_' + df['Operator']
                        + '_' + df['Location_Type'].astype(str) + '_' + df['Session_ID'].astype(str))

    session_metrics = df.groupby(['Operator', 'Operator_Full', 'Operation_Date', 'Shift', 'Location_Type', 'Session_Key']).agg({
        'Start_Time': 'min',
        'End_Time': 'max',
        'MODULE': 'count'
    }).reset_index()

    session_metrics.columns = ['Operator', 'Operator_Full', 'Operation_Date', 'Shift', 'Location_Type', 'Session_Key',
                                'Session_Start', 'Session_End', 'Module_Count']
    
    print(f"   Sessions created: {len(session_metrics):,}")
    
    session_metrics['Duration_Minutes'] = (
        (session_metrics['Session_End'] - session_metrics['Session_Start']).dt.total_seconds() / 60
    )
    
    session_metrics.loc[session_metrics['Module_Count'] == 1, 'Duration_Minutes'] = 0.7
    
    session_metrics['Seconds_Per_Module'] = (
        session_metrics['Duration_Minutes'] * 60 / session_metrics['Module_Count']
    ).round(1)
    
    before_filter = len(session_metrics)
    
    session_metrics = session_metrics[
        (session_metrics['Seconds_Per_Module'] >= Config.MIN_SPEED_SECONDS) & 
        (session_metrics['Seconds_Per_Module'] <= Config.MAX_SPEED_SECONDS)
    ]
    
    after_filter = len(session_metrics)
    
    if before_filter > after_filter:
        print(f"   Filtered: {before_filter - after_filter:,} sessions with unrealistic speeds")
    
    session_metrics['Year_Month'] = session_metrics['Operation_Date'].apply(
        lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else None
    )
    
    print(f"   Final sessions: {len(session_metrics):,}")
    print(f"   Avg speed: {session_metrics['Seconds_Per_Module'].mean():.1f} sec/module")
    
    return session_metrics

def process_putaway_data(df_putaway_unit, df_102, operator_lookup):
    """Process putaway UNIT data - count UNIQUE MIX modules per day"""
    print("\n[4/7] PROCESSING PUTAWAY DATA")
    print("-" * 70)
    
    df = df_putaway_unit.copy()
    print(f"   Raw unit records: {len(df):,} rows")
    
    start_col = 'START TIME' if 'START TIME' in df.columns else None
    end_col = 'END TIME' if 'END TIME' in df.columns else None
    
    if not start_col:
        print("   ❌ Missing required columns!")
        return pd.DataFrame()
    
    df['Start_Time'] = df[start_col].apply(parse_timestamp)
    if end_col:
        df['End_Time'] = df[end_col].apply(parse_timestamp)
    
    valid_count = df['Start_Time'].notna().sum()
    print(f"   Valid timestamps: {valid_count:,} ({valid_count/len(df)*100:.1f}%)")
    
    initial_count = len(df)
    df = df.dropna(subset=['Start_Time'])
    print(f"   After dropping invalid: {len(df):,} (lost {initial_count - len(df):,})")
    
    if len(df) == 0:
        return pd.DataFrame()
    
    df['Operation_Date'] = df['Start_Time'].apply(get_operation_date)
    df['Shift'] = df['Start_Time'].apply(determine_shift)
    
    print(f"   📅 Date range: {df['Operation_Date'].min()} to {df['Operation_Date'].max()}")
    
    # Clean MODULE field
    module_col = 'MODULE' if 'MODULE' in df.columns else 'MODULE#' if 'MODULE#' in df.columns else None
    if module_col:
        df['MODULE_CLEAN'] = df[module_col].astype(str).str.replace(' ', '').str.strip()
        # Filter out null/empty modules
        df = df[df['MODULE_CLEAN'].notna() & (df['MODULE_CLEAN'] != '') & (df['MODULE_CLEAN'] != 'nan')]
        print(f"   After removing null modules: {len(df):,}")
    
    # =========================================================================
    # Merge with 102 to get Mix/Solid but DO NOT filter out Solid/Unknown for Putaway
    # (Putaway units are putaway regardless of Mix/Solid status)
    # =========================================================================
    if len(df_102) > 0:
        pre_merge = len(df)
        df = df.merge(df_102, left_on='MODULE_CLEAN', right_on='Module', how='left')
        print(f"   After merge with 102: {len(df):,}")

        # Count by type
        mix_count = len(df[df['Module_Type'] == 'Mix'])
        solid_count = len(df[df['Module_Type'] == 'Solid'])
        unknown_count = df['Module_Type'].isna().sum()
        print(f"   Module types: Mix={mix_count:,}, Solid={solid_count:,}, Unknown={unknown_count:,}")

        # We NO LONGER filter out Solid or Unknown modules for putaway
        # because the user wants to see all putaway activity, even for recent
        # modules that might not be fully classified in 102.csv yet.
        print(f"   Keeping all {len(df):,} putaway records (Mix + Solid + Unknown)")

    if len(df) == 0:
        return pd.DataFrame()

    # ---- Location type from putaway LOCATION column (e.g. "S1-C07-070-00") ----
    location_col = None
    for col in ['LOCATION', 'Location', 'location']:
        if col in df.columns:
            location_col = col
            break
    raw_loc = df[location_col].apply(_normalize_location_type) if location_col else pd.Series('Other', index=df.index)
    if 'Location_Type_102' in df.columns:
        # Trust the unit destination (LOCATION) over the master where present.
        df['Location_Type'] = raw_loc.where(
            raw_loc != 'Other', df['Location_Type_102'].fillna('Other')
        )
    else:
        df['Location_Type'] = raw_loc
    df['Location_Type'] = df['Location_Type'].fillna('Other')
    loc_counts = df['Location_Type'].value_counts().to_dict()
    print(f"   Putaway location mix: {loc_counts}")
    
    # =========================================================================
    # FIX: Deduplicate - each MODULE counts only ONCE per day
    # =========================================================================
    before_dedup = len(df)
    df = df.sort_values(['MODULE_CLEAN', 'Start_Time'])
    df = df.drop_duplicates(subset=['Operation_Date', 'MODULE_CLEAN'], keep='first')
    print(f"   🔧 FIX: Deduplicated {before_dedup - len(df):,} duplicate module records")
    print(f"   Unique module-day combinations: {len(df):,}")
    
    # Calculate duration for speed metrics
    if end_col and 'End_Time' in df.columns:
        df['Duration_Minutes'] = (df['End_Time'] - df['Start_Time']).dt.total_seconds() / 60
        df['Duration_Minutes'] = df['Duration_Minutes'].clip(lower=0.1)
        df['Seconds_Per_Module'] = df['Duration_Minutes'] * 60
        
        # Filter unrealistic speeds
        before_filter = len(df)
        df = df[
            (df['Seconds_Per_Module'] >= 5) & 
            (df['Seconds_Per_Module'] <= 300)
        ]
        if before_filter > len(df):
            print(f"   Filtered: {before_filter - len(df):,} units with unrealistic speeds")
    else:
        df['Duration_Minutes'] = 0.5
        df['Seconds_Per_Module'] = 30
    
    # Each unique module = 1 count
    df['Module_Count'] = 1
    
    # Extract operator information
    df['Operator'] = df['OPCD'].fillna('Unknown').astype(str).str.strip().str.upper()
    
    if 'Op 1st Name' in df.columns and 'Op 2nd Name' in df.columns:
        df['Operator_Full'] = (
            df['OPCD'].fillna('') + '_' + 
            df['Op 1st Name'].fillna('') + '_' + 
            df['Op 2nd Name'].fillna('')
        ).str.strip('_')
    else:
        df['Operator_Full'] = df['Operator']
    
    df = map_operator_names(df, operator_lookup)
    
    df['Week'] = df['Operation_Date'].apply(lambda x: x.isocalendar()[1] if pd.notna(x) else None)
    df['Year_Month'] = df['Operation_Date'].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else None)
    
    print(f"   Final putaway modules: {len(df):,}")
    print(f"   Unique putaway operators: {df['Operator'].nunique()}")
    
    return df

def _per_type_benchmarks(typed_df, op, default_p25, default_p75, default_daily):
    """Compute per-Location_Type speed and daily-volume benchmarks.

    Returns a dict like ``{'M1': {'p25', 'p50', 'p75', 'daily_median'}, ...}``.
    Types with fewer than 5 records / 3 operator-days fall back to the overall
    putaway/sort benchmarks, so an operator who only does S1 isn't crushed by
    a noisy single-record peer pool.
    """
    out = {}
    if 'Location_Type' not in typed_df.columns or len(typed_df) == 0:
        return out
    for ltype, sub in typed_df.groupby('Location_Type'):
        entry = {
            'p25': default_p25, 'p50': (default_p25 + default_p75) / 2.0, 'p75': default_p75,
            'daily_median': float(default_daily),
            'n_records': len(sub),
        }
        if len(sub) >= 5 and 'Seconds_Per_Module' in sub.columns:
            entry['p25'] = float(sub['Seconds_Per_Module'].quantile(0.25))
            entry['p50'] = float(sub['Seconds_Per_Module'].quantile(0.50))
            entry['p75'] = float(sub['Seconds_Per_Module'].quantile(0.75))
        if 'Operation_Date' in sub.columns and 'Operator' in sub.columns:
            daily = sub.groupby(['Operation_Date', 'Operator'])['Module_Count'].sum()
            if len(daily) >= 3:
                entry['daily_median'] = max(float(daily.median()), 1.0)
        out[ltype] = entry
        print(f"   {op:7s} {ltype:5s} – p25={entry['p25']:.1f}  p50={entry['p50']:.1f}  p75={entry['p75']:.1f} sec/mod | peer daily={entry['daily_median']:.0f} (n={entry['n_records']})")
    return out


def calculate_dynamic_benchmarks(sorting_metrics, putaway_metrics):
    """Calculate data-driven speed percentiles from actual operator data.
    Returns a dict with p25/p50/p75 for sorting and putaway, used by
    the unified calculate_efficiency() function so thresholds auto-calibrate
    every time the pipeline runs.

    Also computes per-Location_Type benchmarks (M1 vs S1) so that operators
    doing forklift/pallet work (S1) are compared against S1 peers rather
    than fast hand-carry (M1) peers — see ``benchmarks['by_type']``.
    """
    print("\n[CALCULATING DYNAMIC BENCHMARKS]")
    print("-" * 70)

    benchmarks = {
        'sort_speed': 45,
        'putaway_speed': 30,
        'sort_p25': Config.FALLBACK_SORT_SPEED_P25,
        'sort_p50': (Config.FALLBACK_SORT_SPEED_P25 + Config.FALLBACK_SORT_SPEED_P75) / 2,
        'sort_p75': Config.FALLBACK_SORT_SPEED_P75,
        'putaway_p25': Config.FALLBACK_PUTAWAY_SPEED_P25,
        'putaway_p50': (Config.FALLBACK_PUTAWAY_SPEED_P25 + Config.FALLBACK_PUTAWAY_SPEED_P75) / 2,
        'putaway_p75': Config.FALLBACK_PUTAWAY_SPEED_P75,
        'sort_hourly_median': Config.FALLBACK_SORT_HOURLY_RATE,
        'putaway_hourly_median': Config.FALLBACK_PUTAWAY_HOURLY_RATE,
        'by_type': {'sort': {}, 'putaway': {}},
    }

    if len(sorting_metrics) >= 5:
        benchmarks['sort_p25'] = sorting_metrics['Seconds_Per_Module'].quantile(0.25)
        benchmarks['sort_p50'] = sorting_metrics['Seconds_Per_Module'].quantile(0.50)
        benchmarks['sort_p75'] = sorting_metrics['Seconds_Per_Module'].quantile(0.75)
        sort_benchmark = sorting_metrics['Seconds_Per_Module'].quantile(Config.BENCHMARK_PERCENTILE)
        benchmarks['sort_speed'] = max(sort_benchmark, 20)
        print(f"   Sorting percentiles  – p25={benchmarks['sort_p25']:.1f}  p50={benchmarks['sort_p50']:.1f}  p75={benchmarks['sort_p75']:.1f} sec/module")
    else:
        print(f"   Sorting: <5 records, using fallback percentiles (p25={benchmarks['sort_p25']}, p75={benchmarks['sort_p75']})")

    if len(putaway_metrics) >= 5:
        benchmarks['putaway_p25'] = putaway_metrics['Seconds_Per_Module'].quantile(0.25)
        benchmarks['putaway_p50'] = putaway_metrics['Seconds_Per_Module'].quantile(0.50)
        benchmarks['putaway_p75'] = putaway_metrics['Seconds_Per_Module'].quantile(0.75)
        putaway_benchmark = putaway_metrics['Seconds_Per_Module'].quantile(Config.BENCHMARK_PERCENTILE)
        benchmarks['putaway_speed'] = max(putaway_benchmark * Config.PUTAWAY_BENCHMARK_ADJUSTMENT, 30)
        print(f"   Putaway percentiles  – p25={benchmarks['putaway_p25']:.1f}  p50={benchmarks['putaway_p50']:.1f}  p75={benchmarks['putaway_p75']:.1f} sec/module")
    else:
        print(f"   Putaway: <5 records, using fallback percentiles (p25={benchmarks['putaway_p25']}, p75={benchmarks['putaway_p75']})")

    # ---- Peer-median daily throughput (modules/operator/day) ----
    # Used to normalise the throughput component so sorting and putaway
    # produce comparable efficiency ranges despite very different volumes.
    benchmarks['sort_daily_median'] = float(Config.FALLBACK_DAILY_VOLUME)
    benchmarks['putaway_daily_median'] = float(Config.FALLBACK_DAILY_VOLUME)

    if len(sorting_metrics) >= 5 and 'Operation_Date' in sorting_metrics.columns and 'Operator' in sorting_metrics.columns:
        sort_daily_agg = sorting_metrics.groupby(['Operation_Date', 'Operator']).agg(
            mods=('Module_Count', 'sum'),
            mins=('Duration_Minutes', 'sum'),
        )
        sort_daily = sort_daily_agg['mods']
        if len(sort_daily) >= 3:
            benchmarks['sort_daily_median'] = max(float(sort_daily.median()), 5)
            print(f"   Sorting  peer median – {benchmarks['sort_daily_median']:.0f} modules/operator/day  (p25={sort_daily.quantile(0.25):.0f}, p75={sort_daily.quantile(0.75):.0f})")

        sort_rates = sort_daily_agg[sort_daily_agg['mins'] > 0]
        if len(sort_rates) >= 3:
            rates = sort_rates['mods'] / (sort_rates['mins'] / 60.0)
            benchmarks['sort_hourly_median'] = max(float(rates.median()), 1.0)
            print(f"   Sorting  peer hourly rate – {benchmarks['sort_hourly_median']:.1f} modules/hour  (p25={rates.quantile(0.25):.1f}, p75={rates.quantile(0.75):.1f})")

    if len(putaway_metrics) >= 5 and 'Operation_Date' in putaway_metrics.columns and 'Operator' in putaway_metrics.columns:
        put_daily_agg = putaway_metrics.groupby(['Operation_Date', 'Operator']).agg(
            mods=('Module_Count', 'sum'),
            mins=('Duration_Minutes', 'sum'),
        )
        put_daily = put_daily_agg['mods']
        if len(put_daily) >= 3:
            benchmarks['putaway_daily_median'] = max(float(put_daily.median()), 5)
            print(f"   Putaway  peer median – {benchmarks['putaway_daily_median']:.0f} modules/operator/day  (p25={put_daily.quantile(0.25):.0f}, p75={put_daily.quantile(0.75):.0f})")

        put_rates = put_daily_agg[put_daily_agg['mins'] > 0]
        if len(put_rates) >= 3:
            rates = put_rates['mods'] / (put_rates['mins'] / 60.0)
            benchmarks['putaway_hourly_median'] = max(float(rates.median()), 1.0)
            print(f"   Putaway  peer hourly rate – {benchmarks['putaway_hourly_median']:.1f} modules/hour  (p25={rates.quantile(0.25):.1f}, p75={rates.quantile(0.75):.1f})")

    # ---- Per-Location_Type benchmarks (M1 vs S1 vs Other) ----
    print("\n   Per-location-type benchmarks (M1 = hand-carry, S1 = forklift/pallet):")
    benchmarks['by_type']['sort'] = _per_type_benchmarks(
        sorting_metrics, 'sort',
        benchmarks['sort_p25'], benchmarks['sort_p75'], benchmarks['sort_daily_median'],
    )
    benchmarks['by_type']['putaway'] = _per_type_benchmarks(
        putaway_metrics, 'putaway',
        benchmarks['putaway_p25'], benchmarks['putaway_p75'], benchmarks['putaway_daily_median'],
    )

    return benchmarks


def _speed_score(speed, peer_median_speed):
    """Speed score (log-compressed): ``anchor × (1 + slope × ln(peer_median / operator_speed))``.

    Matching the peer-median speed scores exactly ``EFFICIENCY_MEDIAN_ANCHOR``.
    Faster than median earns positive log credit; slower than median loses
    log credit symmetrically, so neither extreme dominates the geometric mean.

    Tunable via ``Config.SPEED_LOG_SLOPE`` (0.3 default = gentle/forgiving).
    Output clipped to ``[0, EFFICIENCY_SCORE_CAP]``."""
    if speed <= 0 or peer_median_speed <= 0:
        return 0.0
    anchor = Config.EFFICIENCY_MEDIAN_ANCHOR
    cap = Config.EFFICIENCY_SCORE_CAP
    slope = Config.SPEED_LOG_SLOPE
    ratio = float(peer_median_speed) / float(speed)
    score = anchor * (1.0 + slope * np.log(ratio))
    return float(np.clip(score, 0, cap))


def _throughput_score(daily_count, peer_median_count):
    """Volume score (log-compressed): ``anchor × (1 + slope × ln(operator / peer_median))``.

    Matching the peer-median volume scores exactly ``EFFICIENCY_MEDIAN_ANCHOR``.
    The natural log curve softens both extremes: a 2× peer-median day adds
    only ~``anchor × slope × 0.69`` instead of doubling, and a 0.5× day
    loses a similar amount instead of being halved. This keeps single
    huge days from saturating the geometric mean and prevents tiny days
    from being crushed too hard.

    Tunable via ``Config.VOLUME_LOG_SLOPE`` (0.5 default). Capped between
    0 and ``EFFICIENCY_SCORE_CAP``."""
    if daily_count <= 0 or peer_median_count <= 0:
        return 0.0
    anchor = Config.EFFICIENCY_MEDIAN_ANCHOR
    cap = Config.EFFICIENCY_SCORE_CAP
    slope = Config.VOLUME_LOG_SLOPE
    ratio = float(daily_count) / float(peer_median_count)
    score = anchor * (1.0 + slope * np.log(ratio))
    return float(np.clip(score, 0, cap))


def calculate_efficiency(speed, daily_count, effective_goal, p25, p75,
                         w_speed=0.75, w_throughput=0.25,
                         difficulty_bonus=0,
                         hours_worked=None, peer_hourly_rate=None,
                         min_spread=1, peer_median_speed=None):
    """Efficiency = weighted geometric mean of speed score and volume score.

    ``efficiency = speed_score ** w_speed × volume_score ** w_throughput``
    (weights are normalised to sum to 1).

    With the default weights (0.75 / 0.25) speed dominates: a slow operator
    can't recover by piling on volume, but a fast operator with median
    volume can still score very high. Setting ``w_speed = w_throughput =
    0.5`` recovers the plain geometric mean.

    Parameters
    ----------
    speed : float           – seconds per module (lower = faster)
    daily_count : float     – modules processed this day
    effective_goal : float  – peer-median daily count for this operation
    p25, p75 : float        – legacy speed percentiles; midpoint is used
                              as the peer-median speed when
                              ``peer_median_speed`` is not supplied
    w_speed, w_throughput   – exponents for the weighted geometric mean
    difficulty_bonus : float – additive bonus added after the geometric mean
                               (e.g. +6 for putaway physical demands)
    peer_median_speed : float – true peer-median seconds/module (p50);
                                preferred over the p25/p75 midpoint
    """
    if speed <= 0 or daily_count <= 0 or effective_goal <= 0:
        return 0.0

    if peer_median_speed is None or peer_median_speed <= 0:
        peer_median_speed = (p25 + p75) / 2.0

    s_score = _speed_score(speed, peer_median_speed)
    t_score = _throughput_score(daily_count, effective_goal)

    total_w = float(w_speed) + float(w_throughput)
    if total_w <= 0:
        ws, wt = 0.5, 0.5
    else:
        ws, wt = float(w_speed) / total_w, float(w_throughput) / total_w

    if s_score <= 0 or t_score <= 0:
        return round(float(np.clip(difficulty_bonus, 0, 100)), 1)

    efficiency = (s_score ** ws) * (t_score ** wt) + float(difficulty_bonus)
    return round(float(np.clip(efficiency, 0, 100)), 1)


def _apply_shrinkage(scores, counts, k):
    """Blend raw efficiency scores toward the team median by sample size.

    Rationale: a day with 3 modules carries less signal than one with 80.
    Applying weight ``n/(n+k)`` pulls small-sample scores toward the team
    median, reducing noise without masking real underperformance. Rows
    with zero activity (score == 0) are left untouched — they represent
    "didn't work this operation," not weak performance."""
    if k is None or k <= 0:
        return scores
    active = scores[scores > 0]
    if len(active) < 3:
        return scores
    target = float(active.median())
    n = counts.astype(float).clip(lower=0)
    weights = n / (n + float(k))
    adjusted = weights * scores + (1.0 - weights) * target
    return adjusted.where(scores > 0, 0).round(1)


# Backward-compatible wrappers so any stray calls still work
def calculate_sorting_efficiency(speed, weekly_volume, benchmarks=None):
    """Legacy wrapper – delegates to calculate_efficiency()."""
    if benchmarks is None:
        p25, p75 = Config.FALLBACK_SORT_SPEED_P25, Config.FALLBACK_SORT_SPEED_P75
        p50 = (p25 + p75) / 2.0
        eff_goal = float(Config.FALLBACK_DAILY_VOLUME)
    else:
        p25 = benchmarks.get('sort_p25', Config.FALLBACK_SORT_SPEED_P25)
        p75 = benchmarks.get('sort_p75', Config.FALLBACK_SORT_SPEED_P75)
        p50 = benchmarks.get('sort_p50', (p25 + p75) / 2.0)
        eff_goal = benchmarks.get('sort_daily_median', float(Config.FALLBACK_DAILY_VOLUME))
    return calculate_efficiency(speed, weekly_volume, eff_goal, p25, p75,
                                difficulty_bonus=0, peer_median_speed=p50)


def calculate_putaway_efficiency(speed, weekly_volume, benchmarks=None):
    """Legacy wrapper – delegates to calculate_efficiency()."""
    if benchmarks is None:
        p25, p75 = Config.FALLBACK_PUTAWAY_SPEED_P25, Config.FALLBACK_PUTAWAY_SPEED_P75
        p50 = (p25 + p75) / 2.0
        eff_goal = float(Config.FALLBACK_DAILY_VOLUME)
    else:
        p25 = benchmarks.get('putaway_p25', Config.FALLBACK_PUTAWAY_SPEED_P25)
        p75 = benchmarks.get('putaway_p75', Config.FALLBACK_PUTAWAY_SPEED_P75)
        p50 = benchmarks.get('putaway_p50', (p25 + p75) / 2.0)
        eff_goal = benchmarks.get('putaway_daily_median', float(Config.FALLBACK_DAILY_VOLUME))
    return calculate_efficiency(speed, weekly_volume, eff_goal, p25, p75,
                                difficulty_bonus=Config.PUTAWAY_DIFFICULTY_BONUS,
                                peer_median_speed=p50)

def _aggregate_with_location_split(df, prefix, speed_agg,
                                   by_type_bench, default_p25, default_p50, default_p75, default_goal,
                                   w_speed, w_throughput, difficulty_bonus):
    """Aggregate sorting/putaway rows to one row per (date, shift, operator) while:

    1. Splitting counts by Location_Type so M1 (hand-carry boxes) and S1
       (forklift pallets) stay visible per operator.
    2. Computing efficiency *per location type* against that type's peer
       benchmarks – so an S1 operator isn't compared to an M1 operator.
    3. Rolling typed efficiencies/speeds back up using a volume-weighted
       average, so the operator-day score reflects their actual mix of work.

    Returns a DataFrame with the legacy columns (``{prefix}_Count``,
    ``{prefix}_Duration``, ``{prefix}_Speed``, ``{prefix}_Efficiency``) plus
    new per-type breakdown columns (``{prefix}_M1_Count``, ``{prefix}_S1_Count``,
    ``{prefix}_Other_Count``).
    """
    if len(df) == 0:
        return pd.DataFrame()

    df = df.copy()
    if 'Location_Type' not in df.columns:
        df['Location_Type'] = 'Other'

    typed = df.groupby(['Operation_Date', 'Shift', 'Operator', 'Operator_Full', 'Location_Type']).agg({
        'Module_Count': 'sum',
        'Duration_Minutes': 'sum',
        'Seconds_Per_Module': speed_agg,
    }).reset_index()
    typed.columns = ['Operation_Date', 'Shift', 'Operator', 'Operator_Full', 'Location_Type',
                     'Count', 'Duration', 'Speed']

    def _eff(row):
        if row['Speed'] <= 0 or row['Count'] <= 0:
            return 0.0
        bench = by_type_bench.get(row['Location_Type']) if isinstance(by_type_bench, dict) else None
        # Type-specific benchmark needs a real peer pool; otherwise fall back to overall.
        if bench and bench.get('n_records', 0) >= 5:
            p25, p50, p75 = bench['p25'], bench['p50'], bench['p75']
            goal = bench['daily_median']
        else:
            p25, p50, p75, goal = default_p25, default_p50, default_p75, default_goal
        return calculate_efficiency(
            row['Speed'], row['Count'],
            goal, p25, p75,
            w_speed=w_speed, w_throughput=w_throughput,
            difficulty_bonus=difficulty_bonus,
            peer_median_speed=p50,
        )
    typed['Efficiency'] = typed.apply(_eff, axis=1)

    # Collapse to one row per operator-day.
    def _collapse(g):
        total = float(g['Count'].sum())
        if total > 0:
            speed = float((g['Speed'].astype(float) * g['Count']).sum() / total)
            eff = float((g['Efficiency'].astype(float) * g['Count']).sum() / total)
        else:
            speed = 0.0
            eff = 0.0
        m1 = float(g.loc[g['Location_Type'] == 'M1', 'Count'].sum())
        s1 = float(g.loc[g['Location_Type'] == 'S1', 'Count'].sum())
        other = float(g.loc[~g['Location_Type'].isin(['M1', 'S1']), 'Count'].sum())
        return pd.Series({
            f'{prefix}_Count': total,
            f'{prefix}_Duration': float(g['Duration'].sum()),
            f'{prefix}_Speed': speed,
            f'{prefix}_Efficiency': round(eff, 1),
            f'{prefix}_M1_Count': m1,
            f'{prefix}_S1_Count': s1,
            f'{prefix}_Other_Count': other,
        })

    daily = typed.groupby(['Operation_Date', 'Shift', 'Operator', 'Operator_Full']).apply(_collapse).reset_index()
    # Some pandas versions append an extra integer level from `apply` — drop it.
    extra_cols = [c for c in daily.columns if isinstance(c, str) and c.startswith('level_')]
    if extra_cols:
        daily = daily.drop(columns=extra_cols)
    return daily


def aggregate_daily_metrics(df_sorting, df_putaway, report_start, report_end, benchmarks=None):
    """Aggregate metrics by day, shift, and operator with M1/S1 location split.

    Efficiency is computed *per location type* so that operators handling
    slow S1 forklift pallets aren't penalised against fast M1 hand-carry
    peers — see ``_aggregate_with_location_split``.
    """
    print("\n[5/7] AGGREGATING DAILY METRICS")
    print("-" * 70)

    speed_agg = 'median' if Config.USE_MEDIAN_SPEED else 'mean'

    if benchmarks is None:
        benchmarks = {}
    sort_p25 = benchmarks.get('sort_p25', Config.FALLBACK_SORT_SPEED_P25)
    sort_p75 = benchmarks.get('sort_p75', Config.FALLBACK_SORT_SPEED_P75)
    sort_p50 = benchmarks.get('sort_p50', (sort_p25 + sort_p75) / 2.0)
    put_p25  = benchmarks.get('putaway_p25', Config.FALLBACK_PUTAWAY_SPEED_P25)
    put_p75  = benchmarks.get('putaway_p75', Config.FALLBACK_PUTAWAY_SPEED_P75)
    put_p50  = benchmarks.get('putaway_p50', (put_p25 + put_p75) / 2.0)
    sort_eff_goal = benchmarks.get('sort_daily_median', float(Config.FALLBACK_DAILY_VOLUME))
    put_eff_goal  = benchmarks.get('putaway_daily_median', float(Config.FALLBACK_DAILY_VOLUME))
    by_type = benchmarks.get('by_type', {'sort': {}, 'putaway': {}})

    print(f"\n   Sorting  – efficiency = speed^{Config.SORT_SPEED_WEIGHT:.2f} × volume^{Config.SORT_THROUGHPUT_WEIGHT:.2f} (per-type benchmarks)")
    print(f"             default peer-median speed={sort_p50:.1f} sec  |  default peer volume={sort_eff_goal:.0f} mod/day")
    print(f"   Putaway  – efficiency = speed^{Config.PUTAWAY_SPEED_WEIGHT:.2f} × volume^{Config.PUTAWAY_THROUGHPUT_WEIGHT:.2f}  +{Config.PUTAWAY_DIFFICULTY_BONUS}pt (per-type benchmarks)")
    print(f"             default peer-median speed={put_p50:.1f} sec   |  default peer volume={put_eff_goal:.0f} mod/day")

    # STEP 1: Aggregate sorting data with per-type efficiency.
    sorting_daily = _aggregate_with_location_split(
        df_sorting, 'Sort', speed_agg,
        by_type.get('sort', {}), sort_p25, sort_p50, sort_p75, sort_eff_goal,
        Config.SORT_SPEED_WEIGHT, Config.SORT_THROUGHPUT_WEIGHT, 0,
    )
    if len(sorting_daily) > 0:
        print(f"   Sorting: {len(sorting_daily):,} daily records")

    # STEP 2: Aggregate putaway data with per-type efficiency.
    putaway_daily = _aggregate_with_location_split(
        df_putaway, 'Putaway', speed_agg,
        by_type.get('putaway', {}), put_p25, put_p50, put_p75, put_eff_goal,
        Config.PUTAWAY_SPEED_WEIGHT, Config.PUTAWAY_THROUGHPUT_WEIGHT, Config.PUTAWAY_DIFFICULTY_BONUS,
    )
    if len(putaway_daily) > 0:
        print(f"   Putaway: {len(putaway_daily):,} daily records (by operator)")
        print(f"   Total putaway modules: {putaway_daily['Putaway_Count'].sum():,}")

    # STEP 3: Merge sorting and putaway, filling missing sides with zero counts.
    sort_zero_cols = {
        'Sort_Count': 0, 'Sort_Duration': 0, 'Sort_Speed': 0, 'Sort_Efficiency': 0,
        'Sort_M1_Count': 0, 'Sort_S1_Count': 0, 'Sort_Other_Count': 0,
    }
    put_zero_cols = {
        'Putaway_Count': 0, 'Putaway_Duration': 0, 'Putaway_Speed': 0, 'Putaway_Efficiency': 0,
        'Putaway_M1_Count': 0, 'Putaway_S1_Count': 0, 'Putaway_Other_Count': 0,
    }
    if len(sorting_daily) > 0 and len(putaway_daily) > 0:
        daily_metrics = sorting_daily.merge(
            putaway_daily,
            on=['Operation_Date', 'Shift', 'Operator', 'Operator_Full'],
            how='outer'
        ).fillna(0)
    elif len(sorting_daily) > 0:
        daily_metrics = sorting_daily.copy()
        for col, v in put_zero_cols.items():
            daily_metrics[col] = v
    elif len(putaway_daily) > 0:
        daily_metrics = putaway_daily.copy()
        for col, v in sort_zero_cols.items():
            daily_metrics[col] = v
    else:
        return pd.DataFrame()

    # STEP 4: Calculate volumes per operator within the selected report range
    if report_start is not None and report_end is not None:
        range_data = daily_metrics[
            (daily_metrics['Operation_Date'] >= report_start) &
            (daily_metrics['Operation_Date'] <= report_end)
        ]
    else:
        range_data = daily_metrics

    operator_range_sort = range_data.groupby(['Operator', 'Operator_Full'])['Sort_Count'].sum().reset_index()
    operator_range_sort.columns = ['Operator', 'Operator_Full', 'Week_Sort_Volume']

    operator_range_putaway = range_data.groupby(['Operator', 'Operator_Full'])['Putaway_Count'].sum().reset_index()
    operator_range_putaway.columns = ['Operator', 'Operator_Full', 'Week_Putaway_Volume']

    daily_metrics = daily_metrics.merge(operator_range_sort, on=['Operator', 'Operator_Full'], how='left')
    daily_metrics = daily_metrics.merge(operator_range_putaway, on=['Operator', 'Operator_Full'], how='left')
    daily_metrics['Week_Sort_Volume'] = daily_metrics['Week_Sort_Volume'].fillna(0)
    daily_metrics['Week_Putaway_Volume'] = daily_metrics['Week_Putaway_Volume'].fillna(0)

    sort_eff_data = daily_metrics[daily_metrics['Sort_Efficiency'] > 0]
    putaway_eff_data = daily_metrics[daily_metrics['Putaway_Efficiency'] > 0]
    print(f"\n   Efficiency results (peer-calibrated per location type)")
    if len(sort_eff_data) > 0:
        print(f"   Sorting  – range: {sort_eff_data['Sort_Efficiency'].min():.1f}% - {sort_eff_data['Sort_Efficiency'].max():.1f}%  avg: {sort_eff_data['Sort_Efficiency'].mean():.1f}%")
    if len(putaway_eff_data) > 0:
        print(f"   Putaway  – range: {putaway_eff_data['Putaway_Efficiency'].min():.1f}% - {putaway_eff_data['Putaway_Efficiency'].max():.1f}%  avg: {putaway_eff_data['Putaway_Efficiency'].mean():.1f}%")
    
    # STEP 7: Calculate overall efficiency
    total_count = daily_metrics['Sort_Count'] + daily_metrics['Putaway_Count']
    daily_metrics['Overall_Efficiency'] = np.where(
        total_count > 0,
        (
            (daily_metrics['Sort_Efficiency'] * daily_metrics['Sort_Count'] + 
             daily_metrics['Putaway_Efficiency'] * daily_metrics['Putaway_Count']) / 
            total_count
        ),
        0
    ).round(1)
    
    daily_metrics['Year_Month'] = daily_metrics['Operation_Date'].apply(
        lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else None
    )
    
    print(f"\n   Combined: {len(daily_metrics):,} daily records")
    print(f"   Operators: {daily_metrics['Operator'].nunique()}")
    
    return daily_metrics

def analyze_operator_performance(daily_metrics, report_start, report_end):
    """Analyze operator performance for the selected date range"""
    print("\n[6/7] ANALYZING OPERATORS")
    print("-" * 70)
    
    if len(daily_metrics) == 0:
        return pd.DataFrame()
    
    print(f"   Report period: {report_start} to {report_end}")
    
    recent_data = daily_metrics[
        (daily_metrics['Operation_Date'] >= report_start) &
        (daily_metrics['Operation_Date'] <= report_end)
    ].copy()
    
    if len(recent_data) == 0:
        print("   No data in report range!")
        return pd.DataFrame()
    
    print(f"   Records in report range: {len(recent_data):,}")
    
    is_single_day = (report_start == report_end)
    
    # Make sure the per-type count columns are present so groupby below won't KeyError
    for col in ['Sort_M1_Count', 'Sort_S1_Count', 'Putaway_M1_Count', 'Putaway_S1_Count']:
        if col not in recent_data.columns:
            recent_data[col] = 0

    if is_single_day:
        # Daily-style: aggregate by operator AND shift
        operator_stats = recent_data.groupby(['Operator', 'Operator_Full', 'Shift']).agg({
            'Sort_Count': 'sum',
            'Putaway_Count': 'sum',
            'Sort_M1_Count': 'sum',
            'Sort_S1_Count': 'sum',
            'Putaway_M1_Count': 'sum',
            'Putaway_S1_Count': 'sum',
            'Sort_Speed': 'mean',
            'Putaway_Speed': 'mean',
            'Sort_Efficiency': 'mean',
            'Putaway_Efficiency': 'mean'
        }).reset_index()

        operator_stats.columns = ['Operator', 'Operator_Full', 'Shift',
                                  'Day_Sort_Count', 'Day_Putaway_Count',
                                  'Day_Sort_M1', 'Day_Sort_S1', 'Day_Putaway_M1', 'Day_Putaway_S1',
                                  'Day_Sort_Speed', 'Day_Putaway_Speed',
                                  'Day_Sort_Efficiency', 'Day_Putaway_Efficiency']

        operator_stats['Day_Total_Modules'] = operator_stats['Day_Sort_Count'] + operator_stats['Day_Putaway_Count']

        # Rank within each shift
        operator_stats['Shift_Rank'] = operator_stats.groupby('Shift')['Day_Sort_Efficiency'].rank(ascending=False, method='dense')

        print(f"   Active operators: {len(operator_stats)}")
        print(f"   Shift 1 operators: {len(operator_stats[operator_stats['Shift']=='Shift1'])}")
        print(f"   Shift 2 operators: {len(operator_stats[operator_stats['Shift']=='Shift2'])}")
    else:
        # Multi-day: aggregate by operator across the range
        operator_stats = recent_data.groupby(['Operator', 'Operator_Full']).agg({
            'Sort_Count': 'sum',
            'Putaway_Count': 'sum',
            'Sort_M1_Count': 'sum',
            'Sort_S1_Count': 'sum',
            'Putaway_M1_Count': 'sum',
            'Putaway_S1_Count': 'sum',
            'Sort_Speed': 'mean',
            'Putaway_Speed': 'mean',
            'Operation_Date': ['min', 'max', 'nunique']
        }).reset_index()

        operator_stats.columns = ['Operator', 'Operator_Full', 'Week_Sort_Count', 'Week_Putaway_Count',
                                  'Week_Sort_M1', 'Week_Sort_S1', 'Week_Putaway_M1', 'Week_Putaway_S1',
                                  'Week_Sort_Speed', 'Week_Putaway_Speed',
                                  'First_Date', 'Last_Date', 'Days_Worked']
        
        # Compute efficiency only from records where the operator actually sorted/put away
        sort_records = recent_data[recent_data['Sort_Count'] > 0]
        putaway_records = recent_data[recent_data['Putaway_Count'] > 0]
        
        sort_eff = sort_records.groupby(['Operator', 'Operator_Full'])['Sort_Efficiency'].mean().reset_index()
        sort_eff.columns = ['Operator', 'Operator_Full', 'Week_Sort_Efficiency']
        
        putaway_eff = putaway_records.groupby(['Operator', 'Operator_Full'])['Putaway_Efficiency'].mean().reset_index()
        putaway_eff.columns = ['Operator', 'Operator_Full', 'Week_Putaway_Efficiency']
        
        operator_stats = operator_stats.merge(sort_eff, on=['Operator', 'Operator_Full'], how='left')
        operator_stats = operator_stats.merge(putaway_eff, on=['Operator', 'Operator_Full'], how='left')
        operator_stats['Week_Sort_Efficiency'] = operator_stats['Week_Sort_Efficiency'].fillna(0)
        operator_stats['Week_Putaway_Efficiency'] = operator_stats['Week_Putaway_Efficiency'].fillna(0)
        
        operator_stats['Week_Total_Modules'] = operator_stats['Week_Sort_Count'] + operator_stats['Week_Putaway_Count']
        
        operator_stats = operator_stats.sort_values('Week_Sort_Efficiency', ascending=False).reset_index(drop=True)
        operator_stats['Week_Rank'] = range(1, len(operator_stats) + 1)
        
        print(f"   Active operators (report range): {len(operator_stats)}")
        if len(operator_stats) > 0:
            print(f"   Top sort efficiency: {operator_stats['Week_Sort_Efficiency'].max():.1f}%")
            print(f"   Avg sort efficiency: {operator_stats['Week_Sort_Efficiency'].mean():.1f}%")
    
    return operator_stats

def calculate_daily_efficiency_trend(daily_metrics, report_start, report_end):
    """Calculate daily efficiency for the report range - BOTH sorting and putaway"""
    if len(daily_metrics) == 0:
        return pd.DataFrame()
    
    week_data = daily_metrics[
        (daily_metrics['Operation_Date'] >= report_start) &
        (daily_metrics['Operation_Date'] <= report_end)
    ].copy()
    
    if len(week_data) == 0:
        return pd.DataFrame()
    
    # Separate sorting and putaway data
    sort_data = week_data[week_data['Sort_Count'] > 0].copy()
    putaway_data = week_data[week_data['Putaway_Count'] > 0].copy()
    
    # Aggregate sorting metrics by date
    daily_sort = sort_data.groupby('Operation_Date').agg({
        'Sort_Efficiency': 'mean',
        'Sort_Count': 'sum'
    }).reset_index()
    daily_sort.columns = ['Date', 'Sort_Efficiency', 'Sort_Volume']
    
    # Aggregate putaway metrics by date
    daily_putaway = putaway_data.groupby('Operation_Date').agg({
        'Putaway_Efficiency': 'mean',
        'Putaway_Count': 'sum'
    }).reset_index()
    daily_putaway.columns = ['Date', 'Putaway_Efficiency', 'Putaway_Volume']
    
    # Merge sorting and putaway
    daily_trend = daily_sort.merge(daily_putaway, on='Date', how='outer').fillna(0)
    daily_trend = daily_trend.sort_values('Date')
    
    daily_trend['Day_Name'] = pd.to_datetime(daily_trend['Date']).dt.strftime('%a')
    daily_trend['Date_Short'] = pd.to_datetime(daily_trend['Date']).dt.strftime('%m/%d')
    
    return daily_trend

def _format_op_name(op_full):
    """Extract display name from Operator_Full field."""
    parts = str(op_full).split('_')
    return ' '.join(parts[1:]) if len(parts) > 1 else parts[0]


def _analyze_operation_insights(
    operator_data, daily_analysis, label,
    count_col, eff_col, speed_col, is_multi_day
):
    """Generate smart insights for one operation type (sorting OR putaway).

    Parameters
    ----------
    operator_data : DataFrame  – per-operator-day rows (filtered to this op)
    daily_analysis : DataFrame – per-date-shift aggregation (Volume, Efficiency, Operators, …)
    label : str                – '[SORTING]' or '[PUTAWAY]'
    count_col, eff_col, speed_col : str – column names in operator_data
    is_multi_day : bool        – whether the report spans more than one day
    """
    insights = []
    if len(daily_analysis) == 0:
        return insights

    avg_volume = daily_analysis['Volume'].mean()
    avg_efficiency = daily_analysis['Efficiency'].mean()
    avg_operators = daily_analysis['Operators'].mean()
    avg_per_operator = daily_analysis['Modules_Per_Operator'].mean()

    # Peer-median per-operator daily volume — drives staffing recommendations
    # in lieu of a fixed daily target.
    op_counts = operator_data[operator_data[count_col] > 0][count_col]
    peer_median = float(op_counts.median()) if len(op_counts) >= 3 else float(avg_per_operator)
    peer_median = max(peer_median, 1.0)

    # ------------------------------------------------------------------
    # 1. STAFFING RECOMMENDATION – ideal vs actual headcount, anchored on peer median
    # ------------------------------------------------------------------
    for _, row in daily_analysis.iterrows():
        ideal_ops = max(1, int(np.ceil(row['Volume'] / peer_median)))
        actual_ops = int(row['Operators'])
        staffing_delta = actual_ops - ideal_ops
        if staffing_delta <= -2:
            insights.append({
                'type': 'understaffing',
                'priority': 'high',
                'date': row['Date'],
                'day_name': row['Day_Name'],
                'shift': row['Shift'],
                'message': (
                    f"{label} {row['Day_Name']} {row['Shift']} processed {row['Volume']:,.0f} modules "
                    f"with {actual_ops} operators. At the peer-median pace ({peer_median:.0f} modules/op), "
                    f"the ideal team size was {ideal_ops} — {abs(staffing_delta)} operator(s) short. "
                    f"Efficiency was {row['Efficiency']:.1f}%."
                ),
                'volume': row['Volume'],
                'operators': actual_ops,
                'efficiency': row['Efficiency']
            })
        elif staffing_delta >= 3 and row['Modules_Per_Operator'] < peer_median * 0.5:
            insights.append({
                'type': 'overstaffing',
                'priority': 'low',
                'date': row['Date'],
                'day_name': row['Day_Name'],
                'shift': row['Shift'],
                'message': (
                    f"{label} {row['Day_Name']} {row['Shift']} had {actual_ops} operators for "
                    f"{row['Volume']:,.0f} modules (ideal: {ideal_ops} at peer-median pace). "
                    f"{staffing_delta} extra operator(s) could have been redeployed. "
                    f"Average was only {row['Modules_Per_Operator']:.0f} modules/operator."
                ),
                'volume': row['Volume'],
                'operators': actual_ops,
                'efficiency': row['Efficiency']
            })

    # ------------------------------------------------------------------
    # 3. SPEED OUTLIERS – operators significantly slower than the median
    # ------------------------------------------------------------------
    op_speeds = operator_data[operator_data[speed_col] > 0].copy()
    if len(op_speeds) >= 3:
        median_speed = op_speeds[speed_col].median()
        p75_speed = op_speeds[speed_col].quantile(0.75)
        iqr = p75_speed - op_speeds[speed_col].quantile(0.25)
        slow_threshold = p75_speed + 0.5 * max(iqr, 5)

        slow_ops = (
            op_speeds[op_speeds[speed_col] > slow_threshold]
            .groupby('Operator_Full')[speed_col]
            .mean()
            .sort_values(ascending=False)
        )
        if len(slow_ops) > 0:
            names = ', '.join(_format_op_name(n) for n in slow_ops.index[:4])
            extra = f" (+{len(slow_ops)-4} more)" if len(slow_ops) > 4 else ""
            avg_slow = slow_ops.mean()
            pct_slower = ((avg_slow - median_speed) / median_speed) * 100
            insights.append({
                'type': 'speed_outlier',
                'priority': 'medium',
                'date': daily_analysis['Date'].max(),
                'day_name': 'Period',
                'shift': 'Shift1',
                'message': (
                    f"{label} {names}{extra} averaged {avg_slow:.0f} sec/module — "
                    f"{pct_slower:.0f}% slower than the team median of {median_speed:.0f} sec. "
                    f"Targeted coaching could bring them closer to the median and lift overall efficiency."
                ),
                'volume': float(op_speeds[count_col].sum()),
                'operators': int(op_speeds['Operator'].nunique()),
                'efficiency': float(op_speeds[eff_col].mean())
            })

    # ------------------------------------------------------------------
    # 4. TOP PERFORMER RECOGNITION
    # ------------------------------------------------------------------
    op_summary = operator_data.groupby('Operator_Full').agg({
        count_col: 'sum',
        eff_col: 'mean',
        speed_col: 'mean'
    }).reset_index()
    if len(op_summary) >= 2:
        top = op_summary.nlargest(1, eff_col).iloc[0]
        top_name = _format_op_name(top['Operator_Full'])
        insights.append({
            'type': 'top_performer',
            'priority': 'info',
            'date': daily_analysis['Date'].max(),
            'day_name': 'Period',
            'shift': 'Shift1',
            'message': (
                f"{label} Top performer: {top_name} at {top[eff_col]:.1f}% efficiency, "
                f"{top[speed_col]:.0f} sec/module, {top[count_col]:,.0f} modules total. "
                f"Shadowing or sharing their workflow could lift team performance."
            ),
            'volume': float(top[count_col]),
            'operators': 1,
            'efficiency': float(top[eff_col])
        })

    # ------------------------------------------------------------------
    # 5. VOLUME SPIKE with efficiency drop (improved)
    # ------------------------------------------------------------------
    if is_multi_day:
        for _, row in daily_analysis.iterrows():
            if avg_volume > 0:
                volume_ratio = row['Volume'] / avg_volume
                eff_drop = avg_efficiency - row['Efficiency']
                if volume_ratio > 1.15 and eff_drop > 5:
                    insights.append({
                        'type': 'volume_spike',
                        'priority': 'medium',
                        'date': row['Date'],
                        'day_name': row['Day_Name'],
                        'shift': row['Shift'],
                        'message': (
                            f"{label} {row['Day_Name']} {row['Shift']} saw a {volume_ratio*100-100:.0f}% "
                            f"volume spike ({row['Volume']:,.0f} modules vs avg {avg_volume:,.0f}), "
                            f"and efficiency dropped to {row['Efficiency']:.1f}% (down {eff_drop:.1f}pp). "
                            f"Pre-staging extra operators on high-volume days could prevent this."
                        ),
                        'volume': row['Volume'],
                        'operators': int(row['Operators']),
                        'efficiency': row['Efficiency']
                    })

    # ------------------------------------------------------------------
    # 6. BEST PRACTICE – explain WHY the day was optimal
    # ------------------------------------------------------------------
    daily_analysis_scored = daily_analysis.copy()
    daily_analysis_scored['Staffing_Score'] = (
        daily_analysis_scored['Efficiency'] / max(avg_efficiency, 1) * 0.5 +
        daily_analysis_scored['Modules_Per_Operator'] / max(avg_per_operator, 1) * 0.3 +
        (daily_analysis_scored['Operators'] / max(avg_operators, 1)).clip(0.5, 1.5).apply(
            lambda x: 1 - abs(1 - x)
        ) * 0.2
    )
    best = daily_analysis_scored.loc[daily_analysis_scored['Staffing_Score'].idxmax()]
    ideal_for_best = max(1, int(np.ceil(best['Volume'] / peer_median)))
    pct_vs_median = best['Modules_Per_Operator'] / peer_median * 100
    insights.append({
        'type': 'best_practice',
        'priority': 'info',
        'date': best['Date'],
        'day_name': best['Day_Name'],
        'shift': best['Shift'],
        'message': (
            f"{label} {best['Day_Name']} {best['Shift']} achieved the best balance: "
            f"{int(best['Operators'])} operators (ideal: {ideal_for_best} at peer-median pace) handled "
            f"{best['Volume']:,.0f} modules at {best['Efficiency']:.1f}% efficiency "
            f"({best['Modules_Per_Operator']:.0f} modules/operator, "
            f"{pct_vs_median:.0f}% of peer median). "
            f"Replicate this staffing ratio for similar volumes."
        ),
        'volume': best['Volume'],
        'operators': int(best['Operators']),
        'efficiency': best['Efficiency']
    })

    # ------------------------------------------------------------------
    # 7. SHIFT GAP (improved – adds specific recommendation)
    # ------------------------------------------------------------------
    shift_summary = daily_analysis.groupby('Shift').agg({
        'Volume': 'sum',
        'Efficiency': 'mean',
        'Operators': 'mean',
        'Modules_Per_Operator': 'mean'
    }).reset_index()

    if len(shift_summary) == 2:
        s1 = shift_summary[shift_summary['Shift'] == 'Shift1'].iloc[0]
        s2 = shift_summary[shift_summary['Shift'] == 'Shift2'].iloc[0]
        eff_diff = abs(s1['Efficiency'] - s2['Efficiency'])

        if eff_diff > 5:
            better = 'Shift1' if s1['Efficiency'] > s2['Efficiency'] else 'Shift2'
            worse = 'Shift2' if better == 'Shift1' else 'Shift1'
            b_label = '1st Shift' if better == 'Shift1' else '2nd Shift'
            w_label = '2nd Shift' if better == 'Shift1' else '1st Shift'
            b_eff = max(s1['Efficiency'], s2['Efficiency'])
            w_eff = min(s1['Efficiency'], s2['Efficiency'])
            b_mpo = max(s1['Modules_Per_Operator'], s2['Modules_Per_Operator'])
            w_mpo = min(s1['Modules_Per_Operator'], s2['Modules_Per_Operator'])

            insights.append({
                'type': 'shift_gap',
                'priority': 'medium',
                'date': daily_analysis['Date'].max(),
                'day_name': 'Period',
                'shift': worse,
                'message': (
                    f"{label} {b_label} averaged {b_eff:.1f}% efficiency "
                    f"({b_mpo:.0f} modules/op) vs {w_label} at {w_eff:.1f}% "
                    f"({w_mpo:.0f} modules/op) — a {eff_diff:.1f}pp gap. "
                    f"Pair a top {b_label} operator with {w_label} for cross-shift mentoring."
                ),
                'volume': float(daily_analysis[daily_analysis['Shift'] == worse]['Volume'].sum()),
                'operators': int(daily_analysis[daily_analysis['Shift'] == worse]['Operators'].mean()),
                'efficiency': w_eff
            })

    # ------------------------------------------------------------------
    # 8. SPEED TREND (multi-day only)
    # ------------------------------------------------------------------
    if is_multi_day and len(daily_analysis) >= 2:
        daily_speed = operator_data.groupby('Operation_Date')[speed_col].mean().sort_index()
        if len(daily_speed) >= 2:
            first_speed = daily_speed.iloc[0]
            last_speed = daily_speed.iloc[-1]
            if first_speed > 0:
                pct_change = ((last_speed - first_speed) / first_speed) * 100
                if abs(pct_change) > 8:
                    direction = 'improved' if pct_change < 0 else 'slowed'
                    priority = 'info' if pct_change < 0 else 'medium'
                    insights.append({
                        'type': 'speed_trend',
                        'priority': priority,
                        'date': daily_analysis['Date'].max(),
                        'day_name': 'Period',
                        'shift': 'Shift1',
                        'message': (
                            f"{label} Team speed {direction} by {abs(pct_change):.0f}% across the period "
                            f"(from {first_speed:.0f} to {last_speed:.0f} sec/module). "
                            + (f"Great momentum — keep reinforcing current practices."
                               if pct_change < 0 else
                               f"Investigate fatigue, process changes, or module complexity shifts.")
                        ),
                        'volume': float(daily_analysis['Volume'].sum()),
                        'operators': int(daily_analysis['Operators'].mean()),
                        'efficiency': float(daily_analysis['Efficiency'].mean())
                    })

    # ------------------------------------------------------------------
    # 9. CONSISTENCY (improved – names the best and worst days)
    # ------------------------------------------------------------------
    if is_multi_day and len(daily_analysis) >= 3:
        eff_std = daily_analysis['Efficiency'].std()
        if eff_std > 8:
            worst_day = daily_analysis.loc[daily_analysis['Efficiency'].idxmin()]
            best_day = daily_analysis.loc[daily_analysis['Efficiency'].idxmax()]
            insights.append({
                'type': 'consistency',
                'priority': 'medium',
                'date': worst_day['Date'],
                'day_name': worst_day['Day_Name'],
                'shift': worst_day['Shift'],
                'message': (
                    f"{label} Efficiency ranged from {worst_day['Efficiency']:.1f}% "
                    f"({worst_day['Day_Name']} {worst_day['Shift']}) to "
                    f"{best_day['Efficiency']:.1f}% ({best_day['Day_Name']} {best_day['Shift']}) — "
                    f"std dev {eff_std:.1f}pp. Standardising start-of-shift procedures and "
                    f"workload distribution could tighten this range."
                ),
                'volume': worst_day['Volume'],
                'operators': int(worst_day['Operators']),
                'efficiency': worst_day['Efficiency']
            })

    return insights


def analyze_workforce_optimization(daily_metrics, operator_stats, report_start, report_end):
    """Generate smart, data-driven workforce optimization insights for both
    sorting and putaway.  Works for single-day and multi-day reports."""
    print("\n[ANALYZING WORKFORCE OPTIMIZATION]")
    print("-" * 70)

    if len(daily_metrics) == 0:
        return []

    insights = []
    is_multi_day = (report_start != report_end)

    range_data = daily_metrics[
        (daily_metrics['Operation_Date'] >= report_start) &
        (daily_metrics['Operation_Date'] <= report_end)
    ].copy()

    if len(range_data) == 0:
        return []

    # ---- helper: build daily_analysis for one operation type ----
    def _build_daily_analysis(df, count_col, eff_col):
        da = df.groupby(['Operation_Date', 'Shift']).agg({
            count_col: 'sum',
            eff_col: 'mean',
            'Operator': 'nunique'
        }).reset_index()
        da.columns = ['Date', 'Shift', 'Volume', 'Efficiency', 'Operators']
        da['Day_Name'] = pd.to_datetime(da['Date']).dt.strftime('%A')
        da['Modules_Per_Operator'] = da['Volume'] / da['Operators']
        return da

    # ---- SORTING INSIGHTS ----
    sort_data = range_data[range_data['Sort_Count'] > 0].copy()
    if len(sort_data) > 0:
        da_sort = _build_daily_analysis(sort_data, 'Sort_Count', 'Sort_Efficiency')
        insights += _analyze_operation_insights(
            sort_data, da_sort, '[SORTING]',
            'Sort_Count', 'Sort_Efficiency', 'Sort_Speed', is_multi_day
        )

    # ---- PUTAWAY INSIGHTS ----
    put_data = range_data[range_data['Putaway_Count'] > 0].copy()
    if len(put_data) > 0:
        da_put = _build_daily_analysis(put_data, 'Putaway_Count', 'Putaway_Efficiency')
        insights += _analyze_operation_insights(
            put_data, da_put, '[PUTAWAY]',
            'Putaway_Count', 'Putaway_Efficiency', 'Putaway_Speed', is_multi_day
        )

    # ---- CROSS-TRAINING INSIGHT (cross-operation) ----
    sort_operators = set(sort_data['Operator'].unique()) if len(sort_data) > 0 else set()
    put_operators = set(put_data['Operator'].unique()) if len(put_data) > 0 else set()
    both = sort_operators & put_operators
    sort_only = sort_operators - put_operators
    put_only = put_operators - sort_operators

    if len(sort_only) > 0 and len(put_only) > 0:
        total_ops = len(sort_operators | put_operators)
        cross_pct = len(both) / total_ops * 100 if total_ops > 0 else 0
        sort_only_names = ', '.join(
            _format_op_name(
                range_data[range_data['Operator'] == op]['Operator_Full'].iloc[0]
            ) for op in list(sort_only)[:3]
        )
        put_only_names = ', '.join(
            _format_op_name(
                range_data[range_data['Operator'] == op]['Operator_Full'].iloc[0]
            ) for op in list(put_only)[:3]
        )
        s_extra = f" (+{len(sort_only)-3} more)" if len(sort_only) > 3 else ""
        p_extra = f" (+{len(put_only)-3} more)" if len(put_only) > 3 else ""

        if cross_pct < 40:
            insights.append({
                'type': 'cross_training',
                'priority': 'low',
                'date': report_end,
                'day_name': 'Period',
                'shift': 'Shift1',
                'message': (
                    f"Only {len(both)} of {total_ops} operators ({cross_pct:.0f}%) "
                    f"worked both sorting and putaway. Sort-only: {sort_only_names}{s_extra}. "
                    f"Putaway-only: {put_only_names}{p_extra}. "
                    f"Cross-training improves scheduling flexibility and reduces single-point-of-failure risk."
                ),
                'volume': float(range_data['Sort_Count'].sum() + range_data['Putaway_Count'].sum()),
                'operators': total_ops,
                'efficiency': float(range_data[range_data['Sort_Efficiency'] > 0]['Sort_Efficiency'].mean())
            })

    # ---- SORT & PRIORITISE ----
    priority_order = {'high': 0, 'medium': 1, 'low': 2, 'info': 3}
    insights.sort(key=lambda x: priority_order.get(x['priority'], 9))

    print(f"   Generated {len(insights)} workforce insights")
    for insight in insights[:5]:
        print(f"   [{insight['priority'].upper()}] {insight['message'][:90]}...")

    return insights

def predict_daily_volumes(daily_metrics):
    """ML predictions"""
    if not ML_AVAILABLE or len(daily_metrics) == 0:
        return None, None
    
    print("\n[7/7] ML PREDICTIONS (SHIFT-LEVEL)")
    print("-" * 70)
    
    df = daily_metrics[daily_metrics['Sort_Count'] > 0].copy()
    
    if len(df) == 0:
        print("   No operator data for predictions")
        return None, None
    
    shift_daily = df.groupby(['Operation_Date', 'Shift']).agg({
        'Sort_Count': 'sum',
        'Putaway_Count': 'sum',
        'Sort_Efficiency': 'mean',
        'Operator': 'nunique'
    }).reset_index()
    
    shift_daily.columns = ['Operation_Date', 'Shift', 'Sort_Count', 'Putaway_Count', 
                           'Efficiency', 'Active_Operators']
    
    shift_daily['DayOfWeek'] = pd.to_datetime(shift_daily['Operation_Date']).dt.dayofweek
    shift_daily['WeekOfYear'] = pd.to_datetime(shift_daily['Operation_Date']).dt.isocalendar().week
    shift_daily['Month'] = pd.to_datetime(shift_daily['Operation_Date']).dt.month
    shift_daily['IsWeekend'] = shift_daily['DayOfWeek'].isin([5, 6]).astype(int)
    shift_daily['Is_Shift1'] = (shift_daily['Shift'] == 'Shift1').astype(int)
    
    shift_daily = shift_daily.sort_values(['Shift', 'Operation_Date'])
    
    for col in ['Sort_Count', 'Efficiency', 'Active_Operators']:
        shift_daily[f'{col}_Lag1'] = shift_daily.groupby('Shift')[col].shift(1)
        shift_daily[f'{col}_Lag7'] = shift_daily.groupby('Shift')[col].shift(7)
    
    shift_daily['Sort_Count_MA7'] = shift_daily.groupby('Shift')['Sort_Count'].transform(
        lambda x: x.rolling(7, min_periods=1).mean()
    )
    shift_daily['Sort_Count_MA30'] = shift_daily.groupby('Shift')['Sort_Count'].transform(
        lambda x: x.rolling(30, min_periods=1).mean()
    )
    
    shift_daily = shift_daily.dropna()
    
    if len(shift_daily) < 30:
        print(f"   Insufficient data: {len(shift_daily)} records (need 30+)")
        return None, None
    
    feature_cols = [
        'DayOfWeek', 'WeekOfYear', 'Month', 'IsWeekend', 'Is_Shift1',
        'Sort_Count_Lag1', 'Efficiency_Lag1', 'Active_Operators_Lag1',
        'Sort_Count_MA7', 'Sort_Count_MA30'
    ]
    
    shift_daily['Sort_Count_Next'] = shift_daily.groupby('Shift')['Sort_Count'].shift(-1)
    shift_daily['Efficiency_Next'] = shift_daily.groupby('Shift')['Efficiency'].shift(-1)
    
    train_df = shift_daily.dropna(subset=['Sort_Count_Next', 'Efficiency_Next'])
    
    if len(train_df) < 20:
        print(f"   Insufficient training data: {len(train_df)}")
        return None, None
    
    X = train_df[feature_cols]
    y_volume = train_df['Sort_Count_Next']
    y_efficiency = train_df['Efficiency_Next']
    
    total_training_modules = train_df['Sort_Count'].sum()
    
    model_volume = RandomForestRegressor(n_estimators=100, random_state=42, max_depth=8)
    model_efficiency = RandomForestRegressor(n_estimators=100, random_state=42, max_depth=8)
    
    model_volume.fit(X, y_volume)
    model_efficiency.fit(X, y_efficiency)
    
    print(f"   Trained on: {len(train_df):,} shift-days ({total_training_modules:,} modules)")
    
    latest_date = shift_daily['Operation_Date'].max()
    latest_data = shift_daily[shift_daily['Operation_Date'] == latest_date]
    
    if len(latest_data) == 0:
        print("   No recent data for predictions")
        return None, total_training_modules
    
    pred_volume = model_volume.predict(latest_data[feature_cols])
    pred_efficiency = model_efficiency.predict(latest_data[feature_cols])
    
    result = latest_data[['Shift']].copy()
    result['Predicted_Sort_Volume'] = pred_volume.round(0).astype(int)
    result['Predicted_Efficiency'] = pred_efficiency.round(1)
    result['Prediction_Date'] = latest_date + pd.Timedelta(days=1)
    result['Expected_Operators'] = latest_data['Active_Operators_Lag1'].values
    
    total_row = pd.DataFrame({
        'Shift': ['Total'],
        'Predicted_Sort_Volume': [result['Predicted_Sort_Volume'].sum()],
        'Predicted_Efficiency': [result['Predicted_Efficiency'].mean()],
        'Prediction_Date': [result['Prediction_Date'].iloc[0]],
        'Expected_Operators': [result['Expected_Operators'].sum()]
    })
    
    result = pd.concat([result, total_row], ignore_index=True)
    
    print(f"   Predictions:")
    for _, row in result.iterrows():
        print(f"     {row['Shift']}: {row['Predicted_Sort_Volume']:,.0f} modules @ {row['Predicted_Efficiency']:.1f}% efficiency")
    
    return result, total_training_modules

def calculate_shift_breakdown(daily_metrics, report_start, report_end):
    """Calculate detailed shift breakdown"""
    if len(daily_metrics) == 0:
        return {}, ()
    
    week_data = daily_metrics[
        (daily_metrics['Operation_Date'] >= report_start) &
        (daily_metrics['Operation_Date'] <= report_end)
    ].copy()
    
    if len(week_data) == 0:
        return {}, ()
    
    shift_breakdown = {}
    
    for date in pd.date_range(report_start, report_end):
        date_data = week_data[week_data['Operation_Date'] == date.date()]
        
        if len(date_data) > 0:
            shift_breakdown[date.date()] = {
                'Shift1': date_data[date_data['Shift'] == 'Shift1'],
                'Shift2': date_data[date_data['Shift'] == 'Shift2']
            }
    
    return shift_breakdown, (report_start, report_end)

def generate_email_content(daily_metrics, operator_stats, daily_trend, report_start, report_end, range_label="", workforce_insights=None, operator_speeds=None, predictions=None, training_modules=None, dashboard_html=None, fragment_only=False):
    """Generate HTML email for supervisors. If fragment_only=True, return only the inner content div without html/head/body wrapper."""
    if len(daily_metrics) == 0:
        if fragment_only:
            return '<div class="container"><h1>No data available</h1></div>'
        return "<html><body><h1>No data available</h1></body></html>"
    
    last_monday = report_start
    last_friday = report_end
    
    # Dynamic title based on date range
    num_days = (pd.Timestamp(report_end) - pd.Timestamp(report_start)).days + 1
    start_fmt = pd.Timestamp(report_start).strftime('%B %d, %Y')
    end_fmt = pd.Timestamp(report_end).strftime('%B %d, %Y')
    
    if range_label == "last_week":
        report_title = "Weekly Sorting and Putaway KPI Report"
        report_subtitle = f"WEEK OF {pd.Timestamp(report_start).strftime('%B %d')} - {pd.Timestamp(report_end).strftime('%B %d, %Y')}"
        overview_title = "Weekly Overview"
    elif range_label == "this_week":
        report_title = "Week-to-Date Sorting and Putaway KPI Report"
        report_subtitle = f"{start_fmt} - {end_fmt}"
        overview_title = "Week-to-Date Overview"
    elif range_label == "last_7_days":
        report_title = "7-Day Sorting and Putaway KPI Report"
        report_subtitle = f"{start_fmt} - {end_fmt}"
        overview_title = "7-Day Overview"
    elif range_label == "last_30_days":
        report_title = "30-Day Sorting and Putaway KPI Report"
        report_subtitle = f"{start_fmt} - {end_fmt}"
        overview_title = "30-Day Overview"
    elif range_label == "all_data":
        report_title = "Sorting and Putaway KPI Report - All Data"
        report_subtitle = f"{start_fmt} - {end_fmt}"
        overview_title = "Full Period Overview"
    elif range_label == "yesterday":
        report_title = "Daily Sorting and Putaway KPI Report"
        report_subtitle = f"{start_fmt}"
        overview_title = "Daily Overview"
    elif "custom" in range_label:
        report_title = "Custom Date Sorting and Putaway KPI Report"
        report_subtitle = f"{start_fmt} - {end_fmt}"
        overview_title = "Custom Period Overview"
    else:
        report_title = "Sorting and Putaway KPI Report"
        report_subtitle = f"REPORT PERIOD: {start_fmt} - {end_fmt}"
        overview_title = "Period Overview"
    
    week_data = daily_metrics[
        (daily_metrics['Operation_Date'] >= last_monday) &
        (daily_metrics['Operation_Date'] <= last_friday)
    ]
    
    print(f"\n[REPORT GENERATION]")
    print(f"   Report data records: {len(week_data)}")
    print(f"   Dates in report data: {sorted(week_data['Operation_Date'].unique())}")
    
    week_sort = week_data['Sort_Count'].sum()
    week_putaway = week_data['Putaway_Count'].sum()
    week_sort_efficiency = week_data[week_data['Sort_Count'] > 0]['Sort_Efficiency'].mean()
    week_putaway_efficiency = week_data[week_data['Putaway_Count'] > 0]['Putaway_Efficiency'].mean()
    
    shift_breakdown, week_range = calculate_shift_breakdown(daily_metrics, report_start, report_end)
    
    # Get top 5 sorting performers
    if len(operator_stats) > 0 and 'Week_Sort_Count' in operator_stats.columns:
        top_5_sorting = operator_stats[operator_stats['Week_Sort_Count'] > 0].nlargest(5, 'Week_Sort_Efficiency')
    else:
        top_5_sorting = pd.DataFrame()
    
    # Get top 5 putaway performers
    if len(operator_stats) > 0 and 'Week_Putaway_Count' in operator_stats.columns:
        top_5_putaway = operator_stats[operator_stats['Week_Putaway_Count'] > 0].nlargest(5, 'Week_Putaway_Efficiency')
    else:
        top_5_putaway = pd.DataFrame()
    
    if len(daily_trend) > 0:
        best_day = daily_trend.loc[daily_trend['Sort_Efficiency'].idxmax()]
        best_day_name = f"{best_day['Day_Name']} {best_day['Date_Short']}"
        best_day_eff = best_day['Sort_Efficiency']
        best_day_vol = best_day['Sort_Volume']
    else:
        best_day_name = "N/A"
        best_day_eff = 0
        best_day_vol = 0
    
    if len(daily_trend) > 0:
        high_vol_day = daily_trend.loc[daily_trend['Sort_Volume'].idxmax()]
        high_vol_day_name = f"{high_vol_day['Day_Name']} {high_vol_day['Date_Short']}"
        high_vol_day_count = high_vol_day['Sort_Volume']
    else:
        high_vol_day_name = "N/A"
        high_vol_day_count = 0
    
    active_ops = len(operator_stats[operator_stats['Week_Sort_Count'] > 0]) if len(operator_stats) > 0 and 'Week_Sort_Count' in operator_stats.columns else 0

    if not fragment_only:
        html = f"""
    <html>
    <head>
        <style>
            * {{
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }}
            body {{
                font-family: Arial, Helvetica, sans-serif;
                background: #ffffff;
                padding: 0;
                color: #2c3e50;
            }}
            .container {{
                max-width: 100%;
                margin: 0;
                background: white;
                box-shadow: none;
                overflow: hidden;
            }}
            .header {{
                background: transparent;
                color: #2f3b45;
                padding: 18px 30px;
                text-align: center;
                border-bottom: none;
            }}
            .header h1 {{
                margin: 6px 0 0 0;
                font-size: 1.6em;
                font-weight: 600;
                letter-spacing: 0.6px;
                line-height: 1.05;
            }}
            .header p {{
                margin: 0;
                font-size: 0.9em;
                font-weight: 600;
                color: #4f5b66;
                line-height: 1.05;
            }}
            .section {{
                padding: 20px 30px;
                background: white;
                border-bottom: 1px solid #dee2e6;
            }}
            .section h2 {{
                color: #2c3e50;
                border-bottom: 2px solid #3498db;
                padding-bottom: 8px;
                margin-top: 0;
                font-size: 1.1em;
                letter-spacing: 0.4px;
            }}
            .metrics-grid {{
                display: grid;
                grid-template-columns: 0.15fr 0.35fr 0.15fr 0.35fr;
                gap: 20px;
                margin: 20px 0 0 0;
            }}
            .metric-box {{
                background: white;
                color: #2c3e50;
                padding: 16px 20px;
                border-top: 4px solid #95a5a6;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                min-height: 90px;
            }}
            .metric-box.primary {{ border-top-color: #3498db; }}
            .metric-box.success {{ border-top-color: #28a745; }}
            .metric-box.warning {{ border-top-color: #f4b400; }}
            .metric-box.info {{ border-top-color: #7f8c8d; }}
            .metric-value {{
                font-size: 2.2em;
                font-weight: bold;
                margin: 8px 0 4px 0;
                line-height: 1;
            }}
            .metric-label {{
                font-size: 0.85em;
                color: #7f8c8d;
                text-transform: uppercase;
                letter-spacing: 0.6px;
                font-weight: 600;
            }}
            .metric-sublabel {{
                font-size: 0.95em;
                color: #4f5b66;
                margin-top: 4px;
                font-weight: 500;
            }}
            .top-performers {{
                background-color: #f8f9fa;
                padding: 8px;
                border-radius: 4px;
                margin: 8px 0 0 0;
            }}
            .top-performers-item {{
                padding: 4px 0;
                border-bottom: 1px solid #dee2e6;
                font-size: 1.05em;
                color: #2c3e50;
                font-weight: 600;
            }}
            .top-performers-item:last-child {{ border-bottom: none; }}
            .chart-container {{
                padding: 30px;
                background: white;
                text-align: center;
            }}
            .chart-container > div {{
                display: inline-block;
                text-align: left;
            }}
            .day-card-container {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
                gap: 15px;
                margin: 20px 0 0 0;
            }}
            .day-card {{
                background: white;
                border-radius: 4px;
                padding: 20px;
                border-top: 4px solid #95a5a6;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                text-align: center;
                transition: transform 0.2s, box-shadow 0.2s;
            }}
            .day-card:hover {{
                transform: translateY(-3px);
                box-shadow: 0 8px 15px rgba(0,0,0,0.2);
            }}
            .day-card-header {{
                font-size: 1.1em;
                font-weight: 600;
                color: #2c3e50;
                margin-bottom: 4px;
            }}
            .day-card-date {{
                font-size: 0.95em;
                color: #6c757d;
                margin-bottom: 12px;
            }}
            .day-card-efficiency {{
                font-size: 2.2em;
                font-weight: bold;
                margin: 10px 0;
            }}
            .day-card-efficiency.high {{ color: #28a745; }}
            .day-card-efficiency.medium {{ color: #f4b400; }}
            .day-card-efficiency.low {{ color: #dc3545; }}
            .day-card-volume {{
                font-size: 1.1em;
                color: #4f5b66;
                margin-top: 8px;
                font-weight: 600;
            }}
            .day-card-label {{
                font-size: 0.8em;
                color: #7f8c8d;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                margin-bottom: 4px;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin: 20px 0 0 0;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            th {{
                background-color: #95a5a6;
                color: white;
                padding: 12px;
                text-align: left;
                font-weight: 600;
                font-size: 0.85em;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            td {{
                padding: 10px 12px;
                border-bottom: 1px solid #ecf0f1;
            }}
            tr:hover {{ background-color: #f8f9fa; }}
            .efficiency-high {{ color: #28a745; font-weight: bold; }}
            .efficiency-medium {{ color: #f4b400; font-weight: bold; }}
            .efficiency-low {{ color: #dc3545; font-weight: bold; }}
            .insights-box {{
                background-color: #f8f9fa;
                border-left: 4px solid #3498db;
                padding: 20px;
                margin: 20px 0 0 0;
                border-radius: 4px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            }}
            .insights-box h3 {{ color: #2c3e50; margin-top: 0; }}
            .insight-item {{
                background-color: white;
                padding: 15px;
                margin: 12px 0;
                border-radius: 4px;
                border-left: 3px solid #95a5a6;
            }}
            .insight-item.priority-high {{ border-left-color: #dc3545; background-color: #f8d7da; }}
            .insight-item.priority-medium {{ border-left-color: #f4b400; background-color: #fff3cd; }}
            .insight-item.priority-low {{ border-left-color: #3498db; background-color: #e3f2fd; }}
            .insight-header {{
                font-weight: bold;
                font-size: 0.9em;
                margin-bottom: 8px;
                display: flex;
                align-items: center;
                gap: 8px;
            }}
            .insight-badge {{
                display: inline-block;
                padding: 3px 8px;
                border-radius: 4px;
                font-size: 0.75em;
                font-weight: bold;
                text-transform: uppercase;
            }}
            .badge-high {{ background-color: #dc3545; color: white; }}
            .badge-medium {{ background-color: #f4b400; color: white; }}
            .badge-low {{ background-color: #3498db; color: white; }}
            .badge-info {{ background-color: #28a745; color: white; }}
            .insight-details {{
                font-size: 0.85em;
                color: #6c757d;
                margin-top: 5px;
            }}
            .shift-breakdown {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 20px;
                margin: 20px 0 0 0;
            }}
            .shift-table {{
                background-color: #ffffff;
                border-radius: 4px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                overflow-x: auto;
            }}
            .shift-header {{
                background: #f1f3f5;
                color: #2c3e50;
                padding: 12px;
                font-size: 0.95em;
                font-weight: 600;
                text-align: center;
                border-bottom: 1px solid #cfd4da;
            }}
            .shift-table table {{ width: 100%; margin: 0; table-layout: fixed; }}
            .shift-table th {{
                background-color: #95a5a6;
                color: white;
                padding: 10px 8px;
                font-size: 0.75em;
                font-weight: 600;
            }}
            .shift-table td {{
                padding: 8px 6px;
                text-align: center;
                font-size: 0.85em;
                word-wrap: break-word;
                overflow: hidden;
                text-overflow: ellipsis;
            }}
            .perf-under {{ background-color: #f8d7da; color: #721c24; font-weight: bold; }}
            .perf-at {{ background-color: #fff3cd; color: #856404; font-weight: bold; }}
            .perf-above {{ background-color: #d4edda; color: #155724; font-weight: bold; }}
            .total-row {{ background-color: #e9ecef; font-weight: bold; border-top: 2px solid #95a5a6; }}
            .speed-fast {{ background-color: #d4edda; color: #155724; }}
            .speed-average {{ background-color: #fff3cd; color: #856404; }}
            .speed-slow {{ background-color: #f8d7da; color: #721c24; }}
            .footer {{
                background: transparent;
                text-align: center;
                padding: 20px;
                color: #4f5b66;
                font-size: 0.85em;
                border-top: none;
            }}
            .dashboard-toggle-btn {{
                display: inline-block;
                background: #3498db;
                color: white;
                padding: 12px 30px;
                border: none;
                border-radius: 6px;
                font-size: 1em;
                font-weight: 600;
                cursor: pointer;
                letter-spacing: 0.5px;
                transition: background 0.2s;
                margin: 10px 0;
            }}
            .dashboard-toggle-btn:hover {{ background: #2980b9; }}
            .dashboard-section {{
                display: none;
                padding: 20px 30px;
                background: #f8f9fa;
                border-bottom: 1px solid #dee2e6;
            }}
            .dashboard-section.visible {{ display: block; }}
            .dashboard-grid {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 20px;
            }}
            .dashboard-chart {{
                background: white;
                border-radius: 6px;
                box-shadow: 0 2px 8px rgba(0,0,0,0.08);
                padding: 10px;
                overflow: hidden;
            }}
            .dashboard-chart.full-width {{
                grid-column: 1 / -1;
            }}
        </style>
        <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
    </head>
    <body>"""
    else:
        html = ""

    html += f"""
        <div class="container">
            <div class="header">
                <h1>{report_title}</h1>
                <p>{report_subtitle}</p>
            </div>
        
        <div class="section">
            <h2>{overview_title}</h2>
            <div class="metrics-grid">
                <div class="metric-box primary">
                    <div class="metric-label">SORTING</div>
                    <div class="metric-value" style="font-size: 42px;">{week_sort:,.0f}</div>
                    <div class="metric-sublabel">Modules Sorted</div>
                </div>
                <div class="metric-box info">
                    <div class="metric-label">🏆 TOP 5 SORTING</div>
                    <div class="top-performers">
    """
    
    if len(top_5_sorting) > 0:
        for rank, (idx, row) in enumerate(top_5_sorting.iterrows(), start=1):
            op_name = row['Operator_Full']
            parts = str(op_name).split('_')
            op_name = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]
            html += f'<div class="top-performers-item">{rank}. {op_name}: {row["Week_Sort_Efficiency"]:.1f}%</div>'
    else:
        html += '<div class="top-performers-item">No data available</div>'
    
    html += f"""
                    </div>
                </div>
                <div class="metric-box success">
                    <div class="metric-label">PUTAWAY</div>
                    <div class="metric-value" style="font-size: 42px;">{week_putaway:,.0f}</div>
                    <div class="metric-sublabel">Modules Put Away</div>
                </div>
                <div class="metric-box info">
                    <div class="metric-label">🏆 TOP 5 PUTAWAY</div>
                    <div class="top-performers">
    """
    
    if len(top_5_putaway) > 0:
        for rank, (idx, row) in enumerate(top_5_putaway.iterrows(), start=1):
            op_name = row['Operator_Full']
            parts = str(op_name).split('_')
            op_name = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]
            html += f'<div class="top-performers-item">{rank}. {op_name}: {row["Week_Putaway_Efficiency"]:.1f}%</div>'
    else:
        html += '<div class="top-performers-item">No data available</div>'
    
    html += """
                    </div>
                </div>
            </div>
        </div>
    """
    
    # Daily Performance Cards - only show individual cards for <=10 days
    daily_totals = week_data.groupby('Operation_Date').agg({
        'Sort_Count': 'sum',
        'Putaway_Count': 'sum'
    }).reset_index().sort_values('Operation_Date')
    
    if len(daily_totals) <= 10:
        html += '<div class="section"><h2>Daily Performance</h2>'
        html += '<div class="day-card-container">'
        
        for _, row in daily_totals.iterrows():
            day_name = pd.to_datetime(row['Operation_Date']).strftime('%A')
            date_str = pd.to_datetime(row['Operation_Date']).strftime('%m/%d')
            
            html += f'''<div class="day-card">
                <div class="day-card-header">{day_name}</div>
                <div class="day-card-date">{date_str}</div>
                
                <div class="day-card-label" style="margin-top:10px;">SORTING</div>
                <div class="day-card-volume">{row['Sort_Count']:,.0f}</div>
                
                <div class="day-card-label" style="margin-top:15px;border-top:1px solid #ddd;padding-top:15px;">PUTAWAY</div>
                <div class="day-card-volume">{row['Putaway_Count']:,.0f}</div>
            </div>'''
        
        html += '</div></div>'
    else:
        # For large ranges, show a weekly summary table instead
        html += '<div class="section"><h2>Volume Summary by Week</h2>'
        daily_totals['Week'] = pd.to_datetime(daily_totals['Operation_Date']).dt.isocalendar().week
        daily_totals['Year'] = pd.to_datetime(daily_totals['Operation_Date']).dt.year
        weekly_agg = daily_totals.groupby(['Year', 'Week']).agg({
            'Operation_Date': ['min', 'max'],
            'Sort_Count': 'sum',
            'Putaway_Count': 'sum'
        }).reset_index()
        weekly_agg.columns = ['Year', 'Week', 'Start', 'End', 'Sort_Count', 'Putaway_Count']
        weekly_agg = weekly_agg.sort_values(['Year', 'Week'])
        
        html += '<table><tr><th>Week</th><th>Period</th><th>Sorting</th><th>Putaway</th><th>Total</th></tr>'
        for _, wrow in weekly_agg.iterrows():
            start_d = pd.to_datetime(wrow['Start']).strftime('%m/%d')
            end_d = pd.to_datetime(wrow['End']).strftime('%m/%d')
            total = wrow['Sort_Count'] + wrow['Putaway_Count']
            html += f'<tr><td>W{int(wrow["Week"])}</td><td>{start_d} - {end_d}</td><td>{wrow["Sort_Count"]:,.0f}</td><td>{wrow["Putaway_Count"]:,.0f}</td><td><strong>{total:,.0f}</strong></td></tr>'
        html += '</table></div>'

    # Operator/Shift Performance Tables
    if shift_breakdown and len(shift_breakdown) > 0:
        html += '<div class="section"><h2>Operator/Shift Performance</h2>'

        # Build date list from actual data dates
        date_list = sorted(shift_breakdown.keys())
        
        # Determine table mode: daily columns (<=7 days) or summary only (>7 days)
        show_daily_columns = (len(date_list) <= 7)
        
        # Helper: build operator data from shift_breakdown for a given shift and metric
        def build_operator_data(shift_key, metric_type):
            operators = {}
            if metric_type == 'sort':
                count_col, eff_col, speed_col = 'Sort_Count', 'Sort_Efficiency', 'Sort_Speed'
                m1_col, s1_col = 'Sort_M1_Count', 'Sort_S1_Count'
            else:
                count_col, eff_col, speed_col = 'Putaway_Count', 'Putaway_Efficiency', 'Putaway_Speed'
                m1_col, s1_col = 'Putaway_M1_Count', 'Putaway_S1_Count'

            for date in date_list:
                if date not in shift_breakdown:
                    continue
                shift_data = shift_breakdown[date][shift_key]
                ops = shift_data[shift_data[count_col] > 0]
                for _, row in ops.iterrows():
                    op_full = row['Operator_Full']
                    if op_full not in operators:
                        operators[op_full] = {
                            'daily': {}, 'total': 0,
                            'efficiencies': [], 'speeds': [], 'days_worked': 0,
                            'm1_total': 0, 's1_total': 0,
                        }
                    operators[op_full]['daily'][date] = int(row[count_col])
                    operators[op_full]['total'] += int(row[count_col])
                    operators[op_full]['m1_total'] += int(row.get(m1_col, 0) or 0)
                    operators[op_full]['s1_total'] += int(row.get(s1_col, 0) or 0)
                    operators[op_full]['efficiencies'].append(row[eff_col])
                    operators[op_full]['days_worked'] += 1
                    if row[speed_col] > 0:
                        operators[op_full]['speeds'].append(row[speed_col])
            return sorted(operators.items(), key=lambda x: x[1]['total'], reverse=True)
        
        # Helper: render a shift table
        def render_shift_table(shift_label, op_type, sorted_ops, is_sort=True):
            tbl = f'<div class="shift-table"><div class="shift-header">{shift_label} - {op_type}</div>'

            # Peer-median per-day count for cell heat-mapping (replaces fixed goal)
            all_daily_counts = [c for _, d in sorted_ops for c in d['daily'].values() if c > 0]
            if len(all_daily_counts) >= 3:
                cell_threshold = float(np.median(all_daily_counts))
            else:
                cell_threshold = float(np.mean(all_daily_counts)) if all_daily_counts else 0.0

            # Table header — adds M1 (hand-carry) and S1 (forklift) breakdown columns
            tbl += '<table><tr><th style="width:18%;">Operator</th>'
            if show_daily_columns:
                for date in date_list:
                    day_abbr = pd.to_datetime(date).strftime('%a')
                    tbl += f'<th>{day_abbr}</th>'
            else:
                tbl += '<th>Days</th><th>Avg/Day</th>'
            tbl += '<th title="Boxes carried by hand">M1</th>'
            tbl += '<th title="Pallets handled by forklift">S1</th>'
            tbl += '<th>Total</th><th>Efficiency</th><th style="width:12%;">Speed</th></tr>'

            daily_totals = {date: 0 for date in date_list}
            grand_total = 0
            grand_m1 = 0
            grand_s1 = 0
            all_effs = []
            all_speeds = []

            speed_thresholds = (55, 75) if is_sort else (25, 35)

            for op_full, data in sorted_ops:
                parts = str(op_full).split('_')
                display_name = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]
                avg_eff = np.mean(data['efficiencies']) if data['efficiencies'] else 0
                avg_speed = np.mean(data['speeds']) if data['speeds'] else 0
                eff_class = 'perf-above' if avg_eff > 80 else 'perf-at' if avg_eff >= 65 else 'perf-under'
                speed_class = 'speed-fast' if 0 < avg_speed < speed_thresholds[0] else 'speed-average' if speed_thresholds[0] <= avg_speed < speed_thresholds[1] else 'speed-slow' if avg_speed >= speed_thresholds[1] else ''

                tbl += f'<tr><td style="text-align:left;"><strong>{display_name}</strong></td>'

                if show_daily_columns:
                    for date in date_list:
                        if date in data['daily']:
                            count = data['daily'][date]
                            daily_totals[date] += count
                            if cell_threshold > 0:
                                cell_class = 'perf-above' if count >= cell_threshold * 1.1 else 'perf-at' if count >= cell_threshold * 0.9 else 'perf-under'
                            else:
                                cell_class = ''
                            tbl += f'<td class="{cell_class}"><strong>{count}</strong></td>'
                        else:
                            tbl += '<td>-</td>'
                else:
                    avg_per_day = data['total'] / data['days_worked'] if data['days_worked'] > 0 else 0
                    tbl += f'<td>{data["days_worked"]}</td><td>{avg_per_day:,.0f}</td>'

                grand_total += data['total']
                grand_m1 += data.get('m1_total', 0)
                grand_s1 += data.get('s1_total', 0)
                all_effs.extend(data['efficiencies'])
                all_speeds.extend(data['speeds'])
                speed_display = f'{avg_speed:.0f} sec' if avg_speed > 0 else '-'
                tbl += f'<td>{data.get("m1_total", 0):,.0f}</td>'
                tbl += f'<td>{data.get("s1_total", 0):,.0f}</td>'
                tbl += f'<td><strong>{data["total"]}</strong></td><td class="{eff_class}"><strong>{avg_eff:.1f}%</strong></td><td class="{speed_class}">{speed_display}</td></tr>'

            # Total row
            avg_eff_total = np.mean(all_effs) if all_effs else 0
            avg_speed_total = np.mean(all_speeds) if all_speeds else 0
            tbl += '<tr class="total-row"><td style="text-align:left;"><strong>TOTAL</strong></td>'
            if show_daily_columns:
                for date in date_list:
                    tbl += f'<td><strong>{daily_totals[date]}</strong></td>'
            else:
                total_days = len(date_list)
                avg_total_per_day = grand_total / total_days if total_days > 0 else 0
                tbl += f'<td><strong>{total_days}</strong></td><td><strong>{avg_total_per_day:,.0f}</strong></td>'
            tbl += f'<td><strong>{grand_m1:,.0f}</strong></td><td><strong>{grand_s1:,.0f}</strong></td>'
            tbl += f'<td><strong>{grand_total}</strong></td><td><strong>{avg_eff_total:.1f}%</strong></td><td><strong>{avg_speed_total:.0f} sec</strong></td></tr></table></div>'
            return tbl
        
        # ==========================
        # SORTING TABLES (1st & 2nd Shift)
        # ==========================
        html += '<h3 style="color: #2E75B6; margin: 20px 0 15px 0;">Sorting Operations</h3>'
        html += '<div class="shift-breakdown">'
        
        shift1_sort_ops = build_operator_data('Shift1', 'sort')
        shift2_sort_ops = build_operator_data('Shift2', 'sort')
        
        html += render_shift_table('1st Shift', 'Sorting', shift1_sort_ops, is_sort=True)
        html += render_shift_table('2nd Shift', 'Sorting', shift2_sort_ops, is_sort=True)
        
        html += '</div>'  # Close sorting shift-breakdown
        
        # ==========================
        # PUTAWAY TABLES (1st & 2nd Shift)
        # ==========================
        html += '<h3 style="color: #228b22; margin: 30px 0 15px 0;">Putaway Operations</h3>'
        html += '<div class="shift-breakdown">'
        
        shift1_putaway_ops = build_operator_data('Shift1', 'putaway')
        shift2_putaway_ops = build_operator_data('Shift2', 'putaway')
        
        html += render_shift_table('1st Shift', 'Putaway', shift1_putaway_ops, is_sort=False)
        html += render_shift_table('2nd Shift', 'Putaway', shift2_putaway_ops, is_sort=False)
        
        html += '</div></div>'
    
    # Workforce Insights
    if workforce_insights and len(workforce_insights) > 0:
        html += '<div class="section"><div class="insights-box"><h3>💡 Workforce Optimization Insights</h3>'
        
        # Summary stats
        high_count = sum(1 for i in workforce_insights if i['priority'] == 'high')
        medium_count = sum(1 for i in workforce_insights if i['priority'] == 'medium')
        low_count = sum(1 for i in workforce_insights if i['priority'] == 'low')
        info_count = sum(1 for i in workforce_insights if i['priority'] == 'info')
        
        html += '<div style="display:flex; gap:15px; margin:10px 0 15px 0; flex-wrap:wrap;">'
        if high_count > 0:
            html += f'<span style="background:#dc3545;color:white;padding:4px 12px;border-radius:12px;font-size:0.85em;font-weight:600;">⚠️ {high_count} Critical</span>'
        if medium_count > 0:
            html += f'<span style="background:#f4b400;color:white;padding:4px 12px;border-radius:12px;font-size:0.85em;font-weight:600;">📈 {medium_count} Warning</span>'
        if low_count > 0:
            html += f'<span style="background:#3498db;color:white;padding:4px 12px;border-radius:12px;font-size:0.85em;font-weight:600;">📊 {low_count} Info</span>'
        if info_count > 0:
            html += f'<span style="background:#28a745;color:white;padding:4px 12px;border-radius:12px;font-size:0.85em;font-weight:600;">✅ {info_count} Best Practice</span>'
        html += '</div>'
        
        # Pick only the few best insights to display:
        #   - always include every 'high' priority item (critical, actionable)
        #   - then fill with 'medium' / 'low' / 'info', at most 1 per type for variety
        #   - hard cap at 5 items total so the email stays scannable
        MAX_DISPLAYED_INSIGHTS = 5
        displayed = [i for i in workforce_insights if i['priority'] == 'high']
        seen_types = {i['type'] for i in displayed}
        for prio in ('medium', 'low', 'info'):
            if len(displayed) >= MAX_DISPLAYED_INSIGHTS:
                break
            for insight in workforce_insights:
                if insight['priority'] != prio or insight['type'] in seen_types:
                    continue
                displayed.append(insight)
                seen_types.add(insight['type'])
                if len(displayed) >= MAX_DISPLAYED_INSIGHTS:
                    break
        displayed = displayed[:MAX_DISPLAYED_INSIGHTS]

        for insight in displayed:
            priority_class = f"priority-{insight['priority']}"
            badge_class = f"badge-{insight['priority']}"
            _icon_map = {
                'understaffing': '⚠️', 'volume_spike': '📈',
                'overstaffing': '📊', 'speed_outlier': '🐢', 'top_performer': '🏆',
                'best_practice': '✅', 'shift_gap': '🔀', 'speed_trend': '⏱️',
                'consistency': '📉', 'cross_training': '🔄',
            }
            icon = _icon_map.get(insight['type'], '💡')
            
            shift_display = insight['shift'].replace('Shift1', '1st Shift').replace('Shift2', '2nd Shift')
            date_display = pd.Timestamp(insight['date']).strftime('%a %m/%d') if 'date' in insight else insight.get('day_name', '')
            
            html += f'<div class="insight-item {priority_class}">'
            html += f'<div class="insight-header"><span>{icon}</span><span>{date_display} - {shift_display}</span><span class="insight-badge {badge_class}">{insight["priority"].upper()}</span></div>'
            html += f'<div style="font-size: 14px; margin: 8px 0;">{insight["message"]}</div>'
            html += f'<div class="insight-details" style="display:flex;gap:20px;flex-wrap:wrap;">'
            html += f'<span>📦 Volume: <strong>{insight["volume"]:,.0f}</strong></span>'
            html += f'<span>👥 Operators: <strong>{insight["operators"]}</strong></span>'
            html += f'<span>⚡ Efficiency: <strong>{insight["efficiency"]:.1f}%</strong></span>'
            if insight["operators"] > 0:
                html += f'<span>📊 Per Operator: <strong>{insight["volume"]/insight["operators"]:,.0f}</strong></span>'
            html += '</div></div>'
        
        html += '</div></div>'
    
    # Dashboard section (embedded, hidden by default)
    if dashboard_html:
        html += '<div class="section" style="text-align:center; border-bottom:none;">'
        html += '<button class="dashboard-toggle-btn" onclick="var d=document.getElementById(\'dashboard-panel\');if(d.classList.contains(\'visible\')){d.classList.remove(\'visible\');this.textContent=\'View Detailed Report\';}else{d.classList.add(\'visible\');this.textContent=\'Hide Detailed Report\';window.dispatchEvent(new Event(\'resize\'));}">View Detailed Report</button>'
        html += '</div>'
        html += f'<div id="dashboard-panel" class="dashboard-section">{dashboard_html}</div>'
    
    html += '<div class="footer"><p>2026 Example Logistics All rights reserved - Created by Viktor Berg </p>'
    html += f'<p>Report generated by the Sorting and Putaway KPI System | {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}</p></div>'
    if not fragment_only:
        html += '</div></body></html>'
    else:
        html += '</div>'
    
    return html

# ============================================================================
# PLOTLY VISUALIZATIONS
# ============================================================================

def create_interactive_dashboard(daily_metrics, operator_stats, daily_trend, report_start, report_end, predictions=None, target_comparison=None):
    """Create interactive Plotly dashboard and return as embeddable HTML div string"""
    if not PLOTLY_AVAILABLE:
        print("   ⚠️  Plotly not available, skipping interactive dashboard")
        return None
    
    print("\n[CREATING INTERACTIVE DASHBOARD]")
    print("-" * 70)
    
    week_data = daily_metrics[
        (daily_metrics['Operation_Date'] >= report_start) &
        (daily_metrics['Operation_Date'] <= report_end) &
        ((daily_metrics['Sort_Count'] > 0) | (daily_metrics['Putaway_Count'] > 0))
    ].copy()
    
    if len(week_data) == 0:
        print("   ⚠️  No data for dashboard")
        return None
    
    figures_html = []
    
    # Helper: convert pandas series/values to plain Python lists to avoid Plotly binary encoding
    def to_list(series):
        return [float(x) if isinstance(x, (int, float, np.integer, np.floating)) else x for x in series.tolist()]
    def to_str_list(series):
        return [str(x) for x in series.tolist()]
    
    # ── Chart 1: Volume by Shift (grouped bar) ──
    shift_vol = week_data.groupby('Shift').agg({
        'Sort_Count': 'sum',
        'Putaway_Count': 'sum',
        'Operator': 'nunique'
    }).reset_index()
    shift_vol['Shift_Label'] = shift_vol['Shift'].replace({'Shift1': '1st Shift', 'Shift2': '2nd Shift'})
    
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(x=to_str_list(shift_vol['Shift_Label']), y=to_list(shift_vol['Sort_Count']),
                          name='Sorting', marker_color='#3498db',
                          hovertemplate='%{x}<br>Sorting: %{y:,.0f} modules<extra></extra>'))
    fig1.add_trace(go.Bar(x=to_str_list(shift_vol['Shift_Label']), y=to_list(shift_vol['Putaway_Count']),
                          name='Putaway', marker_color='#2ecc71',
                          hovertemplate='%{x}<br>Putaway: %{y:,.0f} modules<extra></extra>'))
    fig1.update_layout(title='Volume by Shift', barmode='group', height=380,
                       template='plotly_white', legend=dict(orientation='h', y=-0.15),
                       yaxis_title='Modules', margin=dict(t=50, b=60))
    figures_html.append(fig1.to_html(full_html=False, include_plotlyjs=False, div_id='chart_shift_volume'))
    
    # ── Chart 2: Daily Volume Trend (stacked bar + efficiency line) ──
    daily_vol = week_data.groupby('Operation_Date').agg({
        'Sort_Count': 'sum',
        'Putaway_Count': 'sum'
    }).reset_index().sort_values('Operation_Date')
    daily_vol['Date_Label'] = pd.to_datetime(daily_vol['Operation_Date']).dt.strftime('%a %m/%d')
    # Compute efficiency only from records with actual sorting/putaway
    sort_only = week_data[week_data['Sort_Count'] > 0]
    if len(sort_only) > 0:
        daily_eff = sort_only.groupby('Operation_Date')['Sort_Efficiency'].mean().reset_index()
        daily_vol = daily_vol.merge(daily_eff, on='Operation_Date', how='left')
    else:
        daily_vol['Sort_Efficiency'] = 0.0
    daily_vol['Sort_Efficiency'] = daily_vol['Sort_Efficiency'].fillna(0)
    
    put_only = week_data[week_data['Putaway_Count'] > 0]
    if len(put_only) > 0:
        daily_put_eff = put_only.groupby('Operation_Date')['Putaway_Efficiency'].mean().reset_index()
        daily_vol = daily_vol.merge(daily_put_eff, on='Operation_Date', how='left')
    else:
        daily_vol['Putaway_Efficiency'] = 0.0
    daily_vol['Putaway_Efficiency'] = daily_vol['Putaway_Efficiency'].fillna(0)
    
    fig2 = make_subplots(specs=[[{"secondary_y": True}]])
    fig2.add_trace(go.Bar(x=to_str_list(daily_vol['Date_Label']), y=to_list(daily_vol['Sort_Count']),
                          name='Sorting', marker_color='#3498db',
                          hovertemplate='%{x}<br>Sorting: %{y:,.0f}<extra></extra>'), secondary_y=False)
    fig2.add_trace(go.Bar(x=to_str_list(daily_vol['Date_Label']), y=to_list(daily_vol['Putaway_Count']),
                          name='Putaway', marker_color='#2ecc71',
                          hovertemplate='%{x}<br>Putaway: %{y:,.0f}<extra></extra>'), secondary_y=False)
    # Only plot efficiency for days that have sorting data
    eff_dates = [d for d, e in zip(to_str_list(daily_vol['Date_Label']), to_list(daily_vol['Sort_Efficiency'])) if e > 0]
    eff_vals = [e for e in to_list(daily_vol['Sort_Efficiency']) if e > 0]
    if eff_dates:
        fig2.add_trace(go.Scatter(x=eff_dates, y=eff_vals,
                                  name='Sort Efficiency %', mode='lines+markers',
                                  line=dict(color='#e74c3c', width=2), marker=dict(size=7),
                                  hovertemplate='%{x}<br>Sort Efficiency: %{y:.1f}%<extra></extra>'), secondary_y=True)
    # Only plot putaway efficiency for days that have putaway data
    put_eff_dates = [d for d, e in zip(to_str_list(daily_vol['Date_Label']), to_list(daily_vol['Putaway_Efficiency'])) if e > 0]
    put_eff_vals = [e for e in to_list(daily_vol['Putaway_Efficiency']) if e > 0]
    if put_eff_dates:
        fig2.add_trace(go.Scatter(x=put_eff_dates, y=put_eff_vals,
                                  name='Putaway Efficiency %', mode='lines+markers',
                                  line=dict(color='#f39c12', width=2, dash='dash'), marker=dict(size=7),
                                  hovertemplate='%{x}<br>Putaway Efficiency: %{y:.1f}%<extra></extra>'), secondary_y=True)
    fig2.update_layout(title='Daily Volume & Efficiency Trend', barmode='stack', height=380,
                       template='plotly_white', legend=dict(orientation='h', y=-0.15),
                       margin=dict(t=50, b=60))
    fig2.update_yaxes(title_text='Modules', secondary_y=False)
    fig2.update_yaxes(title_text='Efficiency %', secondary_y=True)
    figures_html.append(fig2.to_html(full_html=False, include_plotlyjs=False, div_id='chart_daily_trend'))
    
    # ── Chart 3: Top Operators by Volume (horizontal bar) ──
    op_data = week_data.groupby('Operator_Full').agg({
        'Sort_Count': 'sum',
        'Putaway_Count': 'sum'
    }).reset_index()
    op_data['Display_Name'] = op_data['Operator_Full'].apply(
        lambda x: '_'.join(str(x).split('_')[1:]) if '_' in str(x) else str(x))
    op_data['Total'] = op_data['Sort_Count'] + op_data['Putaway_Count']
    top_ops = op_data.nlargest(10, 'Total')
    top_ops = top_ops.sort_values('Total', ascending=True)
    
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(y=to_str_list(top_ops['Display_Name']), x=to_list(top_ops['Sort_Count']),
                          name='Sorting', orientation='h', marker_color='#3498db',
                          hovertemplate='%{y}<br>Sorting: %{x:,.0f}<extra></extra>'))
    fig3.add_trace(go.Bar(y=to_str_list(top_ops['Display_Name']), x=to_list(top_ops['Putaway_Count']),
                          name='Putaway', orientation='h', marker_color='#2ecc71',
                          hovertemplate='%{y}<br>Putaway: %{x:,.0f}<extra></extra>'))
    fig3.update_layout(title='Top 10 Operators by Volume', barmode='stack', height=420,
                       template='plotly_white', legend=dict(orientation='h', y=-0.1),
                       xaxis_title='Modules', margin=dict(t=50, b=50, l=150))
    figures_html.append(fig3.to_html(full_html=False, include_plotlyjs=False, div_id='chart_top_operators'))
    
    # ── Chart 4: Efficiency Distribution (box plot by shift) ──
    sort_data = week_data[week_data['Sort_Count'] > 0].copy()
    sort_data['Shift_Label'] = sort_data['Shift'].replace({'Shift1': '1st Shift', 'Shift2': '2nd Shift'})
    
    fig4 = go.Figure()
    for shift_label in ['1st Shift', '2nd Shift']:
        shift_eff = sort_data[sort_data['Shift_Label'] == shift_label]['Sort_Efficiency']
        if len(shift_eff) > 0:
            fig4.add_trace(go.Box(y=to_list(shift_eff), name=shift_label, boxmean=True,
                                  marker_color='#3498db' if '1st' in shift_label else '#2ecc71'))
    fig4.update_layout(title='Sorting Efficiency Distribution by Shift', height=380,
                       template='plotly_white', yaxis_title='Efficiency %',
                       margin=dict(t=50, b=50))
    figures_html.append(fig4.to_html(full_html=False, include_plotlyjs=False, div_id='chart_efficiency_dist'))
    
    # ── Chart 5: Speed vs Efficiency Scatter ──
    scatter_data = week_data[(week_data['Sort_Count'] > 0) & (week_data['Sort_Speed'] > 0)].copy()
    scatter_data['Shift_Label'] = scatter_data['Shift'].replace({'Shift1': '1st Shift', 'Shift2': '2nd Shift'})
    scatter_data['Display_Name'] = scatter_data['Operator_Full'].apply(
        lambda x: '_'.join(str(x).split('_')[1:]) if '_' in str(x) else str(x))
    
    fig5 = go.Figure()
    colors = {'1st Shift': '#3498db', '2nd Shift': '#2ecc71'}
    for shift_label in ['1st Shift', '2nd Shift']:
        sd = scatter_data[scatter_data['Shift_Label'] == shift_label]
        if len(sd) > 0:
            max_sort = sd['Sort_Count'].max()
            bubble_sizes = [float(v)/float(max_sort)*25+5 for v in sd['Sort_Count'].tolist()]
            fig5.add_trace(go.Scatter(
                x=to_list(sd['Sort_Speed']), y=to_list(sd['Sort_Efficiency']),
                mode='markers', name=shift_label,
                marker=dict(color=colors.get(shift_label, '#999'), size=bubble_sizes,
                            opacity=0.7, line=dict(width=1, color='white')),
                text=to_str_list(sd['Display_Name']),
                hovertemplate='%{text}<br>Speed: %{x:.1f} sec<br>Efficiency: %{y:.1f}%<extra></extra>'
            ))
    fig5.update_layout(title='Speed vs Efficiency (bubble = volume)', height=380,
                       template='plotly_white', xaxis_title='Avg Speed (sec/module)',
                       yaxis_title='Efficiency %', legend=dict(orientation='h', y=-0.15),
                       margin=dict(t=50, b=60))
    figures_html.append(fig5.to_html(full_html=False, include_plotlyjs=False, div_id='chart_speed_efficiency'))
    
    # ── Chart 6: Workforce Utilization (operators vs volume per day) ──
    workforce = week_data.groupby(['Operation_Date', 'Shift']).agg({
        'Sort_Count': 'sum',
        'Putaway_Count': 'sum',
        'Operator': 'nunique'
    }).reset_index()
    workforce['Shift_Label'] = workforce['Shift'].replace({'Shift1': '1st Shift', 'Shift2': '2nd Shift'})
    workforce['Total_Modules'] = workforce['Sort_Count'] + workforce['Putaway_Count']
    workforce['Modules_Per_Op'] = workforce['Total_Modules'] / workforce['Operator']
    workforce['Date_Label'] = pd.to_datetime(workforce['Operation_Date']).dt.strftime('%a %m/%d')
    
    fig6 = make_subplots(specs=[[{"secondary_y": True}]])
    for shift_label, color in [('1st Shift', '#3498db'), ('2nd Shift', '#2ecc71')]:
        wf = workforce[workforce['Shift_Label'] == shift_label].sort_values('Operation_Date')
        if len(wf) > 0:
            fig6.add_trace(go.Bar(x=to_str_list(wf['Date_Label']), y=to_list(wf['Operator']),
                                  name=f'{shift_label} Operators',
                                  marker_color=color, opacity=0.6,
                                  hovertemplate='%{x}<br>Operators: %{y}<extra></extra>'), secondary_y=False)
            fig6.add_trace(go.Scatter(x=to_str_list(wf['Date_Label']), y=to_list(wf['Modules_Per_Op']),
                                      name=f'{shift_label} Modules/Op',
                                      mode='lines+markers', line=dict(color=color, width=2),
                                      marker=dict(size=7),
                                      hovertemplate='%{x}<br>Modules/Operator: %{y:.0f}<extra></extra>'), secondary_y=True)
    fig6.update_layout(title='Workforce: Headcount vs Productivity', barmode='group', height=380,
                       template='plotly_white', legend=dict(orientation='h', y=-0.2),
                       margin=dict(t=50, b=70))
    fig6.update_yaxes(title_text='Operators', secondary_y=False)
    fig6.update_yaxes(title_text='Modules per Operator', secondary_y=True)
    figures_html.append(fig6.to_html(full_html=False, include_plotlyjs=False, div_id='chart_workforce'))
    
    # Assemble the dashboard HTML div
    dashboard_html = '<div class="dashboard-grid">'
    for i, chart_html in enumerate(figures_html):
        span_class = ' full-width' if i == 2 else ''
        dashboard_html += f'<div class="dashboard-chart{span_class}">{chart_html}</div>'
    dashboard_html += '</div>'
    
    print(f"   ✅ Created interactive dashboard with {len(figures_html)} charts")
    
    return dashboard_html

def create_forecast_visualization(predictions, daily_trend):
    """Create ML forecast visualization"""
    if not PLOTLY_AVAILABLE or predictions is None:
        return None
    
    print("\n[CREATING FORECAST VISUALIZATION]")
    print("-" * 70)
    
    # Create figure
    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=('Volume Forecast', 'Efficiency Forecast'),
        specs=[[{"secondary_y": False}, {"secondary_y": False}]]
    )
    
    if len(daily_trend) > 0:
        # Historical data
        historical = daily_trend.tail(7).copy()
        historical['Date_Str'] = pd.to_datetime(historical['Date']).dt.strftime('%a %m/%d')
        
        # Add historical volume
        fig.add_trace(
            go.Scatter(
                x=historical['Date_Str'],
                y=historical['Sort_Volume'],
                mode='lines+markers',
                name='Historical',
                line=dict(color='#2E75B6', width=2),
                marker=dict(size=8)
            ),
            row=1, col=1
        )
        
        # Add forecast
        if 'Prediction_Date' in predictions.columns:
            forecast_date = predictions['Prediction_Date'].iloc[0]
            forecast_str = pd.Timestamp(forecast_date).strftime('%a %m/%d')
            
            total_forecast = predictions[predictions['Shift'] == 'Total']['Predicted_Sort_Volume'].values[0]
            
            fig.add_trace(
                go.Scatter(
                    x=[historical['Date_Str'].iloc[-1], forecast_str],
                    y=[historical['Sort_Volume'].iloc[-1], total_forecast],
                    mode='lines+markers',
                    name='Forecast',
                    line=dict(color='#FF9800', width=2, dash='dash'),
                    marker=dict(size=10, symbol='star')
                ),
                row=1, col=1
            )
            
            # Add efficiency forecast
            fig.add_trace(
                go.Scatter(
                    x=historical['Date_Str'],
                    y=historical['Sort_Efficiency'],
                    mode='lines+markers',
                    name='Historical',
                    line=dict(color='#2E75B6', width=2),
                    marker=dict(size=8)
                ),
                row=1, col=2
            )
            
            total_eff_forecast = predictions[predictions['Shift'] == 'Total']['Predicted_Efficiency'].values[0]
            
            fig.add_trace(
                go.Scatter(
                    x=[historical['Date_Str'].iloc[-1], forecast_str],
                    y=[historical['Sort_Efficiency'].iloc[-1], total_eff_forecast],
                    mode='lines+markers',
                    name='Forecast',
                    line=dict(color='#FF9800', width=2, dash='dash'),
                    marker=dict(size=10, symbol='star')
                ),
                row=1, col=2
            )
    
    fig.update_layout(
        title='🤖 ML-Based Performance Forecast (Random Forest)',
        showlegend=True,
        height=400,
        template='plotly_white'
    )
    
    fig.update_xaxes(title_text="Date", row=1, col=1)
    fig.update_yaxes(title_text="Modules", row=1, col=1)
    
    fig.update_xaxes(title_text="Date", row=1, col=2)
    fig.update_yaxes(title_text="Efficiency (%)", row=1, col=2)
    
    print(f"   ✅ Created forecast visualization")
    
    return fig

# ============================================================================
# MAIN
# ============================================================================

def generate_daily_html(daily_metrics, operator_stats, report_day, fragment_only=False):
    """Generate HTML report for a single day. If fragment_only=True, return only the inner content div."""
    if len(daily_metrics) == 0:
        if fragment_only:
            return '<div class="container"><h1>No data available</h1></div>'
        return "<html><body><h1>No data available</h1></body></html>"
    
    day_name = pd.Timestamp(report_day).strftime('%A')
    date_display = pd.Timestamp(report_day).strftime('%B %d, %Y')
    
    day_data = daily_metrics[daily_metrics['Operation_Date'] == report_day]
    
    if len(day_data) == 0:
        if fragment_only:
            return f'<div class="container"><h1>No data available for {date_display}</h1></div>'
        return f"<html><body><h1>No data available for {date_display}</h1></body></html>"
    
    # Calculate daily totals
    day_sort = day_data['Sort_Count'].sum()
    day_putaway = day_data['Putaway_Count'].sum()
    day_sort_efficiency = day_data[day_data['Sort_Count'] > 0]['Sort_Efficiency'].mean() if len(day_data[day_data['Sort_Count'] > 0]) > 0 else 0
    day_putaway_efficiency = day_data[day_data['Putaway_Count'] > 0]['Putaway_Efficiency'].mean() if len(day_data[day_data['Putaway_Count'] > 0]) > 0 else 0
    
    # Get shift breakdowns
    shift1_data = day_data[day_data['Shift'] == 'Shift1']
    shift2_data = day_data[day_data['Shift'] == 'Shift2']
    
    shift1_sort = shift1_data['Sort_Count'].sum()
    shift1_putaway = shift1_data['Putaway_Count'].sum()
    shift1_sort_ops = shift1_data[shift1_data['Sort_Count'] > 0]['Operator'].nunique()
    shift1_put_ops = shift1_data[shift1_data['Putaway_Count'] > 0]['Operator'].nunique()
    shift1_sort_eff = shift1_data[shift1_data['Sort_Count'] > 0]['Sort_Efficiency'].mean() if len(shift1_data[shift1_data['Sort_Count'] > 0]) > 0 else 0
    shift1_put_eff = shift1_data[shift1_data['Putaway_Count'] > 0]['Putaway_Efficiency'].mean() if len(shift1_data[shift1_data['Putaway_Count'] > 0]) > 0 else 0
    
    shift2_sort = shift2_data['Sort_Count'].sum()
    shift2_putaway = shift2_data['Putaway_Count'].sum()
    shift2_sort_ops = shift2_data[shift2_data['Sort_Count'] > 0]['Operator'].nunique()
    shift2_put_ops = shift2_data[shift2_data['Putaway_Count'] > 0]['Operator'].nunique()
    shift2_sort_eff = shift2_data[shift2_data['Sort_Count'] > 0]['Sort_Efficiency'].mean() if len(shift2_data[shift2_data['Sort_Count'] > 0]) > 0 else 0
    shift2_put_eff = shift2_data[shift2_data['Putaway_Count'] > 0]['Putaway_Efficiency'].mean() if len(shift2_data[shift2_data['Putaway_Count'] > 0]) > 0 else 0
    
    if not fragment_only:
        html = f"""
    <html>
    <head>
        <style>
            * {{ margin: 0; padding: 0; box-sizing: border-box; }}
            body {{ font-family: Arial, Helvetica, sans-serif; background: #ffffff; padding: 0; color: #2c3e50; }}
            .container {{ max-width: 100%; margin: 0; background: white; box-shadow: none; overflow: hidden; }}
            .header {{ background: transparent; color: #2f3b45; padding: 18px 30px; text-align: center; border-bottom: none; }}
            .header h1 {{ margin: 6px 0 0 0; font-size: 1.6em; font-weight: 600; letter-spacing: 0.6px; line-height: 1.05; }}
            .header p {{ margin: 0; font-size: 0.9em; font-weight: 600; color: #4f5b66; line-height: 1.05; }}
            .section {{ padding: 20px 30px; background: white; border-bottom: 1px solid #dee2e6; }}
            .section h2 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 8px; margin-top: 0; font-size: 1.1em; letter-spacing: 0.4px; }}
            .metrics-grid {{ display: grid; grid-template-columns: 0.15fr 0.35fr 0.15fr 0.35fr; gap: 20px; margin: 20px 0 0 0; }}
            .metric-box {{ background: white; color: #2c3e50; padding: 16px 20px; border-top: 4px solid #95a5a6; box-shadow: 0 4px 6px rgba(0,0,0,0.1); min-height: 90px; }}
            .metric-box.primary {{ border-top-color: #3498db; }}
            .metric-box.success {{ border-top-color: #28a745; }}
            .metric-box.warning {{ border-top-color: #f4b400; }}
            .metric-box.info {{ border-top-color: #7f8c8d; }}
            .metric-value {{ font-size: 2.2em; font-weight: bold; margin: 8px 0 4px 0; line-height: 1; }}
            .metric-label {{ font-size: 0.85em; color: #7f8c8d; text-transform: uppercase; letter-spacing: 0.6px; font-weight: 600; }}
            .metric-sublabel {{ font-size: 0.95em; color: #4f5b66; margin-top: 4px; font-weight: 500; }}
            .top-performers {{ background-color: #f8f9fa; padding: 8px; border-radius: 4px; margin: 8px 0 0 0; }}
            .top-performers-item {{ padding: 4px 0; border-bottom: 1px solid #dee2e6; font-size: 1.05em; color: #2c3e50; font-weight: 600; }}
            .top-performers-item:last-child {{ border-bottom: none; }}
            .shift-comparison {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0 0 0; }}
            .shift-box {{ background: white; border-radius: 4px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 4px solid #3498db; }}
            .shift-box.shift1 {{ border-left-color: #f4b400; }}
            .shift-box.shift2 {{ border-left-color: #6f42c1; }}
            .shift-header-label {{ font-size: 1.1em; font-weight: 600; color: #2c3e50; margin-bottom: 12px; }}
            .shift-stats {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-top: 10px; }}
            .stat-item {{ text-align: center; padding: 10px; background: #f8f9fa; border-radius: 4px; }}
            .stat-value {{ font-size: 1.6em; font-weight: bold; color: #2c3e50; }}
            .stat-label {{ font-size: 0.85em; color: #6c757d; margin-top: 5px; }}
            table {{ border-collapse: collapse; width: 100%; margin: 20px 0 0 0; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            th {{ background-color: #95a5a6; color: white; padding: 12px; text-align: left; font-weight: 600; font-size: 0.85em; text-transform: uppercase; letter-spacing: 0.5px; }}
            td {{ padding: 12px; border-bottom: 1px solid #ecf0f1; }}
            tr:hover {{ background-color: #f8f9fa; }}
            .efficiency-high {{ color: #28a745; font-weight: bold; }}
            .efficiency-medium {{ color: #f4b400; font-weight: bold; }}
            .efficiency-low {{ color: #dc3545; font-weight: bold; }}
            .perf-above {{ background-color: #d4edda; color: #155724; font-weight: bold; }}
            .perf-at {{ background-color: #fff3cd; color: #856404; font-weight: bold; }}
            .perf-under {{ background-color: #f8d7da; color: #721c24; font-weight: bold; }}
            .footer {{ background: transparent; text-align: center; padding: 20px; color: #4f5b66; font-size: 0.85em; border-top: none; }}
        </style>
    </head>
    <body>"""
    else:
        html = ""

    html += f"""
        <div class="container">
            <div class="header">
                <h1>Daily Sorting and Putaway KPI Report</h1>
                <p>{day_name}, {date_display}</p>
            </div>
        
        <div class="section">
            <h2>Daily Overview</h2>
            <div class="metrics-grid">
                <div class="metric-box primary">
                    <div class="metric-label">SORTING</div>
                    <div class="metric-value" style="font-size: 42px;">{day_sort:,.0f}</div>
                    <div class="metric-sublabel">Modules Sorted</div>
                </div>
                <div class="metric-box info">
                    <div class="metric-label">Top Sorting</div>
                    <div class="top-performers">
    """
    
    # Top sorting performers by shift
    for shift_label, shift_name in [('Shift1', '1st'), ('Shift2', '2nd')]:
        top_sort = operator_stats[
            (operator_stats['Shift'] == shift_label) & (operator_stats['Day_Sort_Count'] > 0)
        ].nlargest(3, 'Day_Sort_Efficiency')
        if len(top_sort) > 0:
            html += f'<div style="font-size:0.8em;color:#7f8c8d;margin-top:4px;">{shift_name} Shift:</div>'
            for rank, (idx, row) in enumerate(top_sort.iterrows(), start=1):
                op_name = row['Operator_Full']
                parts = str(op_name).split('_')
                op_name = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]
                html += f'<div class="top-performers-item">{rank}. {op_name}: {row["Day_Sort_Efficiency"]:.1f}%</div>'
    
    html += f"""
                    </div>
                </div>
                <div class="metric-box success">
                    <div class="metric-label">PUTAWAY</div>
                    <div class="metric-value" style="font-size: 42px;">{day_putaway:,.0f}</div>
                    <div class="metric-sublabel">Modules Put Away</div>
                </div>
                <div class="metric-box info">
                    <div class="metric-label">Top Putaway</div>
                    <div class="top-performers">
    """
    
    # Top putaway performers by shift
    for shift_label, shift_name in [('Shift1', '1st'), ('Shift2', '2nd')]:
        top_put = operator_stats[
            (operator_stats['Shift'] == shift_label) & (operator_stats['Day_Putaway_Count'] > 0)
        ].nlargest(3, 'Day_Putaway_Efficiency')
        if len(top_put) > 0:
            html += f'<div style="font-size:0.8em;color:#7f8c8d;margin-top:4px;">{shift_name} Shift:</div>'
            for rank, (idx, row) in enumerate(top_put.iterrows(), start=1):
                op_name = row['Operator_Full']
                parts = str(op_name).split('_')
                op_name = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]
                html += f'<div class="top-performers-item">{rank}. {op_name}: {row["Day_Putaway_Efficiency"]:.1f}%</div>'
    
    html += """
                    </div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h2>Shift Performance Comparison</h2>
            <div class="shift-comparison">
    """
    
    html += f"""
                <div class="shift-box shift1">
                    <div class="shift-header-label" style="color: #FFA000;">1st Shift</div>
                    <div class="shift-stats">
                        <div class="stat-item"><div class="stat-value">{shift1_sort:,.0f}</div><div class="stat-label">Sorting</div></div>
                        <div class="stat-item"><div class="stat-value">{shift1_putaway:,.0f}</div><div class="stat-label">Putaway</div></div>
                        <div class="stat-item"><div class="stat-value">{shift1_sort_eff:.0f}%</div><div class="stat-label">Sort Eff.</div></div>
                        <div class="stat-item"><div class="stat-value">{shift1_put_eff:.0f}%</div><div class="stat-label">Putaway Eff.</div></div>
                        <div class="stat-item"><div class="stat-value">{shift1_sort_ops}</div><div class="stat-label">Sort Ops</div></div>
                        <div class="stat-item"><div class="stat-value">{shift1_put_ops}</div><div class="stat-label">Putaway Ops</div></div>
                    </div>
                </div>
                <div class="shift-box shift2">
                    <div class="shift-header-label" style="color: #7B1FA2;">2nd Shift</div>
                    <div class="shift-stats">
                        <div class="stat-item"><div class="stat-value">{shift2_sort:,.0f}</div><div class="stat-label">Sorting</div></div>
                        <div class="stat-item"><div class="stat-value">{shift2_putaway:,.0f}</div><div class="stat-label">Putaway</div></div>
                        <div class="stat-item"><div class="stat-value">{shift2_sort_eff:.0f}%</div><div class="stat-label">Sort Eff.</div></div>
                        <div class="stat-item"><div class="stat-value">{shift2_put_eff:.0f}%</div><div class="stat-label">Putaway Eff.</div></div>
                        <div class="stat-item"><div class="stat-value">{shift2_sort_ops}</div><div class="stat-label">Sort Ops</div></div>
                        <div class="stat-item"><div class="stat-value">{shift2_put_ops}</div><div class="stat-label">Putaway Ops</div></div>
                    </div>
                </div>
            </div>
        </div>
    """
    
    # Detailed operator tables per shift
    html += '<div class="section"><h2>Operator Performance Details</h2>'
    
    for shift_label, shift_name, color in [('Shift1', '1st Shift', '#FFA000'), ('Shift2', '2nd Shift', '#7B1FA2')]:
        for op_type, count_col, eff_col, speed_col, m1_col, s1_col in [
            ('Sorting', 'Day_Sort_Count', 'Day_Sort_Efficiency', 'Day_Sort_Speed', 'Day_Sort_M1', 'Day_Sort_S1'),
            ('Putaway', 'Day_Putaway_Count', 'Day_Putaway_Efficiency', 'Day_Putaway_Speed', 'Day_Putaway_M1', 'Day_Putaway_S1')
        ]:
            html += f'<h3 style="color: {color}; margin: 20px 0 10px 0;">{shift_name} - {op_type}</h3>'
            # M1 = boxes carried by hand (fast); S1 = pallets handled by forklift (slow).
            html += ('<table><tr><th>Operator</th><th>Modules</th>'
                     '<th title="Boxes carried by hand">M1</th>'
                     '<th title="Pallets handled by forklift">S1</th>'
                     '<th>Efficiency</th><th>Speed (sec/module)</th><th>Rank</th></tr>')

            shift_ops = operator_stats[
                (operator_stats['Shift'] == shift_label) & (operator_stats[count_col] > 0)
            ].sort_values(eff_col, ascending=False)

            shift_ops = shift_ops.copy()
            shift_ops['_rank'] = range(1, len(shift_ops) + 1)

            # Peer-median count drives cell heat-map thresholds
            cell_threshold = float(shift_ops[count_col].median()) if len(shift_ops) >= 3 else 0.0

            total_m1 = float(shift_ops[m1_col].sum()) if m1_col in shift_ops.columns else 0
            total_s1 = float(shift_ops[s1_col].sum()) if s1_col in shift_ops.columns else 0
            total_modules = float(shift_ops[count_col].sum())

            for idx, row in shift_ops.iterrows():
                op_name = row['Operator_Full']
                parts = str(op_name).split('_')
                display_name = '_'.join(parts[1:]) if len(parts) > 1 else parts[0]

                eff_class = 'efficiency-high' if row[eff_col] > 80 else 'efficiency-medium' if row[eff_col] >= 65 else 'efficiency-low'
                if cell_threshold > 0:
                    count_class = 'perf-above' if row[count_col] >= cell_threshold * 1.1 else 'perf-at' if row[count_col] >= cell_threshold * 0.9 else 'perf-under'
                else:
                    count_class = ''

                m1_val = float(row.get(m1_col, 0) or 0)
                s1_val = float(row.get(s1_col, 0) or 0)
                html += (f'<tr><td><strong>{display_name}</strong></td>'
                         f'<td class="{count_class}">{row[count_col]:.0f}</td>'
                         f'<td>{m1_val:,.0f}</td>'
                         f'<td>{s1_val:,.0f}</td>'
                         f'<td class="{eff_class}">{row[eff_col]:.1f}%</td>'
                         f'<td>{row[speed_col]:.1f}</td>'
                         f'<td>#{int(row["_rank"])}</td></tr>')

            if len(shift_ops) == 0:
                html += '<tr><td colspan="7" style="text-align:center;">No data available</td></tr>'
            else:
                html += (f'<tr class="total-row" style="background-color:#f0f0f0;font-weight:bold;">'
                         f'<td><strong>TOTAL</strong></td><td>{total_modules:,.0f}</td>'
                         f'<td>{total_m1:,.0f}</td><td>{total_s1:,.0f}</td>'
                         f'<td colspan="3"></td></tr>')

            html += '</table>'
    
    html += '</div>'
    
    html += f"""
        <div class="footer">
            <p>2026 Example Logistics All rights reserved - Created by Viktor Berg</p>
            <p>Report Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
        </div>"""
    if not fragment_only:
        html += """
    </body>
    </html>
    """
    
    return html

# ============================================================================
# INTERACTIVE REPORT WITH DATE PICKER
# ============================================================================

def generate_interactive_report(daily_metrics, sorting_metrics, putaway_metrics, benchmarks, output_path):
    """
    Generate a single self-contained interactive HTML report with tabbed date ranges
    (Yesterday, Last 7 Days, Last 30 Days) and a custom date picker.
    All data is embedded as JSON for client-side filtering.
    """
    print("\n[GENERATING INTERACTIVE REPORT]")
    print("-" * 70)

    today = pd.Timestamp.now().date()
    today_weekday = pd.Timestamp.now().weekday()

    # Yesterday (or last Friday if Monday)
    if today_weekday == 0:
        yesterday = today - pd.Timedelta(days=3)
    else:
        yesterday = today - pd.Timedelta(days=1)

    seven_days_ago = today - pd.Timedelta(days=7)
    thirty_days_ago = today - pd.Timedelta(days=30)

    # Define preset ranges
    ranges = {
        'last7': (seven_days_ago, today, "last_7_days"),
        'last30': (thirty_days_ago, today, "last_30_days"),
    }

    fragments = {}
    for key, (r_start, r_end, r_label) in ranges.items():
        print(f"   Pre-rendering: {key} ({r_start} to {r_end})")
        dm = aggregate_daily_metrics(sorting_metrics, putaway_metrics, r_start, r_end, benchmarks)
        if len(dm) == 0:
            fragments[key] = '<div class="container"><h1>No data available for this range</h1></div>'
            continue

        is_single = (r_start == r_end)
        op_stats = analyze_operator_performance(dm, r_start, r_end)

        if is_single:
            frag = generate_daily_html(dm, op_stats, r_start, fragment_only=True)
        else:
            d_trend = calculate_daily_efficiency_trend(dm, r_start, r_end)
            wf_insights = analyze_workforce_optimization(dm, op_stats, r_start, r_end) if not is_single else []
            dash_html = None
            if PLOTLY_AVAILABLE and not is_single:
                dash_html = create_interactive_dashboard(dm, op_stats, d_trend, r_start, r_end)
            frag = generate_email_content(
                dm, op_stats, d_trend, r_start, r_end, r_label,
                wf_insights, None, None, None, dash_html, fragment_only=True
            )
        # Make dashboard panel and chart IDs unique per tab to avoid getElementById conflicts
        frag = frag.replace("dashboard-panel", f"dashboard-panel-{key}")
        for chart_id in ["chart_shift_volume", "chart_daily_trend", "chart_top_operators",
                         "chart_efficiency_dist", "chart_speed_efficiency", "chart_workforce"]:
            frag = frag.replace(chart_id, f"{chart_id}_{key}")
        # Strip fragment footer to avoid duplicates (the wrapper has its own footer)
        import re as _re
        frag = _re.sub(r'<div class="footer">.*?</div>\s*</div>', '</div>', frag, flags=_re.DOTALL)
        fragments[key] = frag

    # Build JSON dataset from ALL daily_metrics for custom range filtering
    all_dates = sorted(daily_metrics['Operation_Date'].unique())
    min_date = str(min(all_dates))
    max_date = str(max(all_dates))

    # Aggregate data per day for JSON embedding
    js_data = []
    for date in all_dates:
        day_data = daily_metrics[daily_metrics['Operation_Date'] == date]
        row_obj = {
            'date': str(date),
            'sort_count': int(day_data['Sort_Count'].sum()),
            'putaway_count': int(day_data['Putaway_Count'].sum()),
            'sort_efficiency': round(float(day_data[day_data['Sort_Count'] > 0]['Sort_Efficiency'].mean()) if len(day_data[day_data['Sort_Count'] > 0]) > 0 else 0, 1),
            'putaway_efficiency': round(float(day_data[day_data['Putaway_Count'] > 0]['Putaway_Efficiency'].mean()) if len(day_data[day_data['Putaway_Count'] > 0]) > 0 else 0, 1),
            'sort_speed': round(float(day_data[day_data['Sort_Speed'] > 0]['Sort_Speed'].mean()) if len(day_data[day_data['Sort_Speed'] > 0]) > 0 else 0, 1),
            'putaway_speed': round(float(day_data[day_data['Putaway_Speed'] > 0]['Putaway_Speed'].mean()) if len(day_data[day_data['Putaway_Speed'] > 0]) > 0 else 0, 1),
            'operators': int(day_data['Operator'].nunique()),
            'shift1_sort': int(day_data[day_data['Shift'] == 'Shift1']['Sort_Count'].sum()),
            'shift1_putaway': int(day_data[day_data['Shift'] == 'Shift1']['Putaway_Count'].sum()),
            'shift2_sort': int(day_data[day_data['Shift'] == 'Shift2']['Sort_Count'].sum()),
            'shift2_putaway': int(day_data[day_data['Shift'] == 'Shift2']['Putaway_Count'].sum()),
        }
        js_data.append(row_obj)

    all_data_json = json.dumps(js_data)

    # Build operator-level data for custom range top performers
    op_rows = []
    for _, row in daily_metrics.iterrows():
        op_rows.append({
            'date': str(row['Operation_Date']),
            'operator': str(row.get('Operator_Full', row.get('Operator', ''))),
            'shift': str(row.get('Shift', '')),
            'sort_count': int(row['Sort_Count']) if pd.notna(row['Sort_Count']) else 0,
            'putaway_count': int(row['Putaway_Count']) if pd.notna(row['Putaway_Count']) else 0,
            'sort_efficiency': round(float(row['Sort_Efficiency']), 1) if pd.notna(row['Sort_Efficiency']) else 0,
            'putaway_efficiency': round(float(row['Putaway_Efficiency']), 1) if pd.notna(row['Putaway_Efficiency']) else 0,
            'sort_speed': round(float(row['Sort_Speed']), 1) if pd.notna(row['Sort_Speed']) else 0,
            'putaway_speed': round(float(row['Putaway_Speed']), 1) if pd.notna(row['Putaway_Speed']) else 0,
        })
    op_data_json = json.dumps(op_rows)

    # Build the full interactive HTML
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Interactive Sorting & Putaway KPI Report</title>
    <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: Arial, Helvetica, sans-serif; background: #ffffff; padding: 0; color: #2c3e50; }}
        .container {{ max-width: 100%; margin: 0; background: white; box-shadow: none; overflow: hidden; }}
        .header {{ background: transparent; color: #2f3b45; padding: 18px 30px; text-align: center; border-bottom: none; }}
        .header h1 {{ margin: 6px 0 0 0; font-size: 1.6em; font-weight: 600; letter-spacing: 0.6px; line-height: 1.05; }}
        .header p {{ margin: 0; font-size: 0.9em; font-weight: 600; color: #4f5b66; line-height: 1.05; }}
        .section {{ padding: 20px 30px; background: white; border-bottom: 1px solid #dee2e6; }}
        .section h2 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 8px; margin-top: 0; font-size: 1.1em; letter-spacing: 0.4px; }}
        .metrics-grid {{ display: grid; grid-template-columns: 0.15fr 0.35fr 0.15fr 0.35fr; gap: 20px; margin: 20px 0 0 0; }}
        .metrics-grid-equal {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin: 20px 0 0 0; }}
        .metric-box {{ background: white; color: #2c3e50; padding: 16px 20px; border-top: 4px solid #95a5a6; box-shadow: 0 4px 6px rgba(0,0,0,0.1); min-height: 90px; }}
        .metric-box.primary {{ border-top-color: #3498db; }}
        .metric-box.success {{ border-top-color: #28a745; }}
        .metric-box.warning {{ border-top-color: #f4b400; }}
        .metric-box.info {{ border-top-color: #7f8c8d; }}
        .metric-value {{ font-size: 2.2em; font-weight: bold; margin: 8px 0 4px 0; line-height: 1; }}
        .metric-label {{ font-size: 0.85em; color: #7f8c8d; text-transform: uppercase; letter-spacing: 0.6px; font-weight: 600; }}
        .metric-sublabel {{ font-size: 0.95em; color: #4f5b66; margin-top: 4px; font-weight: 500; }}
        .top-performers {{ background-color: #f8f9fa; padding: 8px; border-radius: 4px; margin: 8px 0 0 0; }}
        .top-performers-item {{ padding: 4px 0; border-bottom: 1px solid #dee2e6; font-size: 1.05em; color: #2c3e50; font-weight: 600; }}
        .top-performers-item:last-child {{ border-bottom: none; }}
        .chart-container {{ padding: 30px; background: white; text-align: center; }}
        .chart-container > div {{ display: inline-block; text-align: left; }}
        table {{ border-collapse: collapse; width: 100%; margin: 20px 0 0 0; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
        th {{ background-color: #95a5a6; color: white; padding: 12px; text-align: left; font-weight: 600; font-size: 0.85em; text-transform: uppercase; letter-spacing: 0.5px; }}
        td {{ padding: 10px 12px; border-bottom: 1px solid #ecf0f1; }}
        tr:hover {{ background-color: #f8f9fa; }}
        .efficiency-high {{ color: #28a745; font-weight: bold; }}
        .efficiency-medium {{ color: #f4b400; font-weight: bold; }}
        .efficiency-low {{ color: #dc3545; font-weight: bold; }}
        .perf-above {{ background-color: #d4edda; color: #155724; font-weight: bold; }}
        .perf-at {{ background-color: #fff3cd; color: #856404; font-weight: bold; }}
        .perf-under {{ background-color: #f8d7da; color: #721c24; font-weight: bold; }}
        .shift-comparison {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0 0 0; }}
        .shift-box {{ background: white; border-radius: 4px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 4px solid #3498db; }}
        .shift-box.shift1 {{ border-left-color: #f4b400; }}
        .shift-box.shift2 {{ border-left-color: #6f42c1; }}
        .shift-header-label {{ font-size: 1.1em; font-weight: 600; color: #2c3e50; margin-bottom: 12px; }}
        .shift-stats {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-top: 10px; }}
        .stat-item {{ text-align: center; padding: 10px; background: #f8f9fa; border-radius: 4px; }}
        .stat-value {{ font-size: 1.6em; font-weight: bold; color: #2c3e50; }}
        .stat-label {{ font-size: 0.85em; color: #6c757d; margin-top: 5px; }}
        .day-card-container {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 15px; margin: 20px 0 0 0; }}
        .day-card {{ background: white; border-radius: 4px; padding: 20px; border-top: 4px solid #95a5a6; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; transition: transform 0.2s, box-shadow 0.2s; }}
        .day-card:hover {{ transform: translateY(-3px); box-shadow: 0 8px 15px rgba(0,0,0,0.2); }}
        .day-card-header {{ font-size: 1.1em; font-weight: 600; color: #2c3e50; margin-bottom: 4px; }}
        .day-card-date {{ font-size: 0.95em; color: #6c757d; margin-bottom: 12px; }}
        .day-card-volume {{ font-size: 1.1em; color: #4f5b66; margin-top: 8px; font-weight: 600; }}
        .day-card-label {{ font-size: 0.8em; color: #7f8c8d; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }}
        .insights-box {{ background-color: #f8f9fa; border-left: 4px solid #3498db; padding: 20px; margin: 20px 0 0 0; border-radius: 4px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
        .insights-box h3 {{ color: #2c3e50; margin-top: 0; }}
        .insight-item {{ background-color: white; padding: 15px; margin: 12px 0; border-radius: 4px; border-left: 3px solid #95a5a6; }}
        .insight-item.priority-high {{ border-left-color: #dc3545; background-color: #f8d7da; }}
        .insight-item.priority-medium {{ border-left-color: #f4b400; background-color: #fff3cd; }}
        .insight-item.priority-low {{ border-left-color: #3498db; background-color: #e3f2fd; }}
        .insight-header {{ font-weight: bold; font-size: 0.9em; margin-bottom: 8px; display: flex; align-items: center; gap: 8px; }}
        .insight-badge {{ display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 0.75em; font-weight: bold; text-transform: uppercase; }}
        .badge-high {{ background-color: #dc3545; color: white; }}
        .badge-medium {{ background-color: #f4b400; color: white; }}
        .badge-low {{ background-color: #3498db; color: white; }}
        .badge-info {{ background-color: #28a745; color: white; }}
        .insight-details {{ font-size: 0.85em; color: #6c757d; margin-top: 5px; }}
        .shift-breakdown {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0 0 0; }}
        .shift-table {{ background-color: #ffffff; border-radius: 4px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); overflow-x: auto; }}
        .shift-header {{ background: #f1f3f5; color: #2c3e50; padding: 12px; font-size: 0.95em; font-weight: 600; text-align: center; border-bottom: 1px solid #cfd4da; }}
        .shift-table table {{ width: 100%; margin: 0; table-layout: fixed; }}
        .shift-table th {{ background-color: #95a5a6; color: white; padding: 10px 8px; font-size: 0.75em; font-weight: 600; }}
        .shift-table td {{ padding: 8px 6px; text-align: center; font-size: 0.85em; word-wrap: break-word; overflow: hidden; text-overflow: ellipsis; }}
        .speed-fast {{ background-color: #d4edda; color: #155724; }}
        .speed-average {{ background-color: #fff3cd; color: #856404; }}
        .speed-slow {{ background-color: #f8d7da; color: #721c24; }}
        .total-row {{ background-color: #e9ecef; font-weight: bold; border-top: 2px solid #95a5a6; }}
        .dashboard-toggle-btn {{ display: inline-block; background: #3498db; color: white; padding: 12px 30px; border: none; border-radius: 6px; font-size: 1em; font-weight: 600; cursor: pointer; letter-spacing: 0.5px; transition: background 0.2s; margin: 10px 0; }}
        .dashboard-toggle-btn:hover {{ background: #2980b9; }}
        .dashboard-section {{ display: none; padding: 20px 30px; background: #f8f9fa; border-bottom: 1px solid #dee2e6; }}
        .dashboard-section.visible {{ display: block; }}
        .dashboard-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .dashboard-chart {{ background: white; border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); padding: 10px; overflow: hidden; }}
        .dashboard-chart.full-width {{ grid-column: 1 / -1; }}

        /* Interactive report specific styles */
        .toggle-container {{
            display: flex;
            justify-content: center;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
            padding: 20px 30px;
            background: #f8f9fa;
            border-bottom: 2px solid #dee2e6;
        }}
        .date-picker-container {{
            display: flex;
            align-items: center;
            gap: 10px;
            margin-left: 20px;
            padding-left: 20px;
            border-left: 2px solid #dee2e6;
        }}
        .date-picker-container label {{
            font-weight: 600;
            color: #2c3e50;
        }}
        .date-picker-container input[type="date"] {{
            padding: 8px 12px;
            border: 2px solid #3498db;
            font-size: 0.95em;
            outline: none;
            transition: all 0.3s ease;
        }}
        .date-picker-container input[type="date"]:focus {{
            border-color: #2980b9;
            box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);
        }}
        .toggle-btn {{
            padding: 12px 30px;
            border: 2px solid #3498db;
            background: white;
            color: #3498db;
            cursor: pointer;
            font-size: 1em;
            font-weight: 600;
            transition: all 0.3s ease;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .toggle-btn:hover {{
            background: #e3f2fd;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(52, 152, 219, 0.2);
        }}
        .toggle-btn.active {{
            background: #3498db;
            color: white;
        }}
        .view-content {{
            display: none;
        }}
        .view-content.active {{
            display: block;
        }}
        .footer {{
            background: transparent;
            text-align: center;
            padding: 20px;
            color: #4f5b66;
            font-size: 0.85em;
            border-top: none;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Interactive Sorting & Putaway KPI Report</h1>
            <p>Data available: {min_date} to {max_date}</p>
        </div>

        <div class="toggle-container">
            <button class="toggle-btn active" onclick="switchView('last7')" id="btn-last7">Last 7 Days</button>
            <button class="toggle-btn" onclick="switchView('last30')" id="btn-last30">Last 30 Days</button>
            <div class="date-picker-container">
                <label>Custom Range:</label>
                <input type="date" id="start-date" value="{min_date}" min="{min_date}" max="{max_date}" />
                <label>to</label>
                <input type="date" id="end-date" value="{max_date}" min="{min_date}" max="{max_date}" />
                <button class="toggle-btn" onclick="applyCustomRange()" id="btn-custom">Apply</button>
            </div>
        </div>

        <div id="view-last7" class="view-content active">
"""

    html += fragments.get('last7', '<div class="container"><h1>No data</h1></div>')

    html += """
        </div>
        <div id="view-last30" class="view-content">
"""

    html += fragments.get('last30', '<div class="container"><h1>No data</h1></div>')

    html += """
        </div>
        <div id="view-custom" class="view-content">
            <div class="section" style="text-align:center; padding:60px;">
                <h2 style="border:none;">Select a custom date range above and click Apply</h2>
            </div>
        </div>

        <div class="footer">
            <p>2026 Example Logistics All rights reserved - Created by Viktor Berg</p>
            <p>Report Generated: """ + pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S') + """</p>
        </div>
    </div>

    <script>
        // Embedded data for client-side filtering
        const allDailyData = """ + all_data_json + """;
        const allOperatorData = """ + op_data_json + """;

        const viewIds = ['last7', 'last30', 'custom'];

        function switchView(mode) {
            viewIds.forEach(id => {
                const el = document.getElementById('view-' + id);
                const btn = document.getElementById('btn-' + id);
                if (el) el.classList.remove('active');
                if (btn) btn.classList.remove('active');
            });
            const activeEl = document.getElementById('view-' + mode);
            const activeBtn = document.getElementById('btn-' + mode);
            if (activeEl) activeEl.classList.add('active');
            if (activeBtn) activeBtn.classList.add('active');

            // Trigger resize so embedded Plotly charts redraw
            setTimeout(() => window.dispatchEvent(new Event('resize')), 100);
        }

        function parseDate(dateStr) {
            const [year, month, day] = dateStr.split('-').map(Number);
            return new Date(year, month - 1, day);
        }

        function formatDate(dateStr) {
            const d = parseDate(dateStr);
            return d.toLocaleDateString('en-US', {month: 'short', day: 'numeric', year: 'numeric'});
        }

        function formatDay(dateStr) {
            const d = parseDate(dateStr);
            return d.toLocaleDateString('en-US', {weekday: 'long'});
        }

        function formatShortDay(dateStr) {
            const d = parseDate(dateStr);
            return d.toLocaleDateString('en-US', {weekday: 'short'});
        }

        function numberWithCommas(x) {
            return x.toString().replace(/\\B(?=(\\d{3})+(?!\\d))/g, ",");
        }

        function effClass(eff) {
            if (eff > 80) return 'efficiency-high';
            if (eff >= 65) return 'efficiency-medium';
            return 'efficiency-low';
        }

        function applyCustomRange() {
            const startDate = document.getElementById('start-date').value;
            const endDate = document.getElementById('end-date').value;

            if (!startDate || !endDate) {
                alert('Please select both start and end dates.');
                return;
            }
            if (parseDate(startDate) > parseDate(endDate)) {
                alert('Start date must be before end date.');
                return;
            }

            const filtered = allDailyData.filter(item => {
                return item.date >= startDate && item.date <= endDate;
            }).sort((a, b) => a.date.localeCompare(b.date));

            if (filtered.length === 0) {
                alert('No data available for the selected date range.');
                return;
            }

            // Filter operator data
            const filteredOps = allOperatorData.filter(item => {
                return item.date >= startDate && item.date <= endDate;
            });

            // Aggregate metrics
            const totalSort = filtered.reduce((s, i) => s + i.sort_count, 0);
            const totalPutaway = filtered.reduce((s, i) => s + i.putaway_count, 0);
            const avgSortEff = filtered.filter(i => i.sort_count > 0).length > 0
                ? filtered.filter(i => i.sort_count > 0).reduce((s, i) => s + i.sort_efficiency, 0) / filtered.filter(i => i.sort_count > 0).length : 0;
            const avgPutEff = filtered.filter(i => i.putaway_count > 0).length > 0
                ? filtered.filter(i => i.putaway_count > 0).reduce((s, i) => s + i.putaway_efficiency, 0) / filtered.filter(i => i.putaway_count > 0).length : 0;
            const totalOperators = new Set(filteredOps.map(o => o.operator)).size;

            // Top 5 sorting operators
            const sortOpMap = {};
            filteredOps.filter(o => o.sort_count > 0).forEach(o => {
                if (!sortOpMap[o.operator]) sortOpMap[o.operator] = {count: 0, effSum: 0, effN: 0};
                sortOpMap[o.operator].count += o.sort_count;
                sortOpMap[o.operator].effSum += o.sort_efficiency;
                sortOpMap[o.operator].effN += 1;
            });
            const top5Sort = Object.entries(sortOpMap)
                .map(([op, d]) => ({name: op.includes('_') ? op.split('_').slice(1).join('_') : op, avgEff: d.effN > 0 ? d.effSum / d.effN : 0, count: d.count}))
                .sort((a, b) => b.avgEff - a.avgEff)
                .slice(0, 5);

            // Top 5 putaway operators
            const putOpMap = {};
            filteredOps.filter(o => o.putaway_count > 0).forEach(o => {
                if (!putOpMap[o.operator]) putOpMap[o.operator] = {count: 0, effSum: 0, effN: 0};
                putOpMap[o.operator].count += o.putaway_count;
                putOpMap[o.operator].effSum += o.putaway_efficiency;
                putOpMap[o.operator].effN += 1;
            });
            const top5Put = Object.entries(putOpMap)
                .map(([op, d]) => ({name: op.includes('_') ? op.split('_').slice(1).join('_') : op, avgEff: d.effN > 0 ? d.effSum / d.effN : 0, count: d.count}))
                .sort((a, b) => b.avgEff - a.avgEff)
                .slice(0, 5);

            const periodLabel = formatDate(startDate) + ' - ' + formatDate(endDate);
            const isSingleDay = (startDate === endDate);

            // Build custom view HTML
            let customHTML = '';

            // Summary metrics
            customHTML += `
                <div class="section">
                    <h2>${isSingleDay ? 'Daily Overview' : 'Custom Period Overview'} - ${periodLabel}</h2>
                    <div class="metrics-grid">
                        <div class="metric-box primary">
                            <div class="metric-label">SORTING</div>
                            <div class="metric-value" style="font-size: 42px;">${numberWithCommas(totalSort)}</div>
                            <div class="metric-sublabel">Modules Sorted</div>
                        </div>
                        <div class="metric-box info">
                            <div class="metric-label">TOP 5 SORTING</div>
                            <div class="top-performers">
                                ${top5Sort.length > 0 ? top5Sort.map((o, i) => `<div class="top-performers-item">${i+1}. ${o.name}: ${o.avgEff.toFixed(1)}%</div>`).join('') : '<div class="top-performers-item">No data available</div>'}
                            </div>
                        </div>
                        <div class="metric-box success">
                            <div class="metric-label">PUTAWAY</div>
                            <div class="metric-value" style="font-size: 42px;">${numberWithCommas(totalPutaway)}</div>
                            <div class="metric-sublabel">Modules Put Away</div>
                        </div>
                        <div class="metric-box info">
                            <div class="metric-label">TOP 5 PUTAWAY</div>
                            <div class="top-performers">
                                ${top5Put.length > 0 ? top5Put.map((o, i) => `<div class="top-performers-item">${i+1}. ${o.name}: ${o.avgEff.toFixed(1)}%</div>`).join('') : '<div class="top-performers-item">No data available</div>'}
                            </div>
                        </div>
                    </div>
                </div>`;

            // Efficiency summary
            customHTML += `
                <div class="section">
                    <h2>Performance Summary</h2>
                    <div class="metrics-grid-equal">
                        <div class="metric-box warning">
                            <div class="metric-label">AVG SORT EFFICIENCY</div>
                            <div class="metric-value ${effClass(avgSortEff)}">${avgSortEff.toFixed(1)}%</div>
                        </div>
                        <div class="metric-box warning">
                            <div class="metric-label">AVG PUTAWAY EFFICIENCY</div>
                            <div class="metric-value ${effClass(avgPutEff)}">${avgPutEff.toFixed(1)}%</div>
                        </div>
                        <div class="metric-box info">
                            <div class="metric-label">ACTIVE OPERATORS</div>
                            <div class="metric-value">${totalOperators}</div>
                        </div>
                        <div class="metric-box info">
                            <div class="metric-label">DAYS WITH DATA</div>
                            <div class="metric-value">${filtered.length}</div>
                        </div>
                    </div>
                </div>`;

            // Daily cards (up to 10 days)
            if (filtered.length <= 10 && filtered.length > 1) {
                customHTML += '<div class="section"><h2>Daily Performance</h2><div class="day-card-container">';
                filtered.forEach(day => {
                    customHTML += `<div class="day-card">
                        <div class="day-card-header">${formatShortDay(day.date)}</div>
                        <div class="day-card-date">${parseDate(day.date).toLocaleDateString('en-US', {month: '2-digit', day: '2-digit'})}</div>
                        <div class="day-card-label" style="margin-top:10px;">SORTING</div>
                        <div class="day-card-volume">${numberWithCommas(day.sort_count)}</div>
                        <div class="day-card-label" style="margin-top:15px;border-top:1px solid #ddd;padding-top:15px;">PUTAWAY</div>
                        <div class="day-card-volume">${numberWithCommas(day.putaway_count)}</div>
                    </div>`;
                });
                customHTML += '</div></div>';
            }

            // Chart (multi-day only)
            if (filtered.length > 1) {
                customHTML += '<div class="section"><h2>Daily Volume Trend</h2><div id="custom-chart" style="width:100%;"></div></div>';
            }

            // Shift comparison
            const shift1Sort = filtered.reduce((s, i) => s + i.shift1_sort, 0);
            const shift1Put = filtered.reduce((s, i) => s + i.shift1_putaway, 0);
            const shift2Sort = filtered.reduce((s, i) => s + i.shift2_sort, 0);
            const shift2Put = filtered.reduce((s, i) => s + i.shift2_putaway, 0);

            const shift1SortOps = new Set(filteredOps.filter(o => o.shift === 'Shift1' && o.sort_count > 0).map(o => o.operator)).size;
            const shift1PutOps = new Set(filteredOps.filter(o => o.shift === 'Shift1' && o.putaway_count > 0).map(o => o.operator)).size;
            const shift2SortOps = new Set(filteredOps.filter(o => o.shift === 'Shift2' && o.sort_count > 0).map(o => o.operator)).size;
            const shift2PutOps = new Set(filteredOps.filter(o => o.shift === 'Shift2' && o.putaway_count > 0).map(o => o.operator)).size;

            const s1SortEffArr = filteredOps.filter(o => o.shift === 'Shift1' && o.sort_count > 0).map(o => o.sort_efficiency);
            const s2SortEffArr = filteredOps.filter(o => o.shift === 'Shift2' && o.sort_count > 0).map(o => o.sort_efficiency);
            const s1SortEff = s1SortEffArr.length > 0 ? s1SortEffArr.reduce((a,b) => a+b, 0) / s1SortEffArr.length : 0;
            const s2SortEff = s2SortEffArr.length > 0 ? s2SortEffArr.reduce((a,b) => a+b, 0) / s2SortEffArr.length : 0;

            const s1PutEffArr = filteredOps.filter(o => o.shift === 'Shift1' && o.putaway_count > 0).map(o => o.putaway_efficiency);
            const s2PutEffArr = filteredOps.filter(o => o.shift === 'Shift2' && o.putaway_count > 0).map(o => o.putaway_efficiency);
            const s1PutEff = s1PutEffArr.length > 0 ? s1PutEffArr.reduce((a,b) => a+b, 0) / s1PutEffArr.length : 0;
            const s2PutEff = s2PutEffArr.length > 0 ? s2PutEffArr.reduce((a,b) => a+b, 0) / s2PutEffArr.length : 0;

            customHTML += `
                <div class="section">
                    <h2>Shift Performance Comparison</h2>
                    <div class="shift-comparison">
                        <div class="shift-box shift1">
                            <div class="shift-header-label" style="color: #FFA000;">1st Shift</div>
                            <div class="shift-stats">
                                <div class="stat-item"><div class="stat-value">${numberWithCommas(shift1Sort)}</div><div class="stat-label">Sorting</div></div>
                                <div class="stat-item"><div class="stat-value">${numberWithCommas(shift1Put)}</div><div class="stat-label">Putaway</div></div>
                                <div class="stat-item"><div class="stat-value">${s1SortEff.toFixed(0)}%</div><div class="stat-label">Sort Eff.</div></div>
                                <div class="stat-item"><div class="stat-value">${s1PutEff.toFixed(0)}%</div><div class="stat-label">Putaway Eff.</div></div>
                                <div class="stat-item"><div class="stat-value">${shift1SortOps}</div><div class="stat-label">Sort Ops</div></div>
                                <div class="stat-item"><div class="stat-value">${shift1PutOps}</div><div class="stat-label">Putaway Ops</div></div>
                            </div>
                        </div>
                        <div class="shift-box shift2">
                            <div class="shift-header-label" style="color: #7B1FA2;">2nd Shift</div>
                            <div class="shift-stats">
                                <div class="stat-item"><div class="stat-value">${numberWithCommas(shift2Sort)}</div><div class="stat-label">Sorting</div></div>
                                <div class="stat-item"><div class="stat-value">${numberWithCommas(shift2Put)}</div><div class="stat-label">Putaway</div></div>
                                <div class="stat-item"><div class="stat-value">${s2SortEff.toFixed(0)}%</div><div class="stat-label">Sort Eff.</div></div>
                                <div class="stat-item"><div class="stat-value">${s2PutEff.toFixed(0)}%</div><div class="stat-label">Putaway Eff.</div></div>
                                <div class="stat-item"><div class="stat-value">${shift2SortOps}</div><div class="stat-label">Sort Ops</div></div>
                                <div class="stat-item"><div class="stat-value">${shift2PutOps}</div><div class="stat-label">Putaway Ops</div></div>
                            </div>
                        </div>
                    </div>
                </div>`;

            // Render into custom view
            const customView = document.getElementById('view-custom');
            customView.innerHTML = customHTML;
            switchView('custom');

            // Draw Plotly chart if multi-day
            if (filtered.length > 1) {
                const dates = filtered.map(d => parseDate(d.date));
                const trace1 = {
                    x: dates, y: filtered.map(d => d.sort_count),
                    type: 'bar', name: 'Sorting',
                    marker: {color: '#3498db'},
                    text: filtered.map(d => numberWithCommas(d.sort_count)),
                    textposition: 'auto',
                    textfont: {color: 'white', size: 10}
                };
                const trace2 = {
                    x: dates, y: filtered.map(d => d.putaway_count),
                    type: 'bar', name: 'Putaway',
                    marker: {color: '#28a745'},
                    text: filtered.map(d => numberWithCommas(d.putaway_count)),
                    textposition: 'auto',
                    textfont: {color: 'white', size: 10}
                };
                const trace3 = {
                    x: dates, y: filtered.map(d => d.sort_efficiency),
                    type: 'scatter', mode: 'lines+markers', name: 'Sort Efficiency %',
                    yaxis: 'y2',
                    line: {color: '#e74c3c', width: 2},
                    marker: {size: 6}
                };
                const trace4 = {
                    x: dates, y: filtered.map(d => d.putaway_efficiency),
                    type: 'scatter', mode: 'lines+markers', name: 'Putaway Efficiency %',
                    yaxis: 'y2',
                    line: {color: '#f39c12', width: 2, dash: 'dash'},
                    marker: {size: 6}
                };
                const layout = {
                    title: {text: `Daily Volume & Efficiency - ${periodLabel}`, font: {size: 18, color: '#2c3e50'}, x: 0.5},
                    barmode: 'group',
                    xaxis: {tickformat: '%b %d', dtick: 86400000, tickangle: -45, rangebreaks: [{bounds: ["sat", "mon"]}]},
                    yaxis: {title: 'Volume'},
                    yaxis2: {title: 'Efficiency %', overlaying: 'y', side: 'right', range: [0, 110]},
                    legend: {orientation: 'h', y: 1.12, x: 0.5, xanchor: 'center'},
                    height: 500,
                    margin: {l: 60, r: 60, t: 80, b: 80},
                    plot_bgcolor: 'white',
                    paper_bgcolor: 'white'
                };
                Plotly.newPlot('custom-chart', [trace1, trace2, trace3, trace4], layout, {displayModeBar: false, responsive: true});
            }
        }
    </script>
</body>
</html>"""

    # Save the report
    output_file = output_path / "Interactive_Sorting_Putaway_Report.html"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"   Interactive report saved: {output_file}")
    return str(output_file)


def main():
    """Main execution - Unified Sorting Pipeline"""
    print("=" * 70)
    print("SORTING AND PUTAWAY KPI - UNIFIED REPORT")
    print("=" * 70)
    
    Config.OUTPUT_PATH.mkdir(parents=True, exist_ok=True)
    Config.load_config()
    
    # Interactive date range selection
    report_start, report_end, range_label = select_date_range()
    
    try:
        operator_lookup = load_operator_names()
        
        print("\n[1/7] LOADING DATA FILES")
        print("-" * 70)
        
        df_102 = pd.read_csv(Config.DATA_PATH / Config.FILE_102_MODULE)
        df_102_processed = classify_mix_solid(df_102)
        
        df_502_path = Config.DATA_PATH / Config.FILE_502_RECEIVING
        if df_502_path.exists():
            df_502 = pd.read_csv(df_502_path)
            print(f"   502 receiving file: {len(df_502):,} rows")
        
        df_810 = pd.read_csv(Config.DATA_PATH / Config.FILE_810_SORTING)
        print(f"   Sorting file: {len(df_810):,} rows")
        
        df_putaway_container_path = Config.DATA_PATH / Config.FILE_PUTAWAY_CONTAINER
        if df_putaway_container_path.exists():
            df_putaway_container = pd.read_csv(df_putaway_container_path)
            print(f"   Putaway container file: {len(df_putaway_container):,} rows")
        
        df_putaway_unit = pd.read_csv(Config.DATA_PATH / Config.FILE_PUTAWAY_UNIT)
        print(f"   Putaway unit file: {len(df_putaway_unit):,} rows")
        
        df_sorting = process_sorting_data(df_810, df_102_processed, operator_lookup)
        sorting_metrics = calculate_sorting_speed_metrics(df_sorting)
        
        putaway_metrics = process_putaway_data(df_putaway_unit, df_102_processed, operator_lookup)
        
        benchmarks = calculate_dynamic_benchmarks(sorting_metrics, putaway_metrics)
        
        # Resolve "all data" range from actual data
        if report_start is None or report_end is None:
            all_dates = []
            if len(sorting_metrics) > 0:
                all_dates.extend(sorting_metrics['Operation_Date'].dropna().tolist())
            if len(putaway_metrics) > 0:
                all_dates.extend(putaway_metrics['Operation_Date'].dropna().tolist())
            if all_dates:
                report_start = min(all_dates)
                report_end = max(all_dates)
                print(f"\n   All data range resolved: {report_start} to {report_end}")
            else:
                print("\n   No data found!")
                return None, None, None
        
        daily_metrics = aggregate_daily_metrics(sorting_metrics, putaway_metrics, report_start, report_end, benchmarks)
        
        if len(daily_metrics) == 0:
            print("\n   No metrics generated!")
            return None, None, None
        
        # Interactive report mode - generate all-in-one HTML with date picker
        if range_label == "interactive":
            output_file = generate_interactive_report(
                daily_metrics, sorting_metrics, putaway_metrics, benchmarks, Config.OUTPUT_PATH
            )
            print("\n" + "=" * 70)
            print("INTERACTIVE REPORT COMPLETE")
            print("=" * 70)
            print(f"\n   Output: {output_file}")
            print(f"   Open in your web browser to use the date picker.")
            print("\n" + "=" * 70)
            Config.save_config()
            return daily_metrics, None, None
        
        is_single_day = (report_start == report_end)
        
        operator_stats = analyze_operator_performance(daily_metrics, report_start, report_end)
        daily_trend = calculate_daily_efficiency_trend(daily_metrics, report_start, report_end)
        predictions, training_modules = predict_daily_volumes(daily_metrics)
        
        workforce_insights = analyze_workforce_optimization(daily_metrics, operator_stats, report_start, report_end)
        
        print("\n[EXPORTING RESULTS]")
        print("-" * 70)
        
        # Create subfolder based on range type
        subfolder_map = {
            'yesterday': 'Daily',
            'this_week': 'Week-to-Date',
            'last_week': 'Weekly',
            'last_30_days': '30-Day',
            'all_data': 'All Data'
        }
        subfolder_name = subfolder_map.get(range_label, 'Custom')
        output_dir = Config.OUTPUT_PATH / subfolder_name
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Dynamic filenames based on range
        start_str = str(report_start)
        end_str = str(report_end)
        
        if range_label == 'yesterday':
            file_prefix = f"Daily_Report_{start_str}"
        elif range_label == 'last_week':
            file_prefix = f"Weekly_Report_{start_str}_to_{end_str}"
        elif range_label == 'this_week':
            file_prefix = f"Week-to-Date_Report_{start_str}_to_{end_str}"
        elif range_label == 'last_7_days':
            file_prefix = f"7-Day_Report_{start_str}_to_{end_str}"
        elif range_label == 'last_30_days':
            file_prefix = f"30-Day_Report_{start_str}_to_{end_str}"
        elif range_label == 'all_data':
            file_prefix = f"Full_Report_{start_str}_to_{end_str}"
        else:
            file_prefix = f"Custom_Report_{start_str}_to_{end_str}"
        
        # Generate dashboard HTML (multi-day only, embedded in report)
        dashboard_html = None
        if PLOTLY_AVAILABLE and not is_single_day:
            dashboard_html = create_interactive_dashboard(
                daily_metrics, operator_stats, daily_trend,
                report_start, report_end,
                predictions, training_modules
            )
        
        # HTML report generation - choose daily or multi-day format
        if is_single_day:
            html_content = generate_daily_html(daily_metrics, operator_stats, report_start)
        else:
            html_content = generate_email_content(
                daily_metrics, operator_stats, daily_trend,
                report_start, report_end, range_label,
                workforce_insights, None, predictions, training_modules,
                dashboard_html
            )
        
        html_filename = f"{file_prefix}.html"
        html_file = output_dir / html_filename
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"   HTML saved: {subfolder_name}/{html_filename}")
        
        Config.save_config()
        
        # Summary
        print("\n" + "=" * 70)
        print("SUMMARY")
        print("=" * 70)
        
        report_data = daily_metrics[
            (daily_metrics['Operation_Date'] >= report_start) &
            (daily_metrics['Operation_Date'] <= report_end)
        ]
        
        print(f"\n   Report range: {report_start} to {report_end} ({range_label})")
        print(f"   Days with data: {report_data['Operation_Date'].nunique()}")
        print(f"   Active operators: {report_data['Operator'].nunique()}")
        print(f"\n   Sorted: {report_data['Sort_Count'].sum():,.0f} modules")
        print(f"   Put away: {report_data['Putaway_Count'].sum():,.0f} modules")
        
        if not is_single_day:
            print(f"\n   Daily Breakdown:")
            daily_totals = report_data.groupby('Operation_Date').agg({
                'Sort_Count': 'sum',
                'Putaway_Count': 'sum'
            }).reset_index()
            for _, row in daily_totals.iterrows():
                day_name = pd.to_datetime(row['Operation_Date']).strftime('%a %m/%d')
                print(f"     {day_name}: Sort={row['Sort_Count']:,.0f}, Putaway={row['Putaway_Count']:,.0f}")
        
        sort_mask = report_data['Sort_Speed'] > 0
        if sort_mask.sum() > 0:
            print(f"\n   Avg sort speed: {report_data[sort_mask]['Sort_Speed'].mean():.1f} sec/module")
            print(f"   Avg sort efficiency: {report_data[sort_mask]['Sort_Efficiency'].mean():.1f}%")
        
        if predictions is not None:
            print(f"\n   ML PREDICTIONS:")
            for _, row in predictions.iterrows():
                print(f"     {row['Shift']}: {row['Predicted_Sort_Volume']:,.0f} modules @ {row['Predicted_Efficiency']:.1f}% efficiency")
        
        print("\n" + "=" * 70)
        print("Report Complete!")
        print("=" * 70)
        
        return daily_metrics, operator_stats, predictions
        
    except Exception as e:
        print(f"\n   ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None, None

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == '--config':
        configure_system_interactive()
    else:
        daily_metrics, operator_stats, predictions = main()