import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import json
import re
import argparse


ON_TIME_TARGET_PCT = 95.0
REQUIRED_COLUMNS = [
    'TRAILER NO', 'TEMP.TRAILER', 'MODULE NO',
    'PLAN SHIP DATE', 'PLAN SHIP TIME',
    'SHIPMENT LOAD DATE', 'SHIPMENT LOAD TIME', 'QTY',
]


def compress_trailer_ranges(trailer_names):
    """Compress a list of temp trailer names into ranges.
    
    E.g. ['AAA17','AAA18','AAA19','AAA20'] -> 'AAA17 - AAA20'
         ['AAA98','AAA99','AAB01','AAB02','ADS2','ADS3','ADS4'] -> 'AAA98 - AAA99, AAB01 - AAB02, ADS2 - ADS4'
    """
    if not trailer_names:
        return ''
    
    # Parse each name into (prefix, number)
    parsed = []
    for name in sorted(trailer_names):
        match = re.match(r'^(.*?)(\d+)$', name)
        if match:
            parsed.append((match.group(1), int(match.group(2)), name))
        else:
            parsed.append((name, None, name))
    
    # Group consecutive items with the same prefix
    ranges = []
    i = 0
    while i < len(parsed):
        prefix, num, original = parsed[i]
        if num is None:
            ranges.append(original)
            i += 1
            continue
        
        # Find consecutive run with same prefix
        start_name = original
        end_name = original
        j = i + 1
        prev_num = num
        while j < len(parsed):
            p2, n2, o2 = parsed[j]
            if p2 == prefix and n2 is not None and n2 == prev_num + 1:
                end_name = o2
                prev_num = n2
                j += 1
            else:
                break
        
        if start_name == end_name:
            ranges.append(start_name)
        else:
            ranges.append(f"{start_name} - {end_name}")
        i = j
    
    return ', '.join(ranges)


def load_and_process_data(data_folder):
    """Load and process shipping data from 202.csv.

    Returns (df, load_info) where load_info summarizes row counts and the
    ship-date span (surfaced in the report footer).
    """
    csv_file = os.path.join(data_folder, '202.csv')

    if not os.path.exists(csv_file):
        raise FileNotFoundError(
            f"Could not find '{csv_file}'. Place the 202.csv export in the "
            f"'{data_folder}' folder (or pass --data-folder)."
        )

    try:
        df = pd.read_csv(csv_file)
    except Exception as exc:
        raise RuntimeError(f"Failed to read '{csv_file}': {exc}") from exc

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"'{csv_file}' is missing required columns: {', '.join(missing)}. "
            f"Found columns: {', '.join(df.columns)}"
        )

    raw_rows = len(df)

    df['PLAN SHIP DATE'] = pd.to_datetime(df['PLAN SHIP DATE'], format='mixed', dayfirst=False, errors='coerce')
    df['SHIPMENT LOAD DATE'] = pd.to_datetime(df['SHIPMENT LOAD DATE'], format='mixed', dayfirst=False, errors='coerce')
    df['QTY'] = pd.to_numeric(df['QTY'], errors='coerce').fillna(0)

    invalid_ship_rows = int(df['PLAN SHIP DATE'].isna().sum())
    df = df.dropna(subset=['PLAN SHIP DATE'])

    latest_ship_date = df['SHIPMENT LOAD DATE'].max()
    if pd.isna(latest_ship_date):
        latest_ship_date = df['PLAN SHIP DATE'].max()

    df['ON_TIME'] = False
    df['EARLY'] = False
    df['IS_SHIPPED'] = ~df['SHIPMENT LOAD DATE'].isna()

    past_mask = df['PLAN SHIP DATE'] <= latest_ship_date
    shipped_past = past_mask & df['IS_SHIPPED']

    df.loc[shipped_past, 'ON_TIME'] = df.loc[shipped_past, 'SHIPMENT LOAD DATE'] == df.loc[shipped_past, 'PLAN SHIP DATE']
    df.loc[shipped_past, 'EARLY'] = df.loc[shipped_past, 'SHIPMENT LOAD DATE'] < df.loc[shipped_past, 'PLAN SHIP DATE']

    df = df[past_mask].copy()

    last_plan_date = df['PLAN SHIP DATE'].max()
    dropped_last_plan_date = None
    if pd.notna(last_plan_date):
        has_unshipped_last = df.loc[
            df['PLAN SHIP DATE'] == last_plan_date, 'IS_SHIPPED'
        ].eq(False).any()
        if has_unshipped_last:
            dropped_last_plan_date = last_plan_date
            df = df[df['PLAN SHIP DATE'] < last_plan_date].copy()

    load_info = {
        'raw_rows': raw_rows,
        'invalid_ship_rows': invalid_ship_rows,
        'usable_rows': len(df),
        'min_ship_date': df['PLAN SHIP DATE'].min() if not df.empty else None,
        'max_ship_date': df['PLAN SHIP DATE'].max() if not df.empty else None,
        'dropped_last_plan_date': dropped_last_plan_date,
    }

    return df, load_info

def fill_missing_dates_shipping(daily_summary, date_col, start_date, end_date):
    """Fill in all missing business days in the range with zero-value rows (skips weekends)."""
    int_cols = ['total_modules', 'total_qty',
                'on_time_qty', 'early_qty', 'late_qty',
                'on_time_trailers', 'early_trailers', 'late_trailers',
                'on_time_modules', 'early_modules', 'late_modules', 'total_trailers']
    pct_cols = ['on_time_pct', 'early_pct', 'late_pct', 'met_plan_pct']

    all_dates = pd.date_range(start=start_date, end=end_date, freq='B')
    if daily_summary.empty:
        empty = {date_col: all_dates}
        for c in int_cols:
            empty[c] = 0
        for c in pct_cols:
            empty[c] = 0.0
        return pd.DataFrame(empty)

    full_df = pd.DataFrame({date_col: all_dates})
    result = full_df.merge(daily_summary, on=date_col, how='left')
    for col in int_cols:
        if col in result.columns:
            result[col] = result[col].fillna(0).astype(int)
    for col in pct_cols:
        if col in result.columns:
            result[col] = result[col].fillna(0.0)
    return result

def build_daily_summary(period_data):
    """Build daily summary data for a provided dataset."""

    if period_data is None or period_data.empty:
        return pd.DataFrame(columns=[
            'PLAN SHIP DATE', 'total_modules', 'total_qty',
            'on_time_qty', 'early_qty', 'late_qty',
            'on_time_trailers', 'early_trailers', 'late_trailers',
            'on_time_modules', 'early_modules', 'late_modules',
            'total_trailers', 'on_time_pct', 'early_pct', 'late_pct',
            'met_plan_pct',
        ])

    def classify_trailers(x):
        """Classify each trailer uniquely using DATE-ONLY comparison. Matches detail view logic."""
        # x is the TRAILER NO series for this date group, x.index are the row indices
        date_rows = period_data.loc[x.index]
        results = {}

        # Module status should be row-level, not trailer-level. A single trailer can contain
        # modules loaded on different dates (e.g., some on-time, some late).
        module_counts = {'on_time': set(), 'early': set(), 'late': set()}
        for _, row in date_rows.iterrows():
            module_no = row.get('MODULE NO')
            if pd.isna(module_no):
                continue

            if bool(row.get('ON_TIME', False)):
                module_counts['on_time'].add(module_no)
            elif bool(row.get('EARLY', False)):
                module_counts['early'].add(module_no)
            else:
                module_counts['late'].add(module_no)
        
        for trailer in date_rows['TRAILER NO'].unique():
            trailer_rows = date_rows[date_rows['TRAILER NO'] == trailer]
            
            # Get plan date (date only from PLAN SHIP DATE) - all rows for this date have same plan date
            plan_date = pd.Timestamp(date_rows['PLAN SHIP DATE'].iloc[0]).date()
            
            # Check if shipped
            is_shipped = trailer_rows['IS_SHIPPED'].any()

            if is_shipped:
                ship_dates = trailer_rows['SHIPMENT LOAD DATE'].dropna().dt.date

                # Deterministic trailer classification for mixed dates:
                # late if any late/unshipped rows; else on-time if any on-time rows; else early.
                if ship_dates.empty:
                    classification = 'late'
                elif any(d > plan_date for d in ship_dates):
                    classification = 'late'
                elif any(d == plan_date for d in ship_dates):
                    classification = 'on_time'
                else:
                    classification = 'early'
            else:
                # Not shipped -> count as late
                classification = 'late'
            
            results[trailer] = classification
        
        return pd.Series({
            'on_time': sum(1 for v in results.values() if v == 'on_time'),
            'early': sum(1 for v in results.values() if v == 'early'),
            'late': sum(1 for v in results.values() if v == 'late'),
            'on_time_modules': len(module_counts['on_time']),
            'early_modules': len(module_counts['early']),
            'late_modules': len(module_counts['late'])
        })
    
    daily_summary = period_data.groupby('PLAN SHIP DATE').agg(
        total_modules=('MODULE NO', 'size'),
        total_qty=('QTY', 'sum'),
        on_time_qty=('QTY', lambda x: x[period_data.loc[x.index, 'ON_TIME']].sum()),
        early_qty=('QTY', lambda x: x[period_data.loc[x.index, 'EARLY']].sum()),
        late_qty=('QTY', lambda x: x[(~period_data.loc[x.index, 'ON_TIME']) & (~period_data.loc[x.index, 'EARLY'])].sum())
    ).reset_index()
    
    # Add trailer classifications (includes module counts)
    trailer_counts = period_data.groupby('PLAN SHIP DATE')['TRAILER NO'].apply(classify_trailers)
    if isinstance(trailer_counts.index, pd.MultiIndex):
        trailer_counts = trailer_counts.unstack(fill_value=0)
    
    daily_summary = daily_summary.merge(trailer_counts, left_on='PLAN SHIP DATE', right_index=True, how='left')
    daily_summary['on_time_trailers'] = daily_summary['on_time'].fillna(0).astype(int)
    daily_summary['early_trailers'] = daily_summary['early'].fillna(0).astype(int)
    daily_summary['late_trailers'] = daily_summary['late'].fillna(0).astype(int)

    # Keep early as its own bucket (do NOT fold into on_time). Late is the
    # remainder so the three buckets always sum to total_modules.
    daily_summary['on_time_modules'] = daily_summary['on_time_modules'].fillna(0).astype(int)
    daily_summary['early_modules'] = daily_summary['early_modules'].fillna(0).astype(int)
    daily_summary['late_modules'] = (
        daily_summary['total_modules']
        - daily_summary['on_time_modules']
        - daily_summary['early_modules']
    ).clip(lower=0).astype(int)

    daily_summary = daily_summary.drop(columns=['on_time', 'early', 'late'], errors='ignore')

    # Recalculate total_trailers from classifications to ensure consistency
    daily_summary['total_trailers'] = (
        daily_summary['on_time_trailers']
        + daily_summary['early_trailers']
        + daily_summary['late_trailers']
    )

    totals = daily_summary['total_modules'].replace(0, pd.NA)
    daily_summary['on_time_pct'] = (daily_summary['on_time_modules'] / totals * 100).round(1).fillna(0.0)
    daily_summary['early_pct'] = (daily_summary['early_modules'] / totals * 100).round(1).fillna(0.0)
    daily_summary['late_pct'] = (daily_summary['late_modules'] / totals * 100).round(1).fillna(0.0)
    # met_plan_pct = on-time + early (i.e. shipped on or before plan).
    daily_summary['met_plan_pct'] = (
        (daily_summary['on_time_modules'] + daily_summary['early_modules']) / totals * 100
    ).round(1).fillna(0.0)

    daily_summary = daily_summary.sort_values('PLAN SHIP DATE').reset_index(drop=True)

    return daily_summary

def create_daily_summary_calendar_month(df, target_month=None, target_year=None):
    """Create daily summary of shipping performance for a calendar month."""
    
    if target_month is None or target_year is None:
        today = pd.Timestamp.today().normalize()
        first_of_this_month = today.replace(day=1)
        last_month = first_of_this_month - timedelta(days=1)
        target_month = last_month.month
        target_year = last_month.year
    
    month_data = df[
        (df['PLAN SHIP DATE'].dt.month == target_month) & 
        (df['PLAN SHIP DATE'].dt.year == target_year)
    ]
    
    daily_summary = build_daily_summary(month_data)
    
    # Fill in all dates in the month up to last day with data
    if not daily_summary.empty:
        month_start = pd.Timestamp(year=target_year, month=target_month, day=1)
        month_end = daily_summary['PLAN SHIP DATE'].max()
        daily_summary = fill_missing_dates_shipping(daily_summary, 'PLAN SHIP DATE', month_start, month_end)
    
    return daily_summary, 'month', target_month, target_year, month_data

def create_daily_summary_past_2_months(df):
    """Create daily summary of shipping performance for the past 2 months (excluding future dates)."""
    
    today = df['SHIPMENT LOAD DATE'].max()
    if pd.isna(today):
        today = df['PLAN SHIP DATE'].max()
    
    latest_date = min(df['PLAN SHIP DATE'].max(), today)
    start_date = latest_date - timedelta(days=60)
    
    period_data = df[
        (df['PLAN SHIP DATE'] >= start_date) & 
        (df['PLAN SHIP DATE'] <= latest_date)
    ]
    
    daily_summary = build_daily_summary(period_data)
    
    # Fill in all dates in the range
    daily_summary = fill_missing_dates_shipping(daily_summary, 'PLAN SHIP DATE', start_date, latest_date)
    
    return daily_summary, 'past_2_months', start_date, latest_date, period_data


def create_previous_month_summary(df, target_month, target_year):
    """Build a one-shot summary for the month before (target_month, target_year),
    used for the month-over-month delta on the On-Time stat card. Returns None
    if no data exists in the prior month."""
    if df is None or df.empty:
        return None
    anchor = pd.Timestamp(year=target_year, month=target_month, day=1)
    prev_last = anchor - timedelta(days=1)
    prev_data = df[
        (df['PLAN SHIP DATE'].dt.month == prev_last.month) &
        (df['PLAN SHIP DATE'].dt.year == prev_last.year)
    ]
    if prev_data.empty:
        return None
    total = int(prev_data.shape[0])
    met_plan = int((prev_data['ON_TIME'] | prev_data['EARLY']).sum())
    return {
        'month': prev_last.month,
        'year': prev_last.year,
        'total_modules': total,
        'met_plan_modules': met_plan,
    }


def _rolling_met_plan_pct(daily_summary, window=7):
    """Return a list of rolling met-plan % values (window of N active days).

    'Met plan' = on_time + early (shipped on or before plan date). Only days
    with modules count toward the window so filled zero-rows don't drag the
    average down. Days with no history yet return None (gap in the line).
    """
    values = []
    active = []
    for _, row in daily_summary.iterrows():
        total = int(row.get('total_modules', 0) or 0)
        if total > 0:
            met = int(row.get('on_time_modules', 0) or 0) + int(row.get('early_modules', 0) or 0)
            active.append((met, total))
            active = active[-window:]
            tot = sum(t for _, t in active)
            on = sum(o for o, _ in active)
            values.append(round(on / tot * 100, 1) if tot > 0 else None)
        else:
            values.append(None)
    return values


def create_visualization(daily_summary, mode, period_info):
    """Stacked bar chart: Early / On-Time / Late, plus a rolling met-plan %
    line and a horizontal target line."""

    fig = go.Figure()
    plot_dates = daily_summary['PLAN SHIP DATE']

    on_time_counts = daily_summary['on_time_modules'].astype(int)
    early_counts = daily_summary['early_modules'].astype(int) if 'early_modules' in daily_summary.columns else (on_time_counts * 0)
    late_counts = daily_summary['late_modules'].astype(int)
    totals = daily_summary['total_modules'].astype(int)
    met_plan = on_time_counts + early_counts
    max_total = int(totals.max()) if len(totals) else 0

    def pct_of(n, t):
        return (n / t * 100) if t > 0 else 0.0

    # Adaptive label positioning on the on-time bar: short bars push the
    # percentage label outside so it stays readable.
    label_threshold = max(1, int(max_total * 0.08))
    on_time_text, on_time_positions = [], []
    for n, t in zip(on_time_counts, totals):
        pct = pct_of(n, t)
        if t <= 0 or n <= 0:
            on_time_text.append('')
            on_time_positions.append('inside')
        elif n < label_threshold:
            on_time_text.append(f"{pct:.1f}%")
            on_time_positions.append('outside')
        else:
            on_time_text.append(f"{pct:.1f}%")
            on_time_positions.append('inside')

    # Per-trace customdata under hovermode='x unified' — each trace prints
    # its own labelled line.
    early_hover = [
        f"{int(n):,} modules ({pct_of(n, t):.1f}%)" if t > 0 else '—'
        for n, t in zip(early_counts, totals)
    ]
    on_time_hover = [
        f"{int(n):,} modules ({pct_of(n, t):.1f}%)" if t > 0 else '—'
        for n, t in zip(on_time_counts, totals)
    ]
    late_hover = [
        f"{int(n):,} modules ({pct_of(n, t):.1f}%)" if t > 0 else '—'
        for n, t in zip(late_counts, totals)
    ]

    fig.add_trace(go.Bar(
        name='Shipped Early',
        x=plot_dates,
        y=early_counts,
        marker_color='#7bc47f',
        text='',
        hovertemplate='<b>Early</b>: %{customdata}<extra></extra>',
        customdata=early_hover,
        showlegend=True,
    ))

    fig.add_trace(go.Bar(
        name='Shipped on Ship Date',
        x=plot_dates,
        y=on_time_counts,
        marker_color='#28a745',
        text=on_time_text,
        textposition=on_time_positions,
        textfont=dict(color='white', size=10, family='Arial'),
        cliponaxis=False,
        hovertemplate='<b>On Time</b>: %{customdata}<extra></extra>',
        customdata=on_time_hover,
        showlegend=True,
    ))

    fig.add_trace(go.Bar(
        name='Shipped Late',
        x=plot_dates,
        y=late_counts,
        marker_color='#dc3545',
        text='',
        hovertemplate='<b>Late</b>: %{customdata}<extra></extra>',
        customdata=late_hover,
        showlegend=True,
    ))

    # Rolling 7-day met-plan % (on-time + early) on a secondary y-axis.
    rolling_pct = _rolling_met_plan_pct(daily_summary, window=7)
    fig.add_trace(go.Scatter(
        name='Rolling 7-day Met-Plan %',
        x=plot_dates,
        y=rolling_pct,
        mode='lines+markers',
        yaxis='y2',
        line=dict(color='#2c3e50', width=2, dash='dot'),
        marker=dict(size=6, color='#2c3e50'),
        hovertemplate='<b>Rolling 7d</b>: %{y:.1f}%<extra></extra>',
        connectgaps=False,
        showlegend=True,
    ))

    # Horizontal target line on the secondary axis.
    if len(plot_dates) > 0:
        fig.add_trace(go.Scatter(
            name=f'Target {ON_TIME_TARGET_PCT:.0f}%',
            x=[plot_dates.iloc[0], plot_dates.iloc[-1]],
            y=[ON_TIME_TARGET_PCT, ON_TIME_TARGET_PCT],
            mode='lines',
            yaxis='y2',
            line=dict(color='#888', width=1.5, dash='dash'),
            hoverinfo='skip',
            showlegend=True,
        ))

    if mode == 'month':
        month, year = period_info
        title_text = f'Shipping KPI - {datetime(year, month, 1).strftime("%B %Y")}'
    else:
        start_date, end_date = period_info
        title_text = f'Shipping KPI - Past 2 Months'
    
    fig.update_layout(
        title=dict(text=title_text, font=dict(size=24, color='#2c3e50'), x=0.5, xanchor='center'),
        xaxis_title='Plan Ship Date',
        yaxis_title='Module Count',
        barmode='stack',
        hovermode='x unified',
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(family='Arial, sans-serif', size=12, color='#2c3e50'),
        xaxis=dict(
            tickformat='%b %d',
            dtick='D1',
            tickangle=-45,
            showgrid=True,
            gridcolor='#ecf0f1',
            linecolor='#bdc3c7',
            linewidth=2,
            rangebreaks=[dict(bounds=["sat", "mon"])]
        ),
        yaxis=dict(
            title='Module Count',
            showgrid=True,
            gridcolor='#ecf0f1',
            linecolor='#bdc3c7',
            linewidth=2
        ),
        yaxis2=dict(
            title='Met-Plan %',
            overlaying='y',
            side='right',
            range=[0, 105],
            showgrid=False,
            ticksuffix='%',
            linecolor='#bdc3c7',
            linewidth=2,
        ),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='center',
            x=0.5,
            font=dict(size=14)
        ),
        height=600,
        autosize=True,
        margin=dict(l=80, r=80, t=150, b=100)
    )

    return fig

def generate_summary_stats(daily_summary, mode, period_info, period_data, prev_month=None):
    """Generate summary statistics for the period.

    Surfaces Early as a separate bucket from On-Time so the HTML can render
    them on independent stat cards. The "met-plan" combined total
    (on-time + early) is also returned for the headline KPI / target
    comparison. ``period_data`` is used for the unique trailer count.
    ``prev_month`` (optional) enables the MoM delta.
    """

    on_time_modules = int(daily_summary['on_time_modules'].sum())
    early_modules = int(daily_summary['early_modules'].sum()) if 'early_modules' in daily_summary.columns else 0
    late_modules = int(daily_summary['late_modules'].sum())
    total_modules = int(daily_summary['total_modules'].sum())
    met_plan_modules = on_time_modules + early_modules

    total_qty = int(daily_summary['total_qty'].sum())
    on_time_qty = int(daily_summary['on_time_qty'].sum())
    early_qty = int(daily_summary['early_qty'].sum()) if 'early_qty' in daily_summary.columns else 0
    late_qty = int(daily_summary['late_qty'].sum())

    on_time_trailers = int(daily_summary['on_time_trailers'].sum())
    early_trailers = int(daily_summary['early_trailers'].sum())
    late_trailers = int(daily_summary['late_trailers'].sum())

    # Trailer-event count: each (date, trailer) classification counts once,
    # summed across days. Reflects the operational "how many trailers did we
    # load this month" rather than unique chassis (some chassis recycle daily).
    total_trailer_events = on_time_trailers + early_trailers + late_trailers

    on_time_modules_pct = (on_time_modules / total_modules * 100) if total_modules > 0 else 0.0
    early_modules_pct = (early_modules / total_modules * 100) if total_modules > 0 else 0.0
    late_modules_pct = (late_modules / total_modules * 100) if total_modules > 0 else 0.0
    met_plan_pct = (met_plan_modules / total_modules * 100) if total_modules > 0 else 0.0

    if mode == 'month':
        month, year = period_info
        period_label = datetime(year, month, 1).strftime('%B %Y')
        month_name = datetime(year, month, 1).strftime('%B')
    else:
        start_date, end_date = period_info
        period_label = f"{start_date.strftime('%b %d, %Y')} - {end_date.strftime('%b %d, %Y')}"
        month_name = None

    avg_daily_modules = round(daily_summary['total_modules'].mean(), 1) if len(daily_summary) > 0 else 0

    mom = None
    if prev_month is not None and prev_month.get('total_modules', 0) > 0:
        prev_pct = prev_month['met_plan_modules'] / prev_month['total_modules'] * 100
        mom = {
            'prev_label': datetime(prev_month['year'], prev_month['month'], 1).strftime('%B %Y'),
            'prev_pct': round(prev_pct, 1),
            'delta_pts': round(met_plan_pct - prev_pct, 1),
        }

    return {
        'period_label': period_label,
        'mode': mode,
        'month_name': month_name,
        'total_trailers': total_trailer_events,
        'total_modules': total_modules,
        'total_qty': total_qty,
        # On-Time bucket (exact date match)
        'on_time_trailers': on_time_trailers,
        'on_time_qty': on_time_qty,
        'on_time_modules': on_time_modules,
        'on_time_modules_pct': on_time_modules_pct,
        # Early bucket (shipped before plan)
        'early_trailers': early_trailers,
        'early_qty': early_qty,
        'early_modules': early_modules,
        'early_modules_pct': early_modules_pct,
        # Late bucket (shipped after plan or unshipped)
        'late_trailers': late_trailers,
        'late_qty': late_qty,
        'late_modules': late_modules,
        'late_modules_pct': late_modules_pct,
        # Combined "met plan" headline
        'met_plan_modules': met_plan_modules,
        'met_plan_pct': met_plan_pct,
        'avg_daily_modules': avg_daily_modules,
        'mom': mom,
        'target_pct': ON_TIME_TARGET_PCT,
    }

def _render_mom_line(stats):
    """Render the month-over-month delta line for the headline stat card."""
    mom = stats.get('mom')
    if not mom:
        return ''
    delta = mom['delta_pts']
    if delta > 0:
        cls, sign = 'mom-up', '+'
    elif delta < 0:
        cls, sign = 'mom-down', ''
    else:
        cls, sign = 'mom-flat', ''
    return (
        f"<div class='mom-line'>vs {mom['prev_label']}: "
        f"<span class='{cls}'>{sign}{delta:.1f} pts</span> "
        f"(was {mom['prev_pct']:.1f}%)</div>"
    )


def _render_stats_block(stats):
    """Render the stats grid for a view: Period, Total, Met-Plan
    (headline w/ MoM), Late."""
    mom_line = _render_mom_line(stats)
    return f"""
        <div class="stats-container">
            <div class="stat-card">
                <h3>Period</h3>
                <div class="value" style="font-size: 1.5em;">{stats['period_label']}</div>
            </div>
            <div class="stat-card">
                <h3>TOTAL MODULES</h3>
                <div class="value">{stats['total_modules']:,}</div>
                <div class="subtext">{stats['total_trailers']:,} trailers | Avg. {stats['avg_daily_modules']:.0f} per day</div>
            </div>
            <div class="stat-card success">
                <h3>MET PLAN (ON-TIME OR EARLY)</h3>
                <div class="value">{stats['met_plan_pct']:.1f}%</div>
                <div class="subtext">{stats['met_plan_modules']:,} modules &middot; target {stats.get('target_pct', ON_TIME_TARGET_PCT):.0f}%</div>
                {mom_line}
            </div>
            <div class="stat-card danger">
                <h3>SHIPPED LATE</h3>
                <div class="value">{stats['late_modules_pct']:.1f}%</div>
                <div class="subtext">{stats['late_modules']:,} modules</div>
            </div>
        </div>
    """


def _render_load_info(load_info):
    """Render the data-load summary line in the footer."""
    if not load_info:
        return ''
    parts = [f"raw rows: {load_info.get('raw_rows', 0):,}"]
    invalid = load_info.get('invalid_ship_rows', 0)
    if invalid:
        parts.append(f"dropped (invalid plan date): {invalid:,}")
    parts.append(f"usable: {load_info.get('usable_rows', 0):,}")
    min_d = load_info.get('min_ship_date')
    max_d = load_info.get('max_ship_date')
    if min_d is not None and max_d is not None and not pd.isna(min_d) and not pd.isna(max_d):
        parts.append(
            f"plan-ship range: {pd.Timestamp(min_d).strftime('%b %d, %Y')} – {pd.Timestamp(max_d).strftime('%b %d, %Y')}"
        )
    dropped = load_info.get('dropped_last_plan_date')
    if dropped is not None and not pd.isna(dropped):
        parts.append(f"excluded last plan date {pd.Timestamp(dropped).strftime('%b %d, %Y')} (had unshipped modules)")
    return f"<p style='margin-top: 8px; font-size: 0.85em; color: #6c757d;'>Data: {' | '.join(parts)}</p>"


def create_html_report(
    fig_month, stats_month, daily_summary_month,
    fig_past2months, stats_past2months, daily_summary_past2months,
    df, load_info=None,
):
    """Create HTML report with visualization and statistics."""

    chart_month_html = fig_month.to_html(include_plotlyjs='cdn', div_id='chart-month', config={'displayModeBar': False, 'responsive': True})
    chart_past2months_html = fig_past2months.to_html(include_plotlyjs=False, div_id='chart-past2months', config={'displayModeBar': False, 'responsive': True})

    daily_summary_all = build_daily_summary(df)
    if not daily_summary_all.empty:
        daily_summary_all = fill_missing_dates_shipping(
            daily_summary_all, 'PLAN SHIP DATE',
            daily_summary_all['PLAN SHIP DATE'].min(),
            daily_summary_all['PLAN SHIP DATE'].max()
        )
    js_data = []
    for _, row in daily_summary_all.iterrows():
        js_data.append({
            'date': row['PLAN SHIP DATE'].strftime('%Y-%m-%d'),
            'total_trailers': int(row['total_trailers']),
            'total_modules': int(row['total_modules']),
            'total_qty': int(row['total_qty']),
            'on_time_trailers': int(row['on_time_trailers']),
            'early_trailers': int(row['early_trailers']),
            'late_trailers': int(row['late_trailers']),
            'on_time_qty': int(row['on_time_qty']),
            'early_qty': int(row['early_qty']),
            'late_qty': int(row['late_qty']),
            'on_time_modules': int(row['on_time_modules']),
            'early_modules': int(row.get('early_modules', 0)),
            'late_modules': int(row['late_modules']),
            'on_time_pct': float(row['on_time_pct']),
            'early_pct': float(row.get('early_pct', 0.0)),
            'late_pct': float(row.get('late_pct', 0.0)),
            'met_plan_pct': float(row.get('met_plan_pct', 0.0)),
        })

    all_data_json = json.dumps(js_data)
    
    # Generate trailer-level details for ALL dates in the full dataset (not just month+last31)
    trailer_details = {}
    for date in df['PLAN SHIP DATE'].dropna().unique():
        date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
        date_data = df[df['PLAN SHIP DATE'] == date].copy()
        
        trailers = []
        for trailer in date_data['TRAILER NO'].dropna().unique():
            trailer_data = date_data[date_data['TRAILER NO'] == trailer]
            
            # Get temp trailer range for this trailer number
            temp_trailers_list = trailer_data['TEMP.TRAILER'].dropna().unique().tolist()
            temp_trailers_range = compress_trailer_ranges(temp_trailers_list)
            
            # Format plan date and time as 'Jan 15 HH:MM'
            if len(trailer_data) > 0:
                plan_date_obj = pd.Timestamp(date)
                plan_time_str = str(trailer_data['PLAN SHIP TIME'].iloc[0]) if pd.notna(trailer_data['PLAN SHIP TIME'].iloc[0]) else ''
                # Convert HH.MM.SS to HH:MM
                if plan_time_str and '.' in plan_time_str:
                    parts = plan_time_str.split('.')
                    plan_time_formatted = f"{parts[0]}:{parts[1]}"
                else:
                    plan_time_formatted = plan_time_str
                plan_datetime = f"{plan_date_obj.strftime('%b %d')} {plan_time_formatted}".strip()
            else:
                plan_datetime = 'N/A'
            
            # Split trailer rows by SHIPMENT LOAD DATE so mixed-date trailers
            # show separate detail rows (e.g. 180 on-time + 37 late).
            shipped_rows = trailer_data[trailer_data['SHIPMENT LOAD DATE'].notna()]
            unshipped_rows = trailer_data[trailer_data['SHIPMENT LOAD DATE'].isna()]
            
            plan_date_only = pd.Timestamp(date).date()
            
            if len(shipped_rows) > 0:
                for ship_date_val, group in shipped_rows.groupby('SHIPMENT LOAD DATE'):
                    ship_date_obj = pd.Timestamp(ship_date_val)
                    actual_date_formatted = ship_date_obj.strftime('%b %d')
                    # Use the first row's time for this ship-date group
                    actual_time_str = str(group['SHIPMENT LOAD TIME'].iloc[0]) if pd.notna(group['SHIPMENT LOAD TIME'].iloc[0]) else ''
                    if actual_time_str and '.' in actual_time_str:
                        t_parts = actual_time_str.split('.')
                        actual_time_formatted = f"{t_parts[0]}:{t_parts[1]}"
                    else:
                        actual_time_formatted = actual_time_str
                    actual_datetime = f"{actual_date_formatted} {actual_time_formatted}".strip()
                    
                    ship_date_only = ship_date_obj.date()
                    if ship_date_only < plan_date_only:
                        status = 'early'
                    elif ship_date_only > plan_date_only:
                        status = 'late'
                    else:
                        status = 'ontime'
                    
                    total_parts = int(group['QTY'].sum())
                    total_modules = int(group['MODULE NO'].nunique())
                    
                    trailers.append({
                        'trailer': str(trailer),
                        'temp_trailers': temp_trailers_range,
                        'plan_datetime': plan_datetime,
                        'actual_datetime': actual_datetime,
                        'modules': total_modules,
                        'parts': total_parts,
                        'status': status
                    })
            
            if len(unshipped_rows) > 0:
                total_parts = int(unshipped_rows['QTY'].sum())
                total_modules = int(unshipped_rows['MODULE NO'].nunique())
                trailers.append({
                    'trailer': str(trailer),
                    'temp_trailers': temp_trailers_range,
                    'plan_datetime': plan_datetime,
                    'actual_datetime': 'Not Shipped',
                    'modules': total_modules,
                    'parts': total_parts,
                    'status': 'notshipped'
                })
        
        trailer_details[date_str] = trailers
    
    trailer_details_json = json.dumps(trailer_details)
    
    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Shipping KPI Report - {stats_month['period_label']}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: Arial, Helvetica, sans-serif;
            background: white;
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
            background: linear-gradient(135deg, #f1f3f5 0%, #dee2e6 100%);
            color: #2f3b45;
            padding: 18px 30px;
            text-align: center;
            border-bottom: 1px solid #cfd4da;
        }}

        .header h1 {{
            font-size: 1.6em;
            margin-bottom: 4px;
            font-weight: 600;
            letter-spacing: 0.6px;
        }}

        .header p {{
            font-size: 0.9em;
            color: #4f5b66;
        }}
        
        .filter-bar {{
            display: flex;
            flex-direction: column;
            gap: 14px;
            padding: 18px 30px 20px;
            background: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
        }}
        .filter-row {{
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 20px;
        }}
        .filter-row.primary   {{ justify-content: center; }}
        .filter-row.secondary {{ justify-content: center; color: #495057; }}
        .filter-group {{
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 8px;
        }}
        .filter-label {{
            font-size: 0.8em;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            color: #6c757d;
            margin-right: 4px;
        }}
        .filter-sep {{ color: #6c757d; font-size: 0.9em; }}
        .filter-group input[type="date"] {{
            padding: 6px 10px;
            border: 1px solid #bdc3c7;
            font-size: 0.9em;
            outline: none;
            transition: border-color 0.2s ease, box-shadow 0.2s ease;
        }}
        .filter-group input[type="date"]:focus {{
            border-color: #3498db;
            box-shadow: 0 0 0 2px rgba(52, 152, 219, 0.15);
        }}

        .toggle-btn {{
            padding: 10px 22px;
            border: 2px solid #3498db;
            background: white;
            color: #3498db;
            cursor: pointer;
            font-size: 0.9em;
            font-weight: 600;
            transition: all 0.2s ease;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .toggle-btn:hover {{
            background: #e3f2fd;
            transform: translateY(-1px);
            box-shadow: 0 3px 6px rgba(52, 152, 219, 0.2);
        }}
        .toggle-btn.active {{ background: #3498db; color: white; }}
        .toggle-btn.hidden {{ display: none; }}
        .toggle-btn.ghost {{
            border-color: #ced4da;
            color: #6c757d;
            background: transparent;
            letter-spacing: 0.3px;
        }}
        .toggle-btn.ghost:hover {{
            background: #e9ecef;
            color: #495057;
            border-color: #adb5bd;
            box-shadow: none;
            transform: none;
        }}
        .toggle-btn.ghost .caret {{
            display: inline-block;
            margin-left: 4px;
            transition: transform 0.2s ease;
        }}
        .toggle-btn.ghost[aria-expanded="true"] .caret {{ transform: rotate(180deg); }}

        .advanced-filters[hidden] {{ display: none; }}
        .advanced-filters {{
            animation: slideDown 0.2s ease;
            padding-top: 4px;
            border-top: 1px dashed #dee2e6;
        }}
        @keyframes slideDown {{
            from {{ opacity: 0; transform: translateY(-4px); }}
            to   {{ opacity: 1; transform: translateY(0); }}
        }}

        .quick-btn {{
            padding: 6px 14px;
            font-size: 0.85em;
            border: 1px solid #bdc3c7;
            background: white;
            color: #495057;
            cursor: pointer;
            font-weight: 500;
            transition: all 0.15s ease;
        }}
        .quick-btn:hover {{ background: #3498db; color: white; border-color: #3498db; }}

        .apply-btn {{
            padding: 6px 14px;
            font-size: 0.85em;
            border: 1px solid #6c757d;
            background: #6c757d;
            color: white;
            cursor: pointer;
            font-weight: 600;
            letter-spacing: 0.3px;
            transition: all 0.2s ease;
        }}
        .apply-btn:hover {{ background: #495057; border-color: #495057; }}
        
        .stats-container {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }}
        
        .stat-card {{
            background: white;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 15px rgba(0,0,0,0.2);
        }}
        
        .stat-card h3 {{
            color: #7f8c8d;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
        }}
        
        .stat-card .value {{
            font-size: 2.5em;
            font-weight: bold;
            color: #2c3e50;
        }}

        .stat-card.success .value {{
            color: #28a745;
        }}

        .stat-card.danger .value {{
            color: #dc3545;
        }}

        .stat-card.early .value {{
            color: #5e9c63;
        }}

        .stat-card.info .value {{
            color: #17a2b8;
        }}

        .mom-line {{
            margin-top: 6px;
            font-size: 0.85em;
            color: #4f5b66;
        }}

        .mom-up {{ color: #1f7a3a; font-weight: 600; }}
        .mom-down {{ color: #b3261e; font-weight: 600; }}
        .mom-flat {{ color: #6c757d; font-weight: 600; }}

        .stat-card .subtext {{
            color: #95a5a6;
            font-size: 0.9em;
            margin-top: 5px;
        }}
        
        .chart-container {{
            padding: 30px;
            background: white;
        }}

        .chart-container > div {{
            width: 100%;
        }}
        
        .details-container {{
            padding: 30px;
            background: #f8f9fa;
        }}
        
        .details-container h2 {{
            color: #2c3e50;
            font-size: 1.1em;
            margin-bottom: 14px;
            padding-bottom: 8px;
            border-bottom: 2px solid #3498db;
            letter-spacing: 0.4px;
        }}
        
        .view-content {{
            display: none;
        }}
        
        .view-content.active {{
            display: block;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        
        th {{
            background: #95a5a6;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.85em;
            letter-spacing: 0.5px;
        }}
        
        td {{
            padding: 12px 15px;
            border-bottom: 1px solid #ecf0f1;
        }}
        
        tr:hover {{
            background: #f8f9fa;
        }}
        
        tfoot {{
            background: #e9ecef;
            font-weight: bold;
            border-top: 3px solid #95a5a6;
        }}
        
        tfoot td {{
            padding: 15px;
            font-size: 1.05em;
            border-bottom: none;
        }}
        
        tr:last-child td {{
            border-bottom: none;
        }}

        th.arrow-col,
        td.expand-cell {{
            width: 28px;
            text-align: center;
            color: #7f8c8d;
        }}

        .expand-arrow {{
            display: inline-block;
            font-size: 0.9em;
            transition: transform 0.2s ease, color 0.2s ease;
        }}

        .data-row.expanded .expand-arrow {{
            transform: rotate(180deg);
            color: #2c3e50;
        }}

        .detail-row td {{
            padding: 0;
            border-bottom: none;
        }}

        .detail-content {{
            max-height: 0;
            opacity: 0;
            overflow: hidden;
            transform: translateY(-6px);
            transition: max-height 0.3s ease, opacity 0.3s ease, transform 0.3s ease, padding 0.3s ease;
            background: #f8f9fa;
            padding: 0;
        }}

        .detail-row.expanded .detail-content {{
            max-height: 600px;
            opacity: 1;
            transform: translateY(0);
            padding: 20px;
        }}
            
            .badge {{
                display: inline-block;
                padding: 5px 12px;
                font-size: 0.85em;
                font-weight: 600;
            }}
            
            .badge-success {{
                background: #d4edda;
                color: #155724;
            }}
            
            .badge-warning {{
                background: #fff3cd;
                color: #856404;
            }}
            
            .badge-danger {{
                background: #f8d7da;
                color: #721c24;
            }}
            
            .badge-secondary {{
                background: #e9ecef;
                color: #6c757d;
            }}
            
            .footer {{
                background: #e9ecef;
                color: #4f5b66;
                text-align: center;
                padding: 20px;
                font-size: 0.9em;
            }}
            
            @media print {{
                body {{
                    background: white;
                    padding: 0;
                }}
                
                .container {{
                    box-shadow: none;
                }}

                .filter-bar {{
                    display: none;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>Example Logistics - Monthly Shipping KPI</h1>
                <p>This dashboard tracks our shipping performance and timeliness.</p>
            </div>
            
            <div class="filter-bar">
                <div class="filter-row primary">
                    <button class="toggle-btn active" onclick="switchView('month')" id="btn-month">{stats_month['month_name']}</button>
                    <button class="toggle-btn" onclick="switchView('past2months')" id="btn-past2months">Past 2 Months</button>
                    <button class="toggle-btn hidden" onclick="switchView('custom')" id="btn-custom-view">Custom Range</button>
                    <button class="toggle-btn ghost" onclick="toggleAdvancedFilters()" id="btn-advanced" aria-expanded="false" aria-controls="advanced-filters">More ranges <span class="caret">&#9662;</span></button>
                </div>
                <div class="filter-row secondary advanced-filters" id="advanced-filters" hidden>
                    <div class="filter-group">
                        <span class="filter-label">Quick pick</span>
                        <button class="quick-btn" onclick="applyQuickRange(7)">Last 7 days</button>
                        <button class="quick-btn" onclick="applyQuickRange(30)">Last 30 days</button>
                        <button class="quick-btn" onclick="applyQuickRange(60)">Last 60 days</button>
                    </div>
                    <div class="filter-group">
                        <span class="filter-label">Custom</span>
                        <input type="date" id="start-date" aria-label="Start date" />
                        <span class="filter-sep">to</span>
                        <input type="date" id="end-date" aria-label="End date" />
                        <button class="apply-btn" onclick="applyCustomRange()" id="btn-custom-apply">Apply</button>
                    </div>
                </div>
            </div>
            
            <div id="view-month" class="view-content active">
                {_render_stats_block(stats_month)}
                <div class="chart-container">
                    {chart_month_html}
                </div>
                <div class="details-container">
                    <h2>Daily Shipping Details</h2>
                    <p style="color: #FF0000; font-size: 20px; margin-bottom: 15px;">Click on any row to see detailed trailer information for a specific day!</p>
                    <table>
                        <thead>
                            <tr>
                                <th class="arrow-col"></th>
                                <th>Date</th>
                                <th>Day of Week</th>
                                <th>Plan Total Trailers</th>
                                <th>Plan Total Modules</th>
                                <th>On Time</th>
                                <th>Early</th>
                                <th>Late</th>
                                <th>Met Plan %</th>
                            </tr>
                        </thead>
                        <tbody>
"""
    
    for idx, row in daily_summary_month.iterrows():
        date_str = row['PLAN SHIP DATE'].strftime('%b %d')
        day_name = row['PLAN SHIP DATE'].strftime('%A')
        ontime_display = int(row['on_time_modules'])
        early_display = int(row.get('early_modules', 0))
        late_display = int(row['late_modules'])
        met_plan_pct = float(row.get('met_plan_pct', 0.0))

        html_content += f"""
                    <tr class="data-row" data-date="{row['PLAN SHIP DATE'].strftime('%Y-%m-%d')}" style="cursor: pointer;" aria-expanded="false">
                        <td class="expand-cell"><span class="expand-arrow">▾</span></td>
                        <td><strong>{date_str}</strong></td>
                        <td>{day_name}</td>
                        <td>{int(row['total_trailers'])}</td>
                        <td>{int(row['total_modules']):,}</td>
                        <td style="color: #28a745; font-weight: bold;">{ontime_display}</td>
                        <td style="color: #5e9c63; font-weight: bold;">{early_display}</td>
                        <td style="color: #dc3545; font-weight: bold;">{late_display}</td>
                        <td><strong>{met_plan_pct:.1f}%</strong></td>
                    </tr>
                    <tr class="detail-row" id="month-detail-{row['PLAN SHIP DATE'].strftime('%Y-%m-%d')}">
                        <td colspan="9">
                            <div class="detail-content">
                                <div class="trailer-details">Loading trailer details...</div>
                            </div>
                        </td>
                    </tr>
"""

    total_all_trailers = daily_summary_month['total_trailers'].sum()
    total_all_modules = daily_summary_month['total_modules'].sum()
    total_on_time_modules = daily_summary_month['on_time_modules'].sum()
    total_early_modules = daily_summary_month['early_modules'].sum() if 'early_modules' in daily_summary_month.columns else 0
    total_late_modules = daily_summary_month['late_modules'].sum()
    overall_pct = ((total_on_time_modules + total_early_modules) / total_all_modules * 100) if total_all_modules > 0 else 0

    html_content += f"""
                    </tbody>
                    <tfoot>
                        <tr>
                            <td></td>
                            <td><strong>TOTAL</strong></td>
                            <td></td>
                            <td>{int(total_all_trailers)}</td>
                            <td>{int(total_all_modules):,}</td>
                            <td>{int(total_on_time_modules):,}</td>
                            <td>{int(total_early_modules):,}</td>
                            <td>{int(total_late_modules):,}</td>
                            <td><strong>{overall_pct:.1f}%</strong></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <div id="view-past2months" class="view-content">
            {_render_stats_block(stats_past2months)}
            <div class="chart-container">
                {chart_past2months_html}
            </div>
            <div class="details-container">
                <h2>Daily Shipping Details</h2>
                <p style="color: #FF0000; font-size: 20px; margin-bottom: 15px;">Click on any row to see detailed trailer information for a specific day!</p>
                <table>
                    <thead>
                        <tr>
                            <th class="arrow-col"></th>
                            <th>Date</th>
                            <th>Day of Week</th>
                            <th>Plan Total Trailers</th>
                            <th>Plan Total Modules</th>
                            <th>On Time</th>
                            <th>Early</th>
                            <th>Late</th>
                            <th>Met Plan %</th>
                        </tr>
                    </thead>
                    <tbody>
"""
    
    for idx, row in daily_summary_past2months.iterrows():
        date_str = row['PLAN SHIP DATE'].strftime('%b %d')
        day_name = row['PLAN SHIP DATE'].strftime('%A')
        ontime_display = int(row['on_time_modules'])
        early_display = int(row.get('early_modules', 0))
        late_display = int(row['late_modules'])
        met_plan_pct = float(row.get('met_plan_pct', 0.0))

        html_content += f"""
                        <tr class="data-row" data-date="{row['PLAN SHIP DATE'].strftime('%Y-%m-%d')}" style="cursor: pointer;" aria-expanded="false">
                            <td class="expand-cell"><span class="expand-arrow">▾</span></td>
                            <td><strong>{date_str}</strong></td>
                            <td>{day_name}</td>
                            <td>{int(row['total_trailers'])}</td>
                            <td>{int(row['total_modules']):,}</td>
                            <td style="color: #28a745; font-weight: bold;">{ontime_display}</td>
                            <td style="color: #5e9c63; font-weight: bold;">{early_display}</td>
                            <td style="color: #dc3545; font-weight: bold;">{late_display}</td>
                            <td><strong>{met_plan_pct:.1f}%</strong></td>
                        </tr>
                        <tr class="detail-row" id="past2months-detail-{row['PLAN SHIP DATE'].strftime('%Y-%m-%d')}">
                            <td colspan="9">
                                <div class="detail-content">
                                    <div class="trailer-details">Loading trailer details...</div>
                                </div>
                            </td>
                        </tr>
"""

    total_all_trailers_2m = daily_summary_past2months['total_trailers'].sum()
    total_all_modules_2m = daily_summary_past2months['total_modules'].sum()
    total_on_time_modules_2m = daily_summary_past2months['on_time_modules'].sum()
    total_early_modules_2m = daily_summary_past2months['early_modules'].sum() if 'early_modules' in daily_summary_past2months.columns else 0
    total_late_modules_2m = daily_summary_past2months['late_modules'].sum()
    overall_pct_2m = ((total_on_time_modules_2m + total_early_modules_2m) / total_all_modules_2m * 100) if total_all_modules_2m > 0 else 0

    html_content += f"""
                    </tbody>
                    <tfoot>
                        <tr>
                            <td></td>
                            <td><strong>TOTAL</strong></td>
                            <td></td>
                            <td>{int(total_all_trailers_2m)}</td>
                            <td>{int(total_all_modules_2m):,}</td>
                            <td>{int(total_on_time_modules_2m):,}</td>
                            <td>{int(total_early_modules_2m):,}</td>
                            <td>{int(total_late_modules_2m):,}</td>
                            <td><strong>{overall_pct_2m:.1f}%</strong></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <div class="footer">
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
            <p style="margin-top: 10px; font-size: 0.9em; color: #7f8c8d;">Dashboard created by Viktor Berg | Built with Python, Plotly, and Pandas</p>
            {_render_load_info(load_info)}
        </div>
        
        <script>
            const allData = {all_data_json};
            const trailerDetails = {trailer_details_json};
            
            // Handle row clicks to show/hide trailer details using event delegation
            document.addEventListener('DOMContentLoaded', function() {{
                // Use event delegation to handle clicks on both static and dynamically created rows
                document.body.addEventListener('click', function(e) {{
                    // Find if the clicked element or its parent is a data-row
                    const row = e.target.closest('.data-row');
                    if (!row) return;
                    
                    const date = row.getAttribute('data-date');
                    
                    // Determine which view the row belongs to
                    let viewPrefix = '';
                    const viewContainer = row.closest('.view-content');
                    if (viewContainer) {{
                        if (viewContainer.id === 'view-month') {{
                            viewPrefix = 'month-';
                        }} else if (viewContainer.id === 'view-last31') {{
                            viewPrefix = 'last31-';
                        }} else if (viewContainer.id === 'view-custom') {{
                            viewPrefix = 'custom-';
                        }}
                    }}
                    
                    const detailId = viewPrefix + 'detail-' + date;
                    const detailRow = document.getElementById(detailId);
                    
                    if (!detailRow) {{
                        console.error('Detail row not found. View:', viewPrefix, 'Date:', date, 'ID:', detailId);
                        return;
                    }}
                    
                    // Toggle display: if not showing (none or empty), show it; otherwise hide it
                    const isExpanded = detailRow.classList.contains('expanded');
                    if (!isExpanded) {{
                        // Show details
                        const details = trailerDetails[date] || [];
                        console.log('Trailer details for', date, ':', details.length, 'trailers');
                        
                        let detailHTML = '<table style="width: 100%; margin: 10px 0;"><thead><tr style="background: #34495e; color: white;"><th>Trailer Number</th><th>Temp Trailers</th><th>Plan Ship Date & Time</th><th>Actual Ship Date & Time</th><th>Total Modules</th></tr></thead><tbody>';
                        
                        if (details.length === 0) {{
                            detailHTML += '<tr><td colspan="5" style="text-align: center; padding: 20px; color: #7f8c8d;">No trailer data available for this date</td></tr>';
                        }} else {{
                            details.forEach(trailer => {{
                                let rowColor = '';
                                if (trailer.status === 'early') {{
                                    rowColor = 'background-color: #d4edda;'; // Light green
                                }} else if (trailer.status === 'ontime') {{
                                    rowColor = 'background-color: #d4edda;'; // Light green
                                }} else if (trailer.status === 'late') {{
                                    rowColor = 'background-color: #f8d7da;'; // Light red
                                }}
                                detailHTML += `<tr style="${{rowColor}}"><td>${{trailer.trailer}}</td><td>${{trailer.temp_trailers || ''}}</td><td>${{trailer.plan_datetime}}</td><td>${{trailer.actual_datetime}}</td><td>${{trailer.modules}}</td></tr>`;
                            }});
                        }}
                        
                        detailHTML += '</tbody></table>';
                        detailRow.querySelector('.trailer-details').innerHTML = detailHTML;
                        detailRow.classList.add('expanded');
                        row.classList.add('expanded');
                        row.setAttribute('aria-expanded', 'true');
                        console.log('Details shown for', date);
                    }} else {{
                        // Hide details
                        detailRow.classList.remove('expanded');
                        row.classList.remove('expanded');
                        row.setAttribute('aria-expanded', 'false');
                        console.log('Details hidden for', date);
                    }}
                }});
            }});
            
            function switchView(mode) {{
                const views = ['view-month', 'view-past2months', 'view-custom'];
                const btns  = ['btn-month', 'btn-past2months', 'btn-custom-view'];
                views.forEach(id => {{ const el = document.getElementById(id); if (el) el.classList.remove('active'); }});
                btns.forEach(id  => {{ const el = document.getElementById(id); if (el) el.classList.remove('active'); }});

                const targetMap = {{month: ['view-month','btn-month'], past2months: ['view-past2months','btn-past2months'], custom: ['view-custom','btn-custom-view']}};
                const pair = targetMap[mode];
                if (!pair) return;
                pair.forEach(id => {{ const el = document.getElementById(id); if (el) el.classList.add('active'); }});

                // Charts inside hidden views were laid out at 0×0 — force a resize
                // now that the view is visible so they fill the container width.
                const activeView = document.getElementById(pair[0]);
                if (activeView && window.Plotly) {{
                    activeView.querySelectorAll('.js-plotly-plot').forEach(el => Plotly.Plots.resize(el));
                }}
            }}

            function toggleAdvancedFilters() {{
                const panel = document.getElementById('advanced-filters');
                const btn = document.getElementById('btn-advanced');
                const hidden = panel.hasAttribute('hidden');
                if (hidden) {{
                    panel.removeAttribute('hidden');
                    btn.setAttribute('aria-expanded', 'true');
                }} else {{
                    panel.setAttribute('hidden', '');
                    btn.setAttribute('aria-expanded', 'false');
                }}
            }}

            function applyCustomRange() {{
                const startDate = document.getElementById('start-date').value;
                const endDate = document.getElementById('end-date').value;

                if (!startDate || !endDate) {{
                    alert('Please select both start and end dates.');
                    return;
                }}

                if (parseDate(startDate) > parseDate(endDate)) {{
                    alert('Start date must be before end date.');
                    return;
                }}

                let customView = document.getElementById('view-custom');
                if (!customView) {{
                    customView = document.createElement('div');
                    customView.id = 'view-custom';
                    customView.className = 'view-content';
                    document.querySelector('.footer').before(customView);
                }}

                filterDataByRange(startDate, endDate);
                document.getElementById('btn-custom-view').classList.remove('hidden');
                switchView('custom');
            }}

            function parseDate(dateStr) {{
                const [year, month, day] = dateStr.split('-').map(Number);
                return new Date(year, month - 1, day);
            }}

            function formatDateISO(d) {{
                const y = d.getFullYear();
                const m = String(d.getMonth() + 1).padStart(2, '0');
                const day = String(d.getDate()).padStart(2, '0');
                return `${{y}}-${{m}}-${{day}}`;
            }}

            function latestDataDate() {{
                if (!allData || allData.length === 0) return new Date();
                for (let i = allData.length - 1; i >= 0; i--) {{
                    if (allData[i].total_modules > 0) return parseDate(allData[i].date);
                }}
                return parseDate(allData[allData.length - 1].date);
            }}

            function applyRangeDates(startStr, endStr) {{
                document.getElementById('start-date').value = startStr;
                document.getElementById('end-date').value = endStr;
                applyCustomRange();
            }}

            function applyQuickRange(days) {{
                const end = latestDataDate();
                const start = new Date(end);
                start.setDate(end.getDate() - (days - 1));
                applyRangeDates(formatDateISO(start), formatDateISO(end));
            }}

            function filterDataByRange(startDate, endDate) {{
                const start = parseDate(startDate);
                const end = parseDate(endDate);
                const filteredData = allData.filter(item => {{
                    const itemDate = parseDate(item.date);
                    return itemDate >= start && itemDate <= end;
                }}).sort((a, b) => parseDate(a.date) - parseDate(b.date));

                if (filteredData.length === 0) {{
                    alert('No data available for the selected date range.');
                    return;
                }}

                const totalTrailers = filteredData.reduce((sum, item) => sum + item.total_trailers, 0);
                const totalModules = filteredData.reduce((sum, item) => sum + item.total_modules, 0);
                const onTimeModules = filteredData.reduce((sum, item) => sum + (item.on_time_modules || 0), 0);
                const earlyModules = filteredData.reduce((sum, item) => sum + (item.early_modules || 0), 0);
                const lateModules = filteredData.reduce((sum, item) => sum + (item.late_modules || 0), 0);
                const metPlanModules = onTimeModules + earlyModules;
                const onTimeModulesPct = totalModules > 0 ? (onTimeModules / totalModules * 100) : 0;
                const earlyModulesPct = totalModules > 0 ? (earlyModules / totalModules * 100) : 0;
                const lateModulesPct = totalModules > 0 ? (lateModules / totalModules * 100) : 0;
                const metPlanPct = totalModules > 0 ? (metPlanModules / totalModules * 100) : 0;
                const avgTotalModules = totalModules / filteredData.length;

                function formatDate(dateStr) {{
                    const d = parseDate(dateStr);
                    return d.toLocaleDateString('en-US', {{month: 'short', day: 'numeric', year: 'numeric'}});
                }};

                function formatDay(dateStr) {{
                    const d = parseDate(dateStr);
                    return d.toLocaleDateString('en-US', {{weekday: 'long'}});
                }};

                const periodLabel = `${{formatDate(startDate)}} - ${{formatDate(endDate)}}`;

                const dates = filteredData.map(item => parseDate(item.date));
                const onTimeData = filteredData.map(item => item.on_time_modules || 0);
                const earlyData = filteredData.map(item => item.early_modules || 0);
                const lateData = filteredData.map(item => item.late_modules || 0);
                const onTimePctLabels = filteredData.map(item => {{
                    const t = item.total_modules || 0;
                    const n = item.on_time_modules || 0;
                    return (t > 0 && n > 0) ? `${{(n / t * 100).toFixed(1)}}%` : '';
                }});

                function pctOf(n, t) {{ return t > 0 ? (n / t * 100) : 0; }}
                function hoverFor(n, t) {{
                    if (t <= 0) return '—';
                    return `${{n.toLocaleString()}} modules (${{pctOf(n, t).toFixed(1)}}%)`;
                }}
                const earlyHover = filteredData.map(item => hoverFor(item.early_modules || 0, item.total_modules || 0));
                const onTimeHover = filteredData.map(item => hoverFor(item.on_time_modules || 0, item.total_modules || 0));
                const lateHover = filteredData.map(item => hoverFor(item.late_modules || 0, item.total_modules || 0));

                // Rolling 7-day met-plan % over active days only.
                const rollingPct = [];
                const window = [];
                for (const item of filteredData) {{
                    const t = item.total_modules || 0;
                    if (t > 0) {{
                        const m = (item.on_time_modules || 0) + (item.early_modules || 0);
                        window.push([m, t]);
                        while (window.length > 7) window.shift();
                        const wm = window.reduce((s, p) => s + p[0], 0);
                        const wt = window.reduce((s, p) => s + p[1], 0);
                        rollingPct.push(wt > 0 ? Number((wm / wt * 100).toFixed(1)) : null);
                    }} else {{
                        rollingPct.push(null);
                    }}
                }}

                const earlyTrace = {{
                    x: dates, y: earlyData, type: 'bar',
                    name: 'Shipped Early', marker: {{color: '#7bc47f'}},
                    hovertemplate: '<b>Early</b>: %{{customdata}}<extra></extra>',
                    customdata: earlyHover, showlegend: true
                }};
                const onTimeTrace = {{
                    x: dates, y: onTimeData, type: 'bar',
                    name: 'Shipped on Ship Date', marker: {{color: '#28a745'}},
                    text: onTimePctLabels, textposition: 'inside',
                    textfont: {{color: 'white', size: 10, family: 'Arial'}},
                    hovertemplate: '<b>On Time</b>: %{{customdata}}<extra></extra>',
                    customdata: onTimeHover, showlegend: true
                }};
                const lateTrace = {{
                    x: dates, y: lateData, type: 'bar',
                    name: 'Shipped Late', marker: {{color: '#dc3545'}},
                    hovertemplate: '<b>Late</b>: %{{customdata}}<extra></extra>',
                    customdata: lateHover, showlegend: true
                }};
                const rollingTrace = {{
                    x: dates, y: rollingPct, type: 'scatter',
                    name: 'Rolling 7-day Met-Plan %',
                    mode: 'lines+markers', yaxis: 'y2',
                    line: {{color: '#2c3e50', width: 2, dash: 'dot'}},
                    marker: {{size: 6, color: '#2c3e50'}},
                    hovertemplate: '<b>Rolling 7d</b>: %{{y:.1f}}%<extra></extra>',
                    connectgaps: false, showlegend: true
                }};
                const targetTrace = {{
                    x: [dates[0], dates[dates.length - 1]],
                    y: [{ON_TIME_TARGET_PCT}, {ON_TIME_TARGET_PCT}],
                    type: 'scatter', mode: 'lines',
                    name: 'Target {int(ON_TIME_TARGET_PCT)}%', yaxis: 'y2',
                    line: {{color: '#888', width: 1.5, dash: 'dash'}},
                    hoverinfo: 'skip', showlegend: true
                }};

                const layout = {{
                    title: {{
                        text: `Shipping KPI - ${{periodLabel}}`,
                        font: {{size: 24, color: '#2c3e50'}},
                        x: 0.5, xanchor: 'center'
                    }},
                    xaxis: {{
                        title: {{text: 'Plan Ship Date'}},
                        tickformat: '%b %d', dtick: 86400000, tickangle: -45,
                        showgrid: true, gridcolor: '#ecf0f1',
                        linecolor: '#bdc3c7', linewidth: 2,
                        rangebreaks: [{{bounds: ["sat", "mon"]}}]
                    }},
                    yaxis: {{
                        title: {{text: 'Module Count'}},
                        showgrid: true, gridcolor: '#ecf0f1',
                        linecolor: '#bdc3c7', linewidth: 2
                    }},
                    yaxis2: {{
                        title: {{text: 'Met-Plan %'}},
                        overlaying: 'y', side: 'right',
                        range: [0, 105], showgrid: false,
                        ticksuffix: '%', linecolor: '#bdc3c7', linewidth: 2
                    }},
                    barmode: 'stack', hovermode: 'x unified',
                    plot_bgcolor: 'white', paper_bgcolor: 'white',
                    height: 600, autosize: true,
                    font: {{family: 'Arial, sans-serif', size: 12, color: '#2c3e50'}},
                    margin: {{l: 80, r: 80, t: 150, b: 100}},
                    legend: {{
                        orientation: 'h', yanchor: 'bottom', y: 1.02,
                        xanchor: 'center', x: 0.5, font: {{size: 14}}
                    }}
                }};

                let bodyRows = '';
                for (const item of filteredData) {{
                    const t = item.total_modules || 0;
                    const n = item.on_time_modules || 0;
                    const e = item.early_modules || 0;
                    const l = item.late_modules || 0;
                    const mpct = t > 0 ? ((n + e) / t * 100) : 0;
                    bodyRows += `
                        <tr class="data-row" data-date="${{item.date}}" style="cursor: pointer;" aria-expanded="false">
                            <td class="expand-cell"><span class="expand-arrow">▾</span></td>
                            <td><strong>${{formatDate(item.date)}}</strong></td>
                            <td>${{formatDay(item.date)}}</td>
                            <td>${{(item.total_trailers || 0).toLocaleString()}}</td>
                            <td>${{t.toLocaleString()}}</td>
                            <td style="color: #28a745; font-weight: bold;">${{n.toLocaleString()}}</td>
                            <td style="color: #5e9c63; font-weight: bold;">${{e.toLocaleString()}}</td>
                            <td style="color: #dc3545; font-weight: bold;">${{l.toLocaleString()}}</td>
                            <td><strong>${{mpct.toFixed(1)}}%</strong></td>
                        </tr>
                        <tr class="detail-row" id="custom-detail-${{item.date}}">
                            <td colspan="9">
                                <div class="detail-content">
                                    <div class="trailer-details">Loading trailer details...</div>
                                </div>
                            </td>
                        </tr>`;
                }}

                document.getElementById('view-custom').innerHTML = `
                    <div class="stats-container">
                        <div class="stat-card">
                            <h3>Period</h3>
                            <div class="value" style="font-size: 1.5em;">${{periodLabel}}</div>
                        </div>
                        <div class="stat-card">
                            <h3>TOTAL MODULES</h3>
                            <div class="value">${{totalModules.toLocaleString()}}</div>
                            <div class="subtext">${{totalTrailers.toLocaleString()}} trailers | Avg. ${{avgTotalModules.toFixed(0)}} per day</div>
                        </div>
                        <div class="stat-card success">
                            <h3>MET PLAN (ON-TIME OR EARLY)</h3>
                            <div class="value">${{metPlanPct.toFixed(1)}}%</div>
                            <div class="subtext">${{metPlanModules.toLocaleString()}} modules &middot; target {int(ON_TIME_TARGET_PCT)}%</div>
                        </div>
                        <div class="stat-card danger">
                            <h3>SHIPPED LATE</h3>
                            <div class="value">${{lateModulesPct.toFixed(1)}}%</div>
                            <div class="subtext">${{lateModules.toLocaleString()}} modules</div>
                        </div>
                    </div>
                    <div class="chart-container">
                        <div id="custom-chart"></div>
                    </div>
                    <div class="details-container">
                        <h2>Daily Shipping Details</h2>
                        <p style="color: #FF0000; font-size: 20px; margin-bottom: 15px;">Click on any row to see detailed trailer information for a specific day!</p>
                        <table>
                            <thead>
                                <tr>
                                    <th class="arrow-col"></th>
                                    <th>Date</th>
                                    <th>Day of Week</th>
                                    <th>Plan Total Trailers</th>
                                    <th>Plan Total Modules</th>
                                    <th>On Time</th>
                                    <th>Early</th>
                                    <th>Late</th>
                                    <th>Met Plan %</th>
                                </tr>
                            </thead>
                            <tbody>${{bodyRows}}</tbody>
                            <tfoot>
                                <tr>
                                    <td></td>
                                    <td><strong>TOTAL</strong></td>
                                    <td></td>
                                    <td>${{totalTrailers.toLocaleString()}}</td>
                                    <td>${{totalModules.toLocaleString()}}</td>
                                    <td>${{onTimeModules.toLocaleString()}}</td>
                                    <td>${{earlyModules.toLocaleString()}}</td>
                                    <td>${{lateModules.toLocaleString()}}</td>
                                    <td><strong>${{metPlanPct.toFixed(1)}}%</strong></td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                `;

                Plotly.newPlot('custom-chart', [earlyTrace, onTimeTrace, lateTrace, rollingTrace, targetTrace], layout, {{
                    displayModeBar: false,
                    responsive: true
                }});
            }}
        </script>
    </body>
</html>
"""
    
    return html_content

def main():
    parser = argparse.ArgumentParser(
        description='Generate Shipping KPI Report with interactive month/past 2 months toggle',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        '--data-folder',
        default='Data',
        help='Path to the data folder (default: Data)'
    )
    
    parser.add_argument(
        '--month',
        type=int,
        default=None,
        help='Target month (1-12)'
    )
    
    parser.add_argument(
        '--year',
        type=int,
        default=None,
        help='Target year (e.g. 2026)'
    )
    
    args = parser.parse_args()

    data_folder = args.data_folder

    print("Loading and processing data...")
    try:
        df, load_info = load_and_process_data(data_folder)
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)

    print(f"Total records processed: {load_info['usable_rows']:,} "
          f"(raw: {load_info['raw_rows']:,}, dropped: {load_info['invalid_ship_rows']:,})")

    print("\nGenerating calendar month report...")
    daily_summary_month, mode_month, month, year, period_data_month = create_daily_summary_calendar_month(
        df, target_month=args.month, target_year=args.year
    )
    prev_month = create_previous_month_summary(df, month, year)
    fig_month = create_visualization(daily_summary_month, mode_month, (month, year))
    stats_month = generate_summary_stats(
        daily_summary_month, mode_month, (month, year), period_data_month, prev_month=prev_month
    )

    if len(daily_summary_month) > 0:
        start_date = daily_summary_month['PLAN SHIP DATE'].min()
        end_date = daily_summary_month['PLAN SHIP DATE'].max()
        print(f"   Period: {start_date.strftime('%b %d')} - {end_date.strftime('%b %d, %Y')}")
    else:
        print(f"   Period: No data for this month")
    print(f"   Days with data: {len(daily_summary_month)}")
    print(f"   Total Trailers: {stats_month['total_trailers']:,}")
    print(f"   Total Modules: {stats_month['total_modules']:,}")
    print(f"   Met Plan: {stats_month['met_plan_modules']:,} ({stats_month['met_plan_pct']:.1f}%)")
    print(f"      On-Time: {stats_month['on_time_modules']:,} ({stats_month['on_time_modules_pct']:.1f}%)")
    print(f"      Early:   {stats_month['early_modules']:,} ({stats_month['early_modules_pct']:.1f}%)")
    print(f"   Late:    {stats_month['late_modules']:,} ({stats_month['late_modules_pct']:.1f}%)")
    if stats_month.get('mom'):
        print(f"   vs {stats_month['mom']['prev_label']}: {stats_month['mom']['delta_pts']:+.1f} pts")

    print("\nGenerating past 2 months report...")
    daily_summary_past2months, mode_past2months, p2_start, p2_end, period_data_past2months = create_daily_summary_past_2_months(df)
    period_info_past2months = (p2_start, p2_end)
    fig_past2months = create_visualization(daily_summary_past2months, mode_past2months, period_info_past2months)
    stats_past2months = generate_summary_stats(
        daily_summary_past2months, mode_past2months, period_info_past2months, period_data_past2months
    )

    print(f"   Period: {stats_past2months['period_label']}")
    print(f"   Days with data: {len(daily_summary_past2months)}")
    print(f"   Total Modules: {stats_past2months['total_modules']:,}")
    print(f"   Met Plan: {stats_past2months['met_plan_modules']:,} ({stats_past2months['met_plan_pct']:.1f}%)")

    html_content = create_html_report(
        fig_month, stats_month, daily_summary_month,
        fig_past2months, stats_past2months, daily_summary_past2months,
        df, load_info,
    )

    html_output_file = f"Shipping_KPI_Report_{year}_{month:02d}.html"
    with open(html_output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"\n[OK] HTML report generated: {html_output_file}")
    print(f"   Open the file in your web browser to view the report.")
    print(f"   Use the toggle buttons to switch between Calendar Month and Past 2 Months views.")

if __name__ == '__main__':
    main()
