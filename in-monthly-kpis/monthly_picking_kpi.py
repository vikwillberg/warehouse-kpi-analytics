import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
from pandas.tseries.holiday import USFederalHolidayCalendar
from pandas.tseries.offsets import CustomBusinessDay
import os
import argparse
import json


def _build_company_holidays(start_year=2015, end_year=2035):
    """US federal holidays plus the company year-end shutdown (Dec 23 – Jan 2)."""
    federal = USFederalHolidayCalendar().holidays(
        start=f'{start_year}-01-01', end=f'{end_year}-12-31'
    )
    shutdown = []
    for year in range(start_year, end_year + 1):
        shutdown.extend(pd.date_range(f'{year}-12-23', f'{year}-12-31'))
        shutdown.extend(pd.date_range(f'{year}-01-01', f'{year}-01-02'))
    return pd.DatetimeIndex(sorted(set(federal) | set(shutdown)))


_COMPANY_HOLIDAYS = _build_company_holidays()
_COMPANY_BDAY = CustomBusinessDay(holidays=_COMPANY_HOLIDAYS)
_COMPANY_HOLIDAY_STRS = [d.strftime('%Y-%m-%d') for d in _COMPANY_HOLIDAYS]

ON_TIME_TARGET_PCT = 95.0
REQUIRED_COLUMNS = [
    'TRAILER NO', 'TEMP.TRAILER', 'MODULE NO',
    'PLAN SHIP DATE', 'PICKING DATE', 'PICKING TIME',
]


def get_business_days_before(ship_date, business_days=2):
    """
    Calculate the picking due date by going back N business days from ship date.
    Skips weekends, US federal holidays, and the company year-end shutdown
    (Dec 23 – Jan 2).
    """
    return (pd.Timestamp(ship_date) - business_days * _COMPANY_BDAY).normalize()

def load_and_process_data(data_folder):
    """Load the 202.csv file and process picking KPI data.

    Returns (df, load_info) where load_info summarizes row counts and any
    rows dropped due to invalid ship dates (surfaced in the report footer).
    """

    file_path = os.path.join(data_folder, '202.csv')

    if not os.path.exists(file_path):
        raise FileNotFoundError(
            f"Could not find '{file_path}'. Place the 202.csv export in the "
            f"'{data_folder}' folder (or pass --data-folder)."
        )

    try:
        df = pd.read_csv(file_path)
    except Exception as exc:
        raise RuntimeError(f"Failed to read '{file_path}': {exc}") from exc

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"'{file_path}' is missing required columns: {', '.join(missing)}. "
            f"Found columns: {', '.join(df.columns)}"
        )

    raw_rows = len(df)

    df['PLAN SHIP DATE'] = pd.to_datetime(df['PLAN SHIP DATE'], format='mixed', dayfirst=False, errors='coerce')
    df['PICKING DATE'] = pd.to_datetime(df['PICKING DATE'], format='mixed', dayfirst=False, errors='coerce')

    invalid_ship_rows = int(df['PLAN SHIP DATE'].isna().sum())
    df = df.dropna(subset=['PLAN SHIP DATE'])

    df['PICKING DUE DATE'] = df['PLAN SHIP DATE'].apply(get_business_days_before)

    today = pd.Timestamp.today().normalize()

    df['ON_TIME'] = False
    df['IS_PICKED'] = ~df['PICKING DATE'].isna()

    past_mask = df['PICKING DUE DATE'] <= today
    df.loc[past_mask & df['IS_PICKED'], 'ON_TIME'] = (
        df.loc[past_mask & df['IS_PICKED'], 'PICKING DATE']
        <= df.loc[past_mask & df['IS_PICKED'], 'PICKING DUE DATE']
    )

    df = df[past_mask].copy()
    df['NOT_PICKED_OVERDUE'] = ~df['IS_PICKED']

    load_info = {
        'raw_rows': raw_rows,
        'invalid_ship_rows': invalid_ship_rows,
        'usable_rows': len(df),
        'min_ship_date': df['PLAN SHIP DATE'].min() if not df.empty else None,
        'max_ship_date': df['PLAN SHIP DATE'].max() if not df.empty else None,
    }

    return df, load_info

def fill_missing_dates(daily_summary, date_col, start_date, end_date):
    """Fill in all missing business days in the range with zero-value rows (skips weekends)."""
    count_cols = ['total_trailers', 'total_modules', 'on_time_modules',
                  'late_modules', 'picked_late_modules', 'not_picked_modules']
    pct_cols = ['on_time_pct', 'late_pct']

    if daily_summary.empty:
        all_dates = pd.date_range(start=start_date, end=end_date, freq='B')
        empty = {date_col: all_dates}
        for c in count_cols:
            empty[c] = 0
        for c in pct_cols:
            empty[c] = 0.0
        return pd.DataFrame(empty)

    all_dates = pd.date_range(start=start_date, end=end_date, freq='B')
    full_df = pd.DataFrame({date_col: all_dates})
    result = full_df.merge(daily_summary, on=date_col, how='left')
    for col in count_cols:
        if col in result.columns:
            result[col] = result[col].fillna(0).astype(int)
    for col in pct_cols:
        if col in result.columns:
            result[col] = result[col].fillna(0.0)
    return result

def build_daily_summary(period_data):
    """Build daily summary of picking performance for a provided dataset.

    Splits "late" into two buckets: picked-late (picked after due date) and
    not-picked (past due date with no pick recorded). Percentages are
    computed from raw counts so late_pct + on_time_pct can only drift by
    rounding of at most 0.1 pt on each side, not compound.
    """
    if period_data.empty:
        return pd.DataFrame(columns=[
            'PLAN SHIP DATE', 'total_trailers', 'total_modules',
            'on_time_modules', 'late_modules', 'picked_late_modules',
            'not_picked_modules', 'on_time_pct', 'late_pct',
        ])

    daily_summary = period_data.groupby('PLAN SHIP DATE').agg(
        total_trailers=('TRAILER NO', 'nunique'),
        total_modules=('MODULE NO', 'size'),
        on_time_modules=('ON_TIME', 'sum'),
        not_picked_modules=('NOT_PICKED_OVERDUE', 'sum'),
    ).reset_index()

    daily_summary['late_modules'] = daily_summary['total_modules'] - daily_summary['on_time_modules']
    daily_summary['picked_late_modules'] = daily_summary['late_modules'] - daily_summary['not_picked_modules']

    totals = daily_summary['total_modules'].replace(0, pd.NA)
    daily_summary['on_time_pct'] = (daily_summary['on_time_modules'] / totals * 100).round(1).fillna(0.0)
    daily_summary['late_pct'] = (daily_summary['late_modules'] / totals * 100).round(1).fillna(0.0)

    daily_summary = daily_summary.sort_values('PLAN SHIP DATE').reset_index(drop=True)

    return daily_summary

def create_daily_summary_calendar_month(df, target_month=None, target_year=None):
    """Create daily summary of picking performance for a calendar month.

    Returns (daily_summary, period_data, mode, month, year). `period_data`
    is the raw rows in-scope so callers can compute unique trailer counts.
    """

    if df.empty:
        current = pd.Timestamp.today().normalize()
        return build_daily_summary(df), df, 'month', current.month, current.year

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

    if not daily_summary.empty:
        month_start = pd.Timestamp(year=target_year, month=target_month, day=1)
        month_end = daily_summary['PLAN SHIP DATE'].max()
        daily_summary = fill_missing_dates(daily_summary, 'PLAN SHIP DATE', month_start, month_end)

    return daily_summary, month_data, 'month', target_month, target_year

def create_daily_summary_past_2_months(df):
    """Create daily summary of picking performance for the past 2 months."""

    if df.empty:
        today = pd.Timestamp.today().normalize()
        start_date = today - timedelta(days=60)
        return build_daily_summary(df), df, 'past_2_months', start_date, today

    today = pd.Timestamp.today().normalize()
    latest_available_date = df['PLAN SHIP DATE'].max()
    latest_date = min(latest_available_date, today) if pd.notna(latest_available_date) else today
    start_date = latest_date - timedelta(days=60)

    period_data = df[
        (df['PLAN SHIP DATE'] >= start_date) &
        (df['PLAN SHIP DATE'] <= latest_date)
    ]

    daily_summary = build_daily_summary(period_data)
    daily_summary = fill_missing_dates(daily_summary, 'PLAN SHIP DATE', start_date, latest_date)

    return daily_summary, period_data, 'past_2_months', start_date, latest_date

def create_previous_month_summary(df, target_month, target_year):
    """Build the daily summary for the month before (target_month, target_year),
    used for month-over-month comparison. Returns None if no data."""
    if df.empty:
        return None
    anchor = pd.Timestamp(year=target_year, month=target_month, day=1)
    prev_last = anchor - timedelta(days=1)
    prev_data = df[
        (df['PLAN SHIP DATE'].dt.month == prev_last.month) &
        (df['PLAN SHIP DATE'].dt.year == prev_last.year)
    ]
    if prev_data.empty:
        return None
    return {
        'month': prev_last.month,
        'year': prev_last.year,
        'total_modules': int(prev_data.shape[0]),
        'on_time_modules': int(prev_data['ON_TIME'].sum()),
    }

def _rolling_on_time_pct(daily_summary, window=7):
    """Return a list of rolling on-time % values (window of N active days).

    Only business days that have modules are counted in the window so empty
    filled-in days don't drag the average to zero. Days with no history yet
    return None (gap in the line) instead of 0.
    """
    values = []
    active = []
    for _, row in daily_summary.iterrows():
        if int(row['total_modules']) > 0:
            active.append((int(row['on_time_modules']), int(row['total_modules'])))
            active = active[-window:]
            tot = sum(t for _, t in active)
            on = sum(o for o, _ in active)
            values.append(round(on / tot * 100, 1) if tot > 0 else None)
        else:
            values.append(None)
    return values


def create_visualization(daily_summary, mode, period_info):
    """Create stacked bar chart: on-time, picked-late, not-picked-overdue,
    plus a rolling-7 on-time % line and a 95% target line."""

    fig = go.Figure()
    plot_dates = daily_summary['PLAN SHIP DATE']

    on_time_counts = daily_summary['on_time_modules'].astype(int)
    picked_late_counts = daily_summary['picked_late_modules'].astype(int)
    not_picked_counts = daily_summary['not_picked_modules'].astype(int)
    totals = daily_summary['total_modules'].astype(int)
    max_total = int(totals.max()) if len(totals) else 0

    # Adaptive label positioning: short bars put their % label outside to stay readable.
    label_threshold = max(1, int(max_total * 0.08))
    on_time_text, on_time_positions = [], []
    for on_time, pct in zip(on_time_counts, daily_summary['on_time_pct']):
        if on_time <= 0:
            on_time_text.append('')
            on_time_positions.append('inside')
        elif on_time < label_threshold:
            on_time_text.append(f"{pct:.1f}%")
            on_time_positions.append('outside')
        else:
            on_time_text.append(f"{pct:.1f}%")
            on_time_positions.append('inside')

    # Per-trace hover: one line each under `hovermode='x unified'` so the
    # tooltip reads like a mini-legend instead of repeating the breakdown.
    def pct_of(n, t):
        return (n / t * 100) if t > 0 else 0.0

    # Unified-hover labels: the trace name is suppressed by <extra></extra>,
    # so each line prefixes its own label ("On Time: …") for clarity.
    on_time_hover = [
        f"{int(n):,} modules ({pct_of(n, t):.1f}%)" if t > 0 else "—"
        for n, t in zip(on_time_counts, totals)
    ]
    picked_late_hover = [
        f"{int(n):,} modules ({pct_of(n, t):.1f}%)" if t > 0 else "—"
        for n, t in zip(picked_late_counts, totals)
    ]
    not_picked_hover = [
        f"{int(n):,} modules ({pct_of(n, t):.1f}%)" if t > 0 else "—"
        for n, t in zip(not_picked_counts, totals)
    ]

    fig.add_trace(go.Bar(
        name='Picked On Time',
        x=plot_dates,
        y=on_time_counts,
        marker_color='#28a745',
        text=on_time_text,
        textposition=on_time_positions,
        textfont=dict(color='white', size=10, family='Arial'),
        cliponaxis=False,
        hovertemplate='<b>On Time</b>: %{customdata}<extra></extra>',
        customdata=on_time_hover,
    ))

    fig.add_trace(go.Bar(
        name='Picked Late',
        x=plot_dates,
        y=picked_late_counts,
        marker_color='#dc3545',
        text='',
        hovertemplate='<b>Picked Late</b>: %{customdata}<extra></extra>',
        customdata=picked_late_hover,
    ))

    fig.add_trace(go.Bar(
        name='Not Picked (overdue)',
        x=plot_dates,
        y=not_picked_counts,
        marker_color='#fd7e14',
        text='',
        hovertemplate='<b>Not Picked</b>: %{customdata}<extra></extra>',
        customdata=not_picked_hover,
    ))

    rolling_vals = _rolling_on_time_pct(daily_summary, window=7)
    fig.add_trace(go.Scatter(
        name='Rolling 7-day On-Time %',
        x=plot_dates,
        y=rolling_vals,
        mode='lines+markers',
        line=dict(color='#2c3e50', width=2, dash='dot'),
        marker=dict(size=5),
        yaxis='y2',
        hovertemplate='<b>7-day rolling avg</b>: %{y:.1f}%<extra></extra>',
        connectgaps=False,
    ))

    fig.add_hline(
        y=ON_TIME_TARGET_PCT,
        line=dict(color='#6c757d', width=1, dash='dash'),
        annotation_text=f'Target {ON_TIME_TARGET_PCT:.0f}%',
        annotation_position='top right',
        annotation_font=dict(color='#6c757d', size=11),
        yref='y2',
    )

    if mode == 'month':
        month, year = period_info
        title_text = f'Picking KPI - {datetime(year, month, 1).strftime("%B %Y")}'
    else:
        title_text = 'Picking KPI - Past 2 Months'

    fig.update_layout(
        title=dict(text=title_text, font=dict(size=24, color='#2c3e50'),
                   x=0.5, xanchor='center'),
        xaxis_title='Plan Ship Date',
        yaxis_title='Number of Modules',
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
            rangebreaks=[
                dict(bounds=["sat", "mon"]),
                dict(values=_COMPANY_HOLIDAY_STRS),
            ],
        ),
        yaxis=dict(
            showgrid=True,
            gridcolor='#ecf0f1',
            linecolor='#bdc3c7',
            linewidth=2,
        ),
        yaxis2=dict(
            title='On-Time %',
            overlaying='y',
            side='right',
            range=[0, 105],
            showgrid=False,
            ticksuffix='%',
        ),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='center',
            x=0.5,
            font=dict(size=13),
        ),
        height=620,
        autosize=True,
        margin=dict(l=80, r=80, t=150, b=100),
    )

    return fig

def generate_summary_stats(daily_summary, period_data, mode, period_info, prev_month=None):
    """Generate summary statistics for the period.

    `period_data` is the raw filtered rows — used for a correct unique trailer
    count. `prev_month` (optional) enables the month-over-month delta shown on
    the On-Time stat card.
    """

    total_modules = int(daily_summary['total_modules'].sum())
    total_on_time = int(daily_summary['on_time_modules'].sum())
    total_late = int(daily_summary['late_modules'].sum())
    total_picked_late = int(daily_summary.get('picked_late_modules', pd.Series([0])).sum())
    total_not_picked = int(daily_summary.get('not_picked_modules', pd.Series([0])).sum())
    overall_pct = (total_on_time / total_modules * 100) if total_modules > 0 else 0.0

    if mode == 'month':
        month, year = period_info
        period_label = datetime(year, month, 1).strftime('%B %Y')
        month_name = datetime(year, month, 1).strftime('%B')
    else:
        start_date, end_date = period_info
        month, year = None, None
        period_label = f'{start_date.strftime("%b %d, %Y")} - {end_date.strftime("%b %d, %Y")}'
        month_name = None

    unique_trailers = int(period_data['TRAILER NO'].nunique()) if (
        period_data is not None and not period_data.empty and 'TRAILER NO' in period_data.columns
    ) else 0

    mom = None
    if prev_month is not None and prev_month['total_modules'] > 0:
        prev_pct = prev_month['on_time_modules'] / prev_month['total_modules'] * 100
        mom = {
            'prev_label': datetime(prev_month['year'], prev_month['month'], 1).strftime('%B %Y'),
            'prev_pct': round(prev_pct, 1),
            'delta_pts': round(overall_pct - prev_pct, 1),
        }

    # Average modules per active (non-empty) day.
    active_days = daily_summary[daily_summary['total_modules'] > 0] if len(daily_summary) else daily_summary
    avg_daily = round(active_days['total_modules'].mean(), 1) if len(active_days) > 0 else 0

    return {
        'period': period_label,
        'mode': mode,
        'month_name': month_name,
        'total_trailers': unique_trailers,
        'total_modules': total_modules,
        'on_time_modules': total_on_time,
        'late_modules': total_late,
        'picked_late_modules': total_picked_late,
        'not_picked_modules': total_not_picked,
        'overall_pct': round(overall_pct, 1),
        'avg_daily_modules': avg_daily,
        'mom': mom,
    }

def _format_trailer_no(val):
    """Render a trailer number defensively (handles strings, floats, NaN)."""
    if pd.isna(val):
        return 'N/A'
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(val)
    if isinstance(val, (int,)):
        return str(val)
    return str(val).strip() or 'N/A'


def create_html_report(
    fig_month, stats_month, daily_summary_month,
    fig_past2months, stats_past2months, daily_summary_past2months,
    df, load_info,
):
    """Create HTML report with visualization and statistics."""

    chart_month_html = fig_month.to_html(include_plotlyjs='cdn', div_id='chart-month', config={'displayModeBar': False, 'responsive': True})
    chart_past2months_html = fig_past2months.to_html(include_plotlyjs=False, div_id='chart-past2months', config={'displayModeBar': False, 'responsive': True})

    # Build container-level detail data grouped by PLAN SHIP DATE for dropdown rows.
    container_details = {}
    for date in df['PLAN SHIP DATE'].dropna().unique():
        date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
        date_data = df[df['PLAN SHIP DATE'] == date].copy()

        containers = []
        for container in date_data['TEMP.TRAILER'].dropna().unique():
            container_data = date_data[date_data['TEMP.TRAILER'] == container]

            trailer_no = _format_trailer_no(container_data['TRAILER NO'].iloc[0])

            picking_due_obj = container_data['PICKING DUE DATE'].iloc[0]
            picking_due_str = picking_due_obj.strftime('%b %d, %Y') if pd.notna(picking_due_obj) else 'N/A'

            picked_rows = container_data[container_data['PICKING DATE'].notna()]
            if len(picked_rows) > 0:
                pick_date_obj = picked_rows['PICKING DATE'].iloc[0]
                pick_time_str = str(picked_rows['PICKING TIME'].iloc[0]) if pd.notna(picked_rows['PICKING TIME'].iloc[0]) else ''
                if pick_time_str and '.' in pick_time_str:
                    parts = pick_time_str.split('.')
                    pick_time_formatted = f"{parts[0]}:{parts[1]}:{parts[2]}" if len(parts) >= 3 else f"{parts[0]}:{parts[1]}"
                else:
                    pick_time_formatted = pick_time_str
                picking_datetime = f"{pick_date_obj.strftime('%b %d, %Y')} {pick_time_formatted}".strip()
                sort_key = pick_date_obj.strftime('%Y-%m-%d') + ' ' + pick_time_formatted
            else:
                picking_datetime = 'Not Picked'
                sort_key = '9999-99-99'

            total_rows = len(container_data)
            on_time_rows = int(container_data['ON_TIME'].sum())
            has_unpicked_overdue = bool((~container_data['IS_PICKED']).any())

            # "Not picked" is the highest-severity status: if ANY row in the container
            # is overdue-and-unpicked, surface that so those modules aren't hidden
            # behind a "mixed" label when some siblings happened to be picked on time.
            if has_unpicked_overdue:
                status = 'not_picked'
            elif on_time_rows == total_rows:
                status = 'on_time'
            elif on_time_rows == 0:
                status = 'late'
            else:
                status = 'mixed'

            containers.append({
                'container': str(container),
                'trailer_no': trailer_no,
                'picking_due': picking_due_str,
                'picking_datetime': picking_datetime,
                'modules': total_rows,
                'status': status,
                'sort_key': sort_key,
            })

        containers.sort(key=lambda x: x['sort_key'])
        for c in containers:
            del c['sort_key']

        container_details[date_str] = containers
    container_details_json = json.dumps(container_details)

    # Per-date trailer IDs (strings) so custom-range JS can compute TRUE unique
    # trailers by union-ing Sets — replacing the broken sum-of-per-day-counts.
    trailers_by_date = {}
    for date, group in df.dropna(subset=['PLAN SHIP DATE']).groupby('PLAN SHIP DATE'):
        date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
        trailers_by_date[date_str] = [
            _format_trailer_no(v) for v in group['TRAILER NO'].dropna().unique()
        ]
    trailers_by_date_json = json.dumps(trailers_by_date)

    daily_summary_all = build_daily_summary(df)
    if not daily_summary_all.empty:
        daily_summary_all = fill_missing_dates(
            daily_summary_all, 'PLAN SHIP DATE',
            daily_summary_all['PLAN SHIP DATE'].min(),
            daily_summary_all['PLAN SHIP DATE'].max(),
        )

    js_data = []
    for _, row in daily_summary_all.iterrows():
        js_data.append({
            'date': row['PLAN SHIP DATE'].strftime('%Y-%m-%d'),
            'total_trailers': int(row['total_trailers']),
            'total_modules': int(row['total_modules']),
            'on_time_modules': int(row['on_time_modules']),
            'late_modules': int(row['late_modules']),
            'picked_late_modules': int(row.get('picked_late_modules', 0)),
            'not_picked_modules': int(row.get('not_picked_modules', 0)),
            'on_time_pct': float(row['on_time_pct']),
        })

    all_data_json = json.dumps(js_data)
    company_holidays_json = json.dumps(_COMPANY_HOLIDAY_STRS)

    def pct_class(pct):
        if pct is None:
            return 'pct-na'
        if pct >= 95:
            return 'pct-great'
        if pct >= 85:
            return 'pct-ok'
        return 'pct-bad'

    def plural(n, word, word_plural=None):
        return f"{n:,} {word}" if n == 1 else f"{n:,} {word_plural or (word + 's')}"

    def render_mom_line(stats):
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

    def render_daily_rows(daily_summary, view_prefix):
        parts = []
        for _, row in daily_summary.iterrows():
            date_str = row['PLAN SHIP DATE'].strftime('%b %d')
            day_name = row['PLAN SHIP DATE'].strftime('%A')
            date_key = row['PLAN SHIP DATE'].strftime('%Y-%m-%d')
            total_modules = int(row['total_modules'])
            on_time_d = int(row['on_time_modules'])
            picked_late_d = int(row.get('picked_late_modules', 0))
            not_picked_d = int(row.get('not_picked_modules', 0))
            pct = float(row['on_time_pct'])
            pct_cls = pct_class(pct) if total_modules > 0 else 'pct-na'
            pct_display = f'{pct:.1f}%' if total_modules > 0 else '&mdash;'
            parts.append(f"""
                    <tr class="data-row" data-date="{date_key}" tabindex="0" role="button" aria-expanded="false">
                        <td class="expand-cell"><span class="expand-arrow">&#9662;</span></td>
                        <td><strong>{date_str}</strong></td>
                        <td>{day_name}</td>
                        <td>{int(row['total_trailers'])}</td>
                        <td>{total_modules:,}</td>
                        <td style="color: #28a745; font-weight: bold;">{on_time_d:,}</td>
                        <td style="color: #dc3545; font-weight: bold;">{picked_late_d:,}</td>
                        <td style="color: #b4590a; font-weight: bold;">{not_picked_d:,}</td>
                        <td class="pct-cell {pct_cls}"><strong>{pct_display}</strong></td>
                    </tr>
                    <tr class="detail-row" id="{view_prefix}-detail-{date_key}">
                        <td colspan="9">
                            <div class="detail-content">
                                <div class="module-details">Loading module details...</div>
                            </div>
                        </td>
                    </tr>""")
        return ''.join(parts)

    def render_tfoot(stats):
        pct = stats['overall_pct']
        cls = pct_class(pct) if stats['total_modules'] > 0 else 'pct-na'
        return f"""
                    <tfoot>
                        <tr>
                            <td></td>
                            <td><strong>TOTAL</strong></td>
                            <td></td>
                            <td>{stats['total_trailers']}</td>
                            <td>{stats['total_modules']:,}</td>
                            <td style="color: #28a745; font-weight: bold;">{stats['on_time_modules']:,}</td>
                            <td style="color: #dc3545; font-weight: bold;">{stats['picked_late_modules']:,}</td>
                            <td style="color: #b4590a; font-weight: bold;">{stats['not_picked_modules']:,}</td>
                            <td class="pct-cell {cls}"><strong>{pct:.1f}%</strong></td>
                        </tr>
                    </tfoot>"""

    month_rows_html = render_daily_rows(daily_summary_month, 'month')
    past2m_rows_html = render_daily_rows(daily_summary_past2months, 'past2months')
    month_tfoot_html = render_tfoot(stats_month)
    past2m_tfoot_html = render_tfoot(stats_past2months)
    month_mom_line = render_mom_line(stats_month)
    past2m_mom_line = render_mom_line(stats_past2months)
    month_trailer_label = plural(stats_month['total_trailers'], 'unique trailer')
    past2m_trailer_label = plural(stats_past2months['total_trailers'], 'unique trailer')

    date_range_str = ''
    if load_info.get('min_ship_date') is not None and load_info.get('max_ship_date') is not None:
        date_range_str = (f"{load_info['min_ship_date'].strftime('%b %d, %Y')} – "
                          f"{load_info['max_ship_date'].strftime('%b %d, %Y')}")

    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Picking KPI Report - {stats_month['period']}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}

        body {{
            font-family: Arial, Helvetica, sans-serif;
            background: white;
            padding: 0;
            color: #2c3e50;
        }}

        .container {{ max-width: 100%; margin: 0; background: white; overflow: hidden; }}

        .header {{
            background: linear-gradient(135deg, #f1f3f5 0%, #dee2e6 100%);
            color: #2f3b45;
            padding: 18px 30px;
            text-align: center;
            border-bottom: 1px solid #cfd4da;
        }}
        .header h1 {{ font-size: 1.6em; margin-bottom: 4px; font-weight: 600; letter-spacing: 0.6px; }}
        .header p  {{ font-size: 0.9em; color: #4f5b66; }}

        .stats-container {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 20px;
            padding: 24px 30px;
            background: #f8f9fa;
        }}
        @media (max-width: 1100px) {{
            .stats-container {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
        }}
        @media (max-width: 620px) {{
            .stats-container {{ grid-template-columns: 1fr; }}
        }}
        .stat-card {{
            background: white;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            border-top: 3px solid transparent;
        }}
        .stat-card:hover {{ transform: translateY(-5px); box-shadow: 0 8px 15px rgba(0,0,0,0.2); }}
        .stat-card h3 {{
            color: #7f8c8d;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }}
        .stat-card .value    {{ font-size: 2.5em; font-weight: bold; color: #2c3e50; line-height: 1.1; }}
        .stat-card.success .value {{ color: #28a745; }}
        .stat-card.danger  .value {{ color: #dc3545; }}
        .stat-card .subtext  {{ color: #95a5a6; font-size: 0.9em; margin-top: 5px; }}

        .mom-line {{
            margin-top: 6px;
            font-size: 0.85em;
            color: #4f5b66;
        }}
        .mom-up   {{ color: #1f7a3a; font-weight: 600; }}
        .mom-down {{ color: #b3261e; font-weight: 600; }}
        .mom-flat {{ color: #6c757d; font-weight: 600; }}

        .chart-container {{ padding: 30px; background: white; }}
        .chart-container > div {{ width: 100%; }}

        .details-container {{ padding: 30px; background: #f8f9fa; }}
        .details-container h2 {{
            color: #2c3e50;
            font-size: 1.1em;
            margin-bottom: 14px;
            padding-bottom: 8px;
            border-bottom: 2px solid #3498db;
            letter-spacing: 0.4px;
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

        .view-content {{ display: none; }}
        .view-content.active {{ display: block; }}

        table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            box-shadow: 0 4px 6px rgba(0,0,0,0.08);
        }}
        thead th {{
            background: #95a5a6;
            color: white;
            padding: 13px 12px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.82em;
            letter-spacing: 0.5px;
            position: sticky;
            top: 0;
            z-index: 2;
        }}
        td {{ padding: 11px 12px; border-bottom: 1px solid #ecf0f1; }}
        tbody tr:hover {{ background: #f8f9fa; }}
        tfoot {{ background: #e9ecef; font-weight: bold; border-top: 3px solid #95a5a6; }}
        tfoot td {{ padding: 14px 12px; font-size: 1.02em; border-bottom: none; }}
        tr:last-child td {{ border-bottom: none; }}

        th.arrow-col, td.expand-cell {{ width: 28px; text-align: center; color: #7f8c8d; }}
        .data-row {{ cursor: pointer; }}
        .data-row:focus {{ outline: 2px solid #3498db; outline-offset: -2px; background: #eaf4fc; }}
        .expand-arrow {{ display: inline-block; font-size: 0.9em; transition: transform 0.2s ease, color 0.2s ease; }}
        .data-row.expanded .expand-arrow {{ transform: rotate(180deg); color: #2c3e50; }}

        .detail-row td {{ padding: 0; border-bottom: none; }}
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

        .pct-cell {{ text-align: left; }}
        .pct-cell.pct-great {{ background: #d4edda; color: #155724; }}
        .pct-cell.pct-ok    {{ background: #fff3cd; color: #856404; }}
        .pct-cell.pct-bad   {{ background: #f8d7da; color: #721c24; }}
        .pct-cell.pct-na    {{ color: #adb5bd; }}

        .footer {{
            background: #e9ecef;
            color: #4f5b66;
            text-align: center;
            padding: 20px;
            font-size: 0.9em;
        }}
        .footer .data-quality {{
            margin-top: 6px;
            font-size: 0.85em;
            color: #6c757d;
        }}

        .helper-hint {{
            color: #c0392b;
            font-size: 0.95em;
            margin-bottom: 15px;
        }}

        @media print {{
            body {{ background: white; padding: 0; }}
            .container {{ box-shadow: none; }}
            .toggle-container, .date-picker-container {{ display: none; }}
            thead th {{ position: static; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Example Logistics - Monthly Picking KPI</h1>
            <p>This dashboard tracks how well we're meeting our picking standard (modules picked 2 business days before ship date)</p>
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
            <div class="stats-container">
                <div class="stat-card">
                    <h3>Period</h3>
                    <div class="value" style="font-size: 1.5em;">{stats_month['period']}</div>
                </div>
                <div class="stat-card">
                    <h3>Total Modules</h3>
                    <div class="value">{stats_month['total_modules']:,}</div>
                    <div class="subtext">{month_trailer_label} &middot; Avg. {stats_month['avg_daily_modules']:.0f} / active day</div>
                </div>
                <div class="stat-card success">
                    <h3>Picked On Time</h3>
                    <div class="value">{stats_month['overall_pct']:.1f}%</div>
                    <div class="subtext">{stats_month['on_time_modules']:,} modules &middot; target {ON_TIME_TARGET_PCT:.0f}%</div>
                    {month_mom_line}
                </div>
                <div class="stat-card danger">
                    <h3>Picked Late / Not Picked</h3>
                    <div class="value">{100 - stats_month['overall_pct']:.1f}%</div>
                    <div class="subtext">{stats_month['picked_late_modules']:,} late &middot; {stats_month['not_picked_modules']:,} not picked</div>
                </div>
            </div>

            <div class="chart-container">{chart_month_html}</div>

            <div class="details-container">
                <h2>Daily Picking Details</h2>
                <p class="helper-hint">Click a row (or press Enter) to see container-level details for that date.</p>
                <table>
                    <thead>
                        <tr>
                            <th class="arrow-col"></th>
                            <th>Date</th>
                            <th>Day</th>
                            <th>Trailers</th>
                            <th>Modules</th>
                            <th>On Time</th>
                            <th>Late</th>
                            <th>Not Picked</th>
                            <th>On Time %</th>
                        </tr>
                    </thead>
                    <tbody>{month_rows_html}
                    </tbody>{month_tfoot_html}
                </table>
            </div>
        </div>

        <div id="view-past2months" class="view-content">
            <div class="stats-container">
                <div class="stat-card">
                    <h3>Period</h3>
                    <div class="value" style="font-size: 1.5em;">{stats_past2months['period']}</div>
                </div>
                <div class="stat-card">
                    <h3>Total Modules</h3>
                    <div class="value">{stats_past2months['total_modules']:,}</div>
                    <div class="subtext">{past2m_trailer_label} &middot; Avg. {stats_past2months['avg_daily_modules']:.0f} / active day</div>
                </div>
                <div class="stat-card success">
                    <h3>Picked On Time</h3>
                    <div class="value">{stats_past2months['overall_pct']:.1f}%</div>
                    <div class="subtext">{stats_past2months['on_time_modules']:,} modules &middot; target {ON_TIME_TARGET_PCT:.0f}%</div>
                    {past2m_mom_line}
                </div>
                <div class="stat-card danger">
                    <h3>Picked Late / Not Picked</h3>
                    <div class="value">{100 - stats_past2months['overall_pct']:.1f}%</div>
                    <div class="subtext">{stats_past2months['picked_late_modules']:,} late &middot; {stats_past2months['not_picked_modules']:,} not picked</div>
                </div>
            </div>

            <div class="chart-container">{chart_past2months_html}</div>

            <div class="details-container">
                <h2>Daily Picking Details</h2>
                <p class="helper-hint">Click a row (or press Enter) to see container-level details for that date.</p>
                <table>
                    <thead>
                        <tr>
                            <th class="arrow-col"></th>
                            <th>Date</th>
                            <th>Day</th>
                            <th>Trailers</th>
                            <th>Modules</th>
                            <th>On Time</th>
                            <th>Late</th>
                            <th>Not Picked</th>
                            <th>On Time %</th>
                        </tr>
                    </thead>
                    <tbody>{past2m_rows_html}
                    </tbody>{past2m_tfoot_html}
                </table>
            </div>
        </div>

        <div class="footer">
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
            <p class="data-quality">Source: Data/202.csv &middot; {load_info['usable_rows']:,} rows analysed ({load_info['invalid_ship_rows']:,} excluded for invalid ship date) &middot; Covers {date_range_str}</p>
            <p style="margin-top: 8px; font-size: 0.85em; color: #7f8c8d;">Dashboard created by Viktor Berg &middot; Built with Python, Plotly and Pandas</p>
        </div>

        <script>
            const allData = {all_data_json};
            const containerDetails = {container_details_json};
            const trailersByDate = {trailers_by_date_json};
            const companyHolidays = {company_holidays_json};
            const targetPct = {ON_TIME_TARGET_PCT};

            function pctClass(pct) {{
                if (pct === null || pct === undefined || isNaN(pct)) return 'pct-na';
                if (pct >= 95) return 'pct-great';
                if (pct >= 85) return 'pct-ok';
                return 'pct-bad';
            }}

            function expandDetailRow(row) {{
                const date = row.getAttribute('data-date');
                const viewContainer = row.closest('.view-content');
                let viewPrefix = '';
                if (viewContainer) {{
                    if (viewContainer.id === 'view-month') viewPrefix = 'month-';
                    else if (viewContainer.id === 'view-past2months') viewPrefix = 'past2months-';
                    else if (viewContainer.id === 'view-custom') viewPrefix = 'custom-';
                }}
                const detailRow = document.getElementById(viewPrefix + 'detail-' + date);
                if (!detailRow) return;

                const isExpanded = detailRow.classList.contains('expanded');
                if (isExpanded) {{
                    detailRow.classList.remove('expanded');
                    row.classList.remove('expanded');
                    row.setAttribute('aria-expanded', 'false');
                    return;
                }}

                const details = containerDetails[date] || [];
                let detailHTML = '<table style="width: 100%; margin: 10px 0;"><thead><tr style="background: #34495e; color: white;"><th>Trailer Number</th><th>Container</th><th>Plan Pick Date</th><th>Actual Pick Date &amp; Time</th><th>Modules</th></tr></thead><tbody>';
                if (details.length === 0) {{
                    detailHTML += '<tr><td colspan="5" style="text-align: center; padding: 20px; color: #7f8c8d;">No container data available for this date</td></tr>';
                }} else {{
                    details.forEach(c => {{
                        let rowColor = '';
                        if (c.status === 'on_time')        rowColor = 'background-color: #d4edda;';
                        else if (c.status === 'late')      rowColor = 'background-color: #f8d7da;';
                        else if (c.status === 'not_picked') rowColor = 'background-color: #ffe5cc;';
                        else                                rowColor = 'background-color: #fff3cd;';
                        detailHTML += `<tr style="${{rowColor}}"><td>${{c.trailer_no}}</td><td>${{c.container}}</td><td>${{c.picking_due}}</td><td>${{c.picking_datetime}}</td><td>${{c.modules}}</td></tr>`;
                    }});
                }}
                detailHTML += '</tbody></table>';
                detailRow.querySelector('.module-details').innerHTML = detailHTML;
                detailRow.classList.add('expanded');
                row.classList.add('expanded');
                row.setAttribute('aria-expanded', 'true');
            }}

            document.addEventListener('DOMContentLoaded', function() {{
                document.body.addEventListener('click', function(e) {{
                    const row = e.target.closest('.data-row');
                    if (!row) return;
                    expandDetailRow(row);
                }});
                // Keyboard support: Enter / Space toggles the detail row on focus.
                document.body.addEventListener('keydown', function(e) {{
                    if (e.key !== 'Enter' && e.key !== ' ') return;
                    const row = e.target.closest('.data-row');
                    if (!row) return;
                    e.preventDefault();
                    expandDetailRow(row);
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
                // Anchor quick-picks to the last day that has modules, not "today" —
                // the CSV may trail the calendar (e.g. report generated weeks later).
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

            function applyCustomRange() {{
                const startDate = document.getElementById('start-date').value;
                const endDate   = document.getElementById('end-date').value;
                if (!startDate || !endDate) {{ alert('Please select both start and end dates.'); return; }}
                if (parseDate(startDate) > parseDate(endDate)) {{ alert('Start date must be before end date.'); return; }}

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

            function rollingAverage(filteredData, window) {{
                const out = [];
                const active = [];
                filteredData.forEach(item => {{
                    if (item.total_modules > 0) {{
                        active.push({{on: item.on_time_modules, tot: item.total_modules}});
                        if (active.length > window) active.shift();
                        const tot = active.reduce((s, x) => s + x.tot, 0);
                        const on  = active.reduce((s, x) => s + x.on, 0);
                        out.push(tot > 0 ? Math.round(on / tot * 1000) / 10 : null);
                    }} else {{
                        out.push(null);
                    }}
                }});
                return out;
            }}

            function filterDataByRange(startDate, endDate) {{
                const start = parseDate(startDate);
                const end   = parseDate(endDate);
                const filteredData = allData.filter(item => {{
                    const d = parseDate(item.date);
                    return d >= start && d <= end;
                }}).sort((a, b) => parseDate(a.date) - parseDate(b.date));

                if (filteredData.length === 0) {{ alert('No data available for the selected date range.'); return; }}

                const totalModules   = filteredData.reduce((s, i) => s + i.total_modules, 0);
                const onTimeModules  = filteredData.reduce((s, i) => s + i.on_time_modules, 0);
                const pickedLateMods = filteredData.reduce((s, i) => s + i.picked_late_modules, 0);
                const notPickedMods  = filteredData.reduce((s, i) => s + i.not_picked_modules, 0);
                const lateModules    = pickedLateMods + notPickedMods;
                const overallPct     = totalModules > 0 ? (onTimeModules / totalModules * 100) : 0;
                const activeDays     = filteredData.filter(i => i.total_modules > 0).length;
                const avgDaily       = activeDays > 0 ? totalModules / activeDays : 0;

                // True unique trailers: union sets from per-day trailer lists.
                const trailerSet = new Set();
                filteredData.forEach(i => {{
                    (trailersByDate[i.date] || []).forEach(t => trailerSet.add(t));
                }});
                const uniqueTrailers = trailerSet.size;

                const formatDate = (s) => parseDate(s).toLocaleDateString('en-US', {{month: 'short', day: 'numeric', year: 'numeric'}});
                const formatDay  = (s) => parseDate(s).toLocaleDateString('en-US', {{weekday: 'long'}});
                const periodLabel = `${{formatDate(startDate)}} - ${{formatDate(endDate)}}`;

                const dates           = filteredData.map(i => parseDate(i.date));
                const onTimeData      = filteredData.map(i => i.on_time_modules);
                const pickedLateData  = filteredData.map(i => i.picked_late_modules);
                const notPickedData   = filteredData.map(i => i.not_picked_modules);
                const percentages     = filteredData.map(i => i.on_time_pct);
                const rollingVals     = rollingAverage(filteredData, 7);
                const maxTotal        = Math.max(...filteredData.map(i => i.total_modules));
                const labelThreshold  = Math.max(1, Math.floor(maxTotal * 0.08));

                const pctOf = (n, t) => t > 0 ? (n / t * 100) : 0;
                const hoverOnTime    = filteredData.map(i => i.total_modules > 0 ? `${{i.on_time_modules.toLocaleString()}} modules (${{pctOf(i.on_time_modules, i.total_modules).toFixed(1)}}%)` : '0');
                const hoverPickedLate = filteredData.map(i => i.total_modules > 0 ? `${{i.picked_late_modules.toLocaleString()}} modules (${{pctOf(i.picked_late_modules, i.total_modules).toFixed(1)}}%)` : '0');
                const hoverNotPicked  = filteredData.map(i => i.total_modules > 0 ? `${{i.not_picked_modules.toLocaleString()}} modules (${{pctOf(i.not_picked_modules, i.total_modules).toFixed(1)}}%)` : '0');

                const traceOnTime = {{
                    x: dates, y: onTimeData, type: 'bar', name: 'Picked On Time',
                    marker: {{color: '#28a745'}},
                    text: onTimeData.map((v, idx) => v > 0 ? `${{percentages[idx].toFixed(1)}}%` : ''),
                    textposition: onTimeData.map(v => v > 0 && v < labelThreshold ? 'outside' : 'inside'),
                    textfont: {{color: 'white', size: 10, family: 'Arial'}},
                    cliponaxis: false,
                    hovertemplate: '<b>On Time</b>: %{{customdata}}<extra></extra>',
                    customdata: hoverOnTime
                }};
                const tracePickedLate = {{
                    x: dates, y: pickedLateData, type: 'bar', name: 'Picked Late',
                    marker: {{color: '#dc3545'}},
                    hovertemplate: '<b>Picked Late</b>: %{{customdata}}<extra></extra>',
                    customdata: hoverPickedLate
                }};
                const traceNotPicked = {{
                    x: dates, y: notPickedData, type: 'bar', name: 'Not Picked (overdue)',
                    marker: {{color: '#fd7e14'}},
                    hovertemplate: '<b>Not Picked</b>: %{{customdata}}<extra></extra>',
                    customdata: hoverNotPicked
                }};
                const traceRolling = {{
                    x: dates, y: rollingVals, type: 'scatter', mode: 'lines+markers',
                    name: 'Rolling 7-day On-Time %',
                    line: {{color: '#2c3e50', width: 2, dash: 'dot'}},
                    marker: {{size: 5}},
                    yaxis: 'y2',
                    connectgaps: false,
                    hovertemplate: '<b>7-day rolling avg</b>: %{{y:.1f}}%<extra></extra>'
                }};

                const layout = {{
                    title: {{text: `Picking KPI - ${{periodLabel}}`, font: {{size: 22, color: '#2c3e50'}}, x: 0.5, xanchor: 'center'}},
                    xaxis: {{
                        title: {{text: 'Plan Ship Date'}}, tickformat: '%b %d', dtick: 86400000, tickangle: -45,
                        rangebreaks: [{{bounds: ['sat', 'mon']}}, {{values: companyHolidays}}]
                    }},
                    yaxis:  {{title: {{text: 'Number of Modules'}}}},
                    yaxis2: {{title: 'On-Time %', overlaying: 'y', side: 'right', range: [0, 105], ticksuffix: '%', showgrid: false}},
                    barmode: 'stack', hovermode: 'x unified',
                    plot_bgcolor: 'white', paper_bgcolor: 'white',
                    height: 620, autosize: true,
                    margin: {{l: 80, r: 80, t: 150, b: 100}},
                    shapes: [{{type: 'line', xref: 'paper', x0: 0, x1: 1, yref: 'y2', y0: targetPct, y1: targetPct, line: {{color: '#6c757d', width: 1, dash: 'dash'}}}}],
                    annotations: [{{xref: 'paper', yref: 'y2', x: 1, y: targetPct, xanchor: 'right', yanchor: 'bottom', text: `Target ${{targetPct}}%`, showarrow: false, font: {{color: '#6c757d', size: 11}}}}]
                }};

                const customView = document.getElementById('view-custom');
                const periodTotalPct = overallPct;
                const periodPctCls = totalModules > 0 ? pctClass(periodTotalPct) : 'pct-na';
                const trailerLabel = uniqueTrailers === 1 ? '1 unique trailer' : `${{uniqueTrailers.toLocaleString()}} unique trailers`;
                customView.innerHTML = `
                    <div class="stats-container">
                        <div class="stat-card">
                            <h3>Period</h3>
                            <div class="value" style="font-size: 1.5em;">${{periodLabel}}</div>
                        </div>
                        <div class="stat-card">
                            <h3>Total Modules</h3>
                            <div class="value">${{totalModules.toLocaleString()}}</div>
                            <div class="subtext">${{trailerLabel}} &middot; Avg. ${{avgDaily.toFixed(0)}} / active day</div>
                        </div>
                        <div class="stat-card success">
                            <h3>Picked On Time</h3>
                            <div class="value">${{overallPct.toFixed(1)}}%</div>
                            <div class="subtext">${{onTimeModules.toLocaleString()}} modules &middot; target ${{targetPct}}%</div>
                        </div>
                        <div class="stat-card danger">
                            <h3>Picked Late / Not Picked</h3>
                            <div class="value">${{(100 - overallPct).toFixed(1)}}%</div>
                            <div class="subtext">${{pickedLateMods.toLocaleString()}} late &middot; ${{notPickedMods.toLocaleString()}} not picked</div>
                        </div>
                    </div>
                    <div class="chart-container"><div id="custom-chart"></div></div>
                    <div class="details-container">
                        <h2>Daily Picking Details</h2>
                        <p class="helper-hint">Click a row (or press Enter) to see container-level details for that date.</p>
                        <table>
                            <thead>
                                <tr>
                                    <th class="arrow-col"></th>
                                    <th>Date</th>
                                    <th>Day</th>
                                    <th>Trailers</th>
                                    <th>Modules</th>
                                    <th>On Time</th>
                                    <th>Late</th>
                                    <th>Not Picked</th>
                                    <th>On Time %</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${{filteredData.map(item => {{
                                    const cls = item.total_modules > 0 ? pctClass(item.on_time_pct) : 'pct-na';
                                    const pctDisp = item.total_modules > 0 ? `${{item.on_time_pct.toFixed(1)}}%` : '&mdash;';
                                    return `
                                        <tr class="data-row" data-date="${{item.date}}" tabindex="0" role="button" aria-expanded="false">
                                            <td class="expand-cell"><span class="expand-arrow">&#9662;</span></td>
                                            <td><strong>${{formatDate(item.date)}}</strong></td>
                                            <td>${{formatDay(item.date)}}</td>
                                            <td>${{(trailersByDate[item.date] || []).length}}</td>
                                            <td>${{item.total_modules.toLocaleString()}}</td>
                                            <td style="color: #28a745; font-weight: bold;">${{item.on_time_modules.toLocaleString()}}</td>
                                            <td style="color: #dc3545; font-weight: bold;">${{item.picked_late_modules.toLocaleString()}}</td>
                                            <td style="color: #b4590a; font-weight: bold;">${{item.not_picked_modules.toLocaleString()}}</td>
                                            <td class="pct-cell ${{cls}}"><strong>${{pctDisp}}</strong></td>
                                        </tr>
                                        <tr class="detail-row" id="custom-detail-${{item.date}}">
                                            <td colspan="9"><div class="detail-content"><div class="module-details">Loading module details...</div></div></td>
                                        </tr>`;
                                }}).join('')}}
                            </tbody>
                            <tfoot>
                                <tr>
                                    <td></td>
                                    <td><strong>TOTAL</strong></td>
                                    <td></td>
                                    <td>${{uniqueTrailers}}</td>
                                    <td>${{totalModules.toLocaleString()}}</td>
                                    <td style="color: #28a745; font-weight: bold;">${{onTimeModules.toLocaleString()}}</td>
                                    <td style="color: #dc3545; font-weight: bold;">${{pickedLateMods.toLocaleString()}}</td>
                                    <td style="color: #b4590a; font-weight: bold;">${{notPickedMods.toLocaleString()}}</td>
                                    <td class="pct-cell ${{periodPctCls}}"><strong>${{periodTotalPct.toFixed(1)}}%</strong></td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                `;

                Plotly.newPlot('custom-chart', [traceOnTime, tracePickedLate, traceNotPicked, traceRolling], layout, {{
                    displayModeBar: false, responsive: true, staticPlot: false, editable: false, scrollZoom: false
                }});
            }}
        </script>
    </body>
    </html>
    """

    return html_content

def main():
    """Main function to generate the report."""
    
    parser = argparse.ArgumentParser(
        description='Generate Picking KPI Report with interactive month/past 2 months toggle',
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
    daily_summary_month, period_data_month, mode_month, month, year = (
        create_daily_summary_calendar_month(df, target_month=args.month, target_year=args.year)
    )
    period_info_month = (month, year)
    prev_month = create_previous_month_summary(df, month, year)
    fig_month = create_visualization(daily_summary_month, mode_month, period_info_month)
    stats_month = generate_summary_stats(
        daily_summary_month, period_data_month, mode_month, period_info_month, prev_month=prev_month
    )

    print(f"   Period: {stats_month['period']}")
    print(f"   Days with data: {len(daily_summary_month)}")
    print(f"   Total Modules: {stats_month['total_modules']:,}")
    print(f"   Unique Trailers: {stats_month['total_trailers']:,}")
    print(f"   On Time: {stats_month['on_time_modules']:,} ({stats_month['overall_pct']:.1f}%)")
    if stats_month['mom']:
        print(f"   vs {stats_month['mom']['prev_label']}: {stats_month['mom']['delta_pts']:+.1f} pts")

    print("\nGenerating past 2 months report...")
    daily_summary_past2months, period_data_past2m, mode_past2months, start_date, end_date = (
        create_daily_summary_past_2_months(df)
    )
    period_info_past2months = (start_date, end_date)
    fig_past2months = create_visualization(daily_summary_past2months, mode_past2months, period_info_past2months)
    stats_past2months = generate_summary_stats(
        daily_summary_past2months, period_data_past2m, mode_past2months, period_info_past2months
    )

    print(f"   Period: {stats_past2months['period']}")
    print(f"   Days with data: {len(daily_summary_past2months)}")
    print(f"   Total Modules: {stats_past2months['total_modules']:,}")
    print(f"   Unique Trailers: {stats_past2months['total_trailers']:,}")
    print(f"   On Time: {stats_past2months['on_time_modules']:,} ({stats_past2months['overall_pct']:.1f}%)")

    html_content = create_html_report(
        fig_month, stats_month, daily_summary_month,
        fig_past2months, stats_past2months, daily_summary_past2months,
        df, load_info,
    )

    html_output_file = f"Picking_KPI_Report_{year}_{month:02d}.html"
    with open(html_output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"\n[OK] HTML report generated: {html_output_file}")
    print(f"   Open the file in your web browser to view the report.")
    print(f"   Use the toggle buttons to switch between Calendar Month and Past 2 Months views.")

if __name__ == "__main__":
    main()
