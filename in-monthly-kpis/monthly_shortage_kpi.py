import argparse
from datetime import datetime, timedelta
import json
import os

import numpy as np
import pandas as pd
import plotly.graph_objects as go


REQUIRED_COLUMNS = ['CUSTOMER ORDER NO.', 'PRODUCT NO.', 'SHIP DATE', 'QUANTITY']

# No agreed business target for shortage yet. Leave as a hook: setting this
# to a number (e.g. 95.0) enables the horizontal target line in the chart.
TARGET_PCT = None


def _safe_json(obj):
    """json.dumps escaped so the output can't close a surrounding <script> block.

    Replaces ``</`` with ``<\\/``; both are equivalent inside a JSON string but
    the escaped form prevents ``</script>``-style sequences in user data from
    breaking the embedding HTML page.
    """
    return json.dumps(obj).replace('</', '<\\/')


def parse_planned_date(order_no):
    if pd.isna(order_no):
        return pd.NaT

    order_no = str(order_no)
    if '_' in order_no:
        date_str = order_no.split('_')[-1]
    else:
        date_str = order_no[8:16] if len(order_no) >= 16 else ''

    date_digits = ''.join(ch for ch in date_str if ch.isdigit())
    if len(date_digits) < 8:
        return pd.NaT

    return pd.to_datetime(date_digits[:8], format='%Y%m%d', errors='coerce')


def classify_shortage(plan_date, ship_date):
    if pd.isna(ship_date):
        return 'not_shipped'
    if ship_date == plan_date:
        return 'normal'
    if ship_date > plan_date:
        return 'delay'
    return 'early'


def build_shortage_note(plan_date, ship_date, shortage_type):
    if shortage_type == 'not_shipped':
        return 'Not shipped'
    if pd.isna(ship_date) or pd.isna(plan_date):
        return ''

    day_diff = (ship_date - plan_date).days
    if shortage_type == 'delay':
        return f"Delay {day_diff} day(s)"
    if shortage_type == 'early':
        return f"Earlier {abs(day_diff)} day(s)"
    return 'On time'


def load_and_process_data(data_folder):
    """Load 201S.csv and process shortage KPI data.

    Returns (df, load_info) where load_info summarizes row counts at each
    drop/aggregation step (surfaced in the HTML footer).
    """
    file_path = os.path.join(data_folder, '201S.csv')

    if not os.path.exists(file_path):
        raise FileNotFoundError(
            f"Could not find '{file_path}'. Place the 201S.csv export in the "
            f"'{data_folder}' folder (or pass --data-folder)."
        )

    try:
        df = pd.read_csv(file_path, encoding='utf-8-sig')
    except Exception as exc:
        raise RuntimeError(f"Failed to read '{file_path}': {exc}") from exc
    df.columns = df.columns.str.replace('\ufeff', '', regex=False).str.strip()

    if 'CUSTOMER ORDER NO.' not in df.columns and 'ORDER NO' in df.columns:
        df['CUSTOMER ORDER NO.'] = df['ORDER NO']
    if 'PRODUCT NO.' not in df.columns and 'PARTS NO' in df.columns:
        df['PRODUCT NO.'] = df['PARTS NO']
    if 'QUANTITY' not in df.columns and 'QTY' in df.columns:
        df['QUANTITY'] = df['QTY']
    if 'SHIP DATE' not in df.columns and 'SHIPMENT LOAD DATE' in df.columns:
        df['SHIP DATE'] = df['SHIPMENT LOAD DATE']

    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        raise ValueError(
            f"'{file_path}' is missing required columns: {', '.join(missing_cols)}. "
            f"Found columns: {', '.join(df.columns)}"
        )

    raw_rows = len(df)

    if 'PLAN SHIP DATE' in df.columns and df['PLAN SHIP DATE'].notna().any():
        df['PLAN SHIP DATE'] = pd.to_datetime(df['PLAN SHIP DATE'], errors='coerce')
    else:
        df['PLAN SHIP DATE'] = df['CUSTOMER ORDER NO.'].apply(parse_planned_date)
    df['SHIP DATE'] = pd.to_datetime(df['SHIP DATE'], format='mixed', dayfirst=False, errors='coerce')
    df['QUANTITY'] = pd.to_numeric(df['QUANTITY'], errors='coerce').fillna(0)

    invalid_plan_rows = int(df['PLAN SHIP DATE'].isna().sum())
    df = df.dropna(subset=['PLAN SHIP DATE']).copy()

    df['CUSTOMER ORDER NO.'] = df['CUSTOMER ORDER NO.'].astype(str).str.strip()
    df['PRODUCT NO.'] = df['PRODUCT NO.'].astype(str).str.strip()

    cancelled_mask = pd.Series(False, index=df.index)
    if 'STATUS' in df.columns:
        cancelled_mask |= df['STATUS'].astype(str).str.contains('cancel', case=False, na=False)
    if 'UC/CNL' in df.columns:
        cancelled_mask |= df['UC/CNL'].astype(str).str.contains('cnl', case=False, na=False)
    cancelled_rows = int(cancelled_mask.sum())
    df = df[~cancelled_mask].copy()

    # Collapse duplicates by aggregating rather than dropping: partial shipments,
    # re-plans, and re-entered rows should contribute their quantity instead of
    # being silently lost. The order is "complete" when its latest shipment loads,
    # so SHIP DATE takes max; QUANTITY is summed; other attributes take the first
    # value from the group.
    key_cols = ['CUSTOMER ORDER NO.', 'PRODUCT NO.']
    other_cols = [c for c in df.columns if c not in key_cols]
    agg_spec = {c: 'first' for c in other_cols}
    agg_spec['SHIP DATE'] = 'max'
    agg_spec['QUANTITY'] = 'sum'
    df = df.groupby(key_cols, as_index=False, sort=False).agg(agg_spec)
    aggregated_rows = len(df)

    module_path = os.path.join(data_folder, '202.csv')
    if os.path.exists(module_path):
        module_df = pd.read_csv(module_path, dtype=str)
        module_df['ORDER NO'] = module_df['ORDER NO'].astype(str).str.strip()
        module_df['PARTS NO'] = module_df['PARTS NO'].astype(str).str.strip()
        module_counts = (
            module_df.groupby(['ORDER NO', 'PARTS NO'])
            .size()
            .reset_index(name='MODULE_COUNT')
        )
        df = df.merge(
            module_counts,
            how='left',
            left_on=['CUSTOMER ORDER NO.', 'PRODUCT NO.'],
            right_on=['ORDER NO', 'PARTS NO']
        )
        df = df.drop(columns=['ORDER NO', 'PARTS NO'], errors='ignore')
    else:
        df['MODULE_COUNT'] = np.nan

    df['MODULE_COUNT'] = pd.to_numeric(df['MODULE_COUNT'], errors='coerce').fillna(0)

    df['PLAN SHIP DATE'] = df['PLAN SHIP DATE'].dt.normalize()
    df['SHIP DATE'] = df['SHIP DATE'].dt.normalize()
    cutoff_date = pd.Timestamp.today().normalize() - timedelta(days=1)
    df = df[df['PLAN SHIP DATE'] <= cutoff_date].copy()

    df['SHORTAGE_TYPE'] = df.apply(
        lambda row: classify_shortage(row['PLAN SHIP DATE'], row['SHIP DATE']),
        axis=1
    )

    label_map = {
        'normal': 'Normal',
        'delay': 'Shortage - Delay',
        'early': 'Earlier',
        'not_shipped': 'Shortage - Not Shipped'
    }
    df['SHORTAGE_LABEL'] = df['SHORTAGE_TYPE'].map(label_map)
    df['SHORTAGE_NOTE'] = df.apply(
        lambda row: build_shortage_note(row['PLAN SHIP DATE'], row['SHIP DATE'], row['SHORTAGE_TYPE']),
        axis=1
    )

    df['NORMAL_ORDERS'] = (df['SHORTAGE_TYPE'] == 'normal').astype(int)
    df['DELAY_ORDERS'] = (df['SHORTAGE_TYPE'] == 'delay').astype(int)
    df['EARLY_ORDERS'] = (df['SHORTAGE_TYPE'] == 'early').astype(int)
    df['NOT_SHIPPED_ORDERS'] = (df['SHORTAGE_TYPE'] == 'not_shipped').astype(int)

    df['NORMAL_QTY'] = np.where(df['SHORTAGE_TYPE'] == 'normal', df['MODULE_COUNT'], 0)
    df['DELAY_QTY'] = np.where(df['SHORTAGE_TYPE'] == 'delay', df['MODULE_COUNT'], 0)
    df['EARLY_QTY'] = np.where(df['SHORTAGE_TYPE'] == 'early', df['MODULE_COUNT'], 0)
    df['NOT_SHIPPED_QTY'] = np.where(df['SHORTAGE_TYPE'] == 'not_shipped', df['MODULE_COUNT'], 0)

    load_info = {
        'raw_rows': raw_rows,
        'invalid_plan_rows': invalid_plan_rows,
        'cancelled_rows': cancelled_rows,
        'aggregated_rows': aggregated_rows,
        'usable_rows': len(df),
        'min_plan_date': df['PLAN SHIP DATE'].min() if not df.empty else None,
        'max_plan_date': df['PLAN SHIP DATE'].max() if not df.empty else None,
    }

    return df, load_info


_DAILY_INT_COLS = [
    'total_orders', 'total_lines', 'total_qty',
    'normal_lines', 'delay_lines', 'early_lines', 'not_shipped_lines',
    'normal_qty', 'delay_qty', 'early_qty', 'not_shipped_qty',
    'shortage_lines', 'shortage_qty',
    'delay_total_lines', 'delay_total_qty',
]
_DAILY_PCT_COLS = ['normal_pct', 'shortage_pct']
_DAILY_ALL_COLS = ['PLAN SHIP DATE', *_DAILY_INT_COLS, *_DAILY_PCT_COLS]


def fill_missing_dates_shortage(daily_summary, date_col, start_date, end_date):
    """Fill in all missing business days in the range with zero-value rows (skips weekends)."""
    all_dates = pd.date_range(start=start_date, end=end_date, freq='B')
    if daily_summary.empty:
        cols = {date_col: all_dates}
        for c in _DAILY_INT_COLS:
            cols[c] = 0
        for c in _DAILY_PCT_COLS:
            cols[c] = 0.0
        return pd.DataFrame(cols)
    # Preserve any real data dates that fall outside the business-day grid
    # (e.g. a plan ship date that lands on a weekend). Without this, the
    # left-merge below would drop those rows, undercounting both the chart
    # and the headline totals that sum this summary.
    actual_dates = pd.DatetimeIndex(daily_summary[date_col].dropna().unique())
    all_dates = all_dates.union(actual_dates).sort_values()
    full_df = pd.DataFrame({date_col: all_dates})
    result = full_df.merge(daily_summary, on=date_col, how='left')
    for col in _DAILY_INT_COLS:
        if col in result.columns:
            result[col] = result[col].fillna(0).astype(int)
    for col in _DAILY_PCT_COLS:
        if col in result.columns:
            result[col] = result[col].fillna(0.0)
    return result


def build_daily_summary(period_data):
    """Aggregate raw rows into per-day shortage classification counts.

    Each post-aggregation row in `period_data` is a unique (order, part) line
    item, so the count columns are named ``*_lines`` to make the unit explicit.
    Distinct order counts (``nunique('CUSTOMER ORDER NO.')``) are computed in
    ``generate_summary_stats`` for the headline stat card.
    """
    if period_data is None or period_data.empty:
        return pd.DataFrame(columns=_DAILY_ALL_COLS)

    daily_summary = period_data.groupby('PLAN SHIP DATE').agg(
        total_orders=('CUSTOMER ORDER NO.', 'nunique'),
        total_lines=('CUSTOMER ORDER NO.', 'size'),
        total_qty=('MODULE_COUNT', 'sum'),
        normal_lines=('NORMAL_ORDERS', 'sum'),
        delay_lines=('DELAY_ORDERS', 'sum'),
        early_lines=('EARLY_ORDERS', 'sum'),
        not_shipped_lines=('NOT_SHIPPED_ORDERS', 'sum'),
        normal_qty=('NORMAL_QTY', 'sum'),
        delay_qty=('DELAY_QTY', 'sum'),
        early_qty=('EARLY_QTY', 'sum'),
        not_shipped_qty=('NOT_SHIPPED_QTY', 'sum')
    ).reset_index()

    # Shortage = delay + not_shipped. Early shipments are NOT a shortage (the
    # customer received parts ahead of plan) — they're tracked as their own
    # neutral category instead. Applied consistently everywhere.
    daily_summary['shortage_lines'] = (
        daily_summary['delay_lines']
        + daily_summary['not_shipped_lines']
    )
    daily_summary['shortage_qty'] = (
        daily_summary['delay_qty']
        + daily_summary['not_shipped_qty']
    )
    # delay_total = delay + not_shipped (kept for the "Delay" stack + table).
    daily_summary['delay_total_lines'] = (
        daily_summary['delay_lines'] + daily_summary['not_shipped_lines']
    )
    daily_summary['delay_total_qty'] = (
        daily_summary['delay_qty'] + daily_summary['not_shipped_qty']
    )
    totals = daily_summary['total_lines'].replace(0, pd.NA)
    daily_summary['normal_pct'] = (
        daily_summary['normal_lines'] / totals * 100
    ).round(1).fillna(0.0)
    daily_summary['shortage_pct'] = (
        daily_summary['shortage_lines'] / totals * 100
    ).round(1).fillna(0.0)

    daily_summary = daily_summary.sort_values('PLAN SHIP DATE').reset_index(drop=True)
    return daily_summary


def create_daily_summary_calendar_month(df, target_month=None, target_year=None):
    """Returns (daily_summary, period_data, mode, month, year). ``period_data``
    is the raw rows in-scope so callers can compute distinct order counts."""
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

    # Fill in all dates in the month up to the last day with data
    if not daily_summary.empty:
        month_start = pd.Timestamp(year=target_year, month=target_month, day=1)
        month_end = daily_summary['PLAN SHIP DATE'].max()
        daily_summary = fill_missing_dates_shortage(daily_summary, 'PLAN SHIP DATE', month_start, month_end)

    return daily_summary, month_data, 'month', target_month, target_year


def create_daily_summary_past_2_months(df):
    end_date = pd.Timestamp.today().normalize() - timedelta(days=1)
    start_date = end_date - timedelta(days=60)

    period_data = df[
        (df['PLAN SHIP DATE'] >= start_date) &
        (df['PLAN SHIP DATE'] <= end_date)
    ]

    daily_summary = build_daily_summary(period_data)

    # Fill in all dates in the range
    daily_summary = fill_missing_dates_shortage(daily_summary, 'PLAN SHIP DATE', start_date, end_date)

    return daily_summary, period_data, 'past_2_months', start_date, end_date


def create_previous_month_summary(df, target_month, target_year):
    """One-shot summary for the month before (target_month, target_year), used
    for the MoM delta on the Normal stat card. Returns None if no data."""
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
    normal = int((prev_data['SHORTAGE_TYPE'] == 'normal').sum())
    return {
        'month': prev_last.month,
        'year': prev_last.year,
        'total_lines': total,
        'normal_lines': normal,
    }


def _rolling_normal_pct(daily_summary, window=7):
    """Return a list of rolling Normal % values (window of N active days).

    Active-day-only window: filled zero-rows are gaps (None) so empty
    business days don't drag the line to 0. Days without enough history
    yet also return None (gap in the line).
    """
    values = []
    active = []
    for _, row in daily_summary.iterrows():
        total = int(row.get('total_lines', 0) or 0)
        if total > 0:
            normal = int(row.get('normal_lines', 0) or 0)
            active.append((normal, total))
            active = active[-window:]
            tot = sum(t for _, t in active)
            on = sum(o for o, _ in active)
            values.append(round(on / tot * 100, 1) if tot > 0 else None)
        else:
            values.append(None)
    return values


def build_shortage_details(df, start_date, end_date):
    details = df[
        (df['PLAN SHIP DATE'] >= start_date) &
        (df['PLAN SHIP DATE'] <= end_date) &
        (df['SHORTAGE_TYPE'] != 'normal')
    ].copy()

    details = details.sort_values(['PLAN SHIP DATE', 'CUSTOMER ORDER NO.'])
    return details


def build_top_shortage_parts(df, start_date, end_date, top_n=None):
    shortage_df = df[
        (df['PLAN SHIP DATE'] >= start_date) &
        (df['PLAN SHIP DATE'] <= end_date) &
        (df['SHORTAGE_TYPE'] != 'normal')
    ]

    summary = shortage_df.groupby('PRODUCT NO.').agg(
        delay_orders=('DELAY_ORDERS', 'sum'),
        early_orders=('EARLY_ORDERS', 'sum'),
        not_shipped_orders=('NOT_SHIPPED_ORDERS', 'sum'),
        delay_qty=('DELAY_QTY', 'sum'),
        early_qty=('EARLY_QTY', 'sum'),
        not_shipped_qty=('NOT_SHIPPED_QTY', 'sum')
    ).reset_index()

    # Shortage = delay + not_shipped. Early shipments keep their own columns
    # (still identified) but are excluded from the shortage totals/ranking.
    summary['shortage_orders'] = summary['delay_orders'] + summary['not_shipped_orders']
    summary['shortage_qty'] = summary['delay_qty'] + summary['not_shipped_qty']

    summary = summary.sort_values(
        ['shortage_orders', 'shortage_qty'], ascending=[False, False]
    )

    if top_n is not None:
        summary = summary.head(top_n)

    return summary


def format_period_label(mode, period_info):
    if mode == 'month':
        month, year = period_info
        return datetime(year, month, 1).strftime('%B %Y')
    start_date, end_date = period_info
    return f"{start_date.strftime('%b %d, %Y')} - {end_date.strftime('%b %d, %Y')}"


def create_visualization(daily_summary, mode, period_info):
    """Stacked bar chart (Normal / Delay / Earlier) plus a rolling 7-day
    Normal % line on a secondary y-axis. No horizontal target line yet —
    set ``TARGET_PCT`` at the top of this module to enable one."""
    fig = go.Figure()
    plot_dates = daily_summary['PLAN SHIP DATE']

    normal_counts = daily_summary['normal_lines'].astype(int)
    delay_total_counts = daily_summary['delay_total_lines'].astype(int)
    early_counts = daily_summary['early_lines'].astype(int)
    delay_only_counts = daily_summary['delay_lines'].astype(int)
    not_shipped_counts = daily_summary['not_shipped_lines'].astype(int)
    totals = daily_summary['total_lines'].astype(int)
    max_total = int(totals.max()) if len(totals) else 0

    def pct_of(n, t):
        return (n / t * 100) if t > 0 else 0.0

    # Adaptive label positioning on the Normal bar: short bars push their
    # percentage label outside so it stays readable.
    label_threshold = max(1, int(max_total * 0.08))
    normal_text, normal_positions = [], []
    for n, t, pct in zip(normal_counts, totals, daily_summary['normal_pct']):
        if t <= 0 or n <= 0:
            normal_text.append('')
            normal_positions.append('inside')
        elif n < label_threshold:
            normal_text.append(f"{pct:.1f}%")
            normal_positions.append('outside')
        else:
            normal_text.append(f"{pct:.1f}%")
            normal_positions.append('inside')

    # Per-trace customdata under hovermode='x unified': each trace prints its
    # own labelled line so the tooltip reads like a mini-legend.
    normal_hover = [
        f"{int(n):,} orders ({pct_of(n, t):.1f}%)" if t > 0 else '—'
        for n, t in zip(normal_counts, totals)
    ]
    delay_hover = [
        f"{int(d):,} delayed + {int(ns):,} not shipped ({pct_of(int(d) + int(ns), t):.1f}%)"
        if t > 0 else '—'
        for d, ns, t in zip(delay_only_counts, not_shipped_counts, totals)
    ]
    early_hover = [
        f"{int(n):,} orders ({pct_of(n, t):.1f}%)" if t > 0 else '—'
        for n, t in zip(early_counts, totals)
    ]

    fig.add_trace(go.Bar(
        name='Normal',
        x=plot_dates,
        y=normal_counts,
        marker_color='#28a745',
        text=normal_text,
        textposition=normal_positions,
        textfont=dict(color='white', size=10, family='Arial'),
        cliponaxis=False,
        hovertemplate='<b>Normal</b>: %{customdata}<extra></extra>',
        customdata=normal_hover,
    ))

    fig.add_trace(go.Bar(
        name='Shortage - Delay',
        x=plot_dates,
        y=delay_total_counts,
        marker_color='#dc3545',
        text='',
        hovertemplate='<b>Delay</b>: %{customdata}<extra></extra>',
        customdata=delay_hover,
    ))

    fig.add_trace(go.Bar(
        name='Earlier',
        x=plot_dates,
        y=early_counts,
        marker_color='#f4b400',
        text='',
        hovertemplate='<b>Earlier</b>: %{customdata}<extra></extra>',
        customdata=early_hover,
    ))

    # Rolling 7-day Normal % on a secondary y-axis.
    rolling_vals = _rolling_normal_pct(daily_summary, window=7)
    fig.add_trace(go.Scatter(
        name='Rolling 7-day Normal %',
        x=plot_dates,
        y=rolling_vals,
        mode='lines+markers',
        yaxis='y2',
        line=dict(color='#2c3e50', width=2, dash='dot'),
        marker=dict(size=5),
        hovertemplate='<b>7-day rolling avg</b>: %{y:.1f}%<extra></extra>',
        connectgaps=False,
    ))

    # Optional horizontal target line — disabled until the business agrees on
    # a number. Set TARGET_PCT at the top of this file to enable.
    if TARGET_PCT is not None and len(plot_dates) > 0:
        fig.add_trace(go.Scatter(
            name=f'Target {TARGET_PCT:.0f}%',
            x=[plot_dates.iloc[0], plot_dates.iloc[-1]],
            y=[TARGET_PCT, TARGET_PCT],
            mode='lines',
            yaxis='y2',
            line=dict(color='#888', width=1.5, dash='dash'),
            hoverinfo='skip',
            showlegend=True,
        ))

    period_label = format_period_label(mode, period_info)
    fig.update_layout(
        title=dict(
            text=f"Shortage KPI - {period_label}",
            font=dict(size=24, color='#2c3e50'),
            x=0.5,
            xanchor='center'
        ),
        xaxis_title='Plan Ship Date',
        yaxis_title='Order Count',
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
            showgrid=True,
            gridcolor='#ecf0f1',
            linecolor='#bdc3c7',
            linewidth=2
        ),
        yaxis2=dict(
            title='Normal %',
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
        height=620,
        autosize=True,
        margin=dict(l=80, r=80, t=150, b=100)
    )

    return fig


def generate_summary_stats(daily_summary, period_label, period_data=None, prev_month=None):
    """Aggregate stats for the report panel.

    ``period_data`` is the raw filtered rows (each row a unique (order, part)
    line item); used for the distinct-order count. ``prev_month`` (optional)
    enables the MoM Normal % delta on the headline stat card.
    """
    total_lines = int(daily_summary['total_lines'].sum())
    total_qty = int(daily_summary['total_qty'].sum())
    normal_lines = int(daily_summary['normal_lines'].sum())
    delay_lines = int(daily_summary['delay_lines'].sum())
    delay_total_lines = int(daily_summary['delay_total_lines'].sum())
    early_lines = int(daily_summary['early_lines'].sum())
    not_shipped_lines = int(daily_summary['not_shipped_lines'].sum())
    # Early shipments are tracked separately and excluded from shortage.
    shortage_lines = delay_lines + not_shipped_lines
    normal_qty = int(daily_summary['normal_qty'].sum())
    delay_qty = int(daily_summary['delay_qty'].sum())
    delay_total_qty = int(daily_summary['delay_total_qty'].sum())
    early_qty = int(daily_summary['early_qty'].sum())
    not_shipped_qty = int(daily_summary['not_shipped_qty'].sum())
    shortage_qty = delay_qty + not_shipped_qty

    normal_pct = (normal_lines / total_lines * 100) if total_lines else 0.0
    delay_pct = (delay_lines / total_lines * 100) if total_lines else 0.0
    early_pct = (early_lines / total_lines * 100) if total_lines else 0.0
    shortage_pct = (shortage_lines / total_lines * 100) if total_lines else 0.0

    day_count = max(len(daily_summary), 1)
    avg_daily_lines = total_lines / day_count
    avg_daily_qty = total_qty / day_count

    # Distinct order count for the headline stat card. The daily summary
    # counts rows (line items) — customers think in orders, so use nunique.
    if period_data is not None and not period_data.empty and 'CUSTOMER ORDER NO.' in period_data.columns:
        total_orders = int(period_data['CUSTOMER ORDER NO.'].nunique())
    else:
        total_orders = 0

    mom = None
    if prev_month is not None and prev_month.get('total_lines', 0) > 0:
        prev_pct = prev_month['normal_lines'] / prev_month['total_lines'] * 100
        mom = {
            'prev_label': datetime(prev_month['year'], prev_month['month'], 1).strftime('%B %Y'),
            'prev_normal_pct': round(prev_pct, 1),
            'delta_pts': round(normal_pct - prev_pct, 1),
        }

    return {
        'period_label': period_label,
        # Distinct order count for the headline card
        'total_orders': total_orders,
        # Line-item counts (each row = one order/part line)
        'total_lines': total_lines,
        'total_qty': total_qty,
        'avg_daily_lines': avg_daily_lines,
        'avg_daily_qty': avg_daily_qty,
        'normal_lines': normal_lines,
        'normal_qty': normal_qty,
        'delay_lines': delay_lines,
        'delay_qty': delay_qty,
        'delay_total_lines': delay_total_lines,
        'delay_total_qty': delay_total_qty,
        'early_lines': early_lines,
        'early_qty': early_qty,
        'not_shipped_lines': not_shipped_lines,
        'not_shipped_qty': not_shipped_qty,
        'shortage_lines': shortage_lines,
        'shortage_qty': shortage_qty,
        'normal_pct': normal_pct,
        'delay_pct': delay_pct,
        'early_pct': early_pct,
        'shortage_pct': shortage_pct,
        'mom': mom,
        'target_pct': TARGET_PCT,
    }


def format_display_date(value, default=''):
    if pd.isna(value):
        return default
    return pd.Timestamp(value).strftime('%b %d, %Y')


def build_daily_summary_json(daily_summary):
    """JSON shape consumed by the custom-range JS. Keys use ``*_lines`` (post-
    aggregation row counts), but the legacy ``*_orders`` aliases are also
    emitted so any external consumers of the embedded JSON keep working."""
    js_data = []
    for _, row in daily_summary.iterrows():
        total_lines = int(row['total_lines'])
        normal_lines = int(row['normal_lines'])
        delay_lines = int(row['delay_lines'])
        early_lines = int(row['early_lines'])
        not_shipped_lines = int(row['not_shipped_lines'])
        delay_total_lines = int(row['delay_total_lines'])
        js_data.append({
            'date': row['PLAN SHIP DATE'].strftime('%Y-%m-%d'),
            'total_orders': int(row['total_orders']),
            'total_lines': total_lines,
            'total_qty': int(row['total_qty']),
            'normal_lines': normal_lines,
            'delay_lines': delay_lines,
            'early_lines': early_lines,
            'not_shipped_lines': not_shipped_lines,
            'normal_qty': int(row['normal_qty']),
            'delay_qty': int(row['delay_qty']),
            'early_qty': int(row['early_qty']),
            'not_shipped_qty': int(row['not_shipped_qty']),
            'delay_total_lines': delay_total_lines,
            'delay_total_qty': int(row['delay_total_qty']),
            'normal_pct': float(row['normal_pct']),
            'shortage_pct': float(row['shortage_pct']),
        })
    return js_data


def build_shortage_details_json(details):
    detail_rows = []
    for _, row in details.iterrows():
        detail_rows.append({
            'plan_date': row['PLAN SHIP DATE'].strftime('%Y-%m-%d'),
            'ship_date': row['SHIP DATE'].strftime('%Y-%m-%d') if pd.notna(row['SHIP DATE']) else '',
            'order_no': row['CUSTOMER ORDER NO.'],
            'part_no': row['PRODUCT NO.'],
            'module_count': int(row['MODULE_COUNT']),
            'shortage_label': row['SHORTAGE_LABEL'],
            'shortage_note': row['SHORTAGE_NOTE']
        })
    return detail_rows


def render_top_parts_table(summary_df):
    if summary_df.empty:
        return '<p style="color: #7f8c8d;">No shortage parts found for this period.</p>'

    rows = ''
    for _, row in summary_df.iterrows():
        rows += (
            '<tr>'
            f"<td><strong>{row['PRODUCT NO.']}</strong></td>"
            f"<td>{int(row['shortage_orders']):,}</td>"
            f"<td>{int(row['shortage_qty']):,}</td>"
            f"<td>{int(row['delay_orders']):,}</td>"
            f"<td>{int(row['delay_qty']):,}</td>"
            f"<td>{int(row['early_orders']):,}</td>"
            f"<td>{int(row['early_qty']):,}</td>"
            f"<td>{int(row['not_shipped_orders']):,}</td>"
            f"<td>{int(row['not_shipped_qty']):,}</td>"
            '</tr>'
        )

    return (
        '<div class="table-scroll">'
        '<table class="top-parts-table">'
        '<thead>'
        '<tr>'
        '<th>Part No.</th>'
        '<th>Shortage Orders</th>'
        '<th>Shortage Module Count</th>'
        '<th>Delay Orders</th>'
        '<th>Delay Module Count</th>'
        '<th>Early Orders</th>'
        '<th>Early Module Count</th>'
        '<th>Not Shipped</th>'
        '<th>Not Shipped Module Count</th>'
        '</tr>'
        '</thead>'
        '<tbody>'
        f"{rows}"
        '</tbody>'
        '</table>'
        '</div>'
    )


def render_shortage_details_table(details_df):
    if details_df.empty:
        return '<p style="color: #7f8c8d;">No shortage orders found for this period.</p>'

    rows = ''
    for _, row in details_df.iterrows():
        plan_date = format_display_date(row['PLAN SHIP DATE'])
        ship_date = format_display_date(row['SHIP DATE'], default='Not shipped')
        rows += (
            '<tr>'
            f"<td><strong>{plan_date}</strong></td>"
            f"<td>{ship_date}</td>"
            f"<td>{row['CUSTOMER ORDER NO.']}</td>"
            f"<td>{row['PRODUCT NO.']}</td>"
            f"<td>{int(row['MODULE_COUNT']):,}</td>"
            f"<td>{row['SHORTAGE_NOTE']}</td>"
            '</tr>'
        )

    return (
        '<div class="table-scroll">'
        '<table>'
        '<thead>'
        '<tr>'
        '<th>Plan Date</th>'
        '<th>Actual Ship Date</th>'
        '<th>Order No.</th>'
        '<th>Part No.</th>'
        '<th>Module Count</th>'
        '<th>Delay / Earlier</th>'
        '</tr>'
        '</thead>'
        '<tbody>'
        f"{rows}"
        '</tbody>'
        '</table>'
        '</div>'
    )


def _render_mom_line(stats):
    """Render the MoM Normal-% delta line shown on the headline stat card."""
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
        f"(was {mom['prev_normal_pct']:.1f}%)</div>"
    )


def _render_load_info(load_info):
    """Render the data-load summary line in the report footer."""
    if not load_info:
        return ''
    parts = [f"raw rows: {load_info.get('raw_rows', 0):,}"]
    invalid = load_info.get('invalid_plan_rows', 0)
    if invalid:
        parts.append(f"dropped (invalid plan date): {invalid:,}")
    cancelled = load_info.get('cancelled_rows', 0)
    if cancelled:
        parts.append(f"cancelled: {cancelled:,}")
    aggregated = load_info.get('aggregated_rows')
    if aggregated is not None:
        parts.append(f"unique (order, part) orders: {aggregated:,}")
    parts.append(f"usable: {load_info.get('usable_rows', 0):,}")
    min_d = load_info.get('min_plan_date')
    max_d = load_info.get('max_plan_date')
    if min_d is not None and max_d is not None and not pd.isna(min_d) and not pd.isna(max_d):
        parts.append(
            f"plan-ship range: {pd.Timestamp(min_d).strftime('%b %d, %Y')} – {pd.Timestamp(max_d).strftime('%b %d, %Y')}"
        )
    return f"<p style='margin-top: 8px; font-size: 0.85em; color: #6c757d;'>Data: {' | '.join(parts)}</p>"


def create_html_report(
    fig_month,
    stats_month,
    daily_summary_month,
    top_parts_month,
    details_month,
    fig_past2months,
    stats_past2months,
    daily_summary_past2months,
    top_parts_past2months,
    details_past2months,
    df,
    month,
    year,
    load_info=None,
):
    chart_month_html = fig_month.to_html(
        include_plotlyjs='cdn',
        div_id='chart-month',
        config={'displayModeBar': False, 'responsive': True}
    )
    chart_past2months_html = fig_past2months.to_html(
        include_plotlyjs=False,
        div_id='chart-past2months',
        config={'displayModeBar': False, 'responsive': True}
    )

    all_daily_summary = build_daily_summary(df)
    if not all_daily_summary.empty:
        all_daily_summary = fill_missing_dates_shortage(
            all_daily_summary, 'PLAN SHIP DATE',
            all_daily_summary['PLAN SHIP DATE'].min(),
            all_daily_summary['PLAN SHIP DATE'].max()
        )
    all_daily_json = _safe_json(build_daily_summary_json(all_daily_summary))

    all_shortage_details = df[df['SHORTAGE_TYPE'] != 'normal'].copy()
    all_details_json = _safe_json(build_shortage_details_json(all_shortage_details))

    top_parts_month_html = render_top_parts_table(top_parts_month)
    top_parts_past2months_html = render_top_parts_table(top_parts_past2months)
    details_month_html = render_shortage_details_table(details_month)
    details_past2months_html = render_shortage_details_table(details_past2months)

    # Shortage = delay + early + not_shipped (one definition, used everywhere).
    month_shortage_lines = stats_month['shortage_lines']
    month_shortage_qty = stats_month['shortage_qty']
    past2months_shortage_lines = stats_past2months['shortage_lines']
    past2months_shortage_qty = stats_past2months['shortage_qty']

    month_mom_line = _render_mom_line(stats_month)
    past2months_mom_line = _render_mom_line(stats_past2months)
    load_info_html = _render_load_info(load_info)

    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Shortage KPI Report - {stats_month['period_label']}</title>
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
            color: #d39e00;
        }}

        .stat-card .subtext {{
            color: #95a5a6;
            font-size: 0.9em;
            margin-top: 5px;
        }}

        .mom-line {{
            margin-top: 6px;
            font-size: 0.85em;
            color: #4f5b66;
        }}
        .mom-up   {{ color: #1f7a3a; font-weight: 600; }}
        .mom-down {{ color: #b3261e; font-weight: 600; }}
        .mom-flat {{ color: #6c757d; font-weight: 600; }}

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

        .details-container p.note {{
            color: #7f8c8d;
            font-size: 14px;
            margin-bottom: 20px;
        }}

        .table-scroll {{
            max-height: 420px;
            overflow-y: auto;
            border-radius: 4px;
        }}

        .table-section {{
            margin-top: 30px;
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

        .top-parts-table th:nth-child(3),
        .top-parts-table td:nth-child(3),
        .top-parts-table th:nth-child(5),
        .top-parts-table td:nth-child(5),
        .top-parts-table th:nth-child(7),
        .top-parts-table td:nth-child(7) {{
            border-right: 3px solid #b0b8bb;
        }}

        th {{
            background: #95a5a6;
            color: white;
            padding: 14px;
            text-align: left;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.8em;
            letter-spacing: 0.5px;
        }}

        td {{
            padding: 12px 14px;
            border-bottom: 1px solid #ecf0f1;
        }}

        tr:hover {{
            background: #f8f9fa;
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
            <h1>Example Logistics - Monthly Shortage KPI</h1>
            <p>This dashboard highlights shortages by comparing planned vs. actual ship dates.</p>
        </div>

        <div class="filter-bar">
            <div class="filter-row primary">
                <button class="toggle-btn active" onclick="switchView('month')" id="btn-month">{datetime(year, month, 1).strftime('%B')}</button>
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
                    <div class="value" style="font-size: 1.5em;">{stats_month['period_label']}</div>
                </div>
                <div class="stat-card">
                    <h3>Total Orders</h3>
                    <div class="value">{stats_month['total_orders']:,}</div>
                    <div class="subtext">{stats_month['total_qty']:,} module count</div>
                </div>
                <div class="stat-card success">
                    <h3>Normal</h3>
                    <div class="value">{stats_month['normal_pct']:.1f}%</div>
                    <div class="subtext">{stats_month['normal_lines']:,} orders | {stats_month['normal_qty']:,} module count</div>
                    {month_mom_line}
                </div>
                <div class="stat-card early">
                    <h3>Early</h3>
                    <div class="value">{stats_month['early_pct']:.1f}%</div>
                    <div class="subtext">{stats_month['early_lines']:,} orders | {stats_month['early_qty']:,} module count</div>
                    <div class="mom-line">Shipped ahead of plan &mdash; not counted as shortage</div>
                </div>
                <div class="stat-card danger">
                    <h3>Shortage</h3>
                    <div class="value">{stats_month['shortage_pct']:.1f}%</div>
                    <div class="subtext">{month_shortage_lines:,} orders | {month_shortage_qty:,} module count</div>
                </div>
            </div>

            <div class="chart-container">
                {chart_month_html}
            </div>
"""

    html_content += """
            <div class="details-container">
"""

    html_content += f"""
                <h2>Shortage Parts Summary</h2>
                <p class="note">Ranked by shortage orders (delay + not shipped), then module count. Early shipments are listed separately and not counted as shortage.</p>
                {top_parts_month_html}
                <div class="table-section">
                    <h2>Off-Plan Order Details</h2>
                    <p class="note">Delayed and not-shipped orders (shortages), plus early shipments shown for visibility (not counted as shortage).</p>
                    {details_month_html}
                </div>
            </div>
        </div>
"""

    html_content += f"""
        <div id="view-past2months" class="view-content">
            <div class="stats-container">
                <div class="stat-card">
                    <h3>Period</h3>
                    <div class="value" style="font-size: 1.5em;">{stats_past2months['period_label']}</div>
                </div>
                <div class="stat-card">
                    <h3>Total Orders</h3>
                    <div class="value">{stats_past2months['total_orders']:,}</div>
                    <div class="subtext">{stats_past2months['total_qty']:,} module count</div>
                </div>
                <div class="stat-card success">
                    <h3>Normal</h3>
                    <div class="value">{stats_past2months['normal_pct']:.1f}%</div>
                    <div class="subtext">{stats_past2months['normal_lines']:,} orders | {stats_past2months['normal_qty']:,} module count</div>
                    {past2months_mom_line}
                </div>
                <div class="stat-card early">
                    <h3>Early</h3>
                    <div class="value">{stats_past2months['early_pct']:.1f}%</div>
                    <div class="subtext">{stats_past2months['early_lines']:,} orders | {stats_past2months['early_qty']:,} module count</div>
                    <div class="mom-line">Shipped ahead of plan &mdash; not counted as shortage</div>
                </div>
                <div class="stat-card danger">
                    <h3>Shortage</h3>
                    <div class="value">{stats_past2months['shortage_pct']:.1f}%</div>
                    <div class="subtext">{past2months_shortage_lines:,} orders | {past2months_shortage_qty:,} module count</div>
                </div>
            </div>

            <div class="chart-container">
                {chart_past2months_html}
            </div>

            <div class="details-container">
                <h2>Shortage Parts Summary</h2>
                <p class="note">Ranked by shortage orders (delay + not shipped), then module count. Early shipments are listed separately and not counted as shortage.</p>
                {top_parts_past2months_html}
                <div class="table-section">
                    <h2>Off-Plan Order Details</h2>
                    <p class="note">Delayed and not-shipped orders (shortages), plus early shipments shown for visibility (not counted as shortage).</p>
                    {details_past2months_html}
                </div>
            </div>
        </div>

        <div class="footer">
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
            <p style="margin-top: 10px; font-size: 0.9em; color: #7f8c8d;">Dashboard created by Viktor Berg | Built with Python, Plotly, and Pandas</p>
            {load_info_html}
        </div>
    </div>
    """

    script_content = """
    <script>
        const allDailyData = __ALL_DAILY_DATA__;
        const allDetailData = __ALL_DETAIL_DATA__;

        function switchView(mode) {
            const views = ['view-month', 'view-past2months', 'view-custom'];
            const btns  = ['btn-month', 'btn-past2months', 'btn-custom-view'];
            views.forEach(id => { const el = document.getElementById(id); if (el) el.classList.remove('active'); });
            btns.forEach(id  => { const el = document.getElementById(id); if (el) el.classList.remove('active'); });

            const targetMap = {month: ['view-month','btn-month'], past2months: ['view-past2months','btn-past2months'], custom: ['view-custom','btn-custom-view']};
            const pair = targetMap[mode];
            if (!pair) return;
            pair.forEach(id => { const el = document.getElementById(id); if (el) el.classList.add('active'); });

            // Charts inside hidden views were laid out at 0×0 — force a resize
            // now that the view is visible so they fill the container width.
            const activeView = document.getElementById(pair[0]);
            if (activeView && window.Plotly) {
                activeView.querySelectorAll('.js-plotly-plot').forEach(el => Plotly.Plots.resize(el));
            }
        }

        function toggleAdvancedFilters() {
            const panel = document.getElementById('advanced-filters');
            const btn = document.getElementById('btn-advanced');
            const hidden = panel.hasAttribute('hidden');
            if (hidden) {
                panel.removeAttribute('hidden');
                btn.setAttribute('aria-expanded', 'true');
            } else {
                panel.setAttribute('hidden', '');
                btn.setAttribute('aria-expanded', 'false');
            }
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

            let customView = document.getElementById('view-custom');
            if (!customView) {
                customView = document.createElement('div');
                customView.id = 'view-custom';
                customView.className = 'view-content';
                document.querySelector('.footer').before(customView);
            }

            if (!filterDataByRange(startDate, endDate)) return;
            document.getElementById('btn-custom-view').classList.remove('hidden');
            switchView('custom');
        }

        function parseDate(dateStr) {
            const [year, month, day] = dateStr.split('-').map(Number);
            return new Date(year, month - 1, day);
        }

        function formatDateISO(d) {
            const y = d.getFullYear();
            const m = String(d.getMonth() + 1).padStart(2, '0');
            const day = String(d.getDate()).padStart(2, '0');
            return `${y}-${m}-${day}`;
        }

        function latestDataDate() {
            if (!allDailyData || allDailyData.length === 0) return new Date();
            for (let i = allDailyData.length - 1; i >= 0; i--) {
                if (allDailyData[i].total_lines > 0) return parseDate(allDailyData[i].date);
            }
            return parseDate(allDailyData[allDailyData.length - 1].date);
        }

        function applyRangeDates(startStr, endStr) {
            document.getElementById('start-date').value = startStr;
            document.getElementById('end-date').value = endStr;
            applyCustomRange();
        }

        function applyQuickRange(days) {
            const end = latestDataDate();
            const start = new Date(end);
            start.setDate(end.getDate() - (days - 1));
            applyRangeDates(formatDateISO(start), formatDateISO(end));
        }

        function formatDate(dateStr) {
            const date = parseDate(dateStr);
            return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        }

        function buildTopPartsTable(parts) {
            if (!parts.length) {
                return '<p style="color: #7f8c8d;">No shortage parts found for this period.</p>';
            }

            const rows = parts.map(part => `
                <tr>
                    <td><strong>${part.part_no}</strong></td>
                    <td>${part.shortage_orders.toLocaleString()}</td>
                    <td>${part.shortage_qty.toLocaleString()}</td>
                    <td>${part.delay_orders.toLocaleString()}</td>
                    <td>${part.delay_qty.toLocaleString()}</td>
                    <td>${part.early_orders.toLocaleString()}</td>
                    <td>${part.early_qty.toLocaleString()}</td>
                    <td>${part.not_shipped_orders.toLocaleString()}</td>
                    <td>${part.not_shipped_qty.toLocaleString()}</td>
                </tr>
            `).join('');

            return `
                <div class="table-scroll">
                    <table class="top-parts-table">
                        <thead>
                            <tr>
                                <th>Part No.</th>
                                <th>Shortage Orders</th>
                                <th>Shortage Module Count</th>
                                <th>Delay Orders</th>
                                <th>Delay Module Count</th>
                                <th>Early Orders</th>
                                <th>Early Module Count</th>
                                <th>Not Shipped</th>
                                <th>Not Shipped Module Count</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rows}
                        </tbody>
                    </table>
                </div>
            `;
        }

        function buildDetailTable(details) {
            if (!details.length) {
                return '<p style="color: #7f8c8d;">No shortage orders found for this period.</p>';
            }

            const rows = details.map(item => `
                <tr>
                    <td><strong>${formatDate(item.plan_date)}</strong></td>
                    <td>${item.ship_date ? formatDate(item.ship_date) : 'Not shipped'}</td>
                    <td>${item.order_no}</td>
                    <td>${item.part_no}</td>
                    <td>${Number(item.module_count || 0).toLocaleString()}</td>
                    <td>${item.shortage_note}</td>
                </tr>
            `).join('');

            return `
                <div class="table-scroll">
                    <table>
                        <thead>
                            <tr>
                                <th>Plan Date</th>
                                <th>Actual Ship Date</th>
                                <th>Order No.</th>
                                <th>Part No.</th>
                                <th>Module Count</th>
                                <th>Delay / Earlier</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rows}
                        </tbody>
                    </table>
                </div>
            `;
        }

        function rollingNormalPct(filteredDaily, window) {
            const out = [];
            const active = [];
            filteredDaily.forEach(item => {
                if (item.total_lines > 0) {
                    active.push({on: item.normal_lines, tot: item.total_lines});
                    if (active.length > window) active.shift();
                    const tot = active.reduce((s, x) => s + x.tot, 0);
                    const on = active.reduce((s, x) => s + x.on, 0);
                    out.push(tot > 0 ? Math.round(on / tot * 1000) / 10 : null);
                } else {
                    out.push(null);
                }
            });
            return out;
        }

        function filterDataByRange(startDate, endDate) {
            const start = parseDate(startDate);
            const end = parseDate(endDate);
            const filteredDaily = allDailyData.filter(item => {
                const itemDate = parseDate(item.date);
                return itemDate >= start && itemDate <= end;
            }).sort((a, b) => parseDate(a.date) - parseDate(b.date));

            if (!filteredDaily.length) {
                alert('No data available for the selected date range.');
                return false;
            }

            const totalLines = filteredDaily.reduce((sum, item) => sum + item.total_lines, 0);
            const totalQty = filteredDaily.reduce((sum, item) => sum + item.total_qty, 0);
            const normalLines = filteredDaily.reduce((sum, item) => sum + item.normal_lines, 0);
            const normalQty = filteredDaily.reduce((sum, item) => sum + item.normal_qty, 0);
            const delayLines = filteredDaily.reduce((sum, item) => sum + item.delay_lines, 0);
            const delayQty = filteredDaily.reduce((sum, item) => sum + item.delay_qty, 0);
            const delayTotalLines = filteredDaily.reduce((sum, item) => sum + item.delay_total_lines, 0);
            const delayTotalQty = filteredDaily.reduce((sum, item) => sum + item.delay_total_qty, 0);
            const earlyLines = filteredDaily.reduce((sum, item) => sum + item.early_lines, 0);
            const earlyQty = filteredDaily.reduce((sum, item) => sum + item.early_qty, 0);
            const notShippedLines = filteredDaily.reduce((sum, item) => sum + item.not_shipped_lines, 0);
            const notShippedQty = filteredDaily.reduce((sum, item) => sum + item.not_shipped_qty, 0);

            // Shortage excludes early shipments; early is tracked on its own.
            const shortageLines = delayLines + notShippedLines;
            const shortageQty = delayQty + notShippedQty;
            const normalPct = totalLines ? (normalLines / totalLines * 100) : 0;
            const earlyPct = totalLines ? (earlyLines / totalLines * 100) : 0;
            const shortagePct = totalLines ? (shortageLines / totalLines * 100) : 0;

            // Distinct order count. Each order maps to exactly one plan ship
            // date, so summing the per-day distinct-order counts gives the
            // exact total for the range (matches the Calendar Month headline).
            const totalOrders = filteredDaily.reduce((sum, item) => sum + (item.total_orders || 0), 0);

            const periodLabel = `${formatDate(startDate)} - ${formatDate(endDate)}`;
            const dates = filteredDaily.map(item => parseDate(item.date));
            const normalData = filteredDaily.map(item => item.normal_lines);
            const delayData = filteredDaily.map(item => item.delay_total_lines);
            const earlyData = filteredDaily.map(item => item.early_lines);
            const totalsArr = filteredDaily.map(item => item.total_lines);
            const maxTotal = totalsArr.length ? Math.max(...totalsArr) : 0;
            const labelThreshold = Math.max(1, Math.floor(maxTotal * 0.08));
            const normalText = filteredDaily.map(item =>
                item.normal_lines > 0 ? `${item.normal_pct.toFixed(1)}%` : ''
            );
            const normalPositions = filteredDaily.map(item => {
                if (item.total_lines <= 0 || item.normal_lines <= 0) return 'inside';
                return item.normal_lines < labelThreshold ? 'outside' : 'inside';
            });
            const pctOf = (n, t) => (t > 0 ? (n / t * 100) : 0);
            const normalHover = filteredDaily.map(item =>
                item.total_lines > 0
                    ? `${item.normal_lines.toLocaleString()} orders (${pctOf(item.normal_lines, item.total_lines).toFixed(1)}%)`
                    : '—'
            );
            const delayHover = filteredDaily.map(item =>
                item.total_lines > 0
                    ? `${item.delay_lines.toLocaleString()} delayed + ${item.not_shipped_lines.toLocaleString()} not shipped (${pctOf(item.delay_lines + item.not_shipped_lines, item.total_lines).toFixed(1)}%)`
                    : '—'
            );
            const earlyHover = filteredDaily.map(item =>
                item.total_lines > 0
                    ? `${item.early_lines.toLocaleString()} orders (${pctOf(item.early_lines, item.total_lines).toFixed(1)}%)`
                    : '—'
            );
            const rollingVals = rollingNormalPct(filteredDaily, 7);

            const trace1 = {
                name: 'Normal',
                x: dates,
                y: normalData,
                type: 'bar',
                marker: { color: '#28a745' },
                text: normalText,
                textposition: normalPositions,
                textfont: { color: 'white', size: 10, family: 'Arial' },
                cliponaxis: false,
                hovertemplate: '<b>Normal</b>: %{customdata}<extra></extra>',
                customdata: normalHover
            };

            const trace2 = {
                name: 'Shortage - Delay',
                x: dates,
                y: delayData,
                type: 'bar',
                marker: { color: '#dc3545' },
                text: '',
                hovertemplate: '<b>Delay</b>: %{customdata}<extra></extra>',
                customdata: delayHover
            };

            const trace3 = {
                name: 'Earlier',
                x: dates,
                y: earlyData,
                type: 'bar',
                marker: { color: '#f4b400' },
                text: '',
                hovertemplate: '<b>Earlier</b>: %{customdata}<extra></extra>',
                customdata: earlyHover
            };

            const trace4 = {
                name: 'Rolling 7-day Normal %',
                x: dates,
                y: rollingVals,
                type: 'scatter',
                mode: 'lines+markers',
                yaxis: 'y2',
                line: { color: '#2c3e50', width: 2, dash: 'dot' },
                marker: { size: 5 },
                hovertemplate: '<b>7-day rolling avg</b>: %{y:.1f}%<extra></extra>',
                connectgaps: false
            };

            const layout = {
                title: {
                    text: `Shortage KPI - ${periodLabel}`,
                    font: { size: 24, color: '#2c3e50' },
                    x: 0.5,
                    xanchor: 'center'
                },
                barmode: 'stack',
                hovermode: 'x unified',
                plot_bgcolor: 'white',
                paper_bgcolor: 'white',
                font: { family: 'Arial, sans-serif', size: 12, color: '#2c3e50' },
                xaxis: {
                    title: { text: 'Plan Ship Date' },
                    tickformat: '%b %d',
                    dtick: 86400000,
                    tickangle: -45,
                    showgrid: true,
                    gridcolor: '#ecf0f1',
                    linecolor: '#bdc3c7',
                    linewidth: 2,
                    rangebreaks: [{bounds: ["sat", "mon"]}]
                },
                yaxis: {
                    title: { text: 'Order Count' },
                    showgrid: true,
                    gridcolor: '#ecf0f1',
                    linecolor: '#bdc3c7',
                    linewidth: 2
                },
                yaxis2: {
                    title: 'Normal %',
                    overlaying: 'y',
                    side: 'right',
                    range: [0, 105],
                    showgrid: false,
                    ticksuffix: '%',
                    linecolor: '#bdc3c7',
                    linewidth: 2
                },
                legend: {
                    orientation: 'h',
                    yanchor: 'bottom',
                    y: 1.02,
                    xanchor: 'center',
                    x: 0.5,
                    font: { size: 14 }
                },
                height: 620,
                autosize: true,
                margin: { l: 80, r: 80, t: 150, b: 100 }
            };

            const filteredDetails = allDetailData.filter(item => {
                if (!item.plan_date) return false;
                const planDate = parseDate(item.plan_date);
                return planDate >= start && planDate <= end;
            }).sort((a, b) => {
                const dateDiff = parseDate(a.plan_date) - parseDate(b.plan_date);
                if (dateDiff !== 0) return dateDiff;
                return (a.order_no || '').localeCompare(b.order_no || '');
            });

            const partMap = {};
            filteredDetails.forEach(item => {
                const part = item.part_no || 'Unknown';
                if (!partMap[part]) {
                    partMap[part] = {
                        part_no: part,
                        shortage_orders: 0,
                        shortage_qty: 0,
                        delay_orders: 0,
                        delay_qty: 0,
                        early_orders: 0,
                        early_qty: 0,
                        not_shipped_orders: 0,
                        not_shipped_qty: 0
                    };
                }

                const moduleCount = Number(item.module_count || 0);
                const label = (item.shortage_label || '').toLowerCase();
                // Shortage totals = delay + not_shipped only. Early rows still
                // populate their own columns but don't count toward shortage.
                if (label.includes('delay')) {
                    partMap[part].delay_orders += 1;
                    partMap[part].delay_qty += moduleCount;
                    partMap[part].shortage_orders += 1;
                    partMap[part].shortage_qty += moduleCount;
                } else if (label.includes('earlier')) {
                    partMap[part].early_orders += 1;
                    partMap[part].early_qty += moduleCount;
                } else if (label.includes('not shipped')) {
                    partMap[part].not_shipped_orders += 1;
                    partMap[part].not_shipped_qty += moduleCount;
                    partMap[part].shortage_orders += 1;
                    partMap[part].shortage_qty += moduleCount;
                }
            });

            const topParts = Object.values(partMap)
                .sort((a, b) => (b.shortage_orders - a.shortage_orders) || (b.shortage_qty - a.shortage_qty));

            const customView = document.getElementById('view-custom');
            customView.innerHTML = `
                <div class="stats-container">
                    <div class="stat-card">
                        <h3>Period</h3>
                        <div class="value" style="font-size: 1.5em;">${periodLabel}</div>
                    </div>
                    <div class="stat-card">
                        <h3>Total Orders</h3>
                        <div class="value">${totalOrders.toLocaleString()}</div>
                        <div class="subtext">${totalQty.toLocaleString()} module count</div>
                    </div>
                    <div class="stat-card success">
                        <h3>Normal</h3>
                        <div class="value">${normalPct.toFixed(1)}%</div>
                        <div class="subtext">${normalLines.toLocaleString()} orders | ${normalQty.toLocaleString()} module count</div>
                    </div>
                    <div class="stat-card early">
                        <h3>Early</h3>
                        <div class="value">${earlyPct.toFixed(1)}%</div>
                        <div class="subtext">${earlyLines.toLocaleString()} orders | ${earlyQty.toLocaleString()} module count</div>
                        <div class="mom-line">Shipped ahead of plan &mdash; not counted as shortage</div>
                    </div>
                    <div class="stat-card danger">
                        <h3>Shortage</h3>
                        <div class="value">${shortagePct.toFixed(1)}%</div>
                        <div class="subtext">${shortageLines.toLocaleString()} orders | ${shortageQty.toLocaleString()} module count</div>
                    </div>
                </div>
                <div class="chart-container">
                    <div id="custom-chart"></div>
                </div>
                <div class="details-container">
                    <h2>Shortage Parts Summary</h2>
                    <p class="note">Ranked by shortage orders (delay + not shipped), then module count. Early shipments are listed separately and not counted as shortage.</p>
                    ${buildTopPartsTable(topParts)}
                    <div class="table-section">
                        <h2>Off-Plan Order Details</h2>
                        <p class="note">Delayed and not-shipped orders (shortages), plus early shipments shown for visibility (not counted as shortage).</p>
                        ${buildDetailTable(filteredDetails)}
                    </div>
                </div>
            `;

            Plotly.newPlot('custom-chart', [trace1, trace2, trace3, trace4], layout, {
                displayModeBar: false,
                responsive: true
            });
            return true;
        }
    </script>
    """

    script_content = script_content.replace('__ALL_DAILY_DATA__', all_daily_json)
    script_content = script_content.replace('__ALL_DETAIL_DATA__', all_details_json)

    html_content += script_content
    html_content += """
    </body>
    </html>
    """

    return html_content


def main():
    parser = argparse.ArgumentParser(
        description='Generate Monthly Shortage KPI Report with interactive toggle views',
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
    df, load_info = load_and_process_data(data_folder)
    print(f"Total records processed: {len(df):,}")

    print("\nGenerating calendar month report...")
    daily_summary_month, period_data_month, mode_month, month, year = create_daily_summary_calendar_month(
        df, target_month=args.month, target_year=args.year
    )
    fig_month = create_visualization(daily_summary_month, mode_month, (month, year))
    prev_month_summary = create_previous_month_summary(df, month, year)
    stats_month = generate_summary_stats(
        daily_summary_month,
        format_period_label(mode_month, (month, year)),
        period_data=period_data_month,
        prev_month=prev_month_summary,
    )

    if len(daily_summary_month) > 0:
        start_month = daily_summary_month['PLAN SHIP DATE'].min()
        end_month = daily_summary_month['PLAN SHIP DATE'].max()
    else:
        start_month = pd.Timestamp(year, month, 1)
        end_month = (start_month + pd.offsets.MonthEnd(0)).normalize()
    top_parts_month = build_top_shortage_parts(df, start_month, end_month)
    details_month = build_shortage_details(df, start_month, end_month)

    print("\nGenerating past 2 months report...")
    (daily_summary_past2months, period_data_past2months, mode_past2months,
     start_past2months, end_past2months) = create_daily_summary_past_2_months(df)
    fig_past2months = create_visualization(daily_summary_past2months, mode_past2months, (start_past2months, end_past2months))
    stats_past2months = generate_summary_stats(
        daily_summary_past2months,
        format_period_label(mode_past2months, (start_past2months, end_past2months)),
        period_data=period_data_past2months,
    )

    top_parts_past2months = build_top_shortage_parts(df, start_past2months, end_past2months)
    details_past2months = build_shortage_details(df, start_past2months, end_past2months)

    html_content = create_html_report(
        fig_month,
        stats_month,
        daily_summary_month,
        top_parts_month,
        details_month,
        fig_past2months,
        stats_past2months,
        daily_summary_past2months,
        top_parts_past2months,
        details_past2months,
        df,
        month,
        year,
        load_info=load_info,
    )

    html_output_file = f"Shortage_KPI_Report_{year}_{month:02d}.html"
    with open(html_output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"\n[OK] HTML report generated successfully: {html_output_file}")
    print("   Open the file in your web browser to view the report.")
    print("   Use the toggle buttons to switch between Calendar Month, Past 2 Months, and Custom Range views.")


if __name__ == '__main__':
    main()
