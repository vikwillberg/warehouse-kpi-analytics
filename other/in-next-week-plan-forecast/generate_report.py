"""
Next Week Plan & Forecast Report Generator
Automates the DS Weekly Order reporting process
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
from pathlib import Path
import plotly.graph_objects as go
import json

# Configuration
DATA_DIR = Path("Data")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

def get_next_business_days(start_date=None, n=5):
    """Get next Monday-Friday work week. If start_date is provided, use it as-is."""
    if start_date is None:
        today = datetime.now().date()
        # Advance to next Monday
        days_until_monday = (7 - today.weekday()) % 7
        if days_until_monday == 0:  # today is Monday, go to next Monday
            days_until_monday = 7
        start_date = today + timedelta(days=days_until_monday)
    
    business_days = []
    current = start_date
    
    # Include start_date itself if it's a weekday
    if current.weekday() < 5:
        business_days.append(current)
    
    while len(business_days) < n:
        current += timedelta(days=1)
        if current.weekday() < 5:
            business_days.append(current)
    return business_days

def map_country_code(supplier_class):
    """Map supplier class codes to country abbreviations"""
    if pd.isna(supplier_class):
        return 'Unknown'
    
    mapping = {
        'JS': 'JS',  # Japan
        'IS': 'IS',  # Indonesia
        'Thai': 'Thai',  # Thailand
        'MS': 'MS',  # Malaysia
        'China': 'China'
    }
    return mapping.get(str(supplier_class), str(supplier_class))

def load_data():
    """Load all data files and prepare dataframes"""
    print("Loading data files...")
    
    # Load part master data (701.CSV)
    parts_df = pd.read_csv(DATA_DIR / "701.CSV")
    # Strip spaces from product codes
    parts_df['MODEL'] = parts_df['MODEL/COLOR CODE'].str.strip()
    
    # Load orders (201-S.csv)
    orders_df = pd.read_csv(DATA_DIR / "201-S.csv")
    orders_df['SHIP DATE'] = pd.to_datetime(orders_df['SHIP DATE'], errors='coerce')
    
    # Load inventory / stock on hand (502-N.csv)
    inventory_df = pd.read_csv(DATA_DIR / "502-N.csv")
    inventory_df['PRODUCT_CLEAN'] = inventory_df['COMM PRODUCT'].str.strip()
    
    # Load MPMT CHECK.xlsx — primary devanning data source
    mpmt_xl = pd.ExcelFile(DATA_DIR / "MPMT CHECK.xlsx")
    mpmt_schedule = mpmt_xl.parse('Schedule')
    mpmt_schedule['DATE'] = pd.to_datetime(mpmt_schedule['DATE'], errors='coerce')
    
    # The container column header changes (sometimes 'CONTIANER', 'CONTAINER', or an actual container ID)
    # It is always the second column (index 1)
    container_col = mpmt_schedule.columns[1]
    mpmt_schedule['CONTAINER'] = mpmt_schedule[container_col].astype(str).str.strip()
    
    mpmt_data = mpmt_xl.parse('data')
    mpmt_data['CONTAINER_CLEAN'] = mpmt_data['CONTAINER NUMBER'].str.strip()
    
    print(f"Loaded {len(parts_df)} parts, {len(orders_df)} orders, "
          f"{len(inventory_df)} inventory records (502), "
          f"{len(mpmt_schedule)} scheduled containers (MPMT)")
    
    return parts_df, orders_df, inventory_df, mpmt_schedule, mpmt_data

def prepare_outbound(orders_df, parts_df, business_days):
    """Prepare outbound orders section"""
    print("\nProcessing outbound orders...")
    
    # Filter orders for next 5 business days
    start_date = business_days[0]
    end_date = business_days[-1]
    
    filtered_orders = orders_df[
        (orders_df['SHIP DATE'].dt.date >= start_date) & 
        (orders_df['SHIP DATE'].dt.date <= end_date)
    ].copy()
    
    if filtered_orders.empty:
        print("No outbound orders for the next 5 business days.")
        return [], 0, 0, 0, business_days

    # Deduplicate identical order rows
    before = len(filtered_orders)
    filtered_orders = filtered_orders.drop_duplicates(
        subset=['PRODUCT NO.', 'CUSTOMER ORDER NO.', 'QUANTITY', 'SHIP DATE']
    )
    dupes = before - len(filtered_orders)
    if dupes > 0:
        print(f"  Removed {dupes} duplicate order rows")

    # Normalize product codes - remove dashes from orders to match parts master
    filtered_orders['PRODUCT_NORMALIZED'] = filtered_orders['PRODUCT NO.'].str.replace('-', '', regex=False)
    
    # Merge with parts master to get ITEM TYPE and PCS/BOX
    filtered_orders = filtered_orders.merge(
        parts_df[['MODEL', 'ITEM TYPE', 'PCS/BOX']],
        left_on='PRODUCT_NORMALIZED',
        right_on='MODEL',
        how='left'
    )
    
    # Calculate boxes
    filtered_orders['BOXES'] = filtered_orders.apply(
        lambda x: x['QUANTITY'] / x['PCS/BOX'] if pd.notna(x['PCS/BOX']) and x['PCS/BOX'] > 0 else 0,
        axis=1
    )
    
    # Separate by item type
    filtered_orders['MODULE_BOXES'] = filtered_orders.apply(
        lambda x: x['BOXES'] if x['ITEM TYPE'] == 'SO' else 0,
        axis=1
    )
    filtered_orders['UNIT_BOXES'] = filtered_orders.apply(
        lambda x: x['BOXES'] if x['ITEM TYPE'] == 'MX' else 0,
        axis=1
    )
    
    # Build supplier code -> name lookup
    sup_lookup = filtered_orders.dropna(subset=['SUPPLIER NAME']).drop_duplicates('SUPPLIER CODE')
    sup_name_map = dict(zip(sup_lookup['SUPPLIER CODE'], sup_lookup['SUPPLIER NAME'].str.strip()))
    
    # Create pivot table: suppliers as rows, dates as columns
    # Group by supplier and date
    pivot_data = filtered_orders.groupby(
        ['SUPPLIER CODE', filtered_orders['SHIP DATE'].dt.date]
    ).agg({
        'BOXES': 'sum',
        'MODULE_BOXES': 'sum',
        'UNIT_BOXES': 'sum'
    }).reset_index()
    
    pivot_data.columns = ['SUPPLIER', 'SHIP_DATE', 'TOTAL_BOXES', 'MODULE_BOXES', 'UNIT_BOXES']
    
    # Map supplier codes to names
    pivot_data['SUPPLIER'] = pivot_data['SUPPLIER'].map(lambda c: sup_name_map.get(c, c))
    
    # Create a proper pivot table structure with suppliers as rows
    suppliers = sorted(pivot_data['SUPPLIER'].unique())
    supplier_summary = []
    
    for supplier in suppliers:
        supplier_data = {'SUPPLIER': supplier}
        supplier_total = 0
        supplier_module = 0
        supplier_unit = 0
        
        for day in business_days:
            day_data = pivot_data[(pivot_data['SUPPLIER'] == supplier) & (pivot_data['SHIP_DATE'] == day)]
            if len(day_data) > 0:
                total = day_data['TOTAL_BOXES'].iloc[0]
                module = day_data['MODULE_BOXES'].iloc[0]
                unit = day_data['UNIT_BOXES'].iloc[0]
            else:
                total = module = unit = 0
            
            date_key = day.strftime('%Y-%m-%d')
            supplier_data[f'{date_key}_TOTAL'] = total
            supplier_data[f'{date_key}_MODULE'] = module
            supplier_data[f'{date_key}_UNIT'] = unit
            
            supplier_total += total
            supplier_module += module
            supplier_unit += unit
        
        supplier_data['GRAND_TOTAL'] = supplier_total
        supplier_data['GRAND_MODULE'] = supplier_module
        supplier_data['GRAND_UNIT'] = supplier_unit
        supplier_summary.append(supplier_data)
    
    # Calculate overall totals
    total_boxes = filtered_orders['BOXES'].sum()
    module_boxes = filtered_orders['MODULE_BOXES'].sum()
    unit_boxes = filtered_orders['UNIT_BOXES'].sum()
    
    print(f"Total boxes: {total_boxes:.0f}, Solid Modules: {module_boxes:.0f}, Mix Modules: {unit_boxes:.0f}")
    
    return supplier_summary, total_boxes, module_boxes, unit_boxes, business_days

def prepare_devanning(mpmt_schedule, mpmt_data, parts_df, business_days):
    """Prepare devanning plan from MPMT CHECK.xlsx.
    Schedule sheet provides warehouse arrival dates.
    data sheet (filtered by WAREHOUSE CLASS='SITE1') provides module details."""
    print("\nProcessing devanning plan...")
    
    # Filter schedule to report week
    start_date = pd.Timestamp(business_days[0])
    end_date = pd.Timestamp(business_days[-1])
    
    week_schedule = mpmt_schedule[
        (mpmt_schedule['DATE'] >= start_date) & 
        (mpmt_schedule['DATE'] <= end_date)
    ].copy()
    
    if len(week_schedule) == 0:
        print("No containers scheduled for devanning period")
        return []
    
    # Get Site 1-only module data from data sheet
    indiana_data = mpmt_data[mpmt_data['WAREHOUSE CLASS'].str.strip() == 'SITE1'].copy()
    
    # Normalize part numbers and join with parts master for ITEM TYPE
    indiana_data['PART_CLEAN'] = indiana_data['PART NUMBER IN ORDER OF MODEL'].str.strip()
    indiana_data = indiana_data.merge(
        parts_df[['MODEL', 'ITEM TYPE']],
        left_on='PART_CLEAN',
        right_on='MODEL',
        how='left'
    )
    
    # Map transport mode
    def map_transport(mode):
        if pd.isna(mode):
            return 'Sea'
        mode_map = {'S': 'Sea', 'A': 'Air', 'O': 'Other'}
        return mode_map.get(str(mode).strip(), 'Sea')
    
    indiana_data['Transport'] = indiana_data['TRANSPORT MODE'].apply(map_transport)
    
    # Classify modules: UNITLOAD CLASS P = Module(Solid), R = Unit(Mix)
    indiana_data['Is_Mix'] = (indiana_data['UNITLOAD CLASS'].str.strip() == 'R').astype(int)
    indiana_data['Is_Solid'] = (indiana_data['UNITLOAD CLASS'].str.strip() == 'P').astype(int)
    
    # Summarize per container from data sheet
    container_detail = indiana_data.groupby(['CONTAINER_CLEAN', 'Transport']).agg({
        'Is_Mix': 'sum',
        'Is_Solid': 'sum',
        'MODULE NUMBER': 'nunique'
    }).reset_index().rename(columns={
        'Is_Mix': 'Mix_Count',
        'Is_Solid': 'Solid_Count',
        'MODULE NUMBER': 'Total_Modules'
    })
    
    # Join schedule (dates) with data (module details)
    merged = week_schedule.merge(
        container_detail,
        left_on='CONTAINER',
        right_on='CONTAINER_CLEAN',
        how='left'
    )
    
    # Fill missing module data for containers without data sheet entries
    merged['Mix_Count'] = merged['Mix_Count'].fillna(0).astype(int)
    merged['Solid_Count'] = merged['Solid_Count'].fillna(0).astype(int)
    merged['Total_Modules'] = merged['Total_Modules'].fillna(0).astype(int)
    merged['Transport'] = merged['Transport'].fillna('Sea')
    has_data = merged['CONTAINER_CLEAN'].notna()
    
    # Group by date for daily sections
    daily_data = []
    for date in sorted(merged['DATE'].unique()):
        date_containers = merged[merged['DATE'] == date]
        
        containers_list = []
        for _, row in date_containers.iterrows():
            containers_list.append({
                'container': row['CONTAINER'],
                'country': 'SITE1',
                'transport': row['Transport'],
                'mix_count': int(row['Mix_Count']),
                'solid_count': int(row['Solid_Count']),
                'total_modules': int(row['Total_Modules']),
                'has_data': pd.notna(row['CONTAINER_CLEAN'])
            })
        
        daily_totals = {
            'mix': int(date_containers['Mix_Count'].sum()),
            'solid': int(date_containers['Solid_Count'].sum()),
            'total_modules': int(date_containers['Total_Modules'].sum()),
            'total_containers': len(date_containers),
            'containers_with_data': int(has_data[date_containers.index].sum())
        }
        
        daily_data.append({
            'date': date,
            'containers': containers_list,
            'daily_totals': daily_totals
        })
    
    total_containers = sum(d['daily_totals']['total_containers'] for d in daily_data)
    print(f"Found {total_containers} containers across {len(daily_data)} dates")
    
    return daily_data

def parse_planned_date(order_no):
    """Extract planned ship date from order number (after underscore)"""
    if pd.isna(order_no):
        return pd.NaT
    
    order_no = str(order_no)
    if '_' in order_no:
        date_str = order_no.split('_')[-1]
    else:
        # Fallback: try positions 8-16 if no underscore
        date_str = order_no[8:16] if len(order_no) >= 16 else ''
    
    # Extract only digits
    date_digits = ''.join(ch for ch in date_str if ch.isdigit())
    if len(date_digits) < 8:
        return pd.NaT
    
    return pd.to_datetime(date_digits[:8], format='%Y%m%d', errors='coerce')


def calculate_shortages(orders_df, inventory_df, mpmt_schedule, mpmt_data, parts_df, business_days):
    """
    Calculate shortages: demand (outbound orders) vs supply (stock on hand + inbound).
    Shortage = demand - (stock + inbound arriving this week). Only positive values shown.
    """
    print("\nCalculating shortages...")
    
    start_date = pd.Timestamp(business_days[0])
    end_date = pd.Timestamp(business_days[-1])
    
    # 1. Demand: orders for the report week, grouped by product
    filtered_orders = orders_df[
        (orders_df['SHIP DATE'] >= start_date) &
        (orders_df['SHIP DATE'] <= end_date)
    ].copy()
    filtered_orders['PRODUCT_CLEAN'] = filtered_orders['PRODUCT NO.'].str.replace('-', '', regex=False).str.strip()
    
    demand = filtered_orders.groupby('PRODUCT_CLEAN')['QUANTITY'].sum().reset_index()
    demand.columns = ['PRODUCT', 'DEMAND']
    
    if len(demand) == 0:
        print("No orders for shortage calculation")
        return pd.DataFrame()
    
    # 2. Stock on hand from 502.csv
    stock = inventory_df.groupby('PRODUCT_CLEAN')['QUANTITY'].sum().reset_index()
    stock.columns = ['PRODUCT', 'STOCK']
    
    # 3. Inbound: quantities arriving this week from MPMT containers (Site 1 only)
    week_containers = set(
        mpmt_schedule[
            (mpmt_schedule['DATE'] >= start_date) &
            (mpmt_schedule['DATE'] <= end_date)
        ]['CONTAINER']
    )
    indiana_data = mpmt_data[mpmt_data['WAREHOUSE CLASS'].str.strip() == 'SITE1']
    inbound = indiana_data[indiana_data['CONTAINER_CLEAN'].isin(week_containers)]
    
    if len(inbound) > 0:
        inbound_qty = inbound.groupby(
            inbound['PART NUMBER IN ORDER OF MODEL'].str.strip()
        )['QUANTITY'].sum().reset_index()
        inbound_qty.columns = ['PRODUCT', 'INBOUND']
    else:
        inbound_qty = pd.DataFrame(columns=['PRODUCT', 'INBOUND'])
    
    # 4. Combine: demand vs (stock + inbound)
    comparison = demand.merge(stock, on='PRODUCT', how='left')
    comparison = comparison.merge(inbound_qty, on='PRODUCT', how='left')
    
    # Fix pandas future warning by inferring objects after fillna
    pd.set_option('future.no_silent_downcasting', True)
    comparison['STOCK'] = comparison['STOCK'].fillna(0).infer_objects(copy=False).astype(int)
    comparison['INBOUND'] = comparison['INBOUND'].fillna(0).infer_objects(copy=False).astype(int)
    
    comparison['SUPPLY'] = comparison['STOCK'] + comparison['INBOUND']
    comparison['SHORTAGE'] = (comparison['DEMAND'] - comparison['SUPPLY']).clip(lower=0)
    
    # 5. Filter to products with shortages
    shortages = comparison[comparison['SHORTAGE'] > 0].copy()
    
    if len(shortages) == 0:
        print("No shortages detected")
        return pd.DataFrame()
    
    # Add PCS/BOX for module count
    shortages = shortages.merge(
        parts_df[['MODEL', 'PCS/BOX']],
        left_on='PRODUCT',
        right_on='MODEL',
        how='left'
    )
    shortages['SHORTAGE_MODULES'] = shortages.apply(
        lambda r: int(r['SHORTAGE'] / r['PCS/BOX']) if pd.notna(r['PCS/BOX']) and r['PCS/BOX'] > 0 else 0,
        axis=1
    )
    
    shortage_df = shortages[['PRODUCT', 'DEMAND', 'STOCK', 'INBOUND', 'SUPPLY', 'SHORTAGE', 'SHORTAGE_MODULES']].copy()
    shortage_df = shortage_df.sort_values('SHORTAGE', ascending=False).reset_index(drop=True)
    
    print(f"Found {len(shortage_df)} parts with shortages (demand > stock + inbound)")
    
    return shortage_df

def create_outbound_chart(outbound_summary, business_days):
    """Create Plotly bar chart for outbound orders by day"""
    fig = go.Figure()
    
    if not outbound_summary or not business_days:
        return fig
        
    dates = []
    daily_module = []
    daily_unit = []
    
    for day in business_days:
        date_key = day.strftime('%Y-%m-%d')
        dates.append(day.strftime('%b %d (%a)'))
        
        day_module = sum(s.get(f'{date_key}_MODULE', 0) for s in outbound_summary)
        day_unit = sum(s.get(f'{date_key}_UNIT', 0) for s in outbound_summary)
        
        daily_module.append(day_module)
        daily_unit.append(day_unit)
    
    hover_data = []
    for i in range(len(dates)):
        total = daily_module[i] + daily_unit[i]
        hover_data.append(
            f"Solid Modules: {daily_module[i]:,.0f} ({daily_module[i] / total * 100 if total > 0 else 0:.1f}%)<br>"
            f"Mix Modules: {daily_unit[i]:,.0f} ({daily_unit[i] / total * 100 if total > 0 else 0:.1f}%)<br>"
            f"Total: {total:,.0f}"
        )
    
    fig.add_trace(go.Bar(
        name='Solid Modules',
        x=dates,
        y=daily_module,
        marker_color='#4f46e5',
        marker_line_width=0,
        text=[f'{v:,.0f}' if v > 0 else '' for v in daily_module],
        textposition='inside',
        textfont=dict(color='white', size=12, family='Inter, sans-serif'),
        hovertemplate='%{customdata}<extra></extra>',
        customdata=hover_data
    ))
    
    fig.add_trace(go.Bar(
        name='Mix Modules',
        x=dates,
        y=daily_unit,
        marker_color='#38bdf8',
        marker_line_width=0,
        text=[f'{v:,.0f}' if v > 0 else '' for v in daily_unit],
        textposition='inside',
        textfont=dict(color='white', size=12, family='Inter, sans-serif'),
        hovertemplate='<extra></extra>'
    ))
    
    fig.update_layout(
        title=None,
        xaxis_title='',
        yaxis_title='Box Count',
        barmode='stack',
        hovermode='x',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#334155'),
        xaxis=dict(tickformat='%b %d (%a)', tickangle=0, showgrid=False, showline=False, tickfont=dict(color='#475569', size=13)),
        yaxis=dict(showgrid=True, gridcolor='#e2e8f0', showline=False, zeroline=False, tickfont=dict(color='#475569', size=13)),
        legend=dict(orientation='h', yanchor='bottom', y=1.05, xanchor='right', x=1, font=dict(size=14, color='#1e293b')),
        height=360,
        margin=dict(l=50, r=20, t=40, b=40)
    )
    
    return fig

def create_devanning_chart(devanning_data, business_days):
    """Create Plotly bar chart for devanning plan by day"""
    fig = go.Figure()
    
    if not business_days:
        return fig
    
    # Build lookup from devanning_data by date
    dev_lookup = {}
    for daily in (devanning_data or []):
        dev_lookup[daily['date']] = daily['daily_totals']
    
    dates = []
    mix_counts = []
    solid_counts = []
    container_counts = []
    
    for day in business_days:
        dates.append(day.strftime('%b %d (%a)'))
        totals = dev_lookup.get(pd.Timestamp(day), {})
        mix_counts.append(totals.get('mix', 0))
        solid_counts.append(totals.get('solid', 0))
        container_counts.append(totals.get('total_containers', 0))
    
    hover_data = []
    for i in range(len(dates)):
        total_modules = mix_counts[i] + solid_counts[i]
        hover_data.append(
            f"Containers: {container_counts[i]}<br>"
            f"Mix Modules: {mix_counts[i]:,}<br>"
            f"Solid Modules: {solid_counts[i]:,}<br>"
            f"Total: {total_modules:,}"
        )
    
    fig.add_trace(go.Bar(
        name='Mix Modules',
        x=dates,
        y=mix_counts,
        marker_color='#f59e0b',
        marker_line_width=0,
        text=[f'{v:,}' if v > 0 else '' for v in mix_counts],
        textposition='inside',
        textfont=dict(color='white', size=12, family='Inter, sans-serif'),
        hovertemplate='%{customdata}<extra></extra>',
        customdata=hover_data
    ))
    
    fig.add_trace(go.Bar(
        name='Solid Modules',
        x=dates,
        y=solid_counts,
        marker_color='#10b981',
        marker_line_width=0,
        text=[f'{v:,}' if v > 0 else '' for v in solid_counts],
        textposition='inside',
        textfont=dict(color='white', size=12, family='Inter, sans-serif'),
        hovertemplate='<extra></extra>'
    ))
    
    fig.update_layout(
        title=None,
        xaxis_title='',
        yaxis_title='Module Count',
        barmode='stack',
        hovermode='x',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, sans-serif', size=14, color='#334155'),
        xaxis=dict(tickformat='%b %d (%a)', tickangle=0, showgrid=False, showline=False, tickfont=dict(color='#475569', size=13)),
        yaxis=dict(showgrid=True, gridcolor='#e2e8f0', showline=False, zeroline=False, tickfont=dict(color='#475569', size=13)),
        legend=dict(orientation='h', yanchor='bottom', y=1.05, xanchor='right', x=1, font=dict(size=14, color='#1e293b')),
        height=360,
        margin=dict(l=50, r=20, t=40, b=40)
    )
    
    return fig

def _report_css():
    """Return the bold, professional CSS for the report."""
    return """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

:root {
    --bg: #eaecf0;
    --card: #ffffff;
    --text: #0f172a;
    --text-2: #334155;
    --text-3: #64748b;
    --border: #d1d9e6;
    --border-light: #e8edf3;
    --out-color: #4338ca;
    --dev-color: #0d9488;
    --short-color: #dc2626;
}

* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Inter', system-ui, sans-serif; background: var(--bg); color: var(--text); min-height: 100vh; line-height: 1.5; }

/* ── Top Navigation ── */
.top-nav {
    background: linear-gradient(135deg, #0d1b2a 0%, #162d4a 60%, #0d1b2a 100%);
    padding: 0 44px;
    display: flex; justify-content: space-between; align-items: center;
    height: 74px;
    position: sticky; top: 0; z-index: 100;
    box-shadow: 0 4px 24px rgba(0,0,0,0.35);
}
.nav-brand { display: flex; align-items: center; gap: 18px; }
.nav-logo {
    width: 44px; height: 44px;
    background: linear-gradient(135deg, #e53e3e 0%, #c53030 100%);
    border-radius: 11px;
    display: flex; align-items: center; justify-content: center;
    color: white; font-weight: 900; font-size: 19px; letter-spacing: -1px;
    box-shadow: 0 4px 14px rgba(229,62,62,0.45);
    flex-shrink: 0;
}
.nav-text { display: flex; flex-direction: column; gap: 2px; }
.nav-title { color: white; font-size: 1.15rem; font-weight: 800; letter-spacing: 0.01em; display: flex; align-items: center; gap: 10px; }
.nav-title span { color: rgba(255,255,255,0.35); font-weight: 400; }
.nav-subtitle { color: rgba(255,255,255,0.45); font-size: 0.73rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.1em; }
.nav-date {
    background: rgba(255,255,255,0.11);
    border: 1px solid rgba(255,255,255,0.22);
    color: white;
    padding: 10px 24px; border-radius: 999px;
    font-size: 0.88rem; font-weight: 700; letter-spacing: 0.03em;
}

/* ── Main ── */
.main-content { max-width: 1520px; margin: 0 auto; padding: 36px 44px; }

/* ── KPI Grid ── */
.kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 20px; margin-bottom: 52px; }
.kpi-card {
    border-radius: 16px; padding: 26px 26px 22px;
    position: relative; overflow: hidden;
    color: white; min-height: 156px;
    display: flex; flex-direction: column;
    box-shadow: 0 8px 28px rgba(0,0,0,0.14);
}
.kpi-icon {
    position: absolute; right: 18px; top: 14px;
    opacity: 0.13; width: 58px; height: 58px;
}
.kpi-c1 { background: linear-gradient(135deg, #4338ca 0%, #7c3aed 100%); }
.kpi-c2 { background: linear-gradient(135deg, #0369a1 0%, #0ea5e9 100%); }
.kpi-c3 { background: linear-gradient(135deg, #b45309 0%, #f59e0b 100%); }
.kpi-c4 { background: linear-gradient(135deg, #047857 0%, #10b981 100%); }
.kpi-c5.ok { background: linear-gradient(135deg, #15803d 0%, #22c55e 100%); }
.kpi-c5.bad { background: linear-gradient(135deg, #b91c1c 0%, #ef4444 100%); }
.kpi-l { font-size: 0.72rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.1em; opacity: 0.75; margin-bottom: 10px; }
.kpi-v { font-size: 2.9rem; font-weight: 900; line-height: 1; letter-spacing: -0.03em; margin-top: auto; }
.kpi-sub { font-size: 0.82rem; font-weight: 500; opacity: 0.65; margin-top: 9px; }

/* ── Sections ── */
.section { margin-bottom: 52px; }
.section-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 22px; }
.section-title {
    font-size: 1.55rem; font-weight: 900; letter-spacing: -0.02em;
    display: flex; align-items: center; gap: 16px;
    padding-left: 18px; border-left: 5px solid currentColor;
    line-height: 1.2;
}
.section-title.out-title  { color: var(--out-color); }
.section-title.dev-title  { color: var(--dev-color); }
.section-title.short-title { color: var(--short-color); }
.section-badge {
    background: white; border: 1.5px solid var(--border);
    padding: 4px 14px; border-radius: 999px;
    font-size: 0.73rem; font-weight: 800; color: var(--text-3);
    letter-spacing: 0.07em; text-transform: uppercase;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
}

/* ── Panels ── */
.panel {
    background: var(--card); border-radius: 16px;
    box-shadow: 0 4px 18px rgba(0,0,0,0.07);
    border: 1px solid var(--border-light);
    overflow: hidden;
}
.panel.out-panel   { border-top: 4px solid var(--out-color); }
.panel.dev-panel   { border-top: 4px solid var(--dev-color); }
.panel.short-panel { border-top: 4px solid var(--short-color); }
.chart-box { padding: 28px 32px 16px; border-bottom: 1px solid var(--border-light); background: #f9fafb; }

/* ── Tables ── */
.table-wrap { width: 100%; overflow-x: auto; }
table.data-table { width: 100%; border-collapse: collapse; }
.data-table th {
    background: #f2f4f8; color: var(--text-2);
    font-size: 0.76rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.09em;
    padding: 16px 24px; text-align: left;
    border-bottom: 2px solid var(--border); white-space: nowrap;
}
.data-table td {
    padding: 14px 24px; font-size: 0.97rem; color: var(--text);
    border-bottom: 1px solid var(--border-light); white-space: nowrap;
}
.data-table tbody tr:hover td { background: #f8fafc; }
.data-table tbody tr:last-child td { border-bottom: none; }
.data-table .r { text-align: right; }
.data-table .c { text-align: center; }
.data-table .fw-bold { font-weight: 700; }
.data-table .text-muted { color: var(--text-3); }

/* Date separator rows */
.row-date td {
    background: #f1f5f9 !important;
    font-weight: 800; color: var(--text-2);
    padding-top: 20px; padding-bottom: 16px;
    font-size: 1rem; letter-spacing: -0.01em;
    border-top: 2px solid var(--border); border-bottom: 1px solid var(--border);
}
.row-date:first-child td { border-top: none; }
.date-badge {
    display: inline-block; background: white;
    border: 1.5px solid var(--border); padding: 3px 10px;
    border-radius: 999px; font-size: 0.72rem; font-weight: 700;
    color: var(--text-3); margin-left: 12px; vertical-align: middle;
}
.row-subtotal td {
    background: #f7f9fc !important; font-weight: 700; color: var(--text-2);
    border-top: 1px dashed var(--border);
    padding-top: 14px; padding-bottom: 14px; font-size: 0.94rem;
}
.row-grand td {
    background: #1e293b !important; color: white !important;
    font-weight: 800; padding: 20px 24px;
    font-size: 1rem; letter-spacing: 0.01em; border: none;
}

/* ── Badges ── */
.badge {
    display: inline-flex; align-items: center; justify-content: center;
    padding: 5px 14px; border-radius: 6px;
    font-size: 0.77rem; font-weight: 800; letter-spacing: 0.05em; text-transform: uppercase;
}
.bg-sea { background: #dbeafe; color: #1d4ed8; }
.bg-air { background: #fef3c7; color: #b45309; }
.text-danger { color: var(--short-color) !important; font-weight: 800; }

/* ── Empty / Success States ── */
.empty-state {
    padding: 56px; text-align: center; color: var(--text-3);
    background: var(--card); border-radius: 16px;
    border: 2px dashed var(--border); font-weight: 600; font-size: 1rem;
}
.empty-success {
    padding: 22px 28px; display: flex; align-items: center; gap: 16px;
    background: #f0fdf4; border: 2px solid #86efac;
    border-radius: 14px; color: #166534; font-weight: 700; font-size: 1rem;
}
.empty-success svg { width: 28px; height: 28px; flex-shrink: 0; }

/* ── Footer ── */
.footer {
    text-align: center; margin-top: 64px; padding: 30px 0 40px;
    border-top: 1px solid var(--border);
    color: var(--text-3); font-size: 0.84rem; font-weight: 500;
}

@media (max-width: 1200px) { .kpi-grid { grid-template-columns: repeat(3, 1fr); } }
@media (max-width: 768px)  { .kpi-grid { grid-template-columns: 1fr 1fr; } .main-content { padding: 24px 20px; } .top-nav { padding: 0 20px; } }

@media print {
    body { background: white; }
    .top-nav { position: static; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .kpi-card, .row-grand td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .main-content { padding: 0; max-width: 100%; }
    .panel { box-shadow: none; border: 1px solid #ccc; page-break-inside: avoid; }
    .chart-box { display: none; }
}
"""


def generate_html_report(outbound_summary, outbound_totals, devanning_data,
                        shortage_df, business_days, report_days):
    """Generate HTML report with modern dashboard styling and Plotly charts"""
    print("\nGenerating HTML report...")

    total_boxes, module_boxes, unit_boxes = outbound_totals

    # Create Plotly charts
    outbound_fig = create_outbound_chart(outbound_summary, business_days)
    outbound_chart_html = outbound_fig.to_html(
        include_plotlyjs='cdn', div_id='outbound-chart',
        config={'displayModeBar': False, 'responsive': True}
    )

    devanning_fig = create_devanning_chart(devanning_data, business_days)
    devanning_chart_html = devanning_fig.to_html(
        include_plotlyjs=False, div_id='devanning-chart',
        config={'displayModeBar': False, 'responsive': True}
    )

    # Calculate summary stats
    date_range = f"{business_days[0].strftime('%b %d')} - {business_days[-1].strftime('%b %d, %Y')}"
    num_suppliers = len(outbound_summary)
    total_containers = sum(len(d['containers']) for d in devanning_data) if devanning_data else 0
    total_devanning_days = len(devanning_data) if devanning_data else 0
    shortage_count = len(shortage_df) if len(shortage_df) > 0 else 0
    total_dev_modules = sum(
        d['daily_totals'].get('total_modules', d['daily_totals']['mix'] + d['daily_totals']['solid'])
        for d in devanning_data
    ) if devanning_data else 0

    css = _report_css()
    L = []  # accumulate HTML lines

    # ── Head ──
    L.append('<!doctype html>')
    L.append('<html lang="en"><head>')
    L.append('<meta charset="utf-8">')
    L.append('<meta name="viewport" content="width=device-width,initial-scale=1">')
    L.append(f'<title>Next Week Plan & Forecast - {date_range}</title>')
    L.append(f'<style>{css}</style>')
    L.append('</head><body>')

    # ── Top Navigation ──
    L.append('<nav class="top-nav">')
    L.append('<div class="nav-brand">')
    L.append('<div class="nav-logo">N</div>')
    L.append('<div class="nav-text">')
    L.append('<div class="nav-title">Example Logistics <span>|</span> Plan &amp; Forecast</div>')
    L.append('<div class="nav-subtitle">Site 1 Warehouse &nbsp;·&nbsp; Weekly Report</div>')
    L.append('</div>')
    L.append('</div>')
    L.append(f'<div class="nav-date">{date_range}</div>')
    L.append('</nav>')

    L.append('<div class="main-content">')

    # ── KPI Grid ──
    short_kpi_cls = "kpi-c5 bad" if shortage_count > 0 else "kpi-c5 ok"
    # SVG icons (filled, white)
    _icon_box      = '<svg class="kpi-icon" viewBox="0 0 24 24" fill="white"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96" stroke="rgba(0,0,0,.15)" stroke-width="1" fill="none"/><line x1="12" y1="22.08" x2="12" y2="12" stroke="rgba(0,0,0,.15)" stroke-width="1"/></svg>'
    _icon_layers   = '<svg class="kpi-icon" viewBox="0 0 24 24" fill="white"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17" fill="none" stroke="white" stroke-width="2.5"/><polyline points="2 12 12 17 22 12" fill="none" stroke="white" stroke-width="2.5"/></svg>'
    _icon_ship     = '<svg class="kpi-icon" viewBox="0 0 24 24" fill="white"><path d="M2 20a2.4 2.4 0 0 0 2 1 2.4 2.4 0 0 0 2-1 2.4 2.4 0 0 1 2-1 2.4 2.4 0 0 1 2 1 2.4 2.4 0 0 0 2 1 2.4 2.4 0 0 0 2-1 2.4 2.4 0 0 1 2-1 2.4 2.4 0 0 1 2 1"/><path d="M4 18l-1-5h18l-2 5"/><path d="M11 13V6a1 1 0 0 1 1-1h3"/><path d="M15 5l2 8"/></svg>'
    _icon_inbox    = '<svg class="kpi-icon" viewBox="0 0 24 24" fill="white"><polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 17.38 4H6.62a2 2 0 0 0-1.77 1.11z"/></svg>'
    _icon_alert    = '<svg class="kpi-icon" viewBox="0 0 24 24" fill="white"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13" stroke="rgba(0,0,0,.3)" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="17" x2="12.01" y2="17" stroke="rgba(0,0,0,.3)" stroke-width="2" stroke-linecap="round"/></svg>'
    _icon_check    = '<svg class="kpi-icon" viewBox="0 0 24 24" fill="white"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01" stroke="rgba(0,0,0,.3)" stroke-width="2" stroke-linecap="round" fill="none"/></svg>'

    L.append('<div class="kpi-grid">')

    L.append('<div class="kpi-card kpi-c1">')
    L.append(_icon_box)
    L.append('<div class="kpi-l">Outbound Boxes</div>')
    L.append(f'<div class="kpi-v">{total_boxes:,.0f}</div>')
    L.append(f'<div class="kpi-sub">From {num_suppliers} suppliers</div>')
    L.append('</div>')

    L.append('<div class="kpi-card kpi-c2">')
    L.append(_icon_layers)
    L.append('<div class="kpi-l">Solid / Mix Modules</div>')
    L.append(f'<div class="kpi-v">{module_boxes:,.0f} <span style="opacity:.45;font-weight:400">/</span> {unit_boxes:,.0f}</div>')
    L.append('<div class="kpi-sub">Module type breakdown</div>')
    L.append('</div>')

    L.append('<div class="kpi-card kpi-c3">')
    L.append(_icon_ship)
    L.append('<div class="kpi-l">Containers In</div>')
    L.append(f'<div class="kpi-v">{total_containers}</div>')
    L.append(f'<div class="kpi-sub">Across {total_devanning_days} day{"s" if total_devanning_days != 1 else ""}</div>')
    L.append('</div>')

    L.append('<div class="kpi-card kpi-c4">')
    L.append(_icon_inbox)
    L.append('<div class="kpi-l">Modules Inbound</div>')
    L.append(f'<div class="kpi-v">{total_dev_modules:,.0f}</div>')
    L.append('<div class="kpi-sub">Total inbound volume</div>')
    L.append('</div>')

    _shortage_icon = _icon_alert if shortage_count > 0 else _icon_check
    L.append(f'<div class="kpi-card {short_kpi_cls}">')
    L.append(_shortage_icon)
    L.append('<div class="kpi-l">Parts Short</div>')
    L.append(f'<div class="kpi-v">{shortage_count}</div>')
    L.append(f'<div class="kpi-sub">{"Part numbers at risk" if shortage_count > 0 else "No shortages detected"}</div>')
    L.append('</div>')

    L.append('</div>')

    # ── OUTBOUND SECTION ──
    L.append('<div class="section">')
    L.append('<div class="section-header">')
    L.append('<div class="section-title out-title">Outbound Plan <span class="section-badge">Shipping</span></div>')
    L.append('</div>')
    
    L.append('<div class="panel out-panel">')
    L.append(f'<div class="chart-box">{outbound_chart_html}</div>')
    L.append('<div class="table-wrap"><table class="data-table">')
    L.append('<thead><tr>')
    L.append('<th style="min-width:160px">Ship Date</th><th>Supplier</th><th class="r">Solid Modules</th><th class="r">Mix Modules</th><th class="r">Total Boxes</th>')
    L.append('</tr></thead><tbody>')

    for day in report_days:
        dk = day.strftime('%Y-%m-%d')
        day_str = day.strftime('%A, %b %d')
        day_suppliers = []
        for sd in outbound_summary:
            t = sd.get(f'{dk}_TOTAL', 0)
            s = sd.get(f'{dk}_MODULE', 0)
            m = sd.get(f'{dk}_UNIT', 0)
            if t > 0:
                day_suppliers.append((sd['SUPPLIER'], s, m, t))
                
        day_total = sum(sd.get(f'{dk}_TOTAL', 0) for sd in outbound_summary)
        day_solid = sum(sd.get(f'{dk}_MODULE', 0) for sd in outbound_summary)
        day_mix = sum(sd.get(f'{dk}_UNIT', 0) for sd in outbound_summary)
        n_sup = len(day_suppliers)
        cnt_label = f'{n_sup} supplier{"s" if n_sup != 1 else ""}' if n_sup > 0 else 'no orders'
        
        L.append(f'<tr class="row-date"><td colspan="5">{day_str}<span class="date-badge">{cnt_label}</span></td></tr>')
        for sup_name, sup_s, sup_m, sup_t in day_suppliers:
            L.append(f'<tr><td></td><td class="fw-bold">{sup_name}</td><td class="r">{sup_s:,.0f}</td><td class="r">{sup_m:,.0f}</td><td class="r fw-bold">{sup_t:,.0f}</td></tr>')
        if n_sup > 0:
            L.append(f'<tr class="row-subtotal"><td colspan="2" class="r">Subtotal</td><td class="r">{day_solid:,.0f}</td><td class="r">{day_mix:,.0f}</td><td class="r">{day_total:,.0f}</td></tr>')

    L.append('</tbody><tfoot><tr class="row-grand">')
    L.append(f'<td colspan="2" class="r">TOTAL OUTBOUND</td><td class="r">{module_boxes:,.0f}</td><td class="r">{unit_boxes:,.0f}</td><td class="r">{total_boxes:,.0f}</td>')
    L.append('</tr></tfoot>')
    L.append('</table></div></div></div>')

    # ── INBOUND SECTION ──
    L.append('<div class="section">')
    L.append('<div class="section-header">')
    L.append('<div class="section-title dev-title">Devanning Plan <span class="section-badge">Receiving</span></div>')
    L.append('</div>')

    if devanning_data:
        L.append('<div class="panel dev-panel">')
        L.append(f'<div class="chart-box">{devanning_chart_html}</div>')
        L.append('<div class="table-wrap"><table class="data-table">')
        L.append('<thead><tr>')
        L.append('<th style="min-width:160px">Arrival Date</th><th>Container</th><th class="c">Transport</th><th class="r">Mix Modules</th><th class="r">Solid Modules</th><th class="r">Total Modules</th>')
        L.append('</tr></thead><tbody>')

        grand_mix = grand_solid = grand_total_modules = 0

        for daily_section in devanning_data:
            date_str = daily_section['date'].strftime('%A, %b %d')
            num_c = len(daily_section['containers'])
            L.append(f'<tr class="row-date"><td colspan="6">{date_str}<span class="date-badge">{num_c} container{"s" if num_c != 1 else ""}</span></td></tr>')

            for c in daily_section['containers']:
                mx = c['mix_count']
                sl = c['solid_count']
                tm = c.get('total_modules', mx + sl)
                tp = c['transport']
                bp_cls = 'bg-air' if tp == 'Air' else 'bg-sea'
                L.append(f'<tr><td></td><td class="fw-bold">{c["container"]}</td><td class="c"><span class="badge {bp_cls}">{tp}</span></td><td class="r">{mx:,}</td><td class="r">{sl:,}</td><td class="r fw-bold">{tm:,}</td></tr>')

            dt = daily_section['daily_totals']
            dmx = dt['mix']
            dsl = dt['solid']
            dtm = dt.get('total_modules', dmx + dsl)
            grand_mix += dmx
            grand_solid += dsl
            grand_total_modules += dtm
            L.append(f'<tr class="row-subtotal"><td colspan="3" class="r">Subtotal</td><td class="r">{dmx:,}</td><td class="r">{dsl:,}</td><td class="r">{dtm:,}</td></tr>')

        L.append('</tbody><tfoot><tr class="row-grand">')
        L.append(f'<td colspan="3" class="r">TOTAL INBOUND</td><td class="r">{grand_mix:,}</td><td class="r">{grand_solid:,}</td><td class="r">{grand_total_modules:,}</td>')
        L.append('</tr></tfoot></table></div></div>')
    else:
        L.append('<div class="empty-state">No containers scheduled during this period.</div>')

    L.append('</div>')

    # ── SHORTAGE SECTION ──
    L.append('<div class="section">')
    L.append('<div class="section-header">')
    L.append('<div class="section-title short-title">Shortage Estimate <span class="section-badge">Inventory</span></div>')
    L.append('</div>')

    if len(shortage_df) > 0:
        L.append('<div class="panel short-panel"><div class="table-wrap"><table class="data-table">')
        L.append('<thead><tr><th>Part Number</th><th class="r">Demand</th><th class="r">Stock</th><th class="r">Inbound</th><th class="r">Supply</th><th class="r text-danger">Shortage</th><th class="r text-danger">Shortage Modules</th></tr></thead>')
        L.append('<tbody>')
        for _, row in shortage_df.iterrows():
            L.append(f'<tr><td class="fw-bold">{row["PRODUCT"]}</td><td class="r text-muted">{row["DEMAND"]:,.0f}</td><td class="r text-muted">{row["STOCK"]:,.0f}</td><td class="r text-muted">{row["INBOUND"]:,.0f}</td><td class="r text-muted">{row["SUPPLY"]:,.0f}</td><td class="r text-danger">{row["SHORTAGE"]:,.0f}</td><td class="r text-danger">{row["SHORTAGE_MODULES"]:,.0f}</td></tr>')
        tot_s = shortage_df['SHORTAGE'].sum()
        tot_sm = shortage_df['SHORTAGE_MODULES'].sum()
        L.append(f'</tbody><tfoot><tr class="row-grand"><td colspan="5" class="r">TOTAL SHORTAGES</td><td class="r text-danger">{tot_s:,.0f}</td><td class="r text-danger">{tot_sm:,.0f}</td></tr></tfoot>')
        L.append('</table></div></div>')
    else:
        L.append('<div class="empty-success">')
        L.append('<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>')
        L.append('<span>No shortages detected for this period. All parts are sufficiently stocked.</span>')
        L.append('</div>')

    L.append('</div>')

    # ── Footer ──
    timestamp = datetime.now().strftime('%B %d, %Y at %I:%M %p')
    L.append(f'<div class="footer"><p>Generated on {timestamp}</p><p style="margin-top: 4px;">Dashboard created by Viktor Berg</p></div>')
    L.append('</div></body></html>')

    return '\n'.join(L)

def main():
    """Main execution function"""
    print("=" * 60)
    print("DS Weekly Order Report Generator")
    print("=" * 60)
    
    # Check for custom start date argument
    start_date = None
    if len(sys.argv) > 1:
        try:
            start_date = datetime.strptime(sys.argv[1], '%Y-%m-%d').date()
            print(f"\nUsing custom start date: {start_date}")
        except ValueError:
            print(f"\nInvalid date format. Use YYYY-MM-DD. Using default (next week).")
    
    # Get next 5 business days
    business_days = get_next_business_days(start_date=start_date)
    print(f"\nReport period: {business_days[0]} to {business_days[-1]}")
    print(f"Business days: {[d.strftime('%m/%d %a') for d in business_days]}")
    
    # Load data
    parts_df, orders_df, inventory_df, mpmt_schedule, mpmt_data = load_data()
    
    # Process each section
    outbound_summary, total_boxes, module_boxes, unit_boxes, _ = prepare_outbound(
        orders_df, parts_df, business_days
    )
    report_days = business_days  # Use business_days as report_days
    
    # Process devanning plan (from MPMT CHECK.xlsx)
    devanning_data = prepare_devanning(
        mpmt_schedule, mpmt_data, parts_df, business_days
    )
    
    shortage_df = calculate_shortages(
        orders_df, inventory_df, mpmt_schedule, mpmt_data, parts_df, business_days
    )
    
    # Generate HTML report
    html_content = generate_html_report(
        outbound_summary,
        (total_boxes, module_boxes, unit_boxes),
        devanning_data,
        shortage_df,
        business_days,
        report_days
    )
    
    # Save report
    date_range = f"{business_days[0].strftime('%m.%d.%Y')} - {business_days[-1].strftime('%m.%d.%Y')}"
    output_filename = f"Next Week Plan and Forecast ({date_range}).html"
    output_path = OUTPUT_DIR / output_filename
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"\n{'=' * 60}")
    print(f"Report generated successfully!")
    print(f"Output file: {output_path}")
    print(f"{'=' * 60}")
    
    # ── Email Summary (copy-paste ready) ──
    period_str = f"{business_days[0].strftime('%b %d')} - {business_days[-1].strftime('%b %d, %Y')}"
    total_containers = sum(len(d['containers']) for d in devanning_data) if devanning_data else 0
    total_devanning_days = len(devanning_data) if devanning_data else 0
    shortage_count = len(shortage_df) if len(shortage_df) > 0 else 0
    
    # Build daily outbound breakdown
    daily_lines = []
    for day in business_days:
        date_key = day.strftime('%Y-%m-%d')
        day_total = sum(s.get(f'{date_key}_TOTAL', 0) for s in outbound_summary)
        day_label = day.strftime('%a %m/%d')
        daily_lines.append(f"  {day_label}: {day_total:,.0f} boxes")
    
    # Build devanning breakdown
    devanning_lines = []
    grand_mix = 0
    grand_solid = 0
    for d in devanning_data:
        date_label = d['date'].strftime('%a %m/%d')
        nc = len(d['containers'])
        dm = d['daily_totals']['mix']
        ds = d['daily_totals']['solid']
        grand_mix += dm
        grand_solid += ds
        devanning_lines.append(f"  {date_label}: {nc} container{'s' if nc != 1 else ''} ({ds} solid, {dm} mix modules)")
    
    print(f"\n{'─' * 60}")
    print(f"Next Week Plan & Forecast ({period_str})")
    print(f"{'─' * 60}\n")
    print(f"Hello Supervisors,\n")
    print(f"Attached is the next week's plan & forecast for {period_str}. Below is a quick summary:\n")
    print(f"Outbound: {total_boxes:,.0f} boxes ({module_boxes:,.0f} solid, {unit_boxes:,.0f} mix modules)")
    for line in daily_lines:
        print(line)
    print()
    print(f"Devanning: {total_containers} containers across {total_devanning_days} days ({grand_solid} solid, {grand_mix} mix modules)")
    for line in devanning_lines:
        print(line)
    print()
    if shortage_count > 0:
        print(f"Shortages: {shortage_count} parts with insufficient stock — see report for details.")
    else:
        print(f"Shortages: None — all parts are sufficiently stocked.")
    print(f"\nHave a great weekend and see you next week!\n")
    print(f"\n{'─' * 60}")
    
    return output_path

if __name__ == "__main__":
    main()
