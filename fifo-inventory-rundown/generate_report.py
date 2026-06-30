"""Weekly FIFO Inventory HTML report generator.

Reads three CSVs (current inventory, prior inventory from archive, shipping plan),
computes KPIs and chart series, emits a single self-contained HTML file.

Usage:
    python generate_report.py
    python generate_report.py --open
    python generate_report.py --inventory path/to/curr.csv --prior path/to/prev.csv
"""

from __future__ import annotations

import argparse
import html
import re
import sys
import webbrowser
from datetime import date, datetime, timedelta
from pathlib import Path
from string import Template

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio


# ─── Paths ───────────────────────────────────────────────────────────────────

PROJECT_ROOT = Path(__file__).resolve().parent
OUT_DIR      = PROJECT_ROOT

# Tab order matches the screenshot mockup: IP, DS, SITE1.
# loc is the WMS LOCATION token (SITE2 / SITE3 / SITE1) — shown as the tab sub-label.
WAREHOUSES = [
    {"code": "IP",  "loc": "SITE2", "folder": "IP data", "label": "IP FIFO Inventory Rundown"},
    {"code": "DS",  "loc": "SITE3", "folder": "DS data", "label": "DS FIFO Inventory Rundown"},
    {"code": "SITE1", "loc": "SITE1", "folder": "IN data", "label": "SITE1 FIFO Inventory Rundown"},
]

CURR_INV_NAME = "current 502.csv"
SHIP_NAME     = "202.csv"
BACKLOG_GLOB  = "201P.*"     # case-insensitive; IP exports as 201P.CSV, DS/SITE1 as 201P.csv
PRICE_FILE    = PROJECT_ROOT / "Inventory Value Revision" / "Parts Inventory Info and Price.xlsx"


# ─── Theme ───────────────────────────────────────────────────────────────────

COLOR_BG       = "#1f232c"   # page (slate gray)
COLOR_PANEL    = "#2a2f3a"   # cards
COLOR_PANEL2   = "#333845"   # alt-row / lot-brief tint
COLOR_HEAD     = "#363b48"   # table thead
COLOR_BORDER   = "#444b5c"   # default border
COLOR_BORDER2  = "#2e333f"   # row dividers (subtle)
COLOR_TEXT     = "#eef2fa"   # ink
COLOR_DIM      = "#9aa6bf"   # labels
COLOR_ACCENT   = "#ff8a4c"   # orange — primary brand accent
COLOR_CYAN     = "#60a5fa"   # secondary numeric accent (links, doc numbers)
COLOR_GOOD     = "#4ade80"
COLOR_BAD      = "#f87171"
COLOR_WARN     = "#fb923c"

COLOR_LINE_CURR   = "#cbd5e1"
COLOR_LINE_PREV   = "#64748b"
COLOR_BAR_CONSUME = "#dc4a4a"
COLOR_BAR_ARRIVE  = "#4ade80"

AGE_RED    = "#f87171"   # ≥90
AGE_ORANGE = "#fb923c"   # 60–90
AGE_YELLOW = "#fbbf24"   # 30–60


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _esc(value) -> str:
    """HTML-escape a value for safe interpolation into both text and attributes."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return html.escape(str(value), quote=True)


# ─── Loaders ─────────────────────────────────────────────────────────────────

# Snapshot files inside each warehouse folder are the WMS dump format
# MODULE_LOC_<LOC>_<COMPANY>_YYYYMMDD_HHMMSS.csv (e.g. MODULE_LOC_SITE1_CUST1_…,
# MODULE_LOC_SITE3_CUST1_…, MODULE_LOC_SITE2_CUST1_…).
# YYYY-MM-DD[_HHMMSS].csv is also accepted as a friendlier manual fallback.
PA_FILENAME_RE = re.compile(r"MODULE_LOC_[A-Z0-9]+_[A-Z0-9]+_(\d{8})_(\d{6})", re.IGNORECASE)
SNAPSHOT_NAME_FORMATS = ("%Y-%m-%d_%H%M%S", "%Y-%m-%d_%H-%M-%S", "%Y-%m-%d")


def parse_snapshot_from_filename(path: Path) -> datetime | None:
    m = PA_FILENAME_RE.search(path.name)
    if m: 
        try:
            return datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")
        except ValueError:
            pass
    for fmt in SNAPSHOT_NAME_FORMATS:
        try:
            return datetime.strptime(path.stem, fmt)
        except ValueError:
            continue
    return None


def snapshot_datetime(path: Path) -> datetime:
    """Filename timestamp if it parses (PA or ISO format), else file mtime."""
    return parse_snapshot_from_filename(path) or datetime.fromtimestamp(path.stat().st_mtime)


def load_inventory(path: Path, prices: pd.Series | None = None) -> tuple[pd.DataFrame, date]:
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]
    df["MODULE#"]      = df["MODULE#"].str.strip()
    df["PRODUCT"]      = df["PRODUCT"].str.strip()
    df["COMM PRODUCT"] = df["COMM PRODUCT"].str.strip()
    df["ORDER NO"]     = df["ORDER NO"].str.strip()
    df["DAMAGE"]       = df["DAMAGE"].str.strip()
    df["PILOT"]        = df["PILOT"].str.strip()
    df["QUANTITY"]     = pd.to_numeric(df["QUANTITY"],     errors="coerce").fillna(0).astype(int)
    df["ETA"]          = pd.to_datetime(df["ETA"],          errors="coerce").dt.date
    df["ARRIVAL DATE"] = pd.to_datetime(df["ARRIVAL DATE"], errors="coerce").dt.date

    snapshot = snapshot_datetime(path).date()
    df["Snapshot_Date"] = snapshot
    df["Age_Days"]      = df["ARRIVAL DATE"].apply(
        lambda d: (snapshot - d).days if pd.notna(d) else None
    )
    df["IsAllocated"]   = df["ORDER NO"].astype(bool)
    df["IsDamaged"]     = df["DAMAGE"].str.upper() == "Y"

    # Attach unit price + value via the no-hyphen COMM PRODUCT key, which equals
    # the price master's col-14 ("Parts No.(with Color)") with hyphens stripped.
    price_key = df["COMM PRODUCT"].astype(str).str.upper()
    if prices is not None and len(prices):
        df["UNIT_PRICE"] = price_key.map(prices).astype(float)
    else:
        df["UNIT_PRICE"] = float("nan")
    df["VALUE"] = (df["QUANTITY"].astype(float) * df["UNIT_PRICE"].fillna(0)).round(2)
    return df, snapshot


def load_prices(path: Path = PRICE_FILE) -> pd.Series:
    """Load the per-part USD unit price from the price master Excel.

    Key = `Parts No.(with Color)` with all hyphens stripped, uppercased — this
    matches inventory's `COMM PRODUCT` column. Duplicates (same color stem with
    multiple price rows) collapse via mean. Returns an empty Series if the file
    is missing — the report degrades to all-zero values rather than crashing.
    """
    if not path.exists():
        print(f"[warn] price file not found: {path} — value-based KPIs will be 0", file=sys.stderr)
        return pd.Series(dtype=float)

    df = pd.read_excel(path, sheet_name="detail", header=11)
    key = (
        df["Parts No.(with Color)"].astype(str).str.strip()
          .str.replace("-", "", regex=False).str.upper()
    )
    price = pd.to_numeric(df["Unit Price"], errors="coerce")
    out = pd.Series(price.values, index=key.values)
    out = out[(out.index != "") & (out.index != "NAN") & out.notna()]
    out = out.groupby(level=0).mean()
    return out


def pick_prior(curr_snapshot: date, folder: Path) -> Path | None:
    """Pick the archived MODULE_LOC_*.csv with timestamp closest to but not
    exceeding curr - 7d. Scans the warehouse folder directly (snapshots and
    current.csv live side by side now). Falls back to the oldest available
    if nothing matches the 7-day target.
    """
    if not folder.exists():
        return None
    candidates = sorted(p for p in folder.glob("MODULE_LOC_*.csv"))
    if not candidates:
        return None

    target = datetime.combine(curr_snapshot, datetime.min.time()) - timedelta(days=7)
    stamped = [(p, snapshot_datetime(p)) for p in candidates]
    eligible = [(p, ts) for p, ts in stamped if ts <= target]

    if eligible:
        return max(eligible, key=lambda x: x[1])[0]
    return min(stamped, key=lambda x: x[1])[0]


def parse_period_time(s: str | None) -> pd.Timedelta | None:
    if not s or not s.strip():
        return None
    parts = s.strip().split(".")
    if len(parts) != 3:
        return None
    try:
        h, m, sec = (int(p) for p in parts)
        return pd.Timedelta(hours=h, minutes=m, seconds=sec)
    except ValueError:
        return None


def load_shipping(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]

    # IP exports use "DOCK CODE" instead of "SUPPLIER CODE", and omit
    # SERIAL NO / LOADING ENTRY columns. Strip whatever is present.
    str_cols = [
        "TEMP.TRAILER", "TRAILER NO", "SUPPLIER CODE", "DOCK CODE", "ORDER NO",
        "M/F RECEIVER", "PALLETIZASION", "SKIDID", "PARTS NO", "MODULE NO",
        "SERIAL NO", "PILOT NO",
    ]
    for c in str_cols:
        if c in df.columns:
            df[c] = df[c].str.strip()

    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0).astype(int)

    date_time_pairs = [
        ("PLAN SHIP DATE",     "PLAN SHIP TIME",     "PLAN_SHIP_DATETIME"),
        ("PICKING DATE",       "PICKING TIME",       "PICKING_DATETIME"),
        ("LOADING SET DATE",   "LOADING SET TIME",   "LOADING_SET_DATETIME"),
        ("LOADING ENTRY DATE", "LOADING ENTRY TIME", "LOADING_ENTRY_DATETIME"),
        ("SHIPMENT LOAD DATE", "SHIPMENT LOAD TIME", "SHIPMENT_LOAD_DATETIME"),
    ]
    # The WMS exports "0001-01-01 / 00.00.00" as a sentinel for "not yet picked /
    # loaded / shipped". Nullify those so downstream "is this row already
    # shipped?" checks don't treat the sentinel as a real past timestamp.
    SENTINEL_DATE = pd.Timestamp("0001-01-01")
    for date_col, time_col, dt_col in date_time_pairs:
        if date_col not in df.columns or time_col not in df.columns:
            continue
        d = pd.to_datetime(df[date_col], errors="coerce")
        d = d.mask(d <= SENTINEL_DATE)
        t = df[time_col].apply(parse_period_time)
        df[dt_col]    = d + t.fillna(pd.Timedelta(0))
        df[date_col]  = d.dt.date

    return df


def load_orders_backlog(path: Path | None) -> pd.DataFrame:
    """Load 201P (forward-looking order backlog) into a normalized frame.

    201P holds orders not yet planned into 202 — disjoint set, verified at 0%
    overlap on CUSTOMER ORDER. Two warehouse schemas:

      DS/SITE1: CUSTOMER ORDER NO., SUPPLIER CODE, SHIP DATE, SHIP TIME, ORDER DATE,
              TYPE, UC/CNL, STATUS, ..., PRODUCT NO., QUANTITY, PILOT NO.
      IP    : ORDER, CUSTOMER ORDER, ORDER DATE, CONSIGNEE, CONSIGNEE NAME,
              PLAN SHIP, SHIP DATE, PRODUCT NO., QUANTITY, ..., STATUS, ...
              — IP's SHIP DATE is the 0001-01-01 sentinel for everything;
              the real planned ship date lives in PLAN SHIP. PRODUCT NO. is
              already in the hyphenated PRODUCT form on IP (unlike 202.csv),
              so no comm-product translation is needed.

    Returns df with columns: PARTS NO, QTY, PLAN_SHIP_DATE, STATUS, ORDER_NO.
    `Skip` rows are filtered out (cancelled), `Shortage` and blank are kept.
    Missing file → empty frame.
    """
    cols = ["PARTS NO", "QTY", "PLAN_SHIP_DATE", "STATUS", "ORDER_NO"]
    if path is None or not path.exists():
        return pd.DataFrame(columns=cols)

    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df.columns = [c.strip() for c in df.columns]

    is_ip = "PLAN SHIP" in df.columns and "CUSTOMER ORDER" in df.columns
    if is_ip:
        date_col, order_col = "PLAN SHIP", "CUSTOMER ORDER"
    else:
        date_col, order_col = "SHIP DATE", "CUSTOMER ORDER NO."

    out = pd.DataFrame({
        "PARTS NO":       df.get("PRODUCT NO.", "").astype(str).str.strip(),
        "QTY":            pd.to_numeric(df.get("QUANTITY", 0), errors="coerce").fillna(0).astype(int),
        "PLAN_SHIP_DATE": pd.to_datetime(df.get(date_col, ""), errors="coerce").dt.date,
        "STATUS":         df.get("STATUS", "").astype(str).str.strip(),
        "ORDER_NO":       df.get(order_col, "").astype(str).str.strip(),
    })

    # WMS sentinel for "no plan ship yet" — drop these (only IP normally hits this,
    # since its SHIP DATE column is always 0001-01-01).
    out = out[out["PLAN_SHIP_DATE"].notna() & (out["PLAN_SHIP_DATE"] > date(1, 1, 1))]
    # Skip = cancelled/won't-ship-this-cycle. Shortage = real demand, can't fulfill;
    # still counts as ordered.
    out = out[out["STATUS"].str.lower() != "skip"]
    return out.reset_index(drop=True)


def find_backlog(folder: Path) -> Path | None:
    """Case-insensitive lookup for 201P.csv / 201P.CSV in a warehouse folder."""
    if not folder.exists():
        return None
    matches = sorted(folder.glob(BACKLOG_GLOB))
    return matches[0] if matches else None


# ─── Calculations ────────────────────────────────────────────────────────────

def iso_week_bounds(d: date) -> tuple[date, date]:
    """Return (Monday, Sunday) bounding the ISO week containing d."""
    monday = d - timedelta(days=d.weekday())
    return monday, monday + timedelta(days=6)


def _bucket_stats(df: pd.DataFrame) -> dict:
    """Lot-weighted stale bucketing on a snapshot frame.

    Every inventory row is bucketed by its OWN Age_Days, not by its part's
    oldest lot. The bucket's $ value is the sum of VALUE across rows whose
    age falls in the band, so $ buckets sum exactly to total inventory value
    and percentages reflect what's actually in each age band (vs. earlier
    per-part bucketing, which over-counted fresh stock of partially-stale
    parts as 'stale').

    Parts count = distinct PRODUCT that has at least one lot in that band.
    These counts can sum to more than the warehouse's distinct part count,
    since the same part can have lots in multiple bands — that's the honest
    answer to 'how many parts have stock in this age band?'.
    """
    empty = {"parts": 0, "qty": 0, "modules": 0, "value": 0.0}
    if df.empty:
        return {b: dict(empty) for b in ("lt30", "b30", "b60", "b90")}
    age = pd.to_numeric(df["Age_Days"], errors="coerce")
    masks = {
        "lt30": (age <= 30) | age.isna(),
        "b30":  (age > 30)  & (age <= 60),
        "b60":  (age > 60)  & (age <= 90),
        "b90":  (age > 90),
    }
    out = {}
    for b, mask in masks.items():
        rows = df.loc[mask]
        out[b] = {
            "parts":   int(rows["PRODUCT"].nunique()),
            "qty":     int(rows["QUANTITY"].sum()),
            "modules": int(rows["MODULE#"].nunique()),
            "value":   float(rows["VALUE"].sum()),
        }
    return out


def compute_kpis(curr: pd.DataFrame, prev: pd.DataFrame, ship: pd.DataFrame, ref_date: date,
                 backlog: pd.DataFrame | None = None) -> dict:
    """KPIs are computed against ref_date (the inventory snapshot date) so the
    output is idempotent for a given input set.

    `backlog` is the 201P frame (forward-looking orders not yet allocated into
    202). It augments This Week's Order in qty/value modes; modules mode is
    unchanged because backlog rows have no MODULE NO assigned yet. Pass None
    or an empty frame to skip — total tile then reflects 202 only.
    """
    week_start, week_end = iso_week_bounds(ref_date)
    if backlog is None:
        backlog = pd.DataFrame(columns=["PARTS NO", "QTY", "PLAN_SHIP_DATE"])

    parts_in_inv = set(curr["PRODUCT"].unique())
    future_ship  = ship[ship["PLAN SHIP DATE"] >= ref_date] if "PLAN SHIP DATE" in ship.columns else ship
    parts_on_ord = set(future_ship["PARTS NO"].unique()) if "PARTS NO" in future_ship.columns else set()
    # Forward-looking 201P backlog parts also count toward "ordered parts"
    if not backlog.empty:
        future_backlog = backlog[backlog["PLAN_SHIP_DATE"] >= ref_date]
        parts_on_ord |= set(future_backlog["PARTS NO"].unique())
    total_parts  = len(parts_in_inv & parts_on_ord)

    oldest = curr.groupby("PRODUCT")["Age_Days"].max()
    stale_90 = int(((oldest > 90)).sum())
    stale_60 = int(((oldest > 60) & (oldest <= 90)).sum())
    stale_30 = int(((oldest > 30) & (oldest <= 60)).sum())

    # Value-based bucketing for the new $ stale strip + WoW deltas.
    curr_buckets = _bucket_stats(curr)
    prev_buckets = _bucket_stats(prev)
    total_value  = float(curr["VALUE"].sum())

    in_week = ship["PLAN SHIP DATE"].between(week_start, week_end) if "PLAN SHIP DATE" in ship.columns else pd.Series([False] * len(ship))
    # 202.csv retains rows after physical loading (SHIPMENT LOAD DATE gets stamped
    # in but the row stays). Treat any row whose shipment has already loaded as
    # of the snapshot as "done" so This Week's Order reflects pending work —
    # otherwise the tile reads as the *historic* weekly total (massively inflated
    # mid-week, since most warehouses load 80–95% of the week's orders by Wed).
    if "SHIPMENT_LOAD_DATETIME" in ship.columns:
        cutoff_ts = pd.Timestamp(ref_date) + pd.Timedelta(days=1)
        already_shipped = ship["SHIPMENT_LOAD_DATETIME"].notna() & (ship["SHIPMENT_LOAD_DATETIME"] < cutoff_ts)
    else:
        already_shipped = pd.Series(False, index=ship.index)
    pending_in_week = in_week & ~already_shipped

    week_order_qty = int(ship.loc[pending_in_week, "QTY"].sum())
    if "MODULE NO" in ship.columns:
        mods = ship.loc[pending_in_week, "MODULE NO"].astype(str).str.strip()
        week_order_mod = int(mods[mods != ""].nunique())
    else:
        week_order_mod = 0

    # Week order $ value — apply prices to shipping plan.  Shipping's PARTS NO
    # has already been normalized to match inventory's PRODUCT, so we lift the
    # PRODUCT→price mapping out of curr (where COMM PRODUCT also lives).
    prod_to_price = (
        curr.dropna(subset=["UNIT_PRICE"]).drop_duplicates("PRODUCT")
            .set_index("PRODUCT")["UNIT_PRICE"]
        if not curr.empty else pd.Series(dtype=float)
    )
    week_order_value = 0.0
    if "PARTS NO" in ship.columns and not curr.empty:
        week_ship = ship.loc[pending_in_week].copy()
        week_ship["_PRICE"] = week_ship["PARTS NO"].map(prod_to_price).fillna(0.0)
        week_order_value = float((week_ship["QTY"] * week_ship["_PRICE"]).sum())

    # Fold in the 201P backlog (orders not yet allocated to a module). Disjoint
    # from 202, verified empirically. Qty and $ tiles grow; modules tile is
    # unchanged because backlog rows pre-date module assignment.
    if not backlog.empty:
        in_week_b = backlog["PLAN_SHIP_DATE"].between(week_start, week_end)
        bl = backlog.loc[in_week_b]
        week_order_qty += int(bl["QTY"].sum())
        if not prod_to_price.empty:
            bl_price = bl["PARTS NO"].map(prod_to_price).fillna(0.0)
            week_order_value += float((bl["QTY"] * bl_price).sum())

    curr_qty   = int(curr["QUANTITY"].sum())
    prev_qty   = int(prev["QUANTITY"].sum())
    inv_change = curr_qty - prev_qty
    curr_mod   = int(curr["MODULE#"].nunique())
    prev_mod   = int(prev["MODULE#"].nunique())
    inv_change_mod = curr_mod - prev_mod
    inv_change_value = float(curr["VALUE"].sum() - prev["VALUE"].sum())

    # qty-weighted average age — the central FIFO health number.
    # Mask NaN-aged rows on both sides; including them with age=0 would deflate
    # the average whenever ARRIVAL DATE is missing.
    age_series = pd.to_numeric(curr["Age_Days"], errors="coerce")
    qty_series = curr["QUANTITY"]
    aged_mask = age_series.notna()
    aged_qty = int(qty_series[aged_mask].sum())
    avg_age = float((age_series[aged_mask] * qty_series[aged_mask]).sum() / aged_qty) if aged_qty else 0.0

    # qty share of stock that's older than 30 days
    stale_qty = int(curr.loc[age_series > 30, "QUANTITY"].sum())
    stale_qty_pct = (stale_qty / curr_qty * 100) if curr_qty else 0.0

    unpriced_parts = int(curr.loc[curr["UNIT_PRICE"].isna(), "PRODUCT"].nunique())

    return {
        "total_parts":    total_parts,
        "stale_90":       stale_90,
        "stale_60":       stale_60,
        "stale_30":       stale_30,
        "week_order_qty": week_order_qty,
        "week_order_mod": week_order_mod,
        "week_order_value": week_order_value,
        "week_start":     week_start,
        "week_end":       week_end,
        "inv_change":     inv_change,
        "inv_change_mod": inv_change_mod,
        "inv_change_value": inv_change_value,
        "curr_qty":       curr_qty,
        "prev_qty":       prev_qty,
        "curr_mod":       curr_mod,
        "prev_mod":       prev_mod,
        "avg_age":        avg_age,
        "stale_qty":      stale_qty,
        "stale_qty_pct":  stale_qty_pct,
        "curr_buckets":   curr_buckets,
        "prev_buckets":   prev_buckets,
        "total_value":    total_value,
        "unpriced_parts": unpriced_parts,
    }


def fmt_dollars(x: float) -> str:
    """Compact USD: $52M, $10.6M, $520K, $45.

    Anything that would round to ≥ 1000K rolls over to M (1dp) so a value
    never reads as "$1,000K" beside another tile that says "$1.0M".
    """
    if x is None or pd.isna(x) or x == 0:
        return "$0"
    sign = "-" if x < 0 else ""
    a = abs(x)
    if a >= 999_500:  return f"{sign}${a/1_000_000:.1f}M"
    if a >= 1_000:    return f"{sign}${a/1_000:.0f}K"
    return f"{sign}${a:.0f}"


def fmt_dollars_k(x: float) -> str:
    """Compact USD with comma separators: $1.3M / $999K / $143K / $45.

    Anything that would round to 1000K or more rolls over to M (1dp), matching
    the tile-level `fmt_dollars` so a single number never reads as "$1,284K"
    or "$1,000K" when the rest of the report says "$1.3M" / "$1.0M". Below
    that threshold, K precision is preserved so the four bucket WoW deltas
    can still be eyeballed against the Since-Last-Week tile total.
    """
    if x is None or pd.isna(x) or x == 0:
        return "$0"
    sign = "-" if x < 0 else ""
    a = abs(x)
    # Threshold is 999_500 (rounds up to 1000K) so the K branch never emits
    # a comma-separated value ≥ 1000K — those become "$1.0M" instead.
    if a >= 999_500:
        return f"{sign}${a/1_000_000:,.1f}M"
    if a >= 1_000:
        return f"{sign}${a/1_000:,.0f}K"
    return f"{sign}${a:.0f}"


def fmt_pct(x: float) -> str:
    """1dp under 100, 0dp at ≥100; collapses near-zero to 0%."""
    if x is None or pd.isna(x):
        return "—"
    if x == 0:
        return "0%"
    if x >= 100:
        return f"{x:.0f}%"
    return f"{x:.1f}%"


def fmt_wow_delta(x: float) -> str:
    """Compact signed delta + 'since last week' suffix, e.g. '+$171K since last week'.

    Single-channel, neutral colour — sign carries direction, no health/concern
    encoding (earlier two-channel red/green/badge designs were dropped because
    the user found the colour mismatch between aged and fresh buckets harder to
    parse than the raw signed number). Same shape applies to every bucket.

    Returns HTML — the literal English phrases are wrapped in `data-i18n` spans
    so the JP language toggle swaps them. The sign + dollar amount is numeric
    and stays as-is in both languages.
    """
    if x is None or pd.isna(x):
        return '<span data-i18n="wow-no-prior">no prior week</span>'
    if x == 0:
        return '<span data-i18n="wow-no-change">no change since last week</span>'
    sign = "+" if x > 0 else "−"   # U+2212 minus, distinct from hyphen
    return f'{sign}{fmt_dollars_k(abs(x))} <span data-i18n="wow-since">since last week</span>'


def fmt_wow_count(x: int | float) -> str:
    """Compact signed integer delta + 'since last week', e.g. '+1,234 since last week'.

    Sibling of fmt_wow_delta for qty and module-count buckets (no $ formatting).
    Same i18n-spanned suffix so the language toggle swaps the literal phrase.
    """
    if x is None or pd.isna(x):
        return '<span data-i18n="wow-no-prior">no prior week</span>'
    if x == 0:
        return '<span data-i18n="wow-no-change">no change since last week</span>'
    sign = "+" if x > 0 else "−"
    return f'{sign}{abs(int(x)):,} <span data-i18n="wow-since">since last week</span>'


def _attr_esc(s: str) -> str:
    """Escape HTML for embedding in a double-quoted attribute while preserving
    tags, so the consumer can recover the original markup via `el.dataset.X`
    and inject it via innerHTML. Only `"` and `&` need escaping; `<`/`>` are
    legal in attribute values per the HTML5 spec."""
    return s.replace("&", "&amp;").replace('"', "&quot;")


# ─── Chart ───────────────────────────────────────────────────────────────────

def _series_by_arrival(df: pd.DataFrame, value: str) -> pd.Series:
    """Return a Series indexed by ARRIVAL DATE with the requested aggregation."""
    if df.empty:
        return pd.Series(dtype="float64")
    if value == "qty":
        return df.groupby("ARRIVAL DATE")["QUANTITY"].sum()
    if value == "modules":
        return df.groupby("ARRIVAL DATE")["MODULE#"].nunique()
    if value == "value":
        return df.groupby("ARRIVAL DATE")["VALUE"].sum()
    raise ValueError(value)


def _extend_flat_zero(s: pd.Series, ref_date: date) -> pd.Series:
    """Extend an inventory line series with a flat run at zero from the day
    after its last arrival through ref_date.

    The line is a per-arrival-date distribution, so a week with no new arrivals
    has no points and the line simply ends early (floating mid-air). Anchoring a
    zero the day after the last arrival and again at the snapshot date draws the
    line dropping to and sitting on zero across the empty week, which reads as
    "nothing arrived — flat at zero" instead of "data missing / cut off".
    """
    if s.empty:
        return s
    # ARRIVAL DATE indices are python date objects; normalize to Timestamp so
    # the appended tail shares the index dtype (a mixed date/Timestamp index
    # can't be sorted).
    s = s.copy()
    s.index = pd.to_datetime(s.index)
    ref_ts = pd.Timestamp(ref_date)
    last = s.index.max()
    if last >= ref_ts:
        return s
    tail = pd.Series({last + pd.Timedelta(days=1): 0.0, ref_ts: 0.0}, dtype="float64")
    return pd.concat([s, tail]).sort_index()


def build_chart(curr: pd.DataFrame, prev: pd.DataFrame, ref_date: date, weeks: int = 16) -> go.Figure:
    """Combo chart with both qty and module-count traces, toggled via buttons.

    Window is anchored to ref_date (the snapshot date), so the chart matches
    the same age math used elsewhere in the report.
    """
    cutoff = ref_date - timedelta(weeks=weeks)
    curr_w = curr[curr["ARRIVAL DATE"].notna() & (curr["ARRIVAL DATE"] >= cutoff)].copy()
    prev_w = prev[prev["ARRIVAL DATE"].notna() & (prev["ARRIVAL DATE"] >= cutoff)].copy()

    prev_modules = set(prev_w["MODULE#"]) if not prev_w.empty else set()
    curr_modules = set(curr_w["MODULE#"]) if not curr_w.empty else set()

    consumed = prev_w[prev_w["MODULE#"].isin(prev_modules - curr_modules)]
    arrived  = curr_w[curr_w["MODULE#"].isin(curr_modules - prev_modules)]

    fig = go.Figure()
    fig.update_layout(
        paper_bgcolor=COLOR_PANEL,
        plot_bgcolor=COLOR_PANEL,
        font=dict(color=COLOR_TEXT, family="Inter, Segoe UI, Roboto, sans-serif", size=12),
        margin=dict(l=44, r=16, t=10, b=36),
        height=300,
        barmode="relative",
        showlegend=False,
        xaxis=dict(gridcolor="rgba(40,52,85,0.35)", zerolinecolor="rgba(40,52,85,0.35)",
                   linecolor="rgba(40,52,85,0.6)",
                   title=None, tickfont=dict(size=11, color=COLOR_DIM),
                   type="date", tickformat="%m-%d",
                   dtick=7 * 24 * 60 * 60 * 1000, tick0=cutoff.strftime("%Y-%m-%d"),
                   # Pin the right edge to the snapshot date so a week with no
                   # arrivals still shows as an empty/flat span instead of the
                   # axis auto-trimming to the last arrival date (which makes a
                   # zero-inbound week look truncated rather than flat).
                   range=[(cutoff - timedelta(days=2)).strftime("%Y-%m-%d"),
                          (ref_date + timedelta(days=2)).strftime("%Y-%m-%d")]),
        yaxis=dict(gridcolor="rgba(40,52,85,0.35)", zerolinecolor="rgba(40,52,85,0.6)",
                   linecolor="rgba(40,52,85,0.6)",
                   title=None, tickfont=dict(size=11, color=COLOR_DIM),
                   # Default mode is VALUE — prefix ticks with $ and SI-compact.
                   # JS clears these when the user toggles to qty/modules.
                   tickprefix="$", tickformat="~s"),
        hovermode="x unified",
        hoverlabel=dict(bgcolor=COLOR_PANEL2, bordercolor=COLOR_BORDER,
                        font=dict(color=COLOR_TEXT, family="Inter, sans-serif")),
    )

    # qty traces (visible by default) — indices 0–3.
    # Bars (consumed/arrived) stay as-is; the inventory lines get a flat-zero
    # tail so a week with no arrivals draws as a line sitting on zero.
    cons_q = -_series_by_arrival(consumed, "qty")
    arr_q  =  _series_by_arrival(arrived,  "qty")
    prev_q =  _extend_flat_zero(_series_by_arrival(prev_w, "qty"), ref_date)
    curr_q =  _extend_flat_zero(_series_by_arrival(curr_w, "qty"), ref_date)

    # module-count traces — indices 4–7
    cons_m = -_series_by_arrival(consumed, "modules")
    arr_m  =  _series_by_arrival(arrived,  "modules")
    prev_m =  _extend_flat_zero(_series_by_arrival(prev_w, "modules"), ref_date)
    curr_m =  _extend_flat_zero(_series_by_arrival(curr_w, "modules"), ref_date)

    # value ($) traces — indices 8–11
    cons_v = -_series_by_arrival(consumed, "value")
    arr_v  =  _series_by_arrival(arrived,  "value")
    prev_v =  _extend_flat_zero(_series_by_arrival(prev_w, "value"), ref_date)
    curr_v =  _extend_flat_zero(_series_by_arrival(curr_w, "value"), ref_date)

    # Default visibility is VALUE (traces 8–11). QTY (0–3) and MODULES (4–7)
    # are drawn hidden — JS swaps visibility on chart-toggle click.
    fig.add_bar(name="Shipped",   x=cons_q.index, y=cons_q.values, marker_color=COLOR_BAR_CONSUME, visible=False, showlegend=False)
    fig.add_bar(name="Arrived",   x=arr_q.index,  y=arr_q.values,  marker_color=COLOR_BAR_ARRIVE,  visible=False, showlegend=False)
    fig.add_scatter(name="Last Week's Inventory", x=prev_q.index, y=prev_q.values, mode="lines", line=dict(color=COLOR_LINE_PREV, dash="dot",   width=2), visible=False, showlegend=False)
    fig.add_scatter(name="Current Inventory",     x=curr_q.index, y=curr_q.values, mode="lines", line=dict(color=COLOR_LINE_CURR, dash="solid", width=3), visible=False, showlegend=False)

    fig.add_bar(name="Shipped",   x=cons_m.index, y=cons_m.values, marker_color=COLOR_BAR_CONSUME, visible=False, showlegend=False)
    fig.add_bar(name="Arrived",   x=arr_m.index,  y=arr_m.values,  marker_color=COLOR_BAR_ARRIVE,  visible=False, showlegend=False)
    fig.add_scatter(name="Last Week's Inventory", x=prev_m.index, y=prev_m.values, mode="lines", line=dict(color=COLOR_LINE_PREV, dash="dot",   width=2), visible=False, showlegend=False)
    fig.add_scatter(name="Current Inventory",     x=curr_m.index, y=curr_m.values, mode="lines", line=dict(color=COLOR_LINE_CURR, dash="solid", width=3), visible=False, showlegend=False)

    fig.add_bar(name="Shipped",   x=cons_v.index, y=cons_v.values, marker_color=COLOR_BAR_CONSUME, visible=True, hovertemplate="$%{y:,.0f}<extra></extra>")
    fig.add_bar(name="Arrived",   x=arr_v.index,  y=arr_v.values,  marker_color=COLOR_BAR_ARRIVE,  visible=True, hovertemplate="$%{y:,.0f}<extra></extra>")
    fig.add_scatter(name="Last Week's Inventory", x=prev_v.index, y=prev_v.values, mode="lines", line=dict(color=COLOR_LINE_PREV, dash="dot",   width=2), visible=True, hovertemplate="$%{y:,.0f}<extra></extra>")
    fig.add_scatter(name="Current Inventory",     x=curr_v.index, y=curr_v.values, mode="lines", line=dict(color=COLOR_LINE_CURR, dash="solid", width=3), visible=True, hovertemplate="$%{y:,.0f}<extra></extra>")

    return fig


# ─── Detail table ────────────────────────────────────────────────────────────

def _age_class(age: int | None) -> str:
    if age is None: return ""
    if age >= 90: return "age-90"
    if age >= 60: return "age-60"
    if age >= 30: return "age-30"
    return ""


def build_detail_table(curr: pd.DataFrame, prev: pd.DataFrame, ship: pd.DataFrame, ref_date: date,
                       backlog: pd.DataFrame | None = None) -> list[dict]:
    week_start, week_end = iso_week_bounds(ref_date)
    if backlog is None:
        backlog = pd.DataFrame(columns=["PARTS NO", "QTY", "PLAN_SHIP_DATE"])

    curr_g = curr.groupby("PRODUCT").agg(
        curr_inv=("QUANTITY", "sum"),
        oldest_age=("Age_Days", "max"),
        lots=("MODULE#", "nunique"),
    )
    prev_g = prev.groupby("PRODUCT")["QUANTITY"].sum().rename("prev_inv")
    df = curr_g.join(prev_g, how="left").fillna({"prev_inv": 0})
    df["change"] = df["curr_inv"] - df["prev_inv"]

    # Per-PART week order qty — exclude rows already shipped as of ref_date,
    # matching the same filter used in compute_kpis for the tile total.
    in_week = ship["PLAN SHIP DATE"].between(week_start, week_end)
    if "SHIPMENT_LOAD_DATETIME" in ship.columns:
        already_shipped = ship["SHIPMENT_LOAD_DATETIME"].notna() & (
            ship["SHIPMENT_LOAD_DATETIME"] < pd.Timestamp(ref_date) + pd.Timedelta(days=1)
        )
    else:
        already_shipped = pd.Series(False, index=ship.index)
    pending_in_week = in_week & ~already_shipped
    wo = ship.loc[pending_in_week].groupby("PARTS NO")["QTY"].sum()

    # Fold per-PART backlog qty (201P this-week) into week_order so the detail
    # table matches the KPI tile.
    if not backlog.empty:
        in_week_b = backlog["PLAN_SHIP_DATE"].between(week_start, week_end)
        wo_b = backlog.loc[in_week_b].groupby("PARTS NO")["QTY"].sum()
        wo = wo.add(wo_b, fill_value=0)

    df = df.join(wo.rename("week_order"), on="PRODUCT").fillna({"week_order": 0})

    # Next pending shipment per PART — any row that hasn't physically loaded
    # yet, earliest PLAN_SHIP_DATETIME wins. Past-dated rows are intentionally
    # kept: a planned ship that's overdue and still not loaded is a real
    # obligation and should surface as the part's Next Order. Sentinel/NaT rows
    # (PLAN SHIP DATE unset) are excluded via .notna() — otherwise groupby.first
    # could pick them column-by-column and produce a frankenrow.
    pending = ship[~already_shipped & ship["PLAN_SHIP_DATETIME"].notna()].sort_values("PLAN_SHIP_DATETIME")
    next_per_part = pending.groupby("PARTS NO").first()
    df = df.join(
        next_per_part[["ORDER NO", "PLAN SHIP DATE", "TRAILER NO"]].rename(
            columns={"ORDER NO": "next_order", "PLAN SHIP DATE": "next_ship", "TRAILER NO": "next_dock"}
        ),
        on="PRODUCT",
    )

    df = df.sort_values("oldest_age", ascending=False, na_position="last").reset_index()

    # FIFO lot breakdown per PRODUCT: one row per ARRIVAL DATE, oldest first
    lots_by_part = {
        prod: g.groupby("ARRIVAL DATE").agg(
            qty=("QUANTITY", "sum"),
            modules=("MODULE#", "nunique"),
            module_list=("MODULE#", lambda s: sorted(s.dropna().astype(str).unique().tolist())),
            allocated=("IsAllocated", "sum"),
            damaged=("IsDamaged", "sum"),
            age=("Age_Days", "max"),
        ).sort_index().reset_index().to_dict("records")
        for prod, g in curr.groupby("PRODUCT")
    }

    rows = []
    for idx, r in df.iterrows():
        age_raw = r["oldest_age"]
        if pd.isna(age_raw):
            age_int   = None
            age_class = ""
            age_disp  = "—"
            badge     = ""
        else:
            age_int   = int(age_raw)
            badge     = "!" if age_int >= 90 else ""
            age_class = _age_class(age_int)
            age_disp  = f"{age_int}d"

        curr_inv = int(r["curr_inv"])
        prev_inv = int(r["prev_inv"])
        change   = int(r["change"])
        change_disp = f"+{change:,}" if change > 0 else (f"{change:,}" if change < 0 else "0")
        change_class = "pos" if change > 0 else ("neg" if change < 0 else "")

        next_ship_raw = r.get("next_ship")
        next_ship_disp = next_ship_raw.strftime("%m-%d") if pd.notna(next_ship_raw) else "—"

        # FIFO lot rows for this part, oldest first
        lot_rows = []
        for lot in lots_by_part.get(r["PRODUCT"], []):
            lot_age = int(lot["age"]) if pd.notna(lot["age"]) else None
            lot_rows.append({
                "arrival":   lot["ARRIVAL DATE"].isoformat() if pd.notna(lot["ARRIVAL DATE"]) else "—",
                "age":       f"{lot_age}d" if lot_age is not None else "—",
                "age_badge": "!" if (lot_age is not None and lot_age >= 90) else "",
                "age_class": _age_class(lot_age),
                "qty":       int(lot["qty"]),
                "modules":   int(lot["modules"]),
                "module_list": list(lot["module_list"]) if lot.get("module_list") is not None else [],
                "allocated": int(lot["allocated"]),
                "damaged":   int(lot["damaged"]),
            })

        rows.append({
            "row_id":      idx,
            "part":        r["PRODUCT"],
            "age_int":     age_int,
            "age":         age_disp,
            "age_badge":   badge,
            "age_class":   age_class,
            "curr_inv":    curr_inv,
            "prev_inv":    prev_inv,
            "curr":        f"{curr_inv:,}",
            "prev":        f"{prev_inv:,}",
            "change":      change_disp,
            "change_class": change_class,
            "is_new":      prev_inv == 0 and curr_inv > 0,
            "week_order_n": int(r["week_order"]),
            "week_order":  f"{int(r['week_order']):,}" if r["week_order"] else "—",
            "next_order":  r.get("next_order") if pd.notna(r.get("next_order")) else "—",
            "next_dock":   r.get("next_dock") if pd.notna(r.get("next_dock")) else "—",
            "next_ship":   next_ship_disp,
            "lots":        int(r["lots"]),
            "lot_rows":    lot_rows,
        })
    return rows


# ─── Template ────────────────────────────────────────────────────────────────

PAGE_TEMPLATE = Template(r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>FIFO Inventory Rundown — $generated_at</title>
<style>
  * { box-sizing: border-box; }
  :root {
    --bg: $bg; --panel: $panel; --panel2: $panel2; --head: $head;
    --border: $border; --border2: $border2;
    --ink: $text; --muted: $dim; --accent: $accent; --primary: $cyan;
    --good: $good; --bad: $bad; --warn: $warn;
    --age90: $age_red; --age60: $age_orange; --age30: $age_yellow;
    --shadow-sm: 0 1px 2px rgba(0, 0, 0, 0.30);
    --shadow-md: 0 4px 14px rgba(0, 0, 0, 0.35), 0 1px 2px rgba(0, 0, 0, 0.25);
    --shadow-lg: 0 12px 36px rgba(0, 0, 0, 0.45), 0 2px 6px rgba(0, 0, 0, 0.25);
    --radius: 12px;
  }
  html, body { background: var(--bg); }
  body { margin: 0; color: var(--ink);
         font-family: "Inter", "Segoe UI", system-ui, -apple-system,
                      "Hiragino Sans", "Yu Gothic UI", "Meiryo", "MS PGothic", sans-serif;
         font-size: 13px; line-height: 1.5;
         -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale;
         background-image:
           radial-gradient(1100px 480px at 8% -10%, rgba(255, 138, 76, 0.05), transparent 60%),
           radial-gradient(900px 420px at 100% 0%, rgba(96, 165, 250, 0.04), transparent 55%);
         background-attachment: fixed; }
  .wrap { max-width: none; margin: 0; padding: 22px 28px 40px; }
  a { color: inherit; }
  ::selection { background: rgba(255, 138, 76, 0.30); color: #fff; }

  /* ── Warehouse tabs ─────────────────────────────────────────── */
  .tabs { display: flex; gap: 4px; align-items: center;
          border-bottom: 1px solid var(--border); padding: 0 4px;
          margin-bottom: 22px; }
  .tab { display: inline-flex; align-items: baseline; gap: 8px;
         padding: 12px 18px; background: transparent; border: none;
         color: var(--muted); cursor: pointer;
         font: 700 13px/1 "Inter", sans-serif; letter-spacing: 0.10em;
         text-transform: uppercase;
         border-bottom: 2px solid transparent;
         transition: color 0.12s, border-color 0.12s, background 0.12s; }
  .tab:hover { color: var(--ink); background: rgba(255,255,255,0.02); }
  .tab .icon { font-size: 11px; line-height: 1; }
  .tab .loc  { color: var(--muted); font-size: 11px; font-weight: 600;
               letter-spacing: 0.10em; padding-left: 4px; }
  .tab.active { color: var(--ink); border-bottom-color: var(--accent); }
  .tab.active .icon { color: var(--good); }
  .tab.active .loc  { color: var(--accent); background: rgba(255, 138, 76, 0.12);
                      border: 1px solid rgba(255, 138, 76, 0.36); border-radius: 5px;
                      padding: 3px 8px; }
  .tab[data-tab="IP"]  .icon { color: #8b5cf6; }
  .tab[data-tab="DS"]  .icon { color: #22c55e; }
  .tab[data-tab="SITE1"] .icon { color: #3b82f6; }
  .tab[data-tab="IP"].active  { border-bottom-color: #8b5cf6; }
  .tab[data-tab="DS"].active  { border-bottom-color: #22c55e; }
  .tab[data-tab="SITE1"].active { border-bottom-color: #3b82f6; }
  .tab[data-tab="IP"].active  .loc { color: #8b5cf6; background: rgba(139, 92, 246, 0.14);
                                     border-color: rgba(139, 92, 246, 0.40); }
  .tab[data-tab="DS"].active  .loc { color: #22c55e; background: rgba(34, 197, 94, 0.14);
                                     border-color: rgba(34, 197, 94, 0.40); }
  .tab[data-tab="SITE1"].active .loc { color: #3b82f6; background: rgba(59, 130, 246, 0.14);
                                     border-color: rgba(59, 130, 246, 0.40); }

  .warehouse-section { display: none; }
  .warehouse-section.active { display: block; }

  /* ── Language toggle (upper right) ──────────────────────────── */
  .lang-toggle { position: fixed; top: 14px; right: 18px; z-index: 100;
                 display: inline-flex; padding: 3px; gap: 2px;
                 background: var(--panel); border: 1px solid var(--border);
                 border-radius: 8px; box-shadow: var(--shadow-md); }
  .lang-toggle button { padding: 6px 12px; border: none; background: transparent;
                        color: var(--muted); border-radius: 6px;
                        font: 700 11px/1 "Inter", "Hiragino Sans", "Yu Gothic UI", sans-serif;
                        letter-spacing: 0.10em; cursor: pointer;
                        transition: background 0.12s, color 0.12s; }
  .lang-toggle button:hover { color: var(--ink); background: rgba(255,255,255,0.04); }
  .lang-toggle button.active { background: rgba(255, 138, 76, 0.16);
                               color: #ffd5b8;
                               box-shadow: inset 0 0 0 1px rgba(255, 138, 76, 0.36); }
  @media print { .lang-toggle { display: none; } }

  /* ── Header ─────────────────────────────────────────────────── */
  .top { display: flex; align-items: center; justify-content: space-between; gap: 24px;
         flex-wrap: wrap; padding-bottom: 16px; margin-bottom: 18px; }
  .title-block { display: flex; align-items: center; gap: 12px; }
  h1 { margin: 0; font-size: 30px; font-weight: 800; letter-spacing: 0.04em;
       color: var(--accent); text-transform: uppercase;
       text-shadow: 0 0 22px rgba(255, 138, 76, 0.20); }
  .warehouse-section[data-wh="IP"]  h1 { color: #8b5cf6; text-shadow: 0 0 22px rgba(139, 92, 246, 0.30); }
  .warehouse-section[data-wh="DS"]  h1 { color: #22c55e; text-shadow: 0 0 22px rgba(34, 197, 94, 0.30); }
  .warehouse-section[data-wh="SITE1"] h1 { color: #3b82f6; text-shadow: 0 0 22px rgba(59, 130, 246, 0.30); }
  .tagline { color: var(--muted); font-size: 14px; margin-top: 4px;
             letter-spacing: 0.01em; }
  .pills { display: flex; gap: 6px; align-items: center; flex-wrap: wrap; }
  .pill { display: inline-flex; align-items: baseline; gap: 8px;
          padding: 8px 14px; border: 1px solid var(--border); border-radius: 8px;
          background-color: var(--panel); color: var(--ink);
          font: 500 13px/1.4 "JetBrains Mono", "Consolas", "Courier New", monospace;
          box-shadow: var(--shadow-sm); }
  .pill.curr { border-color: rgba(255, 138, 76, 0.5);
               box-shadow: 0 0 0 1px rgba(255, 138, 76, 0.10), var(--shadow-sm); }
  .pill .lbl { color: var(--muted); font-size: 10.5px; letter-spacing: 0.12em;
               text-transform: uppercase; font-weight: 700; font-family: "Inter", sans-serif; }
  .pill .delta { color: var(--muted); font-size: 11.5px; padding-left: 8px;
                 border-left: 1px solid var(--border2); margin-left: 2px;
                 font-family: "Inter", sans-serif; font-weight: 500; }
  .pill-arrow { color: var(--muted); font-size: 14px; user-select: none; }

  /* ── KPI strip ──────────────────────────────────────────────── */
  .metrics { display: grid; grid-template-columns: repeat(4, 1fr);
             background: var(--panel); border: 1px solid var(--border); border-radius: var(--radius);
             margin-bottom: 18px; overflow: hidden; box-shadow: var(--shadow-md); }
  .metric { padding: 18px 22px; border-right: 1px solid var(--border2);
            transition: background 0.12s; }
  .metric:last-child { border-right: none; }
  .metric:hover { background: var(--panel2); }
  .metric .lbl { color: var(--muted); font-size: 12.5px; font-weight: 700;
                 letter-spacing: 0.12em; text-transform: uppercase; }
  .metric .val { font-size: 42px; font-weight: 700; line-height: 1.05; margin-top: 10px;
                 font-variant-numeric: tabular-nums; color: var(--ink); letter-spacing: -0.025em; }
  .metric .val .unit { font-size: 15px; font-weight: 500; color: var(--muted);
                       margin-left: 5px; letter-spacing: 0; }
  .metric .sub { color: var(--muted); font-size: 13px; margin-top: 7px; }
  .stales { display: flex; gap: 36px; align-items: baseline; margin-top: 10px; }
  .stales > div .v { font-size: 44px; font-weight: 800; font-variant-numeric: tabular-nums;
                     letter-spacing: -0.028em; line-height: 1; }
  .stales > div .l { display: block; font-size: 13px; color: var(--muted);
                     letter-spacing: 0.05em; margin-top: 7px; font-weight: 600; }
  .stales .b90 .v { color: var(--bad); text-shadow: 0 0 22px rgba(248, 113, 113, 0.25); }
  .stales .b60 .v { color: var(--warn); }
  .stales .b30 .v { color: var(--age30); }

  /* ── Consolidated health row: Total Parts | 4 stale tiles | Week | InvCh ── */
  .metrics.health-row { grid-template-columns: 1fr 4fr 1fr 1fr; }
  /* Outer KPI tiles — big, prominent numbers */
  .metric.compact { padding: 18px 22px; }
  .metric.compact .lbl { font-size: 14px; letter-spacing: 0.10em; }
  .metric.compact .val { font-size: 40px; font-weight: 800; margin-top: 10px;
                         line-height: 1; letter-spacing: -0.035em; }
  .metric.compact .sub { font-size: 15px; margin-top: 10px; }

  /* Stale group: 4-tile sub-grid, sandwiched between KPI tiles, plus a thin
     header strip spanning all 4 columns that names the section so the WoW
     deltas at the bottom of each tile are unambiguously about week-over-week
     change in aged inventory dollars. */
  .stale-group { display: grid; grid-template-columns: repeat(4, 1fr);
                 grid-template-rows: auto 1fr;
                 border-right: 1px solid var(--border2);
                 background: rgba(0, 0, 0, 0.10); }
  .stale-group .stale-header { grid-column: 1 / -1;
                               display: flex; gap: 18px; align-items: baseline;
                               padding: 8px 18px 7px;
                               border-bottom: 1px solid var(--border2);
                               font: 700 10.5px/1.2 "Inter", sans-serif;
                               color: var(--muted); letter-spacing: 0.12em;
                               text-transform: uppercase; }
  .stale-group .stale-header .lead { color: var(--ink); }
  .stale-group .stale-header .sep  { color: var(--border); user-select: none; }
  .stale-group .stale-header .meta { font-weight: 500; letter-spacing: 0.08em;
                                     text-transform: none; font-size: 11px; }
  .stale-group .tile { padding: 14px 18px 16px; border-right: 1px solid var(--border2);
                       transition: background 0.12s; position: relative; }
  .stale-group .tile:last-child { border-right: none; }
  .stale-group .tile:hover { background: rgba(255,255,255,0.02); }
  .stale-group .tile .band { font-size: 14px; color: var(--muted);
                             letter-spacing: 0.10em; font-weight: 700;
                             text-transform: uppercase; }
  .stale-group .tile .pct { font: 800 32px/1 "Inter", sans-serif;
                            font-variant-numeric: tabular-nums; letter-spacing: -0.03em;
                            margin-top: 10px; }
  .stale-group .tile .val { font: 600 18px/1 "Inter", sans-serif;
                            font-variant-numeric: tabular-nums;
                            color: var(--muted); margin-top: 10px; }
  /* qty/modules unit suffix on the bucket .val number. The KPI tiles share
     `.metric .val .unit` (15px) above, but bucket vals are 18px not 40px so
     they need a smaller suffix to keep the unit subordinate to the number. */
  .stale-group .tile .val .unit { font-size: 12px; font-weight: 500;
                                  color: var(--muted); margin-left: 4px;
                                  letter-spacing: 0; }
  /* WoW line: signed dollar delta + 'since last week', neutral colour.
     No green/red — sign carries direction and the user reads the verb in
     their head. Earlier red/green and badge variants were tried and both
     created more confusion than they removed.  No nowrap: the phrase is
     long enough to wrap to 2 lines on narrow tiles, which is fine here. */
  .stale-group .tile .wow { display: block; margin-top: 8px;
                            font: 600 15px/1.3 "JetBrains Mono", "Consolas", monospace;
                            font-variant-numeric: tabular-nums;
                            color: var(--muted); }
  .stale-group .b90 .pct  { color: var(--bad);
                            text-shadow: 0 0 22px rgba(248, 113, 113, 0.20); }
  .stale-group .b60 .pct  { color: var(--warn); }
  .stale-group .b30 .pct  { color: var(--age30); }
  .stale-group .lt30 .pct { color: var(--good); }

  /* Responsive: under 1500px, drop the inline stale group below the KPIs */
  @media (max-width: 1500px) {
    .metrics.health-row { grid-template-columns: 1fr 1fr 1fr; }
    .stale-group { grid-column: 1 / -1; border-right: none;
                   border-top: 1px solid var(--border2); }
  }
  @media (max-width: 880px) {
    .metrics.health-row { grid-template-columns: 1fr 1fr; }
    .stale-group { grid-template-columns: 1fr 1fr; }
    /* tiles are stale-group children 2..5 (after the header). In 2-col mode:
       row 1 = children 2,3 (b90, b60) → bottom border;
       right column = children 3,5 (b60, lt30) → drop right border. */
    .stale-group .tile:nth-child(2),
    .stale-group .tile:nth-child(3) { border-bottom: 1px solid var(--border2); }
    .stale-group .tile:nth-child(3),
    .stale-group .tile:nth-child(5) { border-right: none; }
  }
  .change.pos { color: var(--good); text-shadow: 0 0 22px rgba(74, 222, 128, 0.20); }
  .change.neg { color: var(--bad); text-shadow: 0 0 22px rgba(248, 113, 113, 0.20); }
  /* '+' / '−' lives inline with the big numeric in the Since-Last-Week tile.
     Inherit the surrounding font-size (40px in .metric.compact .val) so the
     sign reads as part of the number, not as a tiny prefix mark. */
  .arrow { font-size: inherit; font-weight: inherit; margin-right: 4px; }

  /* ── Section header ─────────────────────────────────────────── */
  .section { margin-bottom: 18px; }
  .section-h { display: flex; align-items: center; justify-content: space-between;
               margin-bottom: 8px; gap: 16px; padding: 0 6px; }
  .section-h .t { font-size: 11px; font-weight: 700; letter-spacing: 0.14em;
                  text-transform: uppercase; color: var(--muted); }
  .section-h .m { color: var(--muted); font-size: 11.5px; }
  .section-h .right { display: flex; align-items: center; gap: 12px; margin-left: auto; }
  .panel { background: var(--panel); border: 1px solid var(--border); border-radius: var(--radius);
           box-shadow: var(--shadow-md); }

  /* ── Chart panel header (in-panel H1-style title) ───────────── */
  .chart-head { display: flex; align-items: center; justify-content: space-between;
                gap: 16px; padding: 18px 22px 6px; }
  .chart-head .t { font-size: 16px; font-weight: 700; letter-spacing: 0.08em;
                   color: var(--ink); text-transform: uppercase; margin: 0; }
  .chart-head .right { display: flex; align-items: center; gap: 12px; margin-left: auto; }
  .chart-head .m { color: var(--muted); font-size: 11.5px; }
  .chart-legend { display: inline-flex; align-items: center; gap: 14px; margin-left: 18px;
                  font-size: 11.5px; color: var(--muted); }
  .chart-legend .item { display: inline-flex; align-items: center; gap: 6px; }
  .chart-legend .swatch { display: inline-block; width: 12px; height: 12px; border-radius: 2px; }
  .chart-legend .line { display: inline-block; width: 18px; height: 0; }
  .chart-legend .line.solid { border-top: 3px solid #cbd5e1; }
  .chart-legend .line.dot { border-top: 2px dotted #64748b; }

  /* ── Toggle group (chart qty/modules) ───────────────────────── */
  .toggle-group { display: inline-flex; padding: 2px; background: var(--head);
                  border: 1px solid var(--border); border-radius: 7px; gap: 2px; }
  .toggle-group button { padding: 5px 12px; border: none; background: transparent;
                         color: var(--muted); border-radius: 5px;
                         font: 700 10.5px/1 "Inter", sans-serif; letter-spacing: 0.12em;
                         text-transform: uppercase; cursor: pointer;
                         transition: background 0.15s, color 0.15s; }
  .toggle-group button.active { background: rgba(255, 138, 76, 0.16);
                                color: #ffd5b8;
                                box-shadow: inset 0 0 0 1px rgba(255, 138, 76, 0.36); }
  .toggle-group button:not(.active):hover { color: var(--ink); background: rgba(255,255,255,0.03); }

  /* ── Chart ──────────────────────────────────────────────────── */
  .chart-block { overflow: hidden; }
  .chart { padding: 10px 6px 4px; }

  /* ── Filter bar ─────────────────────────────────────────────── */
  .filters { display: flex; gap: 8px; align-items: center; flex-wrap: wrap;
             padding: 12px 14px; border-bottom: 1px solid var(--border2);
             background: var(--panel2); }
  .filters .lbl-inline { color: var(--muted); font-size: 11.5px; font-weight: 600;
                         margin-right: 2px; }
  .search-wrap { position: relative; display: inline-flex; align-items: center; }
  .search-wrap::before { content: ""; position: absolute; left: 10px; top: 50%;
                         transform: translateY(-50%); width: 12px; height: 12px;
                         background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%238696b8' stroke-width='2.4' stroke-linecap='round' stroke-linejoin='round'><circle cx='11' cy='11' r='7'/><line x1='21' y1='21' x2='16.65' y2='16.65'/></svg>");
                         background-repeat: no-repeat; opacity: 0.85; pointer-events: none; }
  .filters input { background: var(--bg); border: 1px solid var(--border); border-radius: 7px;
                   color: var(--ink); padding: 7px 12px 7px 30px; width: 260px;
                   font: inherit; font-size: 12.5px;
                   transition: border-color 0.12s, box-shadow 0.12s; }
  .filters input::placeholder { color: var(--muted); }
  .filters input:focus { outline: none; border-color: var(--accent);
                         box-shadow: 0 0 0 3px rgba(255, 138, 76, 0.18); }
  .filters .btn { background: var(--panel); border: 1px solid var(--border); border-radius: 7px;
                  color: var(--ink); padding: 6px 12px; font: inherit; font-size: 12.5px;
                  font-weight: 500; cursor: pointer; transition: all 0.12s; }
  .filters .btn:hover { background: var(--head); border-color: #3a4870; }
  .filters .btn.active { background: rgba(255, 138, 76, 0.14);
                         color: #ffd5b8; border-color: rgba(255, 138, 76, 0.45);
                         box-shadow: inset 0 0 0 1px rgba(255, 138, 76, 0.18),
                                     0 0 14px rgba(255, 138, 76, 0.08); }
  .filters .btn.reset { color: var(--muted); border-color: transparent; background: transparent; }
  .filters .btn.reset:hover { color: var(--ink); background: var(--head); }
  .filters select.sort { background: var(--bg); border: 1px solid var(--border); border-radius: 7px;
                         color: var(--ink); padding: 6px 28px 6px 12px; font: inherit; font-size: 12.5px;
                         cursor: pointer; appearance: none;
                         background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='none' stroke='%238696b8' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'><polyline points='6 9 12 15 18 9'/></svg>");
                         background-repeat: no-repeat; background-position: right 10px center; }
  .filters select.sort:focus { outline: none; border-color: var(--accent);
                                box-shadow: 0 0 0 3px rgba(255, 138, 76, 0.18); }
  .filters .right { margin-left: auto; color: var(--muted); font-size: 12.5px;
                    font-variant-numeric: tabular-nums; }
  .filters .right b { color: var(--ink); font-weight: 700; }

  /* ── Inventory table ────────────────────────────────────────── */
  .panel.with-table { padding: 0; overflow: hidden; }
  table { width: 100%; border-collapse: collapse; background: var(--panel); }
  thead th { position: sticky; top: 0; background: var(--head); color: var(--muted);
             font-size: 10.5px; font-weight: 700; text-transform: uppercase;
             letter-spacing: 0.12em; text-align: left; padding: 11px 14px;
             border-bottom: 1px solid var(--border); z-index: 2; white-space: nowrap; }
  tbody td { padding: 12px 14px; vertical-align: middle; font-size: 13px; }
  /* Default: every main-row gets a strong divider below — clearly separates parts */
  tbody tr.main-row > td { border-bottom: 1px solid var(--border); }
  tbody tr:last-child > td { border-bottom: none; }
  th.num { text-align: right; }
  td.num { text-align: right; font-variant-numeric: tabular-nums;
           font-family: "JetBrains Mono", "Consolas", monospace; font-size: 13.5px; }
  td.part { font-family: "JetBrains Mono", "Consolas", monospace;
            font-size: 13.5px; font-weight: 600; letter-spacing: -0.01em; color: var(--ink); }

  /* Age — bold rounded badge */
  .age-badge { display: inline-flex; align-items: center; gap: 4px;
               padding: 4px 10px; border-radius: 999px;
               font: 700 12.5px/1 "JetBrains Mono", "Consolas", monospace;
               font-variant-numeric: tabular-nums;
               background: rgba(138, 152, 180, 0.10); color: var(--muted);
               border: 1px solid rgba(138, 152, 180, 0.18); }
  .age-badge .bang { font-family: "Inter", sans-serif; font-weight: 800; margin-left: 1px; }
  .age-30 .age-badge { background: rgba(251, 191, 36, 0.10); color: var(--age30);
                       border-color: rgba(251, 191, 36, 0.32); }
  .age-60 .age-badge { background: rgba(251, 146, 60, 0.12); color: var(--age60);
                       border-color: rgba(251, 146, 60, 0.36); }
  .age-90 .age-badge { background: rgba(248, 113, 113, 0.14); color: #fecaca;
                       border-color: rgba(248, 113, 113, 0.45);
                       box-shadow: 0 0 14px rgba(248, 113, 113, 0.12); }

  /* Severity left edge on row */
  tr.main-row > td:first-child { box-shadow: inset 3px 0 0 transparent; }
  tr.main-row.age-30 > td:first-child { box-shadow: inset 3px 0 0 var(--age30); }
  tr.main-row.age-60 > td:first-child { box-shadow: inset 3px 0 0 var(--age60); }
  tr.main-row.age-90 > td:first-child { box-shadow: inset 3px 0 0 var(--age90); }

  .pos { color: var(--good); font-weight: 600; }
  .neg { color: var(--bad); font-weight: 600; }
  .dim { color: var(--muted); }
  .ship-link { color: var(--primary); text-decoration: none;
               font-family: "JetBrains Mono", "Consolas", monospace; font-size: 13px;
               font-weight: 500; }
  .order-stack { display: flex; flex-direction: column; gap: 2px; line-height: 1.3; }
  .order-stack .dock-line { color: var(--muted); font-size: 11.5px;
                            font-family: "JetBrains Mono", "Consolas", monospace; }
  .new-tag { display: inline-block; padding: 2px 8px; margin-left: 8px; border-radius: 4px;
             background: rgba(74, 222, 128, 0.14); color: var(--good); font-size: 10.5px;
             font-weight: 700; letter-spacing: 0.07em; vertical-align: 1px;
             font-family: "Inter", sans-serif;
             border: 1px solid rgba(74, 222, 128, 0.30); }

  /* Expand affordance */
  tr.main-row { cursor: pointer; transition: background-color 0.1s; }
  tr.main-row:hover > td { background: rgba(96, 165, 250, 0.04); }
  /* Open part: parent row + lot block share the same tint and merge into one unit */
  tr.main-row.open > td { background: var(--panel2); border-bottom-color: var(--border2); }
  tr.main-row .chev { color: var(--muted); font-size: 10px; margin-right: 9px;
                      transition: transform 0.15s; display: inline-block; width: 9px; }
  tr.main-row.open .chev { transform: rotate(90deg); color: var(--accent); }

  /* ── FIFO drilldown — aligned list, oldest first ────────────── */
  tr.lot-row { display: none; }
  tr.lot-row.open { display: table-row; }
  /* Lot-row td matches its parent's open background; bottom border is the strong "end of part" divider */
  tr.lot-row > td { padding: 0 16px 16px; background: var(--panel2);
                    border-bottom: 1px solid var(--border); }
  .lot-strip { padding: 16px 20px 18px;
               background: var(--bg);
               border: 1px solid var(--border);
               border-radius: 10px;
               box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.02); }
  .lot-strip .head { font-size: 11.5px; font-weight: 700; letter-spacing: 0.10em;
                     text-transform: uppercase; color: var(--muted); margin-bottom: 14px;
                     display: flex; gap: 8px; align-items: baseline; flex-wrap: wrap; }
  .lot-strip .head .ord { color: var(--accent); font-family: "JetBrains Mono", "Consolas", monospace;
                          letter-spacing: 0; text-transform: none; font-size: 13px; font-weight: 500; }
  .lot-strip .head .sep { color: var(--border); user-select: none; padding: 0 8px; }
  .lot-strip .head .pill-mini { background: var(--panel); border: 1px solid var(--border2);
                                padding: 3px 9px; border-radius: 999px; color: var(--ink);
                                font-weight: 600; letter-spacing: 0.04em; font-size: 11.5px; }

  .lot-list { background: var(--panel); border: 1px solid var(--border); border-radius: 9px;
              overflow: hidden; }
  .lot-list .lhead, .lot-list .lline {
    display: grid;
    grid-template-columns: 4px 140px 100px 1fr 2.4fr 1fr 1fr;
    align-items: center; gap: 18px;
    padding: 11px 18px 11px 14px;
  }
  .lot-list .lhead { font-size: 11.5px; font-weight: 700; letter-spacing: 0.11em;
                     text-transform: uppercase; color: var(--muted);
                     background: var(--head); border-bottom: 1px solid var(--border);
                     padding-top: 10px; padding-bottom: 10px; }
  .lot-list .lhead .num { text-align: right; }
  .lot-list .lline { font-size: 13.5px; transition: background 0.10s; position: relative; }
  .lot-list .lline + .lline { border-top: 1px solid var(--border2); }
  .lot-list .lline:hover { background: var(--panel2); }

  /* Severity edge per line (left-most 4px column) */
  .lot-list .lline .edge { width: 4px; height: 24px; border-radius: 2px;
                           background: rgba(138, 152, 180, 0.18); }
  .lot-list .lline.age-30 .edge { background: var(--age30); }
  .lot-list .lline.age-60 .edge { background: var(--age60); }
  .lot-list .lline.age-90 .edge { background: var(--age90);
                                  box-shadow: 0 0 10px rgba(248, 113, 113, 0.40); }
  .lot-list .lline.oldest { background: rgba(248, 113, 113, 0.045); }

  .lot-list .ldate { font: 600 13.5px/1 "JetBrains Mono", "Consolas", monospace; color: var(--ink); }
  .lot-list .ldate .yr { color: var(--muted); font-weight: 500; margin-left: 4px; font-size: 12px; }
  .lot-list .lqty { font: 700 14px/1 "JetBrains Mono", "Consolas", monospace; color: var(--ink);
                    text-align: right; font-variant-numeric: tabular-nums; }
  .lot-list .lqty .u { font-weight: 500; font-size: 11.5px; color: var(--muted); margin-left: 3px; }
  .lot-list .lmod { font: 500 12.5px/1.45 "JetBrains Mono", "Consolas", monospace; color: var(--muted);
                    text-align: left; word-break: break-all; }
  .lot-list .lstatus { font: 500 12.5px/1 "Inter", sans-serif; text-align: right;
                       color: var(--muted); }
  .lot-list .lstatus.alloc { color: var(--primary); font-weight: 600; }
  .lot-list .lstatus.unalloc { color: var(--muted); }
  .lot-list .lflags { font: 600 12px/1 "Inter", sans-serif; text-align: right; color: var(--muted); }
  .lot-list .lflags .dmg { color: var(--bad); }

  /* ── Footer ─────────────────────────────────────────────────── */
  footer { color: var(--muted); font-size: 12px; text-align: right; margin-top: 28px;
           padding-top: 16px; border-top: 1px solid var(--border); }
  footer code { font-family: "JetBrains Mono", "Consolas", monospace; color: var(--ink);
                background: var(--panel); padding: 1px 6px; border-radius: 4px;
                border: 1px solid var(--border2); }

  /* ── Print ──────────────────────────────────────────────────── */
  @media print {
    html, body { background: #fff !important; color: #111 !important;
                 background-image: none !important; font-size: 11px; }
    .wrap { max-width: none; padding: 12px; }
    .filters, .toggle-group { display: none; }
    tr.main-row .chev { display: none; }
    tr.lot-row { display: table-row; }
    tr.main-row { cursor: default; }
    .panel, .metrics { background: #fff !important; box-shadow: none !important;
                       border-color: #ddd !important; }
    .metric, .pill, .chip, tbody td, thead th, tr.lot-row > td,
    .filters, .lot-strip { background: #fff !important; color: #111 !important;
                           border-color: #ddd !important; }
    .metric .lbl, .metric .sub, .stales > div .l, .dim, .pill .lbl, .pill .delta,
    .section-h .m, .lot-strip .head, .chip .age, .chip .row2, .chip .meta,
    thead th, footer { color: #555 !important; }
    .age-badge { background: #fff !important; border-color: currentColor !important; }
    .chip { break-inside: avoid; }
  }
</style>
</head>
<body>

<div class="lang-toggle" id="lang-toggle">
  <button data-lang="en" class="active">EN</button>
  <button data-lang="jp">日本語</button>
</div>

<div class="wrap">

  <div class="tabs" id="warehouse-tabs">
    $tab_buttons
  </div>

  $sections

  <footer><span data-i18n="generated">generated</span> <code>$generated_at</code></footer>

</div>

<script>
  (function () {
    // ── i18n dictionary ─────────────────────────────────────────
    const I18N = {
      'report-title':       { en: 'FIFO Inventory Rundown',                     jp: 'FIFO在庫ランダウン' },
      'tagline':            { en: 'Arrival-based FIFO allocation & compliance', jp: '入荷日基準FIFO割当・コンプライアンス' },
      'prev':               { en: 'PREV',                                       jp: '先週' },
      'curr':               { en: 'CURR',                                       jp: '今週' },
      'chart-title':        { en: 'Arrival Date — Module Movement',             jp: '入荷日 — モジュール移動' },
      'consumed':           { en: 'Shipped',                                    jp: '出荷' },
      'arrivals':           { en: 'Arrived',                                    jp: '入荷済' },
      'prev-inv-line':      { en: "Last Week's Inventory",                      jp: '先週在庫' },
      'curr-inv-line':      { en: 'Current Inventory',                          jp: '今週在庫' },
      'last-16w':           { en: 'last 16 weeks',                              jp: '直近16週' },
      'modules':            { en: 'Modules',                                    jp: 'モジュール' },
      'qty':                { en: 'Qty',                                        jp: '数量' },
      'value':              { en: 'Value ($$)',                                 jp: '金額 ($$)' },
      'total-parts':        { en: 'Total Parts',                                jp: '総部品数' },
      'in-both':            { en: 'in both inv & orders',                       jp: '在庫・注文両方' },
      'total-value':        { en: 'Total Inventory $$',                         jp: '在庫総額 $$' },
      'at-current-prices':  { en: 'at current unit prices',                     jp: '現行単価ベース' },
      'stale-parts':        { en: 'Stale Parts',                                jp: '滞留部品' },
      'stale-parts-value':  { en: 'Stale Parts — Value-Weighted',               jp: '滞留部品 — 金額加重' },
      'of-total-inv':       { en: 'of total inventory',                         jp: '総在庫に対する' },
      'wow-aged-title':     { en: 'Aged Inventory — WoW Change ($$)',           jp: '滞留在庫 — 前週比変動 ($$)' },
      'vs-prev-snapshot':   { en: 'vs. previous snapshot, current prices',      jp: '先週比、現行単価' },
      'stale-header-lead':  { en: 'Aged Inventory Composition',                 jp: '滞留在庫の構成' },
      'stale-header-meta':  { en: '% of total $$ on hand & week-over-week change in $$', jp: '在庫総額に対する % ・ 前週からの $$ 変動' },
      'this-week-order':    { en: "This Week's Order",                          jp: '今週の注文' },
      'inv-change':         { en: 'Since Last Week',                           jp: '前週から' },
      'part-no-label':      { en: 'Part#:',                                     jp: '部品番号:' },
      'search-placeholder': { en: 'filter part no…',                            jp: '部品番号で検索…' },
      'stale-90-btn':       { en: 'Stale >90d',                                 jp: '滞留 >90d' },
      'new-stock-btn':      { en: 'New Stock',                                  jp: '新着在庫' },
      'sort-label':         { en: 'Sort:',                                      jp: '並べ替え:' },
      'sort-oldest':        { en: 'Oldest first',                               jp: '古い順' },
      'sort-newest':        { en: 'Newest first',                               jp: '新しい順' },
      'sort-curr-desc':     { en: 'Current Inventory (high)',                   jp: '今週在庫(多い順)' },
      'sort-change-desc':   { en: 'Change (high)',                              jp: '変動(多い順)' },
      'sort-change-asc':    { en: 'Change (low)',                               jp: '変動(少ない順)' },
      'reset':              { en: 'Reset',                                      jp: 'リセット' },
      'parts-suffix':       { en: 'parts',                                      jp: '部品' },
      'col-part-no':        { en: 'Part No',                                    jp: '部品番号' },
      'col-age':            { en: 'Age',                                        jp: '経過日数' },
      'col-curr-inv':       { en: 'Current Inventory',                          jp: '今週在庫' },
      'col-prev-inv':       { en: "Last Week's Inventory",                      jp: '先週在庫' },
      'col-change':         { en: 'Change',                                     jp: '変動' },
      'col-week-order':     { en: "This Week's Order",                          jp: '今週の注文' },
      'col-next-order':     { en: 'Next Order / Doc / Dock',                    jp: '次回注文 / Doc / ドック' },
      'col-next-ship':      { en: 'Next Ship',                                  jp: '次回出荷' },
      'col-lots':           { en: 'Lots',                                       jp: 'ロット' },
      'new-tag':            { en: 'NEW',                                        jp: '新着' },
      'lot-arrival':        { en: 'Arrival',                                    jp: '入荷日' },
      'lot-age':            { en: 'Age',                                        jp: '経過日数' },
      'lot-qty':            { en: 'Qty',                                        jp: '数量' },
      'lot-modules':        { en: 'Modules',                                    jp: 'モジュール' },
      'lot-allocation':     { en: 'Allocation',                                 jp: '割当' },
      'lot-flags':          { en: 'Flags',                                      jp: 'フラグ' },
      'alloc-suffix':       { en: 'alloc',                                      jp: '割当' },
      'unalloc':            { en: 'unalloc',                                    jp: '未割当' },
      'dmg-suffix':         { en: 'dmg',                                        jp: '損傷' },
      'mod-suffix':         { en: 'mod',                                        jp: 'モジュール' },
      'aged-prefix':        { en: 'Aged',                                       jp: '経過' },
      'all-allocated':      { en: 'all allocated',                              jp: '全て割当済' },
      'none-allocated':     { en: 'none allocated',                             jp: '未割当' },
      'allocated-suffix':   { en: 'allocated',                                  jp: '割当済' },
      'damaged-suffix':     { en: 'damaged',                                    jp: '損傷' },
      'no-order':           { en: 'no order',                                   jp: '注文なし' },
      'next-prefix':        { en: 'Next:',                                      jp: '次回:' },
      'dock-prefix':        { en: 'Dock:',                                      jp: 'ドック:' },
      'wo-prefix':          { en: 'WO',                                         jp: '注文' },
      'generated':          { en: 'generated',                                  jp: '生成日時' },
      'age-unknown':        { en: 'age unknown',                                jp: '経過日数不明' },
      'wow-since':          { en: 'since last week',                            jp: '前週から' },
      'wow-no-change':      { en: 'no change since last week',                   jp: '前週から変動なし' },
      'wow-no-prior':       { en: 'no prior week',                              jp: '前週データなし' },
      'prev-delta-same-day':{ en: 'same day',                                   jp: '同日' },
      'unit-qty':           { en: 'qty',                                        jp: '数量' },
      'unit-modules':       { en: 'modules',                                    jp: 'モジュール' }
    };

    // Plotly trace names per language. Trace order: 0–3 are qty, 4–7 are modules,
    // 8–11 are value ($$). Each block is [Shipped, Arrived, Last Week's Inventory, Current Inventory].
    const TRACE_NAMES = {
      en: ['Shipped', 'Arrived', "Last Week's Inventory", 'Current Inventory',
           'Shipped', 'Arrived', "Last Week's Inventory", 'Current Inventory',
           'Shipped', 'Arrived', "Last Week's Inventory", 'Current Inventory'],
      jp: ['出荷', '入荷済', '先週在庫', '今週在庫',
           '出荷', '入荷済', '先週在庫', '今週在庫',
           '出荷', '入荷済', '先週在庫', '今週在庫']
    };

    let currentLang = 'en';

    function applyLang(lang) {
      currentLang = lang;
      document.documentElement.lang = (lang === 'jp') ? 'ja' : 'en';
      document.querySelectorAll('[data-i18n]').forEach(function (el) {
        const t = I18N[el.dataset.i18n];
        if (t && t[lang] != null) el.textContent = t[lang];
      });
      document.querySelectorAll('[data-i18n-placeholder]').forEach(function (el) {
        const t = I18N[el.dataset.i18nPlaceholder];
        if (t && t[lang] != null) el.placeholder = t[lang];
      });
      if (window.Plotly) {
        document.querySelectorAll('.js-plotly-plot').forEach(function (chartDiv) {
          try { window.Plotly.restyle(chartDiv, { name: TRACE_NAMES[lang] }, [0,1,2,3,4,5,6,7,8,9,10,11]); }
          catch (e) { /* chart may not be ready yet */ }
        });
      }
      document.querySelectorAll('.lang-toggle button').forEach(function (b) {
        b.classList.toggle('active', b.dataset.lang === lang);
      });
      try { localStorage.setItem('fifo-rundown-lang', lang); } catch (e) {}
    }

    document.querySelectorAll('.lang-toggle button').forEach(function (b) {
      b.addEventListener('click', function () { applyLang(b.dataset.lang); });
    });

    // ── Tab switching ────────────────────────────────────────────
    const tabs = document.querySelectorAll('#warehouse-tabs .tab');
    const sections = document.querySelectorAll('.warehouse-section');
    tabs.forEach(t => t.addEventListener('click', () => {
      const code = t.dataset.tab;
      tabs.forEach(b => b.classList.toggle('active', b === t));
      sections.forEach(s => s.classList.toggle('active', s.dataset.wh === code));
      // Re-layout any Plotly chart that became visible (it was hidden at draw time).
      const sec = document.querySelector('.warehouse-section.active');
      if (sec && window.Plotly) {
        const chartDiv = sec.querySelector('.js-plotly-plot');
        if (chartDiv) window.Plotly.Plots.resize(chartDiv);
      }
    }));

    // ── Chart-mode (QTY / MODULES / VALUE) ──────────────────────
    // Hoisted out of the per-section forEach: the toggle is global so all
    // three warehouses re-render to the same mode, and the choice persists
    // across tab switches AND page reloads via localStorage.
    let currentMode = 'value';

    function fmtDollars(x) {
      if (x === 0 || x == null || isNaN(x)) return '$$0';
      const sign = x < 0 ? '-' : '';
      const a = Math.abs(x);
      if (a >= 999500) return sign + '$$' + (a / 1e6).toFixed(1) + 'M';
      if (a >= 1e3)    return sign + '$$' + Math.round(a / 1e3) + 'K';
      return sign + '$$' + Math.round(a);
    }

    function fmtDollarsK(x) {
      if (x === 0 || x == null || isNaN(x)) return '$$0';
      const sign = x < 0 ? '-' : '';
      const a = Math.abs(x);
      if (a >= 999500) return sign + '$$' + (a / 1e6).toFixed(1) + 'M';
      if (a >= 1e3)    return sign + '$$' + Math.round(a / 1e3).toLocaleString() + 'K';
      return sign + '$$' + Math.round(a);
    }

    // Unit suffix injected after inventory numbers when the view is qty/modules.
    // Value mode is unit-less (the $$ prefix carries the unit). Skipped on bucket
    // .pct (already %) and bucket .wow (already trailing 'since last week').
    // The span is data-i18n'd so applyLang translates it on lang switch — and
    // setMode re-runs applyLang after every mode swap to catch fresh injections.
    function unitSpan(kpiKey) {
      if (kpiKey === 'qty') return ' <span class="unit" data-i18n="unit-qty">qty</span>';
      if (kpiKey === 'mod') return ' <span class="unit" data-i18n="unit-modules">modules</span>';
      return '';
    }

    function updateKpis(section, kpiKey) {
      const isVal = kpiKey === 'val';
      const suf = unitSpan(kpiKey);
      const wo = section.querySelector('.week-order-val');
      if (wo) {
        const v = parseFloat(wo.dataset[kpiKey]);
        const num = isNaN(v) ? '0' : (isVal ? fmtDollars(v) : v.toLocaleString());
        wo.innerHTML = num + suf;
      }
      const ic = section.querySelector('.inv-change-val');
      // data-no-prior is set when prev snapshot is missing — Python rendered
      // 'no prior week' once; leave it alone across mode swaps.
      if (ic && !ic.dataset.noPrior) {
        const v = parseFloat(ic.dataset[kpiKey]) || 0;
        const sign = v > 0 ? '+' : (v < 0 ? '−' : '·');
        section.querySelector('.inv-change-arrow').textContent = sign;
        const num = isVal ? fmtDollarsK(Math.abs(v)) : Math.abs(v).toLocaleString();
        section.querySelector('.inv-change-num').innerHTML = num + suf;
      }
      // Aged-bucket strip: pct / val are plain text; wow carries i18n spans,
      // so use innerHTML and re-run applyLang afterwards to translate them.
      section.querySelectorAll('.stale-group .tile').forEach(tile => {
        const pct = tile.querySelector('.pct');
        const val = tile.querySelector('.val');
        const wow = tile.querySelector('.wow');
        if (pct && pct.dataset[kpiKey] != null) pct.textContent = pct.dataset[kpiKey];
        if (val && val.dataset[kpiKey] != null) val.innerHTML  = val.dataset[kpiKey] + suf;
        if (wow && wow.dataset[kpiKey] != null) wow.innerHTML  = wow.dataset[kpiKey];
      });
    }

    function setMode(mode) {
      currentMode = mode;
      const kpiKey = mode === 'qty' ? 'qty' : (mode === 'modules' ? 'mod' : 'val');
      const start  = mode === 'qty' ? 0     : (mode === 'modules' ? 4     : 8);
      const vis = Array.from({length: 12}, (_, i) => i >= start && i < start + 4);
      const tickprefix = mode === 'value' ? '$$' : '';
      const tickformat = mode === 'value' ? '~s' : '';
      document.querySelectorAll('.warehouse-section').forEach(section => {
        const toggle = section.querySelector('.chart-toggle');
        if (toggle) {
          toggle.querySelectorAll('button').forEach(b =>
            b.classList.toggle('active', b.dataset.mode === mode));
          const chartDiv = document.getElementById(toggle.dataset.chartDiv);
          if (chartDiv && window.Plotly) {
            try {
              window.Plotly.restyle(chartDiv, { visible: vis });
              window.Plotly.relayout(chartDiv, {
                'yaxis.tickprefix': tickprefix,
                'yaxis.tickformat': tickformat,
              });
            } catch (e) { /* chart may not be ready yet */ }
          }
        }
        updateKpis(section, kpiKey);
      });
      // wow innerHTML injection drops fresh [data-i18n] nodes — translate them.
      applyLang(currentLang);
      try { localStorage.setItem('fifo-rundown-mode', mode); } catch (e) {}
    }

    document.querySelectorAll('.chart-toggle button').forEach(btn => {
      btn.addEventListener('click', () => {
        if (btn.dataset.mode === currentMode) return;
        setMode(btn.dataset.mode);
      });
    });

    // ── Per-warehouse wiring ────────────────────────────────────
    sections.forEach(section => {
      const code = section.dataset.wh;
      const input = section.querySelector('.part-filter');
      const table = section.querySelector('.detail-table');
      const tbody = table.querySelector('tbody');
      const buttons = section.querySelectorAll('.filters .btn[data-filter]');
      const reset = section.querySelector('.reset-filters');
      const counter = section.querySelector('.row-count');
      const sortSel = section.querySelector('.sort-select');
      const filters = { search: '', stale90: false, newStock: false };

      function applyFilters() {
        const q = filters.search.toLowerCase();
        let visible = 0;
        table.querySelectorAll('tr.main-row').forEach(tr => {
          const part = (tr.dataset.part || '').toLowerCase();
          const isStale90 = tr.classList.contains('age-90');
          const isNew = tr.dataset.new === '1';
          let show = true;
          if (q && !part.includes(q)) show = false;
          if (filters.stale90 && !isStale90) show = false;
          if (filters.newStock && !isNew) show = false;
          tr.style.display = show ? '' : 'none';
          const lot = section.querySelector('#lot-' + code + '-' + tr.dataset.rowId);
          if (lot) {
            if (!show) {
              lot.classList.remove('open');
              tr.classList.remove('open');
            }
            lot.style.display = show ? '' : 'none';
          }
          if (show) visible++;
        });
        counter.textContent = visible;
      }

      function sortRows(mode) {
        const pairs = [];
        table.querySelectorAll('tr.main-row').forEach(tr => {
          const lot = section.querySelector('#lot-' + code + '-' + tr.dataset.rowId);
          pairs.push({ tr, lot,
            age:    parseInt(tr.dataset.age   || '-1', 10),
            curr:   parseInt(tr.dataset.curr  || '0',  10),
            change: parseInt(tr.dataset.change|| '0',  10),
          });
        });
        const cmp = {
          'oldest':      (a, b) => b.age - a.age,
          'newest':      (a, b) => a.age - b.age,
          'curr-desc':   (a, b) => b.curr - a.curr,
          'change-desc': (a, b) => b.change - a.change,
          'change-asc':  (a, b) => a.change - b.change,
        }[mode] || ((a, b) => b.age - a.age);
        pairs.sort(cmp);
        const frag = document.createDocumentFragment();
        pairs.forEach(p => { frag.appendChild(p.tr); if (p.lot) frag.appendChild(p.lot); });
        tbody.appendChild(frag);
      }

      if (input) input.addEventListener('input', e => { filters.search = e.target.value; applyFilters(); });
      buttons.forEach(b => b.addEventListener('click', () => {
        b.classList.toggle('active');
        const key = b.dataset.filter;
        if (key === 'stale-90')  filters.stale90  = b.classList.contains('active');
        if (key === 'new-stock') filters.newStock = b.classList.contains('active');
        applyFilters();
      }));
      if (sortSel) sortSel.addEventListener('change', e => sortRows(e.target.value));
      if (reset) reset.addEventListener('click', () => {
        if (input) input.value = '';
        filters.search = ''; filters.stale90 = false; filters.newStock = false;
        buttons.forEach(b => b.classList.remove('active'));
        if (sortSel) sortSel.value = 'oldest';
        sortRows('oldest');
        applyFilters();
      });

      table.querySelectorAll('tr.main-row').forEach(tr => {
        tr.addEventListener('click', () => {
          const lot = section.querySelector('#lot-' + code + '-' + tr.dataset.rowId);
          if (!lot) return;
          tr.classList.toggle('open');
          lot.classList.toggle('open');
        });
      });
    });

    // ── Apply saved language preference (or default to EN) ──────
    let savedLang = 'en';
    try { savedLang = localStorage.getItem('fifo-rundown-lang') || 'en'; } catch (e) {}
    if (savedLang !== 'en') applyLang(savedLang);

    // ── Apply saved chart-mode preference (or default to VALUE) ──
    // The Python template renders the initial state in VALUE mode (chart
    // visibility, KPI tiles, bucket tiles), so no-op if the saved/default
    // is also 'value'. Otherwise flip everything in one shot.
    let savedMode = 'value';
    try { savedMode = localStorage.getItem('fifo-rundown-mode') || 'value'; } catch (e) {}
    if (savedMode !== 'value') setMode(savedMode);
  })();
</script>
</body>
</html>
""")


WAREHOUSE_SECTION_TEMPLATE = Template(r"""
  <div class="warehouse-section${active_class}" data-wh="$wh_code">

    <div class="top">
      <div class="title-block">
        <div>
          <h1>$wh_code <span data-i18n="report-title">FIFO Inventory Rundown</span></h1>
          <div class="tagline" data-i18n="tagline">Arrival-based FIFO allocation &amp; compliance</div>
        </div>
      </div>
      <div class="pills">
        <div class="pill"><span class="lbl" data-i18n="prev">PREV</span>$prev_date<span class="delta">$prev_delta</span></div>
        <span class="pill-arrow">→</span>
        <div class="pill curr"><span class="lbl" data-i18n="curr">CURR</span>$curr_date</div>
      </div>
    </div>

    <div class="section">
      <div class="panel chart-block">
        <div class="chart-head">
          <h2 class="t" data-i18n="chart-title">Arrival Date — Module Movement</h2>
          <div class="chart-legend">
            <span class="item"><span class="swatch" style="background:#dc4a4a"></span><span data-i18n="consumed">Shipped</span></span>
            <span class="item"><span class="swatch" style="background:#4ade80"></span><span data-i18n="arrivals">Arrived</span></span>
            <span class="item"><span class="line dot"></span><span data-i18n="prev-inv-line">Last Week's Inventory</span></span>
            <span class="item"><span class="line solid"></span><span data-i18n="curr-inv-line">Current Inventory</span></span>
          </div>
          <div class="right">
            <div class="m" data-i18n="last-16w">last 16 weeks</div>
            <div class="toggle-group chart-toggle" data-chart-div="$chart_div_id">
              <button data-mode="modules" data-i18n="modules">Modules</button>
              <button data-mode="qty" data-i18n="qty">Qty</button>
              <button data-mode="value" class="active" data-i18n="value">Value ($$)</button>
            </div>
          </div>
        </div>
        <div class="chart">$chart_html</div>
      </div>
    </div>

    <div class="metrics health-row">
      <div class="metric compact">
        <div class="lbl" data-i18n="total-parts">Total Parts</div>
        <div class="val">$total_parts</div>
        <div class="sub" data-i18n="in-both">in both inv &amp; orders</div>
      </div>
      <div class="stale-group">
        <div class="stale-header">
          <span class="lead" data-i18n="stale-header-lead">Aged Inventory Composition</span>
          <span class="sep">·</span>
          <span class="meta" data-i18n="stale-header-meta">% of total $$ on hand &amp; week-over-week change in $$</span>
        </div>
        $stale_tiles
      </div>
      <div class="metric compact">
        <div class="lbl" data-i18n="this-week-order">This Week's Order</div>
        <div class="val week-order-val" data-qty="$week_order_qty_n" data-mod="$week_order_mod_n" data-val="$week_order_value_n">$week_order_initial</div>
        <div class="sub">$week_range</div>
      </div>
      <div class="metric compact">
        <div class="lbl" data-i18n="inv-change">Since Last Week</div>
        <div class="val change inv-change-val $inv_change_class"$inv_change_no_prior_attr data-qty="$inv_change_qty_n" data-mod="$inv_change_mod_n" data-val="$inv_change_value_n"><span class="arrow inv-change-arrow">$inv_change_arrow</span><span class="inv-change-num">$inv_change_disp</span></div>
        <div class="sub">$prev_date_short → $curr_date_short</div>
      </div>
    </div>

    <div class="section">
      <div class="panel with-table">
        <div class="filters">
          <span class="lbl-inline" data-i18n="part-no-label">Part#:</span>
          <span class="search-wrap">
            <input type="text" class="part-filter" placeholder="filter part no…" data-i18n-placeholder="search-placeholder" autocomplete="off" />
          </span>
          <button class="btn" data-filter="stale-90" data-i18n="stale-90-btn">Stale &gt;90d</button>
          <button class="btn" data-filter="new-stock" data-i18n="new-stock-btn">New Stock</button>
          <span class="lbl-inline" style="margin-left:8px;" data-i18n="sort-label">Sort:</span>
          <select class="sort sort-select">
            <option value="oldest"      data-i18n="sort-oldest">Oldest first</option>
            <option value="newest"      data-i18n="sort-newest">Newest first</option>
            <option value="curr-desc"   data-i18n="sort-curr-desc">Current Inventory (high)</option>
            <option value="change-desc" data-i18n="sort-change-desc">Change (high)</option>
            <option value="change-asc"  data-i18n="sort-change-asc">Change (low)</option>
          </select>
          <button class="btn reset reset-filters" data-i18n="reset">Reset</button>
          <span class="right"><b class="row-count">$row_count</b> <span data-i18n="parts-suffix">parts</span></span>
        </div>
        <table class="detail-table">
          <thead>
            <tr>
              <th data-i18n="col-part-no">Part No</th>
              <th data-i18n="col-age">Age</th>
              <th class="num" data-i18n="col-curr-inv">Current Inventory</th>
              <th class="num" data-i18n="col-prev-inv">Last Week's Inventory</th>
              <th class="num" data-i18n="col-change">Change</th>
              <th class="num" data-i18n="col-week-order">This Week's Order</th>
              <th data-i18n="col-next-order">Next Order / Doc / Dock</th>
              <th data-i18n="col-next-ship">Next Ship</th>
              <th class="num" data-i18n="col-lots">Lots</th>
            </tr>
          </thead>
          <tbody>
            $rows
          </tbody>
        </table>
      </div>
    </div>

  </div>
""")


def _render_stale_tile(band_label: str, band_class: str,
                       q: tuple[str, str, str],
                       m: tuple[str, str, str],
                       v: tuple[str, str, str]) -> str:
    """Render one aged-bucket tile carrying all three modes' data for client-side
    swapping. Each mode tuple is `(pct_text, val_text, wow_html)`.

    `wow_html` is HTML (an i18n-spanned phrase plus a signed number), so the
    three wow strings are stashed as attr-escaped values in `data-qty`/`data-mod`/
    `data-val`. JS reads `el.dataset.X` (browser-decoded), sets `innerHTML`,
    and the spans become real DOM. `pct`/`val` are plain text and go through
    `textContent`. Initial textContent is the value-mode form, matching the
    chart toggle's default-active button.
    """
    pct_q, val_q, wow_q = q
    pct_m, val_m, wow_m = m
    pct_v, val_v, wow_v = v
    return (
        f'<div class="tile {band_class}">'
        f'<div class="band">{band_label}</div>'
        f'<div class="pct" data-qty="{_attr_esc(pct_q)}" data-mod="{_attr_esc(pct_m)}" data-val="{_attr_esc(pct_v)}">{pct_v}</div>'
        f'<div class="val" data-qty="{_attr_esc(val_q)}" data-mod="{_attr_esc(val_m)}" data-val="{_attr_esc(val_v)}">{val_v}</div>'
        f'<div class="wow" data-qty="{_attr_esc(wow_q)}" data-mod="{_attr_esc(wow_m)}" data-val="{_attr_esc(wow_v)}">{wow_v}</div>'
        f'</div>'
    )


def _short_date(iso: str, with_year: bool = False) -> str:
    """2025-12-04 → Dec 4 (or 'Dec 4, 2025' when with_year=True)"""
    try:
        d = datetime.strptime(iso, "%Y-%m-%d").date()
        fmt = "%b %-d" if sys.platform != "win32" else "%b %#d"
        if with_year:
            fmt += ", %Y"
        return d.strftime(fmt)
    except (ValueError, TypeError):
        return iso


def render_rows(rows: list[dict], wh_code: str) -> str:
    out = []
    for r in rows:
        is_new   = "1" if r["is_new"] else "0"
        has_ord  = "1" if r["week_order_n"] > 0 else "0"
        new_tag  = ' <span class="new-tag" data-i18n="new-tag">NEW</span>' if r["is_new"] else ""
        part_esc = _esc(r["part"])
        row_id   = _esc(r["row_id"])
        next_order_esc = _esc(r["next_order"])
        next_dock_esc  = _esc(r["next_dock"])

        # Age cell — pill badge, color-coded by severity, "!" appended for ≥90d
        if r["age_int"] is None:
            age_cell = '<span class="dim">—</span>'
        else:
            bang = '<span class="bang">!</span>' if r["age_badge"] else ""
            age_cell = f'<span class="age-badge">{r["age_int"]}d{bang}</span>'

        open_main = ""
        open_lot  = ""

        age_for_sort = r["age_int"] if r["age_int"] is not None else -1
        change_for_sort = r["curr_inv"] - r["prev_inv"]

        if r["next_order"] != "—":
            order_cell = (
                f'<div class="order-stack">'
                f'<span class="ship-link">{next_order_esc}</span>'
                f'<span class="dock-line"><span data-i18n="dock-prefix">Dock:</span> {next_dock_esc}</span>'
                f'</div>'
            )
        else:
            order_cell = '<span class="dim">—</span>'

        out.append(
            f'<tr class="main-row {r["age_class"]}{open_main}" data-part="{part_esc}" '
            f'data-row-id="{row_id}" data-new="{is_new}" data-order="{has_ord}" '
            f'data-age="{age_for_sort}" data-curr="{r["curr_inv"]}" data-change="{change_for_sort}">'
            f'<td class="part"><span class="chev">▸</span>{part_esc}{new_tag}</td>'
            f'<td>{age_cell}</td>'
            f'<td class="num">{_esc(r["curr"])}</td>'
            f'<td class="num dim">{_esc(r["prev"])}</td>'
            f'<td class="num {r["change_class"]}">{_esc(r["change"])}</td>'
            f'<td class="num">{_esc(r["week_order"])}</td>'
            f'<td>{order_cell}</td>'
            f'<td class="dim">{_esc(r["next_ship"])}</td>'
            f'<td class="num">{_esc(r["lots"])}</td>'
            f'</tr>'
        )

        # FIFO lot drilldown — aligned list, oldest first
        lines = [
            '<div class="lhead">'
            '<span></span>'
            '<span data-i18n="lot-arrival">Arrival</span>'
            '<span data-i18n="lot-age">Age</span>'
            '<span class="num" data-i18n="lot-qty">Qty</span>'
            '<span data-i18n="lot-modules">Modules</span>'
            '<span class="num" data-i18n="lot-allocation">Allocation</span>'
            '<span class="num" data-i18n="lot-flags">Flags</span>'
            '</div>'
        ]
        for i, lot in enumerate(r["lot_rows"]):
            lot_age_int = None
            try:
                lot_age_int = int(lot["age"].rstrip("d")) if lot["age"] != "—" else None
            except (ValueError, AttributeError):
                pass
            oldest_cls = " oldest" if i == 0 and lot_age_int is not None and lot_age_int >= 90 else ""
            date_short = _esc(_short_date(lot["arrival"], with_year=True))
            age_int_disp = _esc(lot["age"])
            age_bang = '<span class="bang">!</span>' if lot["age_badge"] else ""
            qty_disp = f'{lot["qty"]:,}'
            mod_ids = lot.get("module_list") or []
            if mod_ids:
                mod_disp = ", ".join(_esc(str(m)) for m in mod_ids)
            else:
                mod_disp = f'{lot["modules"]:,} <span class="dim" data-i18n="mod-suffix">mod</span>'
            if lot["allocated"]:
                alloc_html = (f'<span class="lstatus alloc">'
                              f'{lot["allocated"]}/{lot["modules"]} <span data-i18n="alloc-suffix">alloc</span></span>')
            else:
                alloc_html = '<span class="lstatus unalloc" data-i18n="unalloc">unalloc</span>'
            flag_html = (f'<span class="lflags"><span class="dmg">{lot["damaged"]} <span data-i18n="dmg-suffix">dmg</span></span></span>'
                         if lot["damaged"] else '<span class="lflags">—</span>')

            if lot_age_int is None:
                age_pill = '<span class="dim">—</span>'
            else:
                age_pill = f'<span class="age-badge">{age_int_disp}{age_bang}</span>'

            lines.append(
                f'<div class="lline {lot["age_class"]}{oldest_cls}">'
                f'<span class="edge"></span>'
                f'<span class="ldate">{date_short}</span>'
                f'<span>{age_pill}</span>'
                f'<span class="lqty">{qty_disp}<span class="u">q</span></span>'
                f'<span class="lmod">{mod_disp}</span>'
                f'{alloc_html}'
                f'{flag_html}'
                f'</div>'
            )

        order_meta = (f'<span class="ord">{next_order_esc}</span>'
                      if r["next_order"] != "—" else '<span class="dim" data-i18n="no-order">no order</span>')
        wo_meta    = (f'<span class="sep">·</span><span class="pill-mini"><span data-i18n="wo-prefix">WO</span> {_esc(r["week_order"])}</span>'
                      if r["week_order_n"] else '')

        # Aggregate stats describing the breakdown beneath the strip.
        # lot_rows are sorted oldest-first, so [0] is oldest, [-1] is newest.
        lot_ages = [int(l["age"].rstrip("d")) for l in r["lot_rows"] if l["age"] != "—"]
        lots_allocated = sum(1 for l in r["lot_rows"] if l["allocated"] >= l["modules"] and l["modules"] > 0)
        lots_total     = len(r["lot_rows"])
        damaged_total  = sum(l["damaged"] for l in r["lot_rows"])

        if not lot_ages:
            age_span = '<span class="dim" data-i18n="age-unknown">age unknown</span>'
        elif len(lot_ages) == 1 or lot_ages[0] == lot_ages[-1]:
            age_span = f'<span data-i18n="aged-prefix">Aged</span> <strong>{lot_ages[0]}d</strong>'
        else:
            age_span = f'<span data-i18n="aged-prefix">Aged</span> <strong>{lot_ages[-1]}d–{lot_ages[0]}d</strong>'

        if lots_total == 0:
            alloc_meta = ''
        elif lots_allocated == lots_total:
            alloc_meta = '<span class="sep">·</span><span data-i18n="all-allocated">all allocated</span>'
        elif lots_allocated == 0:
            alloc_meta = '<span class="sep">·</span><span class="dim" data-i18n="none-allocated">none allocated</span>'
        else:
            alloc_meta = f'<span class="sep">·</span><span>{lots_allocated}/{lots_total} <span data-i18n="allocated-suffix">allocated</span></span>'

        damage_meta = (f'<span class="sep">·</span><span style="color:var(--bad);">{damaged_total} <span data-i18n="damaged-suffix">damaged</span></span>'
                       if damaged_total else '')

        out.append(
            f'<tr class="lot-row{open_lot}" id="lot-{wh_code}-{row_id}">'
            f'<td colspan="9">'
            f'<div class="lot-strip">'
            f'<div class="head">'
            f'<span>{age_span}</span>'
            f'{alloc_meta}'
            f'{damage_meta}'
            f'{wo_meta}'
            f'<span class="sep">·</span><span><span data-i18n="next-prefix">Next:</span> {order_meta}</span>'
            f'</div>'
            f'<div class="lot-list">{"".join(lines)}</div>'
            f'</div>'
            f'</td></tr>'
        )

    return "\n".join(out)


# ─── Main ────────────────────────────────────────────────────────────────────

def build_warehouse_section(wh: dict, *, include_plotly_js: bool, active: bool,
                            prices: pd.Series | None = None) -> tuple[str, date]:
    """Run the full pipeline for one warehouse and return (html_fragment, curr_date)."""
    folder = PROJECT_ROOT / wh["folder"]
    inv_path = folder / CURR_INV_NAME
    ship_path = folder / SHIP_NAME

    if not inv_path.exists():
        raise FileNotFoundError(f"inventory file not found: {inv_path}")
    if not ship_path.exists():
        raise FileNotFoundError(f"shipping file not found: {ship_path}")

    print(f"[{wh['code']}] inventory: {inv_path.name}")
    curr, curr_date = load_inventory(inv_path, prices=prices)
    print(f"       {len(curr):,} rows, snapshot {curr_date}")

    prior_path = pick_prior(curr_date, folder)
    if prior_path is None:
        print(f"[{wh['code']}] no MODULE_LOC_*.csv found — using CURR as PREV (chart will be flat)")
        prev, prev_date = curr.copy(), curr_date
    else:
        print(f"[{wh['code']}] prior:     {prior_path.name}")
        prev, prev_date = load_inventory(prior_path, prices=prices)
        print(f"       {len(prev):,} rows, snapshot {prev_date}")

    print(f"[{wh['code']}] shipping:  {ship_path.name}")
    ship = load_shipping(ship_path)
    print(f"       {len(ship):,} rows")

    # IP's shipping CSV uses the no-hyphen "COMM PRODUCT" form for PARTS NO
    # (e.g. "904400825"), while DS/SITE1 use the hyphenated "PRODUCT" form
    # ("90440-0825"). Translate so all downstream joins on PARTS NO == PRODUCT
    # work regardless of warehouse. No-op for DS/SITE1.
    if "PARTS NO" in ship.columns:
        comm_to_prod = {
            c: p for c, p in zip(curr["COMM PRODUCT"], curr["PRODUCT"])
            if c and c != p
        }
        if comm_to_prod:
            ship["PARTS NO"] = ship["PARTS NO"].replace(comm_to_prod)

    backlog_path = find_backlog(folder)
    if backlog_path is None:
        print(f"[{wh['code']}] no 201P backlog file — week-order tile reflects 202 only")
        backlog = load_orders_backlog(None)
    else:
        print(f"[{wh['code']}] backlog:   {backlog_path.name}")
        backlog = load_orders_backlog(backlog_path)
        print(f"       {len(backlog):,} rows (after Skip filter)")

    kpis = compute_kpis(curr, prev, ship, curr_date, backlog=backlog)
    fig  = build_chart(curr, prev, curr_date)
    rows = build_detail_table(curr, prev, ship, curr_date, backlog=backlog)

    chart_div_id = f"movement-chart-{wh['code']}"
    chart_html = pio.to_html(
        fig,
        include_plotlyjs=("inline" if include_plotly_js else False),
        full_html=False,
        config={"displayModeBar": False, "responsive": True},
        div_id=chart_div_id,
    )

    # `has_prev` gates BOTH the Since-Last-Week tile and the bucket WoW lines.
    # When no MODULE_LOC_*.csv exists, prev_date == curr_date and every
    # delta is mechanically 0 — but rendering "·$0" misrepresents a missing
    # comparison as a flat week. Show "no prior week" instead, matching the
    # bucket WoW behavior.
    has_prev = (prev_date != curr_date)

    if has_prev:
        # Default mode is VALUE — initial display text and sign derive from the
        # $ delta. JS updateKpis swaps to qty/modules numbers on toggle.
        inv_change_for_init = kpis["inv_change_value"]
        inv_change_disp  = fmt_dollars_k(abs(inv_change_for_init))
        inv_change_arrow = "+" if inv_change_for_init > 0 else ("−" if inv_change_for_init < 0 else "·")
        inv_change_no_prior_attr = ""
    else:
        inv_change_disp  = '<span data-i18n="wow-no-prior">no prior week</span>'
        inv_change_arrow = ""
        # JS reads this attr to skip the mode-swap update for this tile, so
        # toggling QTY/MODULES/VALUE leaves "no prior week" in place.
        inv_change_no_prior_attr = ' data-no-prior="1"'
    # Last Week's Change is intentionally rendered neutral (no green/red).
    # The bucket WoW row already encodes good/bad with composition awareness
    # (growth-in-stale = bad, shrinkage-in-fresh = bad), and coloring this
    # tile by raw sign would visually contradict those (e.g., a green ▲ here
    # next to a red ▲ on the >90d bucket — both arrows point the same way).
    inv_change_class = ""

    prev_delta_days = (curr_date - prev_date).days
    prev_delta_disp = (
        f"Δ {prev_delta_days}d" if prev_delta_days
        else '<span data-i18n="prev-delta-same-day">same day</span>'
    )

    # Aged-bucket strip — emit pct/val/wow for every mode so the chart toggle
    # can swap them client-side. Per-mode totals come from compute_kpis.
    cb = kpis["curr_buckets"]
    pb = kpis["prev_buckets"]
    total_q = kpis["curr_qty"]
    total_m = kpis["curr_mod"]
    total_v = kpis["total_value"]

    def _pct(num: float, denom: float) -> str:
        return fmt_pct((num / denom * 100) if denom else 0.0)

    def _wow(b: str, kind: str) -> str:
        if not has_prev:
            return '<span data-i18n="wow-no-prior">no prior week</span>'
        delta = cb[b][kind] - pb[b][kind]
        return fmt_wow_delta(delta) if kind == "value" else fmt_wow_count(delta)

    def _tile_modes(b: str) -> tuple[tuple[str, str, str], tuple[str, str, str], tuple[str, str, str]]:
        q = (_pct(cb[b]["qty"],     total_q), f"{cb[b]['qty']:,}",            _wow(b, "qty"))
        m = (_pct(cb[b]["modules"], total_m), f"{cb[b]['modules']:,}",        _wow(b, "modules"))
        v = (_pct(cb[b]["value"],   total_v), fmt_dollars(cb[b]["value"]),    _wow(b, "value"))
        return q, m, v

    stale_tiles = "\n        ".join(
        _render_stale_tile(label, band_class, *_tile_modes(b))
        for b, label, band_class in [
            ("b90",  "&gt;90d",  "b90"),
            ("b60",  "60–90d",   "b60"),
            ("b30",  "30–60d",   "b30"),
            ("lt30", "&lt;30d",  "lt30"),
        ]
    )

    fragment = WAREHOUSE_SECTION_TEMPLATE.substitute(
        wh_code=wh["code"],
        wh_label=wh["label"],
        active_class=" active" if active else "",
        curr_date=curr_date.isoformat(),
        prev_date=prev_date.isoformat(),
        curr_date_short=_short_date(curr_date.isoformat()),
        prev_date_short=_short_date(prev_date.isoformat()),
        prev_delta=prev_delta_disp,
        chart_html=chart_html,
        chart_div_id=chart_div_id,
        total_parts=f"{kpis['total_parts']:,}",
        # 4-tile aged strip — single fragment carrying all three modes.
        stale_tiles=stale_tiles,
        # KPI tiles
        week_order_initial=fmt_dollars(kpis["week_order_value"]),
        week_order_qty_n=kpis["week_order_qty"],
        week_order_mod_n=kpis["week_order_mod"],
        week_order_value_n=f"{kpis['week_order_value']:.2f}",
        week_range=f"{_short_date(kpis['week_start'].isoformat())} → {_short_date(kpis['week_end'].isoformat())}",
        inv_change_disp=inv_change_disp,
        inv_change_arrow=inv_change_arrow,
        inv_change_class=inv_change_class,
        inv_change_no_prior_attr=inv_change_no_prior_attr,
        inv_change_qty_n=kpis["inv_change"],
        inv_change_mod_n=kpis["inv_change_mod"],
        inv_change_value_n=f"{kpis['inv_change_value']:.2f}",
        row_count=len(rows),
        rows=render_rows(rows, wh["code"]),
    )
    return fragment, curr_date


def _resolve_warehouses(spec: str) -> list[dict]:
    """Map a selection string (digit menu choice or comma-separated codes) to WAREHOUSES rows."""
    s = spec.strip().lower()
    menu = {
        "1": ["IP", "DS", "SITE1"], "":  ["IP", "DS", "SITE1"], "all": ["IP", "DS", "SITE1"],
        "2": ["IP"],  "3": ["DS"],  "4": ["SITE1"],
    }
    if s in menu:
        codes = menu[s]
    else:
        codes = [tok.strip().upper() for tok in spec.replace(";", ",").split(",") if tok.strip()]
    valid = {wh["code"] for wh in WAREHOUSES}
    bad = [c for c in codes if c not in valid]
    if not codes or bad:
        raise ValueError(f"unrecognised warehouse selection: {spec!r} (valid: {sorted(valid)} or 1/2/3/4)")
    # preserve WAREHOUSES order
    chosen_set = set(codes)
    return [wh for wh in WAREHOUSES if wh["code"] in chosen_set]


def prompt_warehouse_selection() -> list[dict]:
    """Interactive prompt asking which warehouse(s) to include in the report."""
    print()
    print("Which warehouse(s) to include in the report?")
    print("  1) All three  (IP + DS + SITE1)   [default]")
    print("  2) IP only    (Site 2 / SITE2)")
    print("  3) DS only    (SITE3)")
    print("  4) SITE1 only   (Site 1 / SITE1)")
    while True:
        try:
            choice = input("Selection [1]: ")
        except EOFError:
            # Non-interactive stdin (e.g. piped) — fall back to all three.
            print("  (no stdin — defaulting to all three)")
            return _resolve_warehouses("1")
        try:
            return _resolve_warehouses(choice)
        except ValueError as e:
            print(f"  ⚠ {e}. Try again.")


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Generate the weekly FIFO inventory HTML report.")
    ap.add_argument("--output", type=Path, default=None,
                    help="Output HTML path (default: ./FIFO Inventory Rundown YYYYMMDD.html)")
    ap.add_argument("--open",   action="store_true",
                    help="Open the result in the default browser")
    ap.add_argument("--warehouses", type=str, default=None,
                    help="Skip the interactive prompt. Accepts 'all' or a comma-separated "
                         "list of codes (e.g. 'IP', 'DS,SITE1').")
    args = ap.parse_args(argv)

    if args.warehouses is not None:
        selected = _resolve_warehouses(args.warehouses)
        print(f"[selection] {', '.join(wh['code'] for wh in selected)} (from --warehouses)")
    else:
        selected = prompt_warehouse_selection()
        print(f"[selection] {', '.join(wh['code'] for wh in selected)}")

    # SITE1 is the default-active tab (rightmost) when present, matching the existing
    # report's anchor warehouse. Fall back to the first selected warehouse otherwise.
    selected_codes = [wh["code"] for wh in selected]
    active_code = "SITE1" if "SITE1" in selected_codes else selected_codes[0]

    print(f"[prices] loading {PRICE_FILE.name}")
    prices = load_prices(PRICE_FILE)
    print(f"         {len(prices):,} unique part-keys with USD unit price")

    sections_html = []
    curr_dates: dict[str, date] = {}
    first = True
    for wh in selected:
        fragment, curr_date = build_warehouse_section(
            wh,
            include_plotly_js=first,
            active=(wh["code"] == active_code),
            prices=prices,
        )
        sections_html.append(fragment)
        curr_dates[wh["code"]] = curr_date
        first = False

    tab_buttons = "\n    ".join(
        f'<button class="tab{" active" if wh["code"] == active_code else ""}" '
        f'data-tab="{wh["code"]}">'
        f'<span class="icon">◆</span>{wh["code"]}'
        f'<span class="loc">{wh["loc"]}</span>'
        f'</button>'
        for wh in selected
    )

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    page = PAGE_TEMPLATE.substitute(
        bg=COLOR_BG, panel=COLOR_PANEL, panel2=COLOR_PANEL2, head=COLOR_HEAD,
        border=COLOR_BORDER, border2=COLOR_BORDER2,
        text=COLOR_TEXT, dim=COLOR_DIM,
        accent=COLOR_ACCENT, cyan=COLOR_CYAN, good=COLOR_GOOD, bad=COLOR_BAD, warn=COLOR_WARN,
        age_red=AGE_RED, age_orange=AGE_ORANGE, age_yellow=AGE_YELLOW,
        tab_buttons=tab_buttons,
        sections="\n".join(sections_html),
        generated_at=generated_at,
    )

    OUT_DIR.mkdir(exist_ok=True)
    stamp = curr_dates.get(active_code, max(curr_dates.values())).strftime("%Y%m%d")
    # Single-warehouse reports get a code suffix so they don't overwrite the full report.
    if args.output is not None:
        out_path = args.output
    elif len(selected) == 1:
        out_path = OUT_DIR / f"FIFO Inventory Rundown {selected[0]['code']} {stamp}.html"
    else:
        out_path = OUT_DIR / f"FIFO Inventory Rundown {stamp}.html"
    out_path.write_text(page, encoding="utf-8")
    print(f"[done] wrote {out_path} ({out_path.stat().st_size / 1024:,.0f} KB)")

    if args.open:
        webbrowser.open(out_path.as_uri())
    return 0


if __name__ == "__main__":
    sys.exit(main())
