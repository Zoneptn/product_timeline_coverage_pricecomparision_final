"""
Price Trend view — reads price_history from crop_timeline_coverage.xlsx
(the same workbook Product Coverage and Price Comparison already read)
and shows what changed between the two most recent price-update rounds
per product — the "what should I pay attention to right now" companion
to a full price-history trend chart (a possible later addition).

Kept in its own file rather than folded into price_view.py: this is a
genuinely different kind of analysis (a time-series comparison across
price_history's raw snapshots) from price_view.py's four modes, which
are all about comparing companies' CURRENT prices against each other
for a specific pest/weed/disease target. price_view.py was already
large; this avoids growing it further with an unrelated concern.
"""

import streamlit as st
import pandas as pd

from shared import _format_price
from data_cov import DEFAULT_PATH_COV, load_workbook_cov, get_file_cov

CATEGORY_OPTIONS = ["All", "Herbicide", "Insecticide", "Fungicide"]


def _price_movement_table(price_history: pd.DataFrame) -> pd.DataFrame:
    """For every (product_id, category) with at least 2 distinct
    snapshot dates, compares the two MOST RECENT ones — i.e. 'what
    changed since the last update round' — regardless of whether every
    product was included in every batch (a product updated in January
    and not again until July still gets compared January-vs-July, not
    silently skipped for missing an April entry that never existed).
    Grouped by (product_id, category) rather than product_id alone, in
    case an ID were ever reused across categories.

    Returns one row per product with old/new price, date, and %
    change, sorted biggest movers first (by absolute % change — a
    small currency change on a cheap product can be a huge % swing and
    vice versa, so % is the fairer 'what deserves attention' signal
    than raw currency change alone). Empty DataFrame if price_history
    has fewer than 2 usable snapshots anywhere."""
    required = {"snapshot_date", "product_id", "category", "price"}
    empty_cols = ["category", "company", "trade_name", "common_name", "product_id",
                  "old_date", "old_price", "new_date", "new_price", "change", "pct_change"]
    if price_history.empty or not required.issubset(price_history.columns):
        return pd.DataFrame(columns=empty_cols)

    df = price_history.copy()
    df["snapshot_date"] = pd.to_datetime(df["snapshot_date"], errors="coerce", dayfirst=True)
    df["product_id"] = df["product_id"].astype(str).str.strip()
    df["category"] = df["category"].astype(str).str.strip()
    df = df.dropna(subset=["snapshot_date", "price"])
    if df.empty:
        return pd.DataFrame(columns=empty_cols)

    rows = []
    for (pid, cat), g in df.groupby(["product_id", "category"]):
        distinct_dates = g["snapshot_date"].drop_duplicates().sort_values()
        if len(distinct_dates) < 2:
            continue  # only one snapshot ever recorded — nothing to compare yet
        new_date = distinct_dates.iloc[-1]
        old_date = distinct_dates.iloc[-2]
        # A data-entry duplicate (two rows on the same date for the
        # same product) takes the last-entered one for that date
        # rather than erroring or averaging.
        new_row = g[g["snapshot_date"] == new_date].iloc[-1]
        old_row = g[g["snapshot_date"] == old_date].iloc[-1]
        old_price = pd.to_numeric(old_row["price"], errors="coerce")
        new_price = pd.to_numeric(new_row["price"], errors="coerce")
        if pd.isna(old_price) or pd.isna(new_price):
            continue
        change = new_price - old_price
        pct_change = (change / old_price * 100) if old_price != 0 else None
        rows.append({
            "category": cat,
            "company": new_row.get("company", ""),
            "trade_name": new_row.get("trade_name", ""),
            "common_name": new_row.get("common_name", ""),
            "product_id": pid,
            "old_date": old_date,
            "old_price": old_price,
            "new_date": new_date,
            "new_price": new_price,
            "change": change,
            "pct_change": pct_change,
        })
    out = pd.DataFrame(rows, columns=empty_cols)
    if out.empty:
        return out
    out["_abs_pct"] = out["pct_change"].abs()
    out = out.sort_values("_abs_pct", ascending=False, na_position="last").drop(columns=["_abs_pct"])
    return out.reset_index(drop=True)


def render_price_trend_view():
    st.title("📈 Price Movement")
    st.caption(
        "What changed between the two most recent price-update rounds, "
        "per product — sorted by biggest movers first. Only products "
        "with at least 2 recorded snapshots show up here; a product "
        "with just one price entry so far has nothing to compare yet."
    )

    data_file = get_file_cov()
    if data_file is None:
        st.warning(
            f"No workbook found. Upload one from the sidebar, or place a file "
            f"named `{DEFAULT_PATH_COV}` next to `app.py`."
        )
        st.stop()

    try:
        sheets = load_workbook_cov(data_file)
    except Exception as e:
        st.error(f"Couldn't read the workbook: {e}")
        st.stop()

    price_history = sheets.get("price_history", pd.DataFrame())
    if price_history.empty:
        st.info("No `price_history` sheet found, or it's empty — nothing to show yet.")
        st.stop()

    table = _price_movement_table(price_history)
    if table.empty:
        st.info("No product has at least 2 recorded price snapshots yet — "
                 "check back after the next update round.")
        st.stop()

    # Company options come from every company with at least one recorded
    # price change (the unfiltered `table`), not from `filtered` — kept
    # independent of the Category/Direction picks below, same as
    # Category itself doesn't narrow based on Direction. Cascading
    # filters (where one dropdown's options shift based on another)
    # would be more surprising here than useful.
    company_options = ["All"] + sorted(table["company"].dropna().astype(str).str.strip().unique().tolist())

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        category_choice = st.selectbox("Category", CATEGORY_OPTIONS, key="pm_category")
    with col2:
        company_choice = st.selectbox("Company", company_options, key="pm_company")
    with col3:
        direction_choice = st.radio("Direction", ["All", "Increases only", "Decreases only"],
                                     horizontal=True, key="pm_direction")

    filtered = table.copy()
    if category_choice != "All":
        filtered = filtered[filtered["category"] == category_choice]
    if company_choice != "All":
        filtered = filtered[filtered["company"] == company_choice]
    if direction_choice == "Increases only":
        filtered = filtered[filtered["change"] > 0]
    elif direction_choice == "Decreases only":
        filtered = filtered[filtered["change"] < 0]

    if filtered.empty:
        st.info("No price movements match this filter.")
        st.stop()

    display = filtered.copy()
    display["Old Price"] = display.apply(
        lambda r: f"{_format_price(r['old_price'])} ({r['old_date'].strftime('%Y-%m-%d')})", axis=1)
    display["New Price"] = display.apply(
        lambda r: f"{_format_price(r['new_price'])} ({r['new_date'].strftime('%Y-%m-%d')})", axis=1)
    display["Change"] = display["change"].apply(
        lambda v: f"{'+' if v > 0 else ''}{_format_price(v)}")
    display["% Change"] = display["pct_change"].apply(
        lambda v: f"{'+' if v > 0 else ''}{v:.1f}%" if pd.notna(v) else "—")
    display = display.rename(columns={
        "category": "Category", "company": "Company", "trade_name": "Trade Name",
        "common_name": "Common Name",
    })
    display = display[["Category", "Company", "Trade Name", "Common Name",
                        "Old Price", "New Price", "Change", "% Change"]]

    st.subheader(f"{len(display)} product(s) with a recorded price change")
    st.dataframe(display, use_container_width=True, hide_index=True)
