"""
Loading/caching for crop_timeline_coverage.xlsx. Shared by coverage_view.py
(the Product Coverage board) and price_view.py (Price Comparison) since
both read the exact same workbook — just sliced differently.
"""

import os
import streamlit as st
import pandas as pd

DEFAULT_PATH_COV = "crop_timeline_coverage.xlsx"

SHEET_NAMES_COV = [
    "crop_stage",
    "crop_weeds", "weed_her", "prod_her",
    "crop_pest", "pest_ins", "prod_ins",
    "crop_disease", "disease_fun", "prod_fun",
    "crop_fer", "fertilizer", "prod_fer",
    "price_history",
]

# junction sheet -> (its id column, master sheet, master's id column)
CATEGORY_CONFIG_COV = {
    "weed_her": {"junction_id": "her_id", "master": "prod_her", "master_id": "her_id",
                 "code_col": "hrac_code", "code_label": "HRAC"},
    "pest_ins": {"junction_id": "ins_id", "master": "prod_ins", "master_id": "ins_id",
                 "code_col": "irac_code", "code_label": "IRAC"},
    "disease_fun": {"junction_id": "fun_id", "master": "prod_fun", "master_id": "fun_id",
                     "code_col": "frac_code", "code_label": "FRAC"},
    "fertilizer": {"junction_id": "fer_id", "master": "prod_fer", "master_id": "fer_id",
                   "code_col": "type", "code_label": "Type"},
}

# price_history schema (append-only, one row per product per quarterly
# snapshot): snapshot_date, product_id, category, trade_name,
# common_name, concentration, formulation_type, company, price.
# "category" is one of these exact labels — must match price_history's
# actual values for the fallback join below to find anything.
#
# Fertilizer (prod_fer) is intentionally NOT part of this — it has a
# different shape (price_per_ton/bag_size/price_per_bag, no
# price/size/usage) and isn't tracked in price_history yet.
_PRICE_FALLBACK_CONFIG = {
    "prod_her": {"master_id_col": "her_id", "cost_col": "price_per_rai", "category_label": "Herbicide"},
    "prod_ins": {"master_id_col": "ins_id", "cost_col": "price_per_20l", "category_label": "Insecticide"},
    "prod_fun": {"master_id_col": "fun_id", "cost_col": "price_per_20l", "category_label": "Fungicide"},
}


def _latest_prices_from_history(price_history_df: pd.DataFrame, category_label: str) -> pd.Series:
    """Returns a Series indexed by product_id (as string) -> latest
    price (by snapshot_date) among price_history rows matching this
    category. Empty Series if price_history is missing, empty, lacks
    the needed columns, or has no rows for this category — callers
    treat that as 'no price data available' (shows as blank downstream,
    the same as any other missing-data case in this app), not an
    error."""
    required = {"snapshot_date", "product_id", "category", "price"}
    if price_history_df.empty or not required.issubset(price_history_df.columns):
        return pd.Series(dtype=float)
    df = price_history_df[
        price_history_df["category"].astype(str).str.strip() == category_label
    ].copy()
    if df.empty:
        return pd.Series(dtype=float)
    # dayfirst=True as a defensive backstop: pandas assumes month-first
    # (MM-DD-YYYY) by default, which would silently misread a DD-MM-YYYY
    # string (e.g. a date typed by hand in Excel) for any day-of-month
    # <=12 — no error, just a wrong "latest" snapshot picked from then
    # on. This doesn't help if dates are ALREADY ambiguous both ways in
    # the same column, but it matches the day-first convention used
    # when writing new snapshot rows (see update_price_example.py).
    df["snapshot_date"] = pd.to_datetime(df["snapshot_date"], errors="coerce", dayfirst=True)
    df["product_id"] = df["product_id"].astype(str).str.strip()
    df = df.dropna(subset=["snapshot_date"])
    if df.empty:
        return pd.Series(dtype=float)
    # Keep only the single latest snapshot per product_id — ties (two
    # rows with the identical latest date for the same product, a
    # data-entry duplicate) resolve to whichever pandas keeps first,
    # which is an edge case worth cleaning up in the sheet if it ever
    # happens, not something to silently average or sum.
    latest_idx = df.groupby("product_id")["snapshot_date"].idxmax()
    return df.loc[latest_idx].set_index("product_id")["price"]


def _apply_price_fallback(sheets: dict) -> dict:
    """For prod_her/prod_ins/prod_fun: if a master sheet no longer has
    its own 'price' AND per-unit cost column (price_per_rai /
    price_per_20l) — i.e. they've been fully migrated out in favor of
    price_history — both are computed here instead, live, using the
    exact same formula the old Excel columns used: cost = usage * price
    / size. 'price' itself comes from price_history's latest snapshot
    for that product; 'size'/'usage' still come from the master sheet
    (per the earlier decision that those rarely change and don't need
    per-snapshot tracking).

    Deliberately only activates when BOTH 'price' and the cost column
    are absent — a sheet that still has either one is treated as 'not
    yet migrated' and left completely untouched, so a partial/in-progress
    migration can never silently overwrite a still-valid manually-typed
    price with a possibly-stale price_history value.

    Once this fills in 'price' and the cost column, every other part of
    the app (coverage_view.py, price_view.py) keeps working exactly as
    before without any changes — they just see those columns already
    populated, same names, same meaning, regardless of whether the
    numbers came from the sheet directly or were computed here."""
    price_history = sheets.get("price_history", pd.DataFrame())
    for master_name, cfg in _PRICE_FALLBACK_CONFIG.items():
        master = sheets.get(master_name, pd.DataFrame())
        if master.empty:
            continue
        cost_col = cfg["cost_col"]
        if "price" in master.columns or cost_col in master.columns:
            continue  # not yet migrated (or only partially) — leave alone
        id_col = cfg["master_id_col"]
        if id_col not in master.columns or "size" not in master.columns or "usage" not in master.columns:
            continue  # can't compute without these; leave sheet as-is

        latest_prices = _latest_prices_from_history(price_history, cfg["category_label"])
        master = master.copy()
        product_ids = master[id_col].astype(str).str.strip()
        master["price"] = product_ids.map(latest_prices)
        size_num = pd.to_numeric(master["size"], errors="coerce")
        usage_num = pd.to_numeric(master["usage"], errors="coerce")
        price_num = pd.to_numeric(master["price"], errors="coerce")
        master[cost_col] = usage_num * price_num / size_num
        sheets[master_name] = master
    return sheets


@st.cache_data
def load_workbook_cov(file):
    sheets = {}
    for name in SHEET_NAMES_COV:
        try:
            df = pd.read_excel(file, sheet_name=name)
            df.columns = [c.strip() for c in df.columns]
            df = df.loc[:, ~df.columns.str.startswith("Unnamed:")]
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
            sheets[name] = df
        except ValueError:
            sheets[name] = pd.DataFrame()
    sheets = _apply_price_fallback(sheets)
    return sheets


def get_file_cov():
    st.sidebar.subheader("Product Coverage data source")
    uploaded = st.sidebar.file_uploader(
        "Upload workbook (.xlsx)", type=["xlsx"], key="cov_uploader"
    )
    if st.sidebar.button("🔄 Reload data", key="cov_reload"):
        st.cache_data.clear()
        st.rerun()
    if uploaded is not None:
        return uploaded
    if os.path.exists(DEFAULT_PATH_COV):
        return DEFAULT_PATH_COV
    return None


def load_product_df(sheets: dict, junction_name: str, crop_id) -> pd.DataFrame:
    """Join a junction sheet to its product master sheet, scoped to one
    crop. The master sheet is the single source of truth: any column
    that exists on both sides (e.g. a leftover trade_name kept on the
    junction sheet for readability) is dropped from the junction copy
    before merging, so the master's value always wins. Returns an empty
    DataFrame if either sheet is missing/empty or the id columns aren't
    present."""
    cfg = CATEGORY_CONFIG_COV[junction_name]
    junction = sheets.get(junction_name, pd.DataFrame())
    master = sheets.get(cfg["master"], pd.DataFrame())
    j_id, m_id = cfg["junction_id"], cfg["master_id"]

    if junction.empty or master.empty or j_id not in junction.columns or m_id not in master.columns:
        return pd.DataFrame()

    j = junction[junction["crop_id"] == crop_id].copy()
    if j.empty:
        return j

    overlap = [c for c in j.columns if c in master.columns and c != j_id]
    j = j.drop(columns=overlap)

    if j_id == m_id:
        return j.merge(master, on=j_id, how="left")
    return j.merge(master, left_on=j_id, right_on=m_id, how="left")
