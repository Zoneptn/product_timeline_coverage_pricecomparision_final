"""
Price Comparison view — cross-company shopping view: pick a crop, a
category (Herbicide/Insecticide/Fungicide), then a specific weed/pest/
disease (English or Thai name) — see every company's product for that
target side by side, sorted cheapest or priciest first. Uses the SAME
crop_timeline_coverage.xlsx workbook as the Product Coverage view
(junction + master sheets), just read across ALL companies instead of
one. Fertilizer is intentionally excluded — it isn't tied to a specific
pest/weed/disease target, so "price comparison for a target" doesn't
apply to it the same way.
"""

import streamlit as st
import pandas as pd

from shared import TIER_ORDER, normalize_tier, EFFICIENCY_ORDER, normalize_efficiency, _format_price
from data_cov import DEFAULT_PATH_COV, load_workbook_cov, get_file_cov

PRICE_CATEGORY_CONFIG = {
    "Herbicide (Weed)": {
        "junction": "weed_her", "master": "prod_her",
        "junction_id": "her_id", "master_id": "her_id",
        "window_sheet": "crop_weeds", "target_id_col": "weed_id",
        "name_en_col": "weed_science", "name_th_col": "weed_name_th",
        "cost_col": "price_per_rai", "cost_unit_label": "rai",
        # weed_stage groups weed windows by application timing (e.g.
        # "Early Post", "Late Post") — only crop_weeds has this concept,
        # so this key is intentionally absent from Insect/Disease below.
        "stage_col": "weed_stage",
    },
    "Insecticide (Insect)": {
        "junction": "pest_ins", "master": "prod_ins",
        "junction_id": "ins_id", "master_id": "ins_id",
        "window_sheet": "crop_pest", "target_id_col": "pest_id",
        "name_en_col": "pest_name_en", "name_th_col": "pest_name_th",
        "cost_col": "price_per_20l", "cost_unit_label": "20L tank",
    },
    "Fungicide (Disease)": {
        "junction": "disease_fun", "master": "prod_fun",
        "junction_id": "fun_id", "master_id": "fun_id",
        "window_sheet": "crop_disease", "target_id_col": "disease_id",
        "name_en_col": "disease_name_sc", "name_th_col": "disease_name_th",
        "cost_col": "price_per_20l", "cost_unit_label": "20L tank",
    },
}


def _price_stage_options(sheets: dict, cfg: dict, crop_id) -> list:
    """Distinct spray-timing values (e.g. Early Post/Late Post) for this
    crop, only meaningful when cfg has a 'stage_col' (currently just
    Herbicide). Returns [] for categories without the concept, or if the
    column isn't present in the sheet yet — the caller treats an empty
    list as 'no timing filter to offer'."""
    stage_col = cfg.get("stage_col")
    if not stage_col:
        return []
    window_df = sheets.get(cfg["window_sheet"], pd.DataFrame())
    if window_df.empty or stage_col not in window_df.columns:
        return []
    w = window_df[window_df["crop_id"] == crop_id]
    return sorted({str(v).strip() for v in w[stage_col].dropna() if str(v).strip()})


def _price_target_options(sheets: dict, cfg: dict, crop_id, stage_filter: str = None) -> pd.DataFrame:
    """Distinct (target_id, name_en, name_th) options for the target
    picker, scoped to this crop. A target can have several rows in the
    window sheet (one per pressure window) — dedupe down to one row per
    target_id since only the name is needed here. stage_filter, when
    given and cfg has a 'stage_col', narrows this to only targets that
    have at least one window at that spray timing (e.g. only weeds with
    an 'Early Post' window)."""
    window_df = sheets.get(cfg["window_sheet"], pd.DataFrame())
    if window_df.empty:
        return pd.DataFrame(columns=[cfg["target_id_col"], "name_en", "name_th"])
    w = window_df[window_df["crop_id"] == crop_id].copy()
    stage_col = cfg.get("stage_col")
    if stage_filter and stage_col and stage_col in w.columns:
        w = w[w[stage_col].astype(str).str.strip() == stage_filter]
    if w.empty:
        return pd.DataFrame(columns=[cfg["target_id_col"], "name_en", "name_th"])
    w["name_en"] = w[cfg["name_en_col"]]
    w["name_th"] = w[cfg["name_th_col"]]
    return w[[cfg["target_id_col"], "name_en", "name_th"]].drop_duplicates(
        subset=[cfg["target_id_col"]]
    ).sort_values("name_en")


def _price_comparison_table(sheets: dict, cfg: dict, crop_id, target_id,
                             ascending: bool = True) -> pd.DataFrame:
    """Every company's product linked to this target, across ALL windows
    that target appears in (deduped down to one row per product — price/
    size/usage/tier are product-level attributes and don't vary by
    window, only efficiency might, so efficiency is shown as a mix if it
    does). ascending=True sorts cheapest first, False sorts priciest
    first; rows with no price data always sort to the bottom either way
    so they never get mistaken for the cheapest (or most expensive)
    option."""
    junction = sheets.get(cfg["junction"], pd.DataFrame())
    master = sheets.get(cfg["master"], pd.DataFrame())
    j_id, m_id = cfg["junction_id"], cfg["master_id"]
    empty_cols = ["company", "trade_name", "common_name", "tier", "efficiency",
                  "price", "size", "usage", cfg["cost_col"]]
    if junction.empty or master.empty or j_id not in junction.columns or m_id not in master.columns:
        return pd.DataFrame(columns=empty_cols)

    j = junction[(junction["crop_id"] == crop_id) & (junction[cfg["target_id_col"]] == target_id)].copy()
    if j.empty:
        return pd.DataFrame(columns=empty_cols)

    overlap = [c for c in j.columns if c in master.columns and c != j_id]
    j = j.drop(columns=overlap)
    merged = j.merge(master, left_on=j_id, right_on=m_id, how="left") if j_id != m_id \
        else j.merge(master, on=j_id, how="left")
    if merged.empty or "company" not in merged.columns:
        return pd.DataFrame(columns=empty_cols)

    rows = []
    for pid, g in merged.groupby(m_id, dropna=False):
        first = g.iloc[0]
        raw_efficiencies = [normalize_efficiency(v) for v in g.get("efficiency", pd.Series(dtype=object))]
        efficiencies_present = {e for e in raw_efficiencies if e is not None}
        efficiency_display = (
            "/".join(e for e in EFFICIENCY_ORDER if e in efficiencies_present)
            if efficiencies_present else "Unrated"
        )
        rows.append({
            "company": first.get("company", ""),
            "trade_name": first.get("trade_name", ""),
            "common_name": first.get("common_name", ""),
            "tier": normalize_tier(first.get("tier")),
            "efficiency": efficiency_display,
            "price": first.get("price"),
            "size": first.get("size"),
            "usage": first.get("usage"),
            cfg["cost_col"]: first.get(cfg["cost_col"]),
        })
    out = pd.DataFrame(rows, columns=empty_cols)
    out["_sort_cost"] = pd.to_numeric(out[cfg["cost_col"]], errors="coerce")
    out = out.sort_values("_sort_cost", ascending=ascending, na_position="last").drop(columns=["_sort_cost"])
    return out.reset_index(drop=True)


def render_price_comparison_view():
    st.title("💰 Price Comparison")
    st.caption("Compare every company's product for a specific weed, pest, or disease.")

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

    stage_df_all = sheets["crop_stage"]
    if stage_df_all.empty:
        st.error("`crop_stage` sheet is missing or empty.")
        st.stop()

    crop_lookup = stage_df_all[["crop_id", "crop"]].drop_duplicates()
    crop_name_to_id = dict(zip(crop_lookup["crop"], crop_lookup["crop_id"]))

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()), key="price_crop")
    crop_id = crop_name_to_id[crop_choice]
    with col2:
        category_choice = st.selectbox("Category", list(PRICE_CATEGORY_CONFIG.keys()), key="price_category")
    with col3:
        lang_choice = st.radio("Name language", ["English", "Thai"], horizontal=True, key="price_lang")
    cfg = PRICE_CATEGORY_CONFIG[category_choice]
    name_col = "name_en" if lang_choice == "English" else "name_th"

    # Spray Timing (e.g. Early Post / Late Post) only exists for
    # Herbicide, via crop_weeds' weed_stage column — other categories
    # simply don't get this filter offered at all.
    stage_options = _price_stage_options(sheets, cfg, crop_id)
    stage_filter = None
    if stage_options:
        stage_choice = st.selectbox(
            "Spray Timing", ["All"] + stage_options, key="price_stage"
        )
        stage_filter = None if stage_choice == "All" else stage_choice

    targets = _price_target_options(sheets, cfg, crop_id, stage_filter=stage_filter)
    if targets.empty:
        if stage_filter:
            st.info(f"No {category_choice.lower()} targets found for '{stage_filter}' timing on this crop.")
        else:
            st.info(f"No {category_choice.lower()} targets found for this crop.")
        st.stop()

    target_name_to_id = dict(zip(targets[name_col], targets[cfg["target_id_col"]]))
    col4, col5 = st.columns([3, 1])
    with col4:
        target_choice = st.selectbox(
            f"{category_choice.split(' ')[0]} target", list(target_name_to_id.keys()), key="price_target"
        )
    with col5:
        sort_choice = st.radio("Sort", ["Cheapest first", "Priciest first"],
                                horizontal=True, key="price_sort")
    target_id = target_name_to_id[target_choice]

    table = _price_comparison_table(sheets, cfg, crop_id, target_id,
                                     ascending=(sort_choice == "Cheapest first"))
    if table.empty:
        st.info(f"No products found across any company for {target_choice}.")
        st.stop()

    display = table.copy()
    display["price"] = display["price"].apply(_format_price)
    display[cfg["cost_col"]] = display[cfg["cost_col"]].apply(
        lambda v: _format_price(v) + f"/{cfg['cost_unit_label']}" if _format_price(v) else "—"
    )
    display = display.rename(columns={
        "company": "Company", "trade_name": "Trade Name", "common_name": "Common Name",
        "tier": "Tier", "efficiency": "Efficiency", "price": "Price", "size": "Size", "usage": "Usage",
        cfg["cost_col"]: cfg["cost_col"].replace("_", " ").title(),
    })
    st.subheader(f"{target_choice} — {len(display)} product(s) across {display['Company'].nunique()} company(ies)")
    st.dataframe(display, use_container_width=True, hide_index=True)
