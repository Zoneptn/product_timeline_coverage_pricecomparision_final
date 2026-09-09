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

price_history itself is purely product-level (no crop/pest linkage at
all) — so narrowing "what changed" down to a specific crop, category,
spray timing, or target pest/weed/disease means cross-referencing a
product_id against the same junction sheets (weed_her/pest_ins/
disease_fun) that price_view.py already uses for its target lookups.
Rather than reimplementing that, this file imports and reuses
price_view.py's PRICE_CATEGORY_CONFIG / _price_stage_options /
_price_target_options / _target_ws_ids directly, so the two files can
never quietly disagree about how targets or timings are resolved.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

from shared import _format_price
from data_cov import DEFAULT_PATH_COV, load_workbook_cov, get_file_cov
from price_view import (
    PRICE_CATEGORY_CONFIG, _price_stage_options, _price_target_options, _target_ws_ids,
)

CATEGORY_OPTIONS = ["All", "Herbicide", "Insecticide", "Fungicide"]

# price_trend_view uses short category labels (matching price_history's
# own "category" column) in its UI/filtering; price_view.py's config
# dict is keyed by the longer display labels used in ITS dropdowns.
# This maps one to the other so both files can describe "Herbicide"
# without needing to agree on a single label everywhere.
_CATEGORY_TO_CONFIG_KEY = {
    "Herbicide": "Herbicide (Weed)",
    "Insecticide": "Insecticide (Insect)",
    "Fungicide": "Fungicide (Disease)",
}


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


def _products_for_selection(sheets: dict, cfg: dict, crop_id, target_id=None,
                             stage_filter: str = None) -> set:
    """Product IDs (the master sheet's own id values — e.g. her_id) linked
    to this crop, via the category's junction sheet — optionally
    narrowed to one specific target and/or (Herbicide only) one spray
    timing. target_id=None means 'any target for this crop/category' —
    still respecting stage_filter if given, checked across every weed
    at that timing, not just one. Returns an empty set if the junction
    sheet is missing/empty or lacks the needed columns — callers treat
    that as 'nothing matches', same as any other missing-data case."""
    junction = sheets.get(cfg["junction"], pd.DataFrame())
    j_id = cfg["junction_id"]
    if junction.empty or "crop_id" not in junction.columns or j_id not in junction.columns:
        return set()

    j = junction[junction["crop_id"] == crop_id]
    if target_id is not None:
        j = j[j[cfg["target_id_col"]] == target_id]

    stage_col = cfg.get("stage_col")
    if stage_col and stage_filter and "ws_id" in j.columns:
        window_df = sheets.get(cfg["window_sheet"], pd.DataFrame())
        if not window_df.empty and "ws_id" in window_df.columns and stage_col in window_df.columns:
            w = window_df[
                (window_df["crop_id"] == crop_id)
                & (window_df[stage_col].astype(str).str.strip() == stage_filter)
            ]
            if target_id is not None:
                w = w[w[cfg["target_id_col"]] == target_id]
            valid_ws_ids = set(w["ws_id"].dropna())
            j = j[j["ws_id"].isin(valid_ws_ids)]

    return set(j[j_id].astype(str).str.strip())


def _render_recent_changes(sheets: dict, price_history: pd.DataFrame):
    st.caption(
        "What changed between the two most recent price-update rounds, "
        "per product — sorted by biggest movers first. Only products "
        "with at least 2 recorded snapshots show up here; a product "
        "with just one price entry so far has nothing to compare yet."
    )

    table = _price_movement_table(price_history)
    if table.empty:
        st.info("No product has at least 2 recorded price snapshots yet — "
                 "check back after the next update round.")
        st.stop()

    # --- Crop -> Category -> (Spray Timing, Herbicide only) -> Target ---
    # These four only narrow anything once BOTH crop and category are
    # specific — price_history has no crop/pest linkage of its own, so
    # this cross-references product_id against the junction sheets
    # price_view.py already uses for the same purpose. Left at "All",
    # this section is a no-op and the table behaves exactly as before.
    stage_df_all = sheets.get("crop_stage", pd.DataFrame())
    crop_names = sorted(stage_df_all["crop"].dropna().astype(str).unique().tolist()) \
        if not stage_df_all.empty and "crop" in stage_df_all.columns else []
    crop_lookup = dict(zip(stage_df_all.get("crop", []), stage_df_all.get("crop_id", [])))

    col_a, col_b, col_lang = st.columns([2, 2, 1])
    with col_a:
        crop_choice = st.selectbox("Crop", ["All"] + crop_names, key="pm_crop")
    with col_b:
        category_choice = st.selectbox("Category", CATEGORY_OPTIONS, key="pm_category")
    with col_lang:
        lang_choice = st.radio("Name language", ["English", "Thai"], horizontal=True, key="pm_lang")

    target_choice = "All"
    stage_filter = None
    cfg = None
    target_name_to_id = {}
    crop_id = crop_lookup.get(crop_choice) if crop_choice != "All" else None

    if crop_choice != "All" and category_choice != "All":
        cfg = PRICE_CATEGORY_CONFIG[_CATEGORY_TO_CONFIG_KEY[category_choice]]
        stage_options = _price_stage_options(sheets, cfg, crop_id) if category_choice == "Herbicide" else []

        if stage_options:
            col_c, col_d = st.columns(2)
            with col_c:
                stage_choice = st.selectbox("Spray Timing", ["All"] + stage_options, key="pm_stage")
                stage_filter = None if stage_choice == "All" else stage_choice
            target_slot = col_d
        else:
            target_slot = st

        targets_df = _price_target_options(sheets, cfg, crop_id, stage_filter=stage_filter)
        # Falls back to English if the Thai name column isn't present in
        # this sheet yet (or vice versa) — same defensive pattern used
        # everywhere else names are shown bilingually in this app.
        preferred_col = "name_en" if lang_choice == "English" else "name_th"
        name_col = preferred_col if preferred_col in targets_df.columns else (
            "name_en" if "name_en" in targets_df.columns else (
                targets_df.columns[0] if not targets_df.empty else "name_en"))
        target_options = targets_df[name_col].dropna().astype(str).tolist() if not targets_df.empty else []
        target_name_to_id = dict(zip(targets_df.get(name_col, []), targets_df.get(cfg["target_id_col"], [])))

        target_choice = target_slot.selectbox(
            "Target (pest/weed/disease)", ["All"] + target_options, key="pm_target"
        )

    # Company options come from every company with at least one recorded
    # price change (the unfiltered `table`), not from `filtered` — kept
    # independent of the other filters below, same as Category doesn't
    # narrow based on Direction. Cascading filters (where one dropdown's
    # options shift based on another) would be more surprising here
    # than useful.
    company_options = ["All"] + sorted(table["company"].dropna().astype(str).str.strip().unique().tolist())
    col_e, col_f = st.columns([2, 1])
    with col_e:
        company_choice = st.selectbox("Company", company_options, key="pm_company")
    with col_f:
        direction_choice = st.radio("Direction", ["All", "Increases only", "Decreases only"],
                                     horizontal=True, key="pm_direction")

    # Spray Timing only has a real value when Herbicide is selected —
    # shown as "N/A" otherwise rather than "All", so it doesn't imply a
    # timing concept exists for Insecticide/Fungicide (it doesn't).
    if category_choice == "Herbicide":
        spray_timing_display = stage_filter if stage_filter else "All"
    else:
        spray_timing_display = "N/A"

    st.markdown(
        f"**Crop:** {crop_choice} &nbsp;|&nbsp; "
        f"**Category:** {category_choice} &nbsp;|&nbsp; "
        f"**Spray Timing:** {spray_timing_display} &nbsp;|&nbsp; "
        f"**Company:** {company_choice}"
    )

    filtered = table.copy()
    if category_choice != "All":
        filtered = filtered[filtered["category"] == category_choice]

    if crop_choice != "All" and category_choice != "All":
        target_id = target_name_to_id.get(target_choice) if target_choice != "All" else None
        product_ids = _products_for_selection(sheets, cfg, crop_id, target_id=target_id, stage_filter=stage_filter)
        filtered = filtered[filtered["product_id"].isin(product_ids)]

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


def _search_price_history_by_name(price_history: pd.DataFrame, search_term: str) -> pd.DataFrame:
    """Case-insensitive substring search on price_history's common_name
    column — returns every matching ROW (every snapshot, not just the
    latest), since building a trend line needs the full history, not
    just a current value. Same search style as price_view.py's By
    Chemical Name mode, applied here to price_history instead of the
    master sheets."""
    if price_history.empty or "common_name" not in price_history.columns or not search_term.strip():
        return pd.DataFrame()
    term = search_term.strip().lower()
    mask = price_history["common_name"].astype(str).str.lower().str.contains(term, na=False, regex=False)
    return price_history[mask].copy()


def _render_price_trend_chart(price_history: pd.DataFrame):
    st.caption(
        "Search a chemical/active-ingredient name (e.g. 'copper', "
        "'glyphosate') to see how its price has moved over time — one "
        "line per product, so a company selling two different "
        "formulations of the same chemical gets two separate lines "
        "rather than one misleadingly averaged line. Shows the raw "
        "listed price (฿ per package as recorded), not a per-rai/per-"
        "20L cost — that stays comparable across categories with no "
        "unit-mixing risk, unlike the derived cost metrics used "
        "elsewhere in this app."
    )
    search_term = st.text_input(
        "Search by chemical / common name", key="pt_chart_search",
        placeholder="e.g. copper, glyphosate, mancozeb",
    )
    if not search_term.strip():
        st.info("Type a chemical or active-ingredient name above to search.")
        return

    matches = _search_price_history_by_name(price_history, search_term)
    if matches.empty:
        st.info(f"No price history found matching '{search_term}'.")
        return

    # One identity per (product_id, category, company, trade_name) --
    # category is part of the identity in case a product_id were ever
    # reused across categories, same safeguard used elsewhere.
    identity_cols = ["product_id", "category", "company", "trade_name", "common_name"]
    missing_identity_cols = [c for c in identity_cols if c not in matches.columns]
    if missing_identity_cols:
        st.error(f"price_history is missing expected column(s): {', '.join(missing_identity_cols)}")
        return
    identities = matches[identity_cols].drop_duplicates().copy()
    identities["label"] = identities.apply(
        lambda r: f"{r['company']} — {r['trade_name']} ({r['common_name']}, {r['category']})", axis=1
    )

    default_labels = identities["label"].tolist()[:8]  # cap the first view so it isn't overcrowded by default
    chosen_labels = st.multiselect(
        "Products to plot", identities["label"].tolist(), default=default_labels, key="pt_chart_products",
    )
    if not chosen_labels:
        st.info("Pick at least one product above to plot.")
        return

    fig = go.Figure()
    plotted_any = False
    for label in chosen_labels:
        id_row = identities[identities["label"] == label].iloc[0]
        rows = matches[
            (matches["product_id"].astype(str).str.strip() == str(id_row["product_id"]).strip())
            & (matches["category"] == id_row["category"])
        ].copy()
        rows["snapshot_date"] = pd.to_datetime(rows["snapshot_date"], errors="coerce", dayfirst=True)
        rows["price"] = pd.to_numeric(rows["price"], errors="coerce")
        rows = rows.dropna(subset=["snapshot_date", "price"]).sort_values("snapshot_date")
        if rows.empty:
            continue
        plotted_any = True
        fig.add_trace(go.Scatter(
            x=rows["snapshot_date"], y=rows["price"], mode="lines+markers", name=label,
        ))

    if not plotted_any:
        st.info("None of the selected products have usable price/date data to plot.")
        return

    fig.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis=dict(title="Snapshot date"),
        yaxis=dict(title="Price (฿ per package)"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("Raw snapshots for the plotted products"):
        chosen_ids = {(str(identities[identities["label"] == l].iloc[0]["product_id"]).strip(),
                       identities[identities["label"] == l].iloc[0]["category"]) for l in chosen_labels}
        raw = matches[matches.apply(
            lambda r: (str(r["product_id"]).strip(), r["category"]) in chosen_ids, axis=1
        )].copy()
        raw["snapshot_date"] = pd.to_datetime(raw["snapshot_date"], errors="coerce", dayfirst=True)
        raw = raw.sort_values(["company", "trade_name", "snapshot_date"])
        raw["price"] = raw["price"].apply(_format_price)
        display_cols = [c for c in ["snapshot_date", "category", "company", "trade_name",
                                     "common_name", "price"] if c in raw.columns]
        st.dataframe(raw[display_cols], use_container_width=True, hide_index=True)


def render_price_trend_view():
    st.title("📈 Price Movement")

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

    mode = st.radio(
        "View", ["Recent Changes", "Price Trend Chart"], horizontal=True, key="pm_mode",
        help="Recent Changes: what moved since the last update round, per product. "
             "Price Trend Chart: search a chemical and see its full price history as a line chart.",
    )
    st.divider()

    if mode == "Recent Changes":
        _render_recent_changes(sheets, price_history)
    else:
        _render_price_trend_chart(price_history)
