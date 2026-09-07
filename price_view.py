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
    does, alongside a 'best_efficiency' column used for filtering — see
    render_price_comparison_view's Minimum Efficiency control). ascending=True
    sorts cheapest first, False sorts priciest first; rows with no price
    data always sort to the bottom either way so they never get mistaken
    for the cheapest (or most expensive) option."""
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
        # Best rating this product has anywhere (a product can be
        # Effective against one window and Moderate against another —
        # for filtering purposes we judge it by its best showing, not
        # its worst). Used by the Minimum Efficiency filter below;
        # dropped before the table is displayed.
        best_efficiency = next((e for e in EFFICIENCY_ORDER if e in efficiencies_present), "Unrated")
        rows.append({
            "company": first.get("company", ""),
            "trade_name": first.get("trade_name", ""),
            "common_name": first.get("common_name", ""),
            "tier": normalize_tier(first.get("tier")),
            "efficiency": efficiency_display,
            "best_efficiency": best_efficiency,
            "price": first.get("price"),
            "size": first.get("size"),
            "usage": first.get("usage"),
            cfg["cost_col"]: first.get(cfg["cost_col"]),
        })
    out = pd.DataFrame(rows, columns=empty_cols + ["best_efficiency"])
    out["_sort_cost"] = pd.to_numeric(out[cfg["cost_col"]], errors="coerce")
    out = out.sort_values("_sort_cost", ascending=ascending, na_position="last").drop(columns=["_sort_cost"])
    return out.reset_index(drop=True)


def _all_companies(sheets: dict) -> list:
    """Union of every company appearing anywhere across the three
    master sheets (prod_her/prod_ins/prod_fun), regardless of which
    crop/target/mode is currently selected — used to populate the
    top-level 'Highlight companies' picker so it stays stable and
    available no matter what else on the page is being filtered."""
    companies = set()
    for cfg in PRICE_CATEGORY_CONFIG.values():
        master = sheets.get(cfg["master"], pd.DataFrame())
        if master.empty or "company" not in master.columns:
            continue
        companies.update(str(c).strip() for c in master["company"].dropna() if str(c).strip())
    return sorted(companies)


def _companies_for_category(sheets: dict, cfg: dict) -> list:
    """Companies appearing in just ONE category's master sheet — used
    by Portfolio Cost mode, which is scoped to a single category, so
    offering companies from other categories would just be noise (they
    couldn't have a product to contribute to this comparison anyway)."""
    master = sheets.get(cfg["master"], pd.DataFrame())
    if master.empty or "company" not in master.columns:
        return []
    return sorted({str(c).strip() for c in master["company"].dropna() if str(c).strip()})


def _portfolio_cost_table(sheets: dict, cfg: dict, crop_id, target_ids: list,
                           target_names: dict, companies: list):
    """For each selected company, sums the CHEAPEST available product's
    cost across all selected targets — a company's own total is only
    ever built from real per-target lookups already used elsewhere
    (_price_comparison_table), so this never invents a price. A company
    missing a product for one or more targets still gets a partial
    total (sum of whatever it does cover) plus a 'missing' count/list,
    rather than being silently dropped or blocked from showing anything
    at all — a real gap is exactly the kind of thing this view exists
    to surface. NOT for combining categories: units differ (per rai vs
    per 20L tank), so target_ids must all come from the SAME category.

    IMPORTANT — one product commonly covers MULTIPLE targets (a
    broad-spectrum product handling several pests). If it's picked as
    the cheapest option for more than one selected target, its cost is
    counted ONCE in the total, not once per target — you'd only
    actually buy it once. Deduplication is by (company, trade_name),
    since that's the identity already available from
    _price_comparison_table's output.

    Returns (summary_df, breakdown_df). breakdown_df is long-format,
    one row per (company, target) — the per-target detail behind each
    summary total, for a line-item drill-down. When a product repeats
    across targets for the same company, only its FIRST occurrence
    (in target order) carries a cost; later occurrences show cost=None
    with a note explaining it's the same product already counted —
    this way, naively summing breakdown_df's cost column for one
    company matches that company's total_cost exactly, rather than
    silently double-counting shared products."""
    # Reuses the existing per-target lookup (_price_comparison_table)
    # rather than re-deriving costs — one call per target, keeping each
    # company's cheapest PRODUCT (not just its cost) for that target, so
    # the same product picked for multiple targets can be recognized and
    # deduped below.
    per_target_company_pick = {}  # target_id -> {company: (trade_name, cost)}
    for tid in target_ids:
        t = _price_comparison_table(sheets, cfg, crop_id, tid, ascending=True)
        if t.empty:
            per_target_company_pick[tid] = {}
            continue
        t = t.copy()
        t["_cost_num"] = pd.to_numeric(t[cfg["cost_col"]], errors="coerce")
        t_valid = t.dropna(subset=["_cost_num"])
        picks = {}
        for company, g in t_valid.groupby("company"):
            best_row = g.loc[g["_cost_num"].idxmin()]
            picks[company] = (best_row["trade_name"], best_row["_cost_num"])
        per_target_company_pick[tid] = picks

    n_targets = len(target_ids)
    rows = []
    breakdown_rows = []
    for company in companies:
        products_used = {}  # trade_name -> cost, dict naturally dedupes
        covered = 0
        missing = []
        seen_products = set()
        for tid in target_ids:
            tname = target_names.get(tid, str(tid))
            pick = per_target_company_pick.get(tid, {}).get(company)
            if pick is not None:
                trade_name, cost = pick
                products_used[trade_name] = cost
                covered += 1
                if trade_name in seen_products:
                    breakdown_rows.append({
                        "company": company, "target": tname, "product": trade_name,
                        "cost": None, "note": "Same product as above — already counted once",
                    })
                else:
                    seen_products.add(trade_name)
                    breakdown_rows.append({
                        "company": company, "target": tname, "product": trade_name,
                        "cost": cost, "note": "",
                    })
            else:
                missing.append(tname)
                breakdown_rows.append({
                    "company": company, "target": tname, "product": None,
                    "cost": None, "note": "No product for this target — gap",
                })
        total = sum(products_used.values()) if products_used else None
        rows.append({
            "company": company,
            "covered": covered,
            "total_targets": n_targets,
            "total_cost": total,
            "products_used": "; ".join(sorted(products_used.keys())),
            "missing_targets": "; ".join(missing) if missing else "",
        })
    out = pd.DataFrame(rows)
    breakdown_df = pd.DataFrame(breakdown_rows)
    if out.empty:
        return out, breakdown_df
    # Fully-covered companies first (cheapest first among those), then
    # partial coverage (cheapest partial total first), matching the
    # instinct that "cheapest AND complete" beats "cheapest but missing
    # something" — a low total that hides a real gap shouldn't outrank
    # a slightly higher, fully-covered one.
    out["_fully_covered"] = out["covered"] == out["total_targets"]
    out = out.sort_values(
        ["_fully_covered", "total_cost"], ascending=[False, True], na_position="last"
    ).drop(columns=["_fully_covered"])
    return out.reset_index(drop=True), breakdown_df


# Soft, visible-but-not-garish highlight — distinct from the app's
# other colors (red/green coverage status, tier/efficiency badges) so
# it doesn't get visually confused with any of those.
_HIGHLIGHT_BG = "#FFF3B0"
# Force dark text on the highlighted background regardless of the
# viewer's Streamlit theme — without this, a dark-mode user's default
# white cell text becomes invisible against the light yellow highlight.
_HIGHLIGHT_TEXT = "#1A1A1A"


def _highlight_companies(display_df: pd.DataFrame, company_col: str, highlight_list: list):
    """Returns a pandas Styler that shades entire rows for any company
    in highlight_list, or the plain DataFrame unchanged if nothing's
    selected — st.dataframe renders either one correctly, so callers
    don't need an if/else at the call site. All selected companies
    share one highlight color (a group of 'companies I care about'),
    not a distinct color per company — simpler to read at a glance,
    and avoids picking an arbitrary color palette."""
    if not highlight_list:
        return display_df

    def _row_style(row):
        if row[company_col] in highlight_list:
            return [f"background-color: {_HIGHLIGHT_BG}; color: {_HIGHLIGHT_TEXT}"] * len(row)
        return [""] * len(row)

    return display_df.style.apply(_row_style, axis=1)


def _search_products_by_name(sheets: dict, search_term: str) -> pd.DataFrame:
    """Searches the common_name column directly on ALL THREE master
    sheets (prod_her/prod_ins/prod_fun) at once — a case-insensitive
    substring match, e.g. 'copper' finds every copper-based product
    regardless of which crop or pest it's linked to. No crop/target
    scoping needed here at all: unlike the By Target flow, company/
    trade_name/tier/price are all master-sheet-level attributes with no
    crop dimension, so this is a much simpler, flatter lookup. Returns
    one row per matching product, tagged with which category
    (Herbicide/Insecticide/Fungicide) it came from."""
    empty_cols = ["category", "company", "trade_name", "common_name", "tier",
                  "price", "size", "usage", "cost", "cost_unit_label"]
    if not search_term or not search_term.strip():
        return pd.DataFrame(columns=empty_cols)

    term = search_term.strip().lower()
    rows = []
    for category_choice, cfg in PRICE_CATEGORY_CONFIG.items():
        master = sheets.get(cfg["master"], pd.DataFrame())
        if master.empty or "common_name" not in master.columns:
            continue
        matches = master[master["common_name"].astype(str).str.lower().str.contains(term, na=False, regex=False)]
        category_label = category_choice.split(" ")[0]  # "Herbicide (Weed)" -> "Herbicide"
        for _, r in matches.iterrows():
            rows.append({
                "category": category_label,
                "company": r.get("company", ""),
                "trade_name": r.get("trade_name", ""),
                "common_name": r.get("common_name", ""),
                "tier": normalize_tier(r.get("tier")),
                "price": r.get("price"),
                "size": r.get("size"),
                "usage": r.get("usage"),
                "cost": r.get(cfg["cost_col"]),
                "cost_unit_label": cfg["cost_unit_label"],
            })
    if not rows:
        return pd.DataFrame(columns=empty_cols)
    # Sort by common_name first so identical/similar active ingredients
    # cluster together (the main point of this search — comparing the
    # same chemical across companies) — then by category, since cost
    # units differ across categories and shouldn't be interleaved as if
    # comparable.
    out = pd.DataFrame(rows, columns=empty_cols)
    return out.sort_values(["common_name", "category"]).reset_index(drop=True)


def _render_by_target(sheets: dict, highlight_companies: list):
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

    # Minimum Efficiency filter — lets weak products be excluded from
    # the price comparison entirely, rather than sitting in the table
    # next to genuinely good options. Judges each product by its BEST
    # rating anywhere (a product Effective against one window and
    # Moderate against another still counts as Effective here), not its
    # worst. Unrated is handled separately from the threshold itself,
    # since "not yet assessed" isn't the same claim as "confirmed weak".
    #
    # Minimum Tier filter sits alongside it — simpler than efficiency
    # since tier is a fixed product-level attribute (from the master
    # sheet) that doesn't vary by window/target, so there's no "mix" or
    # "best of" to resolve, and no Unrated concept (normalize_tier
    # always resolves to a real tier, defaulting unknowns to Generic).
    col6, col7, col8 = st.columns([2, 1, 2])
    with col6:
        min_eff_choice = st.selectbox(
            "Minimum Efficiency", ["All", "Moderate or better", "Effective only"],
            key="price_min_eff",
            help="Filters out products whose best rating anywhere falls below this bar.",
        )
    with col7:
        include_unrated = st.checkbox("Include Unrated", value=True, key="price_include_unrated")
    with col8:
        min_tier_choice = st.selectbox(
            "Minimum Tier", ["All", "Medium or better", "Premium only"],
            key="price_min_tier",
        )

    eff_thresholds = {
        "Moderate or better": {"Effective", "Moderate"},
        "Effective only": {"Effective"},
    }
    if min_eff_choice in eff_thresholds:
        mask = table["best_efficiency"].isin(eff_thresholds[min_eff_choice])
        if include_unrated:
            mask = mask | (table["best_efficiency"] == "Unrated")
        table = table[mask]
    elif not include_unrated:
        table = table[table["best_efficiency"] != "Unrated"]

    tier_thresholds = {
        "Medium or better": {"Premium", "Medium"},
        "Premium only": {"Premium"},
    }
    if min_tier_choice in tier_thresholds:
        table = table[table["tier"].isin(tier_thresholds[min_tier_choice])]

    if table.empty:
        st.info(f"No products meet these filters for {target_choice}.")
        st.stop()

    display = table.drop(columns=["best_efficiency"]).copy()
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
    st.dataframe(_highlight_companies(display, "Company", highlight_companies),
                 use_container_width=True, hide_index=True)


def _render_by_chemical_name(sheets: dict, highlight_companies: list):
    st.caption(
        "Search across ALL crops and categories at once — e.g. type "
        "'copper' to see every copper-based product any company sells, "
        "regardless of which crop or pest it's linked to."
    )
    search_term = st.text_input(
        "Search by chemical / common name", key="price_chem_search",
        placeholder="e.g. copper, glyphosate, mancozeb",
    )

    min_tier_choice = st.selectbox(
        "Minimum Tier", ["All", "Medium or better", "Premium only"],
        key="price_chem_min_tier",
    )

    if not search_term.strip():
        st.info("Type a chemical or active-ingredient name above to search.")
        return

    results = _search_products_by_name(sheets, search_term)
    if results.empty:
        st.info(f"No products found matching '{search_term}'.")
        return

    tier_thresholds = {
        "Medium or better": {"Premium", "Medium"},
        "Premium only": {"Premium"},
    }
    if min_tier_choice in tier_thresholds:
        results = results[results["tier"].isin(tier_thresholds[min_tier_choice])]
    if results.empty:
        st.info(f"No products matching '{search_term}' meet this tier filter.")
        return

    display = results.copy()
    display["price"] = display["price"].apply(_format_price)
    display["cost"] = display.apply(
        lambda r: _format_price(r["cost"]) + f"/{r['cost_unit_label']}" if _format_price(r["cost"]) else "—",
        axis=1,
    )
    display = display.drop(columns=["cost_unit_label"]).rename(columns={
        "category": "Category", "company": "Company", "trade_name": "Trade Name",
        "common_name": "Common Name", "tier": "Tier", "price": "Price",
        "size": "Size", "usage": "Usage", "cost": "Cost",
    })
    st.caption(
        "Note: the Cost column's unit differs by category — per rai "
        "(Herbicide) vs per 20L tank (Insecticide/Fungicide) — never "
        "compare Cost values across different Category rows directly."
    )
    st.subheader(f"'{search_term}' — {len(display)} product(s) across "
                 f"{display['Company'].nunique()} company(ies)")
    st.dataframe(_highlight_companies(display, "Company", highlight_companies),
                 use_container_width=True, hide_index=True)


def _render_portfolio_cost(sheets: dict, highlight_companies: list):
    st.caption(
        "Pick several targets in ONE category — see each company's total "
        "cost to cover all of them (using their cheapest product per "
        "target), plus a flag for anyone missing coverage on one or more. "
        "If one product covers multiple selected targets, its cost is "
        "only counted once, not once per target. Note: this can't yet mix "
        "categories into one number, since ฿/rai (Herbicide) and ฿/20L "
        "tank (Insecticide/Fungicide) aren't the same unit — pick one "
        "category per comparison for now."
    )

    stage_df_all = sheets["crop_stage"]
    if stage_df_all.empty:
        st.error("`crop_stage` sheet is missing or empty.")
        st.stop()
    crop_lookup = stage_df_all[["crop_id", "crop"]].drop_duplicates()
    crop_name_to_id = dict(zip(crop_lookup["crop"], crop_lookup["crop_id"]))

    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()), key="pf_crop")
    crop_id = crop_name_to_id[crop_choice]
    with col2:
        category_choice = st.selectbox("Category", list(PRICE_CATEGORY_CONFIG.keys()), key="pf_category")
    with col3:
        lang_choice = st.radio("Name language", ["English", "Thai"], horizontal=True, key="pf_lang")
    cfg = PRICE_CATEGORY_CONFIG[category_choice]
    name_col = "name_en" if lang_choice == "English" else "name_th"

    # Spray Timing (e.g. Pre-emergence / Early Post / Late Post) only
    # exists for Herbicide, same as By Target mode — "All" pools every
    # timing's targets together into one multiselect; picking a specific
    # timing narrows the target list to just that timing's weeds.
    stage_options = _price_stage_options(sheets, cfg, crop_id)
    stage_filter = None
    if stage_options:
        stage_choice = st.selectbox(
            "Spray Timing", ["All"] + stage_options, key="pf_stage"
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
    target_choices = st.multiselect(
        f"{category_choice.split(' ')[0]} targets to include",
        list(target_name_to_id.keys()), key="pf_targets",
    )
    if not target_choices:
        st.info("Pick at least one target above to build a portfolio comparison.")
        return
    target_ids = [target_name_to_id[t] for t in target_choices]
    target_names = {tid: name for name, tid in target_name_to_id.items() if tid in target_ids}

    category_companies = _companies_for_category(sheets, cfg)
    if not category_companies:
        st.info(f"No companies found for {category_choice.lower()} on this crop.")
        st.stop()
    companies_choice = st.multiselect(
        "Companies to compare", category_companies, key="pf_companies",
        default=[c for c in highlight_companies if c in category_companies] or None,
        help="Defaults to whatever you picked in 'Highlight companies' above, if applicable.",
    )
    if not companies_choice:
        st.info("Pick at least one company above to compare.")
        return

    table, breakdown_df = _portfolio_cost_table(sheets, cfg, crop_id, target_ids, target_names, companies_choice)
    if table.empty:
        st.info("No data to show for this selection.")
        return

    display = table.copy()
    display["Coverage"] = display.apply(
        lambda r: f"{r['covered']}/{r['total_targets']}"
        + (" ⚠️" if r["covered"] < r["total_targets"] else " ✅"),
        axis=1,
    )
    display["Total Cost"] = display["total_cost"].apply(
        lambda v: (_format_price(v) + f"/{cfg['cost_unit_label']} (total)") if v is not None and not pd.isna(v) else "—"
    )
    display = display.rename(columns={
        "company": "Company", "products_used": "Products Used",
        "missing_targets": "Missing Targets",
    })
    display = display[["Company", "Coverage", "Total Cost", "Products Used", "Missing Targets"]]

    st.subheader(f"Portfolio cost across {len(target_ids)} {category_choice.split(' ')[0].lower()} target(s)")
    st.dataframe(_highlight_companies(display, "Company", highlight_companies),
                 use_container_width=True, hide_index=True)

    # Per-company drill-down: which target maps to which product, and
    # what it costs — expanders (not another dataframe column) since
    # st.dataframe can't show a nested/expandable detail per row.
    # Ordered the same as the summary table above, so "cheapest first"
    # still applies to which expander you'd naturally open first.
    st.caption("Expand a company below to see exactly which product covers which target.")
    for _, row in table.iterrows():
        company = row["company"]
        with st.expander(f"📋 {company} — {row['covered']}/{row['total_targets']} covered"):
            company_detail = breakdown_df[breakdown_df["company"] == company].copy()
            company_detail["cost"] = company_detail["cost"].apply(
                lambda v: (_format_price(v) + f"/{cfg['cost_unit_label']}") if v is not None and not pd.isna(v) else "—"
            )
            company_detail["product"] = company_detail["product"].fillna("— (no product)")
            company_detail = company_detail.rename(columns={
                "target": "Target", "product": "Product", "cost": "Cost", "note": "Note",
            })[["Target", "Product", "Cost", "Note"]]
            st.dataframe(company_detail, use_container_width=True, hide_index=True)
            total_display = (
                _format_price(row["total_cost"]) + f"/{cfg['cost_unit_label']} (total)"
                if row["total_cost"] is not None and not pd.isna(row["total_cost"]) else "—"
            )
            st.caption(f"**Total: {total_display}** — adds up the Cost column above, "
                       "skipping rows already counted via a shared product.")


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

    mode = st.radio(
        "Compare mode", ["By Target", "By Chemical Name", "Portfolio Cost"], horizontal=True, key="price_mode",
        help="By Target: pick a crop + pest, compare companies. "
             "By Chemical Name: search a chemical (e.g. copper) across everything at once. "
             "Portfolio Cost: total cost across several targets at once, per company.",
    )

    # Shared across both modes via the same widget key, so picking your
    # company(ies) once carries over whether you're in By Target or By
    # Chemical Name — no need to reselect when switching. Sourced from
    # every company across all three master sheets, not just whatever's
    # currently filtered/visible, so the picker stays stable regardless
    # of crop/category/search selections made elsewhere on the page.
    all_companies = _all_companies(sheets)
    highlight_companies = st.multiselect(
        "Highlight companies", all_companies, key="price_highlight_companies",
        help="Selected companies' rows are shaded in the results table below — "
             "e.g. highlight your own company, or a couple of key competitors.",
    )
    st.divider()

    if mode == "By Target":
        _render_by_target(sheets, highlight_companies)
    elif mode == "By Chemical Name":
        _render_by_chemical_name(sheets, highlight_companies)
    else:
        _render_portfolio_cost(sheets, highlight_companies)
