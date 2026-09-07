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
import plotly.graph_objects as go

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

# Canonical spray-timing order — a weed's timings follow the natural
# agronomic sequence within a season (pre-emergence happens before any
# post-emergence treatment; early post before late post), so this
# should drive sorting/grouping wherever Spray Timing appears, not
# alphabetical order (which would incorrectly put "Early Post" before
# "Late Post" but "Late Post" before "Pre-emergence"). Aliases cover
# common real-world spelling variants; anything genuinely unrecognized
# still sorts in — just placed after the three known stages, rather
# than being dropped or crashing.
SPRAY_TIMING_ORDER = ["Pre-emergence", "Early Post", "Late Post"]
_SPRAY_TIMING_ALIASES = {
    "pre-emergence": "Pre-emergence", "pre emergence": "Pre-emergence",
    "preemergence": "Pre-emergence", "pre": "Pre-emergence",
    "early post": "Early Post", "early post-emergence": "Early Post",
    "early postemergence": "Early Post", "early-post": "Early Post", "early": "Early Post",
    "late post": "Late Post", "late post-emergence": "Late Post",
    "late postemergence": "Late Post", "late-post": "Late Post", "late": "Late Post",
}


def _normalize_spray_timing(timing) -> str:
    """Maps a raw spray-timing string to its canonical label
    (Pre-emergence/Early Post/Late Post) when recognized; returns the
    original (stripped) text unchanged when it isn't a known variant,
    so unusual data still displays as-typed rather than being silently
    dropped or blanked."""
    if not timing:
        return ""
    raw = str(timing).strip()
    key = raw.lower().replace("_", "-").replace("  ", " ")
    return _SPRAY_TIMING_ALIASES.get(key, raw)


def _spray_timing_sort_key(timing):
    """Sort key placing timings in the canonical Pre-emergence -> Early
    Post -> Late Post order. Blank/None sorts into the same bucket as
    genuinely unrecognized values — after the three known stages,
    alphabetically among themselves — rather than erroring or randomly
    interleaving with the known sequence."""
    canonical = _normalize_spray_timing(timing)
    if canonical in SPRAY_TIMING_ORDER:
        return (0, SPRAY_TIMING_ORDER.index(canonical))
    return (1, canonical)


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
    values = {str(v).strip() for v in w[stage_col].dropna() if str(v).strip()}
    return sorted(values, key=_spray_timing_sort_key)


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


def _target_ws_ids(sheets: dict, cfg: dict, crop_id, target_id, stage_filter: str = None):
    """Herbicide only (cfg has 'stage_col'): resolves which specific
    ws_id window(s) correspond to this weed under the CURRENT Spray
    Timing filter, so the price lookup can be scoped to just that
    timing rather than silently aggregating every timing regardless of
    what was selected — that silent aggregation was the actual bug
    behind 'All' and a specific timing showing identical numbers.
    Returns None (no restriction — every window for this weed is
    already what the lookup uses without this) when there's no stage
    concept at all (Insect/Disease) or stage_filter is None/'All'."""
    stage_col = cfg.get("stage_col")
    if not stage_col or not stage_filter:
        return None
    window_df = sheets.get(cfg["window_sheet"], pd.DataFrame())
    if window_df.empty or "ws_id" not in window_df.columns:
        return None
    w = window_df[
        (window_df["crop_id"] == crop_id)
        & (window_df[cfg["target_id_col"]] == target_id)
        & (window_df[stage_col].astype(str).str.strip() == stage_filter)
    ]
    return w["ws_id"].dropna().unique().tolist()


def _price_comparison_table(sheets: dict, cfg: dict, crop_id, target_id,
                             ascending: bool = True, ws_ids: list = None) -> pd.DataFrame:
    """Every company's product linked to this target. ws_ids, when given
    (Herbicide only — see _target_ws_ids), restricts this to just the
    window(s) matching the current Spray Timing filter; None means no
    restriction (every window for this target, i.e. 'All' timing).

    For Herbicide specifically (cfg has 'stage_col'), rows are grouped
    by (product, spray timing) rather than by product alone — the SAME
    product can have different efficiency at different timings (tied to
    a specific ws_id in weed_her), so blending them into one row would
    hide that difference. This means with 'All' selected, a product
    used at both Early Post and Late Post shows as two separate rows,
    each tagged with its own Spray Timing — that's what actually lets
    'All' genuinely combine every timing instead of quietly picking
    just one. Insect/Disease have no stage concept, so they keep the
    original one-row-per-product behavior untouched, no Spray Timing
    column at all.

    price/size/usage/tier are product-level attributes and don't vary
    by window, only efficiency (and, for Herbicide, spray_timing) might.
    A 'best_efficiency' column is also included, used for filtering —
    see render_price_comparison_view's Minimum Efficiency control.
    ascending=True sorts cheapest first, False sorts priciest first;
    rows with no price data always sort to the bottom either way so
    they never get mistaken for the cheapest (or most expensive)
    option."""
    junction = sheets.get(cfg["junction"], pd.DataFrame())
    master = sheets.get(cfg["master"], pd.DataFrame())
    j_id, m_id = cfg["junction_id"], cfg["master_id"]
    has_stage = bool(cfg.get("stage_col")) and "ws_id" in junction.columns
    base_cols = ["company", "trade_name", "common_name", "tier", "efficiency",
                 "price", "size", "usage", cfg["cost_col"]]
    empty_cols = base_cols + (["spray_timing"] if has_stage else [])
    if junction.empty or master.empty or j_id not in junction.columns or m_id not in master.columns:
        return pd.DataFrame(columns=empty_cols)

    j = junction[(junction["crop_id"] == crop_id) & (junction[cfg["target_id_col"]] == target_id)].copy()
    if ws_ids is not None and "ws_id" in j.columns:
        j = j[j["ws_id"].isin(ws_ids)]
    if j.empty:
        return pd.DataFrame(columns=empty_cols)

    # Map each junction row's ws_id to its spray-timing label (from the
    # window sheet, e.g. crop_weeds) so it can be shown per row and used
    # as part of the grouping key below.
    if has_stage:
        window_df = sheets.get(cfg["window_sheet"], pd.DataFrame())
        stage_col = cfg["stage_col"]
        if not window_df.empty and "ws_id" in window_df.columns and stage_col in window_df.columns:
            ws_to_stage = window_df.drop_duplicates(subset=["ws_id"]).set_index("ws_id")[stage_col].to_dict()
            j["spray_timing"] = j["ws_id"].map(ws_to_stage)
        else:
            j["spray_timing"] = None

    overlap = [c for c in j.columns if c in master.columns and c != j_id]
    j = j.drop(columns=overlap)
    merged = j.merge(master, left_on=j_id, right_on=m_id, how="left") if j_id != m_id \
        else j.merge(master, on=j_id, how="left")
    if merged.empty or "company" not in merged.columns:
        return pd.DataFrame(columns=empty_cols)

    rows = []
    group_cols = [m_id, "spray_timing"] if has_stage else [m_id]
    for _, g in merged.groupby(group_cols, dropna=False):
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
        row = {
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
        }
        if has_stage:
            row["spray_timing"] = first.get("spray_timing") or "—"
        rows.append(row)
    out = pd.DataFrame(rows, columns=empty_cols + ["best_efficiency"])
    out["_sort_cost"] = pd.to_numeric(out[cfg["cost_col"]], errors="coerce")
    # When multiple distinct timings are present (i.e. 'All' was
    # selected for Herbicide), group rows by timing FIRST — in the
    # canonical Pre-emergence -> Early Post -> Late Post order, not
    # alphabetically or interleaved by price — with cost as the
    # secondary sort within each timing group. A single-timing result
    # (a specific timing was chosen, or no stage concept at all) keeps
    # the original pure cost-based sort, since there's nothing to group.
    if has_stage and out["spray_timing"].nunique() > 1:
        out["_sort_timing"] = out["spray_timing"].apply(_spray_timing_sort_key)
        out = out.sort_values(
            ["_sort_timing", "_sort_cost"], ascending=[True, ascending], na_position="last"
        ).drop(columns=["_sort_timing", "_sort_cost"])
    else:
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


def _portfolio_treatment_windows(sheets: dict, cfg: dict, crop_id, target_ids: list,
                                  stage_filter: str = None) -> list:
    """Expands the selected targets into the actual list of REQUIRED
    treatment occasions to price out and sum.

    For categories with no timing concept (Insect/Disease), this is
    just one entry per target — unchanged from before.

    For Herbicide: if a SPECIFIC timing is selected, still one entry
    per target (scoped to that timing). If 'All' is selected, a weed
    needing treatment at BOTH Early Post and Late Post contributes TWO
    entries here — one per timing — since those are separate, mandatory
    spray occasions in the season, not alternative ways to cover the
    same need. 'All' must therefore SUM the cost of every timing a weed
    requires, not just pick whichever single timing happens to be
    cheapest — picking only the cheapest would silently skip a spray
    application the crop actually needs.

    Ordering: when 'All' is selected, the result is grouped by TIMING
    first (in the canonical Pre-emergence -> Early Post -> Late Post
    order), with every selected weed nested under each timing — i.e.
    all weeds needing a Pre-emergence spray, then all weeds needing an
    Early Post spray, then all needing Late Post — rather than grouped
    by weed first. This ordering is what every downstream table/
    breakdown inherits, since they all iterate this same list.

    Returns a list of (target_id, spray_timing_or_None) tuples."""
    stage_col = cfg.get("stage_col")
    if not stage_col:
        return [(tid, None) for tid in target_ids]
    window_df = sheets.get(cfg["window_sheet"], pd.DataFrame())
    if window_df.empty or stage_col not in window_df.columns:
        return [(tid, None) for tid in target_ids]

    if stage_filter:
        # A specific timing was chosen — one entry per weed that
        # actually has a window at that timing, no grouping decision
        # needed since every entry shares the same single timing.
        windows = []
        for tid in target_ids:
            w = window_df[
                (window_df["crop_id"] == crop_id)
                & (window_df[cfg["target_id_col"]] == tid)
                & (window_df[stage_col].astype(str).str.strip() == stage_filter)
            ]
            if not w.empty:
                windows.append((tid, stage_filter))
        return windows

    # 'All' — first work out which timing(s) each weed actually needs,
    # then re-group the whole list by TIMING (canonical order) with
    # every weed nested underneath, rather than by weed.
    per_target_timings = {}  # target_id -> set of timing labels
    for tid in target_ids:
        w = window_df[(window_df["crop_id"] == crop_id) & (window_df[cfg["target_id_col"]] == tid)]
        timings = {str(v).strip() for v in w[stage_col].dropna() if str(v).strip()}
        per_target_timings[tid] = timings

    all_timings_present = sorted(
        {t for timings in per_target_timings.values() for t in timings},
        key=_spray_timing_sort_key,
    )

    windows = []
    for timing in all_timings_present:
        for tid in target_ids:  # preserves the user's selection order within each timing group
            if timing in per_target_timings.get(tid, set()):
                windows.append((tid, timing))
    # Any weed with NO timing info at all (blank weed_stage) still gets
    # a single generic, timing-less entry, appended after every
    # recognized timing group rather than silently dropped.
    for tid in target_ids:
        if not per_target_timings.get(tid):
            windows.append((tid, None))
    return windows


def _portfolio_target_options(sheets: dict, cfg: dict, crop_id, treatment_windows: list) -> dict:
    """Every available product option per (company, target, timing)
    triple, sorted cheapest first — the raw material behind both the
    auto-cheapest default and the optional manual override (see
    'Customize product picks' in _render_portfolio_cost). Each entry in
    treatment_windows (from _portfolio_treatment_windows) gets its OWN
    scoped price lookup via _target_ws_ids, so a weed requiring two
    timings correctly gets two independent option lists — one per
    timing — rather than one pooled list that would let the cheaper
    timing's product silently substitute for the other."""
    options = {}  # (company, target_id, timing) -> [(trade_name, cost), ...] sorted by cost asc
    for target_id, timing in treatment_windows:
        ws_ids = _target_ws_ids(sheets, cfg, crop_id, target_id, timing)
        t = _price_comparison_table(sheets, cfg, crop_id, target_id, ascending=True, ws_ids=ws_ids)
        if t.empty:
            continue
        t = t.copy()
        t["_cost_num"] = pd.to_numeric(t[cfg["cost_col"]], errors="coerce")
        t_valid = t.dropna(subset=["_cost_num"])
        for company, g in t_valid.groupby("company"):
            pairs = sorted(zip(g["trade_name"], g["_cost_num"]), key=lambda p: p[1])
            options[(company, target_id, timing)] = pairs
    return options


def _default_picks(all_options: dict) -> dict:
    """Cheapest option per (company, target, timing) combo — used as-is
    when 'Customize product picks' is off, and as the starting point
    for any combo the user hasn't manually overridden when it's on."""
    return {key: opts[0] for key, opts in all_options.items() if opts}


def _resolve_portfolio(treatment_windows: list, target_names: dict, companies: list, picks: dict):
    """treatment_windows: list of (target_id, spray_timing_or_None) —
    the REQUIRED line items to price out and sum (see
    _portfolio_treatment_windows; under 'All' timing, one weed can
    contribute more than one entry here, one per timing it needs).
    picks: {(company, target_id, timing): (trade_name, cost)} — the
    resolved product for every combo that has one (auto-cheapest or a
    manual override); a combo absent from picks is a gap. Shared by
    both the default and 'Customize product picks' paths so the
    totals/breakdown logic can never diverge between the two.

    IMPORTANT — dedup rule: the SAME product picked for two DIFFERENT
    targets AT THE SAME TIMING is counted ONCE (one broad-spectrum
    product, one spray pass, covering multiple pests/weeds together).
    But the SAME product picked at DIFFERENT TIMINGS is counted
    SEPARATELY, even for the same target — Early Post and Late Post
    are different spray occasions in the season, so using the same
    product at both means buying/applying it twice. Dedup key is
    therefore (trade_name, timing), not trade_name alone.

    Returns (summary_df, breakdown_df). breakdown_df is long-format,
    one row per (target, timing) per company. When a product repeats
    across DIFFERENT TARGETS at the SAME timing, only its first
    occurrence (in window order) carries a cost; later ones show
    cost=None with a note — so naively summing breakdown_df's cost
    column for one company matches that company's total_cost exactly."""
    n_windows = len(treatment_windows)
    rows = []
    breakdown_rows = []
    for company in companies:
        total = 0.0
        covered = 0
        missing = []
        products_used = set()
        seen = set()  # (trade_name, timing) dedup key
        for target_id, timing in treatment_windows:
            tname = target_names.get(target_id, str(target_id))
            label = f"{tname} ({timing})" if timing else tname
            pick = picks.get((company, target_id, timing))
            if pick is not None:
                trade_name, cost = pick
                covered += 1
                dedup_key = (trade_name, timing)
                if dedup_key in seen:
                    breakdown_rows.append({
                        "company": company, "target": tname, "spray_timing": timing,
                        "product": trade_name, "cost": None,
                        "note": "Same product already counted for this timing",
                    })
                else:
                    seen.add(dedup_key)
                    total += cost
                    products_used.add(trade_name)
                    breakdown_rows.append({
                        "company": company, "target": tname, "spray_timing": timing,
                        "product": trade_name, "cost": cost, "note": "",
                    })
            else:
                missing.append(label)
                breakdown_rows.append({
                    "company": company, "target": tname, "spray_timing": timing,
                    "product": None, "cost": None, "note": "No product for this window — gap",
                })
        rows.append({
            "company": company,
            "covered": covered,
            "total_targets": n_windows,
            "total_cost": total if covered > 0 else None,
            "products_used": "; ".join(sorted(products_used)),
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

    # Scope the price lookup to the currently selected Spray Timing —
    # None (no restriction) when "All" is chosen, or the exact ws_id
    # window(s) for a specific timing otherwise. Without this, "All"
    # and any single timing would silently return identical results,
    # since the underlying junction lookup ignores timing unless told
    # which windows to restrict to.
    ws_ids = _target_ws_ids(sheets, cfg, crop_id, target_id, stage_filter)
    table = _price_comparison_table(sheets, cfg, crop_id, target_id,
                                     ascending=(sort_choice == "Cheapest first"), ws_ids=ws_ids)
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
        "spray_timing": "Spray Timing",
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
        "target/timing by default — see 'Customize product picks' below "
        "to choose a different one), plus a flag for anyone missing "
        "coverage on one or more. For Herbicide with 'All' timing "
        "selected, a weed needing BOTH Early Post and Late Post treatment "
        "adds BOTH as separate required costs (they're different spray "
        "occasions, not alternatives) — picking a single specific timing "
        "instead scopes the total to just that one spray. If the same "
        "product is used to cover multiple DIFFERENT targets at the SAME "
        "timing, its cost is only counted once (one spray pass); the "
        "same product used at DIFFERENT timings is counted separately. "
        "Note: this can't yet mix categories into one number, since "
        "฿/rai (Herbicide) and ฿/20L tank (Insecticide/Fungicide) aren't "
        "the same unit — pick one category per comparison for now."
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

    # Expands the chosen weeds into the actual required treatment
    # occasions — see _portfolio_treatment_windows. With a SPECIFIC
    # timing selected, this is one entry per weed (unchanged). With
    # 'All' selected (Herbicide only), a weed needing both Early Post
    # AND Late Post now correctly contributes TWO entries here, so the
    # total below sums every required spray rather than picking
    # whichever single timing happens to be cheapest.
    treatment_windows = _portfolio_treatment_windows(sheets, cfg, crop_id, target_ids, stage_filter)

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

    all_options = _portfolio_target_options(sheets, cfg, crop_id, treatment_windows)
    picks = _default_picks(all_options)

    # Customize product picks — off by default (auto-cheapest, exactly
    # as before). When on, only (company, target, timing) combos that
    # actually HAVE more than one product option get a dropdown — no
    # point showing a picker where there's nothing to choose between.
    # Picks made here are collected into `picks` and immediately feed
    # the summary/breakdown below, since Streamlit reruns top-to-bottom
    # on every widget interaction and a selectbox's return value
    # already reflects the latest choice at the point it's read.
    customize = st.checkbox(
        "Customize product picks", key="pf_customize",
        help="Off: automatically uses each company's cheapest product per target/timing. "
             "On: shows a dropdown (defaulting to cheapest) for any company/target/timing "
             "that has more than one product option, so you can pick a different one.",
    )
    if customize:
        any_adjustable = False
        for company in companies_choice:
            adjustable_windows = [
                (tid, timing) for tid, timing in treatment_windows
                if len(all_options.get((company, tid, timing), [])) > 1
            ]
            if not adjustable_windows:
                continue
            any_adjustable = True
            with st.expander(f"🔧 Adjust picks — {company}"):
                for tid, timing in adjustable_windows:
                    tname = target_names[tid]
                    row_label = f"{tname} — {timing}" if timing else tname
                    opts = all_options[(company, tid, timing)]
                    labels = [f"{tn} — {_format_price(c)}/{cfg['cost_unit_label']}" for tn, c in opts]
                    pick_key = f"pf_pick_{company}_{tid}_{timing}"
                    chosen_label = st.selectbox(row_label, labels, key=pick_key)
                    picks[(company, tid, timing)] = opts[labels.index(chosen_label)]
        if not any_adjustable:
            st.caption("No company/target/timing in this selection has more than one product option to choose between.")

    table, breakdown_df = _resolve_portfolio(treatment_windows, target_names, companies_choice, picks)
    if table.empty:
        st.info("No data to show for this selection.")
        return

    display = table.copy()
    display["Coverage"] = display.apply(
        lambda r: f"{r['covered']}/{r['total_targets']} windows"
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

    st.subheader(f"Portfolio cost across {len(treatment_windows)} required "
                 f"{category_choice.split(' ')[0].lower()} treatment window(s)")
    st.dataframe(_highlight_companies(display, "Company", highlight_companies),
                 use_container_width=True, hide_index=True)

    # Per-company drill-down: which target maps to which product, and
    # what it costs — expanders (not another dataframe column) since
    # st.dataframe can't show a nested/expandable detail per row.
    # Ordered the same as the summary table above, so "cheapest first"
    # still applies to which expander you'd naturally open first. Spray
    # Timing column only added for Herbicide (cfg has 'stage_col') —
    # Insect/Disease have no timing concept, so it'd just be a blank
    # column there.
    has_stage = bool(cfg.get("stage_col"))
    st.caption("Expand a company below to see exactly which product covers which target/timing.")
    for _, row in table.iterrows():
        company = row["company"]
        with st.expander(f"📋 {company} — {row['covered']}/{row['total_targets']} windows covered"):
            company_detail = breakdown_df[breakdown_df["company"] == company].copy()
            company_detail["cost"] = company_detail["cost"].apply(
                lambda v: (_format_price(v) + f"/{cfg['cost_unit_label']}") if v is not None and not pd.isna(v) else "—"
            )
            company_detail["product"] = company_detail["product"].fillna("— (no product)")
            rename_map = {"target": "Target", "product": "Product", "cost": "Cost", "note": "Note"}
            display_cols = ["Target", "Product", "Cost"]
            if has_stage:
                company_detail["spray_timing"] = company_detail["spray_timing"].fillna("—")
                rename_map["spray_timing"] = "Spray Timing"
                display_cols.append("Spray Timing")
            display_cols.append("Note")
            company_detail = company_detail.rename(columns=rename_map)[display_cols]
            st.dataframe(company_detail, use_container_width=True, hide_index=True)
            total_display = (
                _format_price(row["total_cost"]) + f"/{cfg['cost_unit_label']} (total)"
                if row["total_cost"] is not None and not pd.isna(row["total_cost"]) else "—"
            )
            st.caption(f"**Total: {total_display}** — adds up the Cost column above, "
                       "skipping rows already counted via a shared product.")


def _render_full_treatment_program(sheets: dict, highlight_companies: list):
    st.caption(
        "Compare a company's TOTAL cost across ALL THREE categories side "
        "by side for one crop, using every target in each category by "
        "default (for finer control over which specific targets count, "
        "use Portfolio Cost mode instead — this view is the quick "
        "full-picture version). Categories are shown as separate "
        "charts/tables since ฿/rai (Herbicide) and ฿/20L tank "
        "(Insecticide/Fungicide) aren't the same unit and can't be "
        "summed into one number — never compare bar heights ACROSS the "
        "three charts below, only within one chart."
    )

    stage_df_all = sheets["crop_stage"]
    if stage_df_all.empty:
        st.error("`crop_stage` sheet is missing or empty.")
        st.stop()
    crop_lookup = stage_df_all[["crop_id", "crop"]].drop_duplicates()
    crop_name_to_id = dict(zip(crop_lookup["crop"], crop_lookup["crop_id"]))

    crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()), key="ftp_crop")
    crop_id = crop_name_to_id[crop_choice]

    all_companies_here = _all_companies(sheets)
    companies_choice = st.multiselect(
        "Companies to compare", all_companies_here, key="ftp_companies",
        default=highlight_companies or None,
        help="Defaults to whatever you picked in 'Highlight companies' above, if applicable.",
    )
    if not companies_choice:
        st.info("Pick at least one company above to compare.")
        return

    for category_choice, cfg in PRICE_CATEGORY_CONFIG.items():
        st.subheader(category_choice.split(" ")[0])

        targets = _price_target_options(sheets, cfg, crop_id)
        if targets.empty:
            st.info(f"No {category_choice.lower()} targets found for this crop.")
            st.divider()
            continue
        target_ids = targets[cfg["target_id_col"]].tolist()
        target_names = dict(zip(targets[cfg["target_id_col"]], targets["name_en"]))

        category_companies = _companies_for_category(sheets, cfg)
        relevant_companies = [c for c in companies_choice if c in category_companies]
        if not relevant_companies:
            st.info(f"None of the selected companies have {category_choice.lower()} products for this crop.")
            st.divider()
            continue

        # Reuses the exact same Portfolio Cost machinery (all targets in
        # this category, auto-cheapest per window, correct timing-aware
        # summing for Herbicide) — this view is that same calculation
        # run once per category and laid out side by side, not a
        # separate calculation with its own rules to keep in sync.
        treatment_windows = _portfolio_treatment_windows(sheets, cfg, crop_id, target_ids, stage_filter=None)
        all_options = _portfolio_target_options(sheets, cfg, crop_id, treatment_windows)
        picks = _default_picks(all_options)
        summary, _ = _resolve_portfolio(treatment_windows, target_names, relevant_companies, picks)
        if summary.empty:
            st.info(f"No data to show for {category_choice.lower()}.")
            st.divider()
            continue

        # Two series (Fully covered / Partial coverage) rather than one,
        # so a company missing part of its coverage renders in a
        # visibly different color — a low total that's hiding a gap
        # shouldn't look as good at a glance as a genuinely complete one.
        full_vals, partial_vals = [], []
        for _, row in summary.iterrows():
            cost = row["total_cost"] if row["total_cost"] is not None and not pd.isna(row["total_cost"]) else 0
            if row["covered"] > 0 and row["covered"] == row["total_targets"]:
                full_vals.append(cost)
                partial_vals.append(0)
            else:
                full_vals.append(0)
                partial_vals.append(cost)

        fig = go.Figure()
        fig.add_trace(go.Bar(name="Fully covered", x=summary["company"], y=full_vals,
                              marker_color="#2A9D8F"))
        fig.add_trace(go.Bar(name="Partial coverage", x=summary["company"], y=partial_vals,
                              marker_color="#E76F51"))
        fig.update_layout(
            barmode="group",
            height=320,
            margin=dict(l=10, r=10, t=10, b=10),
            yaxis=dict(title=f"Total cost ({cfg['cost_unit_label']})"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )
        st.plotly_chart(fig, use_container_width=True)

        display = summary.copy()
        display["Coverage"] = display.apply(
            lambda r: f"{r['covered']}/{r['total_targets']} windows"
            + (" ⚠️" if r["covered"] < r["total_targets"] else " ✅"),
            axis=1,
        )
        display["Total Cost"] = display["total_cost"].apply(
            lambda v: (_format_price(v) + f"/{cfg['cost_unit_label']} (total)") if v is not None and not pd.isna(v) else "—"
        )
        display = display.rename(columns={
            "company": "Company", "products_used": "Products Used", "missing_targets": "Missing Targets",
        })
        display = display[["Company", "Coverage", "Total Cost", "Products Used", "Missing Targets"]]
        st.dataframe(_highlight_companies(display, "Company", highlight_companies),
                     use_container_width=True, hide_index=True)
        st.divider()


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
        "Compare mode", ["By Target", "By Chemical Name", "Portfolio Cost", "Full Treatment Program"],
        horizontal=True, key="price_mode",
        help="By Target: pick a crop + pest, compare companies. "
             "By Chemical Name: search a chemical (e.g. copper) across everything at once. "
             "Portfolio Cost: total cost across several targets at once, per company. "
             "Full Treatment Program: all three categories side by side for a whole crop.",
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
    elif mode == "Portfolio Cost":
        _render_portfolio_cost(sheets, highlight_companies)
    else:
        _render_full_treatment_program(sheets, highlight_companies)
