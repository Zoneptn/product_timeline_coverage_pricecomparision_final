"""
Chemical Analysis view — a heatmap answering "does chemical X work
against every weed/pest/disease on this crop?", built by picking
chemicals one at a time so you can compare a handful side by side
rather than seeing everything at once.

Reads crop_timeline.xlsx (via data_threat.py), 3 dedicated long/tidy
sheets — one row per (chemical, target) pairing, kept deliberately
separate from weed_her/pest_ins/disease_fun so this feature can never
affect the Crop Threat & Input chart's hover text, no matter how much
data gets added here:

  weed_matrix    : crop, common_name, weed_name, weed_stage, effectiveness
  insect_matrix  : crop, common_name, insect_name, effectiveness
  disease_matrix : crop, common_name, disease_name, effectiveness

"crop" is the crop's display NAME (matching crop_stage's "crop"
column), not crop_id — kept simple for manual data entry. effectiveness
is Effective / Moderate / Ineffective — a 3-point scale, after the team
first tried binary Yes/No and found it too coarse for this heatmap;
blank/missing shows as Unrated (gray), not assumed Ineffective.

weed_stage (Weed only — e.g. "Pre-emergence", "Early Post", "Late
Post") is the spray timing, same concept as crop_weeds' weed_stage
column used in the Price Comparison view's "Spray Timing" filter.
Optional column: if it's missing from the sheet, the Weed board simply
skips offering the timing filter rather than erroring.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go

from shared import EFFECTIVENESS_ORDER, EFFECTIVENESS_SCORE, normalize_effectiveness, EFFECTIVENESS_LEGEND
from data_threat import DEFAULT_PATH_THREAT, load_workbook_threat, get_file_threat

CHEMICAL_MATRIX_CONFIG = {
    # stage_col is Weed-only, matching Price Comparison's pattern —
    # intentionally absent from Insect/Disease below.
    "Weed": {"sheet": "weed_matrix", "target_col": "weed_name", "stage_col": "weed_stage"},
    "Insect": {"sheet": "insect_matrix", "target_col": "insect_name"},
    "Disease": {"sheet": "disease_matrix", "target_col": "disease_name"},
}

# 4 discrete levels (Unrated / Ineffective / Moderate / Effective),
# matching EFFECTIVENESS_SCORE (0/1/2/3) — a flat-then-jump colorscale
# rather than a smooth gradient, so each cell reads as one of exactly 4
# clear colors, not a blend.
_HEATMAP_COLORSCALE = [
    [0.00, "#BDBDBD"],  # 0 - Unrated
    [0.25, "#BDBDBD"],
    [0.25, "#E63946"],  # 1 - Ineffective
    [0.50, "#E63946"],
    [0.50, "#F6D55C"],  # 2 - Moderate
    [0.75, "#F6D55C"],
    [0.75, "#2A9D8F"],  # 3 - Effective
    [1.00, "#2A9D8F"],
]


def _matrix_stage_options(matrix_df: pd.DataFrame, cfg: dict, crop_choice: str) -> list:
    """Distinct spray-timing values (e.g. Pre-emergence/Early Post/Late
    Post) for this crop, only meaningful when cfg has a 'stage_col'
    (currently just Weed). Returns [] for categories without the
    concept, or if the column isn't present in the sheet yet — the
    caller treats an empty list as 'no timing filter to offer'."""
    stage_col = cfg.get("stage_col")
    if not stage_col or matrix_df.empty or stage_col not in matrix_df.columns or "crop" not in matrix_df.columns:
        return []
    df = matrix_df[matrix_df["crop"].astype(str).str.strip() == str(crop_choice).strip()]
    return sorted({str(v).strip() for v in df[stage_col].dropna() if str(v).strip()})


def _matrix_for_crop(matrix_df: pd.DataFrame, cfg: dict, crop_choice: str,
                      stage_filter: str = None) -> pd.DataFrame:
    """Rows for this crop only, with effectiveness normalized. Returns
    empty if the sheet doesn't exist yet or has no rows for this crop.
    stage_filter, when given and cfg has a 'stage_col', narrows this to
    only rows at that spray timing (e.g. only 'Early Post' rows).

    Always guarantees 'common_name' and 'effectiveness' columns exist on
    the result, even if the underlying sheet is missing one of them
    (e.g. effectiveness hasn't been added to that category's sheet yet)
    — downstream code (the heatmap, the raw-data table) can then safely
    assume both columns are present rather than checking every time. A
    genuinely missing effectiveness column just means every row shows
    as Unrated, same as a blank cell would."""
    target_col = cfg["target_col"]
    if matrix_df.empty or target_col not in matrix_df.columns or "crop" not in matrix_df.columns:
        return pd.DataFrame(columns=["crop", "common_name", target_col, "effectiveness"])
    df = matrix_df[matrix_df["crop"].astype(str).str.strip() == str(crop_choice).strip()].copy()
    stage_col = cfg.get("stage_col")
    if stage_filter and stage_col and stage_col in df.columns:
        df = df[df[stage_col].astype(str).str.strip() == stage_filter]
    if "effectiveness" not in df.columns:
        df["effectiveness"] = pd.NA
    if "common_name" not in df.columns:
        df["common_name"] = pd.NA
    return df


def _build_heatmap(df: pd.DataFrame, target_col: str, chemicals: list) -> go.Figure:
    """chemicals is the ordered list of chemicals to show as rows (the
    order the user picked them in, top to bottom on screen — reversed
    for Plotly since heatmap y-axis renders bottom-to-top by default).
    Columns are every distinct target in df for this crop."""
    targets = sorted(df[target_col].dropna().astype(str).str.strip().unique().tolist())
    if not targets or not chemicals:
        fig = go.Figure()
        fig.update_layout(height=120, title="Pick at least one chemical to see the heatmap")
        return fig

    # Look up rating per (chemical, target); a chemical/target pair with
    # no matching row is Unrated, not an error.
    lookup = {}
    for _, r in df.iterrows():
        chem = str(r.get("common_name", "")).strip()
        tgt = str(r.get(target_col, "")).strip()
        eff = normalize_effectiveness(r.get("effectiveness")) or "Unrated"
        lookup[(chem, tgt)] = eff

    z = []       # numeric score per cell (0/1/2), for coloring
    text = []    # rating word per cell, for the label/hover
    for chem in reversed(chemicals):  # reversed so first-picked ends up on top
        row_z, row_text = [], []
        for tgt in targets:
            eff = lookup.get((chem.strip(), tgt), "Unrated")
            row_z.append(EFFECTIVENESS_SCORE[eff])
            row_text.append(eff)
        z.append(row_z)
        text.append(row_text)

    fig = go.Figure(data=go.Heatmap(
        z=z,
        x=targets,
        y=list(reversed(chemicals)),
        text=text,
        texttemplate="%{text}",
        textfont=dict(size=17),
        colorscale=_HEATMAP_COLORSCALE,
        zmin=0, zmax=3,
        showscale=False,
        hovertemplate="<b>%{y}</b> vs <b>%{x}</b><br>%{text}<extra></extra>",
        xgap=3, ygap=3,
    ))
    fig.update_layout(
        # Extra top margin/height headroom for the rotated, larger x-axis
        # labels (weed/pest/disease names are often long) — automargin
        # lets Plotly grow the margin further still if a name is
        # especially long, rather than clipping it.
        height=max(260, 160 + 60 * len(chemicals)),
        margin=dict(l=10, r=10, t=140, b=10),
        xaxis=dict(
            tickfont=dict(size=17), side="top",
            tickangle=-45, automargin=True,
        ),
        yaxis=dict(tickfont=dict(size=18), automargin=True),
        font=dict(size=17),
    )
    return fig


def render_chemical_analysis_view():
    st.title("🧪 Chemical Analysis")
    st.caption("Compare how a few chemicals perform against every weed, pest, or disease on a crop.")

    data_file = get_file_threat()
    if data_file is None:
        st.warning(
            f"No workbook found. Upload one from the sidebar, or place a file "
            f"named `{DEFAULT_PATH_THREAT}` next to `app.py`."
        )
        st.stop()

    try:
        sheets = load_workbook_threat(data_file)
    except Exception as e:
        st.error(f"Couldn't read the workbook: {e}")
        st.stop()

    stage_df_all = sheets["crop_stage"]
    if stage_df_all.empty:
        st.error("`crop_stage` sheet is missing or empty.")
        st.stop()

    crop_choices = sorted(stage_df_all["crop"].dropna().astype(str).unique().tolist())

    col1, col2 = st.columns([2, 1])
    with col1:
        crop_choice = st.selectbox("Crop", crop_choices, key="chem_crop")
    with col2:
        category_choice = st.selectbox("Category", list(CHEMICAL_MATRIX_CONFIG.keys()), key="chem_category")
    cfg = CHEMICAL_MATRIX_CONFIG[category_choice]
    target_col = cfg["target_col"]

    matrix_df = sheets.get(cfg["sheet"], pd.DataFrame())
    if matrix_df.empty or cfg["sheet"] not in sheets:
        st.info(
            f"No `{cfg['sheet']}` sheet found yet in the workbook. Add it with columns "
            f"`crop, common_name, {target_col}, effectiveness` to use this page for {category_choice.lower()}s."
        )
        st.stop()

    # Spray Timing (e.g. Pre-emergence / Early Post / Late Post) only
    # exists for Weed, via weed_matrix's weed_stage column — other
    # categories simply don't get this filter offered at all.
    stage_options = _matrix_stage_options(matrix_df, cfg, crop_choice)
    stage_filter = None
    if stage_options:
        stage_choice = st.selectbox(
            "Spray Timing", ["All"] + stage_options, key="chem_stage"
        )
        stage_filter = None if stage_choice == "All" else stage_choice

    crop_df = _matrix_for_crop(matrix_df, cfg, crop_choice, stage_filter=stage_filter)
    if crop_df.empty:
        if stage_filter:
            st.info(f"No {category_choice.lower()} chemical data found for {crop_choice} at '{stage_filter}' timing yet.")
        else:
            st.info(f"No {category_choice.lower()} chemical data found for {crop_choice} yet.")
        st.stop()

    chemical_options = sorted(crop_df["common_name"].dropna().astype(str).str.strip().unique().tolist())

    # Selection persists per (crop, category, stage) combo via the
    # widget key itself, so switching any of them naturally resets which
    # chemicals are shown rather than carrying over an unrelated list.
    widget_key = f"chem_pick_{crop_choice}_{category_choice}_{stage_filter or 'All'}"
    chosen = st.multiselect(
        "Chemicals to compare (pick one to start, add more to compare side by side)",
        chemical_options, key=widget_key,
    )

    if not chosen:
        st.info("Pick at least one chemical above to see the heatmap.")
        st.stop()

    fig = _build_heatmap(crop_df, target_col, chosen)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(EFFECTIVENESS_LEGEND)

    with st.expander("Raw data for this crop/category"):
        # Belt-and-suspenders: only select columns that actually exist,
        # even though _matrix_for_crop already guarantees common_name/
        # effectiveness are present — protects against any future column
        # this expander might reference that isn't guaranteed yet.
        display_cols = [c for c in ["common_name", target_col, "effectiveness"] if c in crop_df.columns]
        sort_cols = [c for c in ["common_name", target_col] if c in crop_df.columns]
        st.dataframe(
            crop_df[display_cols].sort_values(sort_cols) if sort_cols else crop_df[display_cols],
            use_container_width=True, hide_index=True,
        )
