"""
Crop Threat & Input view — reads crop_timeline.xlsx.

  crop_stage    : crop_id, crop, stage, stage_th, start_day, end_day
  crop_weeds    : crop_id, ws_id, weed_stage, weed_id, weed_name_en,
                  weed_name_th, weed_science, type, start_day, end_day
  weed_her      : crop_id, ws_id, weed_id, weed_name_en, weed_name_th,
                  common_name, hrac_code, efficiency (optional)
  crop_pest     : crop_id, pest_id, pest_name_en, pest_name_th, order,
                  rank, start_day, end_day
  pest_ins      : crop_id, pest_id, pest_name_th, common_name, irac_code,
                  efficiency (optional)
  crop_disease  : crop_id, disease_id, disease_name_en, disease_name_th,
                  disease_name_sc, type, start_day, end_day
  disease_fun   : crop_id, disease_id, disease_name_th, common_name,
                  frac_code, efficiency (optional)
  fertilizer    : crop_id, crop, formula, start_day, end_day

efficiency (Excellent/Effective/Moderate/Poor/Ineffective — same scale
as the Coverage workbook) rates the ACTIVE INGREDIENT (common_name)
itself against that specific pest/weed/disease — this sheet has no
company/brand dimension, just chemical-to-target links, unlike the
Coverage workbook's per-product junction sheets. Entirely optional:
omit the column, or leave individual cells blank, and it shows as
"Unrated" rather than breaking anything — fill it in gradually.
"""

import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go

from shared import (
    STAGE_COLORS, assign_lanes, maybe_show_rice_fertilizer_note,
    EFFICIENCY_BADGE, normalize_efficiency, EFFICIENCY_LEGEND,
)

DEFAULT_PATH_THREAT = "crop_timeline.xlsx"

SHEET_NAMES_THREAT = [
    "crop_stage", "crop_weeds", "weed_her",
    "crop_pest", "pest_ins",
    "crop_disease", "disease_fun",
    "fertilizer",
]

PALETTE = [
    "#457B9D", "#E76F51", "#2A9D8F", "#E9C46A", "#6A994E",
    "#BC4749", "#9D4EDD", "#F4A261", "#264653", "#A7C957",
]

BOARD_TITLES_THREAT = {
    "Weed": "Weed Control Windows",
    "Insect": "Insect Pressure Windows",
    "Disease": "Disease Pressure Windows",
    "Fertilizer": "Fertilizer Application Windows",
}


@st.cache_data
def load_workbook_threat(file):
    sheets = {}
    for name in SHEET_NAMES_THREAT:
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
    return sheets


def get_file_threat():
    st.sidebar.subheader("Threat & Input data source")
    uploaded = st.sidebar.file_uploader(
        "Upload workbook (.xlsx)", type=["xlsx"], key="threat_uploader"
    )
    if st.sidebar.button("🔄 Reload data", key="threat_reload"):
        st.cache_data.clear()
        st.rerun()
    if uploaded is not None:
        return uploaded
    if os.path.exists(DEFAULT_PATH_THREAT):
        return DEFAULT_PATH_THREAT
    return None


def aggregate_chemicals(merged: pd.DataFrame, group_cols: list,
                         name_col: str, code_col: str, code_label: str,
                         efficiency_col: str = None):
    """efficiency_col, if provided and present, appends an efficiency
    badge to each bullet — e.g. '• glyphosate (HRAC 9) — ✅ Excellent'.
    Here efficiency rates the ACTIVE INGREDIENT (common_name) itself
    against this specific pest/weed/disease, not a branded product —
    this sheet has no company/brand dimension, just chemical-to-target
    links. If the same (name, code) pair appears more than once within a
    window with different efficiency values (a data-entry duplicate),
    the first non-blank value wins rather than trying to merge them."""
    def _agg(g):
        raw_rows = list(zip(
            g[name_col], g[code_col],
            g[efficiency_col] if efficiency_col and efficiency_col in g.columns
            else [None] * len(g)
        ))
        triples = [
            (str(n).strip(), str(c).strip(), eff)
            for n, c, eff in raw_rows
            if pd.notna(n) or pd.notna(c)
        ]
        triples = [t for t in triples if t[0] not in ("", "nan") or t[1] not in ("", "nan")]
        count = len(triples)
        if count:
            lines = []
            for n, c, eff in triples:
                if efficiency_col:
                    badge = EFFICIENCY_BADGE[normalize_efficiency(eff) or "Unrated"]
                    lines.append(f"• {n} ({c}) — {badge}")
                else:
                    lines.append(f"• {n} ({c})")
            chem_html = "<br>".join(lines)
        else:
            chem_html = "—"
        return pd.Series({"chem_count": count, "chem_list_html": chem_html})

    agg = merged.groupby(group_cols, dropna=False).apply(_agg).reset_index()
    return agg


def build_timeline_chart_threat(df: pd.DataFrame, row_col: str, label_col: str,
                                 color_col: str, hover_fn, title: str,
                                 stage_df: pd.DataFrame = None, stage_label_col: str = "stage",
                                 show_legend: bool = True, row_label_map: dict = None,
                                 sort_col: str = None) -> go.Figure:
    if df.empty:
        fig = go.Figure()
        fig.update_layout(height=120, title=f"{title} — no data for this crop")
        return fig

    order_key = sort_col if sort_col else color_col
    order_df = (
        df.groupby(row_col)
        .agg(**{order_key: (order_key, "first"), "start_day": ("start_day", "min")})
        .reset_index()
        .sort_values([order_key, "start_day"])
    )
    row_order = order_df[row_col].tolist()
    row_to_base = {r: i for i, r in enumerate(row_order)}
    n_rows = len(row_order)

    color_values = sorted(df[color_col].dropna().astype(str).unique().tolist())
    color_map = {v: PALETTE[i % len(PALETTE)] for i, v in enumerate(color_values)}
    multi_category = len(color_values) > 1

    fig = go.Figure()
    annotations = []

    STAGE_ROW_Y = -1.3
    top_of_axis = -0.5
    if stage_df is not None and not stage_df.empty:
        sdf = stage_df.sort_values("start_day").reset_index(drop=True)
        for i, srow in sdf.iterrows():
            duration = srow["end_day"] - srow["start_day"]
            fig.add_trace(go.Bar(
                x=[duration], y=[STAGE_ROW_Y], base=[srow["start_day"]],
                orientation="h", width=0.7,
                marker=dict(color=STAGE_COLORS[i % len(STAGE_COLORS)],
                            line=dict(color="white", width=1)),
                hovertemplate=f"<b>{srow[stage_label_col]}</b><br>Day "
                               f"{srow['start_day']}–{srow['end_day']}<extra></extra>",
                showlegend=False,
            ))
            mid = (srow["start_day"] + srow["end_day"]) / 2
            annotations.append(dict(
                x=mid, y=STAGE_ROW_Y, xref="x", yref="y",
                text=str(srow[stage_label_col]), showarrow=False,
                font=dict(color="white", size=17, family="Georgia, serif"),
                xanchor="center", yanchor="middle",
            ))
        top_of_axis = STAGE_ROW_Y - 0.8

    seen_legend = set()
    row_lane_counts = {}
    for row_val, group in df.groupby(row_col):
        lane_map, n_lanes = assign_lanes(group)
        row_lane_counts[row_val] = n_lanes
        base_y = row_to_base[row_val]
        lane_height = min(0.8 / n_lanes, 0.5)

        for idx, lane in lane_map.items():
            row = df.loc[idx]
            duration = row["end_day"] - row["start_day"]
            y_center = base_y + (lane - (n_lanes - 1) / 2) * lane_height
            cat = str(row.get(color_col, ""))
            color = color_map.get(cat, PALETTE[-1])
            show_this_legend = multi_category and cat not in seen_legend
            seen_legend.add(cat)

            fig.add_trace(go.Bar(
                x=[duration],
                y=[y_center],
                base=[row["start_day"]],
                orientation="h",
                width=lane_height * 0.85,
                marker=dict(color=color, line=dict(color="white", width=1)),
                hovertemplate=hover_fn(row),
                name=cat if cat else "—",
                legendgroup=cat,
                showlegend=show_this_legend,
            ))

    total_lane_rows = sum(row_lane_counts.values())

    xaxis = dict(showgrid=True, title=dict(text="Day after planting", font=dict(size=19)),
                 tickfont=dict(size=18))
    if stage_df is not None and not stage_df.empty:
        sdf = stage_df.sort_values("start_day").reset_index(drop=True)
        stage_min = float(sdf["start_day"].min())
        stage_max = float(sdf["end_day"].max())
        span = stage_max - stage_min

        step = 20
        day_ticks = list(range(0, int(stage_max) + 1, step))
        if not day_ticks or day_ticks[-1] != int(stage_max):
            day_ticks.append(int(stage_max))

        xaxis.update(
            tickmode="array",
            tickvals=day_ticks,
            ticktext=[str(t) for t in day_ticks],
            range=[stage_min - span * 0.02, stage_max + span * 0.02],
        )

    y_ticks = [row_to_base[r] for r in row_order]
    if row_label_map:
        y_ticktext = [row_label_map.get(r, r) for r in row_order]
    else:
        y_ticktext = list(row_order)
    if stage_df is not None and not stage_df.empty:
        y_ticks = [STAGE_ROW_Y] + y_ticks
        y_ticktext = ["Crop Stage"] + y_ticktext

    fig.update_layout(
        barmode="overlay",
        height=max(240, 150 + total_lane_rows * 54),
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis=xaxis,
        yaxis=dict(
            tickmode="array",
            tickvals=y_ticks,
            ticktext=y_ticktext,
            range=[n_rows - 0.5, top_of_axis],
            title="",
            tickfont=dict(size=19),
            automargin=True,
        ),
        annotations=annotations,
        showlegend=multi_category and show_legend,
        legend_title_text=color_col,
        legend=dict(font=dict(size=17)),
        hoverlabel=dict(font=dict(size=20), align="left"),
        font=dict(size=17),
    )
    return fig


def weed_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
    weeds = sheets["crop_weeds"]
    her = sheets["weed_her"]
    raw = weeds[weeds["crop_id"] == crop_id].copy()
    her_c = her[her["crop_id"] == crop_id]
    has_efficiency = "efficiency" in her_c.columns
    her_cols = ["ws_id", "weed_id", "common_name", "hrac_code"] + (["efficiency"] if has_efficiency else [])
    merged = raw.merge(her_c[her_cols], on=["ws_id", "weed_id"], how="left")

    group_cols = ["crop_id", "ws_id", "weed_id", "weed_stage", "weed_science",
                  "weed_name_en", "weed_name_th", "type", "start_day", "end_day"]
    agg = aggregate_chemicals(merged, group_cols, "common_name", "hrac_code", "HRAC",
                               efficiency_col="efficiency" if has_efficiency else None)
    df = merged[group_cols].drop_duplicates().merge(agg, on=group_cols)

    name_col = "weed_name_th" if is_thai else "weed_science"
    row_label_map = dict(zip(df["weed_science"], df[name_col]))

    def hover(row):
        return (
            f"<b><i>{row['weed_science']}</i></b><br>"
            f"{row['weed_name_en']} / {row['weed_name_th']}<br>"
            f"Stage: {row.get('weed_stage', '')}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br>"
            f"<br><b>Products:</b><br>{row['chem_list_html']}"
            "<extra></extra>"
        )

    fig = build_timeline_chart_threat(df, row_col="weed_science", label_col="weed_stage",
                                       color_col="type", hover_fn=hover,
                                       title="Weed Control Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       row_label_map=row_label_map)
    detail_cols = ["weed_stage", "weed_science", "weed_name_en", "weed_name_th",
                   "common_name", "hrac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def insect_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
    pest = sheets["crop_pest"]
    ins = sheets["pest_ins"]
    raw = pest[pest["crop_id"] == crop_id].copy()
    has_rank = "rank" in raw.columns
    if has_rank:
        # Unranked rows sort to the bottom instead of crashing/reordering
        # unpredictably.
        raw["rank"] = pd.to_numeric(raw["rank"], errors="coerce").fillna(float("inf"))
    ins_c = ins[ins["crop_id"] == crop_id]
    has_efficiency = "efficiency" in ins_c.columns
    ins_cols = ["pest_id", "common_name", "irac_code"] + (["efficiency"] if has_efficiency else [])
    merged = raw.merge(ins_c[ins_cols], on="pest_id", how="left")

    group_cols = ["crop_id", "pest_id", "pest_name_en", "pest_name_th",
                  "order", "start_day", "end_day"]
    if has_rank:
        group_cols.append("rank")
    agg = aggregate_chemicals(merged, group_cols, "common_name", "irac_code", "IRAC",
                               efficiency_col="efficiency" if has_efficiency else None)
    df = merged[group_cols].drop_duplicates().merge(agg, on=group_cols)

    name_col = "pest_name_th" if is_thai else "pest_name_en"
    row_label_map = dict(zip(df["pest_name_en"], df[name_col]))

    def hover(row):
        rank_line = f"Rank: {int(row['rank'])}<br>" if has_rank and row['rank'] != float("inf") else ""
        return (
            f"<b>{row['pest_name_en']}</b><br>"
            f"{row['pest_name_th']}<br>"
            f"Insect order: {row.get('order', '')}<br>"
            f"{rank_line}"
            f"Day {row['start_day']}–{row['end_day']}<br>"
            f"<br><b>Products:</b><br>{row['chem_list_html']}"
            "<extra></extra>"
        )

    fig = build_timeline_chart_threat(df, row_col="pest_name_en", label_col="pest_name_en",
                                       color_col="order", hover_fn=hover,
                                       sort_col="rank" if has_rank else None,
                                       title="Insect Pressure Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       row_label_map=row_label_map)
    detail_cols = ["pest_name_en", "pest_name_th", "order", "common_name",
                   "irac_code", "start_day", "end_day"]
    if has_rank:
        detail_cols.insert(3, "rank")
    return fig, merged[detail_cols], detail_cols


def disease_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
    dis = sheets["crop_disease"]
    fun = sheets["disease_fun"]
    raw = dis[dis["crop_id"] == crop_id].copy()
    fun_c = fun[fun["crop_id"] == crop_id]
    has_efficiency = "efficiency" in fun_c.columns
    fun_cols = ["disease_id", "common_name", "frac_code"] + (["efficiency"] if has_efficiency else [])
    merged = raw.merge(fun_c[fun_cols], on="disease_id", how="left")

    group_cols = ["crop_id", "disease_id", "disease_name_en", "disease_name_th",
                  "disease_name_sc", "type", "start_day", "end_day"]
    agg = aggregate_chemicals(merged, group_cols, "common_name", "frac_code", "FRAC",
                               efficiency_col="efficiency" if has_efficiency else None)
    df = merged[group_cols].drop_duplicates().merge(agg, on=group_cols)

    name_col = "disease_name_th" if is_thai else "disease_name_en"
    row_label_map = dict(zip(df["disease_name_sc"], df[name_col]))

    def hover(row):
        return (
            f"<b><i>{row['disease_name_sc']}</i></b><br>"
            f"{row['disease_name_en']} / {row['disease_name_th']}<br>"
            f"Day {row['start_day']}–{row['end_day']}<br>"
            f"<br><b>Products:</b><br>{row['chem_list_html']}"
            "<extra></extra>"
        )

    fig = build_timeline_chart_threat(df, row_col="disease_name_sc", label_col="disease_name_sc",
                                       color_col="type", hover_fn=hover,
                                       title="Disease Pressure Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       row_label_map=row_label_map)
    detail_cols = ["disease_name_sc", "disease_name_en", "disease_name_th",
                   "common_name", "frac_code", "type", "start_day", "end_day"]
    return fig, merged[detail_cols], detail_cols


def fertilizer_board_threat(crop_id, sheets, crop_stage_df, stage_label_col):
    is_thai = stage_label_col.endswith("_th")
    fert = sheets["fertilizer"]
    df = fert[fert["crop_id"] == crop_id].copy()

    has_stage = "stage" in df.columns
    has_type = "type" in df.columns and df["type"].notna().any()

    if has_type:
        type_options = sorted(df["type"].dropna().astype(str).unique().tolist())
        selected_types = st.multiselect(
            "Fertilizer type", type_options, default=type_options,
            help="Choose one type, or keep several selected to see them combined "
                 "on the same timeline (e.g. foliar + granular).",
            key="threat_fert_type",
        )
        df = df[df["type"].astype(str).isin(selected_types)] if selected_types else df.iloc[0:0]

    detail_cols = [c for c in ["stage", "type", "formula", "start_day", "end_day"]
                   if c in df.columns]
    detail_df = df[detail_cols].copy()

    if df.empty:
        fig = build_timeline_chart_threat(df, row_col="formula", label_col="formula",
                                           color_col="_none", hover_fn=lambda r: "",
                                           title="Fertilizer Application Windows",
                                           stage_df=crop_stage_df, stage_label_col=stage_label_col)
        return fig, detail_df, detail_cols

    row_col = "stage" if has_stage else "formula"
    if has_stage:
        stage_name_col = "stage_th" if (is_thai and "stage_th" in df.columns) else "stage"
        row_label_map = dict(zip(df[row_col], df[stage_name_col]))
    else:
        row_label_map = None

    group_cols = [c for c in ["crop_id", row_col, "start_day", "end_day"] if c in df.columns]

    def _agg(g):
        if has_type:
            items = [
                f"• {f} ({t})" for f, t in zip(g["formula"], g["type"])
                if pd.notna(f) or pd.notna(t)
            ]
            types_present = sorted({str(t) for t in g["type"].dropna()})
        else:
            items = [f"• {f}" for f in g["formula"] if pd.notna(f)]
            types_present = []
        return pd.Series({
            "formula_list_html": "<br>".join(items) if items else "—",
            "type_combo": " + ".join(types_present) if types_present else "Fertilizer",
        })

    agg = df.groupby(group_cols, dropna=False).apply(_agg).reset_index()
    df_agg = df[group_cols].drop_duplicates().merge(agg, on=group_cols)

    color_col = "type_combo"

    def hover(row):
        parts = []
        if has_stage:
            parts.append(f"<b>{row['stage']}</b>")
        parts.append(f"Day {row['start_day']}–{row['end_day']}")
        parts.append(f"<br><b>Formula:</b><br>{row['formula_list_html']}")
        return "<br>".join(parts) + "<extra></extra>"

    fig = build_timeline_chart_threat(df_agg, row_col=row_col, label_col=row_col,
                                       color_col=color_col, hover_fn=hover,
                                       title="Fertilizer Application Windows",
                                       stage_df=crop_stage_df, stage_label_col=stage_label_col,
                                       show_legend=has_type, row_label_map=row_label_map)
    return fig, detail_df, detail_cols


BOARDS_THREAT = {
    "Weed": weed_board_threat,
    "Insect": insect_board_threat,
    "Disease": disease_board_threat,
    "Fertilizer": fertilizer_board_threat,
}


def render_threat_view():
    st.title("🌾 Crop Threat & Input Dashboard")

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

    crop_lookup = stage_df_all[["crop_id", "crop"]].drop_duplicates()
    crop_name_to_id = dict(zip(crop_lookup["crop"], crop_lookup["crop_id"]))

    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        crop_choice = st.selectbox("Crop", list(crop_name_to_id.keys()), key="threat_crop")
    with col2:
        board_choice = st.selectbox("Board", list(BOARDS_THREAT.keys()), index=0, key="threat_board")
    with col3:
        stage_label_choice = st.radio("Label language", ["English", "Thai"],
                                       horizontal=True, key="threat_lang")
    label_col = "stage" if stage_label_choice == "English" else "stage_th"

    crop_id = crop_name_to_id[crop_choice]

    crop_stage_df = stage_df_all[stage_df_all["crop_id"] == crop_id]
    if crop_stage_df.empty:
        st.warning("No stage data for this crop.")
        st.stop()

    st.subheader(BOARD_TITLES_THREAT[board_choice])
    maybe_show_rice_fertilizer_note(crop_choice, board_choice)
    fig, board_df, detail_cols = BOARDS_THREAT[board_choice](crop_id, sheets, crop_stage_df, label_col)
    st.plotly_chart(fig, use_container_width=True)
    if board_choice != "Fertilizer":
        st.caption(EFFICIENCY_LEGEND)

    if board_df.empty:
        st.info(f"No {board_choice.lower()} data for this crop.")
    else:
        with st.expander(f"{board_choice} detail table (one row per product)"):
            st.dataframe(board_df, use_container_width=True, hide_index=True)
