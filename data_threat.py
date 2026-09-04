"""
Loading/caching for crop_timeline.xlsx. Shared by threat_view.py (the
Crop Threat & Input board) and chemical_analysis_view.py (the chemical
effectiveness heatmap) since both read the exact same workbook.

Sheet list includes the 3 chemical-effectiveness-matrix sheets
(weed_matrix / insect_matrix / disease_matrix) used only by
chemical_analysis_view.py — they're loaded here regardless so both
modules share one cache, but threat_view.py never looks at them. Each
matrix sheet is long/tidy format, one row per (chemical, target) pairing:

  weed_matrix    : crop, common_name, weed_name, weed_stage, effectiveness
  insect_matrix  : crop, common_name, insect_name, effectiveness
  disease_matrix : crop, common_name, disease_name, effectiveness

"crop" here is the crop's display NAME (matching crop_stage's "crop"
column), not crop_id — kept simple for manual data entry since whoever
fills this in shouldn't need to know internal ID numbers. effectiveness
is Yes/No (see shared.py); blank shows as Unrated. weed_stage (Weed
only — e.g. "Pre-emergence"/"Early Post"/"Late Post") is optional and
powers the Spray Timing filter on the Chemical Analysis page. These
sheets are entirely optional — if they don't exist yet, they load as
empty DataFrames and chemical_analysis_view.py shows an appropriate "no
data yet" message rather than erroring.
"""

import os
import streamlit as st
import pandas as pd

DEFAULT_PATH_THREAT = "crop_timeline.xlsx"

SHEET_NAMES_THREAT = [
    "crop_stage", "crop_weeds", "weed_her",
    "crop_pest", "pest_ins",
    "crop_disease", "disease_fun",
    "fertilizer",
    "weed_matrix", "insect_matrix", "disease_matrix",
]


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
