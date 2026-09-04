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
