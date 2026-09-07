"""
SAC Crop Dashboard — entry point
---------------------------------
Sidebar switch between four independent views, each in its own module:

  threat_view.py            -> "Crop Threat & Input"  (reads crop_timeline.xlsx)
  chemical_analysis_view.py -> "Chemical Analysis"     (reads crop_timeline.xlsx;
                                chemical x weed/pest/disease effectiveness heatmap)
  coverage_view.py          -> "Product Coverage"      (reads crop_timeline_coverage.xlsx;
                                also contains the AI analysis feature)
  price_view.py             -> "Price Comparison"      (reads crop_timeline_coverage.xlsx)

Shared constants/helpers (tier, efficiency/effectiveness, price
formatting, chart helpers) live in shared.py — note shared.py has TWO
separate rating systems: the 5-point EFFICIENCY_* scale (Coverage,
Threat & Input, Price Comparison, AI analysis) and the binary
EFFECTIVENESS_* Yes/No scale (Chemical Analysis only). Workbook
loaders shared by more than one view live in their own data_*.py
module: data_threat.py (crop_timeline.xlsx,
used by threat_view.py + chemical_analysis_view.py) and data_cov.py
(crop_timeline_coverage.xlsx, used by coverage_view.py + price_view.py).
This file only wires the four views together — it holds no view logic
itself.
"""

import streamlit as st

from threat_view import render_threat_view
from chemical_analysis_view import render_chemical_analysis_view
from coverage_view import render_coverage_view
from price_view import render_price_comparison_view

st.set_page_config(page_title="Crop Dashboard", layout="wide")

st.sidebar.subheader("View")
view = st.sidebar.radio(
    "Choose a dashboard",
    ["Crop Threat & Input", "Chemical Analysis", "Product Coverage", "Price Comparison"],
    label_visibility="collapsed",
    key="view_switch",
)

if view == "Crop Threat & Input":
    render_threat_view()
elif view == "Chemical Analysis":
    render_chemical_analysis_view()
elif view == "Product Coverage":
    render_coverage_view()
else:
    render_price_comparison_view()
