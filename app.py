"""
Crop Dashboard — entry point
---------------------------------
Sidebar switch between three independent views, each in its own module:

  threat_view.py   -> "Crop Threat & Input"  (reads crop_timeline.xlsx)
  coverage_view.py -> "Product Coverage"     (reads crop_timeline_coverage.xlsx;
                       also contains the AI analysis feature)
  price_view.py    -> "Price Comparison"     (reads crop_timeline_coverage.xlsx)

Shared constants/helpers (tier, efficiency, price formatting, chart
helpers) live in shared.py. The Coverage-workbook loader (used by both
coverage_view.py and price_view.py, since they read the same file) lives
in data_cov.py. This file only wires the three views together — it holds
no view logic itself.
"""

import streamlit as st

from threat_view import render_threat_view
from coverage_view import render_coverage_view
from price_view import render_price_comparison_view

st.set_page_config(page_title="Crop Dashboard", layout="wide")

st.sidebar.subheader("View")
view = st.sidebar.radio(
    "Choose a dashboard",
    ["Crop Threat & Input", "Product Coverage", "Price Comparison"],
    label_visibility="collapsed",
    key="view_switch",
)

if view == "Crop Threat & Input":
    render_threat_view()
elif view == "Product Coverage":
    render_coverage_view()
else:
    render_price_comparison_view()
