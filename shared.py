"""
Shared constants and helpers used across all three views (Threat,
Coverage, Price Comparison). Nothing in here is view-specific — anything
that only one view needs lives in that view's own module instead.
"""

import streamlit as st
import pandas as pd

# =====================================================================
# Timeline chart helpers (used by both build_timeline_chart_threat in
# threat_view.py and build_timeline_chart_cov in coverage_view.py)
# =====================================================================

STAGE_COLORS = [
    "#8ECAE6", "#219EBC", "#023047", "#FFB703", "#FB8500",
    "#A7C957", "#6A994E", "#BC4749", "#9D4EDD", "#264653",
]


def assign_lanes(group: pd.DataFrame):
    """For genuinely different, overlapping time windows within the same
    row (e.g. two distinct spray dates for the same weed)."""
    lanes_end = []
    assignment = {}
    for idx, row in group.sort_values("start_day").iterrows():
        placed = False
        for lane_idx in range(len(lanes_end)):
            if row["start_day"] >= lanes_end[lane_idx]:
                lanes_end[lane_idx] = row["end_day"]
                assignment[idx] = lane_idx
                placed = True
                break
        if not placed:
            lanes_end.append(row["end_day"])
            assignment[idx] = len(lanes_end) - 1
    return assignment, max(len(lanes_end), 1)


RICE_FERTILIZER_NOTE = (
    "🌾 **Rice fertilizer guideline:** total recommended use is approximately "
    "**50–75 kg per rai** across all applications combined."
)


def maybe_show_rice_fertilizer_note(crop_choice: str, board_choice: str):
    if board_choice == "Fertilizer" and "rice" in str(crop_choice).lower():
        st.info(RICE_FERTILIZER_NOTE)


# =====================================================================
# Tier — used by Coverage boards, Price Comparison, and the AI analysis
# prompt. Accepts a range of real-world spellings (e.g. "premium",
# "Mid-Tier", "generic") via .title()-casing; blank/unrecognized values
# fall back to Generic.
# =====================================================================

TIER_ORDER = ["Premium", "Medium", "Generic"]
TIER_BADGE = {"Premium": "🟣 Premium", "Medium": "🟡 Medium", "Generic": "⚪ Generic"}


def normalize_tier(val) -> str:
    s = str(val).strip().title() if pd.notna(val) else ""
    return s if s in TIER_ORDER else "Generic"


# =====================================================================
# Effectiveness (this section) is used ONLY by chemical_analysis_view.py
# (the weed_matrix/insect_matrix/disease_matrix sheets) — deliberately
# binary Yes/No, since the team found a finer scale too hard to assess
# consistently for that quick-comparison heatmap. Everywhere else in the
# app — Coverage boards, Threat & Input hover, Price Comparison, AI
# analysis — still uses the 5-point EFFICIENCY_* scale further below.
# =====================================================================

EFFECTIVENESS_ORDER = ["Yes", "No"]
EFFECTIVENESS_BADGE = {
    "Yes": "✅ Yes",
    "No": "❌ No",
    "Unrated": "❔ Unrated",
}
EFFECTIVENESS_ALIASES = {
    "yes": "Yes", "y": "Yes", "true": "Yes", "1": "Yes",
    "effective": "Yes", "good": "Yes", "excellent": "Yes", "works": "Yes",
    "no": "No", "n": "No", "false": "No", "0": "No",
    "ineffective": "No", "poor": "No", "not effective": "No", "doesn't work": "No",
}


def normalize_effectiveness(val):
    """Returns 'Yes'/'No', or None if blank/unrecognized (caller decides
    how to label that — see 'Unrated' usage below). Used only by
    chemical_analysis_view.py."""
    if pd.isna(val):
        return None
    raw = str(val).strip()
    if not raw:
        return None
    key = raw.lower().replace("_", "-").replace("  ", " ")
    if key in EFFECTIVENESS_ALIASES:
        return EFFECTIVENESS_ALIASES[key]
    title = raw.title()
    return title if title in EFFECTIVENESS_ORDER else None


# Shown as a caption under the Chemical Analysis heatmap so the badges
# have a clear, consistent meaning rather than being left to guesswork.
EFFECTIVENESS_LEGEND = (
    "✅ **Yes** — effective against this target | "
    "❌ **No** — not effective | "
    "❔ **Unrated** — not yet assessed"
)

# Numeric score for heatmap coloring (chemical_analysis_view.py) — 3
# distinct levels so Unrated renders visibly different (gray) from a
# confirmed No (red), rather than the two being visually confused.
EFFECTIVENESS_SCORE = {
    "Yes": 2,
    "Unrated": 1,
    "No": 0,
}


# =====================================================================
# Efficiency (5-point scale) lives on the JUNCTION sheets (weed_her /
# pest_ins / disease_fun in crop_timeline_coverage.xlsx, and the
# equivalent sheets in crop_timeline.xlsx) — not the product master
# sheets. A product's real-world efficiency varies by which pest/weed/
# disease it's up against (e.g. Chemical A might be Excellent against
# Pest A but only Moderate against Pest B), so it belongs at the
# product-x-target pairing, the same place trade_name/common_name
# already sits. Used by Coverage boards, Threat & Input hover, Price
# Comparison, and the AI analysis — NOT by chemical_analysis_view.py,
# which uses the separate binary EFFECTIVENESS_* system above instead.
#
# Expected values: Excellent / Effective / Moderate / Poor / Ineffective
# (5-point scale). Unlike tier, a blank/unrecognized value normalizes to
# "Unrated" (not "Poor" or "Ineffective") — this is new data you're
# likely filling in gradually, and an unrated product shouldn't be
# treated as if it's a known poor performer. Fertilizer has no
# efficiency concept (it's a nutrition schedule, not a pest/weed/disease
# control decision), so it's intentionally not part of this anywhere.
# =====================================================================

EFFICIENCY_ORDER = ["Excellent", "Effective", "Moderate", "Poor", "Ineffective"]
EFFICIENCY_BADGE = {
    "Excellent": "✅ Excellent",
    "Effective": "🟢 Effective",
    "Moderate": "🟨 Moderate",
    "Poor": "🟠 Poor",
    "Ineffective": "🔴 Ineffective",
    "Unrated": "❔ Unrated",
}
EFFICIENCY_ALIASES = {
    "excellent": "Excellent", "strong": "Excellent", "high": "Excellent", "very good": "Excellent",
    "effective": "Effective", "good": "Effective",
    "moderate": "Moderate", "average": "Moderate", "medium": "Moderate", "fair": "Moderate",
    "poor": "Poor", "weak": "Poor", "low": "Poor",
    "ineffective": "Ineffective", "none": "Ineffective", "no effect": "Ineffective", "very poor": "Ineffective",
}


def normalize_efficiency(val):
    """Returns 'Excellent'/'Effective'/'Moderate'/'Poor'/'Ineffective', or
    None if blank/unrecognized (caller decides how to label that — see
    'Unrated' usage below)."""
    if pd.isna(val):
        return None
    raw = str(val).strip()
    if not raw:
        return None
    key = raw.lower().replace("_", "-").replace("  ", " ")
    if key in EFFICIENCY_ALIASES:
        return EFFICIENCY_ALIASES[key]
    title = raw.title()
    return title if title in EFFICIENCY_ORDER else None


# Shown as a caption under the Weed/Insect/Disease charts so the rating
# words have a concrete, consistent meaning rather than being left to
# guesswork. Adjust the percentages here if your team's definition of
# "control" changes — this is the only place they're defined.
EFFICIENCY_LEGEND = (
    "✅ **Excellent** >90% control | "
    "🟢 **Effective** 80–89% control | "
    "🟨 **Moderate** 60–79% control | "
    "🟠 **Poor** 40–59% control | "
    "🔴 **Ineffective** <40% control"
)


# =====================================================================
# Price formatting — used by Coverage boards and Price Comparison.
# =====================================================================

def _format_price(val) -> str:
    """฿1,234 style formatting for a single price cell. Returns '' for
    blank/non-numeric so callers can cleanly skip it."""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    try:
        num = float(val)
    except (TypeError, ValueError):
        return ""
    return f"฿{num:,.0f}"


def _format_cost_range(values, unit_label: str) -> str:
    """Collapses one or more numeric cost values (e.g. every company
    product covering a window) into a single '฿95/rai' or, when the
    covering products disagree on cost, a '฿95–120/rai' range. Returns
    '—' if nothing numeric was found."""
    nums = []
    for v in values:
        try:
            if pd.isna(v):
                continue
        except (TypeError, ValueError):
            pass
        try:
            nums.append(float(v))
        except (TypeError, ValueError):
            continue
    if not nums:
        return "—"
    lo, hi = min(nums), max(nums)
    if lo == hi:
        return f"฿{lo:,.0f}/{unit_label}"
    return f"฿{lo:,.0f}–{hi:,.0f}/{unit_label}"
