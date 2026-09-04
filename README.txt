CROP DASHBOARD — FILE GUIDE
=================================
Last major change: split one big app.py (~2,000 lines) into 6 smaller
files so it's easier to navigate and edit. Behavior is identical to
before the split — this was a reorganization, not a rewrite.


app.py  (39 lines — the entry point)
-------------------------------------
This is the file Streamlit actually runs. It does almost nothing itself:
shows the sidebar menu ("Crop Threat & Input" / "Product Coverage" /
"Price Comparison") and calls the matching view function from the other
files. If the app won't start, this is the first place to check — but
the actual bug is almost always in one of the other files.


shared.py  (constants + helpers used by ALL THREE views)
-----------------------------------------------------------
No Streamlit pages of its own — just the building blocks every view
needs:
  - Tier system (Premium/Medium/Generic) + normalize_tier()
  - Efficiency system (Excellent/Effective/Moderate/Poor/Ineffective)
    + normalize_efficiency() + the badge emojis + the legend caption
    text shown under charts
  - Price formatting helpers (turns a raw number into "฿1,234")
  - Small chart-drawing helpers (stage colors, lane assignment for
    overlapping timeline bars)
If you ever want to change the 5-point efficiency scale, the tier
rules, or how prices are displayed — this is the one file to edit,
and the change applies everywhere automatically.


data_cov.py  (loads crop_timeline_coverage.xlsx)
----------------------------------------------------
Just the loading/caching logic for the Coverage workbook — reads the
Excel file, cleans up column names, caches it so it doesn't re-read on
every click. Shared by coverage_view.py AND price_view.py since they
both read the exact same file, just use it differently.


threat_view.py  (the "Crop Threat & Input" page)
----------------------------------------------------
Reads crop_timeline.xlsx. Shows the Weed / Insect / Disease / Fertilizer
timeline charts with which chemicals (by common_name/active ingredient)
are options for each pest, colored by threat type. This is the
"reference" view — no company/coverage comparison, just what's
biologically relevant and when.


coverage_view.py  (the "Product Coverage" page — the BIG file)
--------------------------------------------------------------------
Reads crop_timeline_coverage.xlsx. This is the main working dashboard:
  - Pick a crop + your company -> see green/red coverage per pest
  - Hover shows which of YOUR products cover that pest, their tier,
    efficiency, and price
  - The "Analyze full coverage" AI button also lives here (calls
    Claude to summarize strong/weak spots across all 4 boards)
This file is intentionally the biggest one — the AI analysis logic
stayed bundled in here instead of its own file because it directly
calls this file's own board functions to build its summary. Splitting
it out would have created an import loop.


price_view.py  (the "Price Comparison" page)
--------------------------------------------------
Also reads crop_timeline_coverage.xlsx (via data_cov.py). Lets you
pick a crop + category (Herbicide/Insecticide/Fungicide) + a specific
pest, and see EVERY company's product for that pest side by side,
sorted cheapest or priciest first. For Herbicide only, there's also a
"Spray Timing" filter (Early Post / Late Post, from crop_weeds'
weed_stage column).


QUICK RULE OF THUMB — "where do I make my change?"
------------------------------------------------------
  Change a badge color, the tier/efficiency scale, price formatting
    -> shared.py

  Change how the Coverage workbook is read/cached
    -> data_cov.py

  Change something on the Threat & Input page
    -> threat_view.py

  Change something on the Product Coverage page OR the AI analysis
    -> coverage_view.py

  Change something on the Price Comparison page
    -> price_view.py

  Change the sidebar menu itself, or which view loads first
    -> app.py
