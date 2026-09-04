SAC CROP DASHBOARD — FILE GUIDE
=================================
Current structure: 8 files, 4 pages (sidebar views). Last major
addition: Chemical Analysis page + its supporting data_threat.py loader.


app.py  (~46 lines — the entry point)
-------------------------------------
This is the file Streamlit actually runs. It does almost nothing itself:
shows the sidebar menu and calls the matching view function from the
other files. Current sidebar order:
  1. Crop Threat & Input
  2. Chemical Analysis
  3. Product Coverage
  4. Price Comparison
If the app won't start, this is the first place to check — but the
actual bug is almost always in one of the other files (see
TROUBLESHOOTING below for how to tell).


shared.py  (constants + helpers used by ALL FOUR views)
-----------------------------------------------------------
No Streamlit pages of its own — just the building blocks every view
needs:
  - Tier system (Premium/Medium/Generic) + normalize_tier()
  - Efficiency system (Excellent/Effective/Moderate/Poor/Ineffective)
    + normalize_efficiency() + badge emojis + EFFICIENCY_LEGEND caption
    + EFFICIENCY_SCORE (numeric version, used by Chemical Analysis'
    heatmap coloring)
  - Price formatting helpers (turns a raw number into "฿1,234")
  - Small chart-drawing helpers (stage colors, lane assignment for
    overlapping timeline bars)
If you ever want to change the 5-point efficiency scale, the tier
rules, or how prices are displayed — this is the one file to edit,
and the change applies everywhere automatically.


data_cov.py  (loads crop_timeline_coverage.xlsx)
----------------------------------------------------
Loading/caching logic for the Coverage workbook. Shared by
coverage_view.py AND price_view.py since they both read the exact same
file, just use it differently.


data_threat.py  (loads crop_timeline.xlsx)
----------------------------------------------------
Loading/caching logic for the Threat workbook. Shared by
threat_view.py AND chemical_analysis_view.py since they both read the
exact same file. Also lists the 3 chemical-matrix sheets (see below) so
they load automatically even though only Chemical Analysis uses them.


threat_view.py  (the "Crop Threat & Input" page)
----------------------------------------------------
Reads crop_timeline.xlsx. Shows the Weed / Insect / Disease / Fertilizer
timeline charts with which chemicals (by common_name/active ingredient)
are options for each pest, colored by threat type. This is the
"reference" view — no company/coverage comparison, just what's
biologically relevant and when.


chemical_analysis_view.py  (the "Chemical Analysis" page)
----------------------------------------------------------------
Also reads crop_timeline.xlsx (via data_threat.py), but from 3
DIFFERENT dedicated sheets — not weed_her/pest_ins/disease_fun:

  weed_matrix    : crop, common_name, weed_name, weed_stage, efficiency
  insect_matrix  : crop, common_name, insect_name, efficiency
  disease_matrix : crop, common_name, disease_name, efficiency

Pick a crop + category (Weed/Insect/Disease), then pick one or more
chemicals — each pick adds a row to a heatmap, columns are every
weed/pest/disease that crop has data for, cells are colored green
(Excellent) to red (Ineffective), gray for Unrated. Weed only, also
has a "Spray Timing" filter (Pre-emergence/Early Post/Late Post, from
weed_matrix's weed_stage column).

These 3 sheets are completely separate from weed_her/pest_ins/
disease_fun used by threat_view.py — adding data here can NEVER affect
that page's hover text, no matter how much you add.

Leave a cell blank if you haven't rated it yet — shows as "Unrated"
(gray), not treated as a bad rating. Only type one of: Excellent,
Effective, Moderate, Poor, Ineffective.


coverage_view.py  (the "Product Coverage" page — the BIG file)
--------------------------------------------------------------------
Reads crop_timeline_coverage.xlsx. This is the main working dashboard:
  - Pick a crop + your company -> see green/red coverage per pest
  - Hover shows which of YOUR products cover that pest, their tier,
    efficiency, and price
  - The "Analyze full coverage" AI button also lives here (calls
    Claude to summarize strong/weak spots across all 4 boards; price
    is deliberately excluded from what the AI sees/discusses)
This file is intentionally the biggest one — the AI analysis logic
stayed bundled in here instead of its own file because it directly
calls this file's own board functions to build its summary. Splitting
it out would have created an import loop.


price_view.py  (the "Price Comparison" page)
--------------------------------------------------
Also reads crop_timeline_coverage.xlsx (via data_cov.py). Lets you
pick a crop + category (Herbicide/Insecticide/Fungicide) + a specific
pest, and see EVERY company's product for that pest side by side,
sorted cheapest or priciest first. Herbicide only, also has a "Spray
Timing" filter (from crop_weeds' weed_stage column — a different sheet
than Chemical Analysis' weed_matrix, but the same concept).


QUICK RULE OF THUMB — "where do I make my change?"
------------------------------------------------------
  Change a badge color, the tier/efficiency scale, price formatting
    -> shared.py

  Change how the Coverage workbook is read/cached
    -> data_cov.py

  Change how the Threat workbook is read/cached, or which sheets load
    -> data_threat.py

  Change something on the Threat & Input page
    -> threat_view.py

  Change something on the Chemical Analysis heatmap
    -> chemical_analysis_view.py

  Change something on the Product Coverage page OR the AI analysis
    -> coverage_view.py

  Change something on the Price Comparison page
    -> price_view.py

  Change the sidebar menu itself, or which view loads first
    -> app.py


TROUBLESHOOTING — "ImportError: cannot import name 'X' from 'shared'"
---------------------------------------------------------------------
This means a view file expects something from shared.py (or data_cov.py
/ data_threat.py) that isn't actually there yet. Two things to check,
in order:

  1. Open the file directly on github.com and Ctrl+F for the missing
     name. If it's genuinely missing, add it and commit to the `main`
     branch (not just save locally — only a real commit+push changes
     what's deployed).

  2. If GitHub already has it (double-check via the file's "Raw" view)
     but the app STILL shows the same error — this is a known Streamlit
     Cloud quirk: auto-redeploy-on-push sometimes reruns your script
     without fully restarting the Python process, so an old cached
     copy of a module lingers in memory even though the file on disk
     is correct. Fix: go to Manage app -> "Reboot app" (a full restart,
     not the automatic "Updated app!" you see after every push). This
     clears the cache and re-imports everything fresh.

If reboot alone doesn't fix it, delete the app from Streamlit Cloud and
redeploy fresh from the repo as a last resort.

This is unrelated to Streamlit Cloud's free-tier resource limits —
that shows a completely different message ("This app has gone over its
resource limits"), not a Python traceback.
