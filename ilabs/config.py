# ── iLab connection ───────────────────────────────────────────────────────────
# The API token is read from the ILAB_TOKEN environment variable (see .env).
#
# IMPORTANT – ILAB_BASE_URL:
#   The REST API is hosted on a SEPARATE server from the web UI.
#   ucsf.ilab.agilent.com is the web application; it does NOT serve /v1/ routes.
#
#   To find the correct API base URL, do ONE of the following:
#     1. Email iLab-support@agilent.com and ask for the API instance URL for UCSF.
#     2. In iLab, go to Administration → API Clients and look for an
#        "API Base URL" or "Instance URL" field on that page.
#     3. Ask your iLab institutional administrator.
#
#   Once you have it, replace the URL below (it will look something like
#   https://ucsf.ilab.agilent.com  OR  https://my.ilab.agilent.com  OR
#   a completely different hostname).
#
ILAB_BASE_URL = "https://api.ilabsolutions.com"

# CALM = Center for Advanced Light Microscopy, numeric ID confirmed from the
# UCSF iLab landing page (/service_center/show_external/5226).
CORE_ID = 5226

# ── Local data file ───────────────────────────────────────────────────────────
DATA_FILE = "ilab_requests_cache.csv"

# ── Team members shown in the "Assigned To" dropdown ─────────────────────────
TEAM_MEMBERS = [
    "Nico Stuurman",
    "Kari Herrington",
    "Julia Martin",
    "Micaela Lasser",
]

# ── Label / tag options shown as checkboxes ───────────────────────────────────
LABELS = [
    "Urgent",
    "Pending Sample",
    "Emailed and Waiting",
    "Refresher",
    "Switching Microscopes",
    "Follow-up Needed",
]

# ── Microscopes shown in the Training tab dropdown ────────────────────────────
MICROSCOPES = [
    "CSU-W1",
    "CSU-22",
    "CREST-C2",
    "AZ100",
    "6D Widefield",
    "TIME LAPSE",
    "QULIPP Time Lapse",
    "OMX",
    "TIRF-STORM",
    "SNOUTY",
    "C-TRAP",
    "CVRI-SD-1",
    "CVRI-SD-2",
]

# ── Core facility options (selects which xlsx schedule file to update) ────────
CORE_OPTIONS = ["CALM", "CVRI"]

# ── Training day dropdown ─────────────────────────────────────────────────────
TRAINING_DAYS = ["MON", "TUES", "WED", "THURS", "FRI", "SAT", "SUN"]

# ── "Class" iLab charge defaults ─────────────────────────────────────────────
# Find Service ID and Price ID by running:  python get_services.py
# Quantity 2 × $100 = $200 total.  Update via ⚙ Preferences in the app.
CLASS_SERVICE_ID = 0   # ← replace with actual service ID
CLASS_PRICE_ID   = 0   # ← replace with actual price ID
CLASS_QUANTITY   = 2
CLASS_UNIT_PRICE = 100.0

# ── States considered "active" — only these are fetched during Sync ──────────
# Remove a state from this list to stop syncing it; add it back to include it.
ACTIVE_STATES = [
    "proposed",
    "requested",
    "draft",
    "processing",
    "financials_approved",
    "financials_rejected",
    "needs_financial_reapproval",
]

# ── Preferences file (xlsx paths, class charge IDs) ───────────────────────────
PREFS_FILE = "prefs.json"
