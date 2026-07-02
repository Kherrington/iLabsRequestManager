# ── iLab connection ───────────────────────────────────────────────────────────
# The API token is read from the ILAB_TOKEN environment variable (see .env).
#
# IMPORTANT – ILAB_BASE_URL
ILAB_BASE_URL = "https://api.ilabsolutions.com"

# Your core's numeric ID — find it by running:  python get_cores.py
# or from the iLab URL: /service_center/show_external/<CORE_ID>
CORE_ID = 0   # ← replace with your core ID

# ── Local data file ───────────────────────────────────────────────────────────
DATA_FILE = "ilab_requests_cache.csv"

# ── Team members shown in the "Assigned To" dropdown ─────────────────────────
TEAM_MEMBERS = [
    "Staff Member 1",
    "Staff Member 2",
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
    "Microscope 1",
    "Microscope 2",
]

# ── Core facility options (selects which xlsx schedule file to update) ────────
CORE_OPTIONS = ["CORE1", "CORE2"]

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
    "researcher_in_agreement",
]

# ── Preferences file (xlsx paths, class charge IDs) ───────────────────────────
PREFS_FILE = "prefs.json"

# ── Local overrides ───────────────────────────────────────────────────────────
# Copy config.py to config_local.py and set your real values there.
# config_local.py is gitignored and will never be committed.
try:
    from config_local import *  # noqa: F401, F403
except ImportError:
    pass
