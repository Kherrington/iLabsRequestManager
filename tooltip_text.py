"""
Hover-popup text for the iLab Service Request Manager.

Edit the strings below to change what the popups say — no changes to app.py
are needed.  Keys must match the ids used in app.py.

Markup:
    {{state_name}}   Highlights the word in that status's row colour,
                     e.g. {{completed}} or {{processing}}.
    \n               New line.
"""

# ── Toolbar buttons ──────────────────────────────────────────────────────────
BUTTONS = {
    "sync_all": (
        "Sync All\n"
        "Reloads the shared records cache (non-iLab data), then pulls the "
        "latest active requests from iLab."
    ),
    "sync_ilab": (
        "Sync iLab\n"
        "Fetches active requests from the iLab API only. Updates the iLab "
        "columns (ID through Status); your local records are kept."
    ),
    "sync_cache": (
        "Sync Records & Cache (non-iLab)\n"
        "Saves your pending edits, then reloads non-iLab columns from share cache file."
        "Class Schedule and Training data is synced with shared files."
         "Does not contact iLab."
    ),
    "sync_classes": (
        "Sync All Classes\n"
        "Two-way sync with the Microscope Intro Course Log: pulls sessions "
        "and students from the log, then adds any local students it lacks."
    ),
    "merge_classes": (
        "Merge Local File & Download History\n"
        "Merges the class_session.json kept in the app folder (visible only "
        "on this computer) into the shared class sessions, downloads the full "
        "history from the Intro Course Log, and adds anything the log lacks."
    ),
    "clear_all": (
        "Clear All Requests\n"
        "Empties the local cache. Run a sync afterwards to fetch fresh data."
    ),
    "import_csv": "Import an iLab export CSV into the table.",
    "export_csv": "Export the current requests to a CSV file.",
    "new_entry": "Add a request by hand that does not come from iLab.",
    "preferences": "Open Preferences (data file, cache sheet, theme, etc.).",
    "user_permissions": "Open the iLab user-permissions page in your browser.",
    "dark_mode": "Toggle between light and dark colour themes.",
}

# ── Merged group header cells above the table ────────────────────────────────
GROUPS = {
    "ilab": (
        "iLab data\n"
        "These columns come straight from iLab and are overwritten each time"
        "you Sync iLab. Edit them in iLab, not here."
    ),
    "records": (
        "Records (non-iLab)\n"
        "These columns are stored in the shared local records file, not in "
        "iLab. Edit them here; changes are shared via Sync Records & Cache."
        "Records are synced to excel files for training and class scheduling."
    ),
}

# ── Column headings (keys are the table column ids) ──────────────────────────
COLUMNS = {
    "request_id":    "iLab request ID. Double-click a row to open it in iLab.",
    "created_at":    "Date the request was submitted in iLab.",
    "owner_name":    "Person who submitted the request.",
    "pi_name":       "Principal investigator / lab the request belongs to.",
    "service_name":  "The iLab service that was requested.",
    "state": (
        "Request status from iLab:\n"
        "If the status is: {{proposed}}  {{requested}} {{financials_approved}}\n"
         "THEN the user has submitted and is waiting an email.\n" 
         "If the status is (or if you emailed then change to): {{processing}}\n"
         "THEN the core has accepted the request and is working on it.\n"
        "Change to {{completed}}  {{cancelled}}\n"
        "WHEN the work is finished or the request is cancelled.\n"
        "Click a heading to sort; click a Status cell to push a new state."
    ),
    "assigned_to":   "Team member responsible for this request. Click a cell to change.",
    "labels":        "Local tags for this request. Click a cell to edit.",
    "core_lab":      "Core Location of microscope (CALM or CVRI).",
    "microscope":    "Microscope the user is being trained on.",
    "training_date": "Date of the training session.",
    "training_day":  "Day of the week of the training.",
    "training_time": "Time of the training session.",
    "class_taken":   "Whether the user has completed the required class.",
}

# ── Status cells (hover a value in the Status column) ────────────────────────
STATES = {
    "proposed":    "{{proposed}}\nA quote/proposal has been made; waiting on the requester.",
    "requested":   "{{requested}}\nSubmitted by the user; not yet picked up by the core.",
    "draft":       "{{draft}}\nSaved by the user but not yet submitted.",
    "processing":  "{{processing}}\nThe core has accepted the request and is working on it.",
    "financials_approved": "{{financials_approved}}\nFunding has been approved; ready to schedule.",
    "needs_financial_reapproval":
        "{{needs_financial_reapproval}}\nThe request changed and funding must be approved again.",
    "completed":   "{{completed}}\nWork is finished. Completed requests drop off after a sync.",
    "cancelled":   "{{cancelled}}\nThe request was cancelled.",
    "core_disagreement": "{{core_disagreement}}\nThe core disagrees with the request as submitted.",
    "disagreement": "{{disagreement}}\nThere is an unresolved disagreement on this request.",
    "researcher_in_agreement": "{{researcher_in_agreement}}\nThe researcher has agreed to the core's changes.",
}

# ── Detail-panel tabs (keys are the tab labels, stripped of padding) ─────────
TABS = {
    "Request Info":   "Details pulled from iLab, plus local fields (assignee, labels, notes).",
    "Form Data":      "Answers the user gave on the iLab request form.",
    "Track Work":     "Checklist for pre- and post-training steps and email templates.",
    "Training":       "Record the training date, microscope and class status, then export to Records.",
    "Class Schedule": "Plan class sessions and the students attending each one.",
}
