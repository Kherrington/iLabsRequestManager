"""
iLab Service Request Manager
Run:  python app.py

Requirements: pip install requests
Set ILAB_TOKEN env var (or add to .env and load before running).
Set CORE_ID in config.py after running: python get_cores.py
"""

import json
import threading
import webbrowser
from datetime import datetime, timezone, date as _date
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import prefs as _prefs
from calendar_widget import CalendarPicker
from xlsx_export import append_training_row, HAS_OPENPYXL
from config import (
    CORE_ID, ILAB_BASE_URL, DATA_FILE, TEAM_MEMBERS, LABELS,
    MICROSCOPES, TRAINING_DAYS, CORE_OPTIONS, ACTIVE_STATES,
)
from data_store import DataStore
from ilabs_client import ILabClient, ILabError

# ── Colour palette for request states ────────────────────────────────────────
STATE_COLORS = {
    "proposed":                "#FFF9C4",
    "requested":               "#E8EAF6",
    "draft":                   "#F5F5F5",
    "processing":              "#BBDEFB",
    "financials_approved":     "#B2EBF2",
    "needs_financial_reapproval": "#FFE0B2",
    "completed":               "#C8E6C9",
    "cancelled":               "#E0E0E0",
    "core_disagreement":       "#FFCDD2",
    "disagreement":            "#FFCDD2",
}


STATE_COLORS_DARK = {
    "proposed":                    "#3d3500",
    "requested":                   "#1a1e40",
    "draft":                       "#303030",
    "processing":                  "#0d2a45",
    "financials_approved":         "#0a2e30",
    "needs_financial_reapproval":  "#3d2200",
    "completed":                   "#0a2d0a",
    "cancelled":                   "#2a2a2a",
    "core_disagreement":           "#3d0f0f",
    "disagreement":                "#3d0f0f",
}


def _lighten_hex(hex_color: str, factor: float = 0.4) -> str:
    """Return hex_color blended factor% toward white."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r = round(r + (255 - r) * factor)
    g = round(g + (255 - g) * factor)
    b = round(b + (255 - b) * factor)
    return f"#{r:02x}{g:02x}{b:02x}"


def _detect_core_from_form(form_data: dict) -> str:
    """Scan form-data keys and values for 'CALM' or 'CVRI'."""
    text = (
        " ".join(str(v) for v in form_data.values()) + " " +
        " ".join(form_data.keys())
    ).upper()
    if "CVRI" in text:
        return "CVRI"
    if "CALM" in text:
        return "CALM"
    return ""


class ILabManagerApp:
    """Main application window."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("iLab Service Request Manager — UCSF")
        self.root.geometry("1420x820")
        self.root.minsize(900, 600)

        _p = _prefs.get_prefs()
        _data_path = str(_p.get("data_file", "") or "").strip() or DATA_FILE
        self._data = DataStore(_data_path)
        self._client: ILabClient | None = None
        self._current_rec: dict | None = None
        self._sort_col = "created_at"
        self._sort_rev = True
        self._core_id: int | None = CORE_ID
        # Shared var used by both the Track Work and Training tabs
        self._class_taken_var = tk.BooleanVar()

        self._dark_mode = False

        self._build_ui()
        self._data.reload()          # pick up any changes written by other machines
        self._refresh_table()
        self._restore_last_sync()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Apply saved theme after UI is fully built
        _dark_pref = str(_p.get("dark_mode", "0")).strip() == "1"
        if _dark_pref:
            self._apply_theme(dark=True)

    # =========================================================================
    # UI construction
    # =========================================================================

    def _build_ui(self) -> None:
        self._build_toolbar()
        self._build_filter_bar()

        paned = ttk.PanedWindow(self.root, orient="vertical")
        paned.pack(fill="both", expand=True, padx=6, pady=(2, 4))
        self._build_table(paned)
        self._build_detail_panel(paned)

        self._build_status_bar()

    # ── Toolbar ───────────────────────────────────────────────────────────────

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self.root, padding=(6, 6, 6, 2))
        bar.pack(fill="x")

        # Sync status indicator — grey idle / yellow working / green ok / red error
        self._sync_indicator = tk.Frame(bar, width=18, height=18, bg="#C8C8C8",
                                        relief="flat")
        self._sync_indicator.pack_propagate(False)
        self._sync_indicator.pack(side="left", padx=(0, 2), pady=1)

        ttk.Button(bar, text="↻  Sync from iLab",      command=self._on_sync).pack(side="left", padx=2)
        ttk.Button(bar, text="Clear All Requests",      command=self._on_clear_all).pack(side="left", padx=2)
        ttk.Button(bar, text="Import iLab Export CSV…", command=self._on_import).pack(side="left", padx=2)
        ttk.Button(bar, text="Export to CSV…",          command=self._on_export).pack(side="left", padx=2)
        ttk.Button(bar, text="＋ New Entry",             command=self._on_add_manual_entry).pack(side="left", padx=2)
        ttk.Button(bar, text="👥 User Permissions",      command=self._on_user_permissions).pack(side="right", padx=2)
        ttk.Button(bar, text="⚙  Preferences",          command=self._on_open_preferences).pack(side="right", padx=(0, 4))
        self._dark_btn = ttk.Button(bar, text="🌙 Dark",  command=self._toggle_dark_mode)
        self._dark_btn.pack(side="right", padx=2)
        ttk.Button(bar, text="⟳  Sync NON-iLab", command=self._on_sync_cache).pack(side="right", padx=2)

        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=10, pady=2)

        ttk.Label(bar, text="Core ID:").pack(side="left")
        core_default = str(self._core_id) if self._core_id else ""
        self._core_id_var = tk.StringVar(value=core_default)
        ttk.Entry(bar, textvariable=self._core_id_var, width=8).pack(side="left", padx=4)
        ttk.Label(bar, text="(run get_cores.py to find yours)", foreground="#888").pack(side="left")

    # ── Filter bar ────────────────────────────────────────────────────────────

    def _build_filter_bar(self) -> None:
        bar = ttk.Frame(self.root, padding=(6, 2, 6, 4))
        bar.pack(fill="x")

        ttk.Label(bar, text="Search:").pack(side="left")
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._refresh_table())
        ttk.Entry(bar, textvariable=self._search_var, width=32).pack(side="left", padx=4)

        ttk.Label(bar, text="State:").pack(side="left", padx=(12, 0))
        self._state_filter = tk.StringVar(value="All")
        cb = ttk.Combobox(
            bar, textvariable=self._state_filter, width=18, state="readonly",
            values=["All", "proposed", "requested", "processing",
                    "financials_approved", "completed", "cancelled"],
        )
        cb.pack(side="left", padx=4)
        cb.bind("<<ComboboxSelected>>", lambda _: self._refresh_table())

        ttk.Label(bar, text="Assigned:").pack(side="left", padx=(12, 0))
        self._assigned_filter = tk.StringVar(value="All")
        cb2 = ttk.Combobox(
            bar, textvariable=self._assigned_filter, width=18, state="readonly",
            values=["All", "Unassigned"] + TEAM_MEMBERS,
        )
        cb2.pack(side="left", padx=4)
        cb2.bind("<<ComboboxSelected>>", lambda _: self._refresh_table())

        self._row_count_var = tk.StringVar()
        ttk.Label(bar, textvariable=self._row_count_var, foreground="#555").pack(side="right", padx=8)

    # ── Main table ────────────────────────────────────────────────────────────

    def _build_table(self, parent: ttk.PanedWindow) -> None:
        frame = ttk.Frame(parent)
        parent.add(frame, weight=3)

        cols = [
            "request_id", "created_at", "owner_name", "pi_name",
            "service_name", "state", "assigned_to", "labels",
            "core_lab", "microscope",
            "training_date", "training_day", "training_time",
            "class_taken",
        ]
        headers = {
            "request_id":    ("ID",          75),
            "created_at":    ("Submitted",    105),
            "owner_name":    ("Requester",   150),
            "pi_name":       ("Lab",         145),
            "service_name":  ("Service",     195),
            "state":         ("Status",      120),
            "assigned_to":   ("Assigned To", 125),
            "labels":        ("Labels",      135),
            "core_lab":      ("Core",         52),
            "microscope":    ("Microscope",  115),
            "training_date": ("Trng Date",    88),
            "training_day":  ("Day",          52),
            "training_time": ("Time",         68),
            "class_taken":   ("Class",        46),
        }

        self._tree = ttk.Treeview(frame, columns=cols, show="headings",
                                  selectmode="browse")
        for col, (header, width) in headers.items():
            self._tree.heading(col, text=header,
                               command=lambda c=col: self._sort_by(c))
            self._tree.column(col, width=width, anchor="w", minwidth=50)

        for state, color in STATE_COLORS.items():
            self._tree.tag_configure(state, background=color)
            self._tree.tag_configure(state + "_alt",
                                     background=_lighten_hex(color))

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right",  fill="y")
        self._tree.pack(fill="both", expand=True)

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self._tree.bind("<ButtonRelease-1>",  self._on_tree_click)
        self._tree.bind("<Double-1>",         self._on_open_in_ilab)

    # ── Detail panel ─────────────────────────────────────────────────────────

    def _build_detail_panel(self, parent: ttk.PanedWindow) -> None:
        outer = ttk.LabelFrame(parent, text="Request Details", padding=4)
        parent.add(outer, weight=2)

        self._build_quick_actions_bar(outer)

        self._notebook = ttk.Notebook(outer)
        self._notebook.pack(fill="both", expand=True)

        self._build_info_tab()
        self._build_form_tab()
        self._build_milestones_tab()
        self._build_training_tab()

    # ── Quick-actions bar (milestone buttons, always visible) ─────────────────

    def _build_quick_actions_bar(self, parent) -> None:
        self._qa_outer = ttk.Frame(parent)
        self._qa_outer.pack(fill="x", pady=(0, 4))

        ttk.Label(self._qa_outer, text="Track Work:",
                  font=("", 9, "bold")).pack(side="left", padx=(4, 8))

        # Scrollable canvas so many milestones don't overflow
        self._qa_canvas = tk.Canvas(self._qa_outer, height=28,
                                     highlightthickness=0)
        qa_hsb = ttk.Scrollbar(self._qa_outer, orient="horizontal",
                                command=self._qa_canvas.xview)
        self._qa_canvas.configure(xscrollcommand=qa_hsb.set)
        # Only show scrollbar when needed; pack canvas first so label stays left
        self._qa_canvas.pack(side="top", fill="x", expand=True)
        qa_hsb.pack(side="top", fill="x")

        self._qa_inner = ttk.Frame(self._qa_canvas)
        self._qa_win   = self._qa_canvas.create_window(
            (0, 0), window=self._qa_inner, anchor="nw")

        self._qa_inner.bind("<Configure>",
            lambda e: self._qa_canvas.configure(
                scrollregion=self._qa_canvas.bbox("all")))
        self._qa_canvas.bind("<Configure>",
            lambda e: self._qa_canvas.itemconfig(self._qa_win, width=e.width))

        ttk.Label(self._qa_inner,
                  text="← select a request",
                  foreground="#888").pack(side="left", padx=6, pady=4)

        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=(0, 4))

    def _build_info_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=10)
        self._notebook.add(tab, text="  Request Info  ")

        left  = ttk.Frame(tab)
        right = ttk.Frame(tab)
        left.pack(side="left", fill="both", expand=True, padx=(0, 16))
        right.pack(side="left", fill="both", expand=True)

        # ── Left: iLab read-only fields ───────────────────────────────────────
        ttk.Label(left, text="From iLab", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        self._info_vars: dict[str, tk.StringVar] = {}
        ilab_fields = [
            ("request_id",  "Request ID"),
            ("name",        "Name"),
            ("state",       "Status"),
            ("created_at",  "Submitted"),
            ("start_on",    "Start Date"),
            ("end_on",      "End Date"),
            ("completed_on","Completed"),
            ("owner_name",  "Requester"),
            ("owner_email", "Email"),
            ("pi_name",     "Lab / PI"),
            ("service_name","Service"),
        ]
        for i, (key, label) in enumerate(ilab_fields, start=1):
            ttk.Label(left, text=label + ":", anchor="e", width=13).grid(
                row=i, column=0, sticky="e", padx=4, pady=2)
            var = tk.StringVar()
            self._info_vars[key] = var
            ttk.Label(left, textvariable=var, anchor="w").grid(
                row=i, column=1, sticky="w", padx=4, pady=2)

        # State push
        sep_row = len(ilab_fields) + 2
        ttk.Separator(left, orient="horizontal").grid(
            row=sep_row, column=0, columnspan=2, sticky="ew", pady=8)
        ttk.Label(left, text="Push state:", anchor="e", width=13).grid(
            row=sep_row+1, column=0, sticky="e", padx=4)
        self._push_state_var = tk.StringVar()
        ttk.Combobox(
            left, textvariable=self._push_state_var, width=18, state="readonly",
            values=["proposed", "processing", "completed", "cancelled"],
        ).grid(row=sep_row+1, column=1, sticky="w", padx=4)
        ttk.Button(left, text="Push to iLab →",
                   command=self._on_push_state).grid(
            row=sep_row+2, column=1, sticky="w", padx=4, pady=4)

        # ── Right: local editable fields ──────────────────────────────────────
        ttk.Label(right, text="Local Fields", font=("", 10, "bold")).pack(
            anchor="w", pady=(0, 8))

        ttk.Label(right, text="Assigned To:").pack(anchor="w")
        self._assigned_var = tk.StringVar()
        ttk.Combobox(
            right, textvariable=self._assigned_var, width=24, state="readonly",
            values=[""] + TEAM_MEMBERS,
        ).pack(anchor="w", pady=(0, 10))

        ttk.Label(right, text="Labels:").pack(anchor="w")
        self._label_vars: dict[str, tk.BooleanVar] = {}
        lbl_frame = ttk.Frame(right)
        lbl_frame.pack(anchor="w", pady=(0, 10))
        for i, label in enumerate(LABELS):
            var = tk.BooleanVar()
            self._label_vars[label] = var
            ttk.Checkbutton(lbl_frame, text=label, variable=var).grid(
                row=i // 2, column=i % 2, sticky="w", padx=4, pady=1)

        ttk.Label(right, text="Notes:").pack(anchor="w")
        self._notes_text = tk.Text(right, height=4, width=34, wrap="word",
                                   font=("", 9))
        self._notes_text.pack(fill="x", pady=(0, 10))

        btn_frame = ttk.Frame(right)
        btn_frame.pack(anchor="w")
        ttk.Button(btn_frame, text="Save Local Changes",
                   command=self._on_save_local).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Open in iLab →",
                   command=self._on_open_in_ilab).pack(side="left")

    def _build_form_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=8)
        self._notebook.add(tab, text="  Form Data  ")

        # Scrollable two-column grid inside a canvas
        canvas = tk.Canvas(tab, highlightthickness=0)
        vsb = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)

        self._form_inner = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=self._form_inner,
                                             anchor="nw")

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        self._form_inner.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(
            int(-1 * (e.delta / 120)), "units"))

        self._form_canvas   = canvas
        self._form_placeholder = ttk.Label(
            self._form_inner,
            text="Select a request to view its form fields.",
            foreground="#888")
        self._form_placeholder.grid(row=0, column=0, padx=12, pady=12)

    def _build_milestones_tab(self) -> None:
        """Track Work tab — custom pre/post-training workflow checklist."""
        tab = ttk.Frame(self._notebook, padding=10)
        self._notebook.add(tab, text="  Track Work  ")

        # ── Workflow BooleanVars ───────────────────────────────────────────────
        self._wf_vars = {
            "wf_laser_safety":       tk.BooleanVar(),
            "wf_emailed":            tk.BooleanVar(),
            "wf_class_scheduled":    tk.BooleanVar(),
            "wf_not_required":       tk.BooleanVar(),
            "wf_training_scheduled": tk.BooleanVar(),
            "wf_post_email":         tk.BooleanVar(),
            "wf_post_listserve":     tk.BooleanVar(),
            "wf_post_approved":      tk.BooleanVar(),
            "wf_post_confirmed":     tk.BooleanVar(),
        }

        def _save_wf():
            """Auto-save when any checkbox is ticked."""
            if not self._current_rec:
                return
            req_id = self._current_rec["request_id"]
            fields = {k: ("1" if v.get() else "0") for k, v in self._wf_vars.items()}
            self._data.update_local_fields(req_id, **fields)
            self._current_rec.update(fields)
            self._update_quick_actions()
            # If "No" (wf_not_required) is checked, disable Class Taken; re-enable if unchecked
            already_submitted = self._current_rec.get("class_taken") == "1"
            if not already_submitted:
                no_class = self._wf_vars["wf_not_required"].get()
                chk_state = "disabled" if no_class else "normal"
                for chk in (getattr(self, "_class_taken_chk", None),
                            getattr(self, "_class_taken_chk_track", None)):
                    if chk:
                        chk.config(state=chk_state)

        def _chk(parent, key, label, url=None):
            """Build one checkbox row, with an optional hyperlink on the label."""
            row = ttk.Frame(parent)
            row.pack(anchor="w", pady=3)
            ttk.Checkbutton(row, variable=self._wf_vars[key],
                            command=_save_wf).pack(side="left")
            if url:
                lbl = ttk.Label(row, text=label,
                                foreground="#1565C0", cursor="hand2")
                lbl.pack(side="left")
                lbl.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))
            else:
                ttk.Label(row, text=label).pack(side="left")

        # ── Two-column layout ─────────────────────────────────────────────────
        cols = ttk.Frame(tab)
        cols.pack(fill="both", expand=True)

        # Pre-Training column
        pre = ttk.LabelFrame(cols, text="Pre-Training", padding=10)
        pre.pack(side="left", fill="both", expand=True, padx=(0, 6))

        _chk(pre, "wf_laser_safety",        "Laser Safety")
        _chk(pre, "wf_emailed",            "Emailed User")
        # Class? — Yes / No on one row
        class_row = ttk.Frame(pre)
        class_row.pack(anchor="w", pady=3)
        ttk.Label(class_row, text="Class?").pack(side="left", padx=(0, 6))
        ttk.Checkbutton(class_row, variable=self._wf_vars["wf_class_scheduled"],
                        command=_save_wf).pack(side="left")
        ttk.Label(class_row, text="Yes").pack(side="left", padx=(0, 10))
        ttk.Checkbutton(class_row, variable=self._wf_vars["wf_not_required"],
                        command=_save_wf).pack(side="left")
        ttk.Label(class_row, text="No").pack(side="left")

        # if yes → Class Taken (indented)
        ct_row = ttk.Frame(pre)
        ct_row.pack(anchor="w", pady=3)
        ttk.Label(ct_row, text="    if yes:").pack(side="left", padx=(0, 4))
        self._class_taken_chk_track = ttk.Checkbutton(
            ct_row, variable=self._class_taken_var,
            command=self._on_class_taken_toggle)
        self._class_taken_chk_track.pack(side="left")
        self._class_taken_lbl_track = ttk.Label(ct_row, text="Class Taken")
        self._class_taken_lbl_track.pack(side="left")

        _chk(pre, "wf_training_scheduled", "Training Scheduled")

        # Post-Training column
        post = ttk.LabelFrame(cols, text="Post-Training", padding=10)
        post.pack(side="left", fill="both", expand=True, padx=(6, 0))

        _chk(post, "wf_post_email",
             "Post-Training Email")
        _chk(post, "wf_post_listserve",
             "List Serve",
             url="https://listsrv.ucsf.edu/")
        _chk(post, "wf_post_approved",
             "Training Approved",
             url="https://ucsf.ilab.agilent.com/sc/5226/"
                 "center-for-advanced-light-microscopy/?tab=people")
        _chk(post, "wf_post_confirmed",
             "Confirmed in iLab")

    # ── Training tab ─────────────────────────────────────────────────────────

    def _build_training_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=10)
        self._notebook.add(tab, text="  Training  ")

        # Variables (_class_taken_var is shared with Track Work tab; created in __init__)
        self._training_core_var  = tk.StringVar()
        self._training_micro_var = tk.StringVar()
        self._training_date_var  = tk.StringVar()
        self._training_day_var   = tk.StringVar()
        self._training_time_var  = tk.StringVar()

        left  = ttk.Frame(tab)
        right = ttk.Frame(tab)
        left.pack(side="left", fill="both", expand=True, padx=(0, 18))
        right.pack(side="left", fill="both", expand=True)

        # ── Left: Export button at the top ───────────────────────────────────
        top_row = ttk.Frame(left)
        top_row.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        ttk.Button(top_row, text="Export to Records →",
                   command=self._on_export_to_schedule).pack(side="left", padx=(0, 10))
        self._exported_lbl = ttk.Label(top_row, text="", foreground="#2E7D32")
        self._exported_lbl.pack(side="left")

        ttk.Separator(left, orient="horizontal").grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        ttk.Label(left, text="Training Details", font=("", 10, "bold")).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 6))

        LW = 15   # label column width

        # Core Lab  (row offset +3 to sit below the new top export strip)
        ttk.Label(left, text="Core Lab:", anchor="e", width=LW).grid(
            row=3, column=0, sticky="e", padx=4, pady=3)
        core_row = ttk.Frame(left)
        core_row.grid(row=3, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        ttk.Combobox(core_row, textvariable=self._training_core_var,
                     values=CORE_OPTIONS, width=8, state="readonly").pack(side="left")
        ttk.Button(core_row, text="Auto-detect",
                   command=self._on_detect_core).pack(side="left", padx=(6, 0))

        # Microscope
        ttk.Label(left, text="Microscope:", anchor="e", width=LW).grid(
            row=4, column=0, sticky="e", padx=4, pady=3)
        ttk.Combobox(left, textvariable=self._training_micro_var,
                     values=MICROSCOPES, width=20).grid(
            row=4, column=1, columnspan=2, sticky="w", padx=4, pady=3)

        # Training Date
        ttk.Label(left, text="Training Date:", anchor="e", width=LW).grid(
            row=5, column=0, sticky="e", padx=4, pady=3)
        date_row = ttk.Frame(left)
        date_row.grid(row=5, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        ttk.Entry(date_row, textvariable=self._training_date_var, width=13).pack(side="left")
        ttk.Button(date_row, text="📅", width=3,
                   command=self._on_training_date_pick).pack(side="left", padx=(4, 0))

        # Training Day
        ttk.Label(left, text="Training Day:", anchor="e", width=LW).grid(
            row=6, column=0, sticky="e", padx=4, pady=3)
        ttk.Combobox(left, textvariable=self._training_day_var,
                     values=TRAINING_DAYS, width=10, state="readonly").grid(
            row=6, column=1, sticky="w", padx=4, pady=3)

        # Training Time
        ttk.Label(left, text="Training Time:", anchor="e", width=LW).grid(
            row=7, column=0, sticky="e", padx=4, pady=3)
        ttk.Entry(left, textvariable=self._training_time_var, width=14).grid(
            row=7, column=1, sticky="w", padx=4, pady=3)

        ttk.Separator(left, orient="horizontal").grid(
            row=8, column=0, columnspan=3, sticky="ew", pady=8)

        # Class Taken
        self._class_taken_chk = ttk.Checkbutton(
            left,
            text="Class Taken  (adds 2 × Class charge, $200 total, to iLab)",
            variable=self._class_taken_var,
            command=self._on_class_taken_toggle,
        )
        self._class_taken_chk.grid(row=9, column=0, columnspan=3, sticky="w", padx=4)
        self._class_status_lbl = ttk.Label(
            left, text="", foreground="#555", wraplength=340)
        self._class_status_lbl.grid(
            row=10, column=0, columnspan=3, sticky="w", padx=24, pady=(0, 4))

        ttk.Separator(left, orient="horizontal").grid(
            row=11, column=0, columnspan=3, sticky="ew", pady=6)

        # Save button at bottom
        ttk.Button(left, text="Save Training Info",
                   command=self._on_save_training).grid(
            row=12, column=0, columnspan=3, sticky="w", padx=4)

        # ── Right: guidance ───────────────────────────────────────────────────
        ttk.Label(right, text="How to use", font=("", 10, "bold")).pack(
            anchor="w", pady=(0, 8))
        guide = (
            "1. Core Lab (CALM / CVRI) — click Auto-detect to\n"
            "   scan form data, or set manually.  This selects\n"
            "   which xlsx file the export appends to.\n\n"
            "2. Microscope — choose from the dropdown\n"
            "   (edit MICROSCOPES in config.py to customise).\n\n"
            "3. Training Date — click 📅 to open the calendar;\n"
            "   click any day to fill the field.  Training Day\n"
            "   is set automatically from the chosen date.\n\n"
            "4. Training Time — type freely (e.g. 10:00 AM).\n\n"
            "5. Save Training Info, then Export to Records →\n"
            "   to append a row to the correct xlsx file.\n"
            "   Set xlsx paths in ⚙ Preferences.\n\n"
            "6. Class Taken — tick when the researcher has\n"
            "   attended a training class.  Adds 2 × Class\n"
            "   charge ($200) to iLab (configure Service ID\n"
            "   and Price ID in ⚙ Preferences)."
        )
        ttk.Label(right, text=guide, justify="left",
                  foreground="#444").pack(anchor="nw")

    # ── Status bar ────────────────────────────────────────────────────────────

    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self.root, relief="sunken", padding=(6, 2))
        bar.pack(fill="x", side="bottom")
        self._status_var = tk.StringVar(value="Ready — click ↻ Sync from iLab to load data.")
        ttk.Label(bar, textvariable=self._status_var, anchor="w").pack(side="left")
        self._last_sync_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self._last_sync_var,
                  anchor="e", foreground="#666").pack(side="right")

    # =========================================================================
    # Table management
    # =========================================================================

    def _refresh_table(self, *_) -> None:
        records = self._data.all_records()

        # ── Apply filters ─────────────────────────────────────────────────────
        q        = self._search_var.get().lower().strip()
        state_f  = self._state_filter.get()
        assign_f = self._assigned_filter.get()

        if q:
            def _match(r):
                haystack = (
                    r.get("request_id","") + " " +
                    r.get("name","") + " " +
                    r.get("owner_name","") + " " +
                    r.get("pi_name","") + " " +
                    r.get("service_name","")
                ).lower()
                return q in haystack
            records = [r for r in records if _match(r)]

        if state_f != "All":
            records = [r for r in records if r.get("state") == state_f]

        if assign_f == "Unassigned":
            records = [r for r in records if not r.get("assigned_to")]
        elif assign_f != "All":
            records = [r for r in records if r.get("assigned_to") == assign_f]

        # ── Sort ──────────────────────────────────────────────────────────────
        records.sort(key=lambda r: r.get(self._sort_col, "") or "",
                     reverse=self._sort_rev)

        # ── Repopulate treeview ───────────────────────────────────────────────
        self._tree.delete(*self._tree.get_children())
        for _row_idx, rec in enumerate(records):
            state = rec.get("state", "")
            raw = (rec.get("created_at") or "")
            # Strip milliseconds and timezone so fromisoformat works on all
            # Python versions (submitted_at looks like "2024-12-20T17:39:44.000-05:00")
            raw = raw[:19].replace("T", " ")
            try:
                date = datetime.fromisoformat(raw).strftime("%d %b %Y")
            except ValueError:
                date = raw[:10]
            self._tree.insert(
                "", "end",
                iid=rec["request_id"],
                values=(
                    rec.get("request_id"),
                    date,
                    rec.get("owner_name", ""),
                    rec.get("pi_name", ""),
                    rec.get("service_name", ""),
                    state,
                    rec.get("assigned_to", ""),
                    rec.get("labels", ""),
                    rec.get("core_lab", ""),
                    rec.get("microscope", ""),
                    rec.get("training_date", ""),
                    rec.get("training_day", ""),
                    rec.get("training_time", ""),
                    "☑" if rec.get("class_taken") == "1" else "☐",
                ),
                tags=(state + ("_alt" if _row_idx % 2 else ""),),
            )

        self._row_count_var.set(f"{len(records)} request(s)")

    def _sort_by(self, col: str) -> None:
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = False
        self._refresh_table()

    # =========================================================================
    # Detail panel population
    # =========================================================================

    def _load_detail(self, rec: dict) -> None:
        req_id    = rec.get("request_id", "")
        form_data = json.loads(rec.get("form_data") or "{}")

        # ── Auto-fill core_lab if blank ───────────────────────────────────────
        if not rec.get("core_lab"):
            core = _detect_core_from_form(form_data)
            if core:
                self._data.update_local_fields(req_id, core_lab=core)
                rec["core_lab"] = core
                # update the tree cell immediately
                try:
                    self._tree.set(req_id, "core_lab", core)
                except Exception:
                    pass

        # ── Info tab ──────────────────────────────────────────────────────────
        for key, var in self._info_vars.items():
            var.set(rec.get(key, "") or "")

        self._assigned_var.set(rec.get("assigned_to", ""))

        active = set((rec.get("labels") or "").split(","))
        for label, var in self._label_vars.items():
            var.set(label in active)

        self._notes_text.delete("1.0", "end")
        self._notes_text.insert("1.0", rec.get("local_notes", ""))

        # ── Form data tab ─────────────────────────────────────────────────────
        for w in self._form_inner.winfo_children():
            w.destroy()

        form_data: dict = json.loads(rec.get("form_data") or "{}")
        if form_data:
            ttk.Label(self._form_inner, text="Field", font=("", 9, "bold"),
                      width=36, anchor="w").grid(row=0, column=0, padx=6, pady=(4,2), sticky="w")
            ttk.Label(self._form_inner, text="Value", font=("", 9, "bold"),
                      anchor="w").grid(row=0, column=1, padx=6, pady=(4,2), sticky="w")
            ttk.Separator(self._form_inner, orient="horizontal").grid(
                row=1, column=0, columnspan=2, sticky="ew", padx=6)
            for i, (fname, fval) in enumerate(form_data.items(), start=2):
                ttk.Label(self._form_inner, text=fname, anchor="nw",
                          wraplength=260).grid(
                    row=i, column=0, padx=6, pady=2, sticky="nw")
                ttk.Label(self._form_inner, text=str(fval), anchor="nw",
                          wraplength=400).grid(
                    row=i, column=1, padx=6, pady=2, sticky="nw")
        else:
            ttk.Label(self._form_inner,
                      text="No form data loaded.\nSync from iLab to fetch form fields.",
                      foreground=ttk.Style().lookup("TLabel", "foreground") or "#888",
                      ).grid(row=0, column=0, padx=12, pady=12)

        self._form_canvas.configure(scrollregion=self._form_canvas.bbox("all"))

        # ── Track Work tab (workflow) + quick-actions bar ────────────────────
        for key, var in self._wf_vars.items():
            var.set(rec.get(key, "0") == "1")
        self._update_quick_actions()

        # ── Training tab ──────────────────────────────────────────────────────
        self._training_core_var.set(rec.get("core_lab", ""))
        self._training_micro_var.set(rec.get("microscope", ""))
        self._training_date_var.set(rec.get("training_date", ""))
        self._training_day_var.set(rec.get("training_day", ""))
        self._training_time_var.set(rec.get("training_time", ""))
        taken = rec.get("class_taken", "0") == "1"
        self._class_taken_var.set(taken)
        self._update_class_taken_ui(taken)
        if not taken and rec.get("wf_not_required", "0") == "1":
            for chk in (getattr(self, "_class_taken_chk", None),
                        getattr(self, "_class_taken_chk_track", None)):
                if chk:
                    chk.config(state="disabled")
        exported = rec.get("schedule_exported", "0") == "1"
        self._exported_lbl.config(
            text="✓ Exported to schedule" if exported else "Not yet exported")


    # =========================================================================
    # Event handlers
    # =========================================================================

    def _update_class_taken_ui(self, submitted: bool) -> None:
        """Grey out / re-enable both Class Taken checkboxes based on submission state."""
        chk_state = "disabled" if submitted else "normal"
        for chk in (getattr(self, "_class_taken_chk", None),
                    getattr(self, "_class_taken_chk_track", None)):
            if chk:
                chk.config(state=chk_state)
        if hasattr(self, "_class_taken_lbl_track"):
            self._class_taken_lbl_track.config(
                text="Class Taken — charge submitted" if submitted else "Class Taken",
                foreground="#9E9E9E" if submitted else "",
            )
        if hasattr(self, "_class_status_lbl"):
            self._class_status_lbl.config(
                text="✓ Charge submitted to iLab — cannot be re-submitted." if submitted else "",
                foreground="#9E9E9E" if submitted else "#555",
            )

    # =========================================================================
    # Quick-actions bar population
    # =========================================================================

    def _update_quick_actions(self) -> None:
        """Rebuild the workflow-progress strip above the tabs."""
        for w in self._qa_inner.winfo_children():
            w.destroy()

        if not self._current_rec:
            ttk.Label(self._qa_inner, text="← select a request",
                      foreground="#888").pack(side="left", padx=6, pady=4)
            self._qa_canvas.configure(scrollregion=self._qa_canvas.bbox("all"))
            return

        rec = self._current_rec

        _PRE = [
            ("wf_laser_safety",       "Laser Safety"),
            ("wf_emailed",            "Emailed"),
            ("wf_class_scheduled",    "Class Sched"),
            ("wf_not_required",       "No Class"),
            ("class_taken",           "Class Taken"),
            ("wf_training_scheduled", "Trng Sched"),
        ]
        _POST = [
            ("wf_post_email",     "Post Email"),
            ("wf_post_listserve", "Listserve"),
            ("wf_post_approved",  "Approved"),
            ("wf_post_confirmed", "Confirmed"),
        ]

        def _chip(label: str, done: bool) -> None:
            fg = "#2E7D32" if done else "#9E9E9E"
            pfx = "✓ " if done else "○ "
            ttk.Label(self._qa_inner, text=pfx + label,
                      foreground=fg, font=("", 8)).pack(
                side="left", padx=3, pady=3)

        ttk.Label(self._qa_inner, text="Pre:",
                  font=("", 8, "bold")).pack(side="left", padx=(6, 2), pady=3)
        for key, lbl in _PRE:
            _chip(lbl, rec.get(key, "0") == "1")

        ttk.Label(self._qa_inner, text="  |  Post:",
                  font=("", 8, "bold")).pack(side="left", padx=(4, 2), pady=3)
        for key, lbl in _POST:
            _chip(lbl, rec.get(key, "0") == "1")

        self._qa_canvas.configure(scrollregion=self._qa_canvas.bbox("all"))

    # =========================================================================
    # Inline cell editing
    # =========================================================================

    # Columns that open a dropdown on click (field → values list)
    _EDITABLE_COLS: dict = {}   # populated in __init__ after config is imported

    def _on_tree_click(self, event) -> None:
        """Open an inline editor when clicking an editable cell in an already-selected row."""
        region = self._tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        row_id = self._tree.identify_row(event.y)
        if not row_id:
            return
        # Only edit on the row that is already selected (second click)
        sel = self._tree.selection()
        if not sel or sel[0] != row_id:
            return
        col_id  = self._tree.identify_column(event.x)
        col_idx = int(col_id[1:]) - 1
        cols    = self._tree["columns"]
        if col_idx < 0 or col_idx >= len(cols):
            return
        col_name = cols[col_idx]

        if col_name == "state":
            self._show_cell_combo(
                row_id, col_id, "state",
                [
                    "completed",            # most common action — top of list
                    "processing",
                    "proposed",
                    "requested",
                    "financials_approved",
                    "needs_financial_reapproval",
                    "cancelled",
                ],
                on_commit=self._on_push_state_inline,
            )
        elif col_name == "assigned_to":
            self._show_cell_combo(row_id, col_id, "assigned_to",
                                  [""] + TEAM_MEMBERS)
        elif col_name == "core_lab":
            self._show_cell_combo(row_id, col_id, "core_lab", [""] + CORE_OPTIONS)
        elif col_name == "microscope":
            self._show_cell_combo(row_id, col_id, "microscope", [""] + MICROSCOPES)
        elif col_name == "training_day":
            self._show_cell_combo(row_id, col_id, "training_day", [""] + TRAINING_DAYS)
        elif col_name == "training_time":
            self._show_cell_entry(row_id, col_id, "training_time")
        elif col_name == "labels":
            self._show_labels_popup(row_id, event.x_root, event.y_root)
        elif col_name == "training_date":
            self._on_training_date_pick()
        elif col_name == "class_taken":
            # Toggle directly without needing the detail panel checkbox
            if self._current_rec and self._current_rec.get("request_id") == row_id:
                if self._current_rec.get("class_taken", "0") == "1":
                    return  # charge already submitted — block re-click
                new_state = self._current_rec.get("class_taken", "0") != "1"
                self._class_taken_var.set(new_state)
                self._on_class_taken_toggle()

    def _show_cell_combo(self, row_id: str, col_id: str,
                         field: str, values: list,
                         on_commit=None) -> None:
        """
        Overlay a Combobox on a Treeview cell.
        on_commit(row_id, value) is called instead of the default local-save
        when provided (used for state pushes to iLab).
        """
        bbox = self._tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        current = self._tree.set(row_id, col_id)
        var     = tk.StringVar(value=current)
        combo   = ttk.Combobox(self._tree, textvariable=var,
                               values=values, state="readonly", font=("", 9))
        combo.place(x=x, y=y, width=max(w, 120), height=h + 2)
        combo.focus_set()
        combo.event_generate("<Down>")

        def _commit(*_):
            val = var.get()
            combo.place_forget()
            combo.destroy()
            if on_commit:
                on_commit(row_id, val)
                return
            rec = self._data.get_record(row_id)
            if rec is None:
                return
            self._data.update_local_fields(row_id, **{field: val})
            rec[field] = val
            self._tree.set(row_id, col_id, val)
            if field == "assigned_to":
                self._assigned_var.set(val)
            elif field == "core_lab":
                self._training_core_var.set(val)
            elif field == "microscope":
                self._training_micro_var.set(val)
            self._set_status(f"Updated {field} for request {row_id}.")

        def _cancel(*_):
            combo.place_forget()
            combo.destroy()

        combo.bind("<<ComboboxSelected>>", _commit)
        combo.bind("<Escape>",             _cancel)
        combo.bind("<FocusOut>",           _cancel)

    def _show_labels_popup(self, row_id: str, x_root: int, y_root: int) -> None:
        """Floating checkbox popup for multi-value Labels editing."""
        rec = self._data.get_record(row_id)
        if rec is None:
            return
        active = set(filter(None, (rec.get("labels") or "").split(",")))

        popup = tk.Toplevel(self.root)
        popup.title("Labels")
        popup.transient(self.root)
        popup.resizable(False, False)
        popup.geometry(f"+{x_root}+{y_root}")

        frame = ttk.Frame(popup, padding=8)
        frame.pack()
        ttk.Label(frame, text="Labels", font=("", 9, "bold")).pack(anchor="w", pady=(0, 4))

        chk_vars: dict[str, tk.BooleanVar] = {}
        for lbl in LABELS:
            v = tk.BooleanVar(value=lbl in active)
            chk_vars[lbl] = v
            ttk.Checkbutton(frame, text=lbl, variable=v).pack(anchor="w")

        def _save():
            selected = ",".join(l for l, v in chk_vars.items() if v.get())
            self._data.update_local_fields(row_id, labels=selected)
            rec["labels"] = selected
            self._tree.set(row_id, "labels", selected)
            # Sync the checkboxes in the Request Info tab
            active_set = set(filter(None, selected.split(",")))
            for lbl, var in self._label_vars.items():
                var.set(lbl in active_set)
            popup.destroy()
            self._set_status(f"Labels updated for request {row_id}.")

        btn = ttk.Frame(frame)
        btn.pack(fill="x", pady=(8, 0))
        ttk.Button(btn, text="Save",   command=_save).pack(side="left", padx=(0, 4))
        ttk.Button(btn, text="Cancel", command=popup.destroy).pack(side="left")
        popup.bind("<Escape>", lambda e: popup.destroy())
        popup.focus_set()

    def _on_row_select(self, _event=None) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        rec = self._data.get_record(sel[0])
        if rec:
            self._current_rec = rec
            self._load_detail(rec)

    def _on_save_local(self) -> None:
        if not self._current_rec:
            return
        req_id      = self._current_rec["request_id"]
        active_lbls = ",".join(l for l, v in self._label_vars.items() if v.get())
        notes       = self._notes_text.get("1.0", "end-1c")
        assigned    = self._assigned_var.get()

        self._data.update_local_fields(req_id,
                                       assigned_to=assigned,
                                       labels=active_lbls,
                                       local_notes=notes)
        # Update in-memory record so the table refreshes correctly
        self._current_rec.update(assigned_to=assigned,
                                 labels=active_lbls,
                                 local_notes=notes)
        self._refresh_table()
        try:
            self._tree.selection_set(req_id)
            self._tree.see(req_id)
        except tk.TclError:
            pass
        self._set_status(f"Saved local changes for request {req_id}.")

    def _on_push_state(self) -> None:
        if not self._current_rec:
            return
        req_id    = self._current_rec["request_id"]
        new_state = self._push_state_var.get()
        if not new_state:
            messagebox.showwarning("No State Selected",
                                   "Choose a state from the dropdown first.")
            return

        # Local-only records: update locally, skip iLab API entirely
        if self._current_rec.get("local_only") == "1":
            if not messagebox.askyesno(
                    "Confirm",
                    f"Set state to '{new_state}' for local entry {req_id}?\n"
                    "(This entry is not linked to iLab.)"):
                return
            self._data.update_field(req_id, "state", new_state)
            self._current_rec["state"] = new_state
            self._info_vars["state"].set(new_state)
            self._refresh_table()
            self._set_status(f"Status → '{new_state}' (local only).")
            if new_state == "completed":
                self.root.after(100, lambda: self._auto_export_on_complete(req_id))
            return

        core_id = self._get_core_id()
        if core_id is None:
            return
        if not messagebox.askyesno(
                "Confirm Push",
                f"Push state '{new_state}' to iLab for request #{req_id}?"):
            return

        def worker():
            try:
                client = self._get_client()
                client.update_service_request(core_id, int(req_id), state=new_state)
                self._data.update_field(req_id, "state", new_state)
                self._current_rec["state"] = new_state
                self.root.after(0, lambda: self._info_vars["state"].set(new_state))
                self.root.after(0, self._refresh_table)
                self.root.after(0, lambda: self._set_status(
                    f"Status → '{new_state}' pushed to iLab for request {req_id}."))
                if new_state == "completed":
                    self.root.after(100, lambda: self._auto_export_on_complete(req_id))
            except ILabError as exc:
                print(f"[iLab API Error] push state={new_state} req={req_id}: {exc}")
                self.root.after(0, lambda e=exc: messagebox.showerror("iLab API Error", str(e)))
            except Exception as exc:
                print(f"[Error] push state={new_state} req={req_id}: {exc}")
                self.root.after(0, lambda e=exc: messagebox.showerror("Error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_milestone_action(self, request_id: str, milestone_id,
                             milestones: list, action: str) -> None:
        core_id = self._get_core_id()
        if core_id is None:
            return

        now_utc = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        field   = "started_on" if action == "started" else "completed_on"

        def worker():
            try:
                client = self._get_client()
                client.update_milestone(core_id, int(request_id), milestone_id,
                                        **{field: now_utc})
                # Update local cache
                for ms in milestones:
                    if ms.get("id") == milestone_id:
                        ms[field] = now_utc
                        break
                self._data.update_milestones(request_id, milestones)
                if self._current_rec and self._current_rec.get("request_id") == request_id:
                    self._current_rec["milestones_data"] = json.dumps(milestones)
                self.root.after(0, lambda: self._set_status(
                    f"Milestone {milestone_id} marked {action}."))
            except ILabError as exc:
                self.root.after(0, lambda e=exc: messagebox.showerror("iLab API Error", str(e)))
            except Exception as exc:
                self.root.after(0, lambda e=exc: messagebox.showerror("Error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_sync_cache(self) -> None:
        """Reload the shared CSV from disk to pick up changes from other machines."""
        self._data.reload()
        self._refresh_table()
        self._set_status("Cache reloaded from disk.")

    def _on_close(self) -> None:
        """Save current state to the shared CSV before exiting."""
        self._data.save()
        self.root.destroy()

    def _on_sync(self) -> None:
        core_id = self._get_core_id()
        if core_id is None:
            return
        self._set_status("Syncing with iLab…")
        self._set_sync_indicator("working")

        def worker():
            try:
                client = self._get_client()
                count = self._data.sync(
                    core_id, client,
                    on_progress=lambda m: self.root.after(0, lambda msg=m: self._set_status(msg)),
                    states=ACTIVE_STATES,
                )
                when = datetime.now().strftime("%b %d  %I:%M %p")
                self.root.after(0, lambda: self._set_sync_indicator("ok"))
                self.root.after(0, lambda: self._set_status(
                    f"Sync complete — {count} active request(s) loaded."))
                self.root.after(0, lambda t=when: self._set_last_sync(t))
                self.root.after(0, self._refresh_table)
            except ILabError as exc:
                msg = str(exc)
                if "404" in msg:
                    msg = (
                        "HTTP 404 — the API server was not found at:\n"
                        f"  {ILAB_BASE_URL}\n\n"
                        "The iLab REST API is hosted on a separate server from the web UI.\n\n"
                        "Steps to fix:\n"
                        "  1. Email iLab-support@agilent.com for the UCSF API instance URL.\n"
                        "  2. Or check Administration → API Clients in iLab for the URL.\n"
                        "  3. Then update ILAB_BASE_URL in config.py and restart."
                    )
                self.root.after(0, lambda: self._set_sync_indicator("error"))
                self.root.after(0, lambda e=msg: messagebox.showerror("iLab API Error", e))
            except Exception as exc:
                self.root.after(0, lambda: self._set_sync_indicator("error"))
                self.root.after(0, lambda e=exc: messagebox.showerror("Sync Error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_export(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile="ilab_requests_export.csv",
        )
        if not path:
            return
        try:
            self._data.export_expanded(path)
            self._set_status(f"Exported to {path}")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    def _on_add_manual_entry(self) -> None:
        """Open a dialog to create a local-only training record."""
        dlg = tk.Toplevel(self.root)
        dlg.title("New Manual Entry")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=16)
        frame.pack(fill="both", expand=True)

        fields_def = [
            ("Requester Name",  "owner_name",    "entry",   None),
            ("PI / Lab",        "pi_name",        "entry",   None),
            ("Service",         "service_name",   "entry",   None),
            ("Assigned To",     "assigned_to",    "combo",   [""] + TEAM_MEMBERS),
            ("Core Lab",        "core_lab",       "combo",   [""] + CORE_OPTIONS),
            ("Microscope",      "microscope",     "combo",   [""] + MICROSCOPES),
            ("Training Date",   "training_date",  "entry",   None),
            ("Training Day",    "training_day",   "combo",   [""] + TRAINING_DAYS),
            ("Training Time",   "training_time",  "entry",   None),
            ("Notes",           "local_notes",    "entry",   None),
        ]

        vars_: dict = {}
        for i, (label, key, widget_type, options) in enumerate(fields_def):
            ttk.Label(frame, text=label + ":").grid(row=i, column=0, sticky="e", padx=4, pady=3)
            var = tk.StringVar()
            vars_[key] = var
            if widget_type == "combo":
                w = ttk.Combobox(frame, textvariable=var, values=options,
                                 width=28, state="readonly")
            else:
                w = ttk.Entry(frame, textvariable=var, width=30)
            w.grid(row=i, column=1, sticky="w", padx=4, pady=3)

        btn_row = ttk.Frame(frame)
        btn_row.grid(row=len(fields_def), column=0, columnspan=2, pady=(12, 0))

        def _save():
            rec_fields = {k: v.get().strip() for k, v in vars_.items()}
            req_id = self._data.add_manual_record(rec_fields)
            self._refresh_table()
            # Select the new row
            try:
                self._tree.selection_set(req_id)
                self._tree.see(req_id)
            except tk.TclError:
                pass
            self._set_status(f"Manual entry {req_id} created.")
            dlg.destroy()

        ttk.Button(btn_row, text="Save", command=_save).pack(side="left", padx=6)
        ttk.Button(btn_row, text="Cancel", command=dlg.destroy).pack(side="left", padx=6)

    def _on_import(self) -> None:
        """Import a CSV exported from the iLab web UI (View All Requests → Export)."""
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Import iLab Export CSV",
        )
        if not path:
            return
        try:
            result = self._data.import_from_ilab_export(path)
        except Exception as exc:
            messagebox.showerror("Import Error", str(exc))
            return

        self._refresh_table()
        n       = result["imported"]
        skipped = result["skipped"]
        fname   = Path(path).name

        if n > 0:
            self._set_status(
                f"Imported {n} request(s) from {fname}"
                + (f"  ({skipped} row(s) skipped — no request ID)" if skipped else "")
                + "  — use ↻ Sync to fetch form data once the API is available."
            )
            return

        # ── Nothing imported — show a diagnostic window ───────────────────────
        ImportDiagnosticDialog(self.root, fname, result)

    # ── Training tab handlers ─────────────────────────────────────────────────

    def _on_detect_core(self) -> None:
        if not self._current_rec:
            return
        form_data = json.loads(self._current_rec.get("form_data") or "{}")
        core = _detect_core_from_form(form_data)
        if core:
            self._training_core_var.set(core)
            self._set_status(f"Core auto-detected: {core}")
        else:
            self._set_status("Could not auto-detect core from form data — set manually.")

    def _on_training_date_pick(self) -> None:
        raw = self._training_date_var.get().strip()
        initial = None
        try:
            initial = _date.fromisoformat(raw)
        except ValueError:
            pass
        CalendarPicker(self.root, self._on_training_date_selected, initial)

    def _on_training_date_selected(self, date_str: str) -> None:
        self._training_date_var.set(date_str)
        day_str = ""
        try:
            d = _date.fromisoformat(date_str)
            weekday = d.weekday()          # 0=Mon … 6=Sun
            if weekday < len(TRAINING_DAYS):
                day_str = TRAINING_DAYS[weekday]
                self._training_day_var.set(day_str)
        except ValueError:
            pass

        if not self._current_rec:
            return
        req_id = self._current_rec["request_id"]
        fields = {"training_date": date_str}
        if day_str:
            fields["training_day"] = day_str
        self._data.update_local_fields(req_id, **fields)
        self._current_rec.update(fields)
        # Reflect immediately in the table
        try:
            self._tree.set(req_id, "training_date", date_str)
            if day_str:
                self._tree.set(req_id, "training_day", day_str)
        except Exception:
            pass

    def _on_save_training(self) -> None:
        if not self._current_rec:
            return
        req_id = self._current_rec["request_id"]
        fields = {
            "core_lab":      self._training_core_var.get(),
            "microscope":    self._training_micro_var.get(),
            "training_date": self._training_date_var.get(),
            "training_day":  self._training_day_var.get(),
            "training_time": self._training_time_var.get(),
        }
        self._data.update_local_fields(req_id, **fields)
        self._current_rec.update(fields)
        self._refresh_table()
        try:
            self._tree.selection_set(req_id)
            self._tree.see(req_id)
        except tk.TclError:
            pass
        self._set_status(f"Training info saved for request {req_id}.")

    def _on_class_taken_toggle(self) -> None:
        if not self._current_rec:
            return
        checked = self._class_taken_var.get()
        req_id  = self._current_rec["request_id"]

        if not checked:
            # Unchecking — clear locally only (iLab charges are NOT deleted)
            self._data.update_local_fields(req_id, class_taken="0")
            self._current_rec["class_taken"] = "0"
            self._class_status_lbl.config(text="")
            self._refresh_table()
            return

        # Checking — attempt to add charge via API
        p = _prefs.get_prefs()
        svc_id     = str(p.get("class_service_id",  "") or "").strip()
        price_id   = str(p.get("class_price_id",    "") or "").strip()
        qty        = int(str(p.get("class_quantity",    "2")   or "2"))
        unit_price = float(str(p.get("class_unit_price", "100") or "100"))
        total      = qty * unit_price

        # ── Configurable cap (set in ⚙ Preferences → Max Charge) ────────────────
        _MAX_CHARGE = float(str(p.get("max_charge", "200") or "200"))
        if total > _MAX_CHARGE:
            self._class_taken_var.set(False)
            messagebox.showerror(
                "Charge Limit Exceeded",
                f"This charge would total  ${total:,.2f}  "
                f"({qty} × ${unit_price:.2f}).\n\n"
                f"The Class charge is capped at ${_MAX_CHARGE:,.2f}.\n\n"
                "Adjust Quantity or Unit Price in  ⚙ Preferences  and try again.",
            )
            return

        if not svc_id or not price_id or svc_id == "0" or price_id == "0":
            # Charge IDs not configured — save locally with a reminder
            self._data.update_local_fields(req_id, class_taken="1")
            self._current_rec["class_taken"] = "1"
            self._class_status_lbl.config(
                text="⚠ Saved locally. Configure Class Service ID and Price ID "
                     "in ⚙ Preferences to submit the charge to iLab.")
            self._refresh_table()
            return

        core_id = self._get_core_id()
        if core_id is None:
            self._class_taken_var.set(False)
            return

        self._class_status_lbl.config(text="Submitting charge to iLab…")

        def worker():
            try:
                client  = self._get_client()
                result  = client.add_charges(core_id, int(req_id), [{
                    "quantity":   qty,
                    "price_id":   int(price_id),
                    "service_id": int(svc_id),
                    "note":       "Class attendance",
                }])
                # Mark each new charge as completed
                raw_charges = result.get("charges") or []
                if isinstance(raw_charges, dict):
                    raw_charges = [raw_charges]
                for ch in raw_charges:
                    ch_id = ch.get("id")
                    if ch_id:
                        try:
                            client.update_charge(
                                core_id, int(req_id), ch_id,
                                status="completed",
                                billing_status="ready_to_bill",
                            )
                        except Exception:
                            pass

                self._data.update_local_fields(req_id, class_taken="1")
                self._current_rec["class_taken"] = "1"
                self.root.after(0, lambda: self._update_class_taken_ui(True))
                self.root.after(0, lambda: self._class_status_lbl.config(
                    text=f"✓ Charge submitted to iLab — cannot be re-submitted.",
                    foreground="#9E9E9E"))
                self.root.after(0, self._refresh_table)
                self.root.after(0, lambda: self._set_status(
                    f"Class charge added to request {req_id} in iLab."))
            except Exception as exc:
                self.root.after(0, lambda: self._class_taken_var.set(False))
                self.root.after(0, lambda: self._class_status_lbl.config(text=""))
                self.root.after(0, lambda e=exc: messagebox.showerror("Charge Error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_export_to_schedule(self, silent: bool = False) -> None:
        """
        Append the current record to the appropriate training schedule xlsx.
        silent=True suppresses the success dialog (used for auto-export on completion).
        """
        if not self._current_rec:
            if not silent:
                messagebox.showwarning("No Selection", "Select a request first.")
            return

        # Save any unsaved training fields first
        self._on_save_training()

        req_id = self._current_rec["request_id"]
        core   = self._training_core_var.get().strip().upper()
        if not core:
            if not silent:
                messagebox.showwarning(
                    "Core Lab Not Set",
                    "Set the Core Lab (CALM or CVRI) before exporting.\n\n"
                    "Use the Auto-detect button or choose manually.",
                )
            else:
                self._set_status(
                    f"⚠ Request {req_id} marked complete but Core Lab not set — "
                    "export to schedule manually from the Training tab.")
            return

        p = _prefs.get_prefs()
        xlsx_path = str(p.get(f"{core.lower()}_xlsx", "") or "").strip()
        if not xlsx_path:
            if not silent:
                messagebox.showwarning(
                    "xlsx Path Not Configured",
                    f"No {core} schedule xlsx file path is set.\n\n"
                    "Open ⚙ Preferences and browse to your\n"
                    f"Training Schedule_{core}_CURRENT.xlsx file.",
                )
            else:
                self._set_status(
                    f"⚠ Request {req_id} marked complete — set {core} xlsx path "
                    "in ⚙ Preferences to enable auto-export.")
            return

        if not HAS_OPENPYXL:
            if not silent:
                messagebox.showerror(
                    "openpyxl Required",
                    "Install openpyxl to export to xlsx:\n\n    pip install openpyxl",
                )
            return

        sheet_name = str(p.get(f"{core.lower()}_sheet", "") or "").strip()

        try:
            result    = append_training_row(self._current_rec, xlsx_path,
                                            sheet_name=sheet_name)
            # Mark as exported so auto-export doesn't run again
            self._data.update_local_fields(req_id, schedule_exported="1")
            self._current_rec["schedule_exported"] = "1"
            self._exported_lbl.config(text="✓ Exported to records")
            fname     = Path(xlsx_path).name
            sheet_lbl = f" → '{sheet_name}'" if sheet_name else ""
            n_written = len(result.get("written", {}))
            n_empty   = len(result.get("empty", []))
            self._set_status(
                f"Row appended to {fname}{sheet_lbl} ({core})  —  "
                f"{n_written} field(s) written, {n_empty} column(s) unmatched.")
            if not silent:
                written_str = "\n".join(
                    f"  {k}: {v}" for k, v in result.get("written", {}).items()
                ) or "  (none)"
                empty_str = ", ".join(result.get("empty", [])) or "none"
                messagebox.showinfo(
                    "Export to Records — Complete",
                    f"Appended to: {fname}"
                    + (f"\nSheet: {sheet_name}" if sheet_name else "")
                    + f"\n\nFields written ({n_written}):\n{written_str}"
                    + f"\n\nUnmatched headers: {empty_str}",
                )
        except Exception as exc:
            if not silent:
                messagebox.showerror("Export Error", str(exc))

    def _auto_export_on_complete(self, req_id: str) -> None:
        """Called after a request is marked completed; exports if not already done."""
        rec = self._data.get_record(req_id)
        if not rec or rec.get("schedule_exported", "0") == "1":
            return
        # Temporarily point _current_rec at this record so the export works
        saved = self._current_rec
        self._current_rec = rec
        self._on_export_to_schedule(silent=True)
        self._current_rec = saved

    def _show_cell_entry(self, row_id: str, col_id: str, field: str) -> None:
        """Overlay a plain Entry widget on a Treeview cell for free-text editing."""
        bbox = self._tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        var   = tk.StringVar(value=self._tree.set(row_id, col_id))
        entry = ttk.Entry(self._tree, textvariable=var, font=("", 9))
        entry.place(x=x, y=y, width=max(w, 90), height=h + 2)
        entry.focus_set()
        entry.select_range(0, "end")

        def _commit(*_):
            val = var.get().strip()
            entry.place_forget()
            entry.destroy()
            rec = self._data.get_record(row_id)
            if rec is None:
                return
            self._data.update_local_fields(row_id, **{field: val})
            rec[field] = val
            self._tree.set(row_id, col_id, val)
            if field == "training_time":
                self._training_time_var.set(val)

        def _cancel(*_):
            entry.place_forget()
            entry.destroy()

        entry.bind("<Return>",   _commit)
        entry.bind("<Tab>",      _commit)
        entry.bind("<Escape>",   _cancel)
        entry.bind("<FocusOut>", _commit)

    def _on_push_state_inline(self, row_id: str, new_state: str) -> None:
        """Push a state change to iLab from a table-cell dropdown."""
        self._data.update_field(row_id, "state", new_state)
        rec = self._data.get_record(row_id)
        if rec:
            rec["state"] = new_state
        self._refresh_table()
        try:
            self._tree.selection_set(row_id)
        except tk.TclError:
            pass

        if rec and rec.get("local_only") == "1":
            self._set_status(f"Status → {new_state} (local only, not pushed to iLab).")
            if new_state == "completed":
                self.root.after(100, lambda: self._auto_export_on_complete(row_id))
            return

        core_id = self._get_core_id()
        if not core_id:
            return

        def worker():
            try:
                client = self._get_client()
                client.update_service_request(core_id, int(row_id), state=new_state)
                self.root.after(0, lambda: self._set_status(
                    f"Status → {new_state} pushed to iLab for #{row_id}."))
                if new_state == "completed":
                    self.root.after(100, lambda: self._auto_export_on_complete(row_id))
            except Exception as exc:
                print(f"[iLab Error] inline push state={new_state} req={row_id}: {exc}")
                self.root.after(0, lambda e=exc: messagebox.showerror(
                    "iLab Error", f"Status updated locally but iLab push failed:\n{e}"))

        threading.Thread(target=worker, daemon=True).start()

    def _on_user_permissions(self) -> None:
        """Open a dialog linking to the iLab people/permissions page."""
        url = "https://ucsf.ilab.agilent.com/sc/5226/center-for-advanced-light-microscopy?tab=people"
        webbrowser.open(url)

        dlg = tk.Toplevel(self.root)
        dlg.title("User Permissions")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        ttk.Label(dlg, text="iLab user permissions page opened in your browser.",
                  padding=(16, 12)).pack()
        ttk.Label(dlg, text=url, foreground="#1565C0",
                  cursor="hand2", padding=(16, 0)).pack()
        ttk.Button(dlg, text="Done", command=dlg.destroy,
                   width=10).pack(pady=12)

        dlg.update_idletasks()
        px = self.root.winfo_rootx() + (self.root.winfo_width()  - dlg.winfo_width())  // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{max(0, px)}+{max(0, py)}")

    def _on_open_preferences(self) -> None:
        PreferencesDialog(self.root, dark=self._dark_mode)

    def _on_open_in_ilab(self, _event=None) -> None:
        if not self._current_rec:
            return
        req_id = self._current_rec.get("request_id", "")
        url = f"{ILAB_BASE_URL}/service_requests/{req_id}"
        webbrowser.open(url)

    # =========================================================================
    # Helpers
    # =========================================================================

    def _toggle_dark_mode(self) -> None:
        self._apply_theme(dark=not self._dark_mode)

    def _apply_theme(self, dark: bool) -> None:
        self._dark_mode = dark
        self._dark_btn.config(text="☀  Light" if dark else "🌙 Dark")

        if dark:
            bg       = "#00003F"
            bg2      = "#171753"   # treeview / canvas interior
            bg3      = "#382B58"   # fields, entries
            fg       = "#1cc2d4"
            fg2      = "#9aabad"
            sel_bg   = "#145C8B"
            sel_fg   = "#f8f3e5"
            border   = "#4C4683"
            state_colors = STATE_COLORS_DARK
        else:
            bg       = "#f0f0f0"
            bg2      = "#ffffff"
            bg3      = "#ffffff"
            fg       = "#000000"
            fg2      = "#555555"
            sel_bg   = "#0078d4"
            sel_fg   = "#ffffff"
            border   = "#cccccc"
            state_colors = STATE_COLORS

        style = ttk.Style(self.root)
        style.theme_use("clam")   # clam is fully configurable on all platforms

        style.configure(".",
            background=bg, foreground=fg,
            fieldbackground=bg3, selectbackground=sel_bg,
            selectforeground=sel_fg, bordercolor=border,
            troughcolor=bg2, arrowcolor=fg,
            insertcolor=fg, relief="flat",
        )
        for widget_class in ("TFrame", "TLabelframe", "TPanedwindow"):
            style.configure(widget_class, background=bg)
        style.configure("TLabelframe.Label", background=bg, foreground=fg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TButton", background=bg3, foreground=fg,
                         bordercolor=border, focuscolor=bg3)
        style.map("TButton",
            background=[("active", sel_bg), ("pressed", sel_bg)],
            foreground=[("active", sel_fg), ("pressed", sel_fg)],
        )
        style.configure("TEntry", fieldbackground=bg3, foreground=fg,
                         bordercolor=border, insertcolor=fg)
        style.configure("TCombobox", fieldbackground=bg3, foreground=fg,
                         selectbackground=sel_bg, selectforeground=sel_fg,
                         bordercolor=border, arrowcolor=fg)
        style.map("TCombobox",
            fieldbackground=[("readonly", bg3)],
            foreground=[("readonly", fg)],
        )
        style.configure("TCheckbutton", background=bg, foreground=fg,
                         focuscolor=bg)
        style.map("TCheckbutton",
            background=[("active", bg)],
            foreground=[("active", fg)],
        )
        style.configure("TNotebook", background=bg, bordercolor=border)
        style.configure("TNotebook.Tab", background=bg3, foreground=fg,
                         padding=(8, 4))
        style.map("TNotebook.Tab",
            background=[("selected", bg2), ("active", sel_bg)],
            foreground=[("selected", fg), ("active", sel_fg)],
        )
        style.configure("TScrollbar", background=bg3, troughcolor=bg2,
                         bordercolor=border, arrowcolor=fg, relief="flat")
        style.configure("TSeparator", background=border)

        # Treeview
        style.configure("Treeview",
            background=bg2, foreground=fg, fieldbackground=bg2,
            bordercolor=border, rowheight=22,
        )
        style.configure("Treeview.Heading",
            background=bg3, foreground=fg, bordercolor=border, relief="flat")
        style.map("Treeview",
            background=[("selected", sel_bg)],
            foreground=[("selected", sel_fg)],
        )
        for state, color in state_colors.items():
            self._tree.tag_configure(state, background=color,
                                     foreground=sel_fg if dark else fg)
            self._tree.tag_configure(state + "_alt",
                                     background=_lighten_hex(color),
                                     foreground=sel_fg if dark else fg)

        # Root window and bare tk widgets
        self.root.configure(bg=bg)

        if hasattr(self, "_qa_canvas"):
            self._qa_canvas.configure(bg=bg, highlightthickness=0)
        if hasattr(self, "_form_canvas"):
            self._form_canvas.configure(bg=bg, highlightthickness=0)
        if hasattr(self, "_notes_text"):
            self._notes_text.configure(
                bg=bg3, fg=fg, insertbackground=fg,
                selectbackground=sel_bg, selectforeground=sel_fg,
            )

        # Persist preference
        p = _prefs.get_prefs()
        p["dark_mode"] = "1" if dark else "0"
        _prefs.save_prefs(p)

    def _get_core_id(self) -> str | None:
        """Return the core ID/slug as entered. May be a number ('1234') or a slug ('CALM')."""
        val = self._core_id_var.get().strip()
        if not val:
            messagebox.showwarning("Core ID Required",
                                   "Enter your Core ID in the toolbar.\n"
                                   "Run  python get_cores.py  to find it.")
            return None
        return val

    def _get_client(self) -> ILabClient:
        if self._client is None:
            self._client = ILabClient(base_url=ILAB_BASE_URL)
        return self._client

    def _on_clear_all(self) -> None:
        if not messagebox.askyesno(
            "Clear All Requests",
            "Delete all cached service requests?\n\n"
            "This only clears the local cache — nothing in iLab is changed.\n"
            "Use  ↻ Sync from iLab  to fetch fresh data afterwards.",
            icon="warning",
        ):
            return
        self._data.clear_all()
        self._current_rec = None
        self._refresh_table()
        self._set_sync_indicator("idle")
        self._last_sync_var.set("")
        self._set_status("Cache cleared — click ↻ Sync from iLab to fetch fresh data.")

    def _set_sync_indicator(self, state: str) -> None:
        _COLORS = {
            "idle":    "#C8C8C8",   # grey
            "working": "#FFD600",   # yellow
            "ok":      "#43A047",   # green
            "error":   "#E53935",   # red
        }
        self._sync_indicator.config(bg=_COLORS.get(state, "#C8C8C8"))

    def _set_status(self, msg: str) -> None:
        self._status_var.set(msg)

    def _set_last_sync(self, when: str) -> None:
        self._last_sync_var.set(f"Last synced: {when}")

    def _restore_last_sync(self) -> None:
        """On startup, show the most recent last_synced timestamp from the cache."""
        records = self._data.all_records()
        if not records:
            return
        latest = max((r.get("last_synced", "") for r in records), default="")
        if not latest:
            return
        try:
            when = datetime.fromisoformat(latest[:19].replace("T", " ")).strftime("%b %d  %I:%M %p")
        except ValueError:
            when = latest[:10]
        self._set_last_sync(when)


# =============================================================================
# Import diagnostic dialog
# =============================================================================

class ImportDiagnosticDialog(tk.Toplevel):
    """
    Shown when CSV import reads 0 rows.  Displays the column headers found in
    the file so the user can report them and we can add a mapping.
    """

    def __init__(self, parent, filename: str, result: dict):
        super().__init__(parent)
        self.title("Import — Nothing Imported")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.minsize(480, 300)

        pad = dict(padx=12, pady=4)

        ttk.Label(
            self,
            text=f"0 rows were imported from  {filename}",
            font=("", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(12, 2))

        enc     = result.get("encoding", "?")
        skipped = result.get("skipped", 0)
        mapped  = result.get("columns_mapped", {})
        form    = result.get("columns_form", [])
        raw     = result.get("columns_raw", [])

        info = (
            f"Encoding detected : {enc}\n"
            f"Rows with no ID   : {skipped}  "
            f"(a column must map to 'Request ID' for a row to import)\n"
            f"Columns found     : {len(raw)}\n"
            f"  ✓ mapped to fields : {len(mapped)}\n"
            f"  → stored as form data : {len(form)}"
        )
        ttk.Label(self, text=info, justify="left").pack(
            anchor="w", **pad)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12, pady=6)

        if not raw:
            ttk.Label(
                self,
                text="No column headers were found — the file may be empty or\n"
                     "use a delimiter that wasn't recognised.",
                foreground="#C62828",
            ).pack(anchor="w", **pad)
        else:
            ttk.Label(
                self,
                text="Column headers found in the file  "
                     "(copy and send these to your admin):",
                foreground="#555",
            ).pack(anchor="w", **pad)

            # Scrollable list of headers with their mapping status
            frame = ttk.Frame(self)
            frame.pack(fill="both", expand=True, padx=12, pady=(0, 4))

            text = tk.Text(frame, height=12, wrap="none", font=("Courier", 9))
            vsb  = ttk.Scrollbar(frame, orient="vertical",   command=text.yview)
            hsb  = ttk.Scrollbar(frame, orient="horizontal", command=text.xview)
            text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            hsb.pack(side="bottom", fill="x")
            vsb.pack(side="right",  fill="y")
            text.pack(fill="both", expand=True)

            text.tag_configure("ok",   foreground="#2E7D32")
            text.tag_configure("form", foreground="#E65100")
            text.tag_configure("none", foreground="#9E9E9E")

            for h in raw:
                if h in mapped:
                    line = f"  ✓  {h!r:40s}  →  {mapped[h]}\n"
                    tag  = "ok"
                elif h in form:
                    line = f"  →  {h!r:40s}  (stored as form data)\n"
                    tag  = "form"
                else:
                    line = f"  –  {h!r}\n"
                    tag  = "none"
                text.insert("end", line, tag)

            text.configure(state="disabled")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12, pady=4)

        if skipped and not mapped:
            advice = (
                "The importer needs a column whose name matches a known 'Request ID'\n"
                "variant (e.g. 'ID', 'Request #', 'Request No.', 'Req ID').\n\n"
                "If none of the columns above look like request IDs, re-export the\n"
                "CSV from iLab and ensure the ID column is included."
            )
        elif skipped:
            advice = (
                f"{skipped} row(s) were found but each was missing a value in the\n"
                "Request ID column.  Check whether the ID column in the CSV is blank."
            )
        else:
            advice = (
                "The file appears to have no data rows, or they are all blank.\n"
                "Try re-exporting from iLab."
            )
        ttk.Label(self, text=advice, justify="left",
                  foreground="#555").pack(anchor="w", padx=12, pady=(0, 4))

        ttk.Button(self, text="Close", command=self.destroy, width=10).pack(
            anchor="e", padx=12, pady=(0, 12))

        # Centre
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, px)}+{max(0, py)}")


# =============================================================================
# Preferences dialog
# =============================================================================

class PreferencesDialog(tk.Toplevel):
    """Modal dialog: xlsx paths and class charge configuration."""

    def __init__(self, parent: tk.Tk, dark: bool = False):
        super().__init__(parent)
        self.title("Preferences")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # ── Theme colours ─────────────────────────────────────────────────────
        if dark:
            self._bg   = "#1e1e1e"
            self._bg3  = "#3c3c3c"
            self._fg   = "#d4d4d4"
            self._fg2  = "#9e9e9e"
        else:
            self._bg   = "#f0f0f0"
            self._bg3  = "#ffffff"
            self._fg   = "#000000"
            self._fg2  = "#555555"

        self.configure(bg=self._bg)

        p = _prefs.get_prefs()
        self._vars: dict[str, tk.StringVar] = {
            k: tk.StringVar(value=str(v)) for k, v in p.items()
        }
        # Make sure all expected keys exist
        for k, default in [
            ("data_file",""),
            ("calm_xlsx",""), ("cvri_xlsx",""),
            ("class_service_id",""), ("class_price_id",""),
            ("class_quantity","2"), ("class_unit_price","100"),
        ]:
            if k not in self._vars:
                self._vars[k] = tk.StringVar(value=default)

        self._original_data_file = self._vars["data_file"].get()

        PAD = dict(padx=8, pady=4)

        # ── Section: shared data file ─────────────────────────────────────────
        ttk.Label(self, text="Shared Data File",
                  font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(12, 4))

        ttk.Label(self, text="Data File (CSV):", width=14, anchor="e").grid(
            row=1, column=0, sticky="e", **PAD)
        ttk.Entry(self, textvariable=self._vars["data_file"], width=40).grid(
            row=1, column=1, sticky="ew", **PAD)
        ttk.Button(self, text="Browse…",
                   command=self._browse_data_file).grid(row=1, column=2, **PAD)

        ttk.Label(
            self,
            text="Point to a OneDrive/shared folder to sync across machines.\n"
                 "Leave blank to use the default file in the app directory.\n"
                 "Restart the app after changing this path.",
            foreground=self._fg2,
        ).grid(row=2, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))

        ttk.Separator(self, orient="horizontal").grid(
            row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=6)

        # ── Section: xlsx files ───────────────────────────────────────────────
        ttk.Label(self, text="Training Schedule Files",
                  font=("", 10, "bold")).grid(
            row=4, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))

        for base_row, (core, path_key, sheet_key) in enumerate([
            ("CALM", "calm_xlsx", "calm_sheet"),
            ("CVRI", "cvri_xlsx", "cvri_sheet"),
        ]):
            r = 5 + base_row * 2   # two grid rows per core (offset after data-file section)

            # xlsx path row
            ttk.Label(self, text=f"{core} xlsx:", width=14, anchor="e").grid(
                row=r, column=0, sticky="e", **PAD)
            ttk.Entry(self, textvariable=self._vars[path_key], width=40).grid(
                row=r, column=1, sticky="w", **PAD)
            ttk.Button(self, text="Browse…",
                       command=lambda k=path_key: self._browse(k)).grid(
                row=r, column=2, **PAD)

            # Sheet name row
            ttk.Label(self, text="Sheet name:", width=14, anchor="e").grid(
                row=r + 1, column=0, sticky="e", padx=8, pady=(0, 6))
            ttk.Entry(self, textvariable=self._vars[sheet_key], width=22).grid(
                row=r + 1, column=1, sticky="w", padx=8, pady=(0, 6))
            ttk.Label(self, text="(blank = first sheet)",
                      foreground=self._fg2).grid(
                row=r + 1, column=2, sticky="w", padx=4, pady=(0, 6))

        ttk.Label(
            self,
            text="Point to your existing Training Schedule xlsx files.\n"
                 "If the file does not exist it will be created with default headers.",
            foreground=self._fg2,
        ).grid(row=9, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))

        ttk.Separator(self, orient="horizontal").grid(
            row=10, column=0, columnspan=4, sticky="ew", padx=10, pady=6)

        # ── Section: class charge ─────────────────────────────────────────────
        ttk.Label(self, text="Class Charge (iLab API)",
                  font=("", 10, "bold")).grid(
            row=11, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))

        for i, (label, key) in enumerate([
            ("Service ID:",     "class_service_id"),
            ("Price ID:",       "class_price_id"),
            ("Quantity:",       "class_quantity"),
            ("Unit Price ($):", "class_unit_price"),
            ("Max Charge ($):", "max_charge"),
        ], start=12):
            ttk.Label(self, text=label, width=14, anchor="e").grid(
                row=i, column=0, sticky="e", **PAD)
            ttk.Entry(self, textvariable=self._vars[key], width=14).grid(
                row=i, column=1, sticky="w", **PAD)

        ttk.Label(
            self,
            text="Run  python get_services.py  to look up Service ID and Price ID.\n"
                 "Leave blank to save Class Taken locally without calling iLab.",
            foreground=self._fg2,
        ).grid(row=17, column=0, columnspan=4, sticky="w", padx=12, pady=(2, 8))

        ttk.Separator(self, orient="horizontal").grid(
            row=18, column=0, columnspan=4, sticky="ew", padx=10, pady=4)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = ttk.Frame(self, padding=(8, 4, 12, 10))
        btn_row.grid(row=19, column=0, columnspan=4, sticky="e")
        ttk.Button(btn_row, text="Save", command=self._save,
                   width=10).pack(side="right", padx=(6, 0))
        ttk.Button(btn_row, text="Cancel", command=self.destroy,
                   width=10).pack(side="right")

        # Centre over parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, px)}+{max(0, py)}")

    def _browse_data_file(self) -> None:
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Locate existing ilab_requests_cache.csv",
        )
        if path:
            self._vars["data_file"].set(path)

    def _browse(self, key: str) -> None:
        label = "CALM" if "calm" in key else "CVRI"
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            title=f"Select {label} Training Schedule xlsx",
        )
        if path:
            self._vars[key].set(path)

    def _save(self) -> None:
        new_data_file = self._vars["data_file"].get().strip()
        _prefs.save_prefs({k: v.get() for k, v in self._vars.items()})
        self.destroy()
        if new_data_file != self._original_data_file:
            messagebox.showinfo(
                "Restart Required",
                "The Data File path has changed.\n\n"
                "Please restart the app to load data from the new location.",
            )


# =============================================================================
# Entry point
# =============================================================================

def main() -> None:
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.25)   # nicer on HiDPI screens
    except tk.TclError:
        pass
    ILabManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
