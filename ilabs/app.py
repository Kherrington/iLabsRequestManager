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

        self._data = DataStore(DATA_FILE)
        self._client: ILabClient | None = None
        self._current_rec: dict | None = None
        self._sort_col = "created_at"
        self._sort_rev = True
        self._core_id: int | None = CORE_ID

        self._build_ui()
        self._refresh_table()

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
        ttk.Button(bar, text="⚙  Preferences",          command=self._on_open_preferences).pack(side="right", padx=(0, 4))

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
            "core_lab", "microscope", "training_date", "class_taken",
        ]
        headers = {
            "request_id":   ("ID",           75),
            "created_at":   ("Submitted",    110),
            "owner_name":   ("Requester",    155),
            "pi_name":      ("PI",           150),
            "service_name": ("Service",      200),
            "state":        ("State",        115),
            "assigned_to":  ("Assigned To",  130),
            "labels":       ("Labels",       140),
            "core_lab":     ("Core",          55),
            "microscope":   ("Microscope",   120),
            "training_date":("Trng Date",     88),
            "class_taken":  ("Class",         48),
        }

        self._tree = ttk.Treeview(frame, columns=cols, show="headings",
                                  selectmode="browse")
        for col, (header, width) in headers.items():
            self._tree.heading(col, text=header,
                               command=lambda c=col: self._sort_by(c))
            self._tree.column(col, width=width, anchor="w", minwidth=50)

        for state, color in STATE_COLORS.items():
            self._tree.tag_configure(state, background=color)

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right",  fill="y")
        self._tree.pack(fill="both", expand=True)

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self._tree.bind("<Double-1>",         self._on_open_in_ilab)

    # ── Detail panel ─────────────────────────────────────────────────────────

    def _build_detail_panel(self, parent: ttk.PanedWindow) -> None:
        outer = ttk.LabelFrame(parent, text="Request Details", padding=4)
        parent.add(outer, weight=2)

        self._notebook = ttk.Notebook(outer)
        self._notebook.pack(fill="both", expand=True)

        self._build_info_tab()
        self._build_form_tab()
        self._build_milestones_tab()
        self._build_training_tab()

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
            ("state",       "State"),
            ("created_at",  "Submitted"),
            ("start_on",    "Start Date"),
            ("end_on",      "End Date"),
            ("completed_on","Completed"),
            ("owner_name",  "Requester"),
            ("owner_email", "Email"),
            ("pi_name",     "PI"),
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
        tab = ttk.Frame(self._notebook, padding=8)
        self._notebook.add(tab, text="  Track Work  ")

        # Outer scrollable container
        canvas = tk.Canvas(tab, highlightthickness=0)
        vsb = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)

        self._ms_inner = ttk.Frame(canvas)
        canvas_win = canvas.create_window((0, 0), window=self._ms_inner, anchor="nw")

        self._ms_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
            lambda e: canvas.itemconfig(canvas_win, width=e.width))

        self._ms_canvas = canvas
        self._ms_placeholder = ttk.Label(
            self._ms_inner,
            text="Select a request to view milestones.",
            foreground="#888")
        self._ms_placeholder.grid(row=0, column=0, padx=12, pady=12)

    # ── Training tab ─────────────────────────────────────────────────────────

    def _build_training_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=10)
        self._notebook.add(tab, text="  Training  ")

        # Variables
        self._training_core_var  = tk.StringVar()
        self._training_micro_var = tk.StringVar()
        self._training_date_var  = tk.StringVar()
        self._training_day_var   = tk.StringVar()
        self._training_time_var  = tk.StringVar()
        self._class_taken_var    = tk.BooleanVar()

        left  = ttk.Frame(tab)
        right = ttk.Frame(tab)
        left.pack(side="left", fill="both", expand=True, padx=(0, 18))
        right.pack(side="left", fill="both", expand=True)

        # ── Left: fields ──────────────────────────────────────────────────────
        ttk.Label(left, text="Training Details", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        LW = 15   # label column width

        # Core Lab
        ttk.Label(left, text="Core Lab:", anchor="e", width=LW).grid(
            row=1, column=0, sticky="e", padx=4, pady=3)
        core_row = ttk.Frame(left)
        core_row.grid(row=1, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        ttk.Combobox(core_row, textvariable=self._training_core_var,
                     values=CORE_OPTIONS, width=8, state="readonly").pack(side="left")
        ttk.Button(core_row, text="Auto-detect",
                   command=self._on_detect_core).pack(side="left", padx=(6, 0))

        # Microscope
        ttk.Label(left, text="Microscope:", anchor="e", width=LW).grid(
            row=2, column=0, sticky="e", padx=4, pady=3)
        ttk.Combobox(left, textvariable=self._training_micro_var,
                     values=MICROSCOPES, width=20).grid(
            row=2, column=1, columnspan=2, sticky="w", padx=4, pady=3)

        # Training Date
        ttk.Label(left, text="Training Date:", anchor="e", width=LW).grid(
            row=3, column=0, sticky="e", padx=4, pady=3)
        date_row = ttk.Frame(left)
        date_row.grid(row=3, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        ttk.Entry(date_row, textvariable=self._training_date_var, width=13).pack(side="left")
        ttk.Button(date_row, text="📅", width=3,
                   command=self._on_training_date_pick).pack(side="left", padx=(4, 0))

        # Training Day
        ttk.Label(left, text="Training Day:", anchor="e", width=LW).grid(
            row=4, column=0, sticky="e", padx=4, pady=3)
        ttk.Combobox(left, textvariable=self._training_day_var,
                     values=TRAINING_DAYS, width=10, state="readonly").grid(
            row=4, column=1, sticky="w", padx=4, pady=3)

        # Training Time
        ttk.Label(left, text="Training Time:", anchor="e", width=LW).grid(
            row=5, column=0, sticky="e", padx=4, pady=3)
        ttk.Entry(left, textvariable=self._training_time_var, width=14).grid(
            row=5, column=1, sticky="w", padx=4, pady=3)

        ttk.Separator(left, orient="horizontal").grid(
            row=6, column=0, columnspan=3, sticky="ew", pady=8)

        # Class Taken
        ttk.Checkbutton(
            left,
            text="Class Taken  (adds 2 × Class charge, $200 total, to iLab)",
            variable=self._class_taken_var,
            command=self._on_class_taken_toggle,
        ).grid(row=7, column=0, columnspan=3, sticky="w", padx=4)
        self._class_status_lbl = ttk.Label(
            left, text="", foreground="#555", wraplength=340)
        self._class_status_lbl.grid(
            row=8, column=0, columnspan=3, sticky="w", padx=24, pady=(0, 4))

        ttk.Separator(left, orient="horizontal").grid(
            row=9, column=0, columnspan=3, sticky="ew", pady=6)

        # Buttons
        btn_row = ttk.Frame(left)
        btn_row.grid(row=10, column=0, columnspan=3, sticky="w", padx=4)
        ttk.Button(btn_row, text="Save Training Info",
                   command=self._on_save_training).pack(side="left", padx=(0, 8))
        ttk.Button(btn_row, text="Export to Schedule →",
                   command=self._on_export_to_schedule).pack(side="left")

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
            "5. Save Training Info, then Export to Schedule →\n"
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
        for rec in records:
            state = rec.get("state", "")
            date  = (rec.get("created_at") or "")[:10]
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
                    "✓" if rec.get("class_taken") == "1" else "",
                ),
                tags=(state,),
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
                          wraplength=260, foreground="#333").grid(
                    row=i, column=0, padx=6, pady=2, sticky="nw")
                ttk.Label(self._form_inner, text=str(fval), anchor="nw",
                          wraplength=400).grid(
                    row=i, column=1, padx=6, pady=2, sticky="nw")
        else:
            ttk.Label(self._form_inner,
                      text="No form data loaded.\nSync from iLab to fetch form fields.",
                      foreground="#888").grid(row=0, column=0, padx=12, pady=12)

        self._form_canvas.configure(scrollregion=self._form_canvas.bbox("all"))

        # ── Milestones tab ────────────────────────────────────────────────────
        milestones: list = json.loads(rec.get("milestones_data") or "[]")
        self._load_milestones(milestones, rec.get("request_id", ""))

        # ── Training tab ──────────────────────────────────────────────────────
        self._training_core_var.set(rec.get("core_lab", ""))
        self._training_micro_var.set(rec.get("microscope", ""))
        self._training_date_var.set(rec.get("training_date", ""))
        self._training_day_var.set(rec.get("training_day", ""))
        self._training_time_var.set(rec.get("training_time", ""))
        taken = rec.get("class_taken", "0") == "1"
        self._class_taken_var.set(taken)
        self._class_status_lbl.config(
            text="✓ Class charge previously submitted." if taken else "")

    def _load_milestones(self, milestones: list, request_id: str) -> None:
        for w in self._ms_inner.winfo_children():
            w.destroy()

        if not milestones:
            ttk.Label(self._ms_inner,
                      text="No milestones found for this request.",
                      foreground="#888").grid(row=0, column=0, padx=12, pady=12)
            return

        for col, (text, w) in enumerate([("Milestone", 260), ("Started", 100),
                                          ("Completed", 100), ("Actions", 200)]):
            ttk.Label(self._ms_inner, text=text, font=("", 9, "bold"),
                      width=w//8).grid(row=0, column=col, padx=6, pady=(4,2), sticky="w")
        ttk.Separator(self._ms_inner, orient="horizontal").grid(
            row=1, column=0, columnspan=4, sticky="ew", padx=6)

        for i, ms in enumerate(milestones, start=2):
            ms_id    = ms.get("id")
            name     = ms.get("name", "—")
            started  = (ms.get("started_on")  or "")[:10] or "—"
            finished = (ms.get("completed_on") or "")[:10] or "—"

            ttk.Label(self._ms_inner, text=name, wraplength=250).grid(
                row=i, column=0, padx=6, pady=3, sticky="w")
            ttk.Label(self._ms_inner, text=started).grid(
                row=i, column=1, padx=6, pady=3)
            ttk.Label(self._ms_inner, text=finished).grid(
                row=i, column=2, padx=6, pady=3)

            btn_frame = ttk.Frame(self._ms_inner)
            btn_frame.grid(row=i, column=3, padx=6, pady=2)
            if ms.get("started_on") is None:
                ttk.Button(
                    btn_frame, text="Mark Started",
                    command=lambda r=request_id, m=ms_id, ml=milestones:
                        self._on_milestone_action(r, m, ml, "started"),
                ).pack(side="left", padx=2)
            if ms.get("completed_on") is None:
                ttk.Button(
                    btn_frame, text="Mark Complete",
                    command=lambda r=request_id, m=ms_id, ml=milestones:
                        self._on_milestone_action(r, m, ml, "completed"),
                ).pack(side="left", padx=2)
            elif finished != "—":
                ttk.Label(btn_frame, text="✓ Done", foreground="#2E7D32").pack()

        self._ms_canvas.configure(scrollregion=self._ms_canvas.bbox("all"))

    # =========================================================================
    # Event handlers
    # =========================================================================

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
                    f"State updated to '{new_state}' for request {req_id}."))
            except ILabError as exc:
                self.root.after(0, lambda e=exc: messagebox.showerror("iLab API Error", str(e)))
            except Exception as exc:
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
                self.root.after(0, lambda: self._load_milestones(milestones, request_id))
                self.root.after(0, lambda: self._set_status(
                    f"Milestone {milestone_id} marked {action}."))
            except ILabError as exc:
                self.root.after(0, lambda e=exc: messagebox.showerror("iLab API Error", str(e)))
            except Exception as exc:
                self.root.after(0, lambda e=exc: messagebox.showerror("Error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

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
        # Auto-populate Training Day from the selected date
        try:
            d = _date.fromisoformat(date_str)
            # Map Python weekday (0=Mon…6=Sun) to TRAINING_DAYS list index
            day_abbrs = TRAINING_DAYS
            weekday   = d.weekday()          # 0=Mon, 6=Sun
            if weekday < len(day_abbrs):
                self._training_day_var.set(day_abbrs[weekday])
        except ValueError:
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
                self.root.after(0, lambda: self._class_status_lbl.config(
                    text=f"✓ Class charge submitted to iLab ({qty} units)."))
                self.root.after(0, self._refresh_table)
                self.root.after(0, lambda: self._set_status(
                    f"Class charge added to request {req_id} in iLab."))
            except Exception as exc:
                self.root.after(0, lambda: self._class_taken_var.set(False))
                self.root.after(0, lambda: self._class_status_lbl.config(text=""))
                self.root.after(0, lambda e=exc: messagebox.showerror("Charge Error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_export_to_schedule(self) -> None:
        if not self._current_rec:
            messagebox.showwarning("No Selection", "Select a request first.")
            return

        # Save any unsaved training fields first
        self._on_save_training()

        core = self._training_core_var.get().strip().upper()
        if not core:
            messagebox.showwarning(
                "Core Lab Not Set",
                "Set the Core Lab (CALM or CVRI) before exporting.\n\n"
                "Use the Auto-detect button or choose manually.",
            )
            return

        p = _prefs.get_prefs()
        xlsx_path = str(p.get(f"{core.lower()}_xlsx", "") or "").strip()
        if not xlsx_path:
            messagebox.showwarning(
                "xlsx Path Not Configured",
                f"No {core} schedule xlsx file path is set.\n\n"
                "Open ⚙ Preferences and browse to your\n"
                f"Training Schedule_{core}_CURRENT.xlsx file.",
            )
            return

        if not HAS_OPENPYXL:
            messagebox.showerror(
                "openpyxl Required",
                "Install openpyxl to export to xlsx:\n\n    pip install openpyxl",
            )
            return

        try:
            append_training_row(self._current_rec, xlsx_path)
            fname = Path(xlsx_path).name
            self._set_status(f"Row appended to {fname} ({core}).")
            messagebox.showinfo(
                "Export Complete",
                f"Training record appended to:\n{xlsx_path}",
            )
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    def _on_open_preferences(self) -> None:
        PreferencesDialog(self.root)

    def _on_open_in_ilab(self, _event=None) -> None:
        if not self._current_rec:
            return
        req_id = self._current_rec.get("request_id", "")
        url = f"{ILAB_BASE_URL}/service_requests/{req_id}"
        webbrowser.open(url)

    # =========================================================================
    # Helpers
    # =========================================================================

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

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("Preferences")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        p = _prefs.get_prefs()
        self._vars: dict[str, tk.StringVar] = {
            k: tk.StringVar(value=str(v)) for k, v in p.items()
        }
        # Make sure all expected keys exist
        for k, default in [
            ("calm_xlsx",""), ("cvri_xlsx",""),
            ("class_service_id",""), ("class_price_id",""),
            ("class_quantity","2"), ("class_unit_price","100"),
        ]:
            if k not in self._vars:
                self._vars[k] = tk.StringVar(value=default)

        PAD = dict(padx=8, pady=4)

        # ── Section: xlsx files ───────────────────────────────────────────────
        ttk.Label(self, text="Training Schedule Files",
                  font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 4))

        for row, (label, key) in enumerate([
            ("CALM xlsx:", "calm_xlsx"),
            ("CVRI xlsx:", "cvri_xlsx"),
        ], start=1):
            ttk.Label(self, text=label, width=16, anchor="e").grid(
                row=row, column=0, sticky="e", **PAD)
            ttk.Entry(self, textvariable=self._vars[key], width=44).grid(
                row=row, column=1, sticky="w", **PAD)
            ttk.Button(self, text="Browse…",
                       command=lambda k=key: self._browse(k)).grid(
                row=row, column=2, **PAD)

        ttk.Label(
            self,
            text="Point to your existing Training Schedule xlsx files.\n"
                 "If the file does not exist it will be created with default headers.",
            foreground="#555",
        ).grid(row=3, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 4))

        ttk.Separator(self, orient="horizontal").grid(
            row=4, column=0, columnspan=3, sticky="ew", padx=10, pady=6)

        # ── Section: class charge ─────────────────────────────────────────────
        ttk.Label(self, text="Class Charge (iLab API)",
                  font=("", 10, "bold")).grid(
            row=5, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 4))

        for i, (label, key) in enumerate([
            ("Service ID:",     "class_service_id"),
            ("Price ID:",       "class_price_id"),
            ("Quantity:",       "class_quantity"),
            ("Unit Price ($):", "class_unit_price"),
            ("Max Charge ($):", "max_charge"),
        ], start=6):
            ttk.Label(self, text=label, width=16, anchor="e").grid(
                row=i, column=0, sticky="e", **PAD)
            ttk.Entry(self, textvariable=self._vars[key], width=14).grid(
                row=i, column=1, sticky="w", **PAD)

        ttk.Label(
            self,
            text="Run  python get_services.py  to look up Service ID and Price ID.\n"
                 "Leave blank to save Class Taken locally without calling iLab.",
            foreground="#555",
        ).grid(row=10, column=0, columnspan=3, sticky="w", padx=12, pady=(2, 8))

        ttk.Separator(self, orient="horizontal").grid(
            row=11, column=0, columnspan=3, sticky="ew", padx=10, pady=4)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = ttk.Frame(self, padding=(8, 4, 12, 10))
        btn_row.grid(row=12, column=0, columnspan=3, sticky="e")
        ttk.Button(btn_row, text="Save", command=self._save,
                   width=10).pack(side="right", padx=(6, 0))
        ttk.Button(btn_row, text="Cancel", command=self.destroy,
                   width=10).pack(side="right")

        # Centre over parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, px)}+{max(0, py)}")

    def _browse(self, key: str) -> None:
        label = "CALM" if "calm" in key else "CVRI"
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            title=f"Select {label} Training Schedule xlsx",
        )
        if path:
            self._vars[key].set(path)

    def _save(self) -> None:
        _prefs.save_prefs({k: v.get() for k, v in self._vars.items()})
        self.destroy()


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
