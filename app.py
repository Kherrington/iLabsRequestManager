"""
iLab Service Request Manager
Run:  python app.py

Requirements: pip install requests
Set ILAB_TOKEN env var (or add to .env and load before running).
Set CORE_ID in config.py after running: python get_cores.py
"""

import html
import json
import re
import threading
import webbrowser
from datetime import datetime, timezone, date as _date
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

try:
    import win32clipboard
    HAS_WIN32CLIPBOARD = True
except ImportError:
    HAS_WIN32CLIPBOARD = False

import prefs as _prefs
from calendar_widget import CalendarPicker
from xlsx_export import append_training_row, append_class_session, HAS_OPENPYXL
from config import (
    CORE_ID, ILAB_BASE_URL, DATA_FILE, TEAM_MEMBERS, LABELS,
    MICROSCOPES, TRAINING_DAYS, CORE_OPTIONS, ACTIVE_STATES,
    COMMON_RESPONSE_IMAGE,
)
from data_store import DataStore
from ilabs_client import ILabClient, ILabError

_CS_SESSION_FILE         = Path(__file__).parent / "class_session.json"
_CALM_WELCOME_FILE       = Path(__file__).parent / "CALM_welcome.txt"
_CVRI_WELCOME_FILE       = Path(__file__).parent / "CVRI_welcome.txt"
_COMMON_RESPONSE_FILE    = Path(__file__).parent / "Common_response.txt"
_COMMON_RESPONSE_IMG_FILE = Path(__file__).parent / "Common_response_image.txt"
_CVRI_ACCESS_FILE        = Path(__file__).parent / "CVRI-Access.txt"

# ── Email-template markup parser ──────────────────────────────────────────────
# Supports: **bold**   *italic*   __underline__
_MARKUP_RE = re.compile(
    r'\*\*(.+?)\*\*|__(.+?)__|(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)',
    re.DOTALL,
)


def _parse_template_markup(text: str) -> list[tuple[str, tuple]]:
    """Return [(chunk, tag_tuple), ...] from marked-up plain text."""
    result: list[tuple[str, tuple]] = []
    last = 0
    for m in _MARKUP_RE.finditer(text):
        if m.start() > last:
            result.append((text[last:m.start()], ()))
        if m.group(1) is not None:
            result.append((m.group(1), ("bold",)))
        elif m.group(2) is not None:
            result.append((m.group(2), ("underline",)))
        else:
            result.append((m.group(3), ("italic",)))
        last = m.end()
    if last < len(text):
        result.append((text[last:], ()))
    return result


def _strip_template_markup(text: str) -> str:
    """Return plain text with all markup markers removed."""
    return _MARKUP_RE.sub(lambda m: m.group(1) or m.group(2) or m.group(3), text)


# ── Rich-text clipboard (preserve bold/italic/underline on paste) ────────────
_HTML_TAG_MAP = {"bold": "b", "italic": "i", "underline": "u"}

# Bare URLs in template text (http(s):// or www.) get turned into real <a> links.
_URL_RE = re.compile(r'(https?://\S+|www\.\S+)')
_URL_TRAILING_PUNCT = '.,;:!?)]}\'\"'


def _split_urls(text: str) -> list[tuple[str, bool]]:
    """Split *text* into (segment, is_link) pieces, trimming trailing
    punctuation (periods, closing parens, etc.) off detected URLs so
    sentence punctuation doesn't get swallowed into the link."""
    parts: list[tuple[str, bool]] = []
    last = 0
    for m in _URL_RE.finditer(text):
        if m.start() > last:
            parts.append((text[last:m.start()], False))
        url, trail = m.group(0), ""
        while url and url[-1] in _URL_TRAILING_PUNCT:
            trail = url[-1] + trail
            url = url[:-1]
        if url:
            parts.append((url, True))
        if trail:
            parts.append((trail, False))
        last = m.end()
    if last < len(text):
        parts.append((text[last:], False))
    return parts


def _runs_to_html(runs: list[tuple[str, str | None]]) -> str:
    """Convert [(text, tag_or_None), ...] runs into an HTML fragment,
    wrapping runs in <b>/<i>/<u> per _HTML_TAG_MAP, preserving line breaks,
    and turning any bare URLs into clickable <a> links."""
    parts = []
    for chunk, tag in runs:
        seg_html = []
        for seg, is_link in _split_urls(chunk):
            esc = html.escape(seg).replace("\n", "<br>\n")
            if is_link:
                href = seg if seg.lower().startswith(("http://", "https://")) else f"https://{seg}"
                seg_html.append(f'<a href="{html.escape(href, quote=True)}">{esc}</a>')
            else:
                seg_html.append(esc)
        chunk_html = "".join(seg_html)
        wrapper = _HTML_TAG_MAP.get(tag)
        if wrapper:
            chunk_html = f"<{wrapper}>{chunk_html}</{wrapper}>"
        parts.append(chunk_html)
    return "".join(parts)


def _build_cf_html(html_fragment: str) -> bytes:
    """Wrap an HTML fragment in the header Windows' CF_HTML clipboard format
    requires (byte offsets to the fragment within the whole payload)."""
    header_tmpl = (
        "Version:0.9\r\n"
        "StartHTML:{:09d}\r\n"
        "EndHTML:{:09d}\r\n"
        "StartFragment:{:09d}\r\n"
        "EndFragment:{:09d}\r\n"
    )
    prefix = "<html><body>\r\n<!--StartFragment-->"
    suffix = "<!--EndFragment-->\r\n</body></html>"

    header_len     = len(header_tmpl.format(0, 0, 0, 0).encode("utf-8"))
    start_html     = header_len
    start_fragment = start_html + len(prefix.encode("utf-8"))
    end_fragment   = start_fragment + len(html_fragment.encode("utf-8"))
    end_html       = end_fragment + len(suffix.encode("utf-8"))

    header = header_tmpl.format(start_html, end_html, start_fragment, end_fragment)
    return (header + prefix + html_fragment + suffix).encode("utf-8")


def _copy_rich_text(plain_text: str, html_fragment: str, tk_root) -> bool:
    """
    Put both plain text and HTML on the clipboard so pasting into a rich-text
    target (Outlook, Word, Gmail compose, etc.) preserves bold/italic/underline.

    Returns True if the HTML format was written (pywin32 available), False if
    it fell back to plain-text only (still copied via Tk's clipboard).
    """
    if not HAS_WIN32CLIPBOARD:
        tk_root.clipboard_clear()
        tk_root.clipboard_append(plain_text)
        return False

    cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
    data    = _build_cf_html(html_fragment)
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text)
        win32clipboard.SetClipboardData(cf_html, data)
    finally:
        win32clipboard.CloseClipboard()
    return True


def _copy_image_to_clipboard(image_path: str | Path) -> bool:
    """Copy image file to clipboard. Returns True if successful, False otherwise."""
    from pathlib import Path
    image_path = Path(image_path)

    if not image_path.exists():
        return False

    if not HAS_WIN32CLIPBOARD:
        return False

    try:
        from PIL import Image
        import io

        # Load image and convert to RGB if needed (removes alpha channel)
        img = Image.open(image_path)
        if img.mode in ("RGBA", "LA", "P"):
            # Create white background
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            background.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
            img = background
        elif img.mode != "RGB":
            img = img.convert("RGB")

        # Convert to BMP format (DIB)
        bmp_buffer = io.BytesIO()
        img.save(bmp_buffer, format="BMP")
        bmp_data = bmp_buffer.getvalue()

        # BMP format: skip the 14-byte file header, use only the DIB data
        dib_data = bmp_data[14:]

        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dib_data)
        finally:
            win32clipboard.CloseClipboard()
        return True
    except ImportError:
        # Fallback if PIL not available
        return False
    except Exception as e:
        return False


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
    "core_disagreement":           "#FFCDD2",
    "disagreement":                "#FFCDD2",
    "researcher_in_agreement":     "#DCEDC8",
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
    "researcher_in_agreement":     "#1a2d0a",
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
        # Resolve relative paths to app directory to avoid duplicate caches
        from pathlib import Path as _Path
        _data_path_obj = _Path(_data_path)
        if not _data_path_obj.is_absolute():
            _data_path = str(_Path(__file__).parent / _data_path)
        self._data = DataStore(_data_path)
        self._client: ILabClient | None = None
        self._current_rec: dict | None = None
        self._sort_col = "created_at"
        self._sort_rev = True
        self._core_id: int | None = CORE_ID
        # Shared var used by both the Track Work and Training tabs
        self._class_taken_var = tk.BooleanVar()

        self._dark_mode = False
        self._autosave_job: str | None = None

        self._build_ui()
        self._data.reload()          # pick up any changes written by other machines
        self._refresh_table()
        self._restore_last_sync()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # ── Auto-save wiring ──────────────────────────────────────────────────
        # Reset the 60-second timer on any DataStore write
        self._data.on_change = self._schedule_autosave
        # Info tab fields (saved only on "Save Local Changes" click)
        for _var in (self._assigned_var,):
            _var.trace_add("write", self._schedule_autosave)
        for _var in self._label_vars.values():
            _var.trace_add("write", self._schedule_autosave)
        self._notes_text.bind("<KeyRelease>", self._schedule_autosave)
        # Training tab fields (saved only on "Save Training Info" click)
        for _var in (self._training_core_var, self._training_micro_var,
                     self._training_date_var, self._training_day_var,
                     self._training_time_var):
            _var.trace_add("write", self._schedule_autosave)
        # Class Schedule session-info fields
        for _var in (self._cs_date_var, self._cs_instructor_var,
                     self._cs_time_var, self._cs_location_var):
            _var.trace_add("write", self._schedule_autosave)

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
        self._build_class_schedule_tab()

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
            if key == "owner_email":
                ttk.Button(left, text="Copy", width=5,
                           command=lambda k=key: self._copy_info_field(k)).grid(
                    row=i, column=2, sticky="w", padx=(2, 4), pady=2)

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

    # ── Email-template viewer / editor ───────────────────────────────────────

    def _open_template_view(self, title: str, path: Path) -> None:
        """Show a formatted read-only view of an email template file."""
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.geometry("620x460")
        dlg.resizable(True, True)
        dlg.transient(self.root)

        # Scrollable text area
        frm = ttk.Frame(dlg)
        frm.pack(fill="both", expand=True, padx=8, pady=(8, 4))
        sb  = ttk.Scrollbar(frm, orient="vertical")
        txt = tk.Text(frm, wrap="word", padx=10, pady=10,
                      state="disabled", yscrollcommand=sb.set,
                      font=("", 10))
        sb.config(command=txt.yview)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        txt.tag_configure("bold",      font=("", 10, "bold"))
        txt.tag_configure("italic",    font=("", 10, "italic"))
        txt.tag_configure("underline", underline=True)

        def _reload():
            content = path.read_text(encoding="utf-8") if path.exists() else (
                f"[No template file yet — click Edit to create {path.name}]")
            txt.config(state="normal")
            txt.delete("1.0", "end")
            for chunk, tags in _parse_template_markup(content):
                txt.insert("end", chunk, tags)
            txt.config(state="disabled")

        _reload()

        def _runs_in_range(start: str, end: str) -> list[tuple[str, str | None]]:
            """Walk txt[start:end] character by character, merging consecutive
            runs that share the same bold/italic/underline tag."""
            runs: list[tuple[str, str | None]] = []
            idx = start
            while txt.compare(idx, "<", end):
                nxt = txt.index(f"{idx}+1c")
                ch  = txt.get(idx, nxt)
                tag = next((t for t in txt.tag_names(idx)
                           if t in ("bold", "italic", "underline")), None)
                if runs and runs[-1][1] == tag:
                    runs[-1] = (runs[-1][0] + ch, tag)
                else:
                    runs.append((ch, tag))
                idx = nxt
            return runs

        def _copy_formatted(event=None):
            """Copy the current selection (or the whole template if nothing is
            selected) to the clipboard, preserving bold/italic/underline for
            pasting into Outlook / Word / Gmail compose etc."""
            try:
                start, end = txt.index("sel.first"), txt.index("sel.last")
            except tk.TclError:
                start, end = "1.0", "end-1c"
            plain = txt.get(start, end)
            if not plain:
                return "break"
            html_fragment = _runs_to_html(_runs_in_range(start, end))
            rich = _copy_rich_text(plain, html_fragment, self.root)
            self._set_status(
                "Copied with formatting." if rich else
                "Copied as plain text (install pywin32 to preserve formatting).")
            return "break"

        txt.bind("<<Copy>>",    _copy_formatted)
        txt.bind("<Control-c>", _copy_formatted)

        # Button bar
        bar = ttk.Frame(dlg)
        bar.pack(fill="x", padx=8, pady=(0, 8))

        def _copy_plain():
            content = path.read_text(encoding="utf-8") if path.exists() else ""
            self.root.clipboard_clear()
            self.root.clipboard_append(_strip_template_markup(content))
            self._set_status(f"'{title}' copied to clipboard.")

        def _copy_image():
            img_filename = _COMMON_RESPONSE_IMG_FILE.read_text(encoding="utf-8").strip() if _COMMON_RESPONSE_IMG_FILE.exists() else ""
            if not img_filename:
                self._set_status("No image configured.")
                return
            img_path = Path(__file__).parent / img_filename
            if not img_path.exists():
                self._set_status(f"Image file not found: {img_path}")
                return
            if _copy_image_to_clipboard(img_path):
                self._set_status(f"Image copied to clipboard.")
            else:
                self._set_status(f"Failed to copy image (requires pywin32).")

        def _edit():
            dlg.destroy()
            self._open_template_edit(title, path)

        ttk.Button(bar, text="Copy (Formatted)", command=_copy_formatted).pack(side="left")
        ttk.Button(bar, text="Copy as Text",     command=_copy_plain).pack(side="left", padx=6)
        # Show Copy Image button if image is configured
        if _COMMON_RESPONSE_IMG_FILE.exists() and _COMMON_RESPONSE_IMG_FILE.read_text(encoding="utf-8").strip():
            ttk.Button(bar, text="Copy Image",   command=_copy_image).pack(side="left", padx=6)
        ttk.Button(bar, text="Edit…",            command=_edit).pack(side="left", padx=6)
        ttk.Button(bar, text="Close",            command=dlg.destroy).pack(side="right")

        dlg.update_idletasks()
        px = self.root.winfo_rootx() + (self.root.winfo_width()  - dlg.winfo_width())  // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{max(0, px)}+{max(0, py)}")

    def _open_template_edit(self, title: str, path: Path) -> None:
        """Open a simple markup editor for an email template file."""
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Edit — {title}")
        dlg.geometry("620x480")
        dlg.resizable(True, True)
        dlg.transient(self.root)

        # Formatting toolbar
        toolbar = ttk.Frame(dlg)
        toolbar.pack(fill="x", padx=8, pady=(8, 2))

        txt_ref: list[tk.Text] = []   # filled after txt is created

        def _wrap(marker: str):
            txt = txt_ref[0]
            try:
                sel = txt.get("sel.first", "sel.last")
                txt.delete("sel.first", "sel.last")
                txt.insert("insert", f"{marker}{sel}{marker}")
            except tk.TclError:
                txt.insert("insert", f"{marker}{marker}")
                # Move cursor between the markers
                idx = txt.index("insert")
                r, c = map(int, idx.split("."))
                txt.mark_set("insert", f"{r}.{c - len(marker)}")

        ttk.Button(toolbar, text="B", width=3,
                   command=lambda: _wrap("**")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="I", width=3,
                   command=lambda: _wrap("*")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="U", width=3,
                   command=lambda: _wrap("__")).pack(side="left", padx=2)
        ttk.Label(toolbar,
                  text="  Select text, then B / I / U   ·   **bold**  *italic*  __underline__",
                  foreground="gray").pack(side="left", padx=6)

        # Text editor
        frm = ttk.Frame(dlg)
        frm.pack(fill="both", expand=True, padx=8, pady=4)
        sb  = ttk.Scrollbar(frm, orient="vertical")
        txt = tk.Text(frm, wrap="word", padx=10, pady=10, undo=True,
                      yscrollcommand=sb.set, font=("", 10))
        sb.config(command=txt.yview)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)
        txt_ref.append(txt)

        content = path.read_text(encoding="utf-8") if path.exists() else ""
        txt.insert("1.0", content)

        # Image selector (if this is Common Response template)
        img_var = None
        if title == "Common Response":
            img_frame = ttk.LabelFrame(dlg, text="Associated Image (optional)", padding=6)
            img_frame.pack(fill="x", padx=8, pady=4)

            saved_img = _COMMON_RESPONSE_IMG_FILE.read_text(encoding="utf-8").strip() if _COMMON_RESPONSE_IMG_FILE.exists() else ""
            img_var = tk.StringVar(value=saved_img)

            ttk.Label(img_frame, text="Image file:").pack(side="left", padx=(0, 4))
            ttk.Entry(img_frame, textvariable=img_var, width=40).pack(side="left", padx=4, fill="x", expand=True)

            def _browse_image():
                f = filedialog.askopenfilename(
                    title="Select Image",
                    filetypes=[("PNG files", "*.png"), ("JPG files", "*.jpg *.jpeg"), ("All files", "*.*")]
                )
                if f:
                    img_var.set(Path(f).name)

            ttk.Button(img_frame, text="Browse…", command=_browse_image, width=10).pack(side="left")

        # Button bar
        bar = ttk.Frame(dlg)
        bar.pack(fill="x", padx=8, pady=(0, 8))

        def _save():
            path.write_text(txt.get("1.0", "end-1c"), encoding="utf-8")
            # Save image path if this is Common Response template
            if title == "Common Response" and 'img_var' in locals():
                img_filename = img_var.get().strip()
                if img_filename:
                    _COMMON_RESPONSE_IMG_FILE.write_text(img_filename, encoding="utf-8")
            self._set_status(f"Saved {path.name}.")
            dlg.destroy()

        def _preview():
            self._open_template_view(title, path)

        ttk.Button(bar, text="Save",    command=_save).pack(side="left")
        ttk.Button(bar, text="Cancel",  command=dlg.destroy).pack(side="left", padx=6)
        ttk.Button(bar, text="Preview", command=_preview).pack(side="right")

        dlg.update_idletasks()
        px = self.root.winfo_rootx() + (self.root.winfo_width()  - dlg.winfo_width())  // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{max(0, px)}+{max(0, py)}")
        txt.focus_set()

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

        _chk(pre, "wf_laser_safety", "Laser Safety")

        # Emailed User + Common Response template link
        emailed_row = ttk.Frame(pre)
        emailed_row.pack(anchor="w", pady=3)
        ttk.Checkbutton(emailed_row, variable=self._wf_vars["wf_emailed"],
                        command=_save_wf).pack(side="left")
        ttk.Label(emailed_row, text="Emailed User").pack(side="left")
        cr_link = ttk.Label(emailed_row, text="  Common Response",
                            foreground="#1565C0", cursor="hand2")
        cr_link.pack(side="left", padx=(8, 0))
        cr_link.bind("<Button-1>",
                     lambda e: self._open_template_view("Common Response",
                                                        _COMMON_RESPONSE_FILE))
        ttk.Button(emailed_row, text="Edit", width=5,
                   command=lambda: self._open_template_edit("Common Response",
                                                            _COMMON_RESPONSE_FILE),
                   ).pack(side="left", padx=4)
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

        # CVRI-Access template
        cvri_access_row = ttk.Frame(pre)
        cvri_access_row.pack(anchor="w", pady=2, padx=(20, 0))
        cvri_access_link = ttk.Label(cvri_access_row, text="CVRI-Access",
                                     foreground="#1565C0", cursor="hand2")
        cvri_access_link.pack(side="left")
        cvri_access_link.bind("<Button-1>",
                              lambda e: self._open_template_view("CVRI-Access",
                                                                 _CVRI_ACCESS_FILE))
        ttk.Button(cvri_access_row, text="Edit", width=5,
                   command=lambda: self._open_template_edit("CVRI-Access",
                                                            _CVRI_ACCESS_FILE),
                   ).pack(side="left", padx=4)

        # Post-Training column
        post = ttk.LabelFrame(cols, text="Post-Training", padding=10)
        post.pack(side="left", fill="both", expand=True, padx=(6, 0))

        _chk(post, "wf_post_email", "Post-Training Email")

        # CALM Welcome template
        calm_row = ttk.Frame(post)
        calm_row.pack(anchor="w", pady=2, padx=(20, 0))
        calm_link = ttk.Label(calm_row, text="CALM Welcome",
                              foreground="#1565C0", cursor="hand2")
        calm_link.pack(side="left")
        calm_link.bind("<Button-1>",
                       lambda e: self._open_template_view("CALM Welcome",
                                                          _CALM_WELCOME_FILE))
        ttk.Button(calm_row, text="Edit", width=5,
                   command=lambda: self._open_template_edit("CALM Welcome",
                                                            _CALM_WELCOME_FILE),
                   ).pack(side="left", padx=4)

        # CVRI Welcome template
        cvri_row = ttk.Frame(post)
        cvri_row.pack(anchor="w", pady=2, padx=(20, 0))
        cvri_link = ttk.Label(cvri_row, text="CVRI Welcome",
                              foreground="#1565C0", cursor="hand2")
        cvri_link.pack(side="left")
        cvri_link.bind("<Button-1>",
                       lambda e: self._open_template_view("CVRI Welcome",
                                                          _CVRI_WELCOME_FILE))
        ttk.Button(cvri_row, text="Edit", width=5,
                   command=lambda: self._open_template_edit("CVRI Welcome",
                                                            _CVRI_WELCOME_FILE),
                   ).pack(side="left", padx=4)
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

    # ── Class Schedule tab ────────────────────────────────────────────────────

    def _build_class_schedule_tab(self) -> None:
        tab = ttk.Frame(self._notebook, padding=8)
        self._notebook.add(tab, text="  Class Schedule  ")

        # All sessions; selected session index
        self._cs_sessions: list[dict] = []
        self._cs_sel_idx: int | None  = None

        # Fields for the currently selected session
        self._cs_date_var       = tk.StringVar()
        self._cs_instructor_var = tk.StringVar()
        self._cs_time_var       = tk.StringVar()
        self._cs_location_var   = tk.StringVar()

        # ── Horizontal pane: sessions list (left) | session detail (right) ────
        paned = ttk.PanedWindow(tab, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # ── Left: upcoming sessions list ──────────────────────────────────────
        left = ttk.Frame(paned, padding=(0, 0, 6, 0))
        paned.add(left, weight=1)

        ttk.Label(left, text="Upcoming Sessions",
                  font=("", 9, "bold")).pack(anchor="w", pady=(0, 3))

        sess_frame = ttk.Frame(left)
        sess_frame.pack(fill="both", expand=True)

        self._cs_sess_tree = ttk.Treeview(
            sess_frame, columns=("date", "instructor"),
            show="headings", selectmode="browse", height=14)
        self._cs_sess_tree.heading("date",       text="Date")
        self._cs_sess_tree.heading("instructor", text="Instructor")
        self._cs_sess_tree.column("date",        width=90,  minwidth=70)
        self._cs_sess_tree.column("instructor",  width=110, minwidth=70)
        self._cs_sess_tree.bind("<<TreeviewSelect>>", self._cs_on_session_select)

        sess_vsb = ttk.Scrollbar(sess_frame, orient="vertical",
                                  command=self._cs_sess_tree.yview)
        self._cs_sess_tree.configure(yscrollcommand=sess_vsb.set)
        sess_vsb.pack(side="right", fill="y")
        self._cs_sess_tree.pack(fill="both", expand=True)

        sess_btns = ttk.Frame(left)
        sess_btns.pack(fill="x", pady=(4, 0))
        ttk.Button(sess_btns, text="+ New Session",
                   command=self._cs_new_session).pack(side="left", padx=(0, 4))
        ttk.Button(sess_btns, text="Delete",
                   command=self._cs_delete_session).pack(side="left")

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=8)

        ttk.Button(left, text="Export Session to Log →",
                   command=self._cs_export).pack(anchor="w")
        self._cs_export_lbl = ttk.Label(left, text="", foreground="#2E7D32")
        self._cs_export_lbl.pack(anchor="w", pady=(2, 0))

        # ── Right: session detail (info + students) ───────────────────────────
        right = ttk.Frame(paned, padding=(6, 0, 0, 0))
        paned.add(right, weight=3)

        # Session info fields
        info = ttk.LabelFrame(right, text="Session Info", padding=8)
        info.pack(fill="x", pady=(0, 6))

        LW = 11
        ttk.Label(info, text="Date:", anchor="e", width=LW).grid(
            row=0, column=0, sticky="e", padx=4, pady=3)
        date_f = ttk.Frame(info)
        date_f.grid(row=0, column=1, sticky="w", padx=4, pady=3)
        ttk.Entry(date_f, textvariable=self._cs_date_var, width=13).pack(side="left")
        ttk.Button(date_f, text="📅", width=3,
                   command=self._cs_pick_date).pack(side="left", padx=(4, 0))

        ttk.Label(info, text="Instructor:", anchor="e", width=LW).grid(
            row=0, column=2, sticky="e", padx=(16, 4), pady=3)
        ttk.Combobox(info, textvariable=self._cs_instructor_var,
                     values=TEAM_MEMBERS, width=20).grid(
            row=0, column=3, sticky="w", padx=4, pady=3)

        ttk.Label(info, text="Time:", anchor="e", width=LW).grid(
            row=1, column=0, sticky="e", padx=4, pady=3)
        ttk.Entry(info, textvariable=self._cs_time_var, width=14).grid(
            row=1, column=1, sticky="w", padx=4, pady=3)

        ttk.Label(info, text="Location:", anchor="e", width=LW).grid(
            row=1, column=2, sticky="e", padx=(16, 4), pady=3)
        ttk.Entry(info, textvariable=self._cs_location_var, width=22).grid(
            row=1, column=3, sticky="w", padx=4, pady=3)

        ttk.Button(info, text="Save Session Info",
                   command=self._cs_save_info).grid(
            row=2, column=0, columnspan=4, sticky="w", padx=4, pady=(6, 0))

        # Students table
        stu_frame = ttk.LabelFrame(right, text="Students in Class", padding=6)
        stu_frame.pack(fill="both", expand=True)

        self._cs_tree = ttk.Treeview(
            stu_frame, columns=("name", "pi"),
            show="headings", selectmode="browse", height=8)
        self._cs_tree.heading("name", text="Student Name")
        self._cs_tree.heading("pi",   text="PI / Lab")
        self._cs_tree.column("name", width=210, minwidth=100)
        self._cs_tree.column("pi",   width=190, minwidth=80)
        self._cs_tree.bind("<Double-1>", self._cs_edit_student)

        stu_vsb = ttk.Scrollbar(stu_frame, orient="vertical",
                                 command=self._cs_tree.yview)
        self._cs_tree.configure(yscrollcommand=stu_vsb.set)
        stu_vsb.pack(side="right", fill="y")
        self._cs_tree.pack(fill="both", expand=True)

        stu_btns = ttk.Frame(stu_frame)
        stu_btns.pack(fill="x", pady=(4, 0))
        ttk.Button(stu_btns, text="+ Add from Selected Request",
                   command=self._cs_add_from_selection).pack(side="left", padx=(0, 6))
        ttk.Button(stu_btns, text="+ Add Manually",
                   command=self._cs_add_student).pack(side="left", padx=(0, 6))
        ttk.Button(stu_btns, text="Remove",
                   command=self._cs_remove_student).pack(side="left")

        self._cs_load()

    # ── Class Schedule handlers ───────────────────────────────────────────────

    def _cs_load(self) -> None:
        """Load all sessions from JSON and populate the sessions list."""
        try:
            raw = json.loads(_CS_SESSION_FILE.read_text(encoding="utf-8"))
            # Support old single-session format (dict) and new list format
            if isinstance(raw, dict):
                raw = [raw]
            self._cs_sessions = raw
        except Exception:
            self._cs_sessions = []
        self._cs_refresh_sessions_tree()

    def _cs_persist(self) -> None:
        """Write all sessions to JSON."""
        try:
            _CS_SESSION_FILE.write_text(
                json.dumps(self._cs_sessions, indent=2, ensure_ascii=False),
                encoding="utf-8")
        except Exception as exc:
            self._set_status(f"Could not save sessions: {exc}")

    def _cs_refresh_sessions_tree(self) -> None:
        """Rebuild the left-hand sessions list."""
        self._cs_sess_tree.delete(*self._cs_sess_tree.get_children())
        for i, s in enumerate(self._cs_sessions):
            self._cs_sess_tree.insert("", "end", iid=str(i),
                                       values=(s.get("date", ""),
                                               s.get("instructor", "")))

    def _cs_on_session_select(self, _event=None) -> None:
        """Load the selected session's info and students into the right panel."""
        sel = self._cs_sess_tree.selection()
        if not sel:
            return
        self._cs_sel_idx = int(sel[0])
        sess = self._cs_sessions[self._cs_sel_idx]
        self._cs_date_var.set(sess.get("date", ""))
        self._cs_instructor_var.set(sess.get("instructor", ""))
        self._cs_time_var.set(sess.get("time", ""))
        self._cs_location_var.set(sess.get("location", ""))
        self._cs_export_lbl.config(text="")
        self._cs_refresh_students_tree()

    def _cs_new_session(self) -> None:
        """Append a blank session and select it."""
        self._cs_sessions.append(
            {"date": "", "instructor": "", "time": "", "location": "",
             "students": []})
        self._cs_persist()
        self._cs_refresh_sessions_tree()
        new_iid = str(len(self._cs_sessions) - 1)
        self._cs_sess_tree.selection_set(new_iid)
        self._cs_sess_tree.see(new_iid)

    def _cs_delete_session(self) -> None:
        if self._cs_sel_idx is None:
            return
        sess = self._cs_sessions[self._cs_sel_idx]
        label = sess.get("date") or f"session {self._cs_sel_idx + 1}"
        if not messagebox.askyesno("Delete Session",
                                   f"Delete {label} and all its students?"):
            return
        self._cs_sessions.pop(self._cs_sel_idx)
        self._cs_sel_idx = None
        self._cs_persist()
        self._cs_refresh_sessions_tree()
        # Clear the right panel
        self._cs_date_var.set("")
        self._cs_instructor_var.set("")
        self._cs_time_var.set("")
        self._cs_location_var.set("")
        self._cs_refresh_students_tree()
        self._cs_export_lbl.config(text="")

    def _cs_save_info(self) -> None:
        """Save the session info fields back into the sessions list."""
        if self._cs_sel_idx is None:
            messagebox.showwarning("No Session Selected",
                                   "Select or create a session first.")
            return
        sess = self._cs_sessions[self._cs_sel_idx]
        sess["date"]       = self._cs_date_var.get().strip()
        sess["instructor"] = self._cs_instructor_var.get().strip()
        sess["time"]       = self._cs_time_var.get().strip()
        sess["location"]   = self._cs_location_var.get().strip()
        self._cs_persist()
        # Refresh the sessions list so date/instructor column updates
        sel_iid = str(self._cs_sel_idx)
        self._cs_refresh_sessions_tree()
        self._cs_sess_tree.selection_set(sel_iid)
        self._set_status("Session info saved.")

    def _cs_refresh_students_tree(self) -> None:
        self._cs_tree.delete(*self._cs_tree.get_children())
        if self._cs_sel_idx is None:
            return
        students = self._cs_sessions[self._cs_sel_idx].get("students", [])
        for i, s in enumerate(students):
            self._cs_tree.insert("", "end", iid=str(i),
                                  values=(s.get("name", ""), s.get("pi", "")))

    def _cs_add_from_selection(self) -> None:
        """Add the currently selected iLab request's requester to the student list."""
        if self._cs_sel_idx is None:
            messagebox.showwarning("No Session Selected",
                                   "Select or create a session first.")
            return
        if not self._current_rec:
            messagebox.showwarning("No Request Selected",
                                   "Select a service request in the table above first,\n"
                                   "then click Add from Selected Request.")
            return
        name = (self._current_rec.get("owner_name") or "").strip()
        pi   = (self._current_rec.get("pi_name")    or "").strip()
        if not name:
            messagebox.showwarning("No Name",
                                   "The selected request has no requester name.")
            return
        # Pre-fill the dialog so the user can confirm / tweak before adding
        self._cs_student_dialog(prefill={"name": name, "pi": pi})

    def _cs_add_student(self) -> None:
        if self._cs_sel_idx is None:
            messagebox.showwarning("No Session Selected",
                                   "Select or create a session first.")
            return
        self._cs_student_dialog()

    def _cs_edit_student(self, _event=None) -> None:
        sel = self._cs_tree.selection()
        if sel:
            self._cs_student_dialog(edit_idx=int(sel[0]))

    def _cs_student_dialog(self, edit_idx: int | None = None,
                            prefill: dict | None = None) -> None:
        students = self._cs_sessions[self._cs_sel_idx].get("students", [])
        existing = students[edit_idx] if edit_idx is not None else (prefill or {})

        dlg = tk.Toplevel(self.root)
        dlg.title("Edit Student" if edit_idx is not None else "Add Student")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        frame = ttk.Frame(dlg, padding=16)
        frame.pack()

        name_var = tk.StringVar(value=existing.get("name", ""))
        pi_var   = tk.StringVar(value=existing.get("pi", ""))

        ttk.Label(frame, text="Student Name:").grid(
            row=0, column=0, sticky="e", padx=4, pady=4)
        name_entry = ttk.Entry(frame, textvariable=name_var, width=30)
        name_entry.grid(row=0, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(frame, text="PI / Lab:").grid(
            row=1, column=0, sticky="e", padx=4, pady=4)
        ttk.Entry(frame, textvariable=pi_var, width=30).grid(
            row=1, column=1, sticky="w", padx=4, pady=4)

        def _save():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Missing Name", "Enter a student name.",
                                       parent=dlg)
                return
            student = {"name": name, "pi": pi_var.get().strip()}
            if edit_idx is not None:
                students[edit_idx] = student
            else:
                students.append(student)
            self._cs_sessions[self._cs_sel_idx]["students"] = students
            self._cs_refresh_students_tree()
            self._cs_persist()
            dlg.destroy()

        btn = ttk.Frame(frame)
        btn.grid(row=2, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(btn, text="Save",   command=_save).pack(side="left", padx=(0, 6))
        ttk.Button(btn, text="Cancel", command=dlg.destroy).pack(side="left")

        name_entry.focus_set()
        name_entry.select_range(0, "end")
        dlg.bind("<Return>", lambda e: _save())

    def _cs_remove_student(self) -> None:
        if self._cs_sel_idx is None:
            return
        sel = self._cs_tree.selection()
        if not sel:
            return
        students = self._cs_sessions[self._cs_sel_idx].get("students", [])
        students.pop(int(sel[0]))
        self._cs_sessions[self._cs_sel_idx]["students"] = students
        self._cs_refresh_students_tree()
        self._cs_persist()

    def _cs_pick_date(self) -> None:
        raw = self._cs_date_var.get().strip()
        initial = None
        try:
            initial = _date.fromisoformat(raw)
        except ValueError:
            pass
        CalendarPicker(self.root, self._cs_date_selected, initial)

    def _cs_date_selected(self, date_str: str) -> None:
        self._cs_date_var.set(date_str)
        # Auto-save the date into the session immediately
        if self._cs_sel_idx is not None:
            self._cs_sessions[self._cs_sel_idx]["date"] = date_str
            self._cs_persist()
            sel_iid = str(self._cs_sel_idx)
            self._cs_refresh_sessions_tree()
            self._cs_sess_tree.selection_set(sel_iid)

    def _cs_export(self) -> None:
        if self._cs_sel_idx is None:
            messagebox.showwarning("No Session Selected",
                                   "Select a session to export.")
            return
        p = _prefs.get_prefs()
        xlsx_path = str(p.get("intro_xlsx", "") or "").strip()
        if not xlsx_path:
            messagebox.showwarning(
                "xlsx Path Not Set",
                "Set the Microscope Intro Course Log path in ⚙ Preferences.",
            )
            return
        if not HAS_OPENPYXL:
            messagebox.showerror(
                "openpyxl Required",
                "Install openpyxl to export to xlsx:\n\n    pip install openpyxl",
            )
            return

        sess = self._cs_sessions[self._cs_sel_idx]
        if not sess.get("students"):
            messagebox.showwarning("No Students",
                                   "Add at least one student before exporting.")
            return

        sheet_name = str(p.get("intro_sheet", "") or "").strip()
        try:
            ok, result = self._run_export_with_retry(
                lambda: append_class_session(sess, xlsx_path, sheet_name=sheet_name),
                silent=False,
            )
            if not ok:
                return
            n   = result.get("rows_written", 0)
            dup = result.get("duplicates_skipped", 0)
            dup_lbl = f", {dup} duplicate(s) skipped" if dup else ""
            self._cs_export_lbl.config(text=f"✓ {n} row(s) written{dup_lbl}")
            self._set_status(
                f"Session exported to {Path(xlsx_path).name}"
                f" — {n} student row(s) written{dup_lbl}.")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    # ── Status bar ────────────────────────────────────────────────────────────

    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self.root, relief="sunken", padding=(6, 2))
        bar.pack(fill="x", side="bottom")
        self._status_var = tk.StringVar(value="Ready — click ↻ Sync from iLab to load data.")
        ttk.Label(bar, textvariable=self._status_var, anchor="w").pack(side="left")
        self._last_sync_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self._last_sync_var,
                  anchor="e", foreground="#666").pack(side="right")
        self._autosave_status_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self._autosave_status_var,
                  anchor="e", foreground="#888").pack(side="right", padx=(0, 12))

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
            try:
                self._tree.selection_set(req_id)
                self._tree.see(req_id)
            except tk.TclError:
                pass
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
                self.root.after(50, lambda r=req_id: (
                    self._tree.selection_set(r), self._tree.see(r)))
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
                self.root.after(0, self._cleanup_completed_local_entries)
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

        # Check if request already has charges that would be exceeded
        core_id = self._get_core_id()
        if core_id is not None:
            try:
                client = self._get_client()
                existing_total = client.get_total_charges(core_id, int(req_id))
                if existing_total + total > _MAX_CHARGE:
                    self._class_taken_var.set(False)
                    messagebox.showerror(
                        "Charge Limit Exceeded",
                        f"Request {req_id} already has ${existing_total:,.2f} in charges.\n\n"
                        f"Adding ${total:,.2f} would total ${existing_total + total:,.2f}, "
                        f"exceeding the maximum of ${_MAX_CHARGE:,.2f}.",
                    )
                    return
            except Exception as e:
                # If we can't check existing charges, show a warning but continue
                if "validate_min_charge" not in str(e):
                    self._class_taken_var.set(False)
                    messagebox.showerror("Error Checking Charges", str(e))
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

        if core_id is None:
            self._class_taken_var.set(False)
            return

        # Lock UI immediately so a second click before the API returns can't submit again
        self._current_rec["class_taken"] = "1"
        self._update_class_taken_ui(True)
        self._class_status_lbl.config(text="Submitting charge to iLab…")

        def worker():
            try:
                client  = self._get_client()
                result  = client.add_charges(core_id, int(req_id), [{
                    "quantity":   qty,
                    "price_id":   int(price_id),
                    "service_id": int(svc_id),
                    "note":       "Class attendance",
                }], max_charge=_MAX_CHARGE)
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
                self.root.after(0, lambda: self._class_status_lbl.config(
                    text="✓ Charge submitted to iLab — cannot be re-submitted.",
                    foreground="#9E9E9E"))
                self.root.after(0, self._refresh_table)
                self.root.after(0, lambda: self._set_status(
                    f"Class charge added to request {req_id} in iLab."))
            except Exception as exc:
                # Revert optimistic lock so user can retry
                self._current_rec["class_taken"] = "0"
                self.root.after(0, lambda: self._class_taken_var.set(False))
                self.root.after(0, lambda: self._update_class_taken_ui(False))
                self.root.after(0, lambda: self._class_status_lbl.config(text=""))
                self.root.after(0, lambda e=exc: messagebox.showerror("Charge Error", str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _on_export_to_schedule(self, silent: bool = False) -> bool:
        """
        Append the current record to the appropriate training schedule xlsx.
        silent=True suppresses the success dialog (used for auto-export on completion).

        Returns True if the row was actually written, False otherwise (so callers
        know not to discard/clear data that was never successfully exported).
        """
        if not self._current_rec:
            if not silent:
                messagebox.showwarning("No Selection", "Select a request first.")
            return False

        if not silent:
            # Save any unsaved training fields first (interactive mode only —
            # in silent/auto-export mode the UI may reflect a different record)
            self._on_save_training()

        req_id = self._current_rec["request_id"]
        # Read core from the record itself so auto-export works even when the
        # Training tab is showing a different request's fields
        core   = self._current_rec.get("core_lab", "").strip().upper()
        if not core and not silent:
            # Fall back to the UI widget in case the user just changed it
            core = self._training_core_var.get().strip().upper()
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
            return False

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
            return False

        if not HAS_OPENPYXL:
            if not silent:
                messagebox.showerror(
                    "openpyxl Required",
                    "Install openpyxl to export to xlsx:\n\n    pip install openpyxl",
                )
            return False

        sheet_name = str(p.get(f"{core.lower()}_sheet", "") or "").strip()

        ok, result = self._run_export_with_retry(
            lambda: append_training_row(self._current_rec, xlsx_path,
                                        sheet_name=sheet_name),
            silent=silent,
        )
        if not ok:
            return False

        # Mark as exported so auto-export doesn't run again
        self._data.update_local_fields(req_id, schedule_exported="1")
        self._current_rec["schedule_exported"] = "1"
        self._exported_lbl.config(text="✓ Exported to records")
        fname     = Path(xlsx_path).name
        sheet_lbl = f" → '{sheet_name}'" if sheet_name else ""

        if result.get("duplicate"):
            self._set_status(
                f"Row already present in {fname}{sheet_lbl} ({core}) — "
                "skipped duplicate.")
            if not silent:
                messagebox.showinfo(
                    "Export to Records",
                    f"A matching row already exists in {fname} — "
                    "skipped to avoid writing a duplicate.",
                )
            return True

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
        return True

    def _run_export_with_retry(self, fn, silent: bool):
        """
        Call the zero-arg *fn* (an xlsx export call). If it fails because the
        target file is open in Excel (PermissionError), offer a Retry/Cancel
        prompt in interactive mode so the user can close Excel and retry
        without losing the data; in silent/auto-export mode, just report the
        failure via the status bar instead of blocking with a dialog.

        Returns (True, result) on success, (False, None) if the export was
        never completed. Non-lock exceptions propagate to the caller.
        """
        while True:
            try:
                return True, fn()
            except PermissionError as exc:
                if silent:
                    self._set_status(f"⚠ Export skipped — {exc}")
                    return False, None
                if not messagebox.askretrycancel("File In Use", str(exc)):
                    return False, None
                # Loop back and retry the save.

    def _cleanup_completed_local_entries(self) -> None:
        """Run after an iLab sync: local-only manual entries don't come back
        from iLab, so sync's own stale-record cleanup never touches them.
        Sweep for any that are marked completed and clear them out (exporting
        to the schedule first if that hasn't happened yet)."""
        for rec in list(self._data.all_records()):
            if rec.get("local_only") == "1" and rec.get("state") == "completed":
                self._auto_export_on_complete(rec["request_id"])

    def _auto_export_on_complete(self, req_id: str) -> None:
        """Called after a request is marked completed; exports if not already done,
        then clears training fields from the record and resets the Training UI.

        Local-only manual entries are deleted outright once export is confirmed
        (rather than just having their training fields wiped), since they have
        no corresponding iLab request and exist only to get into the schedule.
        """
        rec = self._data.get_record(req_id)
        if not rec:
            return
        already_exported = rec.get("schedule_exported", "0") == "1"

        if not already_exported:
            # Temporarily point _current_rec at this record so the export works
            saved = self._current_rec
            self._current_rec = rec
            exported = self._on_export_to_schedule(silent=True)
            self._current_rec = saved

            if not exported:
                # Export didn't happen (e.g. xlsx open in Excel) — keep the
                # training fields intact so nothing is lost; the row will be
                # exported (and only then cleared) next time it succeeds.
                return

        if rec.get("local_only") == "1":
            # Local-only manual entries have no corresponding iLab request —
            # once the row is confirmed written to the schedule there's
            # nothing left worth keeping, so drop the placeholder entirely
            # instead of just wiping its training fields. This also cleans
            # up any local entries that were exported earlier but, for
            # whatever reason, never got cleared (e.g. an older cache).
            self._data.delete_record(req_id)
            if self._current_rec and self._current_rec.get("request_id") == req_id:
                self._current_rec = None
            self._refresh_table()
            self._set_status(
                f"Local entry {req_id} written to schedule and cleared.")
            return

        if already_exported:
            return

        # After export: wipe training fields so the record is clean
        _CLEAR = dict(
            training_date="", training_day="", training_time="",
            microscope="", core_lab="", local_notes="",
        )
        self._data.update_local_fields(req_id, **_CLEAR)
        rec.update(_CLEAR)

        # If this record is still on screen, reset the Training tab widgets
        if self._current_rec and self._current_rec.get("request_id") == req_id:
            self._current_rec.update(_CLEAR)
            self._training_date_var.set("")
            self._training_day_var.set("")
            self._training_time_var.set("")
            self._training_micro_var.set("")
            self._training_core_var.set("")
            try:
                self._notes_text.delete("1.0", "end")
            except Exception:
                pass

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
        # Keep detail panel in sync if this is the currently displayed record
        if self._current_rec and self._current_rec.get("request_id") == row_id:
            self._current_rec["state"] = new_state
            self._info_vars["state"].set(new_state)
        self._refresh_table()
        try:
            self._tree.selection_set(row_id)
            self._tree.see(row_id)
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

    def _copy_info_field(self, key: str) -> None:
        """Copy the current value of an iLab info field (e.g. owner_email) to
        the system clipboard."""
        value = self._info_vars.get(key, tk.StringVar()).get().strip()
        if not value:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self._set_status(f"Copied '{value}' to clipboard.")

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

    def _schedule_autosave(self, *_) -> None:
        """Debounce: reset the 5-second auto-save countdown on any change."""
        if self._autosave_job:
            self.root.after_cancel(self._autosave_job)
        self._autosave_job = self.root.after(5_000, self._do_autosave)
        self._autosave_status_var.set("● unsaved")

    def _do_autosave(self) -> None:
        """Flush all pending UI edits and save to disk."""
        self._autosave_job = None
        # Pause the change callback so the saves below don't re-trigger the timer
        old_on_change = self._data.on_change
        self._data.on_change = None
        try:
            if self._current_rec:
                req_id = self._current_rec["request_id"]
                # Flush Info tab local fields (notes, assigned-to, labels)
                active_lbls = ",".join(l for l, v in self._label_vars.items() if v.get())
                notes       = self._notes_text.get("1.0", "end-1c")
                assigned    = self._assigned_var.get()
                self._data.update_local_fields(req_id, assigned_to=assigned,
                                               labels=active_lbls, local_notes=notes)
                self._current_rec.update(assigned_to=assigned,
                                         labels=active_lbls, local_notes=notes)
                # Flush Training tab fields
                self._on_save_training()
            # Flush Class Schedule session info if a session is selected
            if getattr(self, "_cs_sel_idx", None) is not None:
                self._cs_save_info()
            self._data.save()
        finally:
            self._data.on_change = old_on_change
        when = datetime.now().strftime("%I:%M %p")
        self._autosave_status_var.set(f"✓ auto-saved {when}")

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
            ("intro_xlsx",""), ("intro_sheet",""),
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

        # ── Section: intro course log ─────────────────────────────────────────
        ttk.Label(self, text="Microscope Intro Course Log",
                  font=("", 10, "bold")).grid(
            row=11, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))

        ttk.Label(self, text="Intro Log xlsx:", width=14, anchor="e").grid(
            row=12, column=0, sticky="e", **PAD)
        ttk.Entry(self, textvariable=self._vars["intro_xlsx"], width=40).grid(
            row=12, column=1, sticky="w", **PAD)
        ttk.Button(self, text="Browse…",
                   command=lambda: self._browse("intro_xlsx")).grid(
            row=12, column=2, **PAD)

        ttk.Label(self, text="Sheet name:", width=14, anchor="e").grid(
            row=13, column=0, sticky="e", padx=8, pady=(0, 6))
        ttk.Entry(self, textvariable=self._vars["intro_sheet"], width=22).grid(
            row=13, column=1, sticky="w", padx=8, pady=(0, 6))
        ttk.Label(self, text="(blank = first sheet)",
                  foreground=self._fg2).grid(
            row=13, column=2, sticky="w", padx=4, pady=(0, 6))

        ttk.Label(
            self,
            text="Points to your Microscope Intro Course Log.xlsx.\n"
                 "Used by the Class Schedule tab to log weekly class sessions.",
            foreground=self._fg2,
        ).grid(row=14, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))

        ttk.Separator(self, orient="horizontal").grid(
            row=15, column=0, columnspan=4, sticky="ew", padx=10, pady=6)

        # ── Section: class charge ─────────────────────────────────────────────
        ttk.Label(self, text="Class Charge (iLab API)",
                  font=("", 10, "bold")).grid(
            row=16, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))

        for i, (label, key) in enumerate([
            ("Service ID:",     "class_service_id"),
            ("Price ID:",       "class_price_id"),
            ("Quantity:",       "class_quantity"),
            ("Unit Price ($):", "class_unit_price"),
            ("Max Charge ($):", "max_charge"),
        ], start=17):
            ttk.Label(self, text=label, width=14, anchor="e").grid(
                row=i, column=0, sticky="e", **PAD)
            ttk.Entry(self, textvariable=self._vars[key], width=14).grid(
                row=i, column=1, sticky="w", **PAD)

        ttk.Label(
            self,
            text="Run  python get_services.py  to look up Service ID and Price ID.\n"
                 "Leave blank to save Class Taken locally without calling iLab.",
            foreground=self._fg2,
        ).grid(row=22, column=0, columnspan=4, sticky="w", padx=12, pady=(2, 8))

        ttk.Separator(self, orient="horizontal").grid(
            row=23, column=0, columnspan=4, sticky="ew", padx=10, pady=4)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = ttk.Frame(self, padding=(8, 4, 12, 10))
        btn_row.grid(row=24, column=0, columnspan=4, sticky="e")
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
        if "intro" in key:
            label = "Microscope Intro Course Log"
        elif "calm" in key:
            label = "CALM Training Schedule"
        else:
            label = "CVRI Training Schedule"
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            title=f"Select {label} xlsx",
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
