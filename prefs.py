"""
Persistent user preferences stored in prefs.json.

Usage::

    from prefs import get_prefs, save_prefs
    p = get_prefs()
    p["calm_xlsx"] = "C:/path/to/Training Schedule_CALM_CURRENT.xlsx"
    save_prefs(p)
"""

import json
import os
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
_PREFS_FILE = APP_DIR / "prefs.json"

# Every shared data file (the xlsx schedules, the cache CSV, the email
# templates) lives in this one folder.  The folder name really does begin with
# an invisible U+200E (left-to-right mark).
DEFAULT_SHARED_DIR = "X:\\\u200eiLabdsRequestManager"

# prefs key -> file name inside the shared folder.  These keys are always
# derived from the shared folder; any value stored in prefs.json is ignored.
SHARED_FILES: dict = {
    "data_file":  "ilab_requests_cache.csv",
    "calm_xlsx":  "Training Schedule_CALM_CURRENT.xlsx",
    "cvri_xlsx":  "Training Schedule_CVRI_CURRENT.xlsx",
    "intro_xlsx": "Microscope Intro Course Log.xlsx",
}

_DEFAULTS: dict = {
    "dark_mode":        "0",
    "shared_dir":       DEFAULT_SHARED_DIR,   # one folder holding all shared files
    "data_file":        "",   # derived: <shared_dir>/ilab_requests_cache.csv
    "cache_sheet":      "Cache",   # worksheet used when data_file is an Excel Online URL
    "calm_xlsx":        "",   # derived
    "calm_sheet":       "",   # worksheet name to append to; blank = first sheet
    "cvri_xlsx":        "",   # derived
    "cvri_sheet":       "",
    "intro_xlsx":       "",   # derived (Microscope Intro Course Log)
    "intro_sheet":      "",   # worksheet name; blank = first sheet
    "class_service_id": "",
    "class_price_id":   "",
    "class_quantity":   "2",
    "class_unit_price": "100",
    "max_charge":       "200",
    # ── Auto-sync (minutes; 0 = off) ─────────────────────────────────────────
    "sync_on_open":         "1",  # 1 = run Sync All shortly after the app starts
    "idle_sync_minutes":    "5",  # Records sync this long after the last edit
    "records_sync_minutes": "60", # periodic Records & Cache sync
    "ilab_sync_minutes":    "60", # periodic iLab sync
    "pending_on_lock":      "0",  # 1 = automatic exports write a time-stamped temp file when the xlsx is locked
    "class_session_file":   "",   # blank = class_session.json beside the data file
}

_cache: dict | None = None


def get_prefs() -> dict:
    global _cache
    if _cache is not None:
        return _cache
    data: dict = {}
    if _PREFS_FILE.exists():
        try:
            data = json.loads(_PREFS_FILE.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    _cache = {**_DEFAULTS, **{k: str(v) for k, v in data.items()}}
    _apply_shared_paths(_cache)
    return _cache


def shared_dir(prefs: dict | None = None) -> Path:
    """The single shared folder that holds every data/template file."""
    p = prefs if prefs is not None else get_prefs()
    return Path(str(p.get("shared_dir", "") or DEFAULT_SHARED_DIR).strip())


def shared_file(name: str) -> Path:
    """Path of *name* inside the shared folder."""
    return shared_dir() / name


def _apply_shared_paths(p: dict) -> None:
    base = shared_dir(p)
    for key, fname in SHARED_FILES.items():
        p[key] = str(base / fname)


def make_writable(path: Path) -> None:
    """Clear a read-only attribute so the file can be overwritten."""
    try:
        if path.exists() and not os.access(path, os.W_OK):
            os.chmod(path, 0o666)
    except OSError:
        pass


def save_prefs(prefs: dict) -> None:
    global _cache
    _apply_shared_paths(prefs)
    _cache = prefs
    _PREFS_FILE.write_text(
        json.dumps(prefs, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def abs_path(value: str) -> str:
    """Return *value* as an absolute path (relative paths resolve against the
    app directory).  Blank values and http(s) URLs are returned unchanged."""
    value = str(value or "").strip()
    if not value or value.lower().startswith(("http://", "https://")):
        return value
    p = Path(os.path.expandvars(value)).expanduser()
    if not p.is_absolute():
        p = APP_DIR / p
    return str(p.resolve())


def minutes(key: str) -> int:
    """A non-negative integer minutes preference (invalid/blank -> 0)."""
    try:
        return max(0, int(float(str(get_prefs().get(key, "0")).strip() or 0)))
    except ValueError:
        return 0
