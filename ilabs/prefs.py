"""
Persistent user preferences stored in prefs.json.

Usage::

    from prefs import get_prefs, save_prefs
    p = get_prefs()
    p["calm_xlsx"] = "C:/path/to/Training Schedule_CALM_CURRENT.xlsx"
    save_prefs(p)
"""

import json
from pathlib import Path

_PREFS_FILE = Path(__file__).parent / "prefs.json"

_DEFAULTS: dict = {
    "calm_xlsx":        "",
    "cvri_xlsx":        "",
    "class_service_id": "",
    "class_price_id":   "",
    "class_quantity":   "2",
    "class_unit_price": "100",
    "max_charge":       "200",
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
    return _cache


def save_prefs(prefs: dict) -> None:
    global _cache
    _cache = prefs
    _PREFS_FILE.write_text(
        json.dumps(prefs, indent=2, ensure_ascii=False), encoding="utf-8"
    )
