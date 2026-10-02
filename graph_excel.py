"""
Read/write Excel workbooks in OneDrive / SharePoint directly through the
Microsoft Graph Excel REST API, so nothing has to sync through the OneDrive
client.

A workbook is referenced by its normal browser / "Copy link" URL, e.g.
    https://ucsf-my.sharepoint.com/:x:/r/personal/.../Training%20Schedule.xlsx

Auth is interactive user sign-in via MSAL (a browser window opens the first
time; the token is then cached in ~/.ilabs_manager/msal_cache.json).

Configuration (env var or .env):
    GRAPH_CLIENT_ID   Entra app (public client) ID.  Defaults to Microsoft's
                      "Graph Command Line Tools" ID; if UCSF blocks it, ask IT
                      to register an app with delegated Files.ReadWrite.All and
                      redirect URI http://localhost.
    GRAPH_TENANT      Defaults to "organizations".

Requires:  pip install msal requests
"""

import base64
import os
import re
import threading
import time
from pathlib import Path
from urllib.parse import quote

import requests

try:
    import msal
    HAS_MSAL = True
except ImportError:
    HAS_MSAL = False

GRAPH = "https://graph.microsoft.com/v1.0"
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"   # Graph Command Line Tools
SCOPES = (os.environ.get("GRAPH_SCOPES") or "Files.ReadWrite.All").split()
_CACHE_PATH = Path.home() / ".ilabs_manager" / "msal_cache.json"

TEXT_CELL_LIMIT = 32000      # Excel hard limit is 32,767 chars per cell


def is_graph_url(value: str) -> bool:
    return str(value or "").strip().lower().startswith(("http://", "https://"))


class GraphError(Exception):
    def __init__(self, message: str, status: int = 0, code: str = ""):
        super().__init__(message)
        self.status = status
        self.code = code


# ── Auth ──────────────────────────────────────────────────────────────────────

_auth_lock = threading.Lock()
_app = None
_token_cache = None


def _get_app():
    global _app, _token_cache
    if not HAS_MSAL:
        raise ImportError("msal is required for Excel Online access.\n"
                          "Install it with:  pip install msal")
    if _app is None:
        _token_cache = msal.SerializableTokenCache()
        if _CACHE_PATH.exists():
            try:
                _token_cache.deserialize(_CACHE_PATH.read_text(encoding="utf-8"))
            except Exception:
                pass
        client_id = os.environ.get("GRAPH_CLIENT_ID") or DEFAULT_CLIENT_ID
        tenant = os.environ.get("GRAPH_TENANT") or "organizations"
        _app = msal.PublicClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant}",
            token_cache=_token_cache,
        )
    return _app


def _persist_cache() -> None:
    if _token_cache is not None and _token_cache.has_state_changed:
        _CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        _CACHE_PATH.write_text(_token_cache.serialize(), encoding="utf-8")


def get_token(interactive: bool = True) -> str:
    """Return a valid access token, signing in through the browser if needed."""
    with _auth_lock:
        app = _get_app()
        result = None
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result and interactive:
            result = app.acquire_token_interactive(SCOPES)
        _persist_cache()
        if not result or "access_token" not in result:
            desc = (result or {}).get("error_description", "not signed in")
            raise GraphError(f"Microsoft sign-in failed: {desc}", 401)
        return result["access_token"]


def sign_out() -> None:
    with _auth_lock:
        app = _get_app()
        for acct in app.get_accounts():
            app.remove_account(acct)
        _persist_cache()


# ── HTTP ──────────────────────────────────────────────────────────────────────

def _request(method: str, path: str, *, json=None, params=None):
    url = path if path.startswith("http") else GRAPH + path
    for attempt in range(5):
        resp = requests.request(
            method, url, json=json, params=params, timeout=60,
            headers={"Authorization": f"Bearer {get_token()}"},
        )
        if resp.status_code in (429, 502, 503, 504) and attempt < 4:
            time.sleep(min(float(resp.headers.get("Retry-After", 2 ** attempt)), 30))
            continue
        break
    if resp.status_code == 423:
        raise PermissionError(
            "The workbook is locked (someone may have it open in a way that "
            "blocks edits). Try again in a moment — your data has not been lost.")
    if not resp.ok:
        try:
            err = resp.json().get("error", {})
        except ValueError:
            err = {}
        raise GraphError(
            f"Graph {method} {path.split('?')[0]} failed ({resp.status_code}): "
            f"{err.get('message') or resp.text[:200]}",
            resp.status_code, err.get("code", ""))
    return resp.json() if resp.content else {}


# ── Workbook ──────────────────────────────────────────────────────────────────

def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _cell_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


_item_cache: dict[str, str] = {}
_item_lock = threading.Lock()


def _resolve_item(url: str) -> str:
    """Return the Graph path prefix (/drives/{d}/items/{i}) for a sharing URL."""
    with _item_lock:
        if url in _item_cache:
            return _item_cache[url]
    encoded = "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    item = _request("GET", f"/shares/{encoded}/driveItem",
                    params={"$select": "id,name,parentReference"})
    prefix = f"/drives/{item['parentReference']['driveId']}/items/{item['id']}"
    with _item_lock:
        _item_cache[url] = prefix
    return prefix


class GraphSheet:
    """One worksheet of a workbook in OneDrive/SharePoint."""

    def __init__(self, url: str, sheet_name: str = ""):
        self.url = url.strip()
        self.sheet_name = sheet_name.strip()
        self._prefix: str | None = None
        self._ws: str | None = None

    # -- paths ---------------------------------------------------------------

    def _wb(self) -> str:
        if self._prefix is None:
            self._prefix = _resolve_item(self.url) + "/workbook"
        return self._prefix

    def _sheet_path(self, create: bool = False) -> str:
        if self._ws is None:
            sheets = _request("GET", self._wb() + "/worksheets",
                              params={"$select": "name"})["value"]
            names = [s["name"] for s in sheets]
            if not self.sheet_name:
                self.sheet_name = names[0]
            elif self.sheet_name not in names:
                if not create:
                    self.sheet_name = names[0]
                else:
                    _request("POST", self._wb() + "/worksheets/add",
                             json={"name": self.sheet_name})
            self._ws = f"{self._wb()}/worksheets('{quote(self.sheet_name.replace(chr(39), chr(39) * 2))}')"
        return self._ws

    def _range(self, address: str) -> str:
        return f"{self._sheet_path()}/range(address='{address}')"

    # -- reading -------------------------------------------------------------

    def read_values(self) -> list[list]:
        """All cell values from A1 to the last used cell (formatting-only
        cells are ignored)."""
        ws = self._sheet_path(create=False)
        used = _request("GET", f"{ws}/usedRange(valuesOnly=true)",
                        params={"$select": "address,values"})
        m = re.search(r"!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", used["address"])
        start_col, start_row = m.group(1), int(m.group(2))
        end_col = m.group(3) or start_col
        end_row = int(m.group(4) or start_row)
        if start_col == "A" and start_row == 1:
            return used["values"]
        # Sheet doesn't start at A1: re-read anchored at A1 so row numbers match
        rng = _request("GET", self._range(f"A1:{end_col}{end_row}"),
                       params={"$select": "values"})
        return rng["values"]

    def headers(self) -> list[str]:
        vals = self.read_values()
        if not vals or not any(str(c).strip() for c in vals[0]):
            return []
        hdr = [_cell_str(c).strip() for c in vals[0]]
        while hdr and not hdr[-1]:
            hdr.pop()
        return hdr

    # -- writing -------------------------------------------------------------

    def write_rows(self, first_row: int, rows: list[list], text_format: bool = False) -> None:
        if not rows:
            return
        width = max(len(r) for r in rows)
        rows = [list(r) + [""] * (width - len(r)) for r in rows]
        body = {"values": rows}
        if text_format:
            body["numberFormat"] = [["@"] * width for _ in rows]
        last = first_row + len(rows) - 1
        _request("PATCH", self._range(f"A{first_row}:{_col_letter(width)}{last}"), json=body)

    def delete_rows(self, row_numbers: list[int]) -> None:
        for r in sorted(set(row_numbers), reverse=True):     # bottom-up
            _request("POST", self._range(f"{r}:{r}") + "/delete", json={"shift": "Up"})

    def ensure_exists(self) -> None:
        """Create the worksheet if it is missing (used for the shared cache)."""
        self._sheet_path(create=True)

    # -- helpers used by xlsx_export ------------------------------------------

    def row_signatures(self, ncols: int) -> set:
        sigs = set()
        for row in self.read_values()[1:]:
            vals = tuple(_cell_str(row[c] if c < len(row) else "").strip()
                         for c in range(ncols))
            if any(vals):
                sigs.add(vals)
        return sigs

    def append_row(self, row: list) -> None:
        last_row = len(self.read_values())
        self.write_rows(last_row + 1, [row])

    # -- keyed record store (shared cache) -------------------------------------

    def read_records(self, key: str) -> tuple[list[str], list[dict]]:
        vals = self.read_values()
        if not vals or not any(str(c).strip() for c in vals[0]):
            return [], []
        hdr = [_cell_str(c).strip() for c in vals[0]]
        out = []
        for row in vals[1:]:
            rec = {h: _cell_str(row[i]) if i < len(row) else ""
                   for i, h in enumerate(hdr) if h}
            if rec.get(key):
                out.append(rec)
        return hdr, out

    def sync_records(self, columns: list[str], key: str,
                     upserts: list[dict], deletes: set[str]) -> None:
        """Write *upserts* (matched on *key*) and remove *deletes*, touching
        only those rows so other users' concurrent edits survive."""
        self.ensure_exists()
        vals = self.read_values()
        if not vals or not any(str(c).strip() for c in vals[0]):
            self.write_rows(1, [columns], text_format=True)
            vals = [columns]
        hdr = [_cell_str(c).strip() for c in vals[0]]
        for c in columns:                       # add any missing columns on the right
            if c not in hdr:
                hdr.append(c)
                self.write_rows(1, [hdr], text_format=True)
        kcol = hdr.index(key)
        row_of = {}
        for i, row in enumerate(vals[1:], start=2):
            k = _cell_str(row[kcol]) if kcol < len(row) else ""
            if k:
                row_of[k] = i
        next_row = len(vals) + 1

        appended, errors = [], []
        for rec in upserts:
            line = [str(rec.get(h, "") or "") for h in hdr]
            too_big = [h for h, v in zip(hdr, line) if len(v) > TEXT_CELL_LIMIT]
            if too_big:
                errors.append(f"Request {rec.get(key)}: '{too_big[0]}' is too large "
                              "for an Excel cell and was not pushed.")
                continue
            k = str(rec.get(key))
            if k in row_of:
                self.write_rows(row_of[k], [line], text_format=True)
            else:
                appended.append(line)
        if appended:
            self.write_rows(next_row, appended, text_format=True)
        self.delete_rows([row_of[k] for k in deletes if k in row_of])
        if errors:
            raise ValueError("; ".join(errors))
