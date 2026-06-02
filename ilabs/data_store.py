"""
Local CSV cache for iLab service requests.

Fixed columns come from iLab; local columns (assigned_to, labels, local_notes)
are never overwritten during sync so edits are preserved.

form_data and milestones_data are stored as JSON strings in the CSV.
export_expanded() flattens form_data fields into individual columns.
"""

import csv
import json
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional


def _detect_core(form_data: dict) -> str:
    """
    Detect core facility from form data.
    Checks 'Which Facility do you need Training for?' first, then falls back
    to a general CALM / CVRI text scan across all keys and values.
    """
    facility_keys = {
        "which facility do you need training for",
        "which facility do you need training for?",
        "facility",
        "which facility",
        "core facility",
        "core",
    }
    def _is_calm(v: str) -> bool:
        return "CALM" in v or "NIC" in v   # NIC = old name for CALM

    # Priority: check specific field whose answer is CALM/NIC or CVRI
    for key, val in form_data.items():
        if key.lower().strip().rstrip("?") in {k.rstrip("?") for k in facility_keys}:
            v = str(val).upper()
            if "CVRI" in v:
                return "CVRI"
            if _is_calm(v):
                return "CALM"
    # Fallback: full text scan
    text = (
        " ".join(str(v) for v in form_data.values()) + " " +
        " ".join(form_data.keys())
    ).upper()
    if "CVRI" in text:
        return "CVRI"
    if _is_calm(text):
        return "CALM"
    return ""


_FIXED_COLS = [
    "request_id", "name", "state", "created_at",
    "start_on", "end_on", "completed_on",
    "owner_name", "owner_email", "pi_name", "pi_email",
    "service_name", "last_synced",
]
_LOCAL_COLS = [
    "assigned_to", "labels", "local_notes",
    # Training scheduling fields (never overwritten by iLab sync)
    "core_lab", "microscope",
    "training_date", "training_time", "training_day",
    "class_taken",
    # Workflow progress checkboxes (pre- and post-training)
    "wf_emailed", "wf_class_scheduled", "wf_not_required",
    "wf_training_scheduled",
    "wf_post_email", "wf_post_listserve",
    "wf_post_approved", "wf_post_confirmed",
]
_BLOB_COLS  = ["form_data", "milestones_data"]

ALL_COLS = _FIXED_COLS + _LOCAL_COLS + _BLOB_COLS


class DataStore:
    def __init__(self, filepath: str = "requests_cache.csv"):
        self.filepath = Path(filepath)
        self.records: Dict[str, dict] = {}
        self._load()

    # ── Persistence ───────────────────────────────────────────────────────────

    def _load(self) -> None:
        if not self.filepath.exists():
            return
        with open(self.filepath, newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                self.records[row["request_id"]] = row

    def save(self) -> None:
        with open(self.filepath, "w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=ALL_COLS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(self.records.values())

    # ── Sync from iLab ────────────────────────────────────────────────────────

    def sync(
        self,
        core_id: int,
        client,
        on_progress: Optional[Callable[[str], None]] = None,
        fetch_forms: bool = True,
        states: Optional[list] = None,
    ) -> int:
        """
        Pull service requests from iLab (optionally filtered to active states),
        merging into the local cache while preserving local fields.
        Returns the number of requests fetched.
        """
        def _progress(msg: str) -> None:
            if on_progress:
                on_progress(msg)

        params = {}
        if states:
            params["states"] = ",".join(states)

        _progress("Fetching active service requests…")
        requests_list = client.list_service_requests(core_id, **params)
        total = len(requests_list)
        _progress(f"Found {total} request(s). Fetching details…")

        for i, req in enumerate(requests_list, start=1):
            req_id = str(req.get("id"))
            existing = self.records.get(req_id, {})

            # ── Form data ─────────────────────────────────────────────────────
            form_data: dict = {}
            if fetch_forms:
                _progress(f"[{i}/{total}] Fetching form data for request {req_id}…")
                try:
                    forms = client.list_custom_forms(core_id, int(req_id))
                    for form in forms:
                        for field in (form.get("fields") or []):
                            fname = (field.get("name") or "").strip()
                            fval  = field.get("value")
                            if fname and fval not in (None, "", [], {}):
                                # Flatten list values to a readable string
                                if isinstance(fval, list):
                                    fval = ", ".join(str(v) for v in fval)
                                form_data[fname] = str(fval)
                except Exception:
                    # Keep previously cached form data on failure
                    form_data = json.loads(existing.get("form_data") or "{}")
            else:
                form_data = json.loads(existing.get("form_data") or "{}")

            owner   = req.get("owner") or {}
            lab     = req.get("lab")   or {}
            pi_list = lab.get("principal_investigators") or []
            pi      = pi_list[0] if pi_list else {}

            # ── PI / Lab name: API lab name first, then PI name, then form data
            pi_name_val = lab.get("name", "") or pi.get("name", "")
            if not pi_name_val:
                pi_keys = {"pi name", "pi", "principal investigator",
                           "lab pi", "pi/lab", "lab/pi", "lab name"}
                for fk, fv in form_data.items():
                    if fk.lower().strip() in pi_keys:
                        pi_name_val = str(fv)
                        break

            # ── Core: check "Which Facility" form field ───────────────────────
            existing_core = existing.get("core_lab", "")
            if not existing_core:
                existing_core = _detect_core(form_data)

            existing_assigned = existing.get("assigned_to", "")

            self.records[req_id] = {
                # ── iLab fields (always refreshed from API) ───────────────────
                "request_id":   req_id,
                "name":         req.get("name", ""),
                "state":        req.get("state", ""),
                # submitted_at is the correct field; fall back to created_at for
                # any older cached records that used the wrong key
                "created_at":   req.get("submitted_at") or req.get("created_at") or "",
                "start_on":     req.get("start_on") or "",
                "end_on":       req.get("end_on") or "",
                "completed_on": req.get("completed_on") or "",
                "owner_name":   owner.get("name", ""),
                "owner_email":  owner.get("email", ""),
                "pi_name":      pi_name_val,
                "pi_email":     pi.get("email", ""),
                "service_name": req.get("service_name", ""),
                "last_synced":  datetime.now().isoformat(timespec="seconds"),
                # ── Local fields (never overwritten after first fill) ──────────
                "assigned_to":    existing_assigned,
                "labels":         existing.get("labels", ""),
                "local_notes":    existing.get("local_notes", ""),
                # ── Training fields (always preserved) ────────────────────────
                "core_lab":       existing_core,
                "microscope":     existing.get("microscope", ""),
                "training_date":  existing.get("training_date", ""),
                "training_time":  existing.get("training_time", ""),
                "training_day":   existing.get("training_day", ""),
                "class_taken":    existing.get("class_taken", "0"),
                # ── Workflow progress (always preserved) ──────────────────────
                "wf_emailed":           existing.get("wf_emailed", "0"),
                "wf_class_scheduled":   existing.get("wf_class_scheduled", "0"),
                "wf_not_required":      existing.get("wf_not_required", "0"),
                "wf_training_scheduled":existing.get("wf_training_scheduled", "0"),
                "wf_post_email":        existing.get("wf_post_email", "0"),
                "wf_post_listserve":    existing.get("wf_post_listserve", "0"),
                "wf_post_approved":     existing.get("wf_post_approved", "0"),
                "wf_post_confirmed":    existing.get("wf_post_confirmed", "0"),
                # ── Blobs ─────────────────────────────────────────────────────
                "form_data":       json.dumps(form_data, ensure_ascii=False),
                "milestones_data": existing.get("milestones_data", "[]"),
            }

        # ── Remove records no longer active in iLab ───────────────────────────
        fetched_ids = {str(req.get("id")) for req in requests_list}
        stale = [rid for rid in list(self.records) if rid not in fetched_ids]
        for rid in stale:
            del self.records[rid]
        if stale:
            _progress(f"Removed {len(stale)} inactive request(s) from cache.")

        self.save()
        _progress(f"Sync complete — {total} active request(s), "
                  f"{len(stale)} removed.")
        return total

    # ── CRUD helpers ──────────────────────────────────────────────────────────

    def update_local_fields(self, request_id: str, **fields) -> None:
        """Update only local fields (assigned_to, labels, local_notes)."""
        rec = self.records.get(str(request_id))
        if rec is None:
            return
        for key, value in fields.items():
            if key in _LOCAL_COLS:
                rec[key] = value
        self.save()

    def update_field(self, request_id: str, key: str, value: str) -> None:
        """Update any single field (use for state changes pushed to iLab)."""
        rec = self.records.get(str(request_id))
        if rec is not None:
            rec[key] = value
            self.save()

    def update_milestones(self, request_id: str, milestones: list) -> None:
        rec = self.records.get(str(request_id))
        if rec is not None:
            rec["milestones_data"] = json.dumps(milestones, ensure_ascii=False)
            self.save()

    def clear_all(self) -> None:
        """Remove every cached record and overwrite the CSV with an empty store."""
        self.records.clear()
        self.save()

    def all_records(self) -> List[dict]:
        return list(self.records.values())

    def get_record(self, request_id) -> Optional[dict]:
        return self.records.get(str(request_id))

    # ── Manual import ────────────────────────────────────────────────────────

    def import_from_ilab_export(self, filepath: str) -> dict:
        """
        Import a CSV exported from the iLab web UI ("View All Requests → Export").

        Tries multiple encodings and auto-detects the delimiter.  Columns that
        don't map to a known internal field are stored as form_data entries so
        they still appear in the Form Data tab.

        Returns a diagnostic dict::

            imported        int   – rows successfully imported
            skipped         int   – rows that had no recognisable request ID
            encoding        str   – encoding that successfully opened the file
            columns_raw     list  – every header found in the file
            columns_mapped  dict  – header → internal field (recognised columns)
            columns_form    list  – headers stored as form_data
            columns_none    list  – headers with no mapping at all (blank etc.)
        """
        # ── Column-name map ───────────────────────────────────────────────────
        # Normalised (lowercase, spaces, no underscores) header → internal key.
        # Add more entries here whenever a new iLab column variant is found.
        COLUMN_MAP = {
            # ── Request ID ─────────────────────────────────────────────────
            "id":                       "request_id",
            "request id":               "request_id",
            "request_id":               "request_id",
            "request #":                "request_id",
            "request#":                 "request_id",
            "request no":               "request_id",
            "request no.":              "request_id",
            "request number":           "request_id",
            "req id":                   "request_id",
            "req #":                    "request_id",
            "req no":                   "request_id",
            "no":                       "request_id",
            "no.":                      "request_id",
            "#":                        "request_id",
            # ── Name / title ───────────────────────────────────────────────
            "name":                     "name",
            "request name":             "name",
            "project name":             "name",
            "project":                  "name",
            "title":                    "name",
            "description":              "name",
            # ── State ──────────────────────────────────────────────────────
            "status":                   "state",
            "state":                    "state",
            "request status":           "state",
            "request state":            "state",
            # ── Created / submitted date ───────────────────────────────────
            "created at":               "created_at",
            "created_at":               "created_at",
            "created":                  "created_at",
            "submitted":                "created_at",
            "submitted at":             "created_at",
            "submitted on":             "created_at",
            "submission date":          "created_at",
            "date submitted":           "created_at",
            "date created":             "created_at",
            "request date":             "created_at",
            "open date":                "created_at",
            # ── Start / end ────────────────────────────────────────────────
            "start":                    "start_on",
            "start on":                 "start_on",
            "start date":               "start_on",
            "scheduled start":          "start_on",
            "end":                      "end_on",
            "end on":                   "end_on",
            "end date":                 "end_on",
            "scheduled end":            "end_on",
            "due date":                 "end_on",
            "completed on":             "completed_on",
            "completed_on":             "completed_on",
            "completed":                "completed_on",
            "completion date":          "completed_on",
            "date completed":           "completed_on",
            # ── Owner / requester ──────────────────────────────────────────
            "owner":                    "owner_name",
            "owner name":               "owner_name",
            "requester":                "owner_name",
            "requester name":           "owner_name",
            "user":                     "owner_name",
            "user name":                "owner_name",
            "username":                 "owner_name",
            "submitted by":             "owner_name",
            "lab member":               "owner_name",
            "member":                   "owner_name",
            "researcher":               "owner_name",
            # ── Owner email ────────────────────────────────────────────────
            "owner email":              "owner_email",
            "requester email":          "owner_email",
            "user email":               "owner_email",
            "email":                    "owner_email",
            "e-mail":                   "owner_email",
            # ── PI ─────────────────────────────────────────────────────────
            "pi":                       "pi_name",
            "pi name":                  "pi_name",
            "pi / lab":                 "pi_name",
            "lab / pi":                 "pi_name",
            "lab/pi":                   "pi_name",
            "pi/lab":                   "pi_name",
            "principal investigator":   "pi_name",
            "lab":                      "pi_name",
            "lab name":                 "pi_name",
            "group":                    "pi_name",
            "group name":               "pi_name",
            # ── PI email ───────────────────────────────────────────────────
            "pi email":                 "pi_email",
            "lab email":                "pi_email",
            # ── Service ────────────────────────────────────────────────────
            "service":                  "service_name",
            "service name":             "service_name",
            "core service":             "service_name",
            "service type":             "service_name",
            "type":                     "service_name",
        }

        def _norm(h: str) -> str:
            """Lowercase, collapse whitespace, drop underscores."""
            import re
            return re.sub(r"[\s_]+", " ", h.strip().lower())

        # ── Try encodings in order ────────────────────────────────────────────
        _ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
        fh = None
        used_enc = "utf-8-sig"
        for enc in _ENCODINGS:
            try:
                candidate = open(filepath, newline="", encoding=enc)
                candidate.read(2048)          # probe — will raise on bad encoding
                candidate.seek(0)
                fh = candidate
                used_enc = enc
                break
            except (UnicodeDecodeError, LookupError):
                try:
                    candidate.close()
                except Exception:
                    pass

        if fh is None:
            raise ValueError(
                "Could not open the file with any supported encoding "
                "(tried UTF-8, CP1252, Latin-1).  "
                "Try re-saving the CSV from Excel as UTF-8."
            )

        result = {
            "imported":       0,
            "skipped":        0,
            "encoding":       used_enc,
            "columns_raw":    [],
            "columns_mapped": {},
            "columns_form":   [],
            "columns_none":   [],
        }

        try:
            # ── Sniff delimiter ───────────────────────────────────────────────
            sample = fh.read(4096)
            fh.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            except csv.Error:
                dialect = csv.excel          # fall back to standard comma CSV

            reader = csv.DictReader(fh, dialect=dialect)
            if not reader.fieldnames:
                return result

            result["columns_raw"] = list(reader.fieldnames)

            # ── Build column map ──────────────────────────────────────────────
            col_map: dict[str, str | None] = {}
            for h in reader.fieldnames:
                internal = COLUMN_MAP.get(_norm(h))
                col_map[h] = internal
                if not h.strip():
                    result["columns_none"].append(h)
                elif internal:
                    result["columns_mapped"][h] = internal
                else:
                    result["columns_form"].append(h)

            # ── Import rows ───────────────────────────────────────────────────
            for row in reader:
                # Find request ID
                req_id = None
                for h, internal in col_map.items():
                    if internal == "request_id":
                        v = row.get(h, "").strip()
                        if v:
                            req_id = v
                            break

                if not req_id:
                    result["skipped"] += 1
                    continue

                existing  = self.records.get(req_id, {})
                form_data = json.loads(existing.get("form_data") or "{}")

                new_rec = {
                    # iLab fields — start from existing, will be overwritten below
                    "request_id":      req_id,
                    "name":            existing.get("name", ""),
                    "state":           existing.get("state", ""),
                    "created_at":      existing.get("created_at", ""),
                    "start_on":        existing.get("start_on", ""),
                    "end_on":          existing.get("end_on", ""),
                    "completed_on":    existing.get("completed_on", ""),
                    "owner_name":      existing.get("owner_name", ""),
                    "owner_email":     existing.get("owner_email", ""),
                    "pi_name":         existing.get("pi_name", ""),
                    "pi_email":        existing.get("pi_email", ""),
                    "service_name":    existing.get("service_name", ""),
                    "last_synced":     existing.get("last_synced", "imported"),
                    # Local fields — always preserved
                    "assigned_to":     existing.get("assigned_to", ""),
                    "labels":          existing.get("labels", ""),
                    "local_notes":     existing.get("local_notes", ""),
                    "milestones_data": existing.get("milestones_data", "[]"),
                    # Training fields — always preserved
                    "core_lab":        existing.get("core_lab", ""),
                    "microscope":      existing.get("microscope", ""),
                    "training_date":   existing.get("training_date", ""),
                    "training_time":   existing.get("training_time", ""),
                    "training_day":    existing.get("training_day", ""),
                    "class_taken":     existing.get("class_taken", "0"),
                }

                for h, internal in col_map.items():
                    val = row.get(h, "").strip()
                    if not val:
                        continue
                    if internal and internal != "request_id":
                        # Only fill blanks — never overwrite data from a live sync
                        if not new_rec.get(internal):
                            new_rec[internal] = val
                    elif internal is None and h.strip():
                        # Unknown column → form_data
                        form_data[h.strip()] = val

                new_rec["form_data"] = json.dumps(form_data, ensure_ascii=False)
                self.records[req_id] = new_rec
                result["imported"] += 1

        finally:
            fh.close()

        if result["imported"] > 0:
            self.save()

        return result

    # ── Export ────────────────────────────────────────────────────────────────

    def export_expanded(self, filepath: str) -> None:
        """
        Write a CSV where each form_data key becomes its own column.
        milestones_data is dropped from the export.
        """
        records = list(self.records.values())
        if not records:
            raise ValueError("No records to export.")

        # Collect all unique form field names in encounter order
        seen: set = set()
        form_fields: List[str] = []
        for rec in records:
            for key in json.loads(rec.get("form_data") or "{}"):
                if key not in seen:
                    form_fields.append(key)
                    seen.add(key)

        base_cols = _FIXED_COLS + _LOCAL_COLS   # everything except blobs
        all_cols  = base_cols + form_fields

        with open(filepath, "w", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=all_cols, extrasaction="ignore")
            writer.writeheader()
            for rec in records:
                row = {k: rec.get(k, "") for k in base_cols}
                form_data = json.loads(rec.get("form_data") or "{}")
                for fname in form_fields:
                    row[fname] = form_data.get(fname, "")
                writer.writerow(row)
