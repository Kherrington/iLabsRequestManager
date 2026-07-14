"""
Append a training record row to a Training Schedule xlsx file.

The workbook is expected to have a header row (row 1).  Known headers are
mapped to record fields; form-data columns are matched by key name
(case-insensitive).  If the file does not exist it is created with a
default header row.

Requires:  pip install openpyxl
"""

import json
from pathlib import Path

try:
    import openpyxl
    import openpyxl.styles
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# Maps normalised xlsx header text → internal record field key.
# Add entries here whenever a new column name is encountered.
HEADER_MAP: dict[str, str] = {
    # ── Date / time ───────────────────────────────────────────────────────────
    "date":               "training_date",
    "training date":      "training_date",
    "trainingdate":       "training_date",
    "trng date":          "training_date",
    "day":                "training_day",
    "training day":       "training_day",
    "trainingday":        "training_day",
    "time":               "training_time",
    "training time":      "training_time",
    "trainingtime":       "training_time",
    # ── Microscope / system ───────────────────────────────────────────────────
    "microscope":         "microscope",
    "system":             "microscope",
    "instrument":         "microscope",
    "scope":              "microscope",
    # ── Requester / trainee ───────────────────────────────────────────────────
    "people":             "owner_name",   # common header in CALM/CVRI schedules
    "person":             "owner_name",
    "trainee":            "owner_name",
    "trainee name":       "owner_name",
    "traineename":        "owner_name",
    "requester":          "owner_name",
    "name":               "owner_name",
    "first name":         "owner_name",
    "email":              "owner_email",
    "trainee email":      "owner_email",
    "traineeemail":       "owner_email",
    # ── PI / lab ──────────────────────────────────────────────────────────────
    "pi":                 "pi_name",
    "pi name":            "pi_name",
    "piname":             "pi_name",
    "lab":                "pi_name",
    "lab name":           "pi_name",
    "principal investigator": "pi_name",
    "pi email":           "pi_email",
    "piemail":            "pi_email",
    # ── Core ─────────────────────────────────────────────────────────────────
    "core":               "core_lab",
    "core lab":           "core_lab",
    "corelab":            "core_lab",
    "facility":           "core_lab",
    # ── Class taken ───────────────────────────────────────────────────────────
    "class":              "class_taken",
    "class taken":        "class_taken",
    "classtaken":         "class_taken",
    # ── Notes ─────────────────────────────────────────────────────────────────
    "notes":              "local_notes",
    "note":               "local_notes",
    "comments":           "local_notes",
    # ── Admin ─────────────────────────────────────────────────────────────────
    "request id":         "request_id",
    "requestid":          "request_id",
    "id":                 "request_id",
    "service":            "service_name",
    "request name":       "name",
    "requestname":        "name",
    "title":              "name",
}

# Column order used when creating a brand-new xlsx (or when row 1 is empty)
DEFAULT_HEADERS = [
    "Date", "Day", "Time", "Microscope",
    "People", "Email", "Lab", "PI Email",
    "Core", "Class Taken", "Notes", "Request ID",
]
DEFAULT_FIELDS = [
    "training_date", "training_day", "training_time", "microscope",
    "owner_name",    "owner_email",  "pi_name",       "pi_email",
    "core_lab",      "class_taken",  "local_notes",   "request_id",
]


def _norm(text: str) -> str:
    """Lowercase + collapse whitespace."""
    return " ".join(text.lower().split())


def _existing_row_signatures(ws, headers: list) -> set:
    """Return a set of signature tuples (stripped-string values, aligned to
    *headers*) for every existing data row, used to detect duplicates before
    appending a new row."""
    n = len(headers)
    sigs = set()
    for r in range(2, ws.max_row + 1):
        vals = tuple(str(ws.cell(r, c + 1).value or "").strip() for c in range(n))
        if any(vals):
            sigs.add(vals)
    return sigs


def _save_workbook(wb, xlsx_path: str) -> None:
    """Save *wb*, turning a Windows file-lock error into a message that makes
    the cause (file open in Excel) obvious to the caller/UI."""
    try:
        wb.save(xlsx_path)
    except PermissionError:
        raise PermissionError(
            f"'{Path(xlsx_path).name}' appears to be open in Excel (or another "
            "program) and can't be saved right now. Close the file and try "
            "again — your data has not been lost."
        ) from None


def append_training_row(rec: dict, xlsx_path: str,
                        sheet_name: str = "") -> None:
    """
    Append one training record row to the workbook at *xlsx_path*.

    sheet_name  Name of the worksheet to append to.  If blank or not found,
                the active (first) sheet is used.

    - If the file does not exist it is created with DEFAULT_HEADERS.
    - If row 1 is empty, DEFAULT_HEADERS are written first.
    - Otherwise the existing headers determine column placement.

    Field values applied:
        People / Requester  → owner_name  (requester name)
        Microscope          → microscope
        Date                → training_date
        Day                 → training_day
        Time                → training_time
        Notes               → local_notes
        Class / Class Taken → "Yes" if class_taken == "1", else ""

    Raises:
        ImportError   if openpyxl is not installed.
        Exception     propagated from openpyxl on file errors.
    """
    if not HAS_OPENPYXL:
        raise ImportError(
            "openpyxl is required for xlsx export.\n"
            "Install it with:  pip install openpyxl"
        )

    form_data: dict = json.loads(rec.get("form_data") or "{}")
    p = Path(xlsx_path)

    # ── Open or create workbook ───────────────────────────────────────────────
    if p.exists():
        wb = openpyxl.load_workbook(xlsx_path)
        # Select sheet by name, fall back to active
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active

        # Read headers from row 1, trimming phantom empty columns at the right
        # (formatted Excel files often have max_column >> actual data columns)
        raw_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        headers = [str(h).strip() if h is not None else "" for h in raw_headers]
        # Drop trailing empty headers
        while headers and not headers[-1]:
            headers.pop()

        if not headers:
            _write_headers(ws, DEFAULT_HEADERS)
            headers = DEFAULT_HEADERS
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name or "Training Schedule"
        _write_headers(ws, DEFAULT_HEADERS)
        headers = DEFAULT_HEADERS

    # ── Build row matching header order ───────────────────────────────────────
    row: list = []
    for header in headers:
        if not header:
            row.append("")
            continue
        n     = _norm(header)
        field = HEADER_MAP.get(n)
        if field:
            val = rec.get(field, "") or ""
            # class_taken: store "Yes" / "" rather than "1" / "0"
            if field == "class_taken":
                val = "Yes" if val == "1" else ""
            row.append(val)
        else:
            # Unknown column — try matching a form_data key
            fd_val = (
                form_data.get(header)
                or form_data.get(header.title())
                or next((v for k, v in form_data.items()
                         if k.lower() == header.lower()), "")
            )
            row.append(fd_val or "")

    # ── Skip if this row already exists ───────────────────────────────────────
    # Prefer matching on the Request ID column (unique per iLab request) when
    # the sheet has one; otherwise fall back to an exact full-row match.
    req_col = next((i for i, h in enumerate(headers)
                    if h and HEADER_MAP.get(_norm(h)) == "request_id"), None)
    req_id_val    = str(rec.get("request_id", "") or "").strip()
    existing_sigs = _existing_row_signatures(ws, headers)

    is_dup = False
    if req_col is not None and req_id_val:
        is_dup = any(sig[req_col] == req_id_val for sig in existing_sigs)
    if not is_dup:
        is_dup = tuple(str(v or "").strip() for v in row) in existing_sigs

    if is_dup:
        return {"headers": headers, "written": {}, "empty": [], "duplicate": True}

    ws.append(row)
    _save_workbook(wb, xlsx_path)

    # Return a summary for debugging / status messages
    mapped   = {h: v for h, v in zip(headers, row) if v not in (None, "")}
    unmapped = [h for h, v in zip(headers, row) if v in (None, "")]
    return {"headers": headers, "written": mapped, "empty": unmapped, "duplicate": False}


def _write_headers(ws, headers: list) -> None:
    bold = openpyxl.styles.Font(bold=True)
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c)
        cell.value = h
        cell.font  = bold


# ── Intro Course Log export ───────────────────────────────────────────────────

CLASS_HEADER_MAP: dict[str, str] = {
    "date":              "date",
    "class date":        "date",
    "instructor":        "instructor",
    "teacher":           "instructor",
    "taught by":         "instructor",
    "time":              "time",
    "class time":        "time",
    "location":          "location",
    "room":              "location",
    "where":             "location",
    "student":           "name",
    "student name":      "name",
    "name":              "name",
    "people":            "name",
    "pi":                "pi",
    "pi / lab":          "pi",
    "pi/lab":            "pi",
    "lab":               "pi",
    "pi name":           "pi",
    "principal investigator": "pi",
}

CLASS_DEFAULT_HEADERS = ["Date", "Instructor", "Time", "Location", "Student Name", "PI / Lab"]
CLASS_DEFAULT_FIELDS  = ["date", "instructor", "time", "location", "name", "pi"]


def append_class_session(session: dict, xlsx_path: str,
                         sheet_name: str = "") -> dict:
    """
    Append one row per student to the Microscope Intro Course Log xlsx.

    session keys: date, instructor, time, location,
                  students (list of {"name": str, "pi": str})

    Returns {"rows_written": int}.
    """
    if not HAS_OPENPYXL:
        raise ImportError(
            "openpyxl is required for xlsx export.\n"
            "Install it with:  pip install openpyxl"
        )

    p = Path(xlsx_path)

    if p.exists():
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active
        raw_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        headers = [str(h).strip() if h is not None else "" for h in raw_headers]
        while headers and not headers[-1]:
            headers.pop()
        if not headers:
            _write_headers(ws, CLASS_DEFAULT_HEADERS)
            headers = CLASS_DEFAULT_HEADERS
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name or "Intro Course Log"
        _write_headers(ws, CLASS_DEFAULT_HEADERS)
        headers = CLASS_DEFAULT_HEADERS

    students = session.get("students") or []
    existing_sigs = _existing_row_signatures(ws, headers)
    rows_written = 0
    duplicates_skipped = 0

    for student in students:
        flat = {
            "date":       session.get("date", ""),
            "instructor": session.get("instructor", ""),
            "time":       session.get("time", ""),
            "location":   session.get("location", ""),
            "name":       student.get("name", ""),
            "pi":         student.get("pi", ""),
        }
        row = []
        for header in headers:
            if not header:
                row.append("")
                continue
            field = CLASS_HEADER_MAP.get(_norm(header))
            row.append(flat.get(field, "") if field else "")

        row_sig = tuple(str(v or "").strip() for v in row)
        if row_sig in existing_sigs:
            duplicates_skipped += 1
            continue

        ws.append(row)
        existing_sigs.add(row_sig)
        rows_written += 1

    if rows_written:
        _save_workbook(wb, xlsx_path)
    return {"rows_written": rows_written, "duplicates_skipped": duplicates_skipped}
