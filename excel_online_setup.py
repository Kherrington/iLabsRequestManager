"""
Helper for the Excel Online (Microsoft Graph) integration.

    python excel_online_setup.py test    <workbook-url> [sheet]
        Sign in (browser opens the first time) and print the sheet's headers.

    python excel_online_setup.py migrate <old_cache.csv> <workbook-url> [sheet]
        One-off upload of an existing ilab_requests_cache.csv into the shared
        workbook's cache sheet (default sheet name: Cache).

    python excel_online_setup.py signout
"""

import sys

from graph_excel import GraphSheet, sign_out, get_token


def main(argv: list[str]) -> int:
    if not argv:
        print(__doc__)
        return 1
    cmd = argv[0]
    if cmd == "signout":
        sign_out()
        print("Signed out.")
    elif cmd == "test" and len(argv) >= 2:
        get_token()
        sheet = GraphSheet(argv[1], argv[2] if len(argv) > 2 else "")
        print("Sheet:", sheet.sheet_name or "(first)", "| headers:", sheet.headers())
    elif cmd == "migrate" and len(argv) >= 3:
        import data_store
        src = data_store.DataStore(argv[1])
        dst = data_store.DataStore(argv[2], remote_sheet=argv[3] if len(argv) > 3 else "Cache")
        dst.reload()
        for rec in src.all_records():
            dst.records[rec["request_id"]] = rec
        dst.save()
        dst.flush_remote(120)
        print(dst.remote_error or f"Uploaded {len(src.records)} record(s).")
    else:
        print(__doc__)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
