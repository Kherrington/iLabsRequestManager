#!/usr/bin/env python3
"""
Bulk-add charges to an iLab service request from a CSV file.

Required CSV columns : quantity, price_id, service_id
Optional CSV column  : note

Usage:
    python add_charges_from_csv.py --core-id 1234 --request-id 56789 --csv charges.csv
    python add_charges_from_csv.py --core-id 1234 --request-id 56789 --csv charges.csv --dry-run

Tips:
  - Run get_cores.py to find your core ID.
  - price_id and service_id are the iLab internal IDs for the service price tier.
    Run get_service_requests.py and inspect an existing request, or pull from
    the iLab web interface URL.
  - Use --dry-run to preview the payload before submitting.
"""

import argparse
import csv
import sys

from ilabs_client import ILabClient


REQUIRED_COLUMNS = {"quantity", "price_id", "service_id"}


def load_charges(csv_path: str) -> list:
    with open(csv_path, newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        fieldnames = set(reader.fieldnames or [])
        missing = REQUIRED_COLUMNS - fieldnames
        if missing:
            print(f"Error: CSV is missing required column(s): {sorted(missing)}", file=sys.stderr)
            sys.exit(1)

        charges = []
        for i, row in enumerate(reader, start=2):
            try:
                charge = {
                    "quantity": float(row["quantity"]),
                    "price_id": int(row["price_id"]),
                    "service_id": int(row["service_id"]),
                }
                if row.get("note"):
                    charge["note"] = row["note"].strip()
                charges.append(charge)
            except ValueError as exc:
                print(f"Error on CSV row {i}: {exc}", file=sys.stderr)
                sys.exit(1)

    return charges


def main():
    parser = argparse.ArgumentParser(
        description="Bulk-add charges to an iLab service request from a CSV file"
    )
    parser.add_argument("--core-id", type=int, required=True, help="iLab core ID")
    parser.add_argument("--request-id", type=int, required=True, help="iLab service request ID")
    parser.add_argument("--csv", required=True, metavar="FILE", help="CSV file with charges to add")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be submitted without making any API calls",
    )
    args = parser.parse_args()

    charges = load_charges(args.csv)
    print(f"Loaded {len(charges)} charge(s) from {args.csv}.")

    if args.dry_run:
        print("\nDry run — charges that would be submitted:")
        header = f"  {'qty':>8}  {'service_id':>12}  {'price_id':>10}  note"
        print(header)
        print("  " + "-" * (len(header) - 2))
        for c in charges:
            note = c.get("note", "")
            print(f"  {c['quantity']:>8}  {c['service_id']:>12}  {c['price_id']:>10}  {note}")
        return

    client = ILabClient()
    print(f"Submitting to core {args.core_id}, request {args.request_id}...")
    client.add_charges(args.core_id, args.request_id, charges)
    print(f"Done. {len(charges)} charge(s) added.")


if __name__ == "__main__":
    main()
