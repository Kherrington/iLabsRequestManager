#!/usr/bin/env python3
"""
Export charges from iLab service requests to CSV.

Fetches all requests for a core, then retrieves charges for each.
Use --request-id to pull charges from a single request only.

Usage:
    python get_charges.py --core-id 1234
    python get_charges.py --core-id 1234 --request-id 56789
    python get_charges.py --core-id 1234 --states processing,completed --out charges.csv
    python get_charges.py --core-id 1234 --from-date 2024-01-01 --out charges.csv

Output columns:
    request_id, request_name, request_state,
    charge_id, charge_name, service_id, price_id,
    quantity, status, billing_status, note
"""

import argparse
import csv
import sys

from ilabs_client import ILabClient


FIELDS = [
    "request_id",
    "request_name",
    "request_state",
    "charge_id",
    "charge_name",
    "service_id",
    "price_id",
    "quantity",
    "status",
    "billing_status",
    "note",
]


def main():
    parser = argparse.ArgumentParser(description="Export iLab charges to CSV")
    parser.add_argument("--core-id", type=int, required=True, help="iLab core ID")
    parser.add_argument("--request-id", type=int, help="Pull charges from a single request only")
    parser.add_argument("--states", help="Filter requests by state (comma-separated)")
    parser.add_argument("--from-date", metavar="YYYY-MM-DD")
    parser.add_argument("--to-date", metavar="YYYY-MM-DD")
    parser.add_argument("--out", "-o", default="-", metavar="FILE")
    args = parser.parse_args()

    client = ILabClient()

    if args.request_id:
        req = client.get_service_request(args.core_id, args.request_id)
        requests_data = [req]
    else:
        filters = {}
        if args.states:
            filters["states"] = args.states
        if args.from_date:
            filters["from_date"] = args.from_date
        if args.to_date:
            filters["to_date"] = args.to_date
        print(f"Fetching requests for core {args.core_id}...", file=sys.stderr)
        requests_data = client.list_service_requests(args.core_id, **filters)
        print(f"Found {len(requests_data)} request(s).", file=sys.stderr)

    rows = []
    for req in requests_data:
        req_id = req.get("id")
        req_name = req.get("name")
        req_state = req.get("state")
        print(f"  Fetching charges for request {req_id} ({req_name})...", file=sys.stderr)
        charges = client.list_charges(args.core_id, req_id)
        for charge in charges:
            rows.append(
                {
                    "request_id": req_id,
                    "request_name": req_name,
                    "request_state": req_state,
                    "charge_id": charge.get("id"),
                    "charge_name": charge.get("name"),
                    "service_id": charge.get("service_id"),
                    "price_id": charge.get("price_id"),
                    "quantity": charge.get("quantity"),
                    "status": charge.get("status"),
                    "billing_status": charge.get("billing_status"),
                    "note": charge.get("note"),
                }
            )

    print(f"\nTotal charges: {len(rows)}", file=sys.stderr)

    if not rows:
        return

    if args.out == "-":
        out = sys.stdout
        close_out = False
    else:
        out = open(args.out, "w", newline="", encoding="utf-8")
        close_out = True

    try:
        writer = csv.DictWriter(out, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)
    finally:
        if close_out:
            out.close()

    if close_out:
        print(f"Saved {len(rows)} row(s) to {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
