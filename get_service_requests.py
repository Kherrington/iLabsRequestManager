#!/usr/bin/env python3
"""
Export iLab service requests to CSV.

Usage:
    python get_service_requests.py --core-id 1234
    python get_service_requests.py --core-id 1234 --states processing,completed
    python get_service_requests.py --core-id 1234 --from-date 2024-01-01 --out requests.csv
    python get_service_requests.py --core-id 1234 --query "flow cytometry"

Output columns:
    id, name, state, created_at, start_on, end_on, completed_on,
    projected_cost, service_name, owner_name, owner_email, pi_name, pi_email
"""

import argparse
import csv
import sys

from ilabs_client import ILabClient


FIELDS = [
    "id",
    "name",
    "state",
    "created_at",
    "start_on",
    "end_on",
    "completed_on",
    "projected_cost",
    "service_name",
    "owner_name",
    "owner_email",
    "pi_name",
    "pi_email",
]


def flatten(req: dict) -> dict:
    owner = req.get("owner") or {}
    pi = req.get("principal_investigator") or {}
    return {
        "id": req.get("id"),
        "name": req.get("name"),
        "state": req.get("state"),
        "created_at": req.get("created_at"),
        "start_on": req.get("start_on"),
        "end_on": req.get("end_on"),
        "completed_on": req.get("completed_on"),
        "projected_cost": req.get("projected_cost"),
        "service_name": req.get("service_name"),
        "owner_name": owner.get("name"),
        "owner_email": owner.get("email"),
        "pi_name": pi.get("name"),
        "pi_email": pi.get("email"),
    }


def main():
    parser = argparse.ArgumentParser(description="Export iLab service requests to CSV")
    parser.add_argument("--core-id", type=int, required=True, help="iLab core ID (run get_cores.py to find it)")
    parser.add_argument(
        "--states",
        help=(
            "Comma-separated request states to include. "
            "Options: proposed, processing, completed, cancelled, draft, "
            "financials_approved, financials_rejected, needs_financial_reapproval, requested"
        ),
    )
    parser.add_argument("--from-date", metavar="YYYY-MM-DD", help="Only include requests on or after this date")
    parser.add_argument("--to-date", metavar="YYYY-MM-DD", help="Only include requests on or before this date")
    parser.add_argument("--query", "-q", metavar="TEXT", help="Full-text search across requests")
    parser.add_argument("--out", "-o", default="-", metavar="FILE", help="Output CSV file (default: stdout)")
    args = parser.parse_args()

    client = ILabClient()

    filters = {}
    if args.states:
        filters["states"] = args.states
    if args.from_date:
        filters["from_date"] = args.from_date
    if args.to_date:
        filters["to_date"] = args.to_date
    if args.query:
        filters["q"] = args.query

    print(f"Fetching service requests for core {args.core_id}...", file=sys.stderr)
    data = client.list_service_requests(args.core_id, **filters)
    print(f"Found {len(data)} request(s).", file=sys.stderr)

    if not data:
        return

    rows = [flatten(r) for r in data]

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
