#!/usr/bin/env python3
"""
List all iLab cores accessible with your API token.

Usage:
    python get_cores.py

Prints core IDs and names. You'll need a core ID to use the other scripts.
"""

from ilabs_client import ILabClient


def main():
    client = ILabClient()
    cores = client.list_cores()

    if not cores:
        print("No cores found for this token.")
        return

    print(f"Found {len(cores)} core(s):\n")
    for core in cores:
        print(f"  ID   : {core.get('id')}")
        print(f"  Name : {core.get('name')}")
        homepage = core.get("homepage") or ""
        if homepage:
            print(f"  URL  : {homepage}")
        print()


if __name__ == "__main__":
    main()
