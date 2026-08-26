#!/usr/bin/env python3
"""
Example: Using charge validation to ensure minimum charges and prevent over-charging.

This demonstrates the new validation features added to ILabClient:
  - validate_min_charge(): Ensures a request has at least $200 in charges
  - can_add_charges(): Prevents adding charges that would exceed the limit
  - add_charges(): Now includes max_charge validation
"""

from ilabs_client import ILabClient, ILabError
from prefs import get_prefs

def demo_charge_validation():
    """Demonstrate charge validation workflow."""
    prefs = get_prefs()
    max_charge = float(prefs.get("max_charge", "200"))

    client = ILabClient()
    core_id = 5226  # CALM (update with your core ID)
    request_id = 123456  # Update with a real request ID

    print(f"Max charge limit: ${max_charge:.2f}\n")

    # Get current total
    try:
        total = client.get_total_charges(core_id, request_id)
        print(f"Current charges: ${total:.2f}")
    except Exception as e:
        print(f"Error getting charges: {e}")
        return

    # Validate minimum charge requirement
    print(f"\nValidating minimum charge of ${max_charge:.2f}...")
    try:
        client.validate_min_charge(core_id, request_id, min_charge=max_charge)
        print("✓ Request has sufficient charges")
    except ILabError as e:
        print(f"✗ {e}")
        return

    # Example: Try to add charges that would exceed limit
    print(f"\nAttempting to add charges that exceed ${max_charge:.2f}...")
    new_charges = [
        {
            "quantity": 1,
            "price_id": 999,
            "service_id": 888,
            "note": "Test charge"
        }
    ]

    try:
        # This will check if adding these charges exceeds the max
        can_add = client.can_add_charges(core_id, request_id, new_charges, max_charge=max_charge)
        if can_add:
            print(f"✓ Safe to add charges (total would be within ${max_charge:.2f} limit)")
            # result = client.add_charges(core_id, request_id, new_charges, max_charge=max_charge)
            # print("✓ Charges added successfully")
    except ILabError as e:
        print(f"✗ Cannot add charges: {e}")

if __name__ == "__main__":
    demo_charge_validation()
