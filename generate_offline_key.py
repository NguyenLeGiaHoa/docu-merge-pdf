"""
DocuMergePDF - Seller Utility: Offline Studio Key Generator
Generates air-gapped cryptographic lifetime or monthly license keys for enterprise and offline customers.

Usage:
    python generate_offline_key.py <CUSTOMER_HWID> [lifetime|monthly]

Example:
    python generate_offline_key.py 4A2B8C1D9E0F2A3B4C5D6E7F8091A2B3 lifetime
"""

import sys
import os

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from core.licensing import LicenseSecurity


def main():
    sec = LicenseSecurity()

    if len(sys.argv) < 2:
        print("=" * 70)
        print("  DocuMergePDF — Offline Studio Key Generator (Seller Tool)")
        print("=" * 70)
        print("\nUsage:")
        print("  python generate_offline_key.py <CUSTOMER_HWID> [lifetime|monthly]\n")
        print(f"Current Machine HWID: {sec.hwid}")
        sample_key = LicenseSecurity.generate_offline_studio_key(sec.hwid, "lifetime")
        print(f"Sample Lifetime Key for this machine: {sample_key}\n")
        sys.exit(0)

    target_hwid = sys.argv[1].strip()
    tier = sys.argv[2].strip().lower() if len(sys.argv) > 2 else "lifetime"

    if len(target_hwid) < 8:
        print(f"[!] Error: HWID '{target_hwid}' appears too short (must be at least 8 hex characters).")
        sys.exit(1)

    offline_key = LicenseSecurity.generate_offline_studio_key(target_hwid, tier=tier)

    print("=" * 70)
    print(f"  DOCUMERGEPDF — OFFLINE STUDIO {tier.upper()} LICENSE GENERATED")
    print("=" * 70)
    print(f"Customer Hardware ID : {target_hwid.upper()}")
    print(f"License Tier         : {tier.capitalize()}")
    print(f"Generated Studio Key : {offline_key}")
    print("=" * 70)
    print("\nEmail Template to Send to Customer:")
    print("-" * 70)
    print("Hi there,")
    print("\nThank you for purchasing DocuMergePDF!")
    print(f"Here is your permanent Offline Studio License Key ({tier.capitalize()}) for your air-gapped PC:\n")
    print(f"  License Key: {offline_key}\n")
    print("Instructions to activate:")
    print("1. Open DocuMergePDF on your computer.")
    print("2. Switch to the 'Settings & License' tab.")
    print(f"3. Paste the key ({offline_key}) into the license box and click 'Activate License'.")
    print("4. Your Commercial License will activate instantly with zero internet required.")
    print("\nBest regards,")
    print("DocuMergePDF Customer Support Team")
    print("-" * 70)


if __name__ == "__main__":
    main()

