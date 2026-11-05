"""
DocuMergePDF - Developer Utility to Reset Trial Runs
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

if __name__ == "__main__":
    sec = LicenseSecurity()
    ok = sec.reset_trial()
    if ok:
        print(f"[✓] Successfully reset trial runs for HWID: {sec.hwid}")
        st = sec.get_status()
        print(f"Status: {st['runs_remaining']}/{st['max_trial_runs']} trial sessions remaining.")
    else:
        print("[!] Failed to reset trial state.")

