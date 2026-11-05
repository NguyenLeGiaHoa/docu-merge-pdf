"""
DocuMergePDF - Industrial Legal Immunity & Liability Shield Engine
Enforces UCC § 2-316 (Absolute As-Is Warranty Disclaimer) and UCC § 2-719 ($1.00 USD Liability Ceiling).
Assures 100% user indemnification, non-fiduciary status, and neutral local utility defense.
"""

import os
import json
import time
import hashlib
from typing import Tuple, Dict, Any

LEGAL_DISCLAIMER_SUMMARY = """LEGAL DISCLAIMER, TRADEMARKS & LIABILITY LIMITATION (UCC § 2-316 / § 2-719):
1. NEUTRAL LOCAL UTILITY: DocuMergePDF is strictly an offline document templating and mail-merge utility. The developer does NOT provide legal, educational accreditation, financial, tax, or employment advice.
2. ABSOLUTE "AS IS" DISCLAIMER: The software is provided "AS IS", "WITH ALL FAULTS", without warranty of merchantability, fitness for a particular purpose, or document rendering fidelity.
3. $1.00 USD LIABILITY CEILING: In no event shall developer liability exceed the lesser of fees paid in the past 12 months or $1.00 USD. All consequential, lost profits, student certification disputes, payroll errors, or data loss claims are waived.
4. 100% USER INDEMNIFICATION: User assumes full responsibility for template accuracy, employee/student privacy, data formatting, and backups, and agrees to indemnify and hold developer harmless.
5. TRADEMARK NOTICE: Microsoft, Word, Excel, Office, and Windows are registered trademarks of Microsoft Corporation. DocuMergePDF is an independent desktop utility and is NOT affiliated with, sponsored by, or endorsed by Microsoft Corporation. Compatibility is claimed under nominative fair use."""

LEGAL_VERSION = "2026.1-DOCUMERGE-UCC-SHIELD"


class LegalShield:
    """Manages legal consent, liability disclaimers, and first-run indemnification acceptance."""

    @staticmethod
    def get_terms_path() -> str:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        terms_file = os.path.join(base_dir, "TERMS_OF_SERVICE.txt")
        if os.path.exists(terms_file):
            return terms_file
        return os.path.join(base_dir, "LICENSE.txt")

    @staticmethod
    def read_full_terms() -> str:
        terms_path = LegalShield.get_terms_path()
        if os.path.exists(terms_path):
            try:
                with open(terms_path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception:
                pass
        return LEGAL_DISCLAIMER_SUMMARY

    @staticmethod
    def get_consent_file_path() -> str:
        candidates = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "DocuMergePDF", ".security"),
            os.path.join(os.path.expanduser("~"), ".documergepdf", ".security"),
            os.path.join(os.environ.get("APPDATA", ""), "DocuMergePDF", ".security"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".security")
        ]
        for secure_dir in candidates:
            if not secure_dir or secure_dir.strip() in ["", "\\DocuMergePDF\\.security"]:
                continue
            try:
                os.makedirs(secure_dir, exist_ok=True)
                test_file = os.path.join(secure_dir, ".perm_test")
                with open(test_file, "w", encoding="utf-8") as f:
                    f.write("ok")
                if os.path.exists(test_file):
                    os.remove(test_file)
                return os.path.join(secure_dir, "legal_consent.json")
            except Exception:
                continue
        fallback_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".security")
        os.makedirs(fallback_dir, exist_ok=True)
        return os.path.join(fallback_dir, "legal_consent.json")

    @classmethod
    def has_user_consented(cls) -> bool:
        consent_path = cls.get_consent_file_path()
        if not os.path.exists(consent_path):
            return False
        try:
            with open(consent_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return (
                    data.get("consented") is True and
                    data.get("version") == LEGAL_VERSION
                )
        except Exception:
            return False

    @classmethod
    def record_consent(cls) -> bool:
        consent_path = cls.get_consent_file_path()
        data = {
            "consented": True,
            "version": LEGAL_VERSION,
            "timestamp": int(time.time()),
            "checksum": hashlib.sha256(LEGAL_DISCLAIMER_SUMMARY.encode("utf-8")).hexdigest()[:16]
        }
        try:
            os.makedirs(os.path.dirname(consent_path), exist_ok=True)
            with open(consent_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            return True
        except Exception:
            return False

