"""
DocuMergePDF - Automated Gumroad Release Packaging Script
Builds Tier1_FREE_TRIAL_Upload_To_Gumroad.zip and Tier2_LIFETIME_LICENSE_Upload_To_Gumroad.zip
Run: python package_gumroad.py
"""

import os
import sys
import zipfile

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(ROOT, "GUMROAD_UPLOAD_PACKAGES")
os.makedirs(OUT_DIR, exist_ok=True)

# Locate executable
EXE_PATH = os.path.join(ROOT, "DocuMergePDF.exe")
if not os.path.exists(EXE_PATH):
    dist_exe = os.path.join(ROOT, "dist", "DocuMergePDF.exe")
    if os.path.exists(dist_exe):
        EXE_PATH = dist_exe

DEMO_FILES = [
    os.path.join(ROOT, "demo_data", "Certificate_Template.docx"),
    os.path.join(ROOT, "demo_data", "Student_Roster_100.xlsx"),
    os.path.join(ROOT, "demo_data", "Payslip_Template.docx"),
    os.path.join(ROOT, "demo_data", "HR_Payroll_50.xlsx"),
]

LICENSE_PATH = os.path.join(ROOT, "LICENSE.txt")
TERMS_PATH = os.path.join(ROOT, "TERMS_OF_SERVICE.txt")
README_PATH = os.path.join(ROOT, "README.txt")
FREE_GUIDE_PATH = os.path.join(ROOT, "USER_GUIDE_FREE_TRIAL.md")
LIFETIME_GUIDE_PATH = os.path.join(ROOT, "USER_GUIDE_LIFETIME.md")


def package_tier1():
    out_zip = os.path.join(OUT_DIR, "Tier1_FREE_TRIAL_Upload_To_Gumroad.zip")
    print(f"[*] Packaging Tier 1 (Free Trial): {out_zip}...")
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
        if os.path.exists(EXE_PATH):
            z.write(EXE_PATH, "DocuMergePDF.exe")
        if os.path.exists(LICENSE_PATH):
            z.write(LICENSE_PATH, "LICENSE.txt")
        if os.path.exists(TERMS_PATH):
            z.write(TERMS_PATH, "TERMS_OF_SERVICE.txt")
        if os.path.exists(README_PATH):
            z.write(README_PATH, "README.txt")
        if os.path.exists(FREE_GUIDE_PATH):
            z.write(FREE_GUIDE_PATH, "USER_GUIDE_FREE_TRIAL.md")
        for df in DEMO_FILES:
            if os.path.exists(df):
                z.write(df, os.path.join("demo_data", os.path.basename(df)))
    print(f"[✓] Tier 1 Package created: {out_zip} ({os.path.getsize(out_zip) / (1024*1024):.2f} MB)")


def package_tier2():
    out_zip = os.path.join(OUT_DIR, "Tier2_LIFETIME_LICENSE_Upload_To_Gumroad.zip")
    print(f"[*] Packaging Tier 2 (Lifetime License): {out_zip}...")
    with zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as z:
        if os.path.exists(EXE_PATH):
            z.write(EXE_PATH, "DocuMergePDF.exe")
        if os.path.exists(LICENSE_PATH):
            z.write(LICENSE_PATH, "LICENSE.txt")
        if os.path.exists(TERMS_PATH):
            z.write(TERMS_PATH, "TERMS_OF_SERVICE.txt")
        if os.path.exists(README_PATH):
            z.write(README_PATH, "README.txt")
        if os.path.exists(LIFETIME_GUIDE_PATH):
            z.write(LIFETIME_GUIDE_PATH, "USER_GUIDE_LIFETIME.md")
        if os.path.exists(FREE_GUIDE_PATH):
            z.write(FREE_GUIDE_PATH, "USER_GUIDE_FREE_TRIAL.md")
        for df in DEMO_FILES:
            if os.path.exists(df):
                z.write(df, os.path.join("demo_data", os.path.basename(df)))
    print(f"[✓] Tier 2 Package created: {out_zip} ({os.path.getsize(out_zip) / (1024*1024):.2f} MB)")


def main():
    print("=" * 60)
    print("📦 DOCUMERGEPDF GUMROAD UPLOAD PACKAGE BUILDER")
    print("=" * 60)
    package_tier1()
    package_tier2()
    print("\n[✓] Both distribution archives successfully packaged in:")
    print(f"--> {OUT_DIR}")


if __name__ == "__main__":
    main()

