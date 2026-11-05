import os
import sys
import zipfile
import unittest
import xml.etree.ElementTree as ET
from PIL import Image

root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

"""
DocuMergePDF - Tier 8: Packaging, Distribution & Binary Integrity Tests
Verifies:
1. Portable Executable (.exe) binary integrity, PE magic headers, and size benchmarks
2. Windows Store Package (.msix) container, AppxManifest schema, and asset compliance
3. Gumroad ZIP delivery bundles (Free Trial & Lifetime License) CRC and contents
4. Store marketing assets (resolution, aspect ratio, image dimensions)
5. User documentation and legal notices bundled in release packages
"""


class TestPackagingAndDistribution(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        cls.exe_path = os.path.join(cls.base_dir, "DocuMergePDF.exe")
        cls.msix_path = os.path.join(cls.base_dir, "DocuMergePDF.msix")
        cls.dist_exe_path = os.path.join(cls.base_dir, "dist", "DocuMergePDF.exe")
        cls.gumroad_dir = os.path.join(cls.base_dir, "GUMROAD_UPLOAD_PACKAGES")
        cls.store_assets_dir = os.path.join(cls.base_dir, "STORE_ASSETS")

    # =========================================================================
    # 1. STANDALONE PORTABLE EXECUTABLE (.EXE) INTEGRITY
    # =========================================================================
    def test_standalone_executable_binary_integrity(self):
        """Verifies that the compiled DocuMergePDF.exe exists, has valid PE headers, and proper size."""
        self.assertTrue(os.path.exists(self.exe_path), f"Missing executable at {self.exe_path}")
        size_bytes = os.path.getsize(self.exe_path)
        size_mb = size_bytes / (1024 * 1024)

        # PyInstaller bundle with customtkinter, reportlab, fitz, docx should be ~100MB - 250MB
        self.assertGreater(size_mb, 50.0, f"Executable too small ({size_mb:.2f} MB), incomplete bundling")
        self.assertLess(size_mb, 300.0, f"Executable unusually bloated ({size_mb:.2f} MB)")

        # Verify PE headers: DOS 'MZ' header and 'PE\0\0' signature
        with open(self.exe_path, "rb") as f:
            dos_header = f.read(64)
            self.assertEqual(dos_header[:2], b"MZ", "Executable missing valid DOS 'MZ' header")
            # PE header offset is at 0x3C (60)
            pe_offset = int.from_bytes(dos_header[60:64], byteorder="little")
            f.seek(pe_offset)
            pe_sig = f.read(4)
            self.assertEqual(pe_sig, b"PE\x00\x00", "Executable missing valid PE signature")

    # =========================================================================
    # 2. WINDOWS STORE PACKAGE (.MSIX) INTEGRITY & APPX MANIFEST
    # =========================================================================
    def test_msix_package_structure_and_manifest(self):
        """Verifies that DocuMergePDF.msix is a valid OPC/ZIP package with valid AppxManifest."""
        self.assertTrue(os.path.exists(self.msix_path), f"Missing MSIX package at {self.msix_path}")

        with zipfile.ZipFile(self.msix_path, "r") as z:
            # Check CRC integrity of all entries
            bad_crc = z.testzip()
            self.assertIsNone(bad_crc, f"Corrupted file found in MSIX: {bad_crc}")

            namelist = z.namelist()
            self.assertIn("AppxManifest.xml", namelist, "MSIX missing AppxManifest.xml")
            self.assertIn("AppxBlockMap.xml", namelist, "MSIX missing AppxBlockMap.xml")
            self.assertIn("[Content_Types].xml", namelist, "MSIX missing [Content_Types].xml")
            self.assertIn("DocuMergePDF.exe", namelist, "MSIX missing payload DocuMergePDF.exe")

            # Parse and inspect AppxManifest.xml
            manifest_bytes = z.read("AppxManifest.xml")
            root = ET.fromstring(manifest_bytes)
            # Find Package Identity
            identity = None
            for elem in root.iter():
                if elem.tag.endswith("Identity"):
                    identity = elem
                    break

            self.assertIsNotNone(identity, "AppxManifest missing Identity node")
            self.assertIn("DocuMergePDF", identity.attrib.get("Name", ""))

    # =========================================================================
    # 3. GUMROAD RELEASE PACKAGES VERIFICATION
    # =========================================================================
    def test_gumroad_release_zip_bundles(self):
        """Verifies that both Tier 1 Free Trial and Tier 2 Lifetime packages exist and are healthy."""
        trial_zip = os.path.join(self.gumroad_dir, "Tier1_FREE_TRIAL_Upload_To_Gumroad.zip")
        lifetime_zip = os.path.join(self.gumroad_dir, "Tier2_LIFETIME_LICENSE_Upload_To_Gumroad.zip")

        for zip_path, expected_guide in [
            (trial_zip, "USER_GUIDE_FREE_TRIAL.md"),
            (lifetime_zip, "USER_GUIDE_LIFETIME.md"),
        ]:
            self.assertTrue(os.path.exists(zip_path), f"Missing Gumroad bundle: {zip_path}")
            with zipfile.ZipFile(zip_path, "r") as z:
                self.assertIsNone(z.testzip(), f"CRC error in {zip_path}")
                names = z.namelist()
                self.assertIn("DocuMergePDF.exe", names)
                self.assertIn("LICENSE.txt", names)
                self.assertIn("TERMS_OF_SERVICE.txt", names)
                self.assertIn(expected_guide, names)
                # Verify demo data templates are packed
                self.assertTrue(any("Certificate_Template.docx" in n for n in names))
                self.assertTrue(any("Student_Roster_100.xlsx" in n for n in names))

    # =========================================================================
    # 4. STORE MARKETING ASSETS SPECIFICATION & RESOLUTION AUDIT
    # =========================================================================
    def test_store_assets_dimensions_and_quality(self):
        """Verifies that all 7 store assets strictly adhere to required dimensions."""
        expected_specs = {
            "AppTileIcon_300x300.png": (300, 300),
            "BoxArt_1000x1000.png": (1000, 1000),
            "BoxArt_1080x1080.png": (1080, 1080),
            "Icon_150x150.png": (150, 150),
            "Icon_71x71.png": (71, 71),
            "Screenshot1_Main_Workspace_1920x1080.png": (1920, 1080),
            "Screenshot2_Features_Workflow_1920x1080.png": (1920, 1080),
        }

        for fname, (expected_w, expected_h) in expected_specs.items():
            asset_path = os.path.join(self.store_assets_dir, fname)
            self.assertTrue(os.path.exists(asset_path), f"Missing store asset: {fname}")
            with Image.open(asset_path) as img:
                w, h = img.size
                self.assertEqual(
                    (w, h),
                    (expected_w, expected_h),
                    f"Asset {fname} has incorrect dimensions: got {w}x{h}, expected {expected_w}x{expected_h}",
                )

    # =========================================================================
    # 5. USER DOCUMENTATION INTEGRITY
    # =========================================================================
    def test_user_documentation_integrity(self):
        """Verifies that all end-user manuals and developer guides are complete and valid."""
        docs = [
            "PROJECT_BLUEPRINT.md",
            "TECHNICAL_ARCHITECTURE_RULES.md",
            "USER_GUIDE_FREE_TRIAL.md",
            "USER_GUIDE_LIFETIME.md",
            "README.md",
            "README.txt",
            "GUMROAD_SELLER_SETUP_GUIDE.md",
        ]
        for doc in docs:
            doc_path = os.path.join(self.base_dir, doc)
            self.assertTrue(os.path.exists(doc_path), f"Missing documentation file: {doc}")
            size = os.path.getsize(doc_path)
            self.assertGreater(size, 100, f"Documentation file {doc} appears empty or truncated")


if __name__ == "__main__":
    unittest.main()

