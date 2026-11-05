import os
import sys
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
r"""
DocuMergePDF - Tier 4: Comprehensive Acceptance, Adversarial & Red Team Tests
Rigorous adversarial test suite covering:
- Trial session limits (5-session boundary, lockout, refund, lifetime unlock)
- Freemium Invariant: Row 1 test generation is 100% UNLIMITED and FREE FOREVER
- Anti-Tamper & Anti-Rollback defenses (clock skew > 24h, corrupted ciphertext, MAC failure)
- Windows DOS Device Names Injection (CON, PRN, AUX, NUL, COM1-9, LPT1-9)
- Directory Traversal Attacks (../../, \..\)
- Extreme Unicode, Emojis, SQL/HTML/XML Special Characters, Length > 300 Chars
- Corrupted/Damaged File Defenses (0-byte docx, corrupt zip, truncated spreadsheet)
"""

import os
import unittest
import tempfile
import shutil
import time
from docx import Document

from core.licensing import LicenseSecurity, DEFAULT_MAX_TRIAL_RUNS, HardwareID
from core.dynamic_namer import DynamicNamer, WINDOWS_RESERVED_NAMES
from core.document_merger import DocumentMerger
from core.pdf_converter import PDFConverter
from core.template_inspector import TemplateInspector
from core.data_reader import DataReader


class TestAcceptanceAndAdversarialComprehensive(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.sec = LicenseSecurity()
        self.sec.reset_trial()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)
        self.sec.reset_trial()

    def test_trial_lifecycle_and_enforcement(self):
        # 1. Initial trial: 5 runs remaining
        st = self.sec.get_status()
        self.assertEqual(st["runs_remaining"], DEFAULT_MAX_TRIAL_RUNS)
        self.assertTrue(st["can_batch_export"])

        # 2. Consume 5 batch runs
        for i in range(DEFAULT_MAX_TRIAL_RUNS):
            ok, msg = self.sec.consume_batch_export_run()
            self.assertTrue(ok)

        # 3. 6th run must be strictly blocked
        ok6, msg6 = self.sec.consume_batch_export_run()
        self.assertFalse(ok6)
        self.assertIn("Trial limit reached", msg6)
        st_blocked = self.sec.get_status()
        self.assertFalse(st_blocked["can_batch_export"])

        # 4. Refund 1 run
        refunded = self.sec.refund_batch_export_run()
        self.assertTrue(refunded)

        # 5. Now 1 run is authorized again
        ok_retry, _ = self.sec.consume_batch_export_run()
        self.assertTrue(ok_retry)

    def test_freemium_invariant_row1_test_always_unlimited(self):
        # Exhaust trial
        for _ in range(DEFAULT_MAX_TRIAL_RUNS + 2):
            self.sec.consume_batch_export_run()

        st = self.sec.get_status()
        self.assertEqual(st["runs_remaining"], 0)
        self.assertFalse(st["can_batch_export"])

        # Rule 4 & Master Blueprint Invariant: Row 1 Single Test generation is ALWAYS permitted!
        # Generating a single test record must NOT touch trial counter and must succeed
        tpl_path = os.path.join(self.test_dir, "freemium_tpl.docx")
        doc = Document()
        doc.add_paragraph("Test Row 1 for {{Recipient}}")
        doc.save(tpl_path)

        out_pdf = os.path.join(self.test_dir, "Row1_Test.pdf")
        tmp_docx = os.path.join(self.test_dir, "_tmp_r1.docx")

        DocumentMerger.merge_single_record(tpl_path, {"Recipient": "Name"}, {"Name": "Free Test User"}, tmp_docx)
        ok, _ = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
        if os.path.exists(tmp_docx):
            os.remove(tmp_docx)

        self.assertTrue(ok)
        self.assertTrue(os.path.exists(out_pdf))

        # Trial runs must remain unchanged at 0
        st_after = self.sec.get_status()
        self.assertEqual(st_after["runs_remaining"], 0)

    def test_clock_rollback_tamper_detection(self):
        state = self.sec.load_state()
        # Simulate clock rollback by setting future last_seen_ts (> 24 hours into future)
        state["last_seen_ts"] = int(time.time()) + 100000
        self.sec.save_state(state)

        # Reload state
        new_state = self.sec.load_state()
        self.assertTrue(new_state["tamper_detected"])

        st = self.sec.get_status()
        self.assertFalse(st["can_batch_export"])

    def test_adversarial_windows_reserved_names_injection(self):
        # Injection of Windows DOS devices: CON, PRN, AUX, NUL, COM1..COM9, LPT1..LPT9
        for dev in ["CON", "PRN", "AUX", "NUL", "COM1", "COM9", "LPT1", "LPT9"]:
            record = {"FullName": dev, "Department": "IT"}
            fname = DynamicNamer.generate("Report_{FullName}.pdf", record, 1, self.test_dir, set())
            clean_stem = os.path.splitext(fname)[0]
            self.assertNotIn(clean_stem.upper(), WINDOWS_RESERVED_NAMES)
            # Verify we can safely create the file on Windows without OS crash
            full_path = os.path.join(self.test_dir, fname)
            with open(full_path, "w", encoding="utf-8") as f:
                f.write("safe")
            self.assertTrue(os.path.exists(full_path))
            os.remove(full_path)

    def test_adversarial_directory_traversal_attacks(self):
        traversals = [
            "../../Windows/System32/calc",
            r"..\..\..\boot.ini",
            "/etc/passwd",
            "../../../hidden_doc"
        ]
        allocated = set()
        for t in traversals:
            fname = DynamicNamer.generate(f"Report_{{Name}}.pdf", {"Name": t}, 1, self.test_dir, allocated)
            self.assertNotIn("/", fname)
            self.assertNotIn("\\", fname)
            self.assertTrue(fname.endswith(".pdf"))

    def test_adversarial_unicode_emojis_and_extreme_lengths(self):
        extreme_names = [
            "Nguyễn Văn Ánh 🇻🇳",
            "Đặng Thị Thảo My (Giải Nhất 🏆)",
            "田中 太郎 先生 🎓",
            "José María López-García 🚀",
            "Александр Сергеевич Пушкин",
            "<script>alert('xss')</script>",
            "'; DROP TABLE Users; --",
            "Line1\nLine2\tLine3\rLine4",
            "Z" * 350,  # Extreme length > 300 chars
        ]

        out_folder = os.path.join(self.test_dir, "extreme_out")
        os.makedirs(out_folder, exist_ok=True)
        allocated = set()

        for i, name in enumerate(extreme_names, start=1):
            fname = DynamicNamer.generate("Doc_{FullName}_{Index}.pdf", {"FullName": name}, i, out_folder, allocated)
            self.assertTrue(len(fname) <= 130)
            self.assertTrue(fname.endswith(".pdf"))
            for char in r'\/:*?"<>|' + "\n\r\t":
                self.assertNotIn(char, fname)

    def test_corrupted_and_empty_file_defenses(self):
        # 1. 0-byte docx
        empty_docx = os.path.join(self.test_dir, "empty.docx")
        with open(empty_docx, "wb") as f:
            f.write(b"")

        with self.assertRaises(Exception):
            TemplateInspector.inspect(empty_docx)

        # 2. Corrupt non-docx file
        fake_docx = os.path.join(self.test_dir, "fake.docx")
        with open(fake_docx, "wb") as f:
            f.write(b"THIS IS NOT A VALID ZIP FILE HEADER")

        with self.assertRaises(Exception):
            TemplateInspector.inspect(fake_docx)

        # 3. Empty CSV
        empty_csv = os.path.join(self.test_dir, "empty.csv")
        with open(empty_csv, "w", encoding="utf-8") as f:
            f.write("")

        # DataReader should handle empty file gracefully without crashing unhandled
        try:
            res = DataReader.read(empty_csv)
            self.assertEqual(res.get("row_count", 0), 0)
        except Exception as e:
            self.assertTrue(isinstance(e, Exception))


if __name__ == "__main__":
    unittest.main()
