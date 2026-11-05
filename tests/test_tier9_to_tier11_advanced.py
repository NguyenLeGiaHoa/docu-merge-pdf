"""
DocuMergePDF - Supreme Quality Assurance: Tiers 9, 10 & 11
Tier 9: Standalone Executable Binary Smoke Test & Dynamic Runtime Memory RSS Check
Tier 10: Deep Vector PDF Forensic Inspection & Typography Quality (PyMuPDF)
Tier 11: Real-World Dual Dataset Production Run (Certificates & HR Payroll)
"""

import os
import sys
import unittest
import tempfile
import shutil
import time
import subprocess
import psutil
from docx import Document

root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from core.template_inspector import TemplateInspector
from core.data_reader import DataReader
from core.dynamic_namer import DynamicNamer
from core.document_merger import DocumentMerger
from core.pdf_converter import PDFConverter

try:
    import pymupdf as fitz
    HAS_PYMUPDF = True
except ImportError:
    try:
        import fitz
        HAS_PYMUPDF = True
    except ImportError:
        HAS_PYMUPDF = False


class TestTier9To11Advanced(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.base_dir = root_dir
        self.demo_dir = os.path.join(self.base_dir, "demo_data")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    # =========================================================================
    # TIER 9: STANDALONE PE EXECUTABLE RUNTIME SMOKE TEST
    # =========================================================================
    def test_tier9_standalone_pe_executable_runtime_smoke(self):
        """Verifies compiled DocuMergePDF.exe launches dynamically and operates stably."""
        exe_path = os.path.join(self.base_dir, "DocuMergePDF.exe")
        self.assertTrue(os.path.exists(exe_path), f"Executable not found at {exe_path}")

        proc = subprocess.Popen([exe_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        try:
            # Allow process 3.5 seconds to initialize
            time.sleep(3.5)
            poll = proc.poll()
            self.assertIsNone(poll, f"DocuMergePDF.exe crashed upon launch with exit code: {poll}")

            # Verify memory stability and process tree
            p = psutil.Process(proc.pid)
            mem_mb = p.memory_info().rss / (1024 * 1024)
            self.assertGreater(mem_mb, 1.0, "Suspiciously zero memory usage")
            self.assertLess(mem_mb, 250.0, f"Memory RSS exceeded limit: {mem_mb:.2f} MB")
        finally:
            proc.terminate()
            try:
                proc.wait(timeout=4)
            except subprocess.TimeoutExpired:
                proc.kill()
            if proc.stdout:
                proc.stdout.close()
            if proc.stderr:
                proc.stderr.close()

    # =========================================================================
    # TIER 10: DEEP VECTOR PDF FORENSIC & TYPOGRAPHY INSPECTION
    # =========================================================================
    def test_tier10_deep_vector_pdf_forensics(self):
        """Forensic inspection of generated PDF vector streams using PyMuPDF."""
        self.assertTrue(HAS_PYMUPDF, "PyMuPDF (fitz) is required for deep vector forensics")

        # 1. Create realistic certificate template with rich international typography and tables
        tpl_path = os.path.join(self.test_dir, "forensic_template.docx")
        doc = Document()
        doc.add_paragraph("INTERNATIONAL ACADEMY OF ADVANCED TECHNOLOGY")
        doc.add_paragraph("Global Academic Certification & Honors")
        doc.add_paragraph("OFFICIAL DIPLOMA OF DISTINCTION")
        doc.add_paragraph("Proudly Conferred Upon: {{Student_Name}}")
        doc.add_paragraph("Degree Major: {{Course_Title}}")

        tbl = doc.add_table(rows=2, cols=2)
        tbl.rows[0].cells[0].paragraphs[0].text = "Student ID: {{Citizen_ID}}"
        tbl.rows[0].cells[1].paragraphs[0].text = "Academic Honors: {{Graduation_Rank}}"
        tbl.rows[1].cells[0].paragraphs[0].text = "Conferral Date: {{Issue_Date}}"
        tbl.rows[1].cells[1].paragraphs[0].text = "Accredited Body: {{Institution}}"
        doc.save(tpl_path)

        # 2. Record with rich international multi-script diacritics
        record = {
            "Student_Name": "Alexander Vance-García & François Müller",
            "Course_Title": "Master of Science in Autonomous Artificial Intelligence",
            "Citizen_ID": "STUDENT-2026-998877",
            "Graduation_Rank": "Summa Cum Laude with Highest Distinction",
            "Issue_Date": "2026-10-25",
            "Institution": "International University of Science & Technology"
        }
        mapping = {k: k for k in record.keys()}

        tmp_docx = os.path.join(self.test_dir, "forensic_merged.docx")
        out_pdf = os.path.join(self.test_dir, "Forensic_Result.pdf")

        DocumentMerger.merge_single_record(tpl_path, mapping, record, tmp_docx)
        ok, engine = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
        self.assertTrue(ok, f"Forensic conversion failed: {engine}")

        # 3. Deep Forensic Analysis with PyMuPDF
        with fitz.open(out_pdf) as pdf:
            self.assertEqual(len(pdf), 1, "PDF should have exactly 1 page")
            page = pdf[0]
            extracted_text = page.get_text()

            # Invariant: No raw mustache tag artifact leaks
            self.assertNotIn("{{", extracted_text, "Found unreplaced '{{' tag in PDF text stream!")
            self.assertNotIn("}}", extracted_text, "Found unreplaced '}}' tag in PDF text stream!")

            # Verify key strings are preserved in the text stream
            self.assertIn("STUDENT-2026-998877", extracted_text)
            self.assertIn("2026-10-25", extracted_text)

            # Check page dimensions (should conform to standard A4 or Letter boundaries)
            rect = page.rect
            self.assertGreater(rect.width, 500)
            self.assertGreater(rect.height, 700)

            # Check rendering preview rasterization
            pix = page.get_pixmap(dpi=150)
            self.assertGreater(pix.width, 800)
            self.assertGreater(pix.height, 1000)
            self.assertGreater(len(pix.tobytes()), 10000)

    # =========================================================================
    # TIER 11: REAL-WORLD DUAL DATASET PRODUCTION PIPELINE (10+10 RECORDS)
    # =========================================================================
    def test_tier11_real_world_dual_dataset_pipeline(self):
        """Executes real-world multi-format pipelines: Academic Roster and Corporate Payroll."""
        # --- SUB-PIPELINE A: ACADEMIC CERTIFICATES ---
        cert_tpl = os.path.join(self.demo_dir, "Certificate_Template.docx")
        cert_data_path = os.path.join(self.demo_dir, "Student_Roster_100.xlsx")

        cert_insp = TemplateInspector.inspect(cert_tpl)
        cert_data = DataReader.read(cert_data_path)
        cert_mapping = DataReader.auto_map(cert_insp["placeholders"], cert_data["columns"])

        cert_out_dir = os.path.join(self.test_dir, "Certificates_Output")
        os.makedirs(cert_out_dir, exist_ok=True)
        cert_allocated = set()

        sample_students = cert_data["records"][:10]
        for idx, student in enumerate(sample_students, start=1):
            fname = DynamicNamer.generate(
                "Cert_{FullName}_{CertificateID}.pdf",
                student,
                idx,
                cert_out_dir,
                cert_allocated
            )
            out_pdf = os.path.join(cert_out_dir, fname)
            tmp_docx = os.path.join(cert_out_dir, f"_tmp_cert_{idx}.docx")

            DocumentMerger.merge_single_record(cert_tpl, cert_mapping, student, tmp_docx)
            ok, engine = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
            if os.path.exists(tmp_docx):
                os.remove(tmp_docx)

            self.assertTrue(ok, f"Failed converting certificate for {student['FullName']}")
            self.assertTrue(os.path.exists(out_pdf))
            self.assertGreater(os.path.getsize(out_pdf), 1000)

        self.assertEqual(len(os.listdir(cert_out_dir)), 10)

        # --- SUB-PIPELINE B: CORPORATE PAYSLIPS ---
        pay_tpl = os.path.join(self.demo_dir, "Payslip_Template.docx")
        pay_data_path = os.path.join(self.demo_dir, "HR_Payroll_50.xlsx")

        pay_insp = TemplateInspector.inspect(pay_tpl)
        pay_data = DataReader.read(pay_data_path)
        pay_mapping = DataReader.auto_map(pay_insp["placeholders"], pay_data["columns"])

        pay_out_dir = os.path.join(self.test_dir, "Payslips_Output")
        os.makedirs(pay_out_dir, exist_ok=True)
        pay_allocated = set()

        sample_employees = pay_data["records"][:10]
        for idx, emp in enumerate(sample_employees, start=1):
            fname = DynamicNamer.generate(
                "Payslip_{EmployeeID}_{EmployeeName}.pdf",
                emp,
                idx,
                pay_out_dir,
                pay_allocated
            )
            out_pdf = os.path.join(pay_out_dir, fname)
            tmp_docx = os.path.join(pay_out_dir, f"_tmp_pay_{idx}.docx")

            DocumentMerger.merge_single_record(pay_tpl, pay_mapping, emp, tmp_docx)
            ok, engine = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
            if os.path.exists(tmp_docx):
                os.remove(tmp_docx)

            self.assertTrue(ok, f"Failed converting payslip for {emp['EmployeeName']}")
            self.assertTrue(os.path.exists(out_pdf))
            self.assertGreater(os.path.getsize(out_pdf), 1000)

        self.assertEqual(len(os.listdir(pay_out_dir)), 10)


if __name__ == "__main__":
    unittest.main()
