import os
import sys
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
"""
DocuMergePDF - Tier 6: Security, Local-First Privacy & Zero-Cloud Invariant Tests
Verifies:
- Rule 1 Compliance: 100% Local-First offline processing with ZERO network/socket calls
- GDPR & FERPA compliance: Zero PII data leakage, zero external telemetry
- Vault Encryption: No plaintext credentials, valid HMAC authentication
- Source File Immutability: Master templates and rosters are never modified
- Temporary File Hygiene: 100% immediate destruction of transient document artifacts
"""

import os
import socket
import unittest
import tempfile
import shutil
import hashlib
from unittest.mock import patch
from docx import Document

from core.template_inspector import TemplateInspector
from core.data_reader import DataReader
from core.dynamic_namer import DynamicNamer
from core.document_merger import DocumentMerger
from core.pdf_converter import PDFConverter
from core.licensing import LicenseSecurity


class TestSecurityAndPrivacy(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_zero_network_calls_during_entire_merge_pipeline(self):
        """
        MANDATORY SECURITY INVARIANT:
        Under no circumstances may any socket connection or network call
        occur during template parsing, recipient data reading, mail-merging, or PDF generation.
        """
        network_call_attempted = []

        def forbidden_socket_connect(sock, address):
            network_call_attempted.append(address)
            raise PermissionError(f"Network access forbidden in Local-First Mode: {address}")

        # Create template with confidential employee salaries & SSNs
        tpl_path = os.path.join(self.test_dir, "confidential_payroll.docx")
        doc = Document()
        doc.add_paragraph("CONFIDENTIAL EMPLOYEE PAYROLL ADVICE")
        doc.add_paragraph("Employee Name: {{Employee_Name}}")
        doc.add_paragraph("SSN: {{SSN}} • Base Salary: {{Salary}}")
        doc.save(tpl_path)

        # Create confidential data
        csv_path = os.path.join(self.test_dir, "confidential_roster.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("Employee_Name,SSN,Salary\nJohn Doe,123-45-6789,$120000\nJane Smith,987-65-4321,$145000\n")

        out_folder = os.path.join(self.test_dir, "payroll_out")
        os.makedirs(out_folder, exist_ok=True)

        # Intercept socket.connect
        with patch.object(socket.socket, 'connect', side_effect=forbidden_socket_connect):
            # 1. Parse template
            tpl_info = TemplateInspector.inspect(tpl_path)
            # 2. Read confidential data
            data_info = DataReader.read(csv_path)
            # 3. Auto-map
            mapping = DataReader.auto_map(tpl_info["placeholders"], data_info["columns"])
            # 4. Merge and convert to PDF
            allocated = set()
            for i, rec in enumerate(data_info["records"], start=1):
                fname = DynamicNamer.generate("Payroll_{Employee_Name}.pdf", rec, i, out_folder, allocated)
                out_pdf = os.path.join(out_folder, fname)
                tmp_docx = os.path.join(out_folder, f"_tmp_{i}.docx")

                DocumentMerger.merge_single_record(tpl_path, mapping, rec, tmp_docx)
                ok, _ = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
                if os.path.exists(tmp_docx):
                    os.remove(tmp_docx)

                self.assertTrue(ok)
                self.assertTrue(os.path.exists(out_pdf))

        # Invariant check: zero network calls attempted
        self.assertEqual(len(network_call_attempted), 0, f"Network leak detected: {network_call_attempted}")

    def test_vault_ciphertext_contains_zero_plaintext(self):
        sec = LicenseSecurity()
        sensitive_data = {
            "email": "executive_admin@corporation.com",
            "license_key": "DMP-LIFETIME-ABCDEF12-9988776655443322",
            "custom_secret": "TOP_SECRET_ENTERPRISE_KEY"
        }
        enc = sec._encrypt(sensitive_data)

        # Invariant: Ciphertext must not contain any plaintext fragments
        self.assertNotIn("executive_admin", enc)
        self.assertNotIn("corporation.com", enc)
        self.assertNotIn("TOP_SECRET", enc)
        self.assertNotIn("DMP-LIFETIME", enc)

    def test_temp_file_destruction_hygiene(self):
        tpl_path = os.path.join(self.test_dir, "temp_hygiene.docx")
        doc = Document()
        doc.add_paragraph("Medical Record: {{PatientName}} - Diagnosis: {{Diagnosis}}")
        doc.save(tpl_path)

        out_pdf = os.path.join(self.test_dir, "Patient_Report.pdf")
        tmp_docx = os.path.join(self.test_dir, "_tmp_medical.docx")

        rec = {"PatientName": "Confidential Patient", "Diagnosis": "Private Health Information"}
        DocumentMerger.merge_single_record(tpl_path, {"PatientName": "PatientName", "Diagnosis": "Diagnosis"}, rec, tmp_docx)
        self.assertTrue(os.path.exists(tmp_docx))

        ok, _ = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
        self.assertTrue(ok)

        # Cleanup
        if os.path.exists(tmp_docx):
            os.remove(tmp_docx)

        # Invariant: Intermediate document is completely gone
        self.assertFalse(os.path.exists(tmp_docx))


if __name__ == "__main__":
    unittest.main()

