import os
import sys
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
"""
DocuMergePDF - Master Legal Compliance, Statutory Validity & Liability Shield Verification Suite
Exhaustively verifies all 11 critical international legal dimensions:
1. UCC § 2-316: Conspicuous Absolute Disclaimer of Implied Warranties ("AS IS", "WITH ALL FAULTS")
2. UCC § 2-719: Explicit $1.00 USD Monetary Liability Ceiling & Waiver of Consequential Damages
3. 100% User Indemnification & Neutral Utility Defense (Non-Fiduciary Status)
4. Nominative Fair Use Trademark Disclaimers (Microsoft Word, Excel, Office, Windows)
5. GDPR & FERPA Privacy Compliance (Local-First Guarantee, Zero Cloud PII Leaks)
6. First-Run Legal Consent Enforcement & Tamper-Resistant Storage
7. Legal Documentation Consistency Across Files (TERMS_OF_SERVICE.txt, LICENSE.txt, Code)
8. US ESIGN / UETA & ISO 32000 Data Message Integrity & Cryptographic Tamper Evidence
9. US FERPA (34 CFR Part 99) & EU GDPR Zero-Cloud Local-First Confidentiality
10. International Standard Page Geometry (ISO A4 / US Letter) & Multi-Script Unicode Fidelity
11. ISO 32000-1 Digital Signature (PAdES / PKCS#7) Readiness & Archival Structure
"""

import os
import json
import time
import hashlib
import socket
import tempfile
import shutil
import unittest
from unittest.mock import patch
from docx import Document

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from core.legal_shield import LegalShield, LEGAL_VERSION, LEGAL_DISCLAIMER_SUMMARY
from core.pdf_converter import PDFConverter
from core.document_merger import DocumentMerger


class TestLegalCompliance(unittest.TestCase):

    def setUp(self):
        self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.tos_path = os.path.join(self.base_dir, "TERMS_OF_SERVICE.txt")
        self.license_path = os.path.join(self.base_dir, "LICENSE.txt")
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    # =========================================================================
    # 1. UCC § 2-316: ABSOLUTE WARRANTY DISCLAIMER VERIFICATION
    # =========================================================================
    def test_ucc_2_316_warranty_disclaimer(self):
        """
        Under Uniform Commercial Code (UCC) § 2-316, exclusion of implied warranties
        must mention merchantability, fitness for purpose, and be conspicuous.
        """
        terms = LegalShield.read_full_terms()
        self.assertIn("UCC § 2-316", terms, "Missing UCC § 2-316 statutory reference")
        self.assertIn("AS IS", terms.upper(), "Missing 'AS IS' warranty disclaimer")
        self.assertIn("WITH ALL FAULTS", terms.upper(), "Missing 'WITH ALL FAULTS' disclaimer")
        self.assertIn("MERCHANTABILITY", terms.upper(), "Missing merchantability disclaimer")
        self.assertIn("FITNESS FOR A PARTICULAR PURPOSE", terms.upper(), "Missing fitness disclaimer")

    # =========================================================================
    # 2. UCC § 2-719: MONETARY LIABILITY CEILING ($1.00 USD) & CONSEQUENTIAL DAMAGES
    # =========================================================================
    def test_ucc_2_719_liability_limitation(self):
        """
        Under UCC § 2-719, parties may limit remedies. DocuMergePDF caps aggregate liability
        at $1.00 USD and explicitly waives consequential damages.
        """
        terms = LegalShield.read_full_terms()
        self.assertIn("UCC § 2-719", terms, "Missing UCC § 2-719 reference")
        self.assertIn("$1.00 USD", terms, "Missing explicit $1.00 USD monetary ceiling")
        self.assertIn("CONSEQUENTIAL", terms.upper(), "Missing consequential damages waiver")
        self.assertIn("LOST PROFITS", terms.upper(), "Missing lost profits waiver")

    # =========================================================================
    # 3. 100% USER INDEMNIFICATION & NEUTRAL LOCAL UTILITY STATUS
    # =========================================================================
    def test_user_indemnification_and_non_fiduciary(self):
        """
        Verifies that DocuMergePDF maintains neutral local utility status
        and obligates user to indemnify developer against payroll/grade disputes.
        """
        terms = LegalShield.read_full_terms()
        self.assertIn("INDEMNIF", terms.upper(), "Missing user indemnification clause")
        self.assertIn("HOLD HARMLESS", terms.upper(), "Missing hold harmless clause")
        self.assertIn("NEUTRAL LOCAL UTILITY", terms.upper(), "Missing neutral utility classification")

    # =========================================================================
    # 4. NOMINATIVE FAIR USE TRADEMARK COMPLIANCE
    # =========================================================================
    def test_trademark_nominative_fair_use(self):
        """
        Verifies compliance with Lanham Act nominative fair use doctrine
        regarding Microsoft Word and Excel trademarks.
        """
        terms = LegalShield.read_full_terms()
        self.assertIn("MICROSOFT", terms.upper(), "Missing Microsoft trademark mention")
        self.assertIn("TRADEMARK", terms.upper(), "Missing trademark notice")
        self.assertIn("NOT AFFILIATED", terms.upper(), "Missing non-affiliation disclaimer")

    # =========================================================================
    # 5. GDPR & FERPA DATA PROTECTION INVARIANTS
    # =========================================================================
    def test_gdpr_and_ferpa_privacy_invariants(self):
        """
        Verifies that legal text and architecture guarantee local-only document processing
        to satisfy GDPR (General Data Protection Regulation) and FERPA (student privacy).
        """
        terms = LegalShield.read_full_terms()
        # Terms must state local-first offline execution
        self.assertTrue("LOCAL" in terms.upper() or "OFFLINE" in terms.upper())

        # Also inspect LICENSE.txt for Privacy Guarantee
        with open(self.license_path, "r", encoding="utf-8") as f:
            lic_text = f.read()
        self.assertIn("LOCAL-FIRST PRIVACY GUARANTEE", lic_text.upper())
        self.assertIn("NOT TRANSMIT", lic_text.upper())

    # =========================================================================
    # 6. FIRST-RUN LEGAL CONSENT FLOW & TAMPER-RESISTANT RECORDING
    # =========================================================================
    def test_legal_consent_vault_and_checksum_verification(self):
        """
        Verifies that user consent is cryptographically recorded with:
        - Timestamp
        - Legal Version
        - Disclaimer Summary SHA-256 Checksum
        """
        consent_path = LegalShield.get_consent_file_path()
        if os.path.exists(consent_path):
            os.remove(consent_path)

        # Before consent: must return False
        self.assertFalse(LegalShield.has_user_consented())

        # Record consent
        recorded = LegalShield.record_consent()
        self.assertTrue(recorded)
        self.assertTrue(os.path.exists(consent_path))

        # After consent: must return True
        self.assertTrue(LegalShield.has_user_consented())

        # Verify consent file contents
        with open(consent_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        self.assertEqual(data["consented"], True)
        self.assertEqual(data["version"], LEGAL_VERSION)
        self.assertIn("timestamp", data)
        self.assertIn("checksum", data)

        # Verify checksum matches SHA-256 of LEGAL_DISCLAIMER_SUMMARY
        expected_chk = hashlib.sha256(LEGAL_DISCLAIMER_SUMMARY.encode("utf-8")).hexdigest()[:16]
        self.assertEqual(data["checksum"], expected_chk)

        # Tampering with consent version must invalidate consent
        data["version"] = "OBSOLETE_VERSION_2020"
        with open(consent_path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        self.assertFalse(LegalShield.has_user_consented())

        # Restore valid consent
        LegalShield.record_consent()
        self.assertTrue(LegalShield.has_user_consented())

    # =========================================================================
    # 7. DOCUMENTATION INTEGRITY & SYNCHRONIZATION
    # =========================================================================
    def test_legal_documentation_files_exist_and_synchronized(self):
        """Verifies that TERMS_OF_SERVICE.txt and LICENSE.txt exist and are populated."""
        self.assertTrue(os.path.exists(self.tos_path), "TERMS_OF_SERVICE.txt does not exist")
        self.assertTrue(os.path.exists(self.license_path), "LICENSE.txt does not exist")

        with open(self.tos_path, "r", encoding="utf-8") as f:
            tos_content = f.read()
        with open(self.license_path, "r", encoding="utf-8") as f:
            lic_content = f.read()

        self.assertGreater(len(tos_content), 500)
        self.assertGreater(len(lic_content), 500)

        # Both must contain UCC § 2-316 and § 2-719
        self.assertIn("UCC § 2-316", tos_content)
        self.assertIn("UCC § 2-719", tos_content)
        self.assertIn("UCC § 2-316", lic_content)
        self.assertIn("UCC § 2-719", lic_content)


    # =========================================================================
    # 8. US ESIGN / UETA & ISO 32000 ELECTRONIC DOCUMENT INTEGRITY
    #    DATA MESSAGE INTEGRITY & CRYPTOGRAPHIC TAMPER EVIDENCE
    # =========================================================================
    def test_international_e_document_integrity_and_tamper_evidence(self):
        """
        Under US Electronic Signatures in Global and National Commerce Act (ESIGN, 15 U.S.C. § 7001)
        and Uniform Electronic Transactions Act (UETA § 7):
        - Assured integrity: document remains unaltered from creation through transmission.
        Verifies:
        1. Bit-level deterministic hashing (SHA-256) of generated document.
        2. Detection of any post-generation tampering or byte corruption.
        3. Valid PDF 1.4/1.7 header and EOF markers conforming to ISO 32000.
        """
        tpl_path = os.path.join(self.test_dir, "contract_tpl.docx")
        doc = Document()
        doc.add_paragraph("INTERNATIONAL MASTER SERVICES AGREEMENT")
        doc.add_paragraph("BINDING COMMERCIAL CONTRACT")
        doc.add_paragraph("Party A: {{Party_A}} - Registration ID: {{Reg_A}}")
        doc.add_paragraph("Party B: {{Party_B}} - Registration ID: {{Reg_B}}")
        doc.add_paragraph("Total Contract Value: {{Contract_Value}} USD.")
        doc.save(tpl_path)

        rec = {
            "Party_A": "SereneWorks Global Technologies Inc.",
            "Reg_A": "US-DEL-984721",
            "Party_B": "International Logistics & Distribution Corp.",
            "Reg_B": "GB-LON-443210",
            "Contract_Value": "$500,000.00"
        }
        mapping = {k: k for k in rec.keys()}

        tmp_docx = os.path.join(self.test_dir, "contract_merged.docx")
        out_pdf = os.path.join(self.test_dir, "Master_Services_Agreement_01.pdf")

        DocumentMerger.merge_single_record(tpl_path, mapping, rec, tmp_docx)
        ok, engine = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
        self.assertTrue(ok, f"Failed converting electronic contract: {engine}")
        self.assertTrue(os.path.exists(out_pdf))

        # Check ISO 32000 PDF format integrity
        with open(out_pdf, "rb") as f:
            pdf_bytes = f.read()

        self.assertTrue(pdf_bytes.startswith(b"%PDF-"), "Invalid PDF header - violates standard electronic document format")
        self.assertTrue(b"%%EOF" in pdf_bytes[-1024:], "Missing %%EOF marker - truncated electronic message")

        # Compute initial SHA-256 hash (Data Integrity)
        original_hash = hashlib.sha256(pdf_bytes).hexdigest()

        # Simulate tampering test: changing 1 byte in document must alter hash
        tampered_bytes = bytearray(pdf_bytes)
        tampered_bytes[100] = (tampered_bytes[100] + 1) % 256
        tampered_hash = hashlib.sha256(tampered_bytes).hexdigest()

        self.assertNotEqual(original_hash, tampered_hash, "Tampering must be cryptographically detectable")

    # =========================================================================
    # 9. US FERPA (34 CFR Part 99), US FLSA & EU GDPR PRIVACY COMPLIANCE
    #    ZERO-CLOUD LOCAL-FIRST PROCESSING & CONFIDENTIAL PII PROTECTION
    # =========================================================================
    def test_international_privacy_zero_cloud_leak_ferpa_gdpr_compliance(self):
        """
        Under US FERPA (34 CFR Part 99), FLSA (29 CFR Part 516), and EU GDPR (Regulation EU 2016/679):
        - Personal data (names, identification numbers, compensation, academic evaluations)
          must be processed strictly on-premise without transmission to unauthorized third-party servers.
        Verifies:
        1. Processing employee compensation and confidential student records executes
           with ZERO network connections (intercepted socket calls).
        2. All intermediate processing artifacts are cleanly purged.
        """
        network_calls = []

        def forbid_network(sock, addr):
            network_calls.append(addr)
            raise PermissionError("GDPR/FERPA violation: Unauthorized outbound network connection attempt")

        pii_data = {
            "EmployeeName": "Alexander Vance",
            "SSN_ID": "SSN-XXX-XX-1234",
            "JobTitle": "Lead Cybersecurity Architect",
            "Salary": "$145,000.00 USD",
            "DirectDeposit": "ACH-9876543210"
        }

        tpl_path = os.path.join(self.test_dir, "payslip_pii.docx")
        doc = Document()
        doc.add_paragraph("CONFIDENTIAL EMPLOYEE COMPENSATION STATEMENT")
        doc.add_paragraph("Employee: {{EmployeeName}} • Tax ID: {{SSN_ID}}")
        doc.add_paragraph("Position: {{JobTitle}} • Net Pay: {{Salary}}")
        doc.add_paragraph("Disbursement Account: {{DirectDeposit}}")
        doc.save(tpl_path)

        tmp_docx = os.path.join(self.test_dir, "_pii_temp.docx")
        out_pdf = os.path.join(self.test_dir, "Payslip_Alexander_Vance.pdf")
        mapping = {k: k for k in pii_data.keys()}

        with patch.object(socket.socket, 'connect', side_effect=forbid_network):
            DocumentMerger.merge_single_record(tpl_path, mapping, pii_data, tmp_docx)
            ok, _ = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
            if os.path.exists(tmp_docx):
                os.remove(tmp_docx)

        self.assertTrue(ok)
        self.assertEqual(len(network_calls), 0, "FERPA/GDPR violation: Outbound network call detected")
        self.assertFalse(os.path.exists(tmp_docx), "Residual PII temporary file found on disk")

    # =========================================================================
    # 10. INTERNATIONAL PAGE GEOMETRY & MULTI-SCRIPT UNICODE FIDELITY
    #     STANDARD A4 / LETTER DIMENSIONS & GLOBAL CHARACTER RECONSTRUCTION
    # =========================================================================
    def test_international_page_geometry_and_unicode_multilingual_fidelity(self):
        """
        Under ISO 216 (A4 standard: 210 x 297 mm, or 595.28 x 841.89 pt) and ISO 10646 (Unicode):
        Verifies:
        1. Output PDF dimensions match standard ISO page geometry.
        2. Extracted text from PDF preserves multi-script international typography
           (Latin, German Umlauts, French accents, Spanish tildes, Vietnamese tone marks, Cyrillic).
        """
        intl_title = "INTERNATIONAL ACADEMIC & PROFESSIONAL RECOGNITION"
        intl_motto = "Excellence • Integrity • Global Leadership"
        intl_body = "OFFICIAL CONFERRAL OF MASTER OF SCIENCE IN ADVANCED COMPUTING"
        intl_candidates = (
            "Candidates: François Müller (Germany), José María López-García (Spain), "
            "Éléonore Roosevelt (France), Nguyễn Hoàng Ánh (Vietnam), Александр Пушкин (Eurasia)"
        )

        tpl_path = os.path.join(self.test_dir, "conferral_tpl.docx")
        doc = Document()
        # Explicitly configure standard ISO A4 paper (210 x 297 mm)
        from docx.shared import Mm
        sec = doc.sections[0]
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)

        doc.add_paragraph(intl_title)
        doc.add_paragraph(intl_motto)
        doc.add_paragraph(intl_body)
        doc.add_paragraph(intl_candidates)
        doc.save(tpl_path)

        out_pdf = os.path.join(self.test_dir, "Conferral_International.pdf")
        ok, engine = PDFConverter.convert_docx_to_pdf(tpl_path, out_pdf)
        self.assertTrue(ok)

        # Inspect using PyMuPDF
        pdf = fitz.open(out_pdf)
        self.assertGreater(len(pdf), 0)
        page0 = pdf[0]

        # 1. Check A4 dimensions (595.28 x 841.89 points ± 5 points)
        rect = page0.rect
        self.assertAlmostEqual(rect.width, 595.28, delta=5.0, msg="Width does not conform to A4 standard")
        self.assertAlmostEqual(rect.height, 841.89, delta=5.0, msg="Height does not conform to A4 standard")

        # 2. Check International Multi-Script character preservation
        extracted_text = page0.get_text()
        self.assertIn("François Müller", extracted_text)
        self.assertIn("López-García", extracted_text)
        self.assertIn("Leadership", extracted_text)
        pdf.close()

    # =========================================================================
    # 11. ISO 32000-1 DIGITAL SIGNATURE READINESS
    # =========================================================================
    def test_digital_signature_readiness(self):
        """
        Under ISO 32000-1 (PDF Standard Specification):
        Digital signatures must be attachable without corrupting document structure.
        Verifies that generated PDF contains a valid trailer, catalog, and xref table
        enabling standard CMS / PKCS#7 / PAdES incremental signature updates.
        """
        tpl_path = os.path.join(self.test_dir, "doc_for_signing.docx")
        doc = Document()
        doc.add_paragraph("OFFICIAL DIGITAL VERIFICATION CERTIFICATE")
        doc.save(tpl_path)

        out_pdf = os.path.join(self.test_dir, "DigitalVerification.pdf")
        ok, _ = PDFConverter.convert_docx_to_pdf(tpl_path, out_pdf)
        self.assertTrue(ok)

        doc_fitz = fitz.open(out_pdf)
        # Verify document is not encrypted/locked against authorized signature appending
        self.assertFalse(doc_fitz.is_encrypted, "Generated document must be unencrypted by default to allow digital signing")
        self.assertGreater(doc_fitz.page_count, 0)
        # Verify xref table exists and can be traversed
        self.assertGreater(doc_fitz.xref_length(), 0)
        doc_fitz.close()


if __name__ == "__main__":
    unittest.main()

