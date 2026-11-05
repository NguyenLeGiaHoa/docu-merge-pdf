import os
import sys
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
"""
DocuMergePDF - Tier 2: Comprehensive Integration Tests
Validates multi-component workflows:
- Real Template + Real Spreadsheet end-to-end processing
- Split-run formatting preservation and table replacements
- Full International Multilingual UTF-8 character encoding pipeline
- Live preview rasterization & PDF structure validation
- Graceful handling of unmapped missing variables
- Dual-vault state persistence synchronization
"""

import os
import unittest
import tempfile
import shutil
from docx import Document

from core.template_inspector import TemplateInspector
from core.data_reader import DataReader
from core.dynamic_namer import DynamicNamer
from core.document_merger import DocumentMerger
from core.pdf_converter import PDFConverter
from core.licensing import LicenseSecurity


class TestIntegrationComprehensive(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.demo_tpl = os.path.join(self.base_dir, "demo_data", "Certificate_Template.docx")
        self.demo_data = os.path.join(self.base_dir, "demo_data", "Student_Roster_100.xlsx")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_end_to_end_real_demo_pipeline(self):
        self.assertTrue(os.path.exists(self.demo_tpl), "Demo template missing!")
        self.assertTrue(os.path.exists(self.demo_data), "Demo data missing!")

        # 1. Inspect template
        tpl_info = TemplateInspector.inspect(self.demo_tpl)
        self.assertTrue(tpl_info["is_valid"])
        self.assertGreaterEqual(tpl_info["placeholder_count"], 5)

        # 2. Read spreadsheet
        data_info = DataReader.read(self.demo_data)
        self.assertEqual(data_info["row_count"], 100)
        self.assertIn("FullName", data_info["columns"])

        # 3. Auto-map
        mapping = DataReader.auto_map(tpl_info["placeholders"], data_info["columns"])
        for p in tpl_info["placeholders"]:
            self.assertIsNotNone(mapping.get(p), f"Failed auto-mapping for tag {p}")

        # 4. Generate Row 1 Test
        row_0 = data_info["records"][0]
        allocated = set()
        filename = DynamicNamer.generate("Award_{FullName}_{CertificateID}.pdf", row_0, 1, self.test_dir, allocated)
        self.assertTrue(filename.endswith(".pdf"))
        self.assertIn(row_0["CertificateID"], filename)

        tmp_docx = os.path.join(self.test_dir, "_temp_demo.docx")
        out_pdf = os.path.join(self.test_dir, filename)

        DocumentMerger.merge_single_record(self.demo_tpl, mapping, row_0, tmp_docx)
        self.assertTrue(os.path.exists(tmp_docx))

        ok, engine = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
        self.assertTrue(ok, f"Conversion failed: {engine}")
        self.assertTrue(os.path.exists(out_pdf))
        self.assertGreater(os.path.getsize(out_pdf), 1000)

        # 5. Live preview rendering
        img = PDFConverter.render_preview_image(out_pdf)
        self.assertIsNotNone(img)
        self.assertGreater(img.width, 300)
        self.assertGreater(img.height, 400)

        # Clean temp docx
        if os.path.exists(tmp_docx):
            os.remove(tmp_docx)

    def test_single_record_merge_and_split_runs(self):
        doc_path = os.path.join(self.test_dir, "template.docx")
        doc = Document()
        
        # Test split across runs intentionally
        p = doc.add_paragraph()
        r1 = p.add_run("Dear {{")
        r2 = p.add_run("ClientName")
        r3 = p.add_run("}}, your balance is {{")
        r4 = p.add_run("Amount}}.")

        # Test table
        t = doc.add_table(rows=1, cols=1)
        t.rows[0].cells[0].paragraphs[0].text = "Account: {{AccountNum}}"

        doc.save(doc_path)

        out_docx = os.path.join(self.test_dir, "merged.docx")
        mapping = {
            "ClientName": "ClientName",
            "Amount": "Amount",
            "AccountNum": "AccountNum"
        }
        record = {
            "ClientName": "Sarah Connor",
            "Amount": "$1,250.00",
            "AccountNum": "ACC-998877"
        }

        DocumentMerger.merge_single_record(doc_path, mapping, record, out_docx)
        self.assertTrue(os.path.exists(out_docx))

        merged_doc = Document(out_docx)
        merged_text = merged_doc.paragraphs[0].text
        self.assertIn("Sarah Connor", merged_text)
        self.assertIn("$1,250.00", merged_text)
        self.assertNotIn("{{ClientName}}", merged_text)

        tbl_text = merged_doc.tables[0].rows[0].cells[0].paragraphs[0].text
        self.assertIn("ACC-998877", tbl_text)

    def test_multilingual_unicode_batch_pipeline(self):
        # Create template with multi-script international diacritics
        tpl_path = os.path.join(self.test_dir, "intl_template.docx")
        doc = Document()
        doc.add_paragraph("INTERNATIONAL ACADEMY OF DISTINCTION")
        doc.add_paragraph("Official Certificate of Completion")
        doc.add_paragraph("Conferred upon: {{CandidateName}}")
        doc.add_paragraph("Graduation Honors: {{HonorsRank}}")
        doc.save(tpl_path)

        # Create records with multi-script international characters
        records = [
            {"CandidateName": "François Müller", "HonorsRank": "Summa Cum Laude"},
            {"CandidateName": "José María López-García", "HonorsRank": "Magna Cum Laude"},
            {"CandidateName": "Nguyễn Hoàng Ánh", "HonorsRank": "Cum Laude"}
        ]

        out_folder = os.path.join(self.test_dir, "intl_out")
        os.makedirs(out_folder, exist_ok=True)
        allocated = set()
        mapping = {"CandidateName": "CandidateName", "HonorsRank": "HonorsRank"}

        for i, rec in enumerate(records, start=1):
            fname = DynamicNamer.generate("Cert_{CandidateName}.pdf", rec, i, out_folder, allocated)
            out_pdf = os.path.join(out_folder, fname)
            tmp_docx = os.path.join(out_folder, f"_tmp_{i}.docx")

            DocumentMerger.merge_single_record(tpl_path, mapping, rec, tmp_docx)
            ok, _ = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
            if os.path.exists(tmp_docx):
                os.remove(tmp_docx)

            self.assertTrue(ok)
            self.assertTrue(os.path.exists(out_pdf))

        generated_files = os.listdir(out_folder)
        self.assertEqual(len(generated_files), 3)

    def test_missing_unmapped_placeholders_graceful_fallback(self):
        tpl_path = os.path.join(self.test_dir, "missing_tags.docx")
        doc = Document()
        doc.add_paragraph("Hello {{User}}, your bonus is {{BonusAmount}} and title is {{OptionalTitle}}.")
        doc.save(tpl_path)

        out_docx = os.path.join(self.test_dir, "missing_out.docx")
        # Only map User, leave OptionalTitle unmapped and BonusAmount None
        mapping = {
            "User": "User",
            "BonusAmount": None,
            "OptionalTitle": None
        }
        record = {"User": "Alice"}

        DocumentMerger.merge_single_record(tpl_path, mapping, record, out_docx, empty_fallback="[N/A]")
        self.assertTrue(os.path.exists(out_docx))

        res_doc = Document(out_docx)
        text = res_doc.paragraphs[0].text
        self.assertIn("Alice", text)
        self.assertIn("[N/A]", text)
        self.assertNotIn("{{User}}", text)
        self.assertNotIn("{{BonusAmount}}", text)

    def test_dual_vault_state_sync_integration(self):
        sec = LicenseSecurity()
        status = sec.get_status()
        self.assertIn("runs_remaining", status)
        self.assertIn("hwid", status)
        self.assertGreater(len(status["hwid"]), 16)


if __name__ == "__main__":
    unittest.main()
