import os
import sys
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
"""
DocuMergePDF - Tier 1: Comprehensive Unit Tests
Exhaustive unit test coverage for all core engine modules:
- TemplateInspector (Mustache, Spaced, Single-brace, MergeFields, Tables, Headers, Footers)
- DataReader (Normalization, Delimiters, Encodings, Sheets, Auto-mapping, Error handling)
- DynamicNamer (Sanitization, Windows Reserved Names, Clamping, Collisions, Path Traversal)
- DocumentMerger (Paragraphs, Runs, Split-runs, Tables, Headers, Footers, Non-destructive)
- PDFConverter (ReportLab Tier 3 Fallback, PyMuPDF Preview Rasterizer, Missing File Handling)
- LicenseSecurity (HWID Determinism, AES+HMAC Crypto Roundtrip, Tamper Detection, Keys)
- LegalShield (Terms Reading, Consent Recording, Versioning, SHA-256 Checksums)
"""

import os
import unittest
import tempfile
import shutil
import hashlib
from docx import Document
import pandas as pd

from core.template_inspector import TemplateInspector
from core.data_reader import DataReader
from core.dynamic_namer import DynamicNamer, WINDOWS_RESERVED_NAMES
from core.document_merger import DocumentMerger
from core.pdf_converter import PDFConverter
from core.licensing import LicenseSecurity, HardwareID
from core.legal_shield import LegalShield, LEGAL_VERSION, LEGAL_DISCLAIMER_SUMMARY


class TestUnitComprehensive(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    # =========================================================================
    # 1. TEMPLATE INSPECTOR TESTS
    # =========================================================================
    def test_template_inspector_basic(self):
        doc_path = os.path.join(self.test_dir, "test_doc.docx")
        doc = Document()
        doc.add_paragraph("Hello {{FullName}}, welcome to {{Organization}}!")
        tbl = doc.add_table(rows=1, cols=2)
        tbl.rows[0].cells[0].paragraphs[0].text = "ID: {{StudentID}}"
        tbl.rows[0].cells[1].paragraphs[0].text = "Date: {{EnrollmentDate}}"
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertTrue(res["is_valid"])
        self.assertEqual(res["placeholder_count"], 4)
        self.assertIn("FullName", res["placeholders"])
        self.assertIn("Organization", res["placeholders"])
        self.assertIn("StudentID", res["placeholders"])
        self.assertIn("EnrollmentDate", res["placeholders"])

    def test_template_inspector_spaced_and_single_brace(self):
        doc_path = os.path.join(self.test_dir, "test_spaced.docx")
        doc = Document()
        doc.add_paragraph("Dear {{  Student_Name  }}, your code is {{CourseCode}}.")
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertTrue(res["is_valid"])
        self.assertIn("Student_Name", res["placeholders"])
        self.assertIn("CourseCode", res["placeholders"])

    def test_template_inspector_headers_and_footers(self):
        doc_path = os.path.join(self.test_dir, "test_hf.docx")
        doc = Document()
        section = doc.sections[0]
        header = section.header
        header.paragraphs[0].text = "Institution: {{InstitutionName}}"
        footer = section.footer
        footer.paragraphs[0].text = "Confidential • Document ID: {{DocUUID}}"
        doc.add_paragraph("Body text with {{BodyTag}}")
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertTrue(res["is_valid"])
        self.assertIn("InstitutionName", res["placeholders"])
        self.assertIn("DocUUID", res["placeholders"])
        self.assertIn("BodyTag", res["placeholders"])

    def test_template_inspector_chevrons_and_single_braces(self):
        doc_path = os.path.join(self.test_dir, "test_chevrons.docx")
        doc = Document()
        doc.add_paragraph("Dear «ClientName», your account {AccountID} has balance {{Balance}}.")
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertTrue(res["is_valid"])
        self.assertIn("ClientName", res["placeholders"])
        self.assertIn("AccountID", res["placeholders"])
        self.assertIn("Balance", res["placeholders"])

    def test_template_inspector_no_tags_and_invalid(self):
        doc_path = os.path.join(self.test_dir, "test_empty.docx")
        doc = Document()
        doc.add_paragraph("Plain text without any placeholders.")
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertFalse(res["is_valid"])
        self.assertEqual(res["placeholder_count"], 0)

        with self.assertRaises(FileNotFoundError):
            TemplateInspector.inspect(os.path.join(self.test_dir, "nonexistent.docx"))

    # =========================================================================
    # 2. DYNAMIC NAMER TESTS
    # =========================================================================
    def test_dynamic_namer_sanitization(self):
        bad_name = 'Dr. Smith / Cardiologist: Section * "A" <Special> | Final'
        clean = DynamicNamer.sanitize(bad_name)
        for char in r'\/:*?"<>|':
            self.assertNotIn(char, clean)
        self.assertTrue(len(clean) <= 120)

    def test_dynamic_namer_windows_reserved_names(self):
        # Testing CON, PRN, AUX, NUL, COM1, LPT1
        for res_name in ["CON", "prn", "Aux", "NUL", "COM1", "LPT1"]:
            clean = DynamicNamer.sanitize(res_name)
            self.assertNotIn(clean.upper(), WINDOWS_RESERVED_NAMES)
            self.assertTrue(clean.startswith("Doc_"))

    def test_dynamic_namer_length_clamping(self):
        super_long = "Recipient_" + "X" * 300
        clean = DynamicNamer.sanitize(super_long)
        self.assertLessEqual(len(clean), 120)

    def test_dynamic_namer_collision_avoidance(self):
        allocated = set()
        out_dir = self.test_dir
        pattern = "Report_{Name}_{ID}"
        record = {"Name": "John", "ID": "101"}

        name1 = DynamicNamer.generate(pattern, record, 1, out_dir, allocated)
        self.assertEqual(name1, "Report_John_101.pdf")

        # Second generation with same data must avoid collision
        name2 = DynamicNamer.generate(pattern, record, 2, out_dir, allocated)
        self.assertEqual(name2, "Report_John_101_01.pdf")

        # Third collision
        name3 = DynamicNamer.generate(pattern, record, 3, out_dir, allocated)
        self.assertEqual(name3, "Report_John_101_02.pdf")

    def test_dynamic_namer_pattern_tokens_and_path_traversal(self):
        allocated = set()
        # Path traversal attack attempt
        pattern = "../../SecretFolder/{FullName}_{Index}"
        record = {"FullName": "Attacker/User"}
        name = DynamicNamer.generate(pattern, record, 7, self.test_dir, allocated)
        self.assertNotIn("/", name)
        self.assertNotIn("\\", name)
        self.assertTrue(name.endswith(".pdf"))
        self.assertIn("0007", name)

    def test_dynamic_namer_normalized_key_resolution(self):
        allocated = set()
        fname = DynamicNamer.generate("Cert_{FullName}_{Student_ID}.pdf", {"Full Name": "Charles Darwin", "Student ID": "STU-100"}, 1, self.test_dir, allocated)
        self.assertIn("Charles_Darwin", fname)
        self.assertIn("STU-100", fname)

    # =========================================================================
    # 3. DATA READER TESTS
    # =========================================================================
    def test_data_reader_normalization(self):
        csv_path = os.path.join(self.test_dir, "test.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("FullName,Age,Score,Date\nAlice Cooper,25.0,98.5,2026-05-15\nBob,30.0,,2026-06-01\n")

        res = DataReader.read(csv_path)
        self.assertEqual(res["row_count"], 2)
        self.assertEqual(res["records"][0]["Age"], "25")
        self.assertEqual(res["records"][0]["Score"], "98.5")
        self.assertEqual(res["records"][0]["Date"], "2026-05-15")
        self.assertEqual(res["records"][1]["Score"], "")

    def test_data_reader_float_precision_preservation(self):
        self.assertEqual(DataReader._normalize_value(3.875), "3.875")
        self.assertEqual(DataReader._normalize_value(2026.0), "2026")
        self.assertEqual(DataReader._normalize_value(1250.5), "1250.5")

    def test_data_reader_csv_delimiters_and_encodings(self):
        # Semicolon separated, UTF-8 with BOM (European standard CSV)
        csv_semi = os.path.join(self.test_dir, "semi.csv")
        with open(csv_semi, "w", encoding="utf-8-sig") as f:
            f.write("ID;Employee_Name;Salary\n101;François Müller;45000\n102;Éléonore García;52000\n")

        res = DataReader.read(csv_semi)
        self.assertEqual(res["row_count"], 2)
        self.assertIn("Employee_Name", res["columns"])
        self.assertEqual(res["records"][0]["Employee_Name"], "François Müller")
        self.assertEqual(res["records"][1]["Employee_Name"], "Éléonore García")

    def test_data_reader_automapping(self):
        placeholders = ["FullName", "Student_ID", "GradDate"]
        columns = ["full_name", "Student ID", "Graduation_Date", "Extra"]
        mapping = DataReader.auto_map(placeholders, columns)
        self.assertEqual(mapping["FullName"], "full_name")
        self.assertEqual(mapping["Student_ID"], "Student ID")
        self.assertEqual(mapping["GradDate"], "Graduation_Date")

    def test_data_reader_nonexistent_file(self):
        with self.assertRaises(FileNotFoundError):
            DataReader.read(os.path.join(self.test_dir, "missing.xlsx"))

    # =========================================================================
    # 4. DOCUMENT MERGER TESTS
    # =========================================================================
    def test_document_merger_single_record(self):
        doc_path = os.path.join(self.test_dir, "merger_test.docx")
        doc = Document()
        doc.add_paragraph("Certificate of Completion awarded to {{RecipientName}} on {{IssueDate}}.")
        doc.save(doc_path)

        out_docx = os.path.join(self.test_dir, "merged_single.docx")
        mapping = {"RecipientName": "Name", "IssueDate": "Date"}
        record = {"Name": "Dr. Angela Merkel", "Date": "2026-10-25"}

        DocumentMerger.merge_single_record(doc_path, mapping, record, out_docx)
        self.assertTrue(os.path.exists(out_docx))

        res_doc = Document(out_docx)
        self.assertIn("Dr. Angela Merkel", res_doc.paragraphs[0].text)
        self.assertIn("2026-10-25", res_doc.paragraphs[0].text)
        self.assertNotIn("{{RecipientName}}", res_doc.paragraphs[0].text)

    def test_document_merger_cross_run_split_and_spaces(self):
        doc_path = os.path.join(self.test_dir, "split_runs.docx")
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Payment voucher: {{  ")
        p.add_run("Voucher")
        p.add_run("Code")
        p.add_run("  }} confirmed.")
        doc.save(doc_path)

        out_docx = os.path.join(self.test_dir, "merged_split.docx")
        mapping = {"VoucherCode": "VoucherCode"}
        record = {"VoucherCode": "VOUCH-8899"}

        DocumentMerger.merge_single_record(doc_path, mapping, record, out_docx)
        res_doc = Document(out_docx)
        self.assertIn("VOUCH-8899", res_doc.paragraphs[0].text)
        self.assertNotIn("VoucherCode", res_doc.paragraphs[0].text)

    def test_document_merger_recursive_placeholder_safety(self):
        doc_path = os.path.join(self.test_dir, "recursive_tpl.docx")
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("User: {{")
        p.add_run("Name")
        p.add_run("}}")
        doc.save(doc_path)

        out_docx = os.path.join(self.test_dir, "recursive_out.docx")
        DocumentMerger.merge_single_record(doc_path, {"Name": "Name"}, {"Name": "{{Name}}"}, out_docx)
        self.assertTrue(os.path.exists(out_docx))
        res_doc = Document(out_docx)
        self.assertEqual(res_doc.paragraphs[0].text, "User: {{Name}}")

    def test_document_merger_non_destructive(self):
        doc_path = os.path.join(self.test_dir, "immutable_source.docx")
        doc = Document()
        doc.add_paragraph("Master Template for {{User}}")
        doc.save(doc_path)

        with open(doc_path, "rb") as f:
            orig_hash = hashlib.sha256(f.read()).hexdigest()

        out_docx = os.path.join(self.test_dir, "merged_output.docx")
        DocumentMerger.merge_single_record(doc_path, {"User": "User"}, {"User": "Alice"}, out_docx)

        with open(doc_path, "rb") as f:
            post_hash = hashlib.sha256(f.read()).hexdigest()

        self.assertEqual(orig_hash, post_hash, "Original template was modified! Rule 2 violation!")

    # =========================================================================
    # 5. PDF CONVERTER TESTS
    # =========================================================================
    def test_pdf_converter_reportlab_standalone(self):
        doc_path = os.path.join(self.test_dir, "reportlab_test.docx")
        doc = Document()
        doc.add_paragraph("ReportLab Standalone Engine Test Header")
        doc.add_paragraph("Normal paragraph text with formatting.")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.rows[0].cells[0].paragraphs[0].text = "Header 1"
        tbl.rows[0].cells[1].paragraphs[0].text = "Header 2"
        tbl.rows[1].cells[0].paragraphs[0].text = "Data 1"
        tbl.rows[1].cells[1].paragraphs[0].text = "Data 2"
        doc.save(doc_path)

        out_pdf = os.path.join(self.test_dir, "reportlab_out.pdf")
        ok, engine = PDFConverter.convert_docx_to_pdf(doc_path, out_pdf)
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(out_pdf))
        self.assertGreater(os.path.getsize(out_pdf), 500)

        # PyMuPDF preview
        img = PDFConverter.render_preview_image(out_pdf)
        self.assertIsNotNone(img)
        self.assertGreater(img.width, 100)
        self.assertGreater(img.height, 100)

    def test_pdf_converter_nonexistent_file(self):
        ok, msg = PDFConverter.convert_docx_to_pdf(
            os.path.join(self.test_dir, "ghost.docx"),
            os.path.join(self.test_dir, "ghost.pdf")
        )
        self.assertFalse(ok)
        self.assertIn("does not exist", msg)

    # =========================================================================
    # 6. LICENSING & SECURITY ENGINE TESTS
    # =========================================================================
    def test_licensing_hwid_and_crypto(self):
        hwid = HardwareID.get()
        self.assertGreaterEqual(len(hwid), 16)

        sec = LicenseSecurity()
        state = sec.load_state()
        self.assertEqual(state["hwid"], hwid)

        # Crypto roundtrip
        payload = {"user": "test_agent", "counter": 42, "ts": 1234567890}
        enc = sec._encrypt(payload)
        dec = sec._decrypt(enc)
        self.assertEqual(payload, dec)

    def test_licensing_tamper_detection(self):
        sec = LicenseSecurity()
        payload = {"runs_remaining": 5}
        enc = sec._encrypt(payload)

        # Corrupt the ciphertext
        raw = bytearray(enc.encode("ascii"))
        raw[25] = ord('X') if raw[25] != ord('X') else ord('Y')
        tampered_enc = raw.decode("ascii", errors="ignore")

        dec = sec._decrypt(tampered_enc)
        self.assertIsNone(dec, "Corrupted ciphertext was accepted without MAC failure!")

    def test_licensing_offline_studio_key(self):
        sec = LicenseSecurity()
        hwid = HardwareID.get()
        offline_key = sec.generate_offline_studio_key(hwid, tier="lifetime")
        self.assertTrue(offline_key.startswith("DMP-LIFETIME-"))

        valid, tier = sec.verify_offline_studio_key(offline_key)
        self.assertTrue(valid)
        self.assertEqual(tier, "lifetime")

        # Wrong key
        bad_valid, _ = sec.verify_offline_studio_key("DMP-LIFETIME-INVALID-KEY-12345")
        self.assertFalse(bad_valid)

    def test_data_reader_duplicate_columns(self):
        csv_path = os.path.join(self.test_dir, "dup_cols.csv")
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write("Email,Email,Notes\njohn@example.com,alternate@example.com,VIP\n")
        res = DataReader.read(csv_path)
        self.assertEqual(res["row_count"], 1)
        self.assertEqual(len(res["columns"]), 3)
        self.assertIn("Email", res["columns"])
        self.assertTrue("Email.1" in res["columns"] or "Email_2" in res["columns"])
        self.assertEqual(res["records"][0]["Email"], "john@example.com")

    def test_data_reader_automapping_unrelated_substrings_no_false_positive(self):
        placeholders = ["ID", "Date", "No"]
        columns = ["President", "Candidate", "Economics", "Student_ID", "Graduation_Date"]
        mapping = DataReader.auto_map(placeholders, columns)
        self.assertEqual(mapping["ID"], "Student_ID")
        self.assertEqual(mapping["Date"], "Graduation_Date")
        self.assertIsNone(mapping["No"])

    def test_data_reader_utf8_split_multibyte_stream(self):
        csv_path = os.path.join(self.test_dir, "utf8_boundary.csv")
        prefix = "a" * 8190
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write(f"Header,Comment\n{prefix},Nguyễn Hoàng Ánh\n")
        res = DataReader.read(csv_path)
        self.assertEqual(res["row_count"], 1)
        self.assertIn("Nguyễn Hoàng Ánh", res["records"][0]["Comment"])

    # =========================================================================
    # 7. LEGAL SHIELD ENGINE TESTS
    # =========================================================================
    def test_legal_shield(self):
        terms = LegalShield.read_full_terms()
        self.assertIn("UCC § 2-316", terms)
        self.assertIn("UCC § 2-719", terms)
        self.assertIn("$1.00 USD", terms)
        self.assertIn("Microsoft", terms)

        # Consent record test
        ok = LegalShield.record_consent()
        self.assertTrue(ok)
        self.assertTrue(LegalShield.has_user_consented())

    # =========================================================================
    # 8. ADVANCED MERGEFIELD & DATA READER ENHANCEMENT TESTS
    # =========================================================================
    def test_document_merger_w_fldsimple_unlink(self):
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        doc_path = os.path.join(self.test_dir, "fldsimple_in.docx")
        out_path = os.path.join(self.test_dir, "fldsimple_out.docx")
        doc = Document()
        p = doc.add_paragraph()
        fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="MERGEFIELD RecipientName"><w:r><w:rPr><w:b/></w:rPr><w:t>«RecipientName»</w:t></w:r></w:fldSimple>')
        p._p.append(fld)
        doc.save(doc_path)

        DocumentMerger.merge_single_record(doc_path, {"RecipientName": "NameCol"}, {"NameCol": "Sarah Connor"}, out_path)
        res_doc = Document(out_path)
        self.assertEqual(res_doc.paragraphs[0].text, "Sarah Connor")
        self.assertEqual(res_doc.paragraphs[0].runs[0].text, "Sarah Connor")
        self.assertTrue(res_doc.paragraphs[0].runs[0].bold)
        self.assertEqual(len(res_doc.paragraphs[0]._p.xpath('.//w:fldSimple')), 0)

    def test_document_merger_complex_field_unlink(self):
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        doc_path = os.path.join(self.test_dir, "complex_in.docx")
        out_path = os.path.join(self.test_dir, "complex_out.docx")
        doc = Document()
        xml_complex = rf'''<w:p {nsdecls("w")}>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> MERGEFIELD AwardTitle \* MERGEFORMAT </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>«AwardTitle»</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>'''
        doc._body._element.append(parse_xml(xml_complex))
        doc.save(doc_path)

        DocumentMerger.merge_single_record(doc_path, {"AwardTitle": "AwardCol"}, {"AwardCol": "Excellence in AI"}, out_path)
        res_doc = Document(out_path)
        self.assertEqual(res_doc.paragraphs[0].text, "Excellence in AI")
        self.assertEqual(len(res_doc.paragraphs[0]._p.xpath('.//w:fldChar')), 0)
        self.assertEqual(len(res_doc.paragraphs[0]._p.xpath('.//w:instrText')), 0)

    def test_data_reader_large_number_no_scientific_notation(self):
        val1 = DataReader._normalize_value(1234567.89)
        val2 = DataReader._normalize_value(1000000.0)
        val3 = DataReader._normalize_value(18005550199.0)
        self.assertEqual(val1, "1234567.89")
        self.assertEqual(val2, "1000000")
        self.assertEqual(val3, "18005550199")

    def test_dynamic_namer_date_today_timestamp(self):
        import datetime
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        allocated = set()
        res = DynamicNamer.generate("Invoice_{FullName}_{Today}", {"FullName": "Wayne Enterprises"}, 1, self.test_dir, allocated)
        self.assertEqual(res, f"Invoice_Wayne_Enterprises_{today_str}.pdf")

    def test_textbox_placeholders_detection_and_merging(self):
        from docx.oxml import parse_xml
        doc_path = os.path.join(self.test_dir, "test_txbx.docx")
        out_path = os.path.join(self.test_dir, "test_txbx_out.docx")
        doc = Document()
        doc.add_paragraph("Top body text")
        tb_xml = """<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:drawing>
                <w:txbxContent>
                    <w:p>
                        <w:r><w:t>{{InsideTextBox}}</w:t></w:r>
                    </w:p>
                </w:txbxContent>
            </w:drawing>
        </w:r>"""
        doc.paragraphs[0]._p.append(parse_xml(tb_xml))
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertIn("InsideTextBox", res["placeholders"])

        DocumentMerger.merge_single_record(
            doc_path,
            {"InsideTextBox": "TextBoxField"},
            {"TextBoxField": "Replaced Inside Text Box"},
            out_path
        )
        import zipfile
        with zipfile.ZipFile(out_path) as z:
            xml_str = z.read("word/document.xml").decode("utf-8")
            self.assertIn("Replaced Inside Text Box", xml_str)
            self.assertNotIn("{{InsideTextBox}}", xml_str)

    def test_nested_table_placeholders_detection_and_merging(self):
        doc_path = os.path.join(self.test_dir, "test_nested.docx")
        out_path = os.path.join(self.test_dir, "test_nested_out.docx")
        doc = Document()
        outer_table = doc.add_table(rows=1, cols=1)
        cell = outer_table.cell(0, 0)
        inner_table = cell.add_table(rows=1, cols=1)
        inner_table.cell(0, 0).paragraphs[0].text = "Nested {{NestedField}}"
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertIn("NestedField", res["placeholders"])

        DocumentMerger.merge_single_record(
            doc_path,
            {"NestedField": "SubCol"},
            {"SubCol": "SubValue"},
            out_path
        )
        import zipfile
        with zipfile.ZipFile(out_path) as z:
            xml_str = z.read("word/document.xml").decode("utf-8")
            self.assertIn("SubValue", xml_str)
            self.assertNotIn("{{NestedField}}", xml_str)

    def test_multilingual_vietnamese_tag_inspection_automap_and_naming(self):
        doc_path = os.path.join(self.test_dir, "test_vn_tags.docx")
        doc = Document()
        doc.add_paragraph("Chứng chỉ cho: {{Họ và Tên}}")
        doc.add_paragraph("Ngày sinh: {Ngày_Sinh}")
        doc.add_paragraph("Nơi cư trú: «Địa_Chỉ»")
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertIn("Họ và Tên", res["placeholders"])
        self.assertIn("Ngày_Sinh", res["placeholders"])
        self.assertIn("Địa_Chỉ", res["placeholders"])

        columns = ["Họ và Tên", "Ngày Sinh", "Địa Chỉ", "Mã Số"]
        mapping = DataReader.auto_map(res["placeholders"], columns)
        self.assertEqual(mapping["Họ và Tên"], "Họ và Tên")
        self.assertEqual(mapping["Ngày_Sinh"], "Ngày Sinh")
        self.assertEqual(mapping["Địa_Chỉ"], "Địa Chỉ")

        record = {"Họ và Tên": "Nguyễn Văn An", "Mã Số": "NV-001"}
        allocated = set()
        fname = DynamicNamer.generate("ChungChi_{Họ và Tên}_{Mã Số}.pdf", record, 1, self.test_dir, allocated)
        self.assertIn("Nguyễn_Văn_An", fname)

    def test_mergefield_with_quotes_and_spaces(self):
        from docx.oxml import parse_xml
        doc_path = os.path.join(self.test_dir, "test_mf_spaces.docx")
        doc = Document()
        mf_xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldSimple w:instr="MERGEFIELD &quot;Full Name&quot; \\* MERGEFORMAT"><w:r><w:t>«Full Name»</w:t></w:r></w:fldSimple></w:p>'
        doc._element.body.append(parse_xml(mf_xml))
        doc.save(doc_path)

        res = TemplateInspector.inspect(doc_path)
        self.assertIn("Full Name", res["placeholders"])

    def test_mergefield_complex_empty_display_replacement(self):
        from docx.oxml import parse_xml
        doc_path = os.path.join(self.test_dir, "test_empty_disp.docx")
        out_path = os.path.join(self.test_dir, "test_empty_disp_out.docx")
        doc = Document()
        fld_xml = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText>MERGEFIELD EmptyDisplay</w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>"""
        doc._element.body.append(parse_xml(fld_xml))
        doc.save(doc_path)

        DocumentMerger.merge_single_record(
            doc_path,
            {"EmptyDisplay": "EmptyDisplay"},
            {"EmptyDisplay": "REPLACED_EMPTY_FIELD"},
            out_path
        )
        res_doc = Document(out_path)
        self.assertEqual(res_doc.paragraphs[0].text, "REPLACED_EMPTY_FIELD")

    def test_csv_utf16_powershell_excel_format(self):
        csv_path = os.path.join(self.test_dir, "utf16_roster.csv")
        content = "FullName,Department\nSarah Connor,Security\nJohn Doe,Operations\n"
        with open(csv_path, "w", encoding="utf-16") as f:
            f.write(content)

        data = DataReader.read(csv_path)
        self.assertEqual(data["row_count"], 2)
        self.assertIn("FullName", data["columns"])
        self.assertEqual(data["records"][0]["FullName"], "Sarah Connor")

    def test_excel_temp_lock_file_rejection(self):
        temp_lock = os.path.join(self.test_dir, "~$Payroll.xlsx")
        with open(temp_lock, "w") as f:
            f.write("fake lock")

        with self.assertRaises(ValueError) as ctx:
            DataReader.read(temp_lock)
        self.assertIn("temporary Microsoft Excel lock file", str(ctx.exception))

    def test_dual_vault_anti_rollback_protection(self):
        sec = LicenseSecurity()
        sec.reset_trial()
        # Consume 3 runs
        sec.consume_batch_export_run()
        sec.consume_batch_export_run()
        sec.consume_batch_export_run()
        st = sec.get_status()
        self.assertEqual(st["runs_used"], 3)

        # Simulate tampering by rolling back file vault to state with runs_used=0
        state = sec.load_state()
        old_state = dict(state)
        old_state["runs_used"] = 0
        old_state["signature"] = sec._sign(old_state)
        enc = sec._encrypt(old_state)
        sec._write_file(enc)

        # Reloading must NOT reset to 0, because registry has 3!
        reloaded = sec.load_state()
        self.assertEqual(reloaded["runs_used"], 3)

    def test_dynamic_namer_double_braces_formula(self):
        # Users typing {{FullName}} and {{Index}} in formula must produce clean names
        allocated = set()
        record = {"FullName": "Elizabeth Bennet", "ID": "9876"}
        res = DynamicNamer.generate("Award_{{FullName}}_{{Index}}", record, 5, self.test_dir, allocated)
        self.assertEqual(res, "Award_Elizabeth_Bennet_0005.pdf")
        self.assertNotIn("{", res)
        self.assertNotIn("}", res)

    def test_pdf_converter_sdt_and_textbox_and_headers(self):
        from docx.oxml import parse_xml
        doc = Document()
        sec = doc.sections[0]
        sec.header.paragraphs[0].text = "ACADEMY HEADER 2026"
        sec.footer.paragraphs[0].text = "CONFIDENTIAL FOOTER"

        # Add SDT content control
        sdt = parse_xml('<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sdtContent><w:p><w:r><w:t>Content in SDT Block</w:t></w:r></w:p></w:sdtContent></w:sdt>')
        doc.element.body.insert(0, sdt)

        # Add Textbox inside paragraph
        p = doc.add_paragraph()
        tb = parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:drawing><w:txbxContent><w:p><w:r><w:t>Floating Textbox Content</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r>')
        p._p.append(tb)

        docx_path = os.path.join(self.test_dir, "rich_elements.docx")
        pdf_path = os.path.join(self.test_dir, "rich_elements.pdf")
        doc.save(docx_path)

        ok, eng = PDFConverter.convert_docx_to_pdf(docx_path, pdf_path)
        self.assertTrue(ok)
        self.assertTrue(os.path.exists(pdf_path))

        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        pdf_doc = fitz.open(pdf_path)
        full_text = pdf_doc[0].get_text()
        self.assertIn("ACADEMY HEADER 2026", full_text)
        self.assertIn("CONFIDENTIAL FOOTER", full_text)
        self.assertIn("Content in SDT Block", full_text)
        self.assertIn("Floating Textbox Content", full_text)

    def test_data_reader_empty_csv_and_sheet_resilience(self):
        empty_csv = os.path.join(self.test_dir, "zero_byte.csv")
        with open(empty_csv, "w", encoding="utf-8") as f:
            pass
        res = DataReader.read(empty_csv)
        self.assertEqual(res["row_count"], 0)
        self.assertEqual(res["columns"], [])
        self.assertEqual(res["records"], [])

    def test_dynamic_namer_control_characters_sanitization(self):
        # Control characters (\x00-\x1f, \x7f) must be sanitized to prevent Windows filesystem errors
        name_with_ctrl = "Document\x07_Special\x1f_Test\x7f"
        sanitized = DynamicNamer.sanitize(name_with_ctrl)
        self.assertNotIn("\x07", sanitized)
        self.assertNotIn("\x1f", sanitized)
        self.assertNotIn("\x7f", sanitized)
        out_path = os.path.join(self.test_dir, f"{sanitized}.pdf")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("test")
        self.assertTrue(os.path.exists(out_path))

    def test_demo_templates_character_integrity(self):
        # Verify demo templates contain no corrupted replacement characters (\ufffd)
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        demo_files = [
            os.path.join(base_dir, "demo_data", "Certificate_Template.docx"),
            os.path.join(base_dir, "demo_data", "Payslip_Template.docx"),
        ]
        for df in demo_files:
            if os.path.exists(df):
                d = Document(df)
                full = " ".join(p.text for p in d.paragraphs)
                self.assertNotIn("\ufffd", full, f"Corrupted character found in {df}")


if __name__ == "__main__":
    unittest.main()
