"""
DocuMergePDF - Primary Merge Studio Workspace View
Features Dual File Dropzones, Smart Field Auto-Mapping, Dynamic Filename Formula Builder,
Background Batch Worker with Progress Bar, and 1-Click Row 1 Test Generation.
"""

import os
import re
import sys
import time
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from typing import Optional, Dict, List, Any

from . import styles
from core.template_inspector import TemplateInspector
from core.data_reader import DataReader
from core.dynamic_namer import DynamicNamer
from core.document_merger import DocumentMerger
from core.pdf_converter import PDFConverter
from core.licensing import LicenseSecurity
from core.legal_shield import LegalShield


class MergeView(ctk.CTkScrollableFrame):
    """Primary document merge workflow canvas."""

    def __init__(self, master, app_instance):
        super().__init__(master, fg_color=styles.CANVAS_BG)
        self.app = app_instance
        self.license_engine: LicenseSecurity = app_instance.license_engine

        # State variables
        self.docx_path: Optional[str] = None
        self.data_path: Optional[str] = None
        self.output_dir: Optional[str] = None
        
        self.detected_tags: List[str] = []
        self.data_columns: List[str] = []
        self.data_records: List[Dict[str, str]] = []
        self.data_sheets: List[str] = []
        self.active_sheet: Optional[str] = None
        self.mapping_vars: Dict[str, ctk.StringVar] = {}

        self.is_merging = False
        self.stop_requested = False

        self._build_ui()
        self._load_default_demo_if_available()

    def _build_ui(self):
        # 1. Source Files Dual Card
        card_sources = ctk.CTkFrame(self, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        card_sources.pack(fill="x", padx=20, pady=(15, 10))

        src_inner = ctk.CTkFrame(card_sources, fg_color="transparent")
        src_inner.pack(fill="x", padx=18, pady=16)

        ctk.CTkLabel(src_inner, text="📁 Source Documents & Recipient Data", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY).pack(anchor="w")

        # Dual dropzones container
        dz_box = ctk.CTkFrame(src_inner, fg_color="transparent")
        dz_box.pack(fill="x", pady=(10, 0))

        # Dropzone 1: Word Template
        self.card_docx = ctk.CTkFrame(dz_box, fg_color=styles.CANVAS_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=10)
        self.card_docx.pack(side="left", fill="both", expand=True, padx=(0, 6))

        docx_inner = ctk.CTkFrame(self.card_docx, fg_color="transparent")
        docx_inner.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(docx_inner, text="📄 Word Template (.docx)", font=styles.FONT_CARD_TITLE, text_color=styles.PRIMARY_ACCENT).pack(anchor="w")
        self.lbl_docx_file = ctk.CTkLabel(docx_inner, text="Drag & drop .docx template or browse", font=styles.FONT_BODY, text_color=styles.TEXT_SECONDARY)
        self.lbl_docx_file.pack(anchor="w", pady=(4, 6))
        self.lbl_docx_meta = ctk.CTkLabel(docx_inner, text="Tags detected: None", font=styles.FONT_SMALL, text_color=styles.TEXT_MUTED)
        self.lbl_docx_meta.pack(anchor="w")

        ctk.CTkButton(
            docx_inner,
            text="Browse Template...",
            font=styles.FONT_SMALL,
            height=28,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._browse_docx
        ).pack(anchor="w", pady=(8, 0))

        # Dropzone 2: Data Roster
        self.card_data = ctk.CTkFrame(dz_box, fg_color=styles.CANVAS_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=10)
        self.card_data.pack(side="right", fill="both", expand=True, padx=(6, 0))

        data_inner = ctk.CTkFrame(self.card_data, fg_color="transparent")
        data_inner.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(data_inner, text="📊 Recipient Data (.xlsx / .csv)", font=styles.FONT_CARD_TITLE, text_color=styles.SUCCESS_TEXT).pack(anchor="w")
        self.lbl_data_file = ctk.CTkLabel(data_inner, text="Drag & drop .xlsx / .csv roster or browse", font=styles.FONT_BODY, text_color=styles.TEXT_SECONDARY)
        self.lbl_data_file.pack(anchor="w", pady=(4, 6))
        self.lbl_data_meta = ctk.CTkLabel(data_inner, text="Rows: 0 • Sheet: N/A", font=styles.FONT_SMALL, text_color=styles.TEXT_MUTED)
        self.lbl_data_meta.pack(anchor="w")

        data_btn_row = ctk.CTkFrame(data_inner, fg_color="transparent")
        data_btn_row.pack(fill="x", pady=(8, 0))

        ctk.CTkButton(
            data_btn_row,
            text="Browse Spreadsheet...",
            font=styles.FONT_SMALL,
            height=28,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._browse_data
        ).pack(side="left")

        self.opt_sheet = ctk.CTkOptionMenu(
            data_btn_row,
            values=["Sheet1"],
            width=110,
            height=28,
            font=styles.FONT_SMALL,
            fg_color=styles.SECONDARY_BTN_BG,
            button_color=styles.CARD_BORDER,
            text_color=styles.TEXT_PRIMARY,
            command=self._on_sheet_changed
        )
        self.opt_sheet.pack(side="right")

        # 2. Field Mapping Card
        self.card_mapping = ctk.CTkFrame(self, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        self.card_mapping.pack(fill="x", padx=20, pady=10)

        map_inner = ctk.CTkFrame(self.card_mapping, fg_color="transparent")
        map_inner.pack(fill="x", padx=18, pady=16)

        map_header = ctk.CTkFrame(map_inner, fg_color="transparent")
        map_header.pack(fill="x")
        ctk.CTkLabel(map_header, text="🔗 Field Mapping & Variable Substitution", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY).pack(side="left")
        self.lbl_match_pill = ctk.CTkLabel(map_header, text="0 / 0 MATCHED", font=styles.FONT_BADGE, fg_color=styles.SOFT_PILL_BG, text_color=styles.SOFT_PILL_TEXT, corner_radius=6)
        self.lbl_match_pill.pack(side="right", padx=5)

        self.mapping_rows_container = ctk.CTkFrame(map_inner, fg_color="transparent")
        self.mapping_rows_container.pack(fill="x", pady=(10, 0))

        # Placeholder message
        self.lbl_empty_mapping = ctk.CTkLabel(
            self.mapping_rows_container,
            text="Load a Word Template (.docx) and Excel file above to configure field mappings.",
            font=styles.FONT_BODY,
            text_color=styles.TEXT_SECONDARY
        )
        self.lbl_empty_mapping.pack(pady=10)

        # 3. Dynamic Filename Formula & Output Destination Card
        card_naming = ctk.CTkFrame(self, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        card_naming.pack(fill="x", padx=20, pady=10)

        naming_inner = ctk.CTkFrame(card_naming, fg_color="transparent")
        naming_inner.pack(fill="x", padx=18, pady=16)

        ctk.CTkLabel(naming_inner, text="🏷️ Output Filenaming Formula & Destination", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY).pack(anchor="w")

        formula_box = ctk.CTkFrame(naming_inner, fg_color="transparent")
        formula_box.pack(fill="x", pady=(8, 4))

        ctk.CTkLabel(formula_box, text="Formula Pattern:", font=styles.FONT_BODY_BOLD, text_color=styles.TEXT_PRIMARY).pack(side="left")
        self.entry_formula = ctk.CTkEntry(
            formula_box,
            font=styles.FONT_BODY,
            height=32,
            border_color=styles.CARD_BORDER
        )
        self.entry_formula.pack(side="left", fill="x", expand=True, padx=(10, 10))
        self.entry_formula.insert(0, "Certificate_{FullName}_{CertificateID}.pdf")

        # Variable insertion pills container
        self.pills_container = ctk.CTkFrame(naming_inner, fg_color="transparent")
        self.pills_container.pack(fill="x", pady=(2, 8))

        # Output folder row
        out_row = ctk.CTkFrame(naming_inner, fg_color="transparent")
        out_row.pack(fill="x", pady=(4, 0))

        ctk.CTkLabel(out_row, text="Output Directory:", font=styles.FONT_BODY_BOLD, text_color=styles.TEXT_PRIMARY).pack(side="left")
        self.entry_out_dir = ctk.CTkEntry(
            out_row,
            font=styles.FONT_BODY,
            height=32,
            border_color=styles.CARD_BORDER
        )
        self.entry_out_dir.pack(side="left", fill="x", expand=True, padx=(10, 10))

        ctk.CTkButton(
            out_row,
            text="Choose Folder...",
            font=styles.FONT_SMALL,
            height=32,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._browse_output_dir
        ).pack(side="right")

        # 4. Action Card (Buttons & Progress)
        card_action = ctk.CTkFrame(self, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        card_action.pack(fill="x", padx=20, pady=(10, 25))

        act_inner = ctk.CTkFrame(card_action, fg_color="transparent")
        act_inner.pack(fill="x", padx=18, pady=16)

        btn_row = ctk.CTkFrame(act_inner, fg_color="transparent")
        btn_row.pack(fill="x")

        # Free Row 1 Test Button
        self.btn_test_row1 = ctk.CTkButton(
            btn_row,
            text="👁️ Generate Row 1 Test (Free Forever)",
            font=styles.FONT_BODY_BOLD,
            height=40,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._generate_row1_test
        )
        self.btn_test_row1.pack(side="left", fill="x", expand=True, padx=(0, 8))

        # Full Batch Export Button
        self.btn_batch_merge = ctk.CTkButton(
            btn_row,
            text="🚀 Generate All Individual PDFs (Batch)",
            font=styles.FONT_HEADING,
            height=40,
            fg_color=styles.PRIMARY_ACCENT,
            hover_color=styles.PRIMARY_HOVER,
            text_color="#FFFFFF",
            command=self._start_batch_merge
        )
        self.btn_batch_merge.pack(side="right", fill="x", expand=True, padx=(8, 0))

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(act_inner, height=10, corner_radius=5)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=(14, 6))

        status_line = ctk.CTkFrame(act_inner, fg_color="transparent")
        status_line.pack(fill="x")

        self.lbl_progress_status = ctk.CTkLabel(status_line, text="Ready to merge documents.", font=styles.FONT_SMALL, text_color=styles.TEXT_SECONDARY)
        self.lbl_progress_status.pack(side="left")

        self.btn_open_folder = ctk.CTkButton(
            status_line,
            text="📂 Open Output Folder",
            font=styles.FONT_SMALL,
            height=26,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._open_output_folder
        )
        self.btn_open_folder.pack(side="right")
        self.btn_open_folder.configure(state="disabled")

    def _load_default_demo_if_available(self):
        """Loads demo data automatically if present for a seamless first-time experience."""
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        demo_docx = os.path.join(base_dir, "demo_data", "Certificate_Template.docx")
        demo_xlsx = os.path.join(base_dir, "demo_data", "Student_Roster_100.xlsx")

        if os.path.exists(demo_docx):
            self.load_template(demo_docx)
        if os.path.exists(demo_xlsx):
            self.load_data(demo_xlsx)

    def load_template(self, docx_path: str):
        try:
            insp = TemplateInspector.inspect(docx_path)
            self.docx_path = docx_path
            self.detected_tags = insp["placeholders"]

            self.lbl_docx_file.configure(text=f"✓ Loaded: {os.path.basename(docx_path)}")
            tags_str = ", ".join([f"{{{{{t}}}}}" for t in self.detected_tags[:4]])
            if len(self.detected_tags) > 4:
                tags_str += f" (+{len(self.detected_tags) - 4} more)"
            self.lbl_docx_meta.configure(text=f"{len(self.detected_tags)} Tags: {tags_str}")
            self.card_docx.configure(border_color=styles.CARD_BORDER_ACTIVE)

            # Set default output dir if not chosen
            if not self.output_dir:
                parent = os.path.dirname(docx_path)
                stem = os.path.splitext(os.path.basename(docx_path))[0]
                self.output_dir = os.path.join(parent, f"{stem}_Merged_PDFs")
                self.entry_out_dir.delete(0, "end")
                self.entry_out_dir.insert(0, self.output_dir)

            self._update_mapping_table()
            self._update_pills()
        except Exception as e:
            messagebox.showerror("Template Error", f"Failed to inspect Word template:\n{str(e)}")

    def load_data(self, data_path: str, sheet_name: Optional[str] = None):
        try:
            data = DataReader.read(data_path, sheet_name=sheet_name)
            self.data_path = data_path
            self.data_columns = data["columns"]
            self.data_records = data["records"]
            self.data_sheets = data["sheets"]
            self.active_sheet = data["active_sheet"]

            self.lbl_data_file.configure(text=f"✓ Loaded: {os.path.basename(data_path)}")
            self.lbl_data_meta.configure(text=f"Total: {len(self.data_records)} Rows • {len(self.data_columns)} Columns")
            self.card_data.configure(border_color=styles.SUCCESS_ACCENT)

            # Update sheet dropdown
            self.opt_sheet.configure(values=self.data_sheets)
            self.opt_sheet.set(self.active_sheet)

            self._update_mapping_table()
            self._update_pills()
        except Exception as e:
            messagebox.showerror("Spreadsheet Error", f"Failed to read spreadsheet:\n{str(e)}")

    def _on_sheet_changed(self, sheet: str):
        if self.data_path:
            self.load_data(self.data_path, sheet_name=sheet)

    def _update_mapping_table(self):
        for widget in self.mapping_rows_container.winfo_children():
            widget.destroy()

        if not self.detected_tags:
            ctk.CTkLabel(self.mapping_rows_container, text="No placeholders detected in Word template.", font=styles.FONT_BODY, text_color=styles.TEXT_SECONDARY).pack(pady=10)
            return

        cols = ["(Leave Blank)"] + self.data_columns
        auto_map = DataReader.auto_map(self.detected_tags, self.data_columns)
        matched_count = 0

        self.mapping_vars = {}

        for tag in self.detected_tags:
            row_frame = ctk.CTkFrame(self.mapping_rows_container, fg_color=styles.CANVAS_BG, corner_radius=6, border_color=styles.CARD_BORDER, border_width=1)
            row_frame.pack(fill="x", pady=3)

            r_inner = ctk.CTkFrame(row_frame, fg_color="transparent")
            r_inner.pack(fill="x", padx=10, pady=6)

            # Tag Pill
            ctk.CTkLabel(
                r_inner,
                text=f"{{{{{tag}}}}}",
                font=styles.FONT_BADGE,
                fg_color=styles.SOFT_PILL_BG,
                text_color=styles.SOFT_PILL_TEXT,
                corner_radius=6,
                width=160
            ).pack(side="left")

            ctk.CTkLabel(r_inner, text="➔  substitute with: ", font=styles.FONT_SMALL, text_color=styles.TEXT_SECONDARY).pack(side="left", padx=10)

            # Dropdown menu
            var = ctk.StringVar()
            best_match = auto_map.get(tag)
            if best_match and best_match in cols:
                var.set(best_match)
                matched_count += 1
            else:
                var.set("(Leave Blank)")
            self.mapping_vars[tag] = var

            opt = ctk.CTkOptionMenu(
                r_inner,
                variable=var,
                values=cols,
                font=styles.FONT_SMALL,
                height=26,
                width=180,
                fg_color=styles.SECONDARY_BTN_BG,
                button_color=styles.CARD_BORDER,
                text_color=styles.TEXT_PRIMARY
            )
            opt.pack(side="left")

            # Sample row preview
            sample_val = ""
            if self.data_records and best_match and best_match in self.data_records[0]:
                sample_val = f"Sample: '{self.data_records[0][best_match]}'"
            
            lbl_sample = ctk.CTkLabel(r_inner, text=sample_val, font=styles.FONT_SMALL, text_color=styles.TEXT_MUTED)
            lbl_sample.pack(side="left", padx=15)

        self.lbl_match_pill.configure(text=f"{matched_count} / {len(self.detected_tags)} MATCHED")

    def _update_pills(self):
        for w in self.pills_container.winfo_children():
            w.destroy()

        ctk.CTkLabel(self.pills_container, text="Click to append into formula:", font=styles.FONT_SMALL, text_color=styles.TEXT_SECONDARY).pack(side="left", padx=(0, 8))

        # Standard tokens
        for token in ["{Index}", "{Row}"]:
            ctk.CTkButton(
                self.pills_container,
                text=f"+ {token}",
                font=styles.FONT_SMALL,
                height=24,
                width=60,
                fg_color=styles.SECONDARY_BTN_BG,
                hover_color=styles.SECONDARY_BTN_HOVER,
                text_color=styles.TEXT_PRIMARY,
                command=lambda t=token: self._insert_formula_token(t)
            ).pack(side="left", padx=2)

        # Available record columns
        for col in self.data_columns[:6]:
            tok = f"{{{col}}}"
            ctk.CTkButton(
                self.pills_container,
                text=f"+ {tok}",
                font=styles.FONT_SMALL,
                height=24,
                width=70,
                fg_color=styles.SOFT_PILL_BG,
                hover_color=styles.CARD_BORDER,
                text_color=styles.SOFT_PILL_TEXT,
                command=lambda t=tok: self._insert_formula_token(t)
            ).pack(side="left", padx=2)

    def _insert_formula_token(self, token: str):
        cur = self.entry_formula.get().strip()
        if cur.lower().endswith(".pdf"):
            cur = cur[:-4]
        token_formatted = token if token.startswith("{") else f"{{{token}}}"
        new_val = f"{cur}_{token_formatted}.pdf" if cur else f"{token_formatted}.pdf"
        self.entry_formula.delete(0, "end")
        self.entry_formula.insert(0, new_val)

    def _browse_docx(self):
        path = filedialog.askopenfilename(
            title="Select Word Template",
            filetypes=[("Word Document (*.docx)", "*.docx"), ("All Files (*.*)", "*.*")]
        )
        if path:
            self.load_template(path)

    def _browse_data(self):
        path = filedialog.askopenfilename(
            title="Select Spreadsheet Roster",
            filetypes=[("Excel & CSV (*.xlsx;*.xls;*.csv)", "*.xlsx;*.xls;*.csv"), ("Excel (*.xlsx)", "*.xlsx"), ("CSV (*.csv)", "*.csv"), ("All Files (*.*)", "*.*")]
        )
        if path:
            self.load_data(path)

    def _browse_output_dir(self):
        folder = filedialog.askdirectory(title="Select Destination Folder for Generated PDFs")
        if folder:
            self.output_dir = folder
            self.entry_out_dir.delete(0, "end")
            self.entry_out_dir.insert(0, folder)

    def _get_active_mapping(self) -> Dict[str, Optional[str]]:
        mapping = {}
        for tag, var in self.mapping_vars.items():
            val = var.get()
            mapping[tag] = None if val == "(Leave Blank)" else val
        return mapping

    # =========================================================================
    # ROW 1 TEST GENERATION (100% FREE FOREVER)
    # =========================================================================
    def _generate_row1_test(self):
        if not LegalShield.has_user_consented():
            if hasattr(self.app, "_show_consent_dialog"):
                self.app._show_consent_dialog()
            return

        if not self.docx_path:
            messagebox.showwarning("Missing Template", "Please select a Word template (.docx) first.")
            return

        if not self.data_path:
            messagebox.showwarning("Missing Spreadsheet", "Please select a recipient spreadsheet (.xlsx or .csv) first.")
            return

        if not self.data_records:
            messagebox.showwarning("No Data Records", "The selected spreadsheet does not contain any recipient rows.")
            return

        self.btn_test_row1.configure(state="disabled")
        self.lbl_progress_status.configure(text="Generating Row 1 Test PDF...")

        t = threading.Thread(target=self._generate_row1_test_worker, daemon=True)
        t.start()

    def _generate_row1_test_worker(self):
        tmp_docx = None
        try:
            out_folder = self.entry_out_dir.get().strip() or os.path.dirname(self.docx_path)
            os.makedirs(out_folder, exist_ok=True)

            formula = self.entry_formula.get().strip() or "Sample_Row1_{FullName}.pdf"
            mapping = self._get_active_mapping()
            row_0 = self.data_records[0]

            # Combine raw spreadsheet keys with mapped template tags for formula resolution
            effective_row_0 = dict(row_0)
            for tag, col in mapping.items():
                if col and col in row_0:
                    effective_row_0[tag] = row_0[col]

            import tempfile
            tmp_dir = tempfile.gettempdir()
            allocated = set()
            out_filename = DynamicNamer.generate(formula, effective_row_0, index=1, output_dir=out_folder, allocated_names=allocated)
            out_pdf = os.path.join(out_folder, out_filename)
            tmp_docx = os.path.join(tmp_dir, f"_dmp_test_{os.getpid()}_1.docx")

            # 1. Merge DOCX
            DocumentMerger.merge_single_record(self.docx_path, mapping, row_0, tmp_docx)
            # 2. Convert to PDF
            ok, engine = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)

            def _on_success():
                self.btn_test_row1.configure(state="normal")
                if ok and os.path.exists(out_pdf):
                    self.lbl_progress_status.configure(text=f"✓ Row 1 Test generated: {out_filename} ({engine})")
                    self.btn_open_folder.configure(state="normal")
                    # Immediately load into preview view and switch tab
                    if hasattr(self.app, "preview_view"):
                        self.app.preview_view.load_pdf(out_pdf)
                    if hasattr(self.app, "switch_tab"):
                        self.app.switch_tab("preview")
                else:
                    self.lbl_progress_status.configure(text="Conversion failed.")
                    messagebox.showerror("Conversion Failed", f"PDF converter failed: {engine}")

            self.after(0, _on_success)
        except Exception as e:
            def _on_err(err_str=str(e)):
                self.btn_test_row1.configure(state="normal")
                self.lbl_progress_status.configure(text="Error generating test PDF.")
                messagebox.showerror("Error", f"Failed generating test PDF:\n{err_str}")
            self.after(0, _on_err)
        finally:
            if tmp_docx and os.path.exists(tmp_docx):
                try:
                    os.remove(tmp_docx)
                except Exception:
                    pass

    # =========================================================================
    # FULL BATCH MERGE PIPELINE WITH CANCELLATION SUPPORT
    # =========================================================================
    def _cancel_batch_merge(self):
        if self.is_merging:
            self.stop_requested = True
            self.lbl_progress_status.configure(text="Cancelling batch merge after current file...")
            self.btn_batch_merge.configure(state="disabled", text="Cancelling...")

    def _start_batch_merge(self):
        if self.is_merging:
            return

        if not LegalShield.has_user_consented():
            if hasattr(self.app, "_show_consent_dialog"):
                self.app._show_consent_dialog()
            return

        if not self.docx_path:
            messagebox.showwarning("Missing Template", "Please select a Word template (.docx) first.")
            return

        if not self.data_path:
            messagebox.showwarning("Missing Spreadsheet", "Please select a recipient spreadsheet (.xlsx or .csv) first.")
            return

        if not self.data_records:
            messagebox.showwarning("No Data Records", "The selected spreadsheet does not contain any recipient rows.")
            return

        if not self.detected_tags:
            proceed = messagebox.askyesno(
                "No Placeholders Detected",
                "No dynamic placeholders (such as {{FullName}} or Word MergeFields) were detected in this Word template.\n\nAll generated PDFs will be identical copies of the template.\n\nDo you want to proceed with the batch merge anyway?"
            )
            if not proceed:
                return

        # Check for unmapped tags and warn user (Technical Architecture Rule 3)
        mapping = self._get_active_mapping()
        unmapped = [tag for tag in self.detected_tags if not mapping.get(tag)]
        if unmapped:
            tag_list = ", ".join([f"{{{{{t}}}}}" for t in unmapped[:5]])
            if len(unmapped) > 5:
                tag_list += f" (+{len(unmapped)-5} more)"
            proceed = messagebox.askyesno(
                "Unmapped Placeholders Detected",
                f"The following template placeholder(s) are not mapped to any spreadsheet column:\n\n{tag_list}\n\nUnmapped placeholders will be left blank in output documents.\n\nDo you want to proceed with the batch merge anyway?"
            )
            if not proceed:
                return

        total_records = len(self.data_records)

        # Check Trial / License
        authorized, msg = self.license_engine.consume_batch_export_run()
        if not authorized:
            # Prompt purchase dialog
            res = messagebox.askyesno(
                "Trial Limit Reached",
                f"{msg}\n\nWould you like to upgrade to the Lifetime License ($6.99) now for unlimited merges?"
            )
            if res and hasattr(self.app, "switch_tab"):
                self.app.switch_tab("settings")
            return

        # Start merge in background thread
        self.is_merging = True
        self.stop_requested = False
        self.btn_test_row1.configure(state="disabled")
        self.btn_batch_merge.configure(
            text="⏹️ Cancel Merge",
            fg_color="#EF4444",
            hover_color="#DC2626",
            command=self._cancel_batch_merge
        )
        self.progress_bar.set(0)

        # Update header trial status pill
        if hasattr(self.app, "settings_view"):
            self.app.settings_view.refresh_status()

        t = threading.Thread(target=self._batch_merge_worker, args=(total_records,), daemon=True)
        t.start()

    def _batch_merge_worker(self, total_records: int):
        import tempfile
        tmp_dir = tempfile.gettempdir()
        out_folder = ""
        successful_count = 0
        failed_count = 0
        start_time = time.perf_counter()

        try:
            out_folder = self.entry_out_dir.get().strip() or os.path.dirname(self.docx_path)
            os.makedirs(out_folder, exist_ok=True)

            formula = self.entry_formula.get().strip() or "Certificate_{Index}_{FullName}.pdf"
            mapping = self._get_active_mapping()
            allocated_names = set()

            for i, record in enumerate(self.data_records, start=1):
                if self.stop_requested:
                    break

                effective_record = dict(record)
                for tag, col in mapping.items():
                    if col and col in record:
                        effective_record[tag] = record[col]

                out_filename = DynamicNamer.generate(formula, effective_record, index=i, output_dir=out_folder, allocated_names=allocated_names)
                out_pdf = os.path.join(out_folder, out_filename)
                tmp_docx = os.path.join(tmp_dir, f"_dmp_{os.getpid()}_{i}.docx")

                try:
                    DocumentMerger.merge_single_record(self.docx_path, mapping, record, tmp_docx)
                    ok, engine = PDFConverter.convert_docx_to_pdf(tmp_docx, out_pdf)
                    if os.path.exists(tmp_docx):
                        os.remove(tmp_docx)

                    if ok:
                        successful_count += 1
                    else:
                        failed_count += 1
                except Exception:
                    failed_count += 1
                    if os.path.exists(tmp_docx):
                        try:
                            os.remove(tmp_docx)
                        except Exception:
                            pass

                # Update GUI progress
                pct = i / total_records
                elapsed = time.perf_counter() - start_time
                rate = i / max(0.1, elapsed)
                rem_sec = int((total_records - i) / max(0.1, rate))

                def _update_ui(p=pct, cur_i=i, rem=rem_sec):
                    self.progress_bar.set(p)
                    self.lbl_progress_status.configure(
                        text=f"Generating {cur_i}/{total_records} PDFs... (~{rem}s remaining)"
                    )

                self.after(0, _update_ui)
        except Exception as e:
            # Fatal batch worker error before or during loop
            self.license_engine.refund_batch_export_run()
            err_msg = str(e)
            def _handle_fatal_error():
                self.is_merging = False
                self.btn_batch_merge.configure(
                    state="normal",
                    text="🚀 Generate All Individual PDFs (Batch)",
                    fg_color=styles.PRIMARY_ACCENT,
                    hover_color=styles.PRIMARY_HOVER,
                    command=self._start_batch_merge
                )
                self.btn_test_row1.configure(state="normal")
                self.lbl_progress_status.configure(text=f"Batch error: {err_msg}")
                messagebox.showerror("Batch Export Failed", f"A fatal error occurred during batch setup:\n{err_msg}")
                if hasattr(self.app, "settings_view"):
                    self.app.settings_view.refresh_status()
            self.after(0, _handle_fatal_error)
            return

        total_elapsed = time.perf_counter() - start_time
        self.is_merging = False

        # If completely failed or cancelled with zero files generated, refund the session
        if (successful_count == 0 and total_records > 0) or (self.stop_requested and successful_count == 0):
            self.license_engine.refund_batch_export_run()
            if hasattr(self.app, "settings_view"):
                self.after(0, self.app.settings_view.refresh_status)

        def _finish_ui():
            self.btn_batch_merge.configure(
                state="normal",
                text="🚀 Generate All Individual PDFs (Batch)",
                fg_color=styles.PRIMARY_ACCENT,
                hover_color=styles.PRIMARY_HOVER,
                command=self._start_batch_merge
            )
            self.btn_test_row1.configure(state="normal")
            self.btn_open_folder.configure(state="normal")

            if self.stop_requested:
                self.lbl_progress_status.configure(
                    text=f"⚠️ Batch cancelled. Generated {successful_count} of {total_records} PDFs in {total_elapsed:.1f}s."
                )
                messagebox.showwarning(
                    "Batch Merge Cancelled",
                    f"Batch merge was stopped by user.\nGenerated {successful_count} individual PDFs in:\n{out_folder}"
                )
            elif failed_count > 0:
                self.progress_bar.set(1.0)
                self.lbl_progress_status.configure(
                    text=f"⚠️ Batch completed with warnings: {successful_count} succeeded, {failed_count} failed in {total_elapsed:.1f}s."
                )
                messagebox.showwarning(
                    "Batch Merge Completed with Issues",
                    f"Generated {successful_count} of {total_records} individual PDFs in:\n{out_folder}\n\nNotice: {failed_count} record(s) failed during document generation or conversion. Please check spreadsheet formatting."
                )
                self._open_output_folder()
            else:
                self.progress_bar.set(1.0)
                self.lbl_progress_status.configure(
                    text=f"✓ Complete! Generated {successful_count} individual PDFs in {total_elapsed:.1f}s."
                )
                messagebox.showinfo(
                    "Batch Merge Complete",
                    f"Successfully generated {successful_count} individual PDFs in:\n{out_folder}\n\nTime elapsed: {total_elapsed:.1f} seconds."
                )
                self._open_output_folder()

        self.after(0, _finish_ui)

    def _open_output_folder(self):
        folder = self.entry_out_dir.get().strip()
        if folder and os.path.exists(folder):
            try:
                os.startfile(folder)
            except Exception:
                subprocess.Popen(["explorer", os.path.normpath(folder)])

