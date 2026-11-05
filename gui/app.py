"""
DocuMergePDF - Master Desktop Application Window
Features Modern Editorial Indigo & Paper Light Theme, native drag-and-drop listener,
real-time trial status badge, and seamless navigation between Merge Studio, Live Preview, and Settings.
"""

import os
import sys
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from typing import Optional
from . import styles
from .views_merge import MergeView
from .views_preview import PreviewView
from .views_settings import SettingsView

try:
    from core.licensing import LicenseSecurity
    from core.legal_shield import LegalShield, LEGAL_DISCLAIMER_SUMMARY
except (ImportError, ValueError):
    from ..core.licensing import LicenseSecurity
    from ..core.legal_shield import LegalShield, LEGAL_DISCLAIMER_SUMMARY

try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False


class DocuMergePDFApp(ctk.CTk):
    """Primary CustomTkinter Application for DocuMergePDF."""

    def __init__(self):
        super().__init__()

        # Enforce Editorial Indigo Light Theme
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("DocuMergePDF — Mail Merge Word & Excel to Individual PDFs")
        self.geometry("1040x720")
        self.minsize(880, 620)
        self.configure(fg_color=styles.CANVAS_BG)

        # Set app icon if available
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_dir, "app_icon.ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass

        # Initialize Security & Licensing Engine
        self.license_engine = LicenseSecurity()

        self._build_header()
        self._build_tabs()
        self._build_footer()

        # Check First-Run Legal Shield Consent
        self.after(200, self._check_first_run_consent)

        # Initialize Windows native drag and drop if supported
        if HAS_WINDND:
            try:
                windnd.hook_dropfiles(self, func=self._on_dropped_files, force_unicode=True)
            except Exception:
                try:
                    windnd.hook_dropfiles(self, func=self._on_dropped_files)
                except Exception:
                    pass

    def _build_header(self):
        """Builds top status bar with brand logo and trial counter pill."""
        header_frame = ctk.CTkFrame(
            self,
            fg_color=styles.CARD_BG,
            border_color=styles.CARD_BORDER,
            border_width=1,
            corner_radius=0,
            height=65
        )
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        h_inner = ctk.CTkFrame(header_frame, fg_color="transparent")
        h_inner.pack(fill="both", expand=True, padx=20)

        # Brand Title Left
        title_box = ctk.CTkFrame(h_inner, fg_color="transparent")
        title_box.pack(side="left", pady=10)

        ctk.CTkLabel(
            title_box,
            text="📑 DocuMergePDF",
            font=styles.FONT_TITLE,
            text_color=styles.TEXT_PRIMARY
        ).pack(side="left")

        ctk.CTkLabel(
            title_box,
            text="   |   Word & Excel to Individual PDFs",
            font=styles.FONT_SUBTITLE,
            text_color=styles.TEXT_SECONDARY
        ).pack(side="left")

        # Right Status / Trial Pill
        right_box = ctk.CTkFrame(h_inner, fg_color="transparent")
        right_box.pack(side="right", pady=12)

        self.btn_trial_pill = ctk.CTkButton(
            right_box,
            text="TRIAL: 5/5 SESSIONS LEFT",
            font=styles.FONT_BADGE,
            height=32,
            corner_radius=16,
            fg_color=styles.SOFT_PILL_BG,
            hover_color=styles.CARD_BORDER,
            text_color=styles.PRIMARY_ACCENT,
            command=lambda: self.switch_tab("settings")
        )
        self.btn_trial_pill.pack(side="right")

        self.update_header_trial_badge(self.license_engine.get_status())

    def update_header_trial_badge(self, st: dict):
        if st.get("is_lifetime"):
            self.btn_trial_pill.configure(
                text="🏆 LIFETIME LICENSE ACTIVE",
                fg_color=styles.LIFETIME_GOLD_BG,
                text_color=styles.LIFETIME_GOLD_TEXT
            )
        elif st.get("is_monthly"):
            days = st.get("monthly_days_remaining", 30)
            self.btn_trial_pill.configure(
                text=f"⚡ MONTHLY PASS ({days}d LEFT)",
                fg_color=styles.SUCCESS_BG,
                text_color=styles.SUCCESS_TEXT
            )
        else:
            rem = st.get("runs_remaining", 5)
            max_r = st.get("max_trial_runs", 5)
            self.btn_trial_pill.configure(
                text=f"TRIAL: {rem}/{max_r} SESSIONS LEFT",
                fg_color=styles.SOFT_PILL_BG,
                text_color=styles.PRIMARY_ACCENT
            )

    def _build_tabs(self):
        """Builds tab navigation controls and content views."""
        nav_bar = ctk.CTkFrame(self, fg_color=styles.CANVAS_BG, height=45)
        nav_bar.pack(fill="x", side="top", padx=20, pady=(10, 0))

        self.seg_nav = ctk.CTkSegmentedButton(
            nav_bar,
            values=["📑 Merge Studio", "👁️ Document Preview", "⚙️ Settings & License"],
            font=styles.FONT_BODY_BOLD,
            selected_color=styles.PRIMARY_ACCENT,
            selected_hover_color=styles.PRIMARY_HOVER,
            unselected_color=styles.CARD_BG,
            unselected_hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._on_tab_selected
        )
        self.seg_nav.pack(side="left")
        self.seg_nav.set("📑 Merge Studio")

        # Container for swappable views
        self.view_container = ctk.CTkFrame(self, fg_color="transparent")
        self.view_container.pack(fill="both", expand=True)

        self.merge_view = MergeView(self.view_container, self)
        self.preview_view = PreviewView(self.view_container, self)
        self.settings_view = SettingsView(self.view_container, self)

        self.merge_view.pack(fill="both", expand=True)

    def _on_tab_selected(self, choice: str):
        if "Merge" in choice:
            self.switch_tab("merge")
        elif "Preview" in choice:
            self.switch_tab("preview")
        elif "Settings" in choice:
            self.switch_tab("settings")

    def switch_tab(self, tab_name: str):
        self.merge_view.pack_forget()
        self.preview_view.pack_forget()
        self.settings_view.pack_forget()

        if tab_name == "merge":
            self.seg_nav.set("📑 Merge Studio")
            self.merge_view.pack(fill="both", expand=True)
        elif tab_name == "preview":
            self.seg_nav.set("👁️ Document Preview")
            self.preview_view.pack(fill="both", expand=True)
        elif tab_name == "settings":
            self.seg_nav.set("⚙️ Settings & License")
            self.settings_view.refresh_status()
            self.settings_view.pack(fill="both", expand=True)

    def _build_footer(self):
        footer_frame = ctk.CTkFrame(
            self,
            fg_color=styles.CARD_BG,
            border_color=styles.CARD_BORDER,
            border_width=1,
            corner_radius=0,
            height=32
        )
        footer_frame.pack(fill="x", side="bottom")
        footer_frame.pack_propagate(False)

        f_inner = ctk.CTkFrame(footer_frame, fg_color="transparent")
        f_inner.pack(fill="both", expand=True, padx=20)

        ctk.CTkLabel(
            f_inner,
            text="🔒 100% Local On-Premise Execution • Zero Cloud Risk • GDPR & FERPA Compliant",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_SECONDARY
        ).pack(side="left", pady=6)

        ctk.CTkLabel(
            f_inner,
            text="DocuMergePDF v1.0.0 Commercial Edition",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_MUTED
        ).pack(side="right", pady=6)

    def _on_dropped_files(self, files):
        """Handles native drag and drop files from Windows Explorer."""
        if not files:
            return

        for raw_item in files:
            if isinstance(raw_item, bytes):
                try:
                    file_path = raw_item.decode("utf-8")
                except UnicodeDecodeError:
                    file_path = raw_item.decode("mbcs", errors="replace")
            else:
                file_path = str(raw_item)
            file_path = file_path.strip().strip('"').strip("'")
            if not os.path.exists(file_path):
                continue

            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".docx":
                self.merge_view.load_template(file_path)
                self.switch_tab("merge")
            elif ext in [".xlsx", ".xls", ".csv"]:
                self.merge_view.load_data(file_path)
                self.switch_tab("merge")

    def _check_first_run_consent(self):
        """Displays legal liability disclaimer and user indemnification dialog on first launch."""
        if not LegalShield.has_user_consented():
            self._show_consent_dialog()

    def _show_consent_dialog(self):
        win = ctk.CTkToplevel(self)
        win.title("DocuMergePDF - Terms of Use & Liability Disclaimer")
        win.geometry("620x460")
        win.transient(self)
        win.grab_set()

        lbl_t = ctk.CTkLabel(win, text="Welcome to DocuMergePDF", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY)
        lbl_t.pack(pady=(18, 6))

        lbl_s = ctk.CTkLabel(
            win,
            text="Please review the legal notice and liability limitation before proceeding:",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_SECONDARY
        )
        lbl_s.pack(pady=(0, 10))

        tb = ctk.CTkTextbox(win, font=styles.FONT_SMALL, wrap="word", border_color=styles.CARD_BORDER, border_width=1)
        tb.pack(fill="both", expand=True, padx=20, pady=5)
        tb.insert("1.0", LEGAL_DISCLAIMER_SUMMARY)
        tb.configure(state="disabled")

        def _on_accept():
            LegalShield.record_consent()
            win.destroy()

        btn_box = ctk.CTkFrame(win, fg_color="transparent")
        btn_box.pack(pady=15)

        ctk.CTkButton(
            btn_box,
            text="✓ I Understand & Accept (Proceed)",
            font=styles.FONT_BODY_BOLD,
            height=36,
            width=260,
            fg_color=styles.PRIMARY_ACCENT,
            hover_color=styles.PRIMARY_HOVER,
            command=_on_accept
        ).pack()

