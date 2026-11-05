"""
DocuMergePDF - Settings, Licensing & Legal Compliance View
Handles Gumroad API online verification, offline Hardware-Bound Studio Keys,
trial session monitoring, and UCC Legal Shield inspection.
"""

import os
import sys
import webbrowser
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from typing import Optional
from . import styles
from core.licensing import (
    LicenseSecurity,
    PRICE_MONTHLY_USD,
    PRICE_LIFETIME_USD,
    GUMROAD_BUY_LIFETIME_URL,
    GUMROAD_BUY_MONTHLY_URL,
    open_store_product_page
)
from core.legal_shield import LegalShield, LEGAL_DISCLAIMER_SUMMARY


class SettingsView(ctk.CTkScrollableFrame):
    """Settings and commercial activation panel."""

    def __init__(self, master, app_instance):
        super().__init__(master, fg_color=styles.CANVAS_BG)
        self.app = app_instance
        self.license_engine: LicenseSecurity = app_instance.license_engine

        self._build_ui()
        self.refresh_status()

    def _build_ui(self):
        # 1. License Status Card
        self.card_status = ctk.CTkFrame(self, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        self.card_status.pack(fill="x", padx=25, pady=(20, 15))

        status_inner = ctk.CTkFrame(self.card_status, fg_color="transparent")
        status_inner.pack(fill="x", padx=20, pady=18)

        lbl_head = ctk.CTkLabel(status_inner, text="🔑 License & Commercial Entitlement", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY)
        lbl_head.pack(anchor="w")

        self.lbl_license_desc = ctk.CTkLabel(
            status_inner,
            text="Checking license status...",
            font=styles.FONT_BODY,
            text_color=styles.TEXT_SECONDARY,
            justify="left"
        )
        self.lbl_license_desc.pack(anchor="w", pady=(6, 12))

        # HWID Row
        hwid_frame = ctk.CTkFrame(status_inner, fg_color=styles.CANVAS_BG, corner_radius=8, border_color=styles.CARD_BORDER, border_width=1)
        hwid_frame.pack(fill="x", pady=5)

        hwid_inner = ctk.CTkFrame(hwid_frame, fg_color="transparent")
        hwid_inner.pack(fill="x", padx=12, pady=8)

        ctk.CTkLabel(hwid_inner, text="Hardware ID (HWID):", font=styles.FONT_BODY_BOLD, text_color=styles.TEXT_PRIMARY).pack(side="left")
        self.lbl_hwid = ctk.CTkLabel(hwid_inner, text=self.license_engine.hwid, font=styles.FONT_BODY, text_color=styles.PRIMARY_ACCENT)
        self.lbl_hwid.pack(side="left", padx=10)

        ctk.CTkButton(
            hwid_inner,
            text="📋 Copy HWID",
            width=95,
            height=26,
            font=styles.FONT_SMALL,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._copy_hwid
        ).pack(side="right")

        # 2. License Key Activation Card
        card_act = ctk.CTkFrame(self, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        card_act.pack(fill="x", padx=25, pady=10)

        act_inner = ctk.CTkFrame(card_act, fg_color="transparent")
        act_inner.pack(fill="x", padx=20, pady=18)

        ctk.CTkLabel(act_inner, text="⚡ Activate Commercial License", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY).pack(anchor="w")
        ctk.CTkLabel(
            act_inner,
            text="Paste your Gumroad License Key or Air-Gapped Offline Studio Key (DMP-LIFETIME-...):",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_SECONDARY
        ).pack(anchor="w", pady=(4, 10))

        key_box = ctk.CTkFrame(act_inner, fg_color="transparent")
        key_box.pack(fill="x")

        self.entry_key = ctk.CTkEntry(
            key_box,
            placeholder_text="Enter license key here (e.g. DMP-LIFETIME-... or Gumroad Key)",
            font=styles.FONT_BODY,
            height=36,
            border_color=styles.CARD_BORDER
        )
        self.entry_key.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_activate = ctk.CTkButton(
            key_box,
            text="✓ Activate License",
            font=styles.FONT_BODY_BOLD,
            height=36,
            width=140,
            fg_color=styles.PRIMARY_ACCENT,
            hover_color=styles.PRIMARY_HOVER,
            text_color="#FFFFFF",
            command=self._activate_license
        )
        self.btn_activate.pack(side="right")

        # 3. Upgrade Tiers Cards (Side-by-side)
        tiers_frame = ctk.CTkFrame(self, fg_color="transparent")
        tiers_frame.pack(fill="x", padx=25, pady=10)

        # Lifetime Card
        card_lifetime = ctk.CTkFrame(tiers_frame, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER_ACTIVE, border_width=2, corner_radius=12)
        card_lifetime.pack(side="left", fill="both", expand=True, padx=(0, 8))

        lt_inner = ctk.CTkFrame(card_lifetime, fg_color="transparent")
        lt_inner.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(lt_inner, text="🏆 Lifetime License", font=styles.FONT_HEADING, text_color=styles.PRIMARY_ACCENT).pack(anchor="w")
        ctk.CTkLabel(lt_inner, text=f"{PRICE_LIFETIME_USD} One-Time Payment", font=styles.FONT_CARD_TITLE, text_color=styles.TEXT_PRIMARY).pack(anchor="w", pady=(2, 6))
        ctk.CTkLabel(
            lt_inner,
            text="• Unlimited batch mail merges forever\n• High-res vector Word-to-PDF rendering\n• Dynamic filename formula builder\n• Zero cloud subscriptions",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_SECONDARY,
            justify="left"
        ).pack(anchor="w", pady=(4, 12))

        ctk.CTkButton(
            lt_inner,
            text=f"Buy Lifetime License ({PRICE_LIFETIME_USD})",
            font=styles.FONT_BODY_BOLD,
            fg_color=styles.PRIMARY_ACCENT,
            hover_color=styles.PRIMARY_HOVER,
            text_color="#FFFFFF",
            command=lambda: webbrowser.open(GUMROAD_BUY_LIFETIME_URL)
        ).pack(fill="x")

        # Monthly Card
        card_monthly = ctk.CTkFrame(tiers_frame, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        card_monthly.pack(side="right", fill="both", expand=True, padx=(8, 0))

        mo_inner = ctk.CTkFrame(card_monthly, fg_color="transparent")
        mo_inner.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(mo_inner, text="⚡ 30-Day Project Pass", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY).pack(anchor="w")
        ctk.CTkLabel(mo_inner, text=f"{PRICE_MONTHLY_USD} / month", font=styles.FONT_CARD_TITLE, text_color=styles.TEXT_PRIMARY).pack(anchor="w", pady=(2, 6))
        ctk.CTkLabel(
            mo_inner,
            text="• Perfect for one-off graduation batches\n• 30 days of unlimited batch exports\n• Cancel anytime with zero friction\n• Immediate key delivery",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_SECONDARY,
            justify="left"
        ).pack(anchor="w", pady=(4, 12))

        ctk.CTkButton(
            mo_inner,
            text=f"Get Monthly Pass ({PRICE_MONTHLY_USD})",
            font=styles.FONT_BODY_BOLD,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=lambda: webbrowser.open(GUMROAD_BUY_MONTHLY_URL)
        ).pack(fill="x")

        # 4. Legal & Developer Card
        card_legal = ctk.CTkFrame(self, fg_color=styles.CARD_BG, border_color=styles.CARD_BORDER, border_width=1, corner_radius=12)
        card_legal.pack(fill="x", padx=25, pady=(10, 25))

        legal_inner = ctk.CTkFrame(card_legal, fg_color="transparent")
        legal_inner.pack(fill="x", padx=20, pady=16)

        ctk.CTkLabel(legal_inner, text="🛡️ Legal Protection & Privacy Compliance", font=styles.FONT_HEADING, text_color=styles.TEXT_PRIMARY).pack(anchor="w")
        ctk.CTkLabel(
            legal_inner,
            text="100% Offline execution ensures your sensitive student records and payroll remain confidential on this device.\nProtected under UCC § 2-316 / § 2-719 liability ceiling ($1.00 USD).",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_SECONDARY,
            justify="left"
        ).pack(anchor="w", pady=(4, 10))

        leg_btns = ctk.CTkFrame(legal_inner, fg_color="transparent")
        leg_btns.pack(fill="x")

        ctk.CTkButton(
            leg_btns,
            text="📜 View Legal Terms & Privacy",
            font=styles.FONT_SMALL,
            height=30,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._show_legal_modal
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            leg_btns,
            text="🏪 View on Microsoft Store",
            font=styles.FONT_SMALL,
            height=30,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=open_store_product_page
        ).pack(side="left", padx=(0, 10))

    def refresh_status(self):
        st = self.license_engine.get_status()
        if st["is_lifetime"]:
            desc = "🏆 LIFETIME COMMERCIAL LICENSE ACTIVE\nYou have unlimited, permanent access to all batch document merging features."
        elif st["is_monthly"]:
            desc = f"⚡ 30-DAY MONTHLY PASS ACTIVE ({st['monthly_days_remaining']} days remaining)\nUnlimited batch document merges are active."
        else:
            used = st["runs_used"]
            max_r = st["max_trial_runs"]
            rem = st["runs_remaining"]
            desc = f"⏳ FREE TRIAL: {rem} of {max_r} batch export sessions remaining.\nGenerating Row 1 test documents is 100% UNLIMITED and FREE FOREVER."

        self.lbl_license_desc.configure(text=desc)
        if hasattr(self.app, "update_header_trial_badge"):
            self.app.update_header_trial_badge(st)

    def _copy_hwid(self):
        self.clipboard_clear()
        self.clipboard_append(self.license_engine.hwid)
        messagebox.showinfo("Copied", "Hardware ID copied to clipboard.")

    def _activate_license(self):
        key = self.entry_key.get().strip()
        if not key:
            messagebox.showwarning("Input Required", "Please enter a valid license key.")
            return

        ok, msg = self.license_engine.verify_gumroad_license(key)
        if ok:
            messagebox.showinfo("Activation Successful", msg)
            self.refresh_status()
        else:
            messagebox.showerror("Activation Failed", msg)

    def _show_legal_modal(self):
        terms = LegalShield.read_full_terms()
        win = ctk.CTkToplevel(self)
        win.title("DocuMergePDF - Terms of Service & Privacy Shield")
        win.geometry("640x500")
        win.transient(self)
        win.grab_set()

        tb = ctk.CTkTextbox(win, font=styles.FONT_SMALL, wrap="word")
        tb.pack(fill="both", expand=True, padx=20, pady=20)
        tb.insert("1.0", terms)
        tb.configure(state="disabled")

        ctk.CTkButton(win, text="Close", command=win.destroy).pack(pady=(0, 15))

