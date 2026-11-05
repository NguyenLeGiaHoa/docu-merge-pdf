"""
DocuMergePDF - Live A4 Document Visualizer & Page Preview View
Renders pixel-sharp vector preview cards with realistic page drop shadow,
zoom controls, and system PDF launcher.
"""

import os
import subprocess
import customtkinter as ctk
from PIL import Image, ImageTk
from typing import Optional
from . import styles
from core.pdf_converter import PDFConverter


class PreviewView(ctk.CTkFrame):
    """Interactive preview canvas for generated Row 1 sample PDF."""

    def __init__(self, master, app_instance):
        super().__init__(master, fg_color=styles.CANVAS_BG)
        self.app = app_instance
        self.current_pdf_path: Optional[str] = None
        self.current_pil_image: Optional[Image.Image] = None
        self.current_tk_image: Optional[ImageTk.PhotoImage] = None
        self.zoom_factor = 0.55  # Default display scale for A4

        self._build_ui()

    def _build_ui(self):
        # Top toolbar
        toolbar = ctk.CTkFrame(self, fg_color=styles.CARD_BG, height=50, corner_radius=8, border_color=styles.CARD_BORDER, border_width=1)
        toolbar.pack(fill="x", padx=20, pady=(15, 10))
        toolbar.pack_propagate(False)

        tb_inner = ctk.CTkFrame(toolbar, fg_color="transparent")
        tb_inner.pack(fill="both", expand=True, padx=15)

        self.lbl_title = ctk.CTkLabel(
            tb_inner,
            text="👁️ Live A4 Document Preview (Row 1 Sample)",
            font=styles.FONT_HEADING,
            text_color=styles.TEXT_PRIMARY
        )
        self.lbl_title.pack(side="left", pady=10)

        # Right action buttons
        btn_box = ctk.CTkFrame(tb_inner, fg_color="transparent")
        btn_box.pack(side="right", pady=8)

        self.btn_zoom_out = ctk.CTkButton(
            btn_box,
            text="🔍 -",
            width=40,
            height=30,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._zoom_out
        )
        self.btn_zoom_out.pack(side="left", padx=3)

        self.lbl_zoom = ctk.CTkLabel(
            btn_box,
            text="100%",
            font=styles.FONT_SMALL,
            text_color=styles.TEXT_SECONDARY,
            width=45
        )
        self.lbl_zoom.pack(side="left", padx=2)

        self.btn_zoom_in = ctk.CTkButton(
            btn_box,
            text="🔍 +",
            width=40,
            height=30,
            fg_color=styles.SECONDARY_BTN_BG,
            hover_color=styles.SECONDARY_BTN_HOVER,
            text_color=styles.TEXT_PRIMARY,
            command=self._zoom_in
        )
        self.btn_zoom_in.pack(side="left", padx=3)

        self.btn_open_external = ctk.CTkButton(
            btn_box,
            text="↗️ Open External PDF",
            font=styles.FONT_BODY_BOLD,
            height=30,
            fg_color=styles.PRIMARY_ACCENT,
            hover_color=styles.PRIMARY_HOVER,
            text_color="#FFFFFF",
            command=self._open_external_pdf
        )
        self.btn_open_external.pack(side="left", padx=(10, 0))

        # Main scrollable canvas container
        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color=styles.CANVAS_BG,
            border_width=0
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))

        # Document shadow and canvas holder
        self.canvas_card = ctk.CTkFrame(
            self.scroll_frame,
            fg_color=styles.CARD_BG,
            border_color=styles.CARD_BORDER,
            border_width=1,
            corner_radius=10
        )
        self.canvas_card.pack(pady=20, padx=20)

        # Image label inside card
        self.lbl_image = ctk.CTkLabel(
            self.canvas_card,
            text="\n\n📄 No Document Generated Yet\n\nClick '👁️ Generate Row 1 Test PDF' on the Merge Studio tab\nto instantly inspect your vector-rendered document.\n\n",
            font=styles.FONT_SUBTITLE,
            text_color=styles.TEXT_SECONDARY
        )
        self.lbl_image.pack(padx=25, pady=25)

    def load_pdf(self, pdf_path: str):
        """Loads and rasterizes a PDF into the preview card."""
        if not os.path.exists(pdf_path):
            return

        self.current_pdf_path = pdf_path
        self.lbl_title.configure(text=f"👁️ Preview: {os.path.basename(pdf_path)}")

        pil_img = PDFConverter.render_preview_image(pdf_path, page_num=0, dpi=160)
        if pil_img:
            self.current_pil_image = pil_img
            self._render_image()

    def _render_image(self):
        if not self.current_pil_image:
            return

        orig_w, orig_h = self.current_pil_image.size
        new_w = int(orig_w * self.zoom_factor)
        new_h = int(orig_h * self.zoom_factor)

        resized = self.current_pil_image.resize((new_w, new_h), Image.Resampling.LANCZOS)
        self.current_tk_image = ctk.CTkImage(light_image=resized, dark_image=resized, size=(new_w, new_h))

        self.lbl_image.configure(image=self.current_tk_image, text="")
        pct = int((self.zoom_factor / 0.55) * 100)
        self.lbl_zoom.configure(text=f"{pct}%")

    def _zoom_in(self):
        if self.current_pil_image and self.zoom_factor < 1.2:
            self.zoom_factor += 0.08
            self._render_image()

    def _zoom_out(self):
        if self.current_pil_image and self.zoom_factor > 0.25:
            self.zoom_factor -= 0.08
            self._render_image()

    def _open_external_pdf(self):
        if self.current_pdf_path and os.path.exists(self.current_pdf_path):
            try:
                os.startfile(self.current_pdf_path)
            except Exception:
                try:
                    import webbrowser
                    webbrowser.open(os.path.abspath(self.current_pdf_path))
                except Exception:
                    subprocess.Popen(f'start "" "{os.path.abspath(self.current_pdf_path)}"', shell=True)
