"""
DocuMergePDF - Multi-Tier PDF Conversion & Live Page Preview Engine
Tier 1: Microsoft Word COM Automation (Native Pixel-Perfect Rendering if Word installed)
Tier 2: LibreOffice Headless CLI (Cross-platform headless engine)
Tier 3: Pure Python DOCX-to-PDF Flowable Engine (ReportLab Standalone Fallback - Zero Dependencies)
Live Preview: PyMuPDF high-DPI rasterizer for real-time A4 card display
"""

import os
import sys
import subprocess
import shutil
from typing import Optional, Tuple
from PIL import Image

try:
    import pymupdf as fitz
    HAS_PYMUPDF = True
except ImportError:
    try:
        import fitz  # PyMuPDF fallback
        HAS_PYMUPDF = True
    except ImportError:
        HAS_PYMUPDF = False

# ReportLab imports for standalone pure Python rendering
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from docx import Document
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


class PDFConverter:
    """Multi-tier document conversion engine with automatic fallback and preview renderer."""

    _FONT_INITIALIZED = False
    _REGULAR_FONT = "Helvetica"
    _BOLD_FONT = "Helvetica-Bold"
    _WORD_COM_AVAILABLE: Optional[bool] = None
    _LIBREOFFICE_PATH: Optional[str] = None
    _LIBREOFFICE_CHECKED: bool = False

    @classmethod
    def _init_unicode_fonts(cls):
        """Initializes high-fidelity Unicode TrueType fonts (Arial / Times New Roman / Segoe UI) on Windows."""
        if cls._FONT_INITIALIZED:
            return
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont

            windir = os.environ.get("WINDIR", os.environ.get("SystemRoot", r"C:\Windows"))
            fonts_dir = os.path.join(windir, "Fonts")

            candidates = [
                ("ArialUnicode", os.path.join(fonts_dir, "arial.ttf"),
                 "ArialUnicode-Bold", os.path.join(fonts_dir, "arialbd.ttf"),
                 "ArialUnicode-Italic", os.path.join(fonts_dir, "ariali.ttf"),
                 "ArialUnicode-BoldItalic", os.path.join(fonts_dir, "arialbi.ttf")),
                ("TimesUnicode", os.path.join(fonts_dir, "times.ttf"),
                 "TimesUnicode-Bold", os.path.join(fonts_dir, "timesbd.ttf"),
                 "TimesUnicode-Italic", os.path.join(fonts_dir, "timesi.ttf"),
                 "TimesUnicode-BoldItalic", os.path.join(fonts_dir, "timesbi.ttf")),
                ("SegoeUIUnicode", os.path.join(fonts_dir, "segoeui.ttf"),
                 "SegoeUIUnicode-Bold", os.path.join(fonts_dir, "segoeuib.ttf"),
                 "SegoeUIUnicode-Italic", os.path.join(fonts_dir, "segoeuii.ttf"),
                 "SegoeUIUnicode-BoldItalic", os.path.join(fonts_dir, "segoeuiz.ttf")),
            ]
            for reg_name, reg_path, bold_name, bold_path, it_name, it_path, bi_name, bi_path in candidates:
                if os.path.exists(reg_path) and os.path.exists(bold_path):
                    try:
                        pdfmetrics.registerFont(TTFont(reg_name, reg_path))
                        pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                        act_it = it_name if os.path.exists(it_path) else reg_name
                        if os.path.exists(it_path):
                            pdfmetrics.registerFont(TTFont(it_name, it_path))
                        act_bi = bi_name if os.path.exists(bi_path) else bold_name
                        if os.path.exists(bi_path):
                            pdfmetrics.registerFont(TTFont(bi_name, bi_path))
                        pdfmetrics.registerFontFamily(
                            reg_name,
                            normal=reg_name,
                            bold=bold_name,
                            italic=act_it,
                            boldItalic=act_bi
                        )
                        cls._REGULAR_FONT = reg_name
                        cls._BOLD_FONT = bold_name
                        cls._FONT_INITIALIZED = True
                        return
                    except Exception:
                        continue
        except Exception:
            pass
        cls._FONT_INITIALIZED = True

    @staticmethod
    def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> Tuple[bool, str]:
        """
        Converts docx_path to pdf_path using the highest quality available engine.
        Returns: (success_bool, engine_used_or_error_message)
        """
        if not os.path.exists(docx_path):
            return False, f"Source file does not exist: {docx_path}"

        os.makedirs(os.path.dirname(os.path.abspath(pdf_path)), exist_ok=True)

        # 1. Try Microsoft Word COM Automation (Tier 1)
        if sys.platform == "win32":
            success, msg = PDFConverter._convert_with_word_com(docx_path, pdf_path)
            if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 500:
                return True, "Microsoft Word COM Engine (Tier 1)"

        # 2. Try LibreOffice Headless (Tier 2)
        libreoffice_bin = PDFConverter._find_libreoffice()
        if libreoffice_bin:
            success, msg = PDFConverter._convert_with_libreoffice(docx_path, pdf_path, libreoffice_bin)
            if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 500:
                return True, "LibreOffice Headless Engine (Tier 2)"

        # 3. Fallback to Pure Python ReportLab Engine (Tier 3 - Standalone)
        try:
            PDFConverter._convert_with_reportlab(docx_path, pdf_path)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 200:
                return True, "Pure Python Standalone Engine (Tier 3)"
        except Exception as e:
            return False, f"ReportLab conversion failed: {str(e)}"

        return False, "All conversion engines failed."

    @classmethod
    def _convert_with_word_com(cls, docx_path: str, pdf_path: str) -> Tuple[bool, str]:
        """Uses Windows PowerShell to automate Word COM safely with timeout."""
        if cls._WORD_COM_AVAILABLE is False:
            return False, "Microsoft Word COM automation unavailable."

        abs_docx = os.path.abspath(docx_path).replace("'", "''")
        abs_pdf = os.path.abspath(pdf_path).replace("'", "''")

        ps_script = f"""
$ErrorActionPreference = 'Stop'
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open('{abs_docx}', $false, $true) # ReadOnly
    $doc.SaveAs([ref]'{abs_pdf}', [ref]17) # 17 = wdFormatPDF
    $doc.Close($false)
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    exit 0
}} catch {{
    if ($word) {{
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }}
    exit 1
}}
"""
        try:
            import base64
            encoded_cmd = base64.b64encode(ps_script.encode("utf-16le")).decode("ascii")
            # Run powershell with EncodedCommand for 100% Unicode path fidelity
            res = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded_cmd],
                capture_output=True,
                text=True,
                timeout=20
            )
            success = (res.returncode == 0)
            if success:
                cls._WORD_COM_AVAILABLE = True
            elif cls._WORD_COM_AVAILABLE is None:
                cls._WORD_COM_AVAILABLE = False
            return success, res.stderr
        except Exception as e:
            if cls._WORD_COM_AVAILABLE is None:
                cls._WORD_COM_AVAILABLE = False
            return False, str(e)

    @classmethod
    def _find_libreoffice(cls) -> Optional[str]:
        if cls._LIBREOFFICE_CHECKED:
            return cls._LIBREOFFICE_PATH
        cls._LIBREOFFICE_CHECKED = True
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        candidates = [
            os.path.join(pf, "LibreOffice", "program", "soffice.exe"),
            os.path.join(pf86, "LibreOffice", "program", "soffice.exe"),
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        which = shutil.which("soffice")
        if which:
            cls._LIBREOFFICE_PATH = which
            return which
        for c in candidates:
            if os.path.exists(c):
                cls._LIBREOFFICE_PATH = c
                return c
        return None

    @staticmethod
    def _convert_with_libreoffice(docx_path: str, pdf_path: str, soffice_path: str) -> Tuple[bool, str]:
        out_dir = os.path.dirname(os.path.abspath(pdf_path))
        cmd = [
            soffice_path,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            os.path.abspath(docx_path)
        ]
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            # Libreoffice saves as <basename>.pdf in outdir
            expected_tmp = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
            if os.path.exists(expected_tmp) and expected_tmp != os.path.abspath(pdf_path):
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                os.rename(expected_tmp, pdf_path)
            return (res.returncode == 0), res.stderr
        except Exception as e:
            return False, str(e)

    @classmethod
    def _convert_with_reportlab(cls, docx_path: str, pdf_path: str):
        """Pure Python standalone renderer mapping DOCX paragraphs and tables in sequential document order."""
        doc = Document(docx_path)
        story = []

        # Determine dynamic page size and margins from first section if available
        section = doc.sections[0] if doc.sections else None
        if section:
            page_w = float(section.page_width.pt)
            page_h = float(section.page_height.pt)
            top_margin = max(18.0, float(section.top_margin.pt))
            bottom_margin = max(18.0, float(section.bottom_margin.pt))
            left_margin = max(18.0, float(section.left_margin.pt))
            right_margin = max(18.0, float(section.right_margin.pt))
        else:
            page_w, page_h = 595.28, 841.89  # Default A4
            top_margin, bottom_margin, left_margin, right_margin = 40.0, 40.0, 50.0, 50.0

        avail_width = page_w - (left_margin + right_margin)

        doc_template = SimpleDocTemplate(
            pdf_path,
            pagesize=(page_w, page_h),
            leftMargin=left_margin,
            rightMargin=right_margin,
            topMargin=top_margin,
            bottomMargin=bottom_margin,
        )

        cls._init_unicode_fonts()
        reg_font = cls._REGULAR_FONT
        bold_font = cls._BOLD_FONT

        styles = getSampleStyleSheet()

        primary_dark = colors.HexColor("#111827")
        indigo_accent = colors.HexColor("#4F46E5")
        border_light = colors.HexColor("#E5E7EB")

        base_normal = styles["Normal"]
        base_normal.fontName = reg_font
        base_normal.fontSize = 11
        base_normal.leading = 15
        base_normal.textColor = primary_dark

        def _format_paragraph(p: DocxParagraph) -> Optional[Paragraph]:
            text = p.text.strip()
            if not text:
                return None

            # Detect alignment
            align = TA_LEFT
            if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                align = TA_CENTER
            elif p.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                align = TA_RIGHT
            elif p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                align = TA_JUSTIFY

            # Check heading styles
            style_name = p.style.name.lower() if p.style else ""
            is_heading = "heading" in style_name or "title" in style_name

            font_size = 11
            leading = 15
            font_name = reg_font
            text_color = primary_dark

            if "title" in style_name or "heading 1" in style_name:
                font_size = 18
                leading = 22
                font_name = bold_font
                text_color = indigo_accent
            elif "heading 2" in style_name:
                font_size = 14
                leading = 18
                font_name = bold_font
                text_color = primary_dark
            elif "heading 3" in style_name:
                font_size = 12
                leading = 16
                font_name = bold_font

            # Check if all runs in the paragraph are bold
            non_empty_runs = [r for r in p.runs if r.text.strip()]
            all_bold = bool(non_empty_runs and all(r.bold for r in non_empty_runs))
            if is_heading or all_bold:
                font_name = bold_font

            # Reconstruct HTML formatted paragraph text with run-level typography
            formatted_parts = []
            for r in p.runs:
                clean_run_text = (
                    r.text
                    .replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace("\n", "<br/>")
                )
                if not clean_run_text:
                    continue

                part = clean_run_text
                # Extract run-level size and color if specified
                font_attrs = []
                try:
                    if r.font and r.font.size:
                        r_pt = float(r.font.size.pt)
                        if abs(r_pt - font_size) >= 0.5:
                            font_attrs.append(f'size="{r_pt:.1f}"')
                    if r.font and r.font.color and r.font.color.rgb:
                        rgb = str(r.font.color.rgb)
                        font_attrs.append(f'color="#{rgb}"')
                except Exception:
                    pass

                if font_attrs:
                    part = f'<font {" ".join(font_attrs)}>{part}</font>'

                if r.underline:
                    part = f"<u>{part}</u>"

                if r.bold and r.italic:
                    part = f"<b><i>{part}</i></b>"
                elif r.bold:
                    part = f"<b>{part}</b>"
                elif r.italic:
                    part = f"<i>{part}</i>"

                formatted_parts.append(part)

            p_xml = "".join(formatted_parts) or text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")

            p_style = ParagraphStyle(
                f"Custom_{len(story)}_{id(p)}",
                parent=base_normal,
                fontName=font_name,
                fontSize=font_size,
                leading=leading,
                textColor=text_color,
                alignment=align,
                spaceAfter=6,
            )
            return Paragraph(p_xml, p_style)

        def _collect_flowables_from_element(elem, container_story):
            for child in elem:
                tag = child.tag
                if tag.endswith("p"):
                    p = DocxParagraph(child, doc)
                    # Check for inline images
                    for r in p.runs:
                        for blip in r._r.xpath('.//a:blip/@r:embed'):
                            if blip and blip in doc.part.related_parts:
                                try:
                                    import io
                                    from reportlab.platypus import Image as RLImage
                                    rel_part = doc.part.related_parts[blip]
                                    img_data = io.BytesIO(rel_part.blob)
                                    with Image.open(img_data) as pil_im:
                                        iw, ih = pil_im.size
                                    img_data.seek(0)
                                    max_w = min(float(avail_width), 360.0)
                                    scale = min(1.0, max_w / max(1, iw))
                                    rl_img = RLImage(img_data, width=iw * scale, height=ih * scale)
                                    container_story.append(rl_img)
                                    container_story.append(Spacer(1, 6))
                                except Exception:
                                    pass

                    rl_p = _format_paragraph(p)
                    if rl_p:
                        container_story.append(rl_p)
                    else:
                        container_story.append(Spacer(1, 6))

                    # Check for text in drawing textboxes
                    for tb_p_elem in child.xpath('.//w:txbxContent//w:p'):
                        tb_p = DocxParagraph(tb_p_elem, doc)
                        tb_rl_p = _format_paragraph(tb_p)
                        if tb_rl_p:
                            container_story.append(tb_rl_p)

                    # Check for explicit page break
                    for r in p.runs:
                        if r._r.xpath('.//w:br[@w:type="page"]'):
                            from reportlab.platypus import PageBreak
                            container_story.append(PageBreak())

                elif tag.endswith("tbl"):
                    t = DocxTable(child, doc)
                    table_data = []
                    for row in t.rows:
                        row_cells = []
                        for cell in row.cells:
                            cell_paragraphs = []
                            for cp in cell.paragraphs:
                                c_p_flowable = _format_paragraph(cp)
                                if c_p_flowable:
                                    cell_paragraphs.append(c_p_flowable)
                            if not cell_paragraphs:
                                cell_text = cell.text.strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
                                cell_paragraphs.append(Paragraph(cell_text, base_normal))
                            row_cells.append(cell_paragraphs if len(cell_paragraphs) > 1 else cell_paragraphs[0])
                        table_data.append(row_cells)

                    if table_data:
                        col_count = max(len(r) for r in table_data) if table_data else 1
                        for r in table_data:
                            while len(r) < col_count:
                                r.append(Paragraph("", base_normal))
                        col_w = avail_width / max(1, col_count)
                        rl_table = Table(table_data, colWidths=[col_w] * col_count)
                        rl_table.setStyle(TableStyle([
                            ("BOX", (0, 0), (-1, -1), 1, border_light),
                            ("INNERGRID", (0, 0), (-1, -1), 0.5, border_light),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ]))
                        container_story.append(Spacer(1, 8))
                        container_story.append(rl_table)
                        container_story.append(Spacer(1, 8))

                elif tag.endswith("sdt"):
                    # Unroll content controls
                    content_node = child.find(qn('w:sdtContent'))
                    target_parent = content_node if content_node is not None else child
                    _collect_flowables_from_element(target_parent, container_story)

        # Collect story in natural document body order
        _collect_flowables_from_element(doc.element.body, story)

        # Extract header/footer text for page canvas callbacks if present
        header_text = ""
        footer_text = ""
        if doc.sections:
            sec = doc.sections[0]
            if sec.header and sec.header.paragraphs:
                header_text = " ".join(p.text.strip() for p in sec.header.paragraphs if p.text.strip())
            if sec.footer and sec.footer.paragraphs:
                footer_text = " ".join(p.text.strip() for p in sec.footer.paragraphs if p.text.strip())

        def _draw_decorations(canvas, d):
            canvas.saveState()
            canvas.setFont(reg_font, 9)
            canvas.setFillColor(colors.HexColor("#6B7280"))
            if header_text:
                canvas.drawCentredString(page_w / 2.0, page_h - top_margin / 1.5, header_text[:120])
            if footer_text:
                canvas.drawCentredString(page_w / 2.0, bottom_margin / 2.0, footer_text[:120])
            canvas.restoreState()

        doc_template.build(story, onFirstPage=_draw_decorations, onLaterPages=_draw_decorations)

    @staticmethod
    def render_preview_image(pdf_path: str, page_num: int = 0, dpi: int = 150) -> Optional[Image.Image]:
        """Renders a PDF page to a PIL Image using PyMuPDF for live A4 preview card display."""
        if not HAS_PYMUPDF or not os.path.exists(pdf_path):
            return None

        try:
            doc = fitz.open(pdf_path)
            if page_num >= len(doc):
                page_num = 0
            page = doc[page_num]
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            return img
        except Exception:
            return None

