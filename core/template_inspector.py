"""
DocuMergePDF - Word Template Placeholder Inspector Engine
Detects mustache tags {{FieldName}}, {FieldName}, and Word MergeFields
across paragraphs, runs, tables, headers, and footers.
"""

import os
import re
from typing import List, Dict, Any, Set
from docx import Document


MUSTACHE_PATTERN = re.compile(r"\{\{\s*([^{}\r\n]+?)\s*\}\}")
CHEVRON_PATTERN = re.compile(r"«\s*([^«»\r\n]+?)\s*»")
SINGLE_BRACE_PATTERN = re.compile(r"(?<!\{)\{([^{}\r\n]+?)\}(?!\})")
MERGEFIELD_PATTERN = re.compile(r'MERGEFIELD\s+(?:"([^"]+)"|([^\s\\]+))', re.IGNORECASE)


class TemplateInspector:
    """Inspects Word .docx templates and extracts all dynamic placeholders."""

    @staticmethod
    def inspect(docx_path: str) -> Dict[str, Any]:
        """
        Inspects a .docx file and returns metadata and all detected unique placeholders.
        """
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"Template file not found: {docx_path}")

        doc = Document(docx_path)
        placeholders: List[str] = []
        seen: Set[str] = set()

        def add_tag(tag: str):
            clean_tag = tag.strip()
            if clean_tag and clean_tag not in seen:
                seen.add(clean_tag)
                placeholders.append(clean_tag)

        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn

        # 1. Inspect all paragraphs across document body (including tables, nested tables, textboxes, drawings)
        for p_elem in doc._element.xpath('.//w:p'):
            p = Paragraph(p_elem, doc)
            TemplateInspector._extract_from_text(p.text, add_tag)
            TemplateInspector._extract_from_p_elem_xml(p_elem, add_tag)

        # 2. Inspect Headers & Footers across all sections (including header tables, textboxes, shapes)
        for section in doc.sections:
            for part in [
                section.header, section.first_page_header, section.even_page_header,
                section.footer, section.first_page_footer, section.even_page_footer
            ]:
                if part and part._element is not None:
                    for p_elem in part._element.xpath('.//w:p'):
                        p = Paragraph(p_elem, part)
                        TemplateInspector._extract_from_text(p.text, add_tag)
                        TemplateInspector._extract_from_p_elem_xml(p_elem, add_tag)

        return {
            "path": docx_path,
            "filename": os.path.basename(docx_path),
            "placeholders": placeholders,
            "placeholder_count": len(placeholders),
            "paragraphs_count": len(doc.paragraphs),
            "tables_count": len(doc.tables),
            "sections_count": len(doc.sections),
            "is_valid": len(placeholders) > 0,
        }

    @staticmethod
    def _extract_from_text(text: str, add_callback):
        if not text:
            return
        # Find all {{Mustache}}
        for m in MUSTACHE_PATTERN.findall(text):
            add_callback(m)
        # Find all «Chevrons»
        for c in CHEVRON_PATTERN.findall(text):
            add_callback(c)
        # Find all single braces {Tag}
        for s in SINGLE_BRACE_PATTERN.findall(text):
            add_callback(s)

    @staticmethod
    def _extract_from_p_elem_xml(p_elem, add_callback):
        from docx.oxml.ns import qn
        # 1. Find Word MERGEFIELD in w:fldSimple
        for fld in p_elem.xpath('.//w:fldSimple'):
            instr = fld.get(qn('w:instr'), '')
            if instr:
                m = MERGEFIELD_PATTERN.search(instr)
                if m:
                    tag = (m.group(1) or m.group(2) or "").strip()
                    if tag:
                        add_callback(tag)

        # 2. Find Word MERGEFIELD in complex fields (w:instrText)
        for instr in p_elem.xpath('.//w:instrText'):
            text = instr.text or ''
            if text:
                m = MERGEFIELD_PATTERN.search(text)
                if m:
                    tag = (m.group(1) or m.group(2) or "").strip()
                    if tag:
                        add_callback(tag)

        # 3. Fallback regex match across raw XML
        xml_str = p_elem.xml
        if xml_str:
            for match in MERGEFIELD_PATTERN.finditer(xml_str):
                tag = (match.group(1) or match.group(2) or "").strip()
                if tag and not tag.startswith("&quot;"):
                    add_callback(tag)

