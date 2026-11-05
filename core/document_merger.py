"""
DocuMergePDF - High-Fidelity Word Document Mail-Merge Engine
Clones Word templates, replaces placeholders across paragraphs, runs, tables,
headers, and footers while strictly preserving run typography, colors, and margins.
"""

import os
import re
from typing import Dict, Any, List, Optional
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


MERGEFIELD_PATTERN = re.compile(r'MERGEFIELD\s+(?:"([^"]+)"|([^\s\\]+))', re.IGNORECASE)


class DocumentMerger:
    """Merges row records into a Word template with run-level format preservation."""

    @staticmethod
    def merge_single_record(
        template_path: str,
        mapping: Dict[str, Optional[str]],
        record: Dict[str, str],
        output_docx_path: str,
        empty_fallback: str = ""
    ) -> str:
        """
        Loads template_path, substitutes mapped fields using row record,
        and saves merged document to output_docx_path.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")

        os.makedirs(os.path.dirname(os.path.abspath(output_docx_path)), exist_ok=True)

        doc = Document(template_path)

        # Prepare replacement lookup: {clean_tag: replacement_value}
        replacements: Dict[str, str] = {}
        for tag, col in mapping.items():
            if col and col in record:
                val = record[col]
            else:
                val = empty_fallback
            replacements[tag.strip()] = str(val)

        # 1. Process all body paragraphs (including tables, nested tables, textboxes, shapes)
        for p_elem in doc._element.xpath('.//w:p'):
            p = Paragraph(p_elem, doc)
            DocumentMerger._replace_in_paragraph(p, replacements)

        # 2. Process all headers & footers across all sections (including textboxes and tables)
        for section in doc.sections:
            for part in [
                section.header, section.first_page_header, section.even_page_header,
                section.footer, section.first_page_footer, section.even_page_footer
            ]:
                if part and part._element is not None:
                    for p_elem in part._element.xpath('.//w:p'):
                        p = Paragraph(p_elem, part)
                        DocumentMerger._replace_in_paragraph(p, replacements)

        doc.save(output_docx_path)
        return output_docx_path

    @staticmethod
    def _unlink_and_replace_mergefields(paragraph: Paragraph, replacements: Dict[str, str]):
        """
        Finds Word MERGEFIELDs (both w:fldSimple and complex fields) in paragraph XML,
        substitutes the matching value, and converts them to standard runs (unlinks fields).
        """
        from docx.oxml import OxmlElement
        p_elem = paragraph._p

        # 1. Handle w:fldSimple
        for fld in list(p_elem.xpath('.//w:fldSimple')):
            instr = fld.get(qn('w:instr'), '')
            m = MERGEFIELD_PATTERN.search(instr)
            if m:
                tag = (m.group(1) or m.group(2) or "").strip()
                if tag in replacements:
                    val = replacements[tag]
                    t_elems = list(fld.iter(qn('w:t')))
                    if t_elems:
                        t_elems[0].text = val
                        for extra_t in t_elems[1:]:
                            extra_t.text = ''
                    parent = fld.getparent()
                    if parent is not None:
                        idx = parent.index(fld)
                        child_runs = list(fld.iterchildren(qn('w:r')))
                        if child_runs:
                            for r in child_runs:
                                parent.insert(idx, r)
                                idx += 1
                        else:
                            new_r = OxmlElement('w:r')
                            new_t = OxmlElement('w:t')
                            new_t.text = val
                            new_r.append(new_t)
                            parent.insert(idx, new_r)
                        parent.remove(fld)

        # 2. Handle complex fields (w:fldChar begin/separate/end with w:instrText)
        for instr in list(p_elem.xpath('.//w:instrText')):
            text = instr.text or ''
            m = MERGEFIELD_PATTERN.search(text)
            if m:
                tag = (m.group(1) or m.group(2) or "").strip()
                if tag in replacements:
                    val = replacements[tag]
                    instr_r = instr.getparent()
                    if instr_r is not None:
                        parent_p = instr_r.getparent()
                        if parent_p is not None:
                            runs = list(parent_p.xpath('./w:r'))
                            if instr_r in runs:
                                instr_idx = runs.index(instr_r)
                                begin_r = None
                                for i in range(instr_idx - 1, -1, -1):
                                    if runs[i].xpath('.//w:fldChar[@w:fldCharType="begin"]'):
                                        begin_r = runs[i]
                                        break
                                sep_r = None
                                end_r = None
                                disp_runs = []
                                for i in range(instr_idx + 1, len(runs)):
                                    r = runs[i]
                                    if r.xpath('.//w:fldChar[@w:fldCharType="separate"]'):
                                        sep_r = r
                                    elif r.xpath('.//w:fldChar[@w:fldCharType="end"]'):
                                        end_r = r
                                        break
                                    elif sep_r is not None:
                                        disp_runs.append(r)
                                if disp_runs:
                                    t_list = list(disp_runs[0].iter(qn('w:t')))
                                    if t_list:
                                        t_list[0].text = val
                                    else:
                                        new_t = OxmlElement('w:t')
                                        new_t.text = val
                                        disp_runs[0].append(new_t)
                                    for dr in disp_runs[1:]:
                                        for t in dr.iter(qn('w:t')):
                                            t.text = ''
                                else:
                                    # disp_runs was empty: insert new run before removing
                                    new_r = OxmlElement('w:r')
                                    new_t = OxmlElement('w:t')
                                    new_t.text = val
                                    new_r.append(new_t)
                                    ref_node = end_r if end_r is not None else instr_r
                                    ref_node.addprevious(new_r)

                                for r_to_remove in [begin_r, instr_r, sep_r, end_r]:
                                    if r_to_remove is not None and r_to_remove.getparent() is not None:
                                        r_to_remove.getparent().remove(r_to_remove)

    @staticmethod
    def _replace_in_paragraph(paragraph: Paragraph, replacements: Dict[str, str]):
        """
        Replaces placeholders inside a paragraph while preserving run formatting.
        Handles Word MERGEFIELDs, mustache tags, chevrons, and single braces.
        """
        # Step 0: Handle native Word MERGEFIELD fields
        DocumentMerger._unlink_and_replace_mergefields(paragraph, replacements)

        full_text = paragraph.text
        if not full_text or ("{" not in full_text and "«" not in full_text):
            return

        # Check which placeholders are present
        active_tags = [t for t in replacements if t in full_text or f"{{{{{t}}}}}" in full_text or f"{{{t}}}" in full_text or f"«{t}»" in full_text]
        if not active_tags:
            return

        # Step 1: Try single-run replacements first (fastest and keeps 100% exact styling)
        for tag, val in replacements.items():
            mustache_patterns = [
                f"{{{{{tag}}}}}",
                f"{{{{ {tag} }}}}",
                f"{{{tag}}}",
                f"«{tag}»"
            ]
            for pattern in mustache_patterns:
                for run in paragraph.runs:
                    if pattern in run.text:
                        run.text = run.text.replace(pattern, val)

        # Step 2: Check if any split placeholders remain in the whole paragraph text
        current_text = "".join(r.text for r in paragraph.runs)
        needs_cross_run_replacement = False
        for tag in active_tags:
            if (f"{{{{{tag}}}}}" in current_text or
                f"«{tag}»" in current_text or
                re.search(r"\{\{\s*" + re.escape(tag) + r"\s*\}\}", current_text) or
                re.search(r"(?<!\{)\{" + re.escape(tag) + r"\}(?!})", current_text) or
                re.search(r"«\s*" + re.escape(tag) + r"\s*»", current_text)):
                needs_cross_run_replacement = True
                break

        if not needs_cross_run_replacement:
            return

        # Step 3: Handle cross-run split placeholders
        # Reconstruct by finding spans across runs
        for tag, val in replacements.items():
            patterns = [
                re.compile(r"\{\{\s*" + re.escape(tag) + r"\s*\}\}"),
                re.compile(r"(?<!\{)\{" + re.escape(tag) + r"\}(?!\})"),
                re.compile(r"«\s*" + re.escape(tag) + r"\s*»"),
            ]
            for reg in patterns:
                DocumentMerger._replace_regex_across_runs(paragraph, reg, val)

    @staticmethod
    def _replace_regex_across_runs(paragraph: Paragraph, regex: re.Pattern, replacement: str):
        """Finds regex matches across adjacent runs in a paragraph and replaces them cleanly in reverse order."""
        runs = paragraph.runs
        if not runs:
            return

        full_text = "".join(r.text for r in runs)
        matches = list(regex.finditer(full_text))
        if not matches:
            return

        # Process from right to left so offsets remain valid and replacement cannot trigger recursive loop
        for match in reversed(matches):
            start_idx, end_idx = match.span()

            # Find which runs contain start_idx and end_idx
            char_count = 0
            start_run = -1
            end_run = -1
            start_offset = 0
            end_offset = 0

            for i, r in enumerate(runs):
                r_len = len(r.text)
                if start_run == -1 and char_count + r_len > start_idx:
                    start_run = i
                    start_offset = start_idx - char_count
                if char_count + r_len >= end_idx:
                    end_run = i
                    end_offset = end_idx - char_count
                    break
                char_count += r_len

            if start_run == -1 or end_run == -1:
                continue

            if start_run == end_run:
                # Single run match
                r = runs[start_run]
                r.text = r.text[:start_offset] + replacement + r.text[end_offset:]
            else:
                # Span across runs: Keep start run prefix + replacement, clear middle runs, keep end run suffix
                runs[start_run].text = runs[start_run].text[:start_offset] + replacement
                for mid in range(start_run + 1, end_run):
                    runs[mid].text = ""
                runs[end_run].text = runs[end_run].text[end_offset:]

