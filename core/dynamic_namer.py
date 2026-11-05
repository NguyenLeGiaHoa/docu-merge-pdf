"""
DocuMergePDF - Dynamic Filename Generation, Sanitization & Collision Shield
Enforces Windows path compliance, strips illegal characters, and ensures non-destructive exports.
"""

import os
import re
from typing import Dict, Any, Set

ILLEGAL_CHARS_PATTERN = re.compile(r'[\x00-\x1f\x7f\\/:*?"<>|]')
MAX_FILENAME_LENGTH = 120
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
}


class DynamicNamer:
    """Generates sanitized, collision-free filenames based on row data and patterns."""

    @staticmethod
    def sanitize(name: str) -> str:
        """Strips illegal Windows characters and collapses consecutive underscores/spaces."""
        if not name:
            return "Document"
        # Neutralize relative directory traversal dots (..)
        clean = re.sub(r"\.{2,}", "_", name)
        # Strip stray curly braces from filename
        clean = re.sub(r"[{}]", "", clean)
        # Replace illegal characters with underscore
        clean = ILLEGAL_CHARS_PATTERN.sub("_", clean)
        # Collapse multiple spaces / underscores
        clean = re.sub(r"[_\s]+", "_", clean).strip(" ._")
        if not clean:
            clean = "Document"
        # Clamp length
        if len(clean) > MAX_FILENAME_LENGTH:
            clean = clean[:MAX_FILENAME_LENGTH].rstrip(" ._")
        # Protect against Windows DOS device names (e.g. CON, NUL, PRN, AUX with or without extensions)
        stem = os.path.splitext(clean)[0]
        if clean.upper() in WINDOWS_RESERVED_NAMES or stem.upper() in WINDOWS_RESERVED_NAMES:
            clean = f"Doc_{clean}"
        return clean

    @staticmethod
    def generate(
        pattern: str,
        record: Dict[str, str],
        index: int,
        output_dir: str,
        allocated_names: Set[str],
        ext: str = ".pdf"
    ) -> str:
        """
        Builds a collision-safe destination filename in output_dir.
        `pattern`: string with placeholders like "Certificate_{FullName}_{ID}"
        `record`: normalized row data
        `index`: 1-based row index
        `output_dir`: target directory
        `allocated_names`: set of filenames already reserved in this session
        """
        if not pattern:
            pattern = "Document_{Index}"

        result = pattern

        # Normalize double braces e.g. {{FullName}} -> {FullName} or {{Index}} -> {Index}
        result = re.sub(r"\{\{\s*([^{}\r\n]+?)\s*\}\}", r"{\1}", result)

        # Replace special index variables (case-insensitive with optional whitespace)
        result = re.sub(r"\{\s*Index\s*\}", f"{index:04d}", result, flags=re.IGNORECASE)
        result = re.sub(r"\{\s*Row\s*\}", str(index), result, flags=re.IGNORECASE)

        # Build case-insensitive and normalized lookup for record keys
        norm_lookup = {}
        for key, val in record.items():
            norm_lookup[key.lower()] = str(val)
            norm_k = re.sub(r"[\W_]+", "", key).lower()
            norm_lookup[norm_k] = str(val)

        import datetime
        now = datetime.datetime.now()

        def _sub_token(m):
            tok = m.group(1).strip()
            tok_lower = tok.lower()
            tok_clean = re.sub(r"[\W_]+", "", tok).lower()
            if tok_lower in norm_lookup:
                return norm_lookup[tok_lower]
            if tok_clean in norm_lookup:
                return norm_lookup[tok_clean]
            if tok_lower in ("date", "today"):
                return now.strftime("%Y-%m-%d")
            if tok_lower == "timestamp":
                return now.strftime("%Y%m%d_%H%M%S")
            return ""

        result = re.sub(r"\{([^{}\r\n]+)\}", _sub_token, result)

        # Clean base name
        base_name = DynamicNamer.sanitize(result)
        if not base_name:
            base_name = f"Document_{index:04d}"

        if not ext.startswith("."):
            ext = "." + ext

        # Prevent duplicate extensions (e.g. pattern ending in .pdf)
        if base_name.lower().endswith(ext.lower()):
            base_name = base_name[:-len(ext)].rstrip(" ._")
            if not base_name:
                base_name = f"Document_{index:04d}"

        # Re-check DOS device names on final stem
        if base_name.upper() in WINDOWS_RESERVED_NAMES:
            base_name = f"Doc_{base_name}"

        candidate = f"{base_name}{ext}"
        counter = 1

        # Check against both disk and current batch allocations
        full_path = os.path.join(output_dir, candidate)
        while candidate in allocated_names or os.path.exists(full_path):
            candidate = f"{base_name}_{counter:02d}{ext}"
            full_path = os.path.join(output_dir, candidate)
            counter += 1

        allocated_names.add(candidate)
        return candidate
