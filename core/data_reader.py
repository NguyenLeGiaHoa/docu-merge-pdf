"""
DocuMergePDF - Tabular Recipient Data Reader & Auto-Mapper
Supports Excel (.xlsx, .xls) and CSV (.csv) with auto-encoding detection,
data normalization (dates, numbers, empty cells), and fuzzy column matching.
"""

import os
import csv
import re
import difflib
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd


class DataReader:
    """Reads spreadsheet records and matches them with template placeholders."""

    @staticmethod
    def read(file_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        Reads an Excel or CSV file safely in read-only mode and returns normalized rows.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Data file not found: {file_path}")

        if os.path.basename(file_path).startswith("~$"):
            raise ValueError(f"'{os.path.basename(file_path)}' is a temporary Microsoft Excel lock file. Please select the master spreadsheet instead.")

        ext = os.path.splitext(file_path)[1].lower()
        sheet_names = []

        try:
            if ext in [".xlsx", ".xls", ".xlsm"]:
                with pd.ExcelFile(file_path) as excel_file:
                    sheet_names = excel_file.sheet_names
                    if not sheet_names:
                        sheet_names = ["Sheet1"]
                    target_sheet = sheet_name if sheet_name and sheet_name in sheet_names else sheet_names[0]
                    try:
                        df = pd.read_excel(excel_file, sheet_name=target_sheet)
                    except Exception:
                        df = pd.DataFrame()
            elif ext == ".csv":
                sheet_names = ["CSV Data"]
                target_sheet = "CSV Data"
                df = DataReader._read_csv_robust(file_path)
            else:
                raise ValueError(f"Unsupported file format: {ext}. Please provide .xlsx, .xls, or .csv")
        except PermissionError:
            raise PermissionError(f"The spreadsheet '{os.path.basename(file_path)}' is currently open in Microsoft Excel or locked by another program. Please close it and try again.")

        # Clean column names (strip whitespace and ensure uniqueness)
        raw_cols = [str(c).strip() for c in df.columns]
        cleaned_columns = []
        col_counts = {}
        for c in raw_cols:
            if c not in col_counts:
                col_counts[c] = 1
                cleaned_columns.append(c)
            else:
                col_counts[c] += 1
                cleaned_columns.append(f"{c}_{col_counts[c]}")
        df.columns = cleaned_columns

        # Drop trailing all-empty rows
        df = df.dropna(how="all")

        # Convert records with intelligent normalization
        normalized_records: List[Dict[str, str]] = []
        for _, row in df.iterrows():
            record = {}
            for col in cleaned_columns:
                val = row[col]
                record[col] = DataReader._normalize_value(val)
            # Skip record if entirely blank
            if any(record.values()):
                normalized_records.append(record)

        return {
            "path": file_path,
            "filename": os.path.basename(file_path),
            "columns": cleaned_columns,
            "row_count": len(normalized_records),
            "sheets": sheet_names,
            "active_sheet": target_sheet,
            "records": normalized_records,
            "sample_row": normalized_records[0] if normalized_records else {},
        }

    @staticmethod
    def _read_csv_robust(file_path: str) -> pd.DataFrame:
        """Reads CSV with automatic encoding (including UTF-16 from Windows/Excel) and delimiter detection."""
        # Check raw bytes for UTF-16 BOM or alternating null bytes (Windows PowerShell / Excel Unicode CSV)
        with open(file_path, "rb") as bf:
            raw_prefix = bf.read(4096)

        best_encoding = None
        if raw_prefix.startswith(b"\xff\xfe") or b"\x00" in raw_prefix[:64]:
            best_encoding = "utf-16"
        elif raw_prefix.startswith(b"\xfe\xff"):
            best_encoding = "utf-16-be"
        elif raw_prefix.startswith(b"\xef\xbb\xbf"):
            best_encoding = "utf-8-sig"

        sample_str = ""
        if best_encoding:
            try:
                sample_str = raw_prefix.decode(best_encoding, errors="replace")[:4096]
            except Exception:
                best_encoding = None

        if not best_encoding:
            encodings = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
            for enc in encodings:
                try:
                    sample_str = raw_prefix.decode(enc)
                    best_encoding = enc
                    break
                except (UnicodeDecodeError, Exception):
                    continue
            if not best_encoding:
                best_encoding = "utf-8"
                sample_str = raw_prefix.decode("utf-8", errors="replace")

        delimiters = [",", ";", "\t", "|"]
        # Detect delimiter using csv.Sniffer with all valid delimiters
        try:
            sniffer = csv.Sniffer()
            dialect = sniffer.sniff(sample_str[:2048], delimiters=",;|\t")
            delimiter = dialect.delimiter
        except Exception:
            # Fallback delimiter heuristics
            first_line = sample_str.split("\n")[0] if sample_str else ""
            counts = {d: first_line.count(d) for d in delimiters}
            delimiter = max(counts, key=counts.get)
            if counts[delimiter] == 0:
                delimiter = ","

        try:
            return pd.read_csv(file_path, encoding=best_encoding, sep=delimiter)
        except pd.errors.EmptyDataError:
            return pd.DataFrame()

    @staticmethod
    def _normalize_value(val: Any) -> str:
        """Normalizes numbers, dates, floats, and empty cells into clean display strings."""
        if val is None:
            return ""

        if isinstance(val, pd.Series):
            val = val.iloc[0] if len(val) > 0 else ""

        try:
            if pd.isna(val):
                return ""
        except Exception:
            pass

        # Float/int check preserving precision and whole number formatting
        if isinstance(val, (float, int)):
            if isinstance(val, float) and val.is_integer():
                return str(int(val))
            if isinstance(val, float):
                # Avoid premature scientific notation for financial numbers, IDs, and phone numbers
                if abs(val) < 1e16 and abs(val) > 1e-6:
                    return f"{round(val, 6):.6f}".rstrip("0").rstrip(".")
            return str(val)

        # Dates / Timestamps
        if hasattr(val, "strftime"):
            # Check if it has non-zero time components
            if hasattr(val, "hour") and (val.hour != 0 or val.minute != 0 or val.second != 0):
                return val.strftime("%Y-%m-%d %H:%M")
            return val.strftime("%Y-%m-%d")

        return str(val).strip()

    @staticmethod
    def auto_map(placeholders: List[str], columns: List[str]) -> Dict[str, Optional[str]]:
        """
        Auto-maps template placeholders to spreadsheet columns using exact, normalized,
        word-token, and fuzzy similarity heuristics with full international Unicode support.
        Returns dict: {placeholder: matching_column_name_or_None}
        """
        mapping: Dict[str, Optional[str]] = {}
        col_clean_map = {re.sub(r"[\W_]+", "", c).lower(): c for c in columns}

        for p in placeholders:
            matched_col = None
            p_strip = p.strip()
            p_clean = re.sub(r"[\W_]+", "", p_strip).lower()
            p_words = set(re.findall(r"[^\W_]+", p_strip.lower()))

            # 1. Exact case-insensitive match
            for c in columns:
                if p_strip.lower() == c.strip().lower():
                    matched_col = c
                    break

            # 2. Normalized alphanumeric match (e.g. Student_ID -> StudentID, Full_Name -> Full Name, First_Name -> First Name)
            if not matched_col and p_clean in col_clean_map:
                matched_col = col_clean_map[p_clean]

            # 3. Word-token set match (e.g. "Full Name" -> "full_name" or "Student ID" -> "ID_Student")
            if not matched_col and p_words:
                for c in columns:
                    c_words = set(re.findall(r"[^\W_]+", c.lower()))
                    if p_words and p_words == c_words:
                        matched_col = c
                        break

            # 4. Word-boundary token containment (e.g. "GradDate" -> "Graduation_Date", "ID" in ["Student_ID"])
            # Match if any token in c_words equals or starts with p_clean, or fuzzy acronym match
            if not matched_col:
                for c in columns:
                    c_words = set(re.findall(r"[^\W_]+", c.lower()))
                    if p_clean in c_words or any(w.startswith(p_clean) or p_clean.startswith(w) for w in c_words if len(w) >= 3):
                        matched_col = c
                        break

            # 5. Fuzzy similarity matching using SequenceMatcher
            if not matched_col and columns:
                best_ratio = 0.0
                best_c = None
                for c in columns:
                    c_clean = re.sub(r"[\W_]+", "", c).lower()
                    ratio = difflib.SequenceMatcher(None, p_clean, c_clean).ratio()
                    if ratio > best_ratio:
                        best_ratio = ratio
                        best_c = c
                if best_ratio >= 0.55:
                    matched_col = best_c

            mapping[p] = matched_col

        return mapping
