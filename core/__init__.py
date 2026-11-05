"""
DocuMergePDF Core Engine Package
"""

from .licensing import LicenseSecurity
from .legal_shield import LegalShield, LEGAL_DISCLAIMER_SUMMARY
from .template_inspector import TemplateInspector
from .data_reader import DataReader
from .dynamic_namer import DynamicNamer
from .document_merger import DocumentMerger
from .pdf_converter import PDFConverter

__all__ = [
    "LicenseSecurity",
    "LegalShield",
    "LEGAL_DISCLAIMER_SUMMARY",
    "TemplateInspector",
    "DataReader",
    "DynamicNamer",
    "DocumentMerger",
    "PDFConverter",
]

