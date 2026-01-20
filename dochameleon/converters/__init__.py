"""
Converters module for Dochameleon.
"""

from .latex import compile_latex_to_pdf, clean_latex_auxiliary_files
from .pdf import (
    convert_pdf_to_docx_enhanced,
    make_docx_native,
    enhance_docx_formatting,
    enhance_docx_code_blocks,
    add_hyperlink,
    extract_pdf_hyperlinks,
)
from .docx import convert_docx_to_pdf

__all__ = [
    'compile_latex_to_pdf',
    'clean_latex_auxiliary_files',
    'convert_pdf_to_docx_enhanced',
    'make_docx_native',
    'enhance_docx_formatting',
    'enhance_docx_code_blocks',
    'convert_docx_to_pdf',
    'add_hyperlink',
    'extract_pdf_hyperlinks',
]
