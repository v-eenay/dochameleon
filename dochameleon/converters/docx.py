"""
DOCX to PDF conversion utilities.
"""

from pathlib import Path
from typing import Tuple, Union


def convert_docx_to_pdf(docx_file: Path, output_dir: Path) -> Tuple[bool, Union[Path, str]]:
    """
    Convert DOCX to PDF using docx2pdf (requires MS Word on Windows).
    """
    try:
        import docx2pdf
        
        output_path = output_dir / (docx_file.stem + ".pdf")
        
        docx2pdf.convert(str(docx_file), str(output_path))
        
        if output_path.exists():
            return True, output_path
        else:
            return False, "PDF file was not created"
            
    except Exception as e:
        error_msg = str(e)
        if "win32com" in error_msg.lower() or "word" in error_msg.lower():
            return False, "Microsoft Word is required for DOCX to PDF conversion on Windows"
        return False, error_msg
