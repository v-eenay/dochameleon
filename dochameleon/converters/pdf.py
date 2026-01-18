"""
PDF to DOCX conversion utilities.
"""

from pathlib import Path
from typing import Tuple, Union


def convert_pdf_to_docx_enhanced(pdf_file: Path, output_dir: Path) -> Tuple[bool, Union[Path, str]]:
    """
    Convert PDF to DOCX with enhanced settings for better preservation.
    """
    from pdf2docx import Converter
    
    output_path = output_dir / (pdf_file.stem + ".docx")
    
    try:
        cv = Converter(str(pdf_file))
        
        # Convert with enhanced settings
        cv.convert(
            str(output_path),
            # Parse settings for better quality
            connected_border_tolerance=0.5,  # Better table detection
            min_section_height=20,           # Don't miss small sections
            line_overlap_threshold=0.9,      # Better line detection
            line_break_width_ratio=0.5,      # Preserve line breaks
            line_break_free_space_ratio=0.1,
            new_paragraph_free_space_ratio=0.85,
            # Float image settings
            float_image_ignorable_gap=5,
            page_margin_factor_top=0.5,
            page_margin_factor_bottom=0.5,
        )
        cv.close()
        
        if output_path.exists():
            # Post-process to improve code block formatting
            enhance_docx_code_blocks(output_path)
            return True, output_path
        else:
            return False, "DOCX file was not created"
            
    except Exception as e:
        return False, str(e)


def enhance_docx_code_blocks(docx_path: Path):
    """
    Post-process DOCX to improve code block and monospace text formatting.
    """
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        
        doc = Document(str(docx_path))
        
        # Define monospace font indicators
        mono_indicators = ['Courier', 'Consolas', 'Monaco', 'Mono', 'Code']
        
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                # Check if this looks like code (monospace font)
                font_name = run.font.name or ''
                is_mono = any(ind.lower() in font_name.lower() for ind in mono_indicators)
                
                if is_mono:
                    # Ensure monospace formatting is preserved
                    run.font.name = 'Consolas'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Consolas')
                    
                    # Add light gray background for code
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:fill'), 'F5F5F5')
                    run._element.get_or_add_rPr().append(shading)
        
        # Process tables for code blocks
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            font_name = run.font.name or ''
                            is_mono = any(ind.lower() in font_name.lower() for ind in mono_indicators)
                            if is_mono:
                                run.font.name = 'Consolas'
        
        doc.save(str(docx_path))
        
    except Exception as e:
        # Don't fail if enhancement doesn't work
        print(f"  Note: Could not enhance code formatting: {e}")
