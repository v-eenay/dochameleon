"""
PDF to DOCX conversion utilities - Clean native DOCX output.

This module focuses on producing DOCX files that look like they were
originally created in Microsoft Word, not converted from PDF.
"""

from pathlib import Path
from typing import Tuple, Union, List


def convert_pdf_to_docx_enhanced(pdf_file: Path, output_dir: Path) -> Tuple[bool, Union[Path, str]]:
    """
    Convert PDF to DOCX with clean, native Word-like output.
    """
    from pdf2docx import Converter
    
    output_path = output_dir / (pdf_file.stem + ".docx")
    
    try:
        cv = Converter(str(pdf_file))
        
        # Convert with minimal table detection to avoid fake tables
        cv.convert(
            str(output_path),
            # CRITICAL: High tolerance = fewer false table detections
            connected_border_tolerance=1.0,
            min_section_height=30,
            
            # Text flow
            line_overlap_threshold=0.9,
            line_break_width_ratio=0.5,
            line_break_free_space_ratio=0.15,
            new_paragraph_free_space_ratio=0.85,
            
            # Images
            float_image_ignorable_gap=10,
            
            # Page margins - preserve original
            page_margin_factor_top=0.0,
            page_margin_factor_bottom=0.0,
        )
        cv.close()
        
        if output_path.exists():
            # Aggressively clean the document
            make_docx_native(output_path)
            return True, output_path
        else:
            return False, "DOCX file was not created"
            
    except Exception as e:
        return False, str(e)


def make_docx_native(docx_path: Path):
    """
    Aggressively clean DOCX to look like native Word document.
    """
    try:
        from docx import Document
        from docx.shared import Pt, Inches, Twips
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        
        doc = Document(str(docx_path))
        
        # ============================================
        # 1. SET PROPER PAGE MARGINS
        # ============================================
        set_page_margins(doc)
        
        # ============================================
        # 2. REMOVE ALL FAKE TABLES (wrapper tables)
        # ============================================
        remove_wrapper_tables(doc)
        
        # ============================================
        # 3. REMOVE ALL PARAGRAPH BORDERS
        # ============================================
        remove_all_paragraph_borders(doc)
        
        # ============================================
        # 4. CLEAN HEADING STYLES
        # ============================================
        clean_headings(doc)
        
        # ============================================
        # 5. APPLY CLEAN DOCUMENT STYLES
        # ============================================
        apply_native_styles(doc)
        
        doc.save(str(docx_path))
        
    except Exception as e:
        print(f"  Note: Could not clean document: {e}")


def set_page_margins(doc):
    """
    Set proper Word-standard page margins (1 inch all around).
    """
    from docx.shared import Inches
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    for section in doc.sections:
        # Standard Word margins: 1 inch all around
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        
        # Also set header/footer distances
        section.header_distance = Inches(0.5)
        section.footer_distance = Inches(0.5)


def remove_wrapper_tables(doc):
    """
    Remove tables that are just wrappers (not real data tables).
    
    Wrapper tables are typically:
    - Single row tables
    - Tables where cells contain mostly paragraphs (not tabular data)
    - Tables used as layout containers
    """
    from docx.oxml.ns import qn
    
    tables_to_process = list(doc.tables)
    
    for table in tables_to_process:
        rows = len(table.rows)
        cols = len(table.columns) if table.rows else 0
        
        # Determine if this is a real table or a wrapper
        is_wrapper = False
        
        # Single cell = definitely a wrapper
        if rows == 1 and cols == 1:
            is_wrapper = True
        
        # Single row with few columns = likely a wrapper/layout
        elif rows == 1 and cols <= 2:
            is_wrapper = True
        
        # Check if it looks like a heading wrapper
        elif rows <= 2:
            cell_text = ""
            for row in table.rows:
                for cell in row.cells:
                    cell_text += cell.text.strip()
            # Short text in few rows = likely wrapper
            if len(cell_text) < 200 and rows == 1:
                is_wrapper = True
        
        if is_wrapper:
            # Remove all borders from wrapper tables
            remove_table_borders(table)
            # Also remove cell borders and shading
            for row in table.rows:
                for cell in row.cells:
                    remove_cell_formatting(cell)
        else:
            # Real table - apply clean styling
            apply_clean_table_style(table)


def remove_table_borders(table):
    """
    Remove all borders from a table.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    
    if tblPr is not None:
        # Remove table borders
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            tblPr.remove(tblBorders)
        
        # Add explicit "no borders"
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'nil')
            tblBorders.append(border)
        tblPr.append(tblBorders)
        
        # Remove table indent
        tblInd = tblPr.find(qn('w:tblInd'))
        if tblInd is not None:
            tblPr.remove(tblInd)


def remove_cell_formatting(cell):
    """
    Remove borders and shading from a cell.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    
    if tcPr is not None:
        # Remove cell borders
        tcBorders = tcPr.find(qn('w:tcBorders'))
        if tcBorders is not None:
            tcPr.remove(tcBorders)
        
        # Add explicit "no borders"
        tcBorders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'nil')
            tcBorders.append(border)
        tcPr.append(tcBorders)
        
        # Remove shading
        shd = tcPr.find(qn('w:shd'))
        if shd is not None:
            tcPr.remove(shd)
    
    # Also clean paragraphs within cell
    for paragraph in cell.paragraphs:
        remove_paragraph_borders(paragraph)


def remove_all_paragraph_borders(doc):
    """
    Remove ALL borders from ALL paragraphs.
    This is aggressive but necessary to get clean output.
    """
    from docx.oxml.ns import qn
    
    for paragraph in doc.paragraphs:
        remove_paragraph_borders(paragraph)
    
    # Also remove from table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    remove_paragraph_borders(paragraph)


def remove_paragraph_borders(paragraph):
    """
    Remove all borders and unnecessary shading from a paragraph.
    """
    from docx.oxml.ns import qn
    
    pPr = paragraph._p.find(qn('w:pPr'))
    if pPr is not None:
        # Remove paragraph borders completely
        pBdr = pPr.find(qn('w:pBdr'))
        if pBdr is not None:
            pPr.remove(pBdr)
        
        # Remove paragraph shading (backgrounds)
        shd = pPr.find(qn('w:shd'))
        if shd is not None:
            pPr.remove(shd)
        
        # Remove frames
        framePr = pPr.find(qn('w:framePr'))
        if framePr is not None:
            pPr.remove(framePr)
    
    # Clean runs too
    for run in paragraph.runs:
        rPr = run._element.find(qn('w:rPr'))
        if rPr is not None:
            # Remove run shading (text backgrounds) except for intentional highlights
            shd = rPr.find(qn('w:shd'))
            if shd is not None:
                fill = shd.get(qn('w:fill'))
                # Remove gray backgrounds (artifacts), keep colored highlights
                if fill and fill.upper() in ('F5F5F5', 'F0F0F0', 'EFEFEF', 'FAFAFA', 
                                               'E0E0E0', 'F8F8F8', 'E8E8E8', 'D0D0D0',
                                               'FFFFFF', 'auto'):
                    rPr.remove(shd)


def clean_headings(doc):
    """
    Ensure headings are clean and properly styled.
    """
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    
    for paragraph in doc.paragraphs:
        # Detect headings by style or formatting
        is_heading = False
        style_name = paragraph.style.name if paragraph.style else ''
        
        if 'Heading' in style_name or 'Title' in style_name:
            is_heading = True
        else:
            # Check if it looks like a heading (bold, larger font, short text)
            text = paragraph.text.strip()
            if len(text) < 100 and len(text) > 0:
                for run in paragraph.runs:
                    if run.bold and run.font.size and run.font.size >= Pt(12):
                        is_heading = True
                        break
        
        if is_heading:
            # Ensure heading has no borders or boxes
            pPr = paragraph._p.find(qn('w:pPr'))
            if pPr is not None:
                # Remove any borders
                pBdr = pPr.find(qn('w:pBdr'))
                if pBdr is not None:
                    pPr.remove(pBdr)
                
                # Remove shading
                shd = pPr.find(qn('w:shd'))
                if shd is not None:
                    pPr.remove(shd)
            
            # Set proper heading spacing
            paragraph.paragraph_format.space_before = Pt(12)
            paragraph.paragraph_format.space_after = Pt(6)


def apply_clean_table_style(table):
    """
    Apply clean styling to real data tables.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    # Remove existing borders
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is not None:
        tblPr.remove(tblBorders)
    
    # Add clean, subtle borders
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), 'BFBFBF')  # Light gray
        border.set(qn('w:space'), '0')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    
    # Clean cell formatting
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                # Remove cell shading
                shd = tcPr.find(qn('w:shd'))
                if shd is not None:
                    fill = shd.get(qn('w:fill'))
                    # Remove gray backgrounds
                    if fill and fill.upper() in ('F5F5F5', 'F0F0F0', 'EFEFEF', 'FAFAFA',
                                                   'E0E0E0', 'F8F8F8', 'FFFFFF'):
                        tcPr.remove(shd)


def apply_native_styles(doc):
    """
    Apply clean, native Word styling to the entire document.
    """
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    
    # Set Normal style
    try:
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.line_spacing = 1.15
    except:
        pass
    
    # Clean all paragraphs
    for paragraph in doc.paragraphs:
        # Set proper spacing
        if paragraph.paragraph_format.space_after is None:
            paragraph.paragraph_format.space_after = Pt(8)
        
        # Clean fonts
        for run in paragraph.runs:
            font_name = (run.font.name or '').lower()
            is_mono = any(m in font_name for m in ['courier', 'consolas', 'mono', 'code', 'menlo'])
            
            if is_mono:
                run.font.name = 'Consolas'
            elif not run.font.name or 'times' in font_name.lower():
                run.font.name = 'Calibri'


# ============================================================
# BACKWARDS COMPATIBILITY
# ============================================================

def enhance_docx_formatting(docx_path: Path):
    """Alias for backwards compatibility."""
    make_docx_native(docx_path)


def enhance_docx_code_blocks(docx_path: Path):
    """Alias for backwards compatibility."""
    make_docx_native(docx_path)
