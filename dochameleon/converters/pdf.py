"""
PDF to DOCX conversion utilities - Clean native DOCX output.
"""

from pathlib import Path
from typing import Tuple, Union


def convert_pdf_to_docx_enhanced(pdf_file: Path, output_dir: Path) -> Tuple[bool, Union[Path, str]]:
    """
    Convert PDF to DOCX with clean, native Word-like output.
    
    The goal is to produce documents that look like they were
    originally created in Microsoft Word - clean, professional,
    without unnecessary boxes or frames.
    """
    from pdf2docx import Converter
    
    output_path = output_dir / (pdf_file.stem + ".docx")
    
    try:
        cv = Converter(str(pdf_file))
        
        # Convert with settings optimized for clean Word-like output
        cv.convert(
            str(output_path),
            # Table detection - be conservative to avoid false positives
            connected_border_tolerance=0.5,
            min_section_height=20,
            
            # Text flow - natural paragraph breaks
            line_overlap_threshold=0.9,
            line_break_width_ratio=0.5,
            line_break_free_space_ratio=0.1,
            new_paragraph_free_space_ratio=0.85,
            
            # Images - standard positioning
            float_image_ignorable_gap=5,
            page_margin_factor_top=0.5,
            page_margin_factor_bottom=0.5,
        )
        cv.close()
        
        if output_path.exists():
            # Clean up the document to look more native
            make_docx_native(output_path)
            return True, output_path
        else:
            return False, "DOCX file was not created"
            
    except Exception as e:
        return False, str(e)


def make_docx_native(docx_path: Path):
    """
    Post-process DOCX to create clean, native Word-like appearance.
    
    Removes unnecessary formatting artifacts and applies clean styling
    that makes the document look like it was created in Word.
    """
    try:
        from docx import Document
        from docx.shared import Pt, Inches
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
        
        doc = Document(str(docx_path))
        
        # ============================================
        # CLEAN UP DOCUMENT STYLES
        # ============================================
        setup_native_styles(doc)
        
        # ============================================
        # CLEAN PARAGRAPHS - Remove box artifacts
        # ============================================
        clean_paragraphs(doc)
        
        # ============================================
        # CLEAN TABLES - Only keep real tables
        # ============================================
        clean_tables(doc)
        
        # ============================================
        # APPLY CLEAN FORMATTING
        # ============================================
        apply_clean_formatting(doc)
        
        doc.save(str(docx_path))
        
    except Exception as e:
        print(f"  Note: Could not clean document: {e}")


def setup_native_styles(doc):
    """
    Set up clean, native Word styles.
    """
    from docx.shared import Pt, RGBColor
    from docx.enum.style import WD_STYLE_TYPE
    
    # Set default document font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    # Clean paragraph formatting
    paragraph_format = style.paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(8)
    paragraph_format.line_spacing = 1.15


def clean_paragraphs(doc):
    """
    Clean paragraph formatting - remove unnecessary borders and boxes.
    """
    from docx.oxml.ns import qn
    
    for paragraph in doc.paragraphs:
        pPr = paragraph._p.find(qn('w:pPr'))
        if pPr is not None:
            # Remove paragraph borders (boxes around text)
            pBdr = pPr.find(qn('w:pBdr'))
            if pBdr is not None:
                # Check if this looks like an intentional callout/note box
                if not _is_intentional_box(paragraph):
                    pPr.remove(pBdr)
            
            # Remove unnecessary shading on regular paragraphs
            shd = pPr.find(qn('w:shd'))
            if shd is not None:
                fill = shd.get(qn('w:fill'))
                # Only remove if it's a light/subtle background (probably artifact)
                if fill and fill.upper() in ('F5F5F5', 'F0F0F0', 'EFEFEF', 'FAFAFA', 'E0E0E0'):
                    if not _is_code_block(paragraph):
                        pPr.remove(shd)
        
        # Clean runs within paragraph
        clean_paragraph_runs(paragraph)


def _is_intentional_box(paragraph) -> bool:
    """
    Detect if a paragraph box is intentional (like a callout, warning, note).
    """
    text = paragraph.text.lower().strip()
    intentional_markers = [
        'note:', 'warning:', 'caution:', 'important:', 'tip:',
        '⚠', '📝', '💡', '❗', '✓', '✗'
    ]
    return any(marker in text for marker in intentional_markers)


def _is_code_block(paragraph) -> bool:
    """
    Detect if paragraph is a code block that should keep formatting.
    """
    text = paragraph.text.strip()
    
    # Check for code-like content
    code_patterns = [
        'def ', 'class ', 'import ', 'from ', 'return ',
        'function ', 'const ', 'let ', 'var ',
        '#!/', '<?php', '<html', '```',
        '>>>', '$ ', 'pip install', 'npm install'
    ]
    
    # Check font
    for run in paragraph.runs:
        font_name = (run.font.name or '').lower()
        if any(mono in font_name for mono in ['courier', 'consolas', 'mono', 'code']):
            return True
    
    return any(pattern in text for pattern in code_patterns)


def clean_paragraph_runs(paragraph):
    """
    Clean text runs - remove unnecessary background shading.
    """
    from docx.oxml.ns import qn
    
    for run in paragraph.runs:
        rPr = run._element.find(qn('w:rPr'))
        if rPr is not None:
            # Only remove shading from non-code text
            font_name = (run.font.name or '').lower()
            is_code_font = any(mono in font_name for mono in ['courier', 'consolas', 'mono', 'code'])
            
            if not is_code_font:
                shd = rPr.find(qn('w:shd'))
                if shd is not None:
                    fill = shd.get(qn('w:fill'))
                    # Remove subtle gray backgrounds (artifacts)
                    if fill and fill.upper() in ('F5F5F5', 'F0F0F0', 'EFEFEF', 'FAFAFA', 'E0E0E0', 'F8F8F8'):
                        rPr.remove(shd)


def clean_tables(doc):
    """
    Clean tables - remove single-cell "fake" tables used as boxes.
    Convert real tables to clean Word table style.
    """
    from docx.oxml.ns import qn
    from docx.shared import Pt
    
    for table in doc.tables:
        rows = len(table.rows)
        cols = len(table.columns) if table.rows else 0
        
        # Check if this is a real table or a fake box
        if rows == 1 and cols == 1:
            # Single cell - likely a box artifact, clean it
            cell = table.rows[0].cells[0]
            clean_single_cell_table(table, cell)
        else:
            # Real table - apply clean styling
            apply_clean_table_style(table)


def clean_single_cell_table(table, cell):
    """
    Clean single-cell tables that are just boxes.
    Remove borders unless it's intentionally a callout.
    """
    from docx.oxml.ns import qn
    
    cell_text = cell.text.lower().strip()
    
    # Keep borders for intentional callouts
    if any(marker in cell_text for marker in ['note:', 'warning:', 'caution:', 'important:', 'tip:']):
        return
    
    # Remove table borders for regular single-cell "boxes"
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is not None:
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            tblPr.remove(tblBorders)
    
    # Remove cell borders
    tc = cell._tc
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is not None:
        tcBorders = tcPr.find(qn('w:tcBorders'))
        if tcBorders is not None:
            tcPr.remove(tcBorders)
        
        # Remove cell shading
        shd = tcPr.find(qn('w:shd'))
        if shd is not None:
            fill = shd.get(qn('w:fill'))
            if fill and fill.upper() in ('F5F5F5', 'F0F0F0', 'EFEFEF', 'FAFAFA', 'E0E0E0', 'F8F8F8', 'FFFFFF'):
                tcPr.remove(shd)


def apply_clean_table_style(table):
    """
    Apply clean, professional Word table styling.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    
    # Remove any indent/shadow effects
    tblInd = tblPr.find(qn('w:tblInd'))
    if tblInd is not None:
        tblPr.remove(tblInd)
    
    # Set clean borders
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is None:
        tblBorders = OxmlElement('w:tblBorders')
        tblPr.append(tblBorders)
    
    # Apply subtle, professional borders
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = tblBorders.find(qn(f'w:{border_name}'))
        if border is None:
            border = OxmlElement(f'w:{border_name}')
            tblBorders.append(border)
        
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')  # Thin border
        border.set(qn('w:color'), 'BFBFBF')  # Light gray - professional look
        border.set(qn('w:space'), '0')


def apply_clean_formatting(doc):
    """
    Apply final clean formatting touches.
    """
    from docx.shared import Pt
    from docx.oxml.ns import qn
    
    for paragraph in doc.paragraphs:
        # Ensure consistent paragraph spacing
        if paragraph.paragraph_format.space_after is None:
            paragraph.paragraph_format.space_after = Pt(8)
        
        # Clean up any remaining font issues
        for run in paragraph.runs:
            # Normalize fonts - use Calibri for body text
            font_name = (run.font.name or '').lower()
            is_mono = any(mono in font_name for mono in ['courier', 'consolas', 'mono', 'code', 'menlo'])
            
            if is_mono:
                # Keep as Consolas for code
                run.font.name = 'Consolas'
            elif not run.font.name or run.font.name == 'Times New Roman':
                # Default to Calibri for cleaner look
                run.font.name = 'Calibri'


# ============================================================
# ENHANCED FORMATTING FUNCTIONS (for documents that need it)
# ============================================================

def enhance_docx_formatting(docx_path: Path):
    """
    Alias for make_docx_native.
    Kept for backwards compatibility.
    """
    make_docx_native(docx_path)


def enhance_docx_code_blocks(docx_path: Path):
    """
    Legacy function - now calls make_docx_native.
    Kept for backwards compatibility.
    """
    make_docx_native(docx_path)
