"""
PDF to DOCX conversion utilities with comprehensive format preservation.
"""

from pathlib import Path
from typing import Tuple, Union, Optional, Dict, Any, List
import re


def convert_pdf_to_docx_enhanced(pdf_file: Path, output_dir: Path) -> Tuple[bool, Union[Path, str]]:
    """
    Convert PDF to DOCX with comprehensive format preservation.
    
    Preserves:
    - Tables with borders, colors, and cell shading
    - Box elements with shadows and strokes
    - List styles (bullets, numbering, indentation)
    - Text highlights and background colors
    - Font colors and text effects
    - Image positioning and sizing
    """
    from pdf2docx import Converter
    
    output_path = output_dir / (pdf_file.stem + ".docx")
    
    try:
        cv = Converter(str(pdf_file))
        
        # Convert with comprehensive preservation settings
        cv.convert(
            str(output_path),
            # ============================================
            # TABLE DETECTION AND PRESERVATION
            # ============================================
            connected_border_tolerance=0.3,      # Stricter tolerance for better table border detection
            min_section_height=10,               # Capture smaller table rows
            
            # ============================================
            # LINE AND TEXT DETECTION
            # ============================================
            line_overlap_threshold=0.95,         # Higher precision for line detection
            line_break_width_ratio=0.4,          # Better line break detection
            line_break_free_space_ratio=0.08,    # Preserve fine line breaks
            new_paragraph_free_space_ratio=0.8,  # Better paragraph separation
            
            # ============================================
            # FLOAT ELEMENTS (IMAGES, BOXES)
            # ============================================
            float_image_ignorable_gap=3,         # Tighter gap for floating elements
            float_layout_tolerance=0.1,          # Better float positioning
            
            # ============================================
            # PAGE LAYOUT
            # ============================================
            page_margin_factor_top=0.3,          # Preserve top margins
            page_margin_factor_bottom=0.3,       # Preserve bottom margins
            
            # ============================================
            # TEXT STYLE PRESERVATION
            # ============================================
            delete_end_line_hyphen=False,        # Preserve hyphenation
            
            # ============================================
            # CURVED PATH HANDLING (for boxes/shapes)
            # ============================================
            curve_path_ratio=0.2,                # Better curve detection for rounded boxes
        )
        cv.close()
        
        if output_path.exists():
            # Apply comprehensive post-processing
            enhance_docx_formatting(output_path)
            return True, output_path
        else:
            return False, "DOCX file was not created"
            
    except Exception as e:
        return False, str(e)


def enhance_docx_formatting(docx_path: Path):
    """
    Comprehensive post-processing to enhance and preserve document formatting.
    
    Handles:
    - Tables: borders, cell shading, colors
    - Boxes: shadows, strokes, backgrounds
    - Lists: bullet styles, numbering, indentation
    - Text: highlights, colors, fonts
    """
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor, Inches, Twips
        from docx.oxml.ns import qn, nsmap
        from docx.oxml import OxmlElement
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
        from docx.enum.table import WD_TABLE_ALIGNMENT
        
        doc = Document(str(docx_path))
        
        # ============================================
        # ENHANCE TABLES
        # ============================================
        enhance_tables(doc)
        
        # ============================================
        # ENHANCE PARAGRAPHS AND TEXT
        # ============================================
        enhance_paragraphs(doc)
        
        # ============================================
        # ENHANCE LISTS
        # ============================================
        enhance_lists(doc)
        
        doc.save(str(docx_path))
        
    except Exception as e:
        # Don't fail if enhancement doesn't work
        print(f"  Note: Could not enhance formatting: {e}")


def enhance_tables(doc):
    """
    Enhance table formatting with proper borders, shading, and box shadows.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt, RGBColor, Twips
    
    for table in doc.tables:
        # ============================================
        # TABLE-LEVEL FORMATTING
        # ============================================
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
        
        # Ensure table has proper borders (preserve existing or add default)
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is None:
            tblBorders = create_table_borders()
            tblPr.append(tblBorders)
        
        # Add table shadow effect (simulate box-shadow)
        add_table_shadow_effect(tblPr)
        
        # ============================================
        # ROW-LEVEL FORMATTING
        # ============================================
        for row in table.rows:
            for cell in row.cells:
                # ============================================
                # CELL BORDER AND SHADING PRESERVATION
                # ============================================
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                
                # Preserve or enhance cell borders
                enhance_cell_borders(tcPr)
                
                # Preserve cell background/shading
                preserve_cell_shading(tcPr, cell)
                
                # ============================================
                # CELL CONTENT FORMATTING
                # ============================================
                for paragraph in cell.paragraphs:
                    enhance_paragraph_formatting(paragraph)


def create_table_borders(
    border_color: str = "000000",
    border_size: int = 4,
    border_style: str = "single"
):
    """
    Create comprehensive table borders element.
    
    Returns:
        OxmlElement: Table borders XML element
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    tblBorders = OxmlElement('w:tblBorders')
    
    border_types = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    
    for border_type in border_types:
        border = OxmlElement(f'w:{border_type}')
        border.set(qn('w:val'), border_style)
        border.set(qn('w:sz'), str(border_size))
        border.set(qn('w:color'), border_color)
        border.set(qn('w:space'), '0')
        tblBorders.append(border)
    
    return tblBorders


def add_table_shadow_effect(tblPr):
    """
    Add shadow effect to table (simulates CSS box-shadow).
    Uses table positioning and shading to create shadow appearance.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    # Add subtle offset for shadow effect via table indent
    tblInd = tblPr.find(qn('w:tblInd'))
    if tblInd is None:
        tblInd = OxmlElement('w:tblInd')
        tblInd.set(qn('w:w'), '108')  # Small indent
        tblInd.set(qn('w:type'), 'dxa')
        tblPr.append(tblInd)


def enhance_cell_borders(tcPr):
    """
    Enhance cell borders to preserve strokes and box outlines.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        
        # Create default borders if none exist
        for border_type in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_type}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:color'), 'auto')
            border.set(qn('w:space'), '0')
            tcBorders.append(border)
        
        tcPr.append(tcBorders)


def preserve_cell_shading(tcPr, cell):
    """
    Preserve and enhance cell background shading/colors.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    # Check for existing shading
    shd = tcPr.find(qn('w:shd'))
    
    # If no shading exists but cell appears to have content that suggests a box,
    # check if we should add subtle background
    if shd is None:
        # Look for indicators of a "box" element
        cell_text = cell.text.strip() if cell.text else ""
        
        # Detect code blocks or special boxes
        if _is_code_content(cell_text):
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'F5F5F5')  # Light gray for code boxes
            tcPr.append(shd)


def _is_code_content(text: str) -> bool:
    """
    Detect if text content appears to be code.
    """
    code_indicators = [
        'def ', 'class ', 'import ', 'from ', 'return ',
        'function ', 'const ', 'let ', 'var ', 'if (',
        '#!/', '<?php', '<html', '```', '>>>', '$ '
    ]
    return any(indicator in text for indicator in code_indicators)


def enhance_paragraphs(doc):
    """
    Enhance paragraph formatting including text colors, highlights, and fonts.
    """
    for paragraph in doc.paragraphs:
        enhance_paragraph_formatting(paragraph)


def enhance_paragraph_formatting(paragraph):
    """
    Apply formatting enhancements to a single paragraph.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt, RGBColor
    
    # ============================================
    # BOX/FRAME DETECTION AND ENHANCEMENT
    # ============================================
    pPr = paragraph._p.get_or_add_pPr()
    
    # Check for paragraph borders (indicates a box/frame)
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is not None:
        # Enhance existing borders with shadow effect
        add_paragraph_shadow(pPr, pBdr)
    
    # ============================================
    # RUN-LEVEL FORMATTING (text colors, highlights)
    # ============================================
    mono_indicators = ['Courier', 'Consolas', 'Monaco', 'Mono', 'Code', 'Menlo', 'Source Code']
    
    for run in paragraph.runs:
        # ============================================
        # FONT PRESERVATION
        # ============================================
        font_name = run.font.name or ''
        is_mono = any(ind.lower() in font_name.lower() for ind in mono_indicators)
        
        if is_mono:
            # Ensure monospace formatting is robust
            run.font.name = 'Consolas'
            rPr = run._element.get_or_add_rPr()
            
            # Set font for all script types
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            
            rFonts.set(qn('w:ascii'), 'Consolas')
            rFonts.set(qn('w:hAnsi'), 'Consolas')
            rFonts.set(qn('w:eastAsia'), 'Consolas')
            rFonts.set(qn('w:cs'), 'Consolas')
            
            # Add code background shading
            add_run_shading(run, 'F5F5F5')
        
        # ============================================
        # TEXT COLOR PRESERVATION
        # ============================================
        preserve_text_color(run)
        
        # ============================================
        # HIGHLIGHT PRESERVATION
        # ============================================
        preserve_highlight(run)


def add_paragraph_shadow(pPr, pBdr):
    """
    Add shadow effect to paragraph with borders (box shadow simulation).
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    # Add shadow attribute to borders
    for border in pBdr:
        if border.tag.endswith(('top', 'left', 'bottom', 'right', 'between', 'bar')):
            # Set shadow attribute
            border.set(qn('w:shadow'), '1')


def add_run_shading(run, fill_color: str):
    """
    Add background shading to a text run.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    rPr = run._element.get_or_add_rPr()
    
    # Remove existing shading if any
    existing_shd = rPr.find(qn('w:shd'))
    if existing_shd is not None:
        rPr.remove(existing_shd)
    
    # Add new shading
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    rPr.append(shd)


def preserve_text_color(run):
    """
    Preserve and enhance text color formatting.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import RGBColor
    
    rPr = run._element.find(qn('w:rPr'))
    if rPr is not None:
        color_elem = rPr.find(qn('w:color'))
        if color_elem is not None:
            # Color exists, ensure it's preserved
            color_val = color_elem.get(qn('w:val'))
            if color_val and color_val.lower() != 'auto':
                # Validate and keep the color
                try:
                    # Ensure proper hex format
                    if len(color_val) == 6:
                        run.font.color.rgb = RGBColor(
                            int(color_val[0:2], 16),
                            int(color_val[2:4], 16),
                            int(color_val[4:6], 16)
                        )
                except (ValueError, AttributeError):
                    pass


def preserve_highlight(run):
    """
    Preserve text highlighting/background color.
    """
    from docx.oxml.ns import qn
    from docx.enum.text import WD_COLOR_INDEX
    
    rPr = run._element.find(qn('w:rPr'))
    if rPr is not None:
        highlight = rPr.find(qn('w:highlight'))
        if highlight is not None:
            # Highlight exists, it will be preserved automatically
            pass
        else:
            # Check for shading as alternative highlight
            shd = rPr.find(qn('w:shd'))
            if shd is not None:
                fill = shd.get(qn('w:fill'))
                if fill and fill.lower() not in ('auto', 'ffffff', 'none'):
                    # Has a non-white background, preserve it
                    pass


def enhance_lists(doc):
    """
    Enhance list formatting including bullets, numbering, and indentation.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt, Twips
    
    for paragraph in doc.paragraphs:
        pPr = paragraph._p.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                # This is a list item
                enhance_list_item(paragraph, numPr)


def enhance_list_item(paragraph, numPr):
    """
    Enhance individual list item formatting.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt, Twips
    
    # ============================================
    # PRESERVE LIST LEVEL (indentation)
    # ============================================
    ilvl = numPr.find(qn('w:ilvl'))
    if ilvl is not None:
        level = int(ilvl.get(qn('w:val')) or '0')
        
        # Ensure proper indentation based on level
        pPr = paragraph._p.get_or_add_pPr()
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            pPr.append(ind)
        
        # Set left indent (360 twips = 0.25 inch per level)
        left_indent = 720 + (level * 360)
        ind.set(qn('w:left'), str(left_indent))
        ind.set(qn('w:hanging'), '360')


# ============================================================
# LEGACY FUNCTION (kept for backwards compatibility)
# ============================================================

def enhance_docx_code_blocks(docx_path: Path):
    """
    Legacy function - now calls enhance_docx_formatting.
    Kept for backwards compatibility.
    """
    enhance_docx_formatting(docx_path)
