"""
Conversion pipeline functions for Dochameleon.
"""

import shutil
from pathlib import Path
from typing import Tuple

from .utils import find_files
from .converters import (
    compile_latex_to_pdf,
    clean_latex_auxiliary_files,
    convert_pdf_to_docx_enhanced,
    convert_docx_to_pdf,
)


# ============================================================
# SINGLE FILE CONVERSIONS
# ============================================================

def convert_single_tex_to_pdf(tex_file: Path, output_dir: Path) -> bool:
    """Convert a single .tex file to PDF."""
    print(f"\n📄 {tex_file.name}")
    result_ok, result = compile_latex_to_pdf(tex_file, output_dir)
    
    if result_ok:
        print(f"   ✓ Created: {result.name}")
        clean_latex_auxiliary_files(tex_file, output_dir)
        return True
    else:
        print(f"   ✗ Failed: {result}")
        return False


def convert_single_tex_to_docx(tex_file: Path, output_dir: Path) -> bool:
    """Convert a single .tex file to DOCX (via PDF, PDF not kept)."""
    # Create temp dir for intermediate PDF
    temp_dir = output_dir / "_temp_pdf"
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"\n📄 {tex_file.name}")
    
    # Step 1: Compile to PDF
    print("   Compiling LaTeX → PDF...")
    pdf_ok, pdf_result = compile_latex_to_pdf(tex_file, temp_dir)
    
    if not pdf_ok:
        print(f"   ✗ LaTeX compilation failed: {pdf_result}")
        try:
            shutil.rmtree(temp_dir)
        except:
            pass
        return False
    
    # Step 2: Convert PDF to DOCX
    print("   Converting PDF → DOCX...")
    docx_ok, docx_result = convert_pdf_to_docx_enhanced(pdf_result, output_dir)
    
    # Clean up
    clean_latex_auxiliary_files(tex_file, temp_dir)
    try:
        shutil.rmtree(temp_dir)
    except:
        pass
    
    if docx_ok:
        print(f"   ✓ Created: {docx_result.name}")
        return True
    else:
        print(f"   ✗ Conversion failed: {docx_result}")
        return False


def convert_single_pdf_to_docx(pdf_file: Path, output_dir: Path) -> bool:
    """Convert a single .pdf file to DOCX."""
    print(f"\n📄 {pdf_file.name}")
    result_ok, result = convert_pdf_to_docx_enhanced(pdf_file, output_dir)
    
    if result_ok:
        print(f"   ✓ Created: {result.name}")
        return True
    else:
        print(f"   ✗ Failed: {result}")
        return False


def convert_single_docx_to_pdf(docx_file: Path, output_dir: Path) -> bool:
    """Convert a single .docx file to PDF."""
    print(f"\n📄 {docx_file.name}")
    result_ok, result = convert_docx_to_pdf(docx_file, output_dir)
    
    if result_ok:
        print(f"   ✓ Created: {result.name}")
        return True
    else:
        print(f"   ✗ Failed: {result}")
        return False


# ============================================================
# BATCH CONVERSIONS (for directory processing)
# ============================================================

def convert_tex_to_pdf(input_dir: Path, output_dir: Path) -> Tuple[int, int]:
    """Convert all .tex files to PDF."""
    tex_files = find_files(input_dir, 'tex')
    success, failed = 0, 0
    
    for tex_file in tex_files:
        print(f"\n📄 {tex_file.name}")
        result_ok, result = compile_latex_to_pdf(tex_file, output_dir)
        
        if result_ok:
            print(f"   ✓ Created: {result.name}")
            success += 1
            clean_latex_auxiliary_files(tex_file, output_dir)
        else:
            print(f"   ✗ Failed: {result}")
            failed += 1
    
    return success, failed


def convert_tex_to_docx(input_dir: Path, output_dir: Path) -> Tuple[int, int]:
    """Convert all .tex files to DOCX (via PDF, PDF not kept)."""
    tex_files = find_files(input_dir, 'tex')
    success, failed = 0, 0
    
    # Create temp dir for intermediate PDFs
    temp_dir = output_dir / "_temp_pdf"
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    for tex_file in tex_files:
        print(f"\n📄 {tex_file.name}")
        
        # Step 1: Compile to PDF
        print("   Compiling LaTeX → PDF...")
        pdf_ok, pdf_result = compile_latex_to_pdf(tex_file, temp_dir)
        
        if not pdf_ok:
            print(f"   ✗ LaTeX compilation failed: {pdf_result}")
            failed += 1
            continue
        
        # Step 2: Convert PDF to DOCX
        print("   Converting PDF → DOCX...")
        docx_ok, docx_result = convert_pdf_to_docx_enhanced(pdf_result, output_dir)
        
        if docx_ok:
            print(f"   ✓ Created: {docx_result.name}")
            success += 1
        else:
            print(f"   ✗ Conversion failed: {docx_result}")
            failed += 1
        
        # Clean up
        clean_latex_auxiliary_files(tex_file, temp_dir)
    
    # Remove temp directory with intermediate PDFs
    try:
        shutil.rmtree(temp_dir)
    except:
        pass
    
    return success, failed


def convert_pdf_to_docx(input_dir: Path, output_dir: Path) -> Tuple[int, int]:
    """Convert all .pdf files to DOCX."""
    pdf_files = find_files(input_dir, 'pdf')
    success, failed = 0, 0
    
    for pdf_file in pdf_files:
        print(f"\n📄 {pdf_file.name}")
        result_ok, result = convert_pdf_to_docx_enhanced(pdf_file, output_dir)
        
        if result_ok:
            print(f"   ✓ Created: {result.name}")
            success += 1
        else:
            print(f"   ✗ Failed: {result}")
            failed += 1
    
    return success, failed


def convert_docx_to_pdf_batch(input_dir: Path, output_dir: Path) -> Tuple[int, int]:
    """Convert all .docx files to PDF."""
    docx_files = find_files(input_dir, 'docx')
    success, failed = 0, 0
    
    for docx_file in docx_files:
        print(f"\n📄 {docx_file.name}")
        result_ok, result = convert_docx_to_pdf(docx_file, output_dir)
        
        if result_ok:
            print(f"   ✓ Created: {result.name}")
            success += 1
        else:
            print(f"   ✗ Failed: {result}")
            failed += 1
    
    return success, failed
