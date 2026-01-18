"""
Command-line interface for Dochameleon.
"""

import argparse
from pathlib import Path

from .packages import check_and_install_packages, check_latex_installed
from .utils import find_files
from .pipeline import (
    convert_tex_to_pdf,
    convert_tex_to_docx,
    convert_pdf_to_docx,
    convert_docx_to_pdf_batch,
)


def print_header():
    """Print program header."""
    print()
    print("╔" + "═" * 58 + "╗")
    print("║" + " Dochameleon - Universal Document Converter ".center(58) + "║")
    print("║" + " LaTeX ↔ PDF ↔ DOCX ".center(58) + "║")
    print("╚" + "═" * 58 + "╝")
    print()


def print_menu():
    """Print conversion options menu."""
    print("Available conversions:")
    print()
    print("  [1] TEX  → PDF   (LaTeX to PDF)")
    print("  [2] TEX  → DOCX  (LaTeX to Word, preserves formatting)")
    print("  [3] PDF  → DOCX  (PDF to Word)")
    print("  [4] DOCX → PDF   (Word to PDF, requires MS Word)")
    print()
    print("  [0] Exit")
    print()


def get_user_choice() -> str:
    """Get user's conversion choice."""
    while True:
        choice = input("Enter your choice (0-4): ").strip()
        if choice in ['0', '1', '2', '3', '4']:
            return choice
        print("Invalid choice. Please enter 0-4.")


def run_conversion(mode: str, input_dir: Path, output_dir: Path, packages: dict):
    """Run the specified conversion mode."""
    
    # Check requirements
    if mode in ['tex2pdf', 'tex2docx']:
        if not check_latex_installed():
            print("\n✗ Error: LaTeX (pdflatex) is not installed.")
            print("  Install MiKTeX or TeX Live to compile LaTeX files.")
            return
    
    if mode in ['tex2docx', 'pdf2docx']:
        if not packages['pdf2docx']:
            print("\n✗ Error: pdf2docx is required but not available.")
            return
    
    if mode == 'docx2pdf':
        if not packages['docx2pdf']:
            print("\n✗ Error: docx2pdf is required but not available.")
            return
    
    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Find source files
    ext_map = {
        'tex2pdf': 'tex',
        'tex2docx': 'tex',
        'pdf2docx': 'pdf',
        'docx2pdf': 'docx'
    }
    source_ext = ext_map[mode]
    source_files = find_files(input_dir, source_ext)
    
    if not source_files:
        print(f"\n✗ No .{source_ext} files found in {input_dir}")
        return
    
    # Show what will be converted
    target_ext = 'pdf' if mode in ['tex2pdf', 'docx2pdf'] else 'docx'
    print(f"\nConverting {len(source_files)} .{source_ext} file(s) to .{target_ext}")
    print(f"Output directory: {output_dir}")
    print("-" * 50)
    
    # Run conversion
    if mode == 'tex2pdf':
        success, failed = convert_tex_to_pdf(input_dir, output_dir)
    elif mode == 'tex2docx':
        success, failed = convert_tex_to_docx(input_dir, output_dir)
    elif mode == 'pdf2docx':
        success, failed = convert_pdf_to_docx(input_dir, output_dir)
    elif mode == 'docx2pdf':
        success, failed = convert_docx_to_pdf_batch(input_dir, output_dir)
    else:
        print(f"Unknown mode: {mode}")
        return
    
    # Summary
    print()
    print("-" * 50)
    print(f"\n✓ Completed: {success} successful, {failed} failed")
    print(f"Output: {output_dir}")


def interactive_mode(input_dir: Path, output_dir: Path, packages: dict):
    """Run in interactive mode with menu."""
    print_menu()
    
    choice = get_user_choice()
    
    if choice == '0':
        print("\nGoodbye!")
        return
    
    # Map choices to modes
    mode_map = {
        '1': 'tex2pdf',
        '2': 'tex2docx',
        '3': 'pdf2docx',
        '4': 'docx2pdf'
    }
    
    mode = mode_map[choice]
    run_conversion(mode, input_dir, output_dir, packages)


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Dochameleon - Universal Document Converter: LaTeX ↔ PDF ↔ DOCX"
    )
    parser.add_argument(
        "--mode", "-m",
        choices=['tex2pdf', 'tex2docx', 'pdf2docx', 'docx2pdf'],
        help="Conversion mode (if not specified, runs interactive menu)"
    )
    parser.add_argument(
        "--input", "-i",
        default=".",
        help="Input directory (default: current directory)"
    )
    parser.add_argument(
        "--output", "-o",
        default="./converted",
        help="Output directory (default: ./converted)"
    )
    
    args = parser.parse_args()
    
    print_header()
    
    # Check and install packages
    print("Checking requirements...")
    packages = check_and_install_packages()
    
    latex_available = check_latex_installed()
    if latex_available:
        print("  ✓ LaTeX (pdflatex) is available")
    else:
        print("  ⚠ LaTeX not found (needed for .tex files)")
    
    print()
    
    input_dir = Path(args.input).resolve()
    output_dir = Path(args.output).resolve()
    
    print(f"Input directory:  {input_dir}")
    print(f"Output directory: {output_dir}")
    print()
    
    if args.mode:
        # Direct mode
        run_conversion(args.mode, input_dir, output_dir, packages)
    else:
        # Interactive mode
        interactive_mode(input_dir, output_dir, packages)


if __name__ == "__main__":
    main()
