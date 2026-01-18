"""
Command-line interface for Dochameleon.
"""

import argparse
import os
from pathlib import Path

from .packages import check_and_install_packages, check_latex_installed
from .pipeline import (
    convert_single_tex_to_pdf,
    convert_single_tex_to_docx,
    convert_single_pdf_to_docx,
    convert_single_docx_to_pdf,
)

# Default output directory (relative to script location)
DEFAULT_OUTPUT_DIR = Path(__file__).parent.parent / "output"


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


def get_input_file(expected_ext: str) -> Path:
    """Prompt user for input file path and validate it."""
    while True:
        file_path = input(f"\nEnter path to .{expected_ext} file: ").strip()
        
        # Remove quotes if present
        file_path = file_path.strip('"').strip("'")
        
        if not file_path:
            print(f"  ✗ Please enter a file path.")
            continue
        
        path = Path(file_path)
        
        if not path.exists():
            print(f"  ✗ File not found: {path}")
            continue
        
        if not path.is_file():
            print(f"  ✗ Not a file: {path}")
            continue
        
        if path.suffix.lower() != f".{expected_ext}":
            print(f"  ✗ Expected .{expected_ext} file, got {path.suffix}")
            continue
        
        return path.resolve()


def get_output_dir() -> Path:
    """Prompt user for output directory, with fallback to default."""
    output_path = input(f"\nEnter output folder (press Enter for default): ").strip()
    
    # Remove quotes if present
    output_path = output_path.strip('"').strip("'")
    
    if not output_path:
        output_dir = DEFAULT_OUTPUT_DIR
        print(f"  Using default: {output_dir}")
    else:
        output_dir = Path(output_path).resolve()
    
    # Create directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)
    
    return output_dir


def run_single_conversion(mode: str, input_file: Path, output_dir: Path, packages: dict):
    """Run conversion for a single file."""
    
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
    
    # Show what will be converted
    target_ext = 'pdf' if mode in ['tex2pdf', 'docx2pdf'] else 'docx'
    print(f"\nConverting: {input_file.name} → .{target_ext}")
    print(f"Output directory: {output_dir}")
    print("-" * 50)
    
    # Run conversion
    if mode == 'tex2pdf':
        success = convert_single_tex_to_pdf(input_file, output_dir)
    elif mode == 'tex2docx':
        success = convert_single_tex_to_docx(input_file, output_dir)
    elif mode == 'pdf2docx':
        success = convert_single_pdf_to_docx(input_file, output_dir)
    elif mode == 'docx2pdf':
        success = convert_single_docx_to_pdf(input_file, output_dir)
    else:
        print(f"Unknown mode: {mode}")
        return
    
    # Summary
    print()
    print("-" * 50)
    if success:
        print(f"\n✓ Conversion completed successfully")
        print(f"Output: {output_dir}")
    else:
        print(f"\n✗ Conversion failed")


def interactive_mode(packages: dict):
    """Run in interactive mode with menu."""
    print_menu()
    
    choice = get_user_choice()
    
    if choice == '0':
        print("\nGoodbye!")
        return
    
    # Map choices to modes and expected extensions
    mode_map = {
        '1': ('tex2pdf', 'tex'),
        '2': ('tex2docx', 'tex'),
        '3': ('pdf2docx', 'pdf'),
        '4': ('docx2pdf', 'docx')
    }
    
    mode, expected_ext = mode_map[choice]
    
    # Get input file from user
    input_file = get_input_file(expected_ext)
    
    # Get output directory from user
    output_dir = get_output_dir()
    
    # Run conversion
    run_single_conversion(mode, input_file, output_dir, packages)


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
        help="Input file path"
    )
    parser.add_argument(
        "--output", "-o",
        help="Output directory (default: ./output)"
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
    
    if args.mode and args.input:
        # Direct mode with file
        input_file = Path(args.input).resolve()
        
        if not input_file.exists():
            print(f"✗ Error: File not found: {input_file}")
            return
        
        if args.output:
            output_dir = Path(args.output).resolve()
        else:
            output_dir = DEFAULT_OUTPUT_DIR
        
        output_dir.mkdir(parents=True, exist_ok=True)
        
        run_single_conversion(args.mode, input_file, output_dir, packages)
    else:
        # Interactive mode
        interactive_mode(packages)


if __name__ == "__main__":
    main()
