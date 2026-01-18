"""
Package management utilities for Dochameleon.
"""

import subprocess
import sys


def install_package(package: str) -> bool:
    """Install a Python package using pip."""
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", package, "-q"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return True
    except:
        return False


def check_and_install_packages() -> dict:
    """Check and install required packages, return availability status."""
    packages = {
        'pdf2docx': False,
        'docx': False,
        'docx2pdf': False,
    }
    
    # Check pdf2docx
    try:
        from pdf2docx import Converter
        packages['pdf2docx'] = True
    except ImportError:
        print("  Installing pdf2docx...")
        if install_package("pdf2docx"):
            packages['pdf2docx'] = True
            print("  ✓ pdf2docx installed")
        else:
            print("  ✗ Failed to install pdf2docx")
    
    # Check python-docx
    try:
        from docx import Document
        packages['docx'] = True
    except ImportError:
        print("  Installing python-docx...")
        if install_package("python-docx"):
            packages['docx'] = True
            print("  ✓ python-docx installed")
        else:
            print("  ✗ Failed to install python-docx")
    
    # Check docx2pdf
    try:
        import docx2pdf
        packages['docx2pdf'] = True
    except ImportError:
        print("  Installing docx2pdf...")
        if install_package("docx2pdf"):
            packages['docx2pdf'] = True
            print("  ✓ docx2pdf installed")
        else:
            print("  ✗ Failed to install docx2pdf")
    
    return packages


def check_latex_installed() -> bool:
    """Check if LaTeX (pdflatex) is installed."""
    try:
        result = subprocess.run(
            ["pdflatex", "--version"],
            capture_output=True,
            text=True,
            check=True
        )
        return True
    except:
        return False
