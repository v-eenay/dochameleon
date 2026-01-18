#!/usr/bin/env python3
"""
Dochameleon - Universal Document Converter
Entry point for the application.
"""

import sys


def main():
    """Main entry point - launches GUI by default, CLI with --cli flag."""
    if "--cli" in sys.argv:
        sys.argv.remove("--cli")
        from dochameleon.cli import main as cli_main
        cli_main()
    else:
        from dochameleon.gui import run_gui
        run_gui()


if __name__ == "__main__":
    main()
