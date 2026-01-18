# Dochameleon

Universal document converter for LaTeX, PDF, and Word formats.

## Features

- TEX → PDF
- TEX → DOCX
- PDF → DOCX
- DOCX → PDF

## Requirements

- Python 3.7+
- LaTeX distribution (MiKTeX or TeX Live) for .tex files
- Microsoft Word (Windows only, for DOCX → PDF)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### GUI Mode (Default)

```bash
python main.py
```

Features:
- Drag and drop file or click to browse
- Select output format
- Choose output folder (defaults to `./output`)

### CLI Mode

```bash
python main.py --cli
```

The CLI will:
1. Display available conversion options
2. Ask for the input file path
3. Ask for the output folder (press Enter for default `./output`)

### Direct CLI Mode

```bash
python main.py --cli --mode tex2pdf --input ./input/document.tex --output ./output
python main.py --cli --mode tex2docx --input ./input/document.tex
python main.py --cli --mode pdf2docx --input ./docs/file.pdf --output ./converted
python main.py --cli --mode docx2pdf --input ./docs/report.docx
```

### Options

| Option | Description |
|--------|-------------|
| `--cli` | Run in command-line mode instead of GUI |
| `-m, --mode` | Conversion mode: `tex2pdf`, `tex2docx`, `pdf2docx`, `docx2pdf` |
| `-i, --input` | Input file path |
| `-o, --output` | Output directory (default: `./output`) |

## Project Structure

```
dochameleon/
├── main.py                 # Entry point (GUI default, --cli for CLI)
├── requirements.txt
├── dochameleon/
│   ├── __init__.py
│   ├── cli.py              # Command-line interface
│   ├── gui.py              # Qt GUI interface
│   ├── packages.py         # Package management
│   ├── pipeline.py         # Conversion pipelines
│   ├── utils.py            # Utility functions
│   └── converters/
│       ├── __init__.py
│       ├── latex.py        # LaTeX to PDF
│       ├── pdf.py          # PDF to DOCX
│       └── docx.py         # DOCX to PDF
├── input/                  # Sample input files
└── icons/
```

## Author

Vinay Koirala
