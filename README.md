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

### Interactive Mode

```bash
python main.py
```

### Direct Mode

```bash
python main.py --mode tex2pdf --input ./input --output ./output
python main.py --mode tex2docx --input ./input --output ./output
python main.py --mode pdf2docx --input ./input --output ./output
python main.py --mode docx2pdf --input ./input --output ./output
```

### Options

| Option | Description |
|--------|-------------|
| `-m, --mode` | Conversion mode: `tex2pdf`, `tex2docx`, `pdf2docx`, `docx2pdf` |
| `-i, --input` | Input directory (default: current directory) |
| `-o, --output` | Output directory (default: `./converted`) |

## Project Structure

```
dochameleon/
├── main.py                 # Entry point
├── requirements.txt
├── dochameleon/
│   ├── __init__.py
│   ├── cli.py              # Command-line interface
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
