# PDF to Word Convertor

This VS Code-ready desktop application converts PDF files into visually faithful Microsoft Word `.docx` documents.

It is designed for PDFs produced by TeX/LaTeX and other layout-sensitive sources. Each PDF page is rendered at high resolution and placed on a matching, marginless Word page. That preserves the source fonts, LaTeX rendering, diagrams, spacing, page size, and ordering far more reliably than a text reflow conversion.

The result is visually faithful, but the text and equations are preserved as page images rather than individually editable Word objects.

## Requirements

- Python 3.10 or newer
- Tkinter, which is included with most desktop Python installations
- PyMuPDF
- python-docx

The app runs on Windows, macOS, and Linux. Microsoft Word is not required to perform the conversion.

## Installation

From this repository directory:

```bash
python -m venv .venv
```

Activate the environment.

On Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

On macOS or Linux:

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Run the desktop app

```bash
python app.py
```

The app supports:

- single-file conversion with a selectable output path;
- sequential batch conversion, saving each `.docx` beside its source PDF;
- configurable rendering resolution from 72 to 600 DPI;
- overwrite confirmation before an existing Word file is replaced;
- background conversion so the interface stays responsive;
- cancellation between PDF pages and between batch items.

## Command-line conversion

The conversion engine can also be used without the GUI:

```bash
python converter.py "/path/to/input.pdf" --output "/path/to/output.docx"
```

The default is 300 DPI:

```bash
python converter.py "/path/to/input.pdf" --dpi 300
```

If `--output` is omitted, the Word file is written beside the input PDF with the same filename stem.

## Example

```bash
python converter.py "/Users/zhy/Downloads/MSE 401 HW2.pdf" \
  --output "/Users/zhy/Documents/Codex/2026-09-07/h-e-l/outputs/MSE 401 HW2.docx"
```

## Project layout

```text
PDF-to-Word-Convertor/
├── .vscode/
├── app.py
├── converter.py
├── PDF-to-Word-Convertor.code-workspace
├── README.md
└── requirements.txt
```

