# PDF to Word Converter

Convert PDFs to Word documents while preserving LaTeX, equations, fonts, diagrams, spacing, and page order. **Preserve exact PDF appearance** is enabled by default.

## Requirements

- Python 3.10 or newer
- Tkinter
- Microsoft Word only if editable Word conversion is needed

## macOS setup

Open Terminal or the VS Code terminal:

```bash
cd ~/PDF-to-Word-Convertor
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python app.py
```

## Windows setup

Open PowerShell or the VS Code terminal:

```powershell
cd "$env:USERPROFILE\PDF-to-Word-Convertor"
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python app.py
```

If PowerShell blocks activation, run this once in the same terminal:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Using the app

1. Select the PDF file.
2. Choose the output `.docx` location.
3. Keep **Preserve exact PDF appearance** checked for the best result with LaTeX, equations, fonts, and diagrams.
4. Click **Convert to Word**.

The exact-appearance mode creates a Word page image for each PDF page. The result is visually faithful, but the text and equations are not individually editable.

Uncheck **Preserve exact PDF appearance** only when you need editable Word text. This mode uses Microsoft Word's native PDF importer and requires the desktop version of Word.

## macOS Word permission

If editable conversion is used on macOS, allow Python or Terminal to control Microsoft Word under:

**System Settings → Privacy & Security → Automation**

If Word reports that it rejected the macOS **Save As** command, keep **Preserve exact PDF appearance** checked. That mode avoids Word automation and is recommended for equation-heavy PDFs.
