# PDF to Word Converter

## Requirements

- Python 3.10 or newer
- Tkinter
- Microsoft Word for the native editable mode on Windows; macOS can use the editable fallback
- A LaTeX installation with `pdflatex`, `xelatex`, or `lualatex` for the LaTeX Workspace

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

Uncheck **Preserve exact PDF appearance** only when you need editable Word text. Windows uses Microsoft Word's native PDF importer. On macOS, the app tries Word first and automatically falls back to `pdf2docx` if Word is unavailable or rejects PDF opening, automation, or Save As.

## LaTeX Workspace

The workspace is the editable source-first workflow. It lets you edit the `.tex` source beside a live PDF preview, compile without changing the source file, and export the compiled PDF to Word.

Start it from the converter with **Open LaTeX Workspace**, or run:

```bash
python latex_workspace.py
```

Then:

1. Open the original `.tex` file.
2. Keep **Auto compile** enabled, or click **Compile / Preview** after editing.
3. If the source is a body fragment without a preamble, the workspace supplies an `exam` wrapper automatically.
4. If figures referenced by `\includegraphics` are missing, the workspace looks for a same-name PDF and recovers the embedded figures into its temporary build directory. The original `.tex` and PDF are not modified.
5. Use **Export Exact Word** to preserve every rendered page exactly, or **Export Editable Word** to create a reflowable Word document.

The `.tex` file remains the source of truth, so equations stay as real LaTeX and can be edited without rewriting them. Exact Word export is visually faithful but page content is not individually editable; editable Word export can reflow equations and layout.

## macOS Word permission

If editable conversion uses Word on macOS, allow Python or Terminal to control Microsoft Word under:

**System Settings → Privacy & Security → Automation**

If Word reports that it rejected the macOS PDF-open or **Save As** command, the app automatically tries the editable fallback. For the most faithful equations and fonts, keep **Preserve exact PDF appearance** checked; editable conversion cannot recreate the original LaTeX source exactly.
