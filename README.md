# PDF to Word Converter

A Windows and macOS desktop application that converts PDF files to Word `.docx` documents. Exact page-preserving conversion is enabled by default; editable Microsoft Word conversion is available as an opt-in mode.

## Requirements

- Windows 10 or Windows 11, or a macOS version supported by your Word installation
- Python 3.10 or newer
- The desktop version of Microsoft Word for Windows or Mac
- `pywin32` on Windows only

The app uses Tkinter, which is included with standard Python distributions. On macOS, use a Python distribution that includes Tcl/Tk support. `PyMuPDF` and `python-docx` are used by the optional exact-appearance mode.

Microsoft Word must be installed and activated on the same computer as the application. The app does not use `pdf2docx` or reconstruct the PDF itself.

## Installation

```text
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

On macOS, activate the environment with:

```text
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

## Running

```text
python app.py
```

Select a PDF, adjust the output filename if needed, and click **Convert to Word**. The **Add PDFs** and **Convert All** controls support sequential batch conversion while reusing one Microsoft Word application instance.

**Preserve exact PDF appearance** is enabled by default. Keep it enabled for equation-heavy or font-sensitive PDFs. This creates a DOCX with one high-resolution PDF page image per Word page, preserving the original equations, embedded fonts, diagrams, and page geometry. The page contents are visually faithful but are not individually editable. Uncheck it only when you need Word's editable text conversion.

For example, the supplied TeX-generated homework uses embedded Latin Modern text and math fonts. Word's editable PDF importer may substitute fonts and flatten equations into ordinary text or shapes; the exact-appearance mode retains the original Latin Modern rendering.

The default output keeps the source filename and changes only the extension. Spaces, capitalization, parentheses, and Unicode characters are preserved.

## Platform backends

- Windows uses `win32com.client.DispatchEx("Word.Application")`, with COM initialized inside the worker thread.
- macOS uses the built-in `/usr/bin/osascript` command to control Microsoft Word for Mac through its AppleScript dictionary. No additional Python package is required.

The first macOS conversion may show an automation permission prompt. If conversion is blocked, allow the application or Python to control Microsoft Word in **System Settings > Privacy & Security > Automation**. The app closes each source document after saving. It does not quit a Word session that was already open; if it started an idle Word session itself, it only quits it when no documents remain.

Some Word for Mac releases have an AppleScript compatibility problem where Word opens a PDF but rejects the scripted **Save As** command. The app identifies this as a Word/AppleScript issue instead of incorrectly reporting a folder-permission problem. Update Microsoft Word and retry; if the issue remains, use **File > Save As > Word Document (.docx)** directly in Word.

## How conversion works

When exact appearance is disabled, the application opens the PDF in Microsoft Word using the platform's native automation interface, lets Word perform its native PDF reflow/import, and saves the result as a DOCX file. Conversion quality should therefore be approximately equivalent to manually opening the PDF in Word and saving it as DOCX.

The exact-appearance mode uses PyMuPDF only to render each PDF page and `python-docx` only to place those rendered pages into a DOCX. It does not attempt to reconstruct PDF text or equations. A PDF does not generally contain the original LaTeX source, so truly editable Word equations require the original `.tex` source or a separate math-OCR workflow.

The interface remains responsive because Word automation runs in a background thread. COM is initialized and uninitialized inside that worker on Windows, and source documents are closed even when conversion fails.

## Limitations

- Complex PDFs may still change layout because Word's PDF import is a reflow conversion. This is the best editable-DOCX fidelity Word provides, but it is not guaranteed to be pixel-identical to the source PDF.
- Scanned PDFs may require OCR before their text becomes editable.
- Password-protected PDFs may fail to open.
- Some equations, fonts, and complicated diagrams may not remain fully editable.
- If exact visual identity is more important than editability, a page-as-image DOCX workflow is required; that is a separate mode from Word's native editable conversion.
- The editable Word mode requires Windows or macOS and the desktop version of Microsoft Word; exact-appearance mode does not use Word's PDF importer.
- A cancellation request safely stops after the current Word operation; it does not forcibly terminate Word while it is saving.

Technical exception details are written to `pdf_to_word_converter.log` beside the application when the directory is writable.

## Packaging with PyInstaller

After testing the source on the target platform, a single-file GUI executable can be built with:

```text
pyinstaller --noconsole --onefile --name "PDF to Word Converter" app.py
```

`pywin32` is imported lazily by `converter.py`, which keeps the source compatible with PyInstaller and allows the GUI to run on macOS without the Windows COM package. If a specific Windows PyInstaller environment needs additional pywin32 collection hooks, install the latest PyInstaller and rebuild in that Windows environment.
