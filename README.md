# PDF to Word Converter

A small Windows desktop application that converts PDF files to editable Microsoft Word `.docx` documents through Microsoft Word's own PDF conversion engine.

## Requirements

- Windows 10 or Windows 11
- Python 3.10 or newer
- The desktop version of Microsoft Word
- `pywin32`

Microsoft Word must be installed and activated on the same Windows computer as the application. The app does not use `pdf2docx` or reconstruct the PDF itself.

## Installation

```text
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Running

```text
python app.py
```

Select a PDF, adjust the output filename if needed, and click **Convert to Word**. The **Add PDFs** and **Convert All** controls support sequential batch conversion while reusing one Microsoft Word instance.

The default output keeps the source filename and changes only the extension. Spaces, capitalization, parentheses, and Unicode characters are preserved.

## How conversion works

The application opens the PDF in Microsoft Word using COM automation, lets Word perform its native PDF reflow/import, and saves the result as a DOCX file. Conversion quality should therefore be approximately equivalent to manually opening the PDF in Word and saving it as DOCX.

The interface remains responsive because Word automation runs in a background thread. COM is initialized and uninitialized inside that worker, and Word is closed even when conversion fails.

## Limitations

- Complex PDFs may still change layout because Word's PDF import is a reflow conversion.
- Scanned PDFs may require OCR before their text becomes editable.
- Password-protected PDFs may fail to open.
- Some equations, fonts, and complicated diagrams may not remain fully editable.
- This application requires Windows and the desktop version of Microsoft Word.
- A cancellation request safely stops after the current Word operation; it does not forcibly terminate Word while it is saving.

Technical exception details are written to `pdf_to_word_converter.log` beside the application when the directory is writable.

## Packaging with PyInstaller

After testing the source on Windows, a single-file GUI executable can be built with:

```text
pyinstaller --noconsole --onefile --name "PDF to Word Converter" app.py
```

`pywin32` is imported lazily by `converter.py`, which keeps the source compatible with PyInstaller and allows the GUI to display a helpful message on unsupported operating systems. If a specific PyInstaller environment needs additional pywin32 collection hooks, install the latest PyInstaller and rebuild in that Windows environment.

