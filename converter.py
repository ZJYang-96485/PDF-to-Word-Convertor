"""Microsoft Word based PDF-to-DOCX conversion.

The conversion is intentionally delegated to Microsoft Word.  Word's PDF
importer is responsible for reflowing the document, which generally produces
the same result as opening the PDF manually in Word and saving it as DOCX.

This module does not import pywin32 at module import time. That keeps the GUI
able to start on macOS and other systems without failing with an ImportError.
"""

from __future__ import annotations

import gc
import logging
import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable


LOGGER = logging.getLogger(__name__)
WD_FORMAT_XML_DOCUMENT = 16
StatusCallback = Callable[[str], None]


class PDFConversionError(Exception):
    """Base exception for expected conversion failures.

    ``user_message`` is intentionally separate from the technical exception
    text so the GUI can show a friendly explanation while logging the original
    details for troubleshooting.
    """

    def __init__(self, message: str, user_message: str | None = None) -> None:
        super().__init__(message)
        self.user_message = user_message or message


class PDFValidationError(PDFConversionError):
    """The input or output path is invalid."""


class WordNotInstalledError(PDFConversionError):
    """Microsoft Word or the pywin32 COM bridge could not be started."""


class PDFOpenError(PDFConversionError):
    """Word could not import the PDF."""


class OutputExistsError(PDFConversionError):
    """The requested output already exists and overwrite was not approved."""


class OutputError(PDFConversionError):
    """The converted DOCX could not be written."""


class ConversionCancelledError(PDFConversionError):
    """Conversion was cancelled before Word began opening the PDF."""


class ConversionDependencyError(PDFConversionError):
    """An optional dependency required by a selected conversion mode is missing."""


def _absolute_path(path_value: str | os.PathLike[str], description: str) -> Path:
    """Return an absolute path and convert path parsing errors to useful errors."""

    try:
        return Path(os.path.abspath(os.fspath(path_value)))
    except (TypeError, ValueError, OSError) as exc:
        raise PDFValidationError(
            f"Invalid {description} path: {path_value!r}",
            f"The {description} filename is not valid.",
        ) from exc


def normalise_pdf_path(pdf_path: str | os.PathLike[str]) -> Path:
    """Validate and normalize a PDF input path."""

    path = _absolute_path(pdf_path, "PDF")
    if not path.exists() or not path.is_file():
        raise PDFValidationError(
            f"PDF does not exist or is not a file: {path}",
            "The selected PDF file could not be found.",
        )
    if path.suffix.lower() != ".pdf":
        raise PDFValidationError(
            f"Input is not a PDF: {path}",
            "Please select a file with a .pdf extension.",
        )
    return path


def normalise_output_path(
    pdf_path: str | os.PathLike[str], output_path: str | os.PathLike[str] | None = None
) -> Path:
    """Return an absolute DOCX path, deriving it from the PDF when omitted."""

    input_path = _absolute_path(pdf_path, "PDF")
    if output_path is None or not str(output_path).strip():
        return input_path.with_suffix(".docx")

    path = _absolute_path(output_path, "output")
    if path.suffix == "":
        path = path.with_suffix(".docx")
    if path.suffix.lower() != ".docx":
        raise PDFValidationError(
            f"Output is not a DOCX path: {path}",
            "The output filename must have a .docx extension.",
        )
    return path


def _prepare_conversion_paths(
    pdf_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str] | None,
    overwrite: bool,
) -> tuple[Path, Path]:
    """Validate both paths and the output destination before starting Word."""

    pdf = normalise_pdf_path(pdf_path)
    docx = normalise_output_path(pdf, output_path)

    if docx.exists() and not overwrite:
        raise OutputExistsError(
            f"Output already exists: {docx}",
            f"{docx.name} already exists. Choose a different output filename or approve overwrite.",
        )
    if not docx.parent.exists():
        raise OutputError(
            f"Output folder does not exist: {docx.parent}",
            "The Word file could not be saved because the output folder does not exist.",
        )
    if docx.exists() and not os.access(docx, os.W_OK):
        raise OutputError(
            f"Output is not writable: {docx}",
            "The Word file could not be saved because the existing file is read-only.",
        )
    if not os.access(docx.parent, os.W_OK):
        raise OutputError(
            f"Output folder is not writable: {docx.parent}",
            "The Word file could not be saved. Check your permission to write to this folder.",
        )
    return pdf, docx


class WordPDFConverter:
    """Own one isolated Word COM instance and reuse it for multiple files.

    Instances must be started, used, and closed on the same thread.  The GUI
    creates this object inside its worker thread so COM initialization and
    cleanup happen in the correct apartment.
    """

    def __init__(self, visible: bool = False) -> None:
        self.visible = visible
        self._word = None
        self._pythoncom = None
        self._com_initialized = False

    def __enter__(self) -> "WordPDFConverter":
        self.start()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()

    def start(self) -> None:
        """Initialize COM and start a private Word instance."""

        if self._word is not None:
            return

        if platform.system() != "Windows":
            raise WordNotInstalledError(
                "Microsoft Word COM automation is only available on Windows.",
                "This application requires Windows and the desktop version of Microsoft Word.",
            )

        try:
            import pythoncom  # type: ignore[import-not-found]
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise WordNotInstalledError(
                "pywin32 is not installed; Python could not import the Word COM bridge.",
                "Microsoft Word could not be started. Please install the application requirements first.",
            ) from exc

        self._pythoncom = pythoncom
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = self.visible
            self._word.DisplayAlerts = 0
        except Exception as exc:
            LOGGER.exception("Could not initialize Microsoft Word COM automation.")
            self.close()
            raise WordNotInstalledError(
                f"Could not start Microsoft Word: {exc}",
                "Microsoft Word could not be started. Please make sure the desktop version of Microsoft Word is installed.",
            ) from exc

    def convert_pdf(
        self,
        pdf_path: str | os.PathLike[str],
        output_path: str | os.PathLike[str] | None = None,
        *,
        overwrite: bool = False,
        status_callback: StatusCallback | None = None,
    ) -> str:
        """Convert one PDF and return the absolute output DOCX path.

        ``overwrite`` must be explicitly enabled by the caller.  The GUI asks
        for confirmation before starting its worker thread and passes that
        decision here as a second safety check.
        """

        pdf, docx = _prepare_conversion_paths(pdf_path, output_path, overwrite)

        self.start()
        doc = None
        try:
            self._notify(status_callback, "Opening PDF in Microsoft Word...")
            try:
                doc = self._word.Documents.Open(
                    str(pdf),
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                )
            except Exception as exc:
                LOGGER.exception("Word could not open PDF: %s", pdf)
                raise PDFOpenError(
                    f"Word could not open {pdf}: {exc}",
                    "Microsoft Word could not open this PDF. The PDF may be corrupted, password-protected, or unsupported.",
                ) from exc

            self._notify(status_callback, "Converting PDF...")
            self._notify(status_callback, "Saving Word document...")
            try:
                doc.SaveAs2(str(docx), FileFormat=WD_FORMAT_XML_DOCUMENT)
            except Exception as exc:
                LOGGER.exception("Word could not save DOCX: %s", docx)
                raise OutputError(
                    f"Word could not save {docx}: {exc}",
                    "The Word file could not be saved. Make sure the output file is not already open and that you have permission to write to this folder.",
                ) from exc

            return str(docx)
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    LOGGER.exception("Could not close Word document cleanly: %s", pdf)
                finally:
                    doc = None

    @staticmethod
    def _notify(callback: StatusCallback | None, message: str) -> None:
        if callback is not None:
            try:
                callback(message)
            except Exception:
                LOGGER.exception("Status callback failed.")

    def close(self) -> None:
        """Close Word and uninitialize COM, swallowing cleanup-only failures."""

        word = self._word
        self._word = None
        try:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    LOGGER.exception("Could not quit Microsoft Word cleanly.")
        finally:
            word = None
            # Release temporary COM dispatch wrappers before uninitializing
            # the apartment so WINWORD.EXE does not remain referenced.
            gc.collect()
            if self._com_initialized and self._pythoncom is not None:
                try:
                    self._pythoncom.CoUninitialize()
                except Exception:
                    LOGGER.exception("Could not uninitialize COM cleanly.")
                finally:
                    self._com_initialized = False
                    self._pythoncom = None


_MAC_CONVERSION_SCRIPT = r'''on run argv
    set inputPath to item 1 of argv
    set outputPath to item 2 of argv
    set sourceDoc to missing value
    set startedAt to current date

    tell application id "com.microsoft.Word"
        set previousAlerts to display alerts
        set display alerts to alerts none
        try
            try
                open (POSIX file (my inputPath))
            on error errMsg number errNum
                error ("PDF_OPEN|" & errMsg) number errNum
            end try

            try
                repeat until (count of documents) > 0
                    if ((current date) - startedAt) > 120 then
                        error "Timed out waiting for Word to finish opening the PDF."
                    end if
                    delay 0.5
                end repeat
                set sourceDoc to active document
            on error errMsg number errNum
                error ("PDF_OPEN|" & errMsg) number errNum
            end try

            try
                set outputFile to (POSIX file (my outputPath)) as text
                save as active document file name outputFile file format format document
            on error errMsg number errNum
                try
                    close (my sourceDoc) saving no
                end try
                error ("DOCX_SAVE|" & errMsg) number errNum
            end try

            try
                close (my sourceDoc) saving no
            on error errMsg number errNum
                error ("DOC_CLOSE|" & errMsg) number errNum
            end try
            set display alerts to previousAlerts
        on error errMsg number errNum
            try
                if (my sourceDoc) is not missing value then close (my sourceDoc) saving no
            end try
            try
                set display alerts to previousAlerts
            end try
            error errMsg number errNum
        end try
    end tell
end run'''

_MAC_QUIT_IF_IDLE_SCRIPT = r'''tell application id "com.microsoft.Word"
    if (count of documents) is 0 then quit
end tell'''
MAC_AUTOMATION_TIMEOUT_SECONDS = 15 * 60


class MacWordConverter:
    """Convert PDFs through Microsoft Word for Mac using ``osascript``.

    Microsoft Word for Mac exposes an AppleScript dictionary rather than the
    Windows COM API. The script opens and saves one document at a time, so the
    same Word application can be reused for batch conversion.
    """

    def __init__(self, visible: bool = False) -> None:
        # Word's Mac automation requires the app to be visible/active. Keep
        # the argument for API parity with the Windows converter.
        self.visible = visible
        self._osascript_path: str | None = None
        self._started = False
        self._owns_word_app = False

    def __enter__(self) -> "MacWordConverter":
        self.start()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()

    def start(self) -> None:
        """Check for osascript and Microsoft Word without importing pywin32."""

        if self._started:
            return
        if platform.system() != "Darwin":
            raise WordNotInstalledError(
                "Microsoft Word for Mac automation is only available on macOS.",
                "This conversion backend is available on macOS with Microsoft Word installed.",
            )

        self._osascript_path = shutil.which("osascript")
        if not self._osascript_path:
            raise WordNotInstalledError(
                "The macOS osascript command was not found.",
                "macOS scripting support could not be found. Please make sure this is a standard macOS installation.",
            )

        try:
            probe = subprocess.run(
                [self._osascript_path, "-e", 'id of application "Microsoft Word"'],
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError as exc:
            raise WordNotInstalledError(
                f"Could not run osascript: {exc}",
                "Microsoft Word could not be detected. Please install the desktop version of Microsoft Word for Mac.",
            ) from exc

        if probe.returncode != 0:
            details = (probe.stderr or probe.stdout).strip()
            raise WordNotInstalledError(
                f"Microsoft Word could not be detected by osascript: {details}",
                "Microsoft Word could not be found. Please install the desktop version of Microsoft Word for Mac.",
            )

        self._owns_word_app = not self._word_process_is_running()
        self._started = True

    @staticmethod
    def _word_process_is_running() -> bool:
        """Best-effort check used to avoid quitting a user's existing Word app."""

        try:
            result = subprocess.run(
                ["pgrep", "-f", "Microsoft Word.app"],
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError:
            # Be conservative if process inspection is unavailable.
            return True
        return result.returncode == 0

    def convert_pdf(
        self,
        pdf_path: str | os.PathLike[str],
        output_path: str | os.PathLike[str] | None = None,
        *,
        overwrite: bool = False,
        status_callback: StatusCallback | None = None,
    ) -> str:
        """Convert one PDF with Word for Mac and return the absolute DOCX path."""

        pdf, docx = _prepare_conversion_paths(pdf_path, output_path, overwrite)
        self.start()
        self._notify(status_callback, "Opening PDF in Microsoft Word...")
        self._notify(status_callback, "Converting PDF...")
        self._notify(status_callback, "Saving Word document...")

        try:
            result = subprocess.run(
                [
                    self._osascript_path or "osascript",
                    "-e",
                    _MAC_CONVERSION_SCRIPT,
                    str(pdf),
                    str(docx),
                ],
                capture_output=True,
                text=True,
                check=False,
                timeout=MAC_AUTOMATION_TIMEOUT_SECONDS,
            )
        except subprocess.TimeoutExpired as exc:
            LOGGER.exception("Word for Mac automation timed out for %s.", pdf)
            raise PDFOpenError(
                f"Word for Mac automation timed out for {pdf}: {exc}",
                "Microsoft Word took too long to convert this PDF. Complete any Word dialog that is open, then try again with a smaller or simpler PDF.",
            ) from exc
        except OSError as exc:
            LOGGER.exception("Could not run Word for Mac automation.")
            raise PDFOpenError(
                f"Could not run osascript for {pdf}: {exc}",
                "Microsoft Word could not be controlled. Check that Microsoft Word is installed and allow this app to control it in System Settings > Privacy & Security > Automation.",
            ) from exc

        if result.returncode != 0:
            details = (result.stderr or result.stdout).strip()
            LOGGER.error("Word for Mac conversion failed for %s: %s", pdf, details)
            error_text = details.lower()
            if "docx_save|" in error_text:
                if "-1708" in error_text or (
                    "save as" in error_text
                    and ("doesn't understand" in error_text or "不理解" in error_text)
                ):
                    raise OutputError(
                        f"Word for Mac does not expose Save As through AppleScript for {docx}: {details}",
                        "Word opened the PDF, but this installed version of Word for Mac rejected the Save As command used by macOS automation. This is a Word/AppleScript compatibility issue, not necessarily a folder-permission problem. Update Word and try again; if it persists, open the PDF in Word and choose File > Save As > Word Document (.docx).",
                    )
                raise OutputError(
                    f"Word for Mac could not save {docx}: {details}",
                    "The Word file could not be saved. Make sure the output file is not already open and that you have permission to write to this folder.",
                )
            if "not authorized" in error_text or "not permitted" in error_text or "automation" in error_text:
                raise PDFOpenError(
                    f"macOS automation permission was denied: {details}",
                    "macOS blocked control of Microsoft Word. Allow this application to control Microsoft Word in System Settings > Privacy & Security > Automation, then try again.",
                )
            raise PDFOpenError(
                f"Word for Mac could not open or convert {pdf}: {details}",
                "Microsoft Word could not open this PDF. The PDF may be corrupted, password-protected, unsupported, or macOS may need permission to control Word.",
            )

        if not docx.exists():
            raise OutputError(
                f"Word for Mac reported success but the output was not found: {docx}",
                "Microsoft Word reported that conversion finished, but the Word file could not be found in the selected output location.",
            )

        return str(docx)

    @staticmethod
    def _notify(callback: StatusCallback | None, message: str) -> None:
        if callback is not None:
            try:
                callback(message)
            except Exception:
                LOGGER.exception("Status callback failed.")

    def close(self) -> None:
        """Close only an idle Word app started by this converter."""

        if self._started and self._owns_word_app and self._osascript_path:
            try:
                result = subprocess.run(
                    [self._osascript_path, "-e", _MAC_QUIT_IF_IDLE_SCRIPT],
                    capture_output=True,
                    text=True,
                    check=False,
                    timeout=10,
                )
                if result.returncode != 0:
                    LOGGER.warning("Could not close idle Word for Mac: %s", result.stderr.strip())
            except (OSError, subprocess.TimeoutExpired):
                LOGGER.exception("Could not close idle Word for Mac cleanly.")
        self._started = False
        self._osascript_path = None
        self._owns_word_app = False


VISUAL_RENDER_DPI = 300


class VisualFidelityConverter:
    """Create a visually faithful DOCX by placing one rendered PDF page per page.

    This mode intentionally preserves the PDF page as a high-resolution image.
    It is the reliable option for TeX equations, embedded fonts, diagrams, and
    exact page geometry when the result does not need individually editable text.
    """

    def __init__(self, dpi: int = VISUAL_RENDER_DPI) -> None:
        if dpi <= 0:
            raise ValueError("dpi must be positive")
        self.dpi = dpi
        self._started = False

    def __enter__(self) -> "VisualFidelityConverter":
        self.start()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()

    def start(self) -> None:
        """Keep the converter API compatible with the Word backends."""

        self._started = True

    def convert_pdf(
        self,
        pdf_path: str | os.PathLike[str],
        output_path: str | os.PathLike[str] | None = None,
        *,
        overwrite: bool = False,
        status_callback: StatusCallback | None = None,
    ) -> str:
        """Render a PDF into a page-sized image DOCX and return its path."""

        pdf, docx = _prepare_conversion_paths(pdf_path, output_path, overwrite)
        self.start()

        try:
            import pymupdf as fitz  # type: ignore[import-not-found]
        except ImportError as exc:
            try:
                import fitz  # type: ignore[import-not-found]
            except ImportError:
                raise ConversionDependencyError(
                    "PyMuPDF is required for exact-appearance conversion.",
                    "The exact-appearance mode needs the PyMuPDF package. Run `pip install -r requirements.txt`, then try again.",
                ) from exc

        try:
            from docx import Document  # type: ignore[import-not-found]
            from docx.enum.section import WD_SECTION  # type: ignore[import-not-found]
            from docx.enum.text import WD_ALIGN_PARAGRAPH  # type: ignore[import-not-found]
            from docx.shared import Inches, Pt  # type: ignore[import-not-found]
        except ImportError as exc:
            raise ConversionDependencyError(
                "python-docx is required for exact-appearance conversion.",
                "The exact-appearance mode needs the python-docx package. Run `pip install -r requirements.txt`, then try again.",
            ) from exc

        try:
            source = fitz.open(str(pdf))
        except Exception as exc:
            LOGGER.exception("Could not read PDF for visual conversion: %s", pdf)
            raise PDFOpenError(
                f"Could not render PDF {pdf}: {exc}",
                "The PDF could not be rendered. It may be corrupted, password-protected, or unsupported.",
            ) from exc

        with source:
            if source.page_count == 0:
                raise PDFOpenError(
                    f"The PDF has no pages: {pdf}",
                    "The selected PDF does not contain any pages.",
                )

            document = Document()
            with tempfile.TemporaryDirectory(prefix="pdf_to_word_pages_") as image_dir:
                for page_index, page in enumerate(source):
                    page_width = float(page.rect.width)
                    page_height = float(page.rect.height)
                    if page_width <= 0 or page_height <= 0:
                        raise PDFOpenError(
                            f"Invalid page dimensions on page {page_index + 1} of {pdf}",
                            "The PDF contains a page with invalid dimensions and could not be converted.",
                        )

                    self._notify(
                        status_callback,
                        f"Rendering page {page_index + 1} of {source.page_count}...",
                    )
                    scale = self.dpi / 72.0
                    pixmap = page.get_pixmap(
                        matrix=fitz.Matrix(scale, scale),
                        alpha=False,
                    )
                    image_path = Path(image_dir) / f"page-{page_index + 1:04d}.png"
                    pixmap.save(str(image_path))

                    if page_index == 0:
                        section = document.sections[0]
                        paragraph = document.add_paragraph()
                    else:
                        previous_paragraph = document.paragraphs[-1]
                        section = document.add_section(WD_SECTION.NEW_PAGE)
                        section_break = document.paragraphs[-1]
                        paragraph = document.add_paragraph()
                        self._move_section_break(section_break, previous_paragraph)

                    self._configure_page(section, page_width, page_height, Inches)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1
                    paragraph.paragraph_format.keep_together = True
                    run = paragraph.add_run()
                    run.add_picture(
                        str(image_path),
                        width=Inches(page_width / 72.0),
                        height=Inches(page_height / 72.0),
                    )

                self._notify(status_callback, "Building Word document...")
                try:
                    document.save(str(docx))
                except Exception as exc:
                    LOGGER.exception("Could not save visual DOCX: %s", docx)
                    raise OutputError(
                        f"Could not save visual DOCX {docx}: {exc}",
                        "The Word file could not be saved. Check that the output folder is writable and that the file is not open.",
                    ) from exc

        return str(docx)

    @staticmethod
    def _configure_page(section, width_points: float, height_points: float, inches) -> None:
        """Set page geometry to the source PDF page with no added margins."""

        section.page_width = inches(width_points / 72.0)
        section.page_height = inches(height_points / 72.0)
        section.top_margin = inches(0)
        section.bottom_margin = inches(0)
        section.left_margin = inches(0)
        section.right_margin = inches(0)
        section.header_distance = inches(0)
        section.footer_distance = inches(0)
        section.gutter = inches(0)

    @staticmethod
    def _move_section_break(section_break, preceding_paragraph) -> None:
        """Attach a section break to the page image paragraph without a spacer."""

        section_properties = section_break._p.get_or_add_pPr().sectPr
        section_break._p.get_or_add_pPr().remove(section_properties)
        preceding_paragraph._p.get_or_add_pPr().append(section_properties)
        section_break._element.getparent().remove(section_break._element)

    @staticmethod
    def _notify(callback: StatusCallback | None, message: str) -> None:
        if callback is not None:
            try:
                callback(message)
            except Exception:
                LOGGER.exception("Status callback failed.")

    def close(self) -> None:
        self._started = False


def create_word_converter(visible: bool = False) -> WordPDFConverter | MacWordConverter:
    """Create the native Microsoft Word automation backend for this platform."""

    system = platform.system()
    if system == "Windows":
        return WordPDFConverter(visible=visible)
    if system == "Darwin":
        return MacWordConverter(visible=visible)
    raise WordNotInstalledError(
        f"Unsupported operating system: {system}",
        "This application requires Windows or macOS with the desktop version of Microsoft Word.",
    )


def convert_pdf_using_word(
    pdf_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str] | None = None,
    visible: bool = False,
    *,
    overwrite: bool = False,
    status_callback: StatusCallback | None = None,
) -> str:
    """Convenience wrapper using the native Word backend for this platform."""

    with create_word_converter(visible=visible) as converter:
        return converter.convert_pdf(
            pdf_path,
            output_path,
            overwrite=overwrite,
            status_callback=status_callback,
        )
