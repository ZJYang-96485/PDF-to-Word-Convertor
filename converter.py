"""Microsoft Word based PDF-to-DOCX conversion.

The conversion is intentionally delegated to Microsoft Word.  Word's PDF
importer is responsible for reflowing the document, which generally produces
the same result as opening the PDF manually in Word and saving it as DOCX.

This module does not import pywin32 at module import time.  That keeps the GUI
able to start on non-Windows systems and lets it show a useful Windows-only
message instead of failing with an ImportError.
"""

from __future__ import annotations

import gc
import logging
import os
import platform
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

        pdf = normalise_pdf_path(pdf_path)
        docx = normalise_output_path(pdf, output_path)

        if docx.exists() and not overwrite:
            raise OutputExistsError(
                f"Output already exists: {docx}",
                f"{docx.name} already exists. Choose a different output filename or approve overwrite.",
            )
        if docx.parent and not docx.parent.exists():
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


def convert_pdf_using_word(
    pdf_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str] | None = None,
    visible: bool = False,
    *,
    overwrite: bool = False,
    status_callback: StatusCallback | None = None,
) -> str:
    """Convenience wrapper for converting one PDF with Word."""

    with WordPDFConverter(visible=visible) as converter:
        return converter.convert_pdf(
            pdf_path,
            output_path,
            overwrite=overwrite,
            status_callback=status_callback,
        )
