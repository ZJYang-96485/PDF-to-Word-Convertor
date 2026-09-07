"""Tkinter desktop application for converting PDFs to Word documents."""

from __future__ import annotations

import logging
import os
import platform
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

from converter import (
    MacWordConverter,
    PDFConversionError,
    VisualFidelityConverter,
    WordPDFConverter,
    create_word_converter,
    normalise_output_path,
    normalise_pdf_path,
)


LOG_PATH = Path(__file__).resolve().with_name("pdf_to_word_converter.log")


def configure_logging() -> None:
    """Write technical details to a small local log without affecting the UI."""

    try:
        logging.basicConfig(
            filename=str(LOG_PATH),
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        )
    except OSError:
        # The application can still run from a read-only install directory.
        logging.basicConfig(level=logging.INFO)


LOGGER = logging.getLogger(__name__)


@dataclass
class BatchItem:
    pdf_path: str
    status: str = "Ready"


@dataclass
class BatchJob:
    item: BatchItem
    output_path: str
    overwrite: bool


class PDFConverterApp:
    """Main window and thread-safe orchestration for single/batch conversion."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.platform = platform.system()
        self.is_supported = self.platform in {"Windows", "Darwin"}
        self.selected_pdf: str | None = None
        self.batch_items: list[BatchItem] = []
        # Exact appearance is the safe default for LaTeX, equations, fonts,
        # and diagrams. Native Word reflow remains available as an opt-in.
        self.visual_mode_var = tk.BooleanVar(value=True)
        self.is_converting = False
        self.cancel_event = threading.Event()
        self.close_when_done = False

        self.root.title("PDF to Word Converter")
        self.root.geometry("700x500")
        self.root.minsize(620, 450)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.selected_path_var = tk.StringVar(value="No PDF selected")
        self.output_path_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Ready")
        self._build_ui()
        self.output_path_var.trace_add("write", self._on_output_changed)
        self._update_button_states()

        if not self.is_supported:
            self.root.after(150, self._show_platform_required)

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=20)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        style = ttk.Style(self.root)
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
        style.configure("Section.TLabelframe", padding=12)
        style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 9))

        header = ttk.Frame(main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="PDF to Word Converter", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Preserve PDF page appearance by default; optionally use Word's editable conversion on Windows or macOS.",
            style="Subtitle.TLabel",
            wraplength=620,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Button(
            header,
            text="Open LaTeX Workspace",
            command=self._open_latex_workspace,
        ).grid(row=0, column=1, rowspan=2, sticky="e", padx=(12, 0))

        single = ttk.LabelFrame(main, text="Single PDF", style="Section.TLabelframe")
        single.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        single.columnconfigure(1, weight=1)

        self.select_button = ttk.Button(single, text="Select PDF File", command=self._select_pdf)
        self.select_button.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        ttk.Label(single, text="Selected file:").grid(row=0, column=1, sticky="w", pady=(0, 8))

        self.selected_path_entry = ttk.Entry(
            single, textvariable=self.selected_path_var, state="readonly"
        )
        self.selected_path_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        ttk.Label(single, text="Output:").grid(row=2, column=0, sticky="w", padx=(0, 10))
        self.output_entry = ttk.Entry(single, textvariable=self.output_path_var)
        self.output_entry.grid(row=2, column=1, sticky="ew", padx=(0, 8))
        self.output_browse_button = ttk.Button(
            single, text="Browse Output Location", command=self._browse_output
        )
        self.output_browse_button.grid(row=2, column=2, sticky="e")

        self.visual_mode_checkbutton = ttk.Checkbutton(
            single,
            text="Preserve exact PDF appearance (recommended; avoids Word automation)",
            variable=self.visual_mode_var,
        )
        self.visual_mode_checkbutton.grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )
        ttk.Label(
            single,
            text="Enabled by default for LaTeX, equations, fonts, diagrams, and page ordering. Uncheck only when you need editable Word text.",
            style="Subtitle.TLabel",
            wraplength=620,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(3, 0))

        batch = ttk.LabelFrame(main, text="Batch conversion", style="Section.TLabelframe")
        batch.grid(row=2, column=0, sticky="nsew", pady=(0, 12))
        batch.columnconfigure(0, weight=1)
        batch.rowconfigure(1, weight=1)

        batch_toolbar = ttk.Frame(batch)
        batch_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.add_button = ttk.Button(batch_toolbar, text="Add PDFs", command=self._add_pdfs)
        self.add_button.grid(row=0, column=0, sticky="w")
        self.clear_button = ttk.Button(batch_toolbar, text="Clear List", command=self._clear_batch)
        self.clear_button.grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(
            batch_toolbar,
            text="Files are saved beside their source PDFs by default.",
            style="Subtitle.TLabel",
        ).grid(row=0, column=2, sticky="e", padx=(12, 0))
        batch_toolbar.columnconfigure(2, weight=1)

        tree_frame = ttk.Frame(batch)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.batch_tree = ttk.Treeview(
            tree_frame, columns=("file", "status"), show="headings", height=5
        )
        self.batch_tree.heading("file", text="File")
        self.batch_tree.heading("status", text="Status")
        self.batch_tree.column("file", width=430, anchor="w", stretch=True)
        self.batch_tree.column("status", width=130, anchor="w", stretch=False)
        self.batch_tree.grid(row=0, column=0, sticky="nsew")
        batch_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.batch_tree.yview)
        batch_scroll.grid(row=0, column=1, sticky="ns")
        self.batch_tree.configure(yscrollcommand=batch_scroll.set)

        actions = ttk.Frame(main)
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        actions.columnconfigure(2, weight=1)
        self.convert_button = ttk.Button(
            actions,
            text="Convert to Word",
            style="Primary.TButton",
            command=self._begin_single_conversion,
        )
        self.convert_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.convert_all_button = ttk.Button(
            actions,
            text="Convert All",
            style="Primary.TButton",
            command=self._begin_batch_conversion,
        )
        self.convert_all_button.grid(row=0, column=1, sticky="ew", padx=5)
        self.cancel_button = ttk.Button(
            actions, text="Cancel", command=self._request_cancel, state="disabled"
        )
        self.cancel_button.grid(row=0, column=2, sticky="ew", padx=(5, 0))

        progress_frame = ttk.Frame(main)
        progress_frame.grid(row=4, column=0, sticky="ew")
        progress_frame.columnconfigure(1, weight=1)
        ttk.Label(progress_frame, text="Status:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=0, column=1, sticky="w")
        self.progress = ttk.Progressbar(progress_frame, mode="indeterminate")
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(7, 0))

    def _show_platform_required(self) -> None:
        messagebox.showerror(
            "Windows or macOS Required",
            "This application requires Windows or macOS with desktop Microsoft Word because it uses Word's native PDF conversion engine.",
            parent=self.root,
        )
        self.status_var.set("Windows/macOS and Microsoft Word are required")

    def _open_latex_workspace(self) -> None:
        """Launch the source-aware editor and live PDF preview."""

        workspace_script = Path(__file__).resolve().with_name("latex_workspace.py")
        try:
            subprocess.Popen([sys.executable, str(workspace_script)])
        except OSError as exc:
            messagebox.showerror(
                "Could not open LaTeX Workspace",
                f"The LaTeX workspace could not be started: {exc}",
                parent=self.root,
            )

    def _select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self._set_selected_pdf(path)

    def _set_selected_pdf(self, path: str) -> None:
        try:
            pdf = normalise_pdf_path(path)
        except PDFConversionError as exc:
            messagebox.showerror("Invalid PDF", exc.user_message, parent=self.root)
            return

        self.selected_pdf = str(pdf)
        self.selected_path_var.set(str(pdf))
        self.output_path_var.set(str(normalise_output_path(pdf)))
        self.status_var.set("Ready")
        self._update_button_states()

    def _browse_output(self) -> None:
        if not self.selected_pdf:
            return
        initial = self.output_path_var.get() or str(normalise_output_path(self.selected_pdf))
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Choose Word output location",
            initialfile=Path(initial).name,
            initialdir=str(Path(initial).parent),
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.output_path_var.set(path)

    def _add_pdfs(self) -> None:
        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="Add PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not paths:
            return

        existing = {os.path.normcase(os.path.abspath(item.pdf_path)) for item in self.batch_items}
        invalid: list[str] = []
        for path in paths:
            try:
                pdf = normalise_pdf_path(path)
            except PDFConversionError:
                invalid.append(path)
                continue
            key = os.path.normcase(str(pdf))
            if key not in existing:
                self.batch_items.append(BatchItem(str(pdf)))
                existing.add(key)

        self._refresh_batch_tree()
        self.status_var.set(f"{len(self.batch_items)} PDF(s) ready")
        self._update_button_states()
        if invalid:
            messagebox.showwarning(
                "Some files were skipped",
                "Only existing PDF files can be added. One or more selected files were skipped.",
                parent=self.root,
            )

    def _clear_batch(self) -> None:
        self.batch_items.clear()
        self._refresh_batch_tree()
        self.status_var.set("Ready")
        self._update_button_states()

    def _refresh_batch_tree(self) -> None:
        for child in self.batch_tree.get_children():
            self.batch_tree.delete(child)
        for index, item in enumerate(self.batch_items):
            self.batch_tree.insert(
                "", "end", iid=str(index), values=(Path(item.pdf_path).name, item.status)
            )

    def _begin_single_conversion(self) -> None:
        if not self.selected_pdf:
            return

        output = self.output_path_var.get().strip()
        try:
            output_path = normalise_output_path(self.selected_pdf, output)
        except PDFConversionError as exc:
            messagebox.showerror("Invalid output", exc.user_message, parent=self.root)
            return

        overwrite = self._confirm_overwrite(output_path)
        if overwrite is None:
            self.status_var.set("Conversion cancelled")
            return
        self._start_conversion(
            [BatchJob(BatchItem(self.selected_pdf), str(output_path), overwrite)],
            batch=False,
            visual_mode=self.visual_mode_var.get(),
        )

    def _begin_batch_conversion(self) -> None:
        jobs: list[BatchJob] = []
        for item in self.batch_items:
            item.status = "Waiting"
            output_path = normalise_output_path(item.pdf_path)
            overwrite = self._confirm_overwrite(output_path)
            if overwrite is None:
                item.status = "Skipped"
            else:
                jobs.append(BatchJob(item, str(output_path), overwrite))
        self._refresh_batch_tree()

        if not jobs:
            self.status_var.set("No files to convert")
            return
        self._start_conversion(jobs, batch=True, visual_mode=self.visual_mode_var.get())

    def _confirm_overwrite(self, output_path: Path) -> bool | None:
        if not output_path.exists():
            return False
        if messagebox.askyesno(
            "Overwrite existing file?",
            f"{output_path.name} already exists.\n\nDo you want to overwrite it?",
            parent=self.root,
        ):
            return True
        return None

    def _start_conversion(
        self, jobs: list[BatchJob], *, batch: bool, visual_mode: bool
    ) -> None:
        self.is_converting = True
        self.close_when_done = False
        self.cancel_event.clear()
        self._set_busy(True)
        self.status_var.set(
            "Preparing exact-appearance Word document..."
            if visual_mode
            else "Starting Microsoft Word..."
        )
        worker = threading.Thread(
            target=self._conversion_worker,
            args=(jobs, batch, visual_mode),
            daemon=True,
            name="pdf-to-word-converter",
        )
        worker.start()

    def _conversion_worker(
        self, jobs: list[BatchJob], batch: bool, visual_mode: bool
    ) -> None:
        converter: WordPDFConverter | MacWordConverter | VisualFidelityConverter | None = None
        completed: list[tuple[BatchJob, str]] = []
        failures: list[tuple[BatchJob, Exception]] = []
        cancelled = False

        try:
            converter = (
                VisualFidelityConverter()
                if visual_mode
                else create_word_converter(visible=False)
            )
            converter.start()
            for job_index, job in enumerate(jobs):
                if self.cancel_event.is_set():
                    cancelled = True
                    self._post(self._set_job_status, job, "Cancelled")
                    continue

                filename = Path(job.item.pdf_path).name
                self._post(self._set_job_status, job, "Converting...")

                def report(message: str, file_name: str = filename) -> None:
                    prefix = f"{file_name}: " if batch else ""
                    self._post(self.status_var.set, prefix + message)

                try:
                    result = converter.convert_pdf(
                        job.item.pdf_path,
                        job.output_path,
                        overwrite=job.overwrite,
                        status_callback=report,
                    )
                    completed.append((job, result))
                    self._post(self._set_job_status, job, "Complete")
                except Exception as exc:  # Keep batch conversion moving file-by-file.
                    LOGGER.exception("Conversion failed for %s", job.item.pdf_path)
                    failures.append((job, exc))
                    self._post(self._set_job_status, job, "Failed")

                if self.cancel_event.is_set():
                    cancelled = True
                    # The current Word operation has finished; do not begin another file.
                    for remaining in jobs[job_index + 1 :]:
                        self._post(self._set_job_status, remaining, "Cancelled")
                    break
        except Exception as exc:
            LOGGER.exception("Could not start or run Microsoft Word conversion.")
            # A startup failure affects every file that has not run yet.
            for job in jobs:
                if not any(existing_job is job for existing_job, _ in completed + failures):
                    failures.append((job, exc))
                    self._post(self._set_job_status, job, "Failed")
        finally:
            if converter is not None:
                converter.close()
            self._post(self._conversion_finished, jobs, completed, failures, cancelled, batch)

    def _conversion_finished(
        self,
        jobs: list[BatchJob],
        completed: list[tuple[BatchJob, str]],
        failures: list[tuple[BatchJob, Exception]],
        cancelled: bool,
        batch: bool,
    ) -> None:
        self.is_converting = False
        self._set_busy(False)

        if self.close_when_done:
            self.root.destroy()
            return

        if not batch:
            if completed:
                self.status_var.set("Conversion complete.")
                self._show_success_dialog(completed[0][1])
            elif failures:
                self.status_var.set("Conversion failed.")
                self._show_error(failures[0][1])
            elif cancelled:
                self.status_var.set("Conversion cancelled")
            return

        complete_count = len(completed)
        failed_count = len(failures)
        cancelled_count = sum(1 for job in jobs if job.item.status == "Cancelled")
        self.status_var.set(
            f"Batch complete: {complete_count} complete, {failed_count} failed"
        )
        summary = (
            f"Completed: {complete_count}\n"
            f"Failed: {failed_count}\n"
            f"Skipped or cancelled: {cancelled_count + sum(1 for item in self.batch_items if item.status == 'Skipped')}"
        )
        if cancelled:
            summary += "\n\nThe batch was stopped after the current file finished."
        if failures:
            summary += "\n\nFailed files are marked in the list and technical details are in the log file."
            if not completed and all(isinstance(error, PDFConversionError) for _, error in failures):
                summary += f"\n\n{failures[0][1].user_message}"
        messagebox.showinfo("Batch Conversion Complete", summary, parent=self.root)

    def _show_success_dialog(self, docx_path: str) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Conversion Complete")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

        frame = ttk.Frame(dialog, padding=20)
        frame.grid(row=0, column=0, sticky="nsew")
        ttk.Label(frame, text="Conversion Complete", font=("Segoe UI", 13, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            frame,
            text="Word document created successfully:",
        ).grid(row=1, column=0, sticky="w", pady=(12, 4))
        path_entry = ttk.Entry(frame, width=72)
        path_entry.insert(0, docx_path)
        path_entry.configure(state="readonly")
        path_entry.grid(row=2, column=0, sticky="ew")

        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=0, sticky="e", pady=(18, 0))
        ttk.Button(buttons, text="Open Word File", command=lambda: self._open_word_file(docx_path)).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(buttons, text="Open Folder", command=lambda: self._open_folder(docx_path)).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(buttons, text="Close", command=dialog.destroy).grid(row=0, column=2)
        dialog.update_idletasks()
        self._center_window(dialog)
        dialog.grab_set()
        dialog.focus_set()

    def _open_word_file(self, docx_path: str) -> None:
        try:
            if self.platform == "Windows":
                if not hasattr(os, "startfile"):
                    raise OSError("Windows file launching is unavailable.")
                os.startfile(docx_path)  # type: ignore[attr-defined]
            elif self.platform == "Darwin":
                subprocess.Popen(["open", docx_path])
            else:
                raise OSError("Opening files is supported on Windows and macOS only.")
        except OSError as exc:
            LOGGER.exception("Could not open Word file: %s", docx_path)
            messagebox.showerror("Could not open file", str(exc), parent=self.root)

    def _open_folder(self, docx_path: str) -> None:
        try:
            if self.platform == "Windows":
                subprocess.Popen(["explorer", "/select,", os.path.normpath(docx_path)])
            elif self.platform == "Darwin":
                subprocess.Popen(["open", "-R", docx_path])
            else:
                raise OSError("Opening folders is supported on Windows and macOS only.")
        except OSError as exc:
            LOGGER.exception("Could not open Explorer for: %s", docx_path)
            messagebox.showerror("Could not open folder", str(exc), parent=self.root)

    def _show_error(self, error: Exception) -> None:
        if isinstance(error, PDFConversionError):
            messagebox.showerror("Conversion Error", error.user_message, parent=self.root)
        else:
            messagebox.showerror(
                "Conversion Error",
                "The conversion could not be completed. Technical details were written to the log file.",
                parent=self.root,
            )

    def _request_cancel(self) -> None:
        if not self.is_converting:
            return
        self.cancel_event.set()
        self.cancel_button.state(["disabled"])
        self.status_var.set("Cancellation requested; finishing the current file...")

    def _set_busy(self, busy: bool) -> None:
        controls = [
            self.select_button,
            self.output_browse_button,
            self.add_button,
            self.clear_button,
            self.convert_button,
            self.convert_all_button,
            self.visual_mode_checkbutton,
        ]
        for control in controls:
            if busy:
                control.state(["disabled"])
            else:
                control.state(["!disabled"])
        if busy:
            self.output_entry.configure(state="disabled")
            self.cancel_button.state(["!disabled"])
            self.progress.start(12)
        else:
            self.output_entry.configure(state="normal")
            self.cancel_button.state(["disabled"])
            self.progress.stop()
            self._update_button_states()

    def _update_button_states(self, *_args: Any) -> None:
        if self.is_converting:
            return
        if not self.is_supported:
            for control in (
                self.select_button,
                self.output_browse_button,
                self.add_button,
                self.clear_button,
                self.convert_button,
                self.convert_all_button,
            ):
                control.state(["disabled"])
            return

        self.select_button.state(["!disabled"])
        self.output_browse_button.state(["!disabled"] if self.selected_pdf else ["disabled"])
        self.add_button.state(["!disabled"])
        self.clear_button.state(["!disabled"] if self.batch_items else ["disabled"])
        self.convert_button.state(
            ["!disabled"]
            if self.selected_pdf and self.output_path_var.get().strip()
            else ["disabled"]
        )
        self.convert_all_button.state(["!disabled"] if self.batch_items else ["disabled"])

    def _on_output_changed(self, *_args: Any) -> None:
        self._update_button_states()

    def _set_job_status(self, job: BatchJob, status: str) -> None:
        job.item.status = status
        self._refresh_batch_tree()

    def _post(self, callback: Callable[..., None], *args: Any) -> None:
        try:
            self.root.after(0, callback, *args)
        except tk.TclError:
            # The window may have been destroyed after the worker completed.
            pass

    def _on_close(self) -> None:
        if not self.is_converting:
            self.root.destroy()
            return
        if self.close_when_done:
            return
        should_close = messagebox.askyesno(
            "Conversion in progress",
            "A conversion is currently running. Exit after it finishes?",
            parent=self.root,
        )
        if should_close:
            self.close_when_done = True
            self.cancel_event.set()
            self.status_var.set("Finishing the current file before exiting...")

    @staticmethod
    def _center_window(window: tk.Toplevel) -> None:
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = max((window.winfo_screenwidth() - width) // 2, 0)
        y = max((window.winfo_screenheight() - height) // 2, 0)
        window.geometry(f"{width}x{height}+{x}+{y}")


def main() -> None:
    configure_logging()
    root = tk.Tk()
    PDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
