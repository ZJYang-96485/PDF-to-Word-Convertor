"""Tkinter desktop application for high-fidelity PDF-to-Word conversion."""

from __future__ import annotations

import logging
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Iterable

from converter import (
    ConversionCancelledError,
    PDFConversionError,
    convert_pdf_to_docx,
    normalise_output_path,
    normalise_pdf_path,
)


LOGGER = logging.getLogger(__name__)
LOG_PATH = Path(__file__).resolve().with_name("pdf_to_word_converter.log")


@dataclass
class BatchItem:
    pdf_path: Path
    tree_id: str


@dataclass
class ConversionJob:
    pdf_path: Path
    output_path: Path
    tree_id: str | None = None


def configure_logging() -> None:
    try:
        logging.basicConfig(
            filename=str(LOG_PATH),
            level=logging.INFO,
            format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        )
    except OSError:
        logging.basicConfig(level=logging.INFO)


class PDFConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PDF to Word Convertor")
        self.root.geometry("780x620")
        self.root.minsize(700, 520)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.selected_pdf: Path | None = None
        self.batch_items: dict[str, BatchItem] = {}
        self.cancel_event = threading.Event()
        self.is_converting = False
        self.close_when_done = False

        self.selected_path_var = tk.StringVar(value="No PDF selected")
        self.output_path_var = tk.StringVar(value="")
        self.dpi_var = tk.IntVar(value=300)
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()
        self._update_button_states()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main = ttk.Frame(self.root, padding=20)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        style = ttk.Style(self.root)
        style.configure("Title.TLabel", font=("TkDefaultFont", 20, "bold"))
        style.configure("Subtitle.TLabel", font=("TkDefaultFont", 10))
        style.configure("Primary.TButton", padding=(16, 8))

        header = ttk.Frame(main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="PDF to Word Convertor", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text=(
                "Preserve LaTeX, fonts, diagrams, spacing, and page ordering "
                "by placing each rendered PDF page into Word."
            ),
            style="Subtitle.TLabel",
            wraplength=700,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        settings = ttk.LabelFrame(main, text="Conversion settings", padding=12)
        settings.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        settings.columnconfigure(1, weight=1)
        ttk.Label(settings, text="Rendering resolution:").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        self.dpi_spinbox = ttk.Spinbox(
            settings,
            from_=72,
            to=600,
            increment=1,
            textvariable=self.dpi_var,
            width=8,
        )
        self.dpi_spinbox.grid(row=0, column=1, sticky="w")
        ttk.Label(settings, text="DPI; 300 is recommended for high-quality output.").grid(
            row=0, column=2, sticky="w", padx=(10, 0)
        )

        single = ttk.LabelFrame(main, text="Single PDF", padding=12)
        single.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        single.columnconfigure(1, weight=1)
        ttk.Button(single, text="Select PDF", command=self._select_pdf).grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        self.selected_path_entry = ttk.Entry(
            single, textvariable=self.selected_path_var, state="readonly"
        )
        self.selected_path_entry.grid(row=0, column=1, columnspan=2, sticky="ew")
        ttk.Label(single, text="Output Word file:").grid(
            row=1, column=0, sticky="w", pady=(10, 0), padx=(0, 10)
        )
        self.output_entry = ttk.Entry(single, textvariable=self.output_path_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", pady=(10, 0))
        self.output_browse_button = ttk.Button(
            single, text="Browse", command=self._browse_output
        )
        self.output_browse_button.grid(row=1, column=2, sticky="e", pady=(10, 0), padx=(8, 0))

        batch = ttk.LabelFrame(main, text="Batch conversion", padding=12)
        batch.grid(row=3, column=0, sticky="nsew", pady=(0, 12))
        batch.columnconfigure(0, weight=1)
        batch.rowconfigure(1, weight=1)
        toolbar = ttk.Frame(batch)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.add_button = ttk.Button(toolbar, text="Add PDFs", command=self._add_pdfs)
        self.add_button.grid(row=0, column=0, sticky="w")
        self.clear_button = ttk.Button(toolbar, text="Clear List", command=self._clear_batch)
        self.clear_button.grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(
            toolbar,
            text="Batch outputs are saved beside their source PDFs.",
            style="Subtitle.TLabel",
        ).grid(row=0, column=2, sticky="e", padx=(12, 0))
        toolbar.columnconfigure(2, weight=1)

        tree_frame = ttk.Frame(batch)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        self.batch_tree = ttk.Treeview(
            tree_frame, columns=("file", "status"), show="headings", height=6
        )
        self.batch_tree.heading("file", text="PDF file")
        self.batch_tree.heading("status", text="Status")
        self.batch_tree.column("file", width=520, anchor="w", stretch=True)
        self.batch_tree.column("status", width=140, anchor="w", stretch=False)
        self.batch_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.batch_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.batch_tree.configure(yscrollcommand=scrollbar.set)

        actions = ttk.Frame(main)
        actions.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        actions.columnconfigure(2, weight=1)
        self.convert_button = ttk.Button(
            actions,
            text="Convert Selected",
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

        status_frame = ttk.Frame(main)
        status_frame.grid(row=5, column=0, sticky="ew")
        status_frame.columnconfigure(1, weight=1)
        ttk.Label(status_frame, text="Status:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(status_frame, textvariable=self.status_var).grid(row=0, column=1, sticky="w")
        self.progress = ttk.Progressbar(status_frame, mode="indeterminate")
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(7, 0))

    def _select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            pdf = normalise_pdf_path(path)
        except PDFConversionError as exc:
            messagebox.showerror("Invalid PDF", str(exc), parent=self.root)
            return
        self.selected_pdf = pdf
        self.selected_path_var.set(str(pdf))
        self.output_path_var.set(str(normalise_output_path(pdf)))
        self.status_var.set("Ready")
        self._update_button_states()

    def _browse_output(self) -> None:
        if not self.selected_pdf:
            return
        current = Path(self.output_path_var.get() or normalise_output_path(self.selected_pdf))
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Choose Word output location",
            initialdir=str(current.parent),
            initialfile=current.name,
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
        existing = {item.pdf_path for item in self.batch_items.values()}
        for raw_path in paths:
            try:
                pdf = normalise_pdf_path(raw_path)
            except PDFConversionError:
                continue
            if pdf in existing:
                continue
            tree_id = f"item-{len(self.batch_items) + 1}"
            self.batch_items[tree_id] = BatchItem(pdf, tree_id)
            self.batch_tree.insert("", "end", iid=tree_id, values=(str(pdf), "Ready"))
            existing.add(pdf)
        self._update_button_states()

    def _clear_batch(self) -> None:
        for tree_id in list(self.batch_items):
            self.batch_tree.delete(tree_id)
        self.batch_items.clear()
        self._update_button_states()

    def _begin_single_conversion(self) -> None:
        if not self.selected_pdf:
            return
        try:
            output = normalise_output_path(self.selected_pdf, self.output_path_var.get())
        except PDFConversionError as exc:
            messagebox.showerror("Invalid output", str(exc), parent=self.root)
            return
        overwrite = self._confirm_overwrite([output])
        if overwrite is None:
            return
        self._start_jobs([ConversionJob(self.selected_pdf, output)], overwrite)

    def _begin_batch_conversion(self) -> None:
        jobs = [
            ConversionJob(item.pdf_path, normalise_output_path(item.pdf_path), item.tree_id)
            for item in self.batch_items.values()
        ]
        if not jobs:
            return
        overwrite = self._confirm_overwrite([job.output_path for job in jobs])
        if overwrite is None:
            return
        self._start_jobs(jobs, overwrite)

    def _confirm_overwrite(self, outputs: Iterable[Path]) -> bool | None:
        existing = [path for path in outputs if path.exists()]
        if not existing:
            return False
        names = "\n".join(f"- {path.name}" for path in existing[:8])
        if len(existing) > 8:
            names += f"\n- and {len(existing) - 8} more"
        approved = messagebox.askyesno(
            "Overwrite existing Word files?",
            f"These output files already exist:\n{names}\n\nOverwrite them?",
            parent=self.root,
        )
        return True if approved else None

    def _start_jobs(self, jobs: list[ConversionJob], overwrite: bool) -> None:
        self.is_converting = True
        self.cancel_event.clear()
        self.status_var.set("Starting conversion...")
        self.progress.start(10)
        self._update_button_states()
        worker = threading.Thread(
            target=self._run_jobs,
            args=(jobs, overwrite),
            daemon=True,
        )
        worker.start()

    def _run_jobs(self, jobs: list[ConversionJob], overwrite: bool) -> None:
        completed: list[Path] = []
        try:
            for job in jobs:
                if self.cancel_event.is_set():
                    break

                self._set_tree_status(job.tree_id, "Converting")

                def notify(message: str, name=job.pdf_path.name) -> None:
                    self.root.after(0, self.status_var.set, f"{name}: {message}")

                output = convert_pdf_to_docx(
                    job.pdf_path,
                    job.output_path,
                    dpi=self._current_dpi(),
                    overwrite=overwrite,
                    status_callback=notify,
                    cancel_event=self.cancel_event,
                )
                completed.append(output)
                self._set_tree_status(job.tree_id, "Converted")

            cancelled = self.cancel_event.is_set()
            self.root.after(0, self._worker_done, completed, cancelled, None)
        except ConversionCancelledError:
            self.root.after(0, self._worker_done, completed, True, None)
        except Exception as exc:
            LOGGER.exception("Conversion failed")
            self.root.after(0, self._worker_done, completed, False, str(exc))

    def _current_dpi(self) -> int:
        try:
            dpi = int(self.dpi_var.get())
        except (TypeError, ValueError):
            raise PDFConversionError("Rendering DPI must be a whole number between 72 and 600.")
        if not 72 <= dpi <= 600:
            raise PDFConversionError("Rendering DPI must be between 72 and 600.")
        return dpi

    def _set_tree_status(self, tree_id: str | None, status: str) -> None:
        if tree_id:
            self.root.after(0, self.batch_tree.set, tree_id, "status", status)

    def _worker_done(
        self, completed: list[Path], cancelled: bool, error_message: str | None
    ) -> None:
        self.is_converting = False
        self.progress.stop()
        self._update_button_states()

        if error_message:
            self.status_var.set("Conversion failed")
            messagebox.showerror("Conversion failed", error_message, parent=self.root)
        elif cancelled:
            self.status_var.set(f"Cancelled after {len(completed)} file(s)")
        else:
            self.status_var.set(f"Completed {len(completed)} file(s)")
            if completed:
                messagebox.showinfo(
                    "Conversion complete",
                    "Created:\n" + "\n".join(str(path) for path in completed),
                    parent=self.root,
                )

        if self.close_when_done:
            self.root.destroy()

    def _request_cancel(self) -> None:
        if self.is_converting:
            self.cancel_event.set()
            self.status_var.set("Cancelling after the current page...")

    def _update_button_states(self) -> None:
        normal = "disabled" if self.is_converting else "normal"
        self.convert_button.configure(state=normal if self.selected_pdf else "disabled")
        self.convert_all_button.configure(
            state=normal if self.batch_items else "disabled"
        )
        self.add_button.configure(state=normal)
        self.clear_button.configure(state=normal if self.batch_items else "disabled")
        self.output_browse_button.configure(
            state=normal if self.selected_pdf else "disabled"
        )
        self.dpi_spinbox.configure(state="disabled" if self.is_converting else "normal")
        self.cancel_button.configure(state="normal" if self.is_converting else "disabled")

    def _on_close(self) -> None:
        if not self.is_converting:
            self.root.destroy()
            return
        if messagebox.askyesno(
            "Cancel conversion?",
            "A conversion is still running. Cancel after the current page and close?",
            parent=self.root,
        ):
            self.close_when_done = True
            self.cancel_event.set()


def main() -> None:
    configure_logging()
    root = tk.Tk()
    PDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
