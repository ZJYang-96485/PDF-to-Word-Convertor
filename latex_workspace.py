"""Small local Overleaf-style workspace for LaTeX projects.

The workspace keeps the user's source file as the editable authority.  It
compiles a temporary copy so generated files and recovered PDF assets do not
pollute the source folder, displays the resulting PDF beside the source, and
hands the compiled PDF to the existing Word export backends.
"""

from __future__ import annotations

import base64
import logging
import os
import platform
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from converter import (
    PDFConversionError,
    Pdf2DocxConverter,
    VisualFidelityConverter,
    create_word_converter,
)


LOGGER = logging.getLogger(__name__)
INCLUDE_GRAPHICS_RE = re.compile(
    r"\\includegraphics(?:\s*\[[^]]*\])?\s*\{([^}]+)\}"
)


class LatexWorkspaceError(Exception):
    """Expected project or compilation error with a user-facing message."""

    def __init__(self, message: str, user_message: str | None = None) -> None:
        super().__init__(message)
        self.user_message = user_message or message


@dataclass
class CompileResult:
    pdf_path: Path
    log: str
    recovered_assets: list[str]
    build_dir: Path


@dataclass
class ExportResult:
    output_path: str
    mode: str


def _command_path(name: str) -> str | None:
    return shutil.which(name)


def _is_standalone_tex(source: str) -> bool:
    return "\\documentclass" in source and "\\begin{document}" in source


def _graphics_names(source: str) -> list[str]:
    names: list[str] = []
    for raw_name in INCLUDE_GRAPHICS_RE.findall(source):
        name = raw_name.strip()
        if name and name not in names:
            names.append(name)
    return names


def _find_graphic(source_dir: Path, requested_name: str) -> Path | None:
    requested = Path(requested_name)
    candidates = [source_dir / requested]
    if requested.suffix == "":
        candidates.extend(source_dir / f"{requested_name}{ext}" for ext in (".pdf", ".png", ".jpg", ".jpeg"))
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def _copy_existing_graphics(source_dir: Path, build_dir: Path, source: str) -> None:
    for requested_name in _graphics_names(source):
        existing = _find_graphic(source_dir, requested_name)
        if existing is None:
            continue
        destination = build_dir / Path(requested_name)
        destination.parent.mkdir(parents=True, exist_ok=True)
        if not destination.exists():
            shutil.copy2(existing, destination)


def _recover_graphics_from_pdf(
    source: str,
    reference_pdf: Path | None,
    build_dir: Path,
) -> list[str]:
    """Recover missing includegraphics files from a matching reference PDF."""

    missing_names = [
        name
        for name in _graphics_names(source)
        if not _find_graphic(build_dir, name)
    ]
    if not missing_names or reference_pdf is None or not reference_pdf.is_file():
        return []

    try:
        import pymupdf as fitz  # type: ignore[import-not-found]
    except ImportError:
        try:
            import fitz  # type: ignore[import-not-found]
        except ImportError as exc:
            raise LatexWorkspaceError(
                "PyMuPDF is required to recover missing figures from a reference PDF.",
                "Install the project requirements, or place the original figure files beside the .tex file.",
            ) from exc

    recovered: list[str] = []
    try:
        document = fitz.open(str(reference_pdf))
        with document:
            image_refs: list[int] = []
            for page in document:
                for image in page.get_images(full=True):
                    xref = int(image[0])
                    if xref not in image_refs:
                        image_refs.append(xref)

            if len(image_refs) < len(missing_names):
                raise LatexWorkspaceError(
                    f"The reference PDF contains {len(image_refs)} embedded figure(s), but the source needs {len(missing_names)}.",
                    "The source refers to figures that could not all be recovered from the matching PDF.",
                )

            for requested_name, xref in zip(missing_names, image_refs):
                destination = build_dir / Path(requested_name)
                destination.parent.mkdir(parents=True, exist_ok=True)
                pixmap = fitz.Pixmap(document, xref)
                try:
                    # The source usually specifies .jpg.  Pixmap.save() writes
                    # the format implied by that extension.
                    pixmap.save(str(destination))
                finally:
                    pixmap = None
                recovered.append(destination.name)
    except LatexWorkspaceError:
        raise
    except Exception as exc:
        LOGGER.exception("Could not recover figures from %s", reference_pdf)
        raise LatexWorkspaceError(
            f"Could not recover figures from {reference_pdf}: {exc}",
            "The matching PDF was found, but its figures could not be extracted for compilation.",
        ) from exc
    return recovered


def _wrapper_source(body_name: str) -> str:
    return f"""\\documentclass[10pt]{{exam}}
\\usepackage[margin=1in]{{geometry}}
\\usepackage{{amsmath,amssymb,graphicx}}
\\pointformat{{[\\thepoints]}}
\\printanswers
\\begin{{document}}
\\input{{{body_name}}}
\\end{{document}}
"""


def _compile_command(main_name: str, engine: str, build_dir: Path) -> list[str]:
    if engine == "auto" and _command_path("latexmk"):
        return [
            "latexmk",
            "-pdf",
            "-interaction=nonstopmode",
            "-halt-on-error",
            "-file-line-error",
            "-outdir=" + str(build_dir),
            main_name,
        ]

    selected = engine
    if selected == "auto":
        selected = "pdflatex"
    if not _command_path(selected):
        raise LatexWorkspaceError(
            f"The selected TeX compiler was not found: {selected}",
            f"Install {selected}, or choose a compiler that is already installed.",
        )
    return [
        selected,
        "-interaction=nonstopmode",
        "-halt-on-error",
        "-file-line-error",
        "-output-directory=" + str(build_dir),
        main_name,
    ]


def compile_latex(
    tex_path: Path,
    source_text: str,
    reference_pdf: Path | None = None,
    engine: str = "auto",
) -> CompileResult:
    """Compile a source buffer in an isolated temporary project directory."""

    if not tex_path.is_file():
        raise LatexWorkspaceError(
            f"The TeX source does not exist: {tex_path}",
            "Choose an existing .tex source file first.",
        )

    temp_root = Path(tempfile.mkdtemp(prefix="pdf_to_word_latex_"))
    project_dir = temp_root / "project"
    build_dir = temp_root / "build"
    project_dir.mkdir()
    build_dir.mkdir()

    standalone = _is_standalone_tex(source_text)
    if standalone:
        main_name = tex_path.name
        main_path = project_dir / main_name
        main_path.write_text(source_text, encoding="utf-8")
    else:
        body_name = "source_body.tex"
        (project_dir / body_name).write_text(source_text, encoding="utf-8")
        main_name = "main.tex"
        (project_dir / main_name).write_text(_wrapper_source(body_name), encoding="utf-8")

    _copy_existing_graphics(tex_path.parent, project_dir, source_text)
    recovered_assets = _recover_graphics_from_pdf(source_text, reference_pdf, project_dir)

    command = _compile_command(main_name, engine, build_dir)
    log_parts: list[str] = []
    try:
        first = subprocess.run(
            command,
            cwd=str(project_dir),
            capture_output=True,
            text=True,
            check=False,
        )
        log_parts.extend([first.stdout, first.stderr])

        # A direct engine needs a second pass for references and page layout.
        if first.returncode == 0 and command[0] != "latexmk":
            second = subprocess.run(
                command,
                cwd=str(project_dir),
                capture_output=True,
                text=True,
                check=False,
            )
            log_parts.extend([second.stdout, second.stderr])
            first = second
    except OSError as exc:
        raise LatexWorkspaceError(
            f"Could not start the TeX compiler: {exc}",
            "The TeX compiler could not be started. Check the compiler selection and your TeX installation.",
        ) from exc

    pdf_path = build_dir / Path(main_name).with_suffix(".pdf").name
    log = "\n".join(part for part in log_parts if part)
    if first.returncode != 0 or not pdf_path.is_file():
        raise LatexWorkspaceError(
            f"LaTeX compilation failed for {tex_path}:\n{log}",
            "LaTeX could not compile this source. Review the compiler log below for the line that needs attention.",
        )
    return CompileResult(pdf_path, log, recovered_assets, build_dir)


class LatexWorkspaceApp:
    """Tkinter UI for source editing, live PDF preview, and Word export."""

    def __init__(self, root: tk.Tk, initial_tex: str | None = None) -> None:
        self.root = root
        self.root.title("LaTeX Workspace")
        self.root.geometry("1440x900")
        self.root.minsize(980, 650)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.tex_path: Path | None = None
        self.reference_pdf: Path | None = None
        self.preview_pdf: Path | None = None
        self.preview_pages = 0
        self.preview_page_index = 0
        self.preview_image: tk.PhotoImage | None = None
        self.compile_result: CompileResult | None = None
        self.last_compiled_text = ""
        self.compile_thread: threading.Thread | None = None
        self.job_thread: threading.Thread | None = None
        self.compile_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.auto_compile = tk.BooleanVar(value=True)
        self.engine = tk.StringVar(value="auto")
        self.status_var = tk.StringVar(value="Open a .tex file to begin")
        self.path_var = tk.StringVar(value="No source selected")
        self.reference_var = tk.StringVar(value="Reference PDF: automatic")
        self.page_var = tk.StringVar(value="No preview")
        self._compile_timer: str | None = None

        self._build_ui()
        self.root.after(100, self._poll_queue)
        if initial_tex:
            self.root.after(100, lambda: self.open_tex(Path(initial_tex)))

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self.root, padding=(10, 8))
        toolbar.grid(row=0, column=0, sticky="ew")
        for column in range(8):
            toolbar.columnconfigure(column, weight=1 if column == 7 else 0)

        ttk.Button(toolbar, text="Open .tex", command=self._choose_tex).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(toolbar, text="Save Source", command=self.save_source).grid(row=0, column=1, padx=6)
        ttk.Button(toolbar, text="Compile / Preview", command=self.start_compile).grid(row=0, column=2, padx=6)
        ttk.Checkbutton(toolbar, text="Auto compile", variable=self.auto_compile).grid(row=0, column=3, padx=6)
        ttk.Label(toolbar, text="Engine:").grid(row=0, column=4, padx=(12, 3))
        ttk.Combobox(
            toolbar,
            textvariable=self.engine,
            values=("auto", "pdflatex", "xelatex", "lualatex"),
            state="readonly",
            width=10,
        ).grid(row=0, column=5, padx=3)
        ttk.Button(toolbar, text="Set Reference PDF", command=self._choose_reference_pdf).grid(row=0, column=6, padx=6)
        ttk.Label(toolbar, textvariable=self.status_var).grid(row=0, column=7, sticky="e", padx=(12, 0))

        content = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        content.grid(row=1, column=0, sticky="nsew", padx=10)

        source_frame = ttk.Frame(content, padding=(0, 0, 6, 0))
        source_frame.rowconfigure(2, weight=1)
        source_frame.columnconfigure(0, weight=1)
        ttk.Label(source_frame, textvariable=self.path_var).grid(row=0, column=0, sticky="w")
        ttk.Label(source_frame, text="Edit the original LaTeX source. The preview compiles a temporary copy.").grid(
            row=1, column=0, sticky="w", pady=(4, 6)
        )
        source_text_frame = ttk.Frame(source_frame)
        source_text_frame.grid(row=2, column=0, sticky="nsew")
        source_text_frame.rowconfigure(0, weight=1)
        source_text_frame.columnconfigure(0, weight=1)
        self.editor = tk.Text(
            source_text_frame,
            undo=True,
            wrap="none",
            font=("Menlo", 12),
            padx=10,
            pady=10,
        )
        self.editor.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(source_text_frame, orient="vertical", command=self.editor.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(source_text_frame, orient="horizontal", command=self.editor.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.editor.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.editor.bind("<KeyRelease>", self._on_source_changed)
        content.add(source_frame, weight=1)

        preview_frame = ttk.Frame(content, padding=(6, 0, 0, 0))
        preview_frame.rowconfigure(1, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_toolbar = ttk.Frame(preview_frame)
        preview_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(preview_toolbar, text="Previous", command=self._previous_page).pack(side="left")
        ttk.Button(preview_toolbar, text="Next", command=self._next_page).pack(side="left", padx=6)
        ttk.Label(preview_toolbar, textvariable=self.page_var).pack(side="left", padx=8)
        ttk.Label(preview_toolbar, textvariable=self.reference_var).pack(side="right")

        self.preview_canvas = tk.Canvas(preview_frame, background="#d7d7d7", highlightthickness=0)
        self.preview_canvas.grid(row=1, column=0, sticky="nsew")
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_canvas.yview)
        preview_scroll.grid(row=1, column=1, sticky="ns")
        self.preview_canvas.configure(yscrollcommand=preview_scroll.set)
        content.add(preview_frame, weight=1)

        bottom = ttk.Frame(self.root, padding=(10, 8))
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, text="Compiler log").grid(row=0, column=0, sticky="w")
        self.log_text = tk.Text(bottom, height=6, wrap="word", state="disabled", font=("Menlo", 10))
        self.log_text.grid(row=1, column=0, sticky="ew", pady=(4, 8))
        action_bar = ttk.Frame(bottom)
        action_bar.grid(row=2, column=0, sticky="ew")
        ttk.Button(action_bar, text="Export Exact Word", command=lambda: self.export_word(exact=True)).pack(side="left")
        ttk.Button(action_bar, text="Export Editable Word", command=lambda: self.export_word(exact=False)).pack(
            side="left", padx=8
        )
        ttk.Button(action_bar, text="Open Preview PDF", command=self.open_preview_pdf).pack(side="left")

    def _choose_tex(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Open LaTeX source",
            filetypes=[("LaTeX source", "*.tex"), ("All files", "*.*")],
        )
        if path:
            self.open_tex(Path(path))

    def open_tex(self, path: Path) -> None:
        try:
            source = path.read_text(encoding="utf-8")
        except (OSError, UnicodeError) as exc:
            messagebox.showerror("Could not open source", str(exc), parent=self.root)
            return
        self.tex_path = path.resolve()
        self.editor.delete("1.0", "end")
        self.editor.insert("1.0", source)
        self.last_compiled_text = ""
        candidate_pdf = self.tex_path.with_suffix(".pdf")
        self.reference_pdf = candidate_pdf if candidate_pdf.is_file() else None
        self.path_var.set(str(self.tex_path))
        self.reference_var.set(
            f"Reference PDF: {self.reference_pdf.name}" if self.reference_pdf else "Reference PDF: automatic"
        )
        self._append_log(f"Opened {self.tex_path}\n")
        self.status_var.set("Ready to compile")
        self._schedule_compile()

    def _choose_reference_pdf(self) -> None:
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Choose a reference PDF for missing figures",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self.reference_pdf = Path(path).resolve()
            self.reference_var.set(f"Reference PDF: {self.reference_pdf.name}")
            self._schedule_compile()

    def _on_source_changed(self, _event: tk.Event) -> None:
        if self.auto_compile.get():
            self._schedule_compile()

    def _schedule_compile(self) -> None:
        if self._compile_timer is not None:
            self.root.after_cancel(self._compile_timer)
        self._compile_timer = self.root.after(1200, self.start_compile)

    def save_source(self) -> None:
        if self.tex_path is None:
            messagebox.showinfo("No source", "Open a .tex file first.", parent=self.root)
            return
        try:
            self.tex_path.write_text(self.editor.get("1.0", "end-1c"), encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Could not save source", str(exc), parent=self.root)
            return
        self.status_var.set("Source saved")
        self._append_log(f"Saved {self.tex_path}\n")

    def start_compile(self) -> None:
        self._compile_timer = None
        if self.tex_path is None:
            return
        if self.compile_thread and self.compile_thread.is_alive():
            return
        source_text = self.editor.get("1.0", "end-1c")
        self.status_var.set("Compiling LaTeX...")
        self.compile_thread = threading.Thread(
            target=self._compile_worker,
            args=(self.tex_path, source_text, self.reference_pdf, self.engine.get()),
            daemon=True,
        )
        self.compile_thread.start()

    def _compile_worker(self, tex_path: Path, source_text: str, reference_pdf: Path | None, engine: str) -> None:
        try:
            result = compile_latex(tex_path, source_text, reference_pdf, engine)
            self.compile_queue.put(("compile_ok", result))
        except Exception as exc:
            self.compile_queue.put(("compile_error", exc))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.compile_queue.get_nowait()
                if kind == "compile_ok":
                    self._compile_finished(payload)  # type: ignore[arg-type]
                elif kind == "compile_error":
                    self._compile_failed(payload)  # type: ignore[arg-type]
                elif kind == "status":
                    self.status_var.set(str(payload))
                elif kind == "export_ok":
                    self._export_finished(payload)  # type: ignore[arg-type]
                elif kind == "export_error":
                    self._export_failed(payload)  # type: ignore[arg-type]
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _compile_finished(self, result: CompileResult) -> None:
        self.compile_result = result
        self.preview_pdf = result.pdf_path
        self.preview_page_index = 0
        self.last_compiled_text = self.editor.get("1.0", "end-1c")
        self.status_var.set("Compiled successfully")
        self._append_log(result.log)
        if result.recovered_assets:
            self._append_log("Recovered figures from reference PDF: " + ", ".join(result.recovered_assets) + "\n")
        self._load_preview_page()

    def _compile_failed(self, exc: Exception) -> None:
        self.status_var.set("Compilation failed")
        self._append_log(str(exc))
        if isinstance(exc, LatexWorkspaceError):
            self._append_log(f"\n{exc.user_message}\n")

    def _load_preview_page(self) -> None:
        if self.preview_pdf is None:
            return
        try:
            import pymupdf as fitz  # type: ignore[import-not-found]
        except ImportError:
            try:
                import fitz  # type: ignore[import-not-found]
            except ImportError:
                self._append_log("PyMuPDF is required for the PDF preview.\n")
                return
        try:
            document = fitz.open(str(self.preview_pdf))
            with document:
                self.preview_pages = document.page_count
                if self.preview_pages == 0:
                    return
                page = document.load_page(self.preview_page_index)
                available_width = max(self.preview_canvas.winfo_width() - 28, 500)
                scale = min(1.35, available_width / float(page.rect.width))
                pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                encoded = base64.b64encode(pixmap.tobytes("png")).decode("ascii")
            self.preview_image = tk.PhotoImage(data=encoded)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(14, 14, anchor="nw", image=self.preview_image)
            self.preview_canvas.configure(
                scrollregion=(0, 0, self.preview_image.width() + 28, self.preview_image.height() + 28)
            )
            self.page_var.set(f"Page {self.preview_page_index + 1} of {self.preview_pages}")
        except Exception as exc:
            self._append_log(f"Could not render preview: {exc}\n")

    def _previous_page(self) -> None:
        if self.preview_pages and self.preview_page_index > 0:
            self.preview_page_index -= 1
            self._load_preview_page()

    def _next_page(self) -> None:
        if self.preview_pages and self.preview_page_index + 1 < self.preview_pages:
            self.preview_page_index += 1
            self._load_preview_page()

    def open_preview_pdf(self) -> None:
        if self.preview_pdf and self.preview_pdf.is_file():
            self._open_path(self.preview_pdf)

    def export_word(self, exact: bool) -> None:
        if self.preview_pdf is None or self.compile_result is None:
            messagebox.showinfo("Compile first", "Compile the LaTeX source before exporting Word.", parent=self.root)
            return
        current_text = self.editor.get("1.0", "end-1c")
        if current_text != self.last_compiled_text:
            messagebox.showinfo("Compile first", "The source has changed since the last preview. Compile it again before exporting.", parent=self.root)
            return
        default_name = (self.tex_path.stem if self.tex_path else "document") + ".docx"
        output = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export Word document",
            initialfile=default_name,
            initialdir=str(self.tex_path.parent if self.tex_path else Path.cwd()),
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
        )
        if not output:
            return
        output_path = Path(output)
        if output_path.exists() and not messagebox.askyesno("Overwrite file", f"Replace {output_path.name}?", parent=self.root):
            return

        self.status_var.set("Exporting Word document...")
        self.job_thread = threading.Thread(
            target=self._export_worker,
            args=(self.preview_pdf, output_path, exact),
            daemon=True,
        )
        self.job_thread.start()

    def _export_worker(self, pdf_path: Path, output_path: Path, exact: bool) -> None:
        try:
            if exact:
                converter = VisualFidelityConverter()
                mode = "exact appearance"
            elif platform.system() == "Darwin" and not _command_path("osascript"):
                converter = Pdf2DocxConverter()
                mode = "editable pdf2docx fallback"
            else:
                converter = create_word_converter()
                mode = "editable Word conversion"
            try:
                result = converter.convert_pdf(
                    pdf_path,
                    output_path,
                    overwrite=True,
                    status_callback=self._thread_status,
                )
            finally:
                converter.close()
            self.compile_queue.put(("export_ok", ExportResult(result, mode)))
        except Exception as exc:
            self.compile_queue.put(("export_error", exc))

    def _thread_status(self, message: str) -> None:
        self.compile_queue.put(("status", message))

    def _export_finished(self, result: ExportResult) -> None:
        self.status_var.set("Word export complete")
        self._append_log(f"Created {result.mode}: {result.output_path}\n")
        if messagebox.askyesno("Export complete", f"Open {Path(result.output_path).name}?", parent=self.root):
            self._open_path(Path(result.output_path))

    def _export_failed(self, exc: Exception) -> None:
        self.status_var.set("Word export failed")
        message = exc.user_message if isinstance(exc, PDFConversionError) else str(exc)
        self._append_log(f"Export failed: {exc}\n")
        messagebox.showerror("Word export failed", message, parent=self.root)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    @staticmethod
    def _open_path(path: Path) -> None:
        try:
            if platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            elif platform.system() == "Windows":
                os.startfile(str(path))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except OSError as exc:
            LOGGER.warning("Could not open %s: %s", path, exc)

    def _on_close(self) -> None:
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    initial_tex = sys.argv[1] if len(sys.argv) > 1 else None
    LatexWorkspaceApp(root, initial_tex=initial_tex)
    root.mainloop()


if __name__ == "__main__":
    main()
