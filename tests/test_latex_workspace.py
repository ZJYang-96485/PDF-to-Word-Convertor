from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement

from latex_workspace import (
    CompileResult,
    LatexSourceDocxConverter,
    LatexWorkspaceError,
    _deduplicate_placeholder_labels,
    _find_reference_pdf,
    _mark_roman_math,
    _tex4ht_source,
)


def _math_run(text: str) -> object:
    run = OxmlElement("m:r")
    node = OxmlElement("m:t")
    node.text = text
    run.append(node)
    return run


class RomanMathTests(unittest.TestCase):
    def test_standard_operators_remain_native_tex4ht_functions(self) -> None:
        source = r"$\ln x + \exp(-x) + \mathrm{II}$"

        marked = _mark_roman_math(source)

        self.assertIn(r"\ln x", marked)
        self.assertIn(r"\exp(-x)", marked)
        self.assertEqual(marked.count("PDFRomanS"), 1)
        self.assertRegex(marked, r"\\mathrm\{\\text\{PDFRomanS\d{3}\}II")

    def test_explicit_roman_letters_merge_into_one_omml_run(self) -> None:
        document = Document()
        paragraph = document.add_paragraph()
        formula = OxmlElement("m:oMath")
        for value in ("PDFRomanS001", "I", "I", "PDFRomanE001"):
            formula.append(_math_run(value))
        paragraph._p.append(formula)

        LatexSourceDocxConverter._restore_math_roman_styles(document)

        namespace = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
        runs = formula.xpath("./m:r", namespaces=namespace)
        self.assertEqual(
            ["".join(run.xpath("./m:t/text()", namespaces=namespace)) for run in runs],
            ["II"],
        )
        self.assertEqual(runs[0].xpath("./m:rPr/m:scr/@m:val", namespaces=namespace), ["roman"])
        self.assertEqual(runs[0].xpath("./m:rPr/m:sty/@m:val", namespaces=namespace), ["p"])

    def test_repeated_greater_than_is_exported_as_one_relation(self) -> None:
        transformed = _tex4ht_source(r"Energy $\Delta\varepsilon>>kT$ is large.")

        self.assertIn(r"$\Delta\varepsilon\gg{}kT$", transformed)


class PlaceholderLabelTests(unittest.TestCase):
    def test_repeated_active_placeholder_labels_are_removed(self) -> None:
        source = (
            "\\begin{figure}\\label{fig:placeholder}First\\end{figure}\n"
            "\\begin{figure}\\label { fig:placeholder }Second\\end{figure}\n"
        )

        transformed, removed = _deduplicate_placeholder_labels(source)

        self.assertEqual(removed, 1)
        self.assertEqual(transformed.count("fig:placeholder"), 1)
        self.assertIn("Second", transformed)

    def test_commented_placeholder_labels_do_not_count_as_active(self) -> None:
        source = (
            "% \\label{fig:placeholder}\n"
            "Text \\% literal \\label{fig:placeholder}\n"
            "More text \\label{fig:placeholder} % \\label{fig:placeholder}\n"
        )

        transformed, removed = _deduplicate_placeholder_labels(source)

        self.assertEqual(removed, 1)
        self.assertEqual(transformed.count("fig:placeholder"), 3)
        self.assertTrue(transformed.startswith("% \\label{fig:placeholder}"))


class ReferencePdfTests(unittest.TestCase):
    def test_spaced_solution_pdf_matches_compact_tex_name(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tex_path = root / "Quiz9.tex"
            question_pdf = root / "Quiz 9.pdf"
            solution_pdf = root / "Quiz 9 Solution.pdf"
            tex_path.touch()
            question_pdf.touch()
            solution_pdf.touch()

            matched = _find_reference_pdf(tex_path, r"\begin{solution}Answer\end{solution}")

            self.assertEqual(matched, solution_pdf)

    def test_question_pdf_wins_when_source_has_no_solutions(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tex_path = root / "Quiz9.tex"
            question_pdf = root / "Quiz 9.pdf"
            solution_pdf = root / "Quiz 9 Solution.pdf"
            tex_path.touch()
            question_pdf.touch()
            solution_pdf.touch()

            matched = _find_reference_pdf(tex_path, "Question only")

            self.assertEqual(matched, question_pdf)


class IntegralLimitTests(unittest.TestCase):
    def test_source_integral_limits_are_read_in_equation_order(self) -> None:
        source = (
            r"$\int_0^{l_z}\int_0^{l_y}\int_0^{l_x} dx\,dy\,dz$ "
            r"$\int_{-\infty}^{+\infty} f(x)\,dx$"
        )
        converter = LatexSourceDocxConverter(
            CompileResult(Path("main.pdf"), "", [], Path("build")),
            Path("source.tex"),
            source,
        )

        self.assertEqual(
            converter._source_integral_limits(),
            [("0", "l_z"), ("0", "l_y"), ("0", "l_x"), (r"-\infty", r"+\infty")],
        )


class PostprocessValidationTests(unittest.TestCase):
    def test_private_equation_markers_are_rejected(self) -> None:
        document = Document()
        paragraph = document.add_paragraph()
        formula = OxmlElement("m:oMath")
        formula.append(_math_run("PDFEq001"))
        paragraph._p.append(formula)

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "leaked.docx"
            document.save(path)
            with self.assertRaises(LatexWorkspaceError):
                LatexSourceDocxConverter._validate_postprocessed_docx(path)


if __name__ == "__main__":
    unittest.main()
