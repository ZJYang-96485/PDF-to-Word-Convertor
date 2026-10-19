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
