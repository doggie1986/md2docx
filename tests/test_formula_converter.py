# -*- coding: utf-8 -*-
"""formula_converter 的自检。

跑法（用项目内的 venv）::

    .venv\\Scripts\\python.exe -m unittest discover -s tests -v

这台机器装了 Office，所以定位样式表那几条是真跑；换一台没装 Office 的机器，
``locate_xslt()`` 会抛 XsltNotFoundError，这些用例会整类跳过而不是失败。
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from lxml import etree  # noqa: E402

import formula_converter as fc  # noqa: E402

try:
    XSLT = fc.locate_xslt()
except fc.XsltNotFoundError:  # pragma: no cover - 取决于机器上有没有 Office
    XSLT = None


@unittest.skipIf(XSLT is None, "本机找不到 MML2OMML.XSL，公式用例无法真跑")
class ConvertTest(unittest.TestCase):
    def setUp(self) -> None:
        self.converter = fc.FormulaConverter(XSLT)

    # --- 结构：行内是 m:oMath，块级外面多一层 m:oMathPara -------------------
    def test_inline_returns_omath(self) -> None:
        result = self.converter.convert(r"\frac{a}{b}")
        self.assertTrue(result.ok, result.problem)
        self.assertEqual(result.omml.tag, fc.M + "oMath")

    def test_display_wraps_omathpara(self) -> None:
        result = self.converter.convert(r"\frac{a}{b}", display=True)
        self.assertEqual(result.omml.tag, fc.M + "oMathPara")
        self.assertEqual(result.omml[0].tag, fc.M + "oMath")

    # --- 落笔前体检的两处声明 ----------------------------------------------
    def test_every_math_run_declares_formula_font(self) -> None:
        result = self.converter.convert(r"\sum_{i=1}^{n} i^2 + \sqrt{x}")
        runs = result.omml.findall(".//" + fc.M + "r")
        self.assertTrue(runs, "转换结果里一个 m:r 都没有")
        for run in runs:
            fonts = run.find(fc.W + "rPr/" + fc.W + "rFonts")
            self.assertIsNotNone(fonts, "有 m:r 没带 w:rFonts，Word 会用正文字体画它")
            self.assertEqual(fonts.get(fc.W + "ascii"), "Cambria Math")

    def test_m_space_is_preserved(self) -> None:
        result = self.converter.convert(r"\beta \alpha")
        texts = result.omml.findall(".//" + fc.M + "t")
        self.assertTrue(texts)
        for text in texts:
            self.assertEqual(text.get("{http://www.w3.org/XML/1998/namespace}space"), "preserve")

    def test_font_is_configurable(self) -> None:
        converter = fc.FormulaConverter(XSLT, formula_font="Times New Roman")
        result = converter.convert(r"x^2")
        fonts = result.omml.find(".//" + fc.W + "rFonts")
        self.assertEqual(fonts.get(fc.W + "ascii"), "Times New Roman")

    # --- 各种真实写法 --------------------------------------------------------
    def test_matrix_and_absolute_value_survive(self) -> None:
        result = self.converter.convert(
            r"\left|\frac{1}{2}\right| \begin{pmatrix} a & b \\ c & d \end{pmatrix}"
        )
        self.assertTrue(result.ok, result.problem)
        self.assertGreater(self.converter.count_text_nodes(result.omml), 4)

    def test_conversion_is_deterministic(self) -> None:
        """同一段 LaTeX 两次转换必须逐字节相同，否则对账时说不清哪一次是对的。"""
        latex = r"E = mc^2 \cdot \int_0^1 x\,dx"
        first = self.converter.convert(latex, display=True)
        second = self.converter.convert(latex, display=True)
        self.assertEqual(
            etree.tostring(first.omml), etree.tostring(second.omml)
        )

    # --- 失败要交回原因，不能抛也不能产出半个框 ----------------------------
    def test_empty_formula_is_reported_not_raised(self) -> None:
        result = self.converter.convert("   ")
        self.assertFalse(result.ok)
        self.assertIn("空的", result.problem)

    def test_broken_formula_reports_chinese_reason(self) -> None:
        result = self.converter.convert(r"\frac{1}{")
        self.assertFalse(result.ok)
        self.assertTrue(result.problem, "失败却没给原因")
        self.assertIn(r"\frac{1}{", result.failure_report())

    def test_strict_helper_raises(self) -> None:
        with self.assertRaises(fc.FormulaConversionError):
            self.converter.convert_or_raise(r"\frac{1}{")


@unittest.skipIf(XSLT is None, "本机找不到 MML2OMML.XSL")
class LocateTest(unittest.TestCase):
    def test_default_search_finds_office(self) -> None:
        self.assertTrue(XSLT.is_file())
        self.assertEqual(XSLT.name.upper(), "MML2OMML.XSL")

    def test_explicit_bad_path_lists_that_path(self) -> None:
        """用户指了路径却不存在时，报错要念他给的那一条，而不是本机默认路径清单。"""
        with self.assertRaises(fc.XsltNotFoundError) as caught:
            fc.locate_xslt(r"D:\不存在\MML2OMML.XSL")
        message = str(caught.exception)
        self.assertIn(r"D:\不存在", message)
        self.assertIn("--xslt-path", message)

    def test_not_found_message_lists_every_searched_path(self) -> None:
        """找不到样式表时，报错要把找过的位置全念出来，只说"找不到"等于没报。"""
        original = fc.default_search_dirs
        fc.default_search_dirs = lambda: [Path(r"D:\甲目录"), Path(r"D:\乙目录")]
        try:
            with self.assertRaises(fc.XsltNotFoundError) as caught:
                fc.locate_xslt(None)
        finally:
            fc.default_search_dirs = original
        message = str(caught.exception)
        self.assertIn(r"D:\甲目录", message)
        self.assertIn(r"D:\乙目录", message)
        self.assertIn("MML2OMML.XSL", message)

    def test_directory_also_accepted(self) -> None:
        self.assertEqual(fc.locate_xslt(str(XSLT.parent)), XSLT)


if __name__ == "__main__":
    unittest.main()
