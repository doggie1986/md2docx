# -*- coding: utf-8 -*-
"""gui/math_html.py 的自检：一条 LaTeX 在预览里长什么样、体检念哪一句。

跑法（用项目内的 venv）::

    .venv\\Scripts\\python.exe -m unittest discover -s tests -v

这一层不需要 Qt，也不需要 Word：只吃 latex2mathml 交回的 MathML，所以命令行与
无界面环境都跑得完。这里钉的是**实测的形状**而不是"应该长什么样"：

* 行内分式折成 ``(a)/(b)``、块级分式摆成两行表格——因为 Qt 会把行内 ``<table>``
  提成块级、把那一行打断；
* 上标下标用 ``<sup>`` / ``<sub>``，帽子与向量用组合符号，根号写成 ``√( … )``
  （``text-decoration:overline`` 在这台机器上一点墨都不出）；
* 语法体检与渲染是**同一份判法**：坏公式在预览里退回原文并挂 ``math-broken``，
  校验念的是同一句原因。

`tests/test_gui_smoke.py` 里另有两条钉 :mod:`gui.theme.preview_css` 的格子线，
那一层要 Qt，所以不在这个文件里。
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import md_parser as mp  # noqa: E402
from gui import math_html  # noqa: E402

#: 体检与渲染共用的那一批样本：既有画得出来的，也有四种"认不出来"的写法。
#: 空串与纯空白是**唯一的例外**（见 EmptyFormulaTest），解析器根本产不出它们。
SAMPLES = [
    r"E=mc^2",
    r"x^2_i",
    r"a_1^2",
    r"\alpha",
    r"\frac{a}{b}",
    r"\frac{1}{2} + \frac{3}{4}",
    r"\sqrt{3}",
    r"\sqrt[3]{x}",
    r"\hat n",
    r"\vec{AB}",
    r"\overline{AB}",
    r"\underline{x}",
    r"\dot{x}",
    r"\tilde{x}",
    r"\bar{x}",
    r"\sum_{i=1}^{n} i",
    r"\int_0^1 x \mathrm{d}x",
    r"\prod_{k=1}^{n} k",
    r"\lim_{n \to \infty} a_n",
    r"\infty",
    r"\neq",
    r"\emptyset",
    r"\cdots",
    r"\binom{n}{k}",
    r"\Vert v\Vert",
    r"\left( x \right)",
    r"\mathbb{R}",
    r"\text{中文}",
    "a \\, b",
    r"\frac{\partial f}{\partial x}",
    r"\begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}",
    r"\begin{cases} a & b \end{cases}",
    r"\begin{aligned} x &= 1 \end{aligned}",
    "a & b",
    "a < b",
    "<script>",
    "a%b",
    "{a}",
    r"\{a\}",
    "a}",
    "",
    "   ",
    r"\frac{a}",
    r"{\frac{a}{b}",
    r"\unknownmacro{x}",
    "{}",
    "$",
    "a$",
    r"\sqrt",
    r"\begin{aligned} x",
    "a\\3",
    r"\hline",
]


class InlineShapeTest(unittest.TestCase):
    """行内：只能画在一行之内的东西。"""

    def test_sup_and_sub(self) -> None:
        self.assertEqual(math_html.render_math(r"E=mc^2"), "E=mc<sup>2</sup>")

    def test_sub_and_sup_together_puts_sub_first(self) -> None:
        # ``x^2_i`` 交回的是 <sub> 在前：MathML 把它收成 msubsup，一格里同时带上下标。
        self.assertEqual(math_html.render_math(r"x^2_i"), "x<sub>i</sub><sup>2</sup>")
        self.assertEqual(math_html.render_math(r"a_1^2"), "a<sub>1</sub><sup>2</sup>")

    def test_inline_fraction_folds_into_one_line(self) -> None:
        # 块级才摆两行表格；行内一旦出 <table> 就被 Qt 提成块级，那一句话被劈开。
        self.assertEqual(math_html.render_math(r"\frac{a}{b}"), "(a)/(b)")
        self.assertEqual(math_html.render_math(r"\frac{1}{2} + \frac{3}{4}"), "(1)/(2)+(3)/(4)")

    def test_sqrt_writes_as_radical_sign_with_parentheses(self) -> None:
        self.assertEqual(math_html.render_math(r"\sqrt{3}"), "√(3)")
        self.assertEqual(math_html.render_math(r"\sqrt[3]{x}"), "√<sup>3</sup>(x)")

    def test_accents_use_combining_marks(self) -> None:
        # 帽子/向量/上横线画得出来，靠的是 Unicode 组合符号而不是 text-decoration。
        self.assertEqual(math_html.render_math(r"\hat n"), "n\u0302")
        self.assertEqual(math_html.render_math(r"\vec{AB}"), "AB\u20d7")
        self.assertEqual(math_html.render_math(r"\overline{AB}"), "A\u0305B\u0305")
        self.assertEqual(math_html.render_math(r"\bar{x}"), "x\u0305")
        self.assertEqual(math_html.render_math(r"\dot{x}"), "x\u0307")
        self.assertEqual(math_html.render_math(r"\tilde{x}"), "x\u0303")

    def test_operators_and_greek_come_back_as_unicode(self) -> None:
        for source, expect in [(r"\alpha", "α"), (r"\infty", "∞"), (r"\neq", "≠"),
                               (r"\emptyset", "∅"), (r"\cdots", "⋯"),
                               (r"\mathbb{R}", "ℝ"), (r"\Vert", "‖")]:
            self.assertEqual(math_html.render_math(source), expect, source)

    def test_big_operators_keep_limits_as_sub_and_sup(self) -> None:
        self.assertEqual(math_html.render_math(r"\sum_{i=1}^{n} i"),
                         "∑<sub>i=1</sub><sup>n</sup>i")
        self.assertEqual(math_html.render_math(r"\prod_{k=1}^{n} k"),
                         "∏<sub>k=1</sub><sup>n</sup>k")
        self.assertEqual(math_html.render_math(r"\int_0^1 x \mathrm{d}x"),
                         "∫<sub>0</sub><sup>1</sup>xdx")
        self.assertEqual(math_html.render_math(r"\lim_{n \to \infty} a_n"),
                         "lim<sub>n→∞</sub>a<sub>n</sub>")

    def test_matrix_folds_into_one_line(self) -> None:
        self.assertEqual(
            math_html.render_math(r"\begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}"),
            "(1, 2 ; 3, 4)")

    def test_binom_is_a_fraction_in_parentheses(self) -> None:
        self.assertEqual(math_html.render_math(r"\binom{n}{k}"), "((n)/(k))")

    def test_left_right_and_text_and_spacing(self) -> None:
        self.assertEqual(math_html.render_math(r"\left( x \right)"), "(x)")
        self.assertEqual(math_html.render_math(r"\text{中文}"), "中文")
        self.assertEqual(math_html.render_math("a \\, b"), "a\u00a0b")
        self.assertEqual(math_html.render_math(r"\{a\}"), "{a}")

    def test_partial_derivative(self) -> None:
        self.assertEqual(math_html.render_math(r"\frac{\partial f}{\partial x}"),
                         "(∂f)/(∂x)")


class DisplayShapeTest(unittest.TestCase):
    """块级：分式与矩阵真的上下叠起来。"""

    def test_display_fraction_is_a_two_row_table(self) -> None:
        self.assertEqual(
            math_html.render_math(r"\frac{a}{b}", display=True),
            "<table class=\"math-frac\" cellspacing=\"0\" cellpadding=\"0\">"
            "<tr><td class=\"math-num\">a</td></tr>"
            "<tr><td class=\"math-den\">b</td></tr></table>")

    def test_display_matrix_is_a_table_in_parentheses(self) -> None:
        self.assertEqual(
            math_html.render_math(r"\begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}", display=True),
            "(<table class=\"math-matrix\" cellspacing=\"0\" cellpadding=\"0\">"
            "<tr><td>1</td><td>2</td></tr><tr><td>3</td><td>4</td></tr></table>)")

    def test_display_binom_wraps_a_table(self) -> None:
        self.assertIn("<table class=\"math-frac\"",
                      math_html.render_math(r"\binom{n}{k}", display=True))

    def test_display_cases_becomes_a_brace_and_a_table(self) -> None:
        html = math_html.render_math(r"\begin{cases} a & b \end{cases}", display=True)
        self.assertTrue(html.startswith("{"), html)
        self.assertIn("<table class=\"math-matrix\"", html)

    def test_what_does_not_need_a_grid_is_identical_in_both_modes(self) -> None:
        # 上下标与根号在两种模式下同一个形状——差别只出现在分式与矩阵上。
        for source in (r"E=mc^2", r"\sqrt{3}", r"\alpha", r"\sum_{i=1}^{n} i"):
            self.assertEqual(math_html.render_math(source),
                             math_html.render_math(source, display=True), source)


class LatexProblemTest(unittest.TestCase):
    """语法体检：四种判法各自那句原因。"""

    def test_good_writings_get_no_complaint(self) -> None:
        for source in (r"\alpha", "a \\, b", r"\frac{a}{b}", r"\sqrt{3}",
                       r"\begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}",
                       r"\begin{cases} a & b \end{cases}", "a < b", "a%b", "{a}",
                       r"\{a\}", r"\lim_{n \to \infty} a_n"):
            self.assertEqual(math_html.latex_problem(source), "", source)

    def test_empty_body(self) -> None:
        for source in ("", "   "):
            self.assertEqual(math_html.latex_problem(source),
                             "公式是空的（定界符之间没有内容）")

    def test_macro_missing_its_groups(self) -> None:
        self.assertEqual(math_html.latex_problem(r"\frac{a}"),
                         "\\frac 只跟了 1 组 {…}，它需要两组")

    def test_unbalanced_braces(self) -> None:
        self.assertEqual(math_html.latex_problem(r"{\frac{a}{b}"),
                         "括号不成对：有 1 个 { 没关上")
        self.assertEqual(math_html.latex_problem("a}"), "括号不成对：多出来 1 个 }")

    def test_unrecognised_macro(self) -> None:
        self.assertEqual(math_html.latex_problem(r"\unknownmacro{x}"),
                         "有一个宏或转义这一版认不出来：\\unknownmacro"
                         "（进了 Word 会原样显示成文本）")
        self.assertEqual(math_html.latex_problem(r"\hline"),
                         "有一个宏或转义这一版认不出来：\\hline（进了 Word 会原样显示成文本）")

    def test_bare_backslash_is_named_as_such(self) -> None:
        self.assertEqual(math_html.latex_problem("a\\3"),
                         "有一个宏或转义这一版认不出来：一个反斜杠（\\ 后面接不上宏名）"
                         "（进了 Word 会原样显示成文本）")

    def test_nothing_visible(self) -> None:
        self.assertEqual(math_html.latex_problem("{}"), "这段公式里没有能显示的内容")

    def test_stray_delimiter(self) -> None:
        for source in ("$", "a$"):
            self.assertEqual(math_html.latex_problem(source),
                             "公式正文里有一个多余的 $（多半是定界符少配了一个，回去看这一行的 $）")

    def test_converter_itself_gave_up(self) -> None:
        self.assertEqual(math_html.latex_problem(r"\sqrt"),
                         "latex2mathml 没能读懂这段 LaTeX（NoAvailableTokensError）")
        self.assertEqual(math_html.latex_problem(r"\begin{aligned} x"),
                         "latex2mathml 没能读懂这段 LaTeX（MissingEndError）")

    def test_bare_ampersand_names_the_character_and_the_fix(self) -> None:
        # latex2mathml 交回 <mi>&</mi>，那不是合法 XML（& 在 XML 里必须写成 &amp;）。
        # 这一句在本轮之前会让体检**自己抛穿**：抛穿的那一条由 test_nothing_raises 钉着。
        for source in ("a & b", r"\begin{aligned} x &= 1 \end{aligned}"):
            self.assertEqual(
                math_html.latex_problem(source),
                "公式正文里有一个没转义的 &（LaTeX 里它是表格对齐符，这一版转不成公式；"
                "要显示这个符号写成 \\&）", source)

    def test_percent_comment_truncates_without_complaint(self) -> None:
        # LaTeX 的注释语法：``a%b`` 读作 a。这不是坏公式，体检不该念一句。
        self.assertEqual(math_html.latex_problem("a%b"), "")
        self.assertEqual(math_html.render_math("a%b"), "a")


class BrokenFallbackTest(unittest.TestCase):
    """坏公式退回原文，并且**一定是转义过的**原文。"""

    def test_broken_gets_the_warning_class(self) -> None:
        self.assertEqual(math_html.render_math(r"\frac{a}"),
                         "<span class=\"math-broken\">\\frac{a}</span>")
        self.assertEqual(math_html.render_math("{}"),
                         "<span class=\"math-broken\">{}</span>")

    def test_escape_is_html_safe(self) -> None:
        # 退回的是文本，可它要进 QTextBrowser 的富文本里：不转义就等于让公式里
        # 那段尖括号在预览区里开一个真标签。
        self.assertEqual(math_html.render_math("<script>"), "&lt;script&gt;")
        self.assertEqual(math_html.render_math("a < b"), "a&lt;b")
        self.assertEqual(math_html.render_math("a & b"),
                         "<span class=\"math-broken\">a &amp; b</span>")
        for source in SAMPLES:
            for display in (False, True):
                html = math_html.render_math(source, display=display)
                self.assertNotIn("<script>", html, "%r display=%r" % (source, display))
                self.assertNotIn("<b>", html, "%r display=%r" % (source, display))

    def test_empty_source_renders_nothing(self) -> None:
        # 空公式没有"原文"可退回；解析器产不出它（见 SameJudgementTest 那条），
        # 所以这一格是体检单方面念得出、预览无处画的那一种。
        self.assertEqual(math_html.render_math(""), "")
        self.assertEqual(math_html.render_math("   ", display=True), "")


class SameJudgementTest(unittest.TestCase):
    """预览与校验必须念同一句：坏公式在预览里挂 math-broken，在校验里有原因。"""

    def test_broken_span_and_problem_agree(self) -> None:
        for source in SAMPLES:
            if not source.strip():
                continue
            for display in (False, True):
                broken = "math-broken" in math_html.render_math(source, display=display)
                complained = bool(math_html.latex_problem(source))
                self.assertEqual(broken, complained,
                                 "%r display=%r broken=%r problem=%r"
                                 % (source, display, broken, complained))

    def test_nothing_raises(self) -> None:
        # 这一条是本轮那处缺陷的回归：一个裸的 & 会让 lxml 抛 XMLSyntaxError，
        # 而异常从 latex_problem 逃出去就把整轮校验带崩、把预览那一次渲染整个废掉。
        for source in SAMPLES + [None, "\\begin", "%", "&", "\\end{array}", "{}{}"]:
            for display in (False, True):
                self.assertIsInstance(math_html.render_math(source, display=display), str)
            self.assertIsInstance(math_html.latex_problem(source), str)

    def test_parser_never_feeds_an_empty_formula(self) -> None:
        # 上面那条豁免的理由在这里，而不是"两边不一致也没关系"：
        # 定界符中间空着的时候，解析器根本不产出数学节点（实测三例都是 0 条）。
        for text in ("$ $", "$$", "$ $$"):
            parsed = mp.parse(text)
            self.assertEqual([], list(parsed.annotated(mp.MathRun, mp.BlockMath)), text)
        # 同一行里"有内容的"与"空着的"并存时，只有前者成为节点，且它必然非空。
        nodes = list(mp.parse("$\\alpha$ $ $").annotated(mp.MathRun, mp.BlockMath))
        self.assertEqual(1, len(nodes))
        self.assertEqual(r"\alpha", nodes[0][0].latex)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
