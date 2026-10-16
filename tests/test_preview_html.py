# -*- coding: utf-8 -*-
"""gui/preview_html.py 的自检：预览区那一整块 HTML 是怎么拼出来的。

跑法（用项目内的 venv）::

    .venv\\Scripts\\python.exe -m unittest discover -s tests -v

这一层不碰 Qt，交回的就是一段纯文本 HTML，所以命令行也能挂上去。钉的四件事：

* **形状**：每一种块画成哪一种标签（标题 ``h1~h6``、分式在块级是 ``table``、代码块带
  ``data-lang``、分隔线是 ``hr``）——预览少画一种块就会变成"解析器认下了、屏幕上没有"；
* **不认识的块要念出来**：交回类名而不是悄悄丢掉，丢掉的那一段没人知道去哪儿查；
* **一律转义**：正文里的 ``<`` 与 ``&`` 不许变成标签，坏公式退回原文时也不许把标签带进预览区；
* **读不动只交回一句话**：``render_file`` 不抛——预览区坏了不该把主窗口带崩。

公式本身画成什么样由 :mod:`tests.test_math_html` 钉，这里只钉"套在外层的那壳"。
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from html import escape
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import md_parser as mp  # noqa: E402
from gui import preview_html as ph  # noqa: E402

B = chr(92)  # 一个反斜杠：写成 B + "frac" 比 "\\\\" 好读，也不会被 shell 吃掉


def html_of(*blocks) -> str:
    return ph.render_blocks(list(blocks))


def para(*runs) -> mp.Paragraph:
    return mp.Paragraph(list(runs))


def table(rows, aligns=(), **kwargs) -> mp.Table:
    """搭一张表。每格写字符串就是一段纯文本，写一串 run 就是"表格里嵌公式"那一类。

    ``md_parser.Table.rows`` 的形状是 ``list[list[list[Inline]]]``——一格里装的是行内
    片段列表，直接把 ``TextRun`` 当格子交进去会在渲染层炸成 TypeError。
    """
    built = [[[mp.TextRun(cell)] if isinstance(cell, str) else list(cell) for cell in row]
             for row in rows]
    return mp.Table(built, list(aligns), **kwargs)


class NoticeTest(unittest.TestCase):
    """预览顶部那一句"这里和 Word 里差在哪"。"""

    def setUp(self) -> None:
        self._dir = tempfile.TemporaryDirectory()
        self.base = Path(self._dir.name)
        self.addCleanup(self._dir.cleanup)

    def write(self, name: str, content: str) -> Path:
        path = self.base / name
        path.write_text(content, encoding="utf-8")
        return path

    def test_notice_says_the_three_differences(self) -> None:
        # 这几处是量出来的差别，不是含糊过去的说法：行内分式折成一行、认不出的写法退回
        # 原文、图片要进了 Word 才真嵌进去、字号分页归转换那一步。文案改了要连着改这一句，
        # 别让它变成假承诺。
        for phrase in ["近似预览", "行内公式里的分式与矩阵只能折成一行",
                       "退回 LaTeX 原文",
                       "图片在这里只写名字，进了 Word 才嵌成真的图",
                       "字号与分页以转换时的样式设置为准"]:
            self.assertIn(phrase, ph.NOTICE)

    def test_render_file_wraps_notice_then_name_then_path(self) -> None:
        path = self.write("doc.md", "# 标题\n\n正文。\n")
        result = ph.render_file(path)
        self.assertTrue(result.ok, result.problem)
        self.assertTrue(result.html.startswith("<div class=\"notice\">"), result.html[:60])
        # 文件名那一行不带内联 style：样式表里 h3 本来就是 margin: 10px 0 4px，多写一句
        # 反而把整份声明顶掉了（连"标题用黑体"都没了）——理由写在 gui/preview_html.py 的 _ALIGN_WORDS 上。
        self.assertIn("<h3>doc.md</h3>", result.html)
        self.assertIn("<p class=\"notice\">%s</p>" % escape(str(path)), result.html)
        self.assertIn("<h1>标题</h1>", result.html)
        self.assertIn("<p>正文。</p>", result.html)

    def test_the_body_comes_after_the_wrapper(self) -> None:
        # 正文排在最后：预览区是一整块连续文本，文件名与说明压在最上面才读得对次序。
        path = self.write("order.md", "第一行。\n\n第二行。\n")
        html = ph.render_file(path).html
        self.assertLess(html.index("<h3"), html.index("<p>第一行。</p>"))
        self.assertLess(html.index("<p>第一行。</p>"), html.index("<p>第二行。</p>"))

    def test_render_blocks_has_no_wrapper(self) -> None:
        # 没有文件也能画（"先看一眼"那条路和测试都走这一条），所以顶部那一叠不该出现。
        html = html_of(mp.Heading(1, [mp.TextRun("只有正文")]))
        self.assertEqual("<h1>只有正文</h1>", html)
        self.assertNotIn("notice", html)

    def test_notice_is_escaped_like_the_rest(self) -> None:
        path = self.write("esc.md", "正文\n")
        result = ph.render_file(path)
        self.assertIn(escape(ph.NOTICE), result.html)


class HeadingBlockTest(unittest.TestCase):
    def test_each_level_maps_to_its_own_tag(self) -> None:
        for level in range(1, 7):
            self.assertEqual("<h%d>题</h%d>" % (level, level),
                             html_of(mp.Heading(level, [mp.TextRun("题")])))

    def test_levels_outside_one_to_six_are_clamped(self) -> None:
        # Qt 富文本只认 h1~h6：写个 h9 进去就画不出标题，宁可收到边界上也不能丢这一段。
        self.assertEqual("<h6>题</h6>", html_of(mp.Heading(9, [mp.TextRun("题")])))
        self.assertEqual("<h1>题</h1>", html_of(mp.Heading(0, [mp.TextRun("题")])))
        self.assertEqual("<h1>题</h1>", html_of(mp.Heading(-3, [mp.TextRun("题")])))

    def test_inline_inside_a_heading_is_still_rendered(self) -> None:
        html = html_of(mp.Heading(2, [mp.TextRun("见 "), mp.CodeSpan("config"),
                                      mp.TextRun("，"), mp.MathRun(B + "alpha")]))
        self.assertEqual("<h2>见 <code>config</code>，<span class=\"math\">α</span></h2>", html)


class ParagraphInlineTest(unittest.TestCase):
    def test_bold_and_italic_nest_in_the_written_order(self) -> None:
        self.assertEqual("<p><strong>粗</strong></p>", html_of(para(mp.TextRun("粗", bold=True))))
        self.assertEqual("<p><em>斜</em></p>", html_of(para(mp.TextRun("斜", italic=True))))
        self.assertEqual("<p><em><strong>又粗又斜</strong></em></p>",
                         html_of(para(mp.TextRun("又粗又斜", bold=True, italic=True))))

    def test_link_wraps_the_already_styled_text(self) -> None:
        html = html_of(para(mp.TextRun("点这里", link="https://x.test/a")))
        self.assertEqual("<p><a href=\"https://x.test/a\">点这里</a></p>", html)

    def test_plain_paragraph(self) -> None:
        self.assertEqual("<p>就是一句话。</p>", html_of(para(mp.TextRun("就是一句话。"))))

    def test_empty_paragraph_stays_a_pair_of_tags(self) -> None:
        # 空段落也要留一个 <p>：整段消失会让人以为预览漏抽了一行。
        self.assertEqual("<p></p>", html_of(para()))

    def test_code_span_is_escaped(self) -> None:
        self.assertEqual("<p><code>&lt;tag&gt;</code></p>", html_of(para(mp.CodeSpan("<tag>"))))


class ImagePlaceholderTest(unittest.TestCase):
    """图片在预览里只写一句"这里有张什么图"——这一层不读盘，也不嵌图（嵌是 Word 那一步的事）。

    那句文案**不能**照抄 Word 的占位：那一份说的是"没嵌进去"，在这里就是假话。
    """

    def test_the_placeholder_uses_the_alt_first(self) -> None:
        html = html_of(para(mp.ImageRef("施工图", "a/b/清汤.png")))
        self.assertEqual("<p><span class=\"placeholder\">%s</span></p>"
                         % escape(ph.IMAGE_MARK % "施工图"), html)
        self.assertNotIn("没嵌进去", html)

    def test_without_alt_it_falls_back_to_the_url(self) -> None:
        html = html_of(para(mp.ImageRef("", "a/b/清汤.png")))
        self.assertIn("a/b/清汤.png", html)
        self.assertIn("placeholder", html)

    def test_the_name_is_escaped(self) -> None:
        # 占位文案里带 < 也不能变成标签：那一格是文件里读出来的任意文本，不算可信输入。
        html = html_of(para(mp.ImageRef("<b>图</b>", "x.png")))
        self.assertIn("&lt;b&gt;图&lt;/b&gt;", html)
        self.assertNotIn("<b>", html)


class CodeBlockTest(unittest.TestCase):
    def test_language_becomes_a_data_attribute(self) -> None:
        self.assertEqual("<pre data-lang=\"python\"><code>print(1)</code></pre>",
                         html_of(mp.CodeBlock("print(1)", "python")))

    def test_no_language_means_no_attribute(self) -> None:
        self.assertEqual("<pre><code>plain</code></pre>", html_of(mp.CodeBlock("plain")))

    def test_the_code_is_escaped_not_executed(self) -> None:
        html = html_of(mp.CodeBlock("<script>alert(1)</script>", "html"))
        self.assertIn("&lt;script&gt;", html)
        self.assertNotIn("<script>", html)

    def test_newlines_survive(self) -> None:
        # 代码块要能读出"这是几行"，把换行压成空格就等于把代码画成一句话。
        self.assertIn("a\nb", html_of(mp.CodeBlock("a\nb", "")))


class MathBlockTest(unittest.TestCase):
    """外层那一壳：行内是 span、块级多一层 div.math-block。"""

    def setUp(self) -> None:
        self._dir = tempfile.TemporaryDirectory()
        self.addCleanup(self._dir.cleanup)

    def write(self, content: str, name: str = "m.md") -> Path:
        path = Path(self._dir.name) / name
        path.write_text(content, encoding="utf-8")
        return path

    def test_inline_math_wraps_in_a_span(self) -> None:
        html = html_of(para(mp.TextRun("值 "), mp.MathRun(B + "frac{a}{b}")))
        self.assertIn("<span class=\"math\">", html)
        self.assertIn("(a)/(b)", html)  # 行内折成一行，NOTICE 里已经说明过这一点

    def test_block_math_wraps_in_a_div_and_uses_the_table_shape(self) -> None:
        html = html_of(mp.BlockMath(B + "frac{a}{b}"))
        self.assertTrue(html.startswith("<div class=\"math math-block\">"), html)
        self.assertIn("<table class=\"math-frac\"", html)
        self.assertTrue(html.endswith("</div>"), html)

    def test_a_broken_formula_falls_back_inside_the_paragraph(self) -> None:
        # 体检没过 → 原样退回 LaTeX 并挂 math-broken，而不是画成一个看不见的空格。
        html = html_of(para(mp.TextRun("这里 "), mp.MathRun(B + "frac{a}")))
        self.assertIn("math-broken", html)
        self.assertIn(B + "frac{a}", html)

    def test_a_bare_ampersand_does_not_abort_the_preview(self) -> None:
        # 这一条是本轮修掉的那处崩溃的回归：裸 & 让 lxml 抛 XMLSyntaxError，而主窗口
        # 调 render_file 那一行没包 try，一条坏公式就把整份预览带走（屏幕上一片空白）。
        result = ph.render_file(self.write("$a & b$ 与 $" + B + "frac{a}$\n"))
        self.assertTrue(result.ok, result.problem)
        self.assertIn("math-broken", result.html)
        self.assertIn("a &amp; b", result.html)


class TableTest(unittest.TestCase):
    def test_first_row_is_header_when_one_is_present(self) -> None:
        self.assertEqual("<table><tr><th>a</th><th>b</th></tr><tr><td>1</td><td>2</td></tr></table>",
                         html_of(table([["a", "b"], ["1", "2"]], ["none", "none"])))

    def test_no_header_row_means_every_cell_is_td(self) -> None:
        self.assertEqual("<table><tr><td>1</td><td>2</td></tr></table>",
                         html_of(table([["1", "2"]], ["none", "none"], header_present=False)))

    def test_alignment_words_become_align_attributes(self) -> None:
        """对齐走 ``align`` 属性而不是 ``style="text-align:…"``：后者在 Qt 里把整格样式声明顶掉，
        一张带对齐的表就**一条格线都没有**（真开窗量的：345 个格线色像素 → 0）。

        ``justify`` 不是 ``align`` 认的值、Markdown 也问不出它，所以和认不出的词一样按默认摆。
        """
        html = html_of(table([["h"] * 4, ["l", "c", "r", "j"]],
                             ["left", "center", "right", "justify"]))
        self.assertIn("<td align=\"left\">l</td>", html)
        self.assertIn("<td align=\"center\">c</td>", html)
        self.assertIn("<td align=\"right\">r</td>", html)
        self.assertIn("<td>j</td>", html)
        self.assertNotIn("style=", html)

    def test_alignment_is_case_and_space_tolerant(self) -> None:
        self.assertIn("<td align=\"center\">v</td>",
                      html_of(table([["h"], ["v"]], [" CENTER "])))

    def test_unknown_alignment_is_dropped_not_invented(self) -> None:
        # 认不出的写法不硬编一个样式：宁可按默认摆，也不猜人想干什么。
        self.assertEqual("<table><tr><th>h</th></tr><tr><td>v</td></tr></table>",
                         html_of(table([["h"], ["v"]], ["middle"])))

    def test_missing_align_entry_leaves_the_cell_plain(self) -> None:
        # 行比 aligns 长（解析器补齐过列）时不许 IndexError，也不许凭空给一个对齐。
        html = html_of(table([["h", "h2"], ["v", "v2"]], ["center"]))
        self.assertIn("<th align=\"center\">h</th>", html)
        self.assertIn("<td>v2</td>", html)

    def test_inline_content_inside_cells_is_rendered(self) -> None:
        self.assertEqual("<table><tr><th><span class=\"math\">α</span></th></tr></table>",
                         html_of(table([[[mp.MathRun(B + "alpha")]],], ["none"])))

    def test_a_cell_can_hold_several_inline_runs(self) -> None:
        # 一格里"文字 + 行内公式 + 代码"是这份解析器的常态写法，预览必须按次序接起来。
        cell = [mp.TextRun("均值为 "), mp.MathRun(B + "mu"), mp.TextRun(" 伏"), mp.CodeSpan("U0")]
        self.assertEqual("<table><tr><th>均值为 <span class=\"math\">μ</span> 伏<code>U0</code>"
                         "</th></tr></table>", html_of(table([[cell]], ["none"])))

    def test_a_cell_can_hold_an_image_placeholder(self) -> None:
        # 校验那四项里"表格单元格里的图片路径"是最容易漏的一格，预览同样要画到。
        html = html_of(table([[[mp.ImageRef("截面", "nope.png")]],
                              [[mp.ImageRef("", "nope.png")]]], ["none"]))
        self.assertEqual("<table><tr><th><span class=\"placeholder\">%s</span></th></tr>"
                         "<tr><td><span class=\"placeholder\">%s</span></td></tr></table>"
                         % (escape(ph.IMAGE_MARK % "截面"),
                            escape(ph.IMAGE_MARK % "nope.png")), html)


class ListTest(unittest.TestCase):
    def test_unordered_and_ordered_tags(self) -> None:
        unordered = mp.ListBlock(False, [mp.ListItem([mp.TextRun("one")]),
                                         mp.ListItem([mp.TextRun("two")])])
        self.assertEqual("<ul><li>one</li><li>two</li></ul>", html_of(unordered))
        ordered = mp.ListBlock(True, [mp.ListItem([mp.TextRun("first")])])
        self.assertEqual("<ol><li>first</li></ol>", html_of(ordered))

    def test_start_attribute_only_when_it_is_not_the_default(self) -> None:
        items = [mp.ListItem([mp.TextRun("five")])]
        self.assertEqual("<ol start=\"5\"><li>five</li></ol>",
                         html_of(mp.ListBlock(True, items, start=5)))
        for start in (0, 1):
            self.assertEqual("<ol><li>five</li></ol>",
                             html_of(mp.ListBlock(True, items, start=start)))

    def test_nested_children_use_ul_when_the_parent_item_is_not_ordered(self) -> None:
        listing = mp.ListBlock(False, [
            mp.ListItem([mp.TextRun("one")]),
            mp.ListItem([mp.TextRun("two")], [mp.ListItem([mp.TextRun("sub")])]),
        ])
        self.assertEqual("<ul><li>one</li><li>two<ul><li>sub</li></ul></li></ul>", html_of(listing))

    def test_the_nested_tag_comes_from_the_children_not_the_parent_row(self) -> None:
        # ``1. a`` 底下挂一个圆点子表：解析器给孩子的 ordered 是 False（它们那一层的记号），
        # 父行那个 True 说的是"我这一层是序号"。拿父行的记号当子层的记号，预览就会把
        # 圆点演成 1. ——而 Word 那一条用的是 child.ordered（docx_builder._write_list_item）。
        listing = mp.ListBlock(True, [
            mp.ListItem([mp.TextRun("a")], [mp.ListItem([mp.TextRun("i")])], True),
            mp.ListItem([mp.TextRun("b")],
                        [mp.ListItem([mp.TextRun("j")], ordered=True),
                         mp.ListItem([mp.TextRun("k")], ordered=True)], True),
        ])
        self.assertEqual("<ol><li>a<ul><li>i</li></ul></li>"
                         "<li>b<ol><li>j</li><li>k</li></ol></li></ol>", html_of(listing))

    def test_three_levels_deep_still_nests(self) -> None:
        # 只走一层的话第三层就凭空消失：解析器认下了、Word 写出来了，预览却没有——
        # 这正是本模块开头不许的那种错，所以钉整串而不是钉"里面有 g"。
        listing = mp.ListBlock(False, [mp.ListItem([mp.TextRun("p")],
                                                   [mp.ListItem([mp.TextRun("c")],
                                                                [mp.ListItem([mp.TextRun("g")])])])])
        self.assertEqual("<ul><li>p<ul><li>c<ul><li>g</li></ul></li></ul></li></ul>",
                         html_of(listing))

    def test_a_nested_list_numbers_from_one(self) -> None:
        # 子层的起始号总是 1（与 Word 同一口径），所以 start 那一格不该被带进嵌套层：
        # 把外层的 start="5" 抄到子表上，屏幕上念的序号就不是这份稿子里的那几个数了。
        child = mp.ListItem([mp.TextRun("i")], ordered=True)
        listing = mp.ListBlock(True, [mp.ListItem([mp.TextRun("a")], [child], True)], start=5)
        self.assertEqual("<ol start=\"5\"><li>a<ol><li>i</li></ol></li></ol>", html_of(listing))

    def test_inline_inside_an_item_is_rendered(self) -> None:
        listing = mp.ListBlock(False, [mp.ListItem([mp.TextRun("x"), mp.CodeSpan("y")])])
        self.assertEqual("<ul><li>x<code>y</code></li></ul>", html_of(listing))


class NestedListFromRealMarkdownTest(unittest.TestCase):
    """嵌套列表按真写法过一遍：解析器吃进去几层，预览就得画出几层。

    手搭的节点只能证明渲染函数认这个形状，认不出"解析器到底会不会交回三层"——
    上一版的缺陷恰好长在这一格上（写测试的人以为只有两层，于是预览少了也没人喊）。
    """

    def html(self, text: str) -> str:
        return ph.render_blocks(mp.parse(text).blocks)

    def test_three_indent_levels_all_appear(self) -> None:
        html = self.html("- 主干\n  - 分支\n    - 细节\n")
        self.assertEqual("<ul><li>主干<ul><li>分支<ul><li>细节</li></ul></li></ul></li></ul>", html)

    def test_a_bullet_under_a_numbered_item_stays_a_bullet(self) -> None:
        self.assertEqual("<ol><li>第一步<ul><li>注意</li></ul></li></ol>",
                         self.html("1. 第一步\n   - 注意\n"))

    def test_several_children_of_one_item_are_all_rendered(self) -> None:
        html = self.html("- 父\n  - 甲\n    - 甲一\n  - 乙\n    - 乙一\n")
        for name in ["父", "甲", "乙", "甲一", "乙一"]:
            self.assertIn(name, html)
        self.assertEqual(html.count("<li>"), 5)

    def test_nesting_inside_a_blockquote_survives(self) -> None:
        # 引用块走的是 _blocks → _block → _list 这一整条，深度不该在这一层被削掉。
        self.assertIn("<blockquote><ul><li>外<ul><li>里</li></ul></li></ul></blockquote>",
                      self.html("> - 外\n>   - 里\n"))


class QuoteTest(unittest.TestCase):
    def test_quote_wraps_whole_blocks(self) -> None:
        quote = mp.QuoteBlock([mp.Paragraph([mp.TextRun("quoted")])])
        self.assertEqual("<blockquote><p>quoted</p></blockquote>", html_of(quote))

    def test_quote_can_hold_a_heading_and_a_list(self) -> None:
        quote = mp.QuoteBlock([
            mp.Heading(3, [mp.TextRun("小标")]),
            mp.ListBlock(False, [mp.ListItem([mp.TextRun("a")])]),
        ])
        self.assertEqual("<blockquote><h3>小标</h3><ul><li>a</li></ul></blockquote>", html_of(quote))

    def test_several_quotes_are_separate_boxes(self) -> None:
        self.assertEqual("<blockquote><p>一</p></blockquote>"
                         "<blockquote><p>二</p></blockquote>",
                         html_of(mp.QuoteBlock([para(mp.TextRun("一"))]),
                                 mp.QuoteBlock([para(mp.TextRun("二"))])))


class ThematicBreakTest(unittest.TestCase):
    def test_it_becomes_an_hr(self) -> None:
        self.assertEqual("<hr>", html_of(mp.ThematicBreak()))

    def test_three_breaks_are_three_lines(self) -> None:
        self.assertEqual("<hr><hr><hr>", html_of(mp.ThematicBreak(), mp.ThematicBreak(),
                                                 mp.ThematicBreak()))


class UnknownBlockTest(unittest.TestCase):
    def test_an_unrecognized_block_names_its_class(self) -> None:
        # 解析器认下的块如果在预览里被悄悄丢掉，就变成"预览少了一段"而没人知道去哪查。
        class 某种新块:  # noqa: N801 - 这个类名就是要出现在交回的那一句里
            pass

        self.assertEqual("<p class=\"placeholder\">（预览还不认识这一种块：某种新块）</p>",
                         html_of(某种新块()))

    def test_a_string_block_does_not_raise(self) -> None:
        self.assertIn("预览还不认识这一种块：str", html_of("not a block"))


class EscapingTest(unittest.TestCase):
    def test_angle_brackets_in_text_never_become_tags(self) -> None:
        self.assertEqual("<p>a &lt; b &gt; c</p>", html_of(para(mp.TextRun("a < b > c"))))

    def test_ampersand_in_text_is_escaped(self) -> None:
        self.assertEqual("<p>A &amp; B</p>", html_of(para(mp.TextRun("A & B"))))

    def test_link_url_is_escaped_for_the_attribute(self) -> None:
        # href 里的 & 与引号必须转义：一个引号就能闭合属性、把标签写进预览区。
        html = html_of(para(mp.TextRun("x", link='http://t.test/?a=1&b="2"')))
        self.assertEqual("<p><a href=\"http://t.test/?a=1&amp;b=&quot;2&quot;\">x</a></p>", html)

    def test_bold_text_is_escaped_before_it_is_wrapped(self) -> None:
        self.assertEqual("<p><strong>&lt;b&gt;</strong></p>",
                         html_of(para(mp.TextRun("<b>", bold=True))))

    def test_no_sample_lets_a_script_tag_through(self) -> None:
        blocks = [
            para(mp.TextRun("<script>")),
            para(mp.CodeSpan("<script>")),
            para(mp.ImageRef("<script>", "x.png")),
            para(mp.TextRun("x", link="javascript:<script>")),
            mp.CodeBlock("<script>"),
            mp.Heading(1, [mp.TextRun("<script>")]),
            mp.BlockMath("<script>"),
        ]
        html = html_of(*blocks)
        self.assertNotIn("<script>", html)


class StatsAndWarningsPassThroughTest(unittest.TestCase):
    def setUp(self) -> None:
        self._dir = tempfile.TemporaryDirectory()
        self.base = Path(self._dir.name)
        self.addCleanup(self._dir.cleanup)

    def write(self, content: str, name: str = "s.md") -> Path:
        path = self.base / name
        path.write_text(content, encoding="utf-8")
        return path

    def test_stats_are_the_parsers_own_numbers(self) -> None:
        # 预览不自己数"有几条公式"：数两遍就会出现"预览说 3 条、日志说 2 条"。
        path = self.write("# 标题\n\n$\\alpha$ 与\n\n$$E=mc^2$$\n\n"
                         "| a | b |\n|---|---|\n| 1 | 2 |\n")
        result = ph.render_file(path)
        self.assertEqual(dict(mp.parse_file(path).stats), dict(result.stats))
        self.assertEqual(1, result.stats["tables"])
        self.assertEqual(1, result.stats["headings"])

    def test_stats_keys_are_the_ones_the_log_line_reads(self) -> None:
        result = ph.render_file(self.write("# t\n\n正文 $\\alpha$\n", "keys.md"))
        for key in ["blocks", "tables", "formulas_inline", "formulas_display", "headings",
                    "code_blocks", "links", "images", "encoding"]:
            self.assertIn(key, result.stats, key)

    def test_warnings_are_the_parsers_list(self) -> None:
        path = self.write("| a | b |\n|---|---|\n| 1 |\n", "w.md")
        self.assertEqual(list(mp.parse_file(path).warnings), list(ph.render_file(path).warnings))

    def test_html_of_a_parsed_document_has_every_block(self) -> None:
        parsed = mp.parse("# 标题\n\n正文 $\\alpha$\n\n$$a+b$$\n\n"
                          "| a |\n|---|\n| 1 |\n\n- one\n\n> 引\n\n```py\nx=1\n```\n\n---\n")
        html = ph.render_blocks(parsed.blocks)
        for piece in ["<h1>标题</h1>", "<span class=\"math\">α</span>",
                      "<div class=\"math math-block\">", "<table>", "<ul><li>one</li></ul>",
                      "<blockquote>", "<pre data-lang=\"py\">", "<hr>"]:
            self.assertIn(piece, html, piece)


class RenderFileProblemTest(unittest.TestCase):
    """读不动那两种情况：交回一句话，不抛，也不留下半成品 HTML。"""

    def setUp(self) -> None:
        self._dir = tempfile.TemporaryDirectory()
        self.base = Path(self._dir.name)
        self.addCleanup(self._dir.cleanup)

    def test_missing_file(self) -> None:
        result = ph.render_file(self.base / "nope.md")
        self.assertFalse(result.ok)
        self.assertTrue(result.problem.startswith("文件已经不在了："), result.problem)
        self.assertIn("nope.md", result.problem)
        self.assertEqual("", result.html)
        self.assertEqual({}, result.stats)

    def test_a_directory_is_not_a_file(self) -> None:
        # is_file() 那一关顺便挡住了"传进来一个目录"：它不是坏了，是根本没法预览。
        self.assertFalse(ph.render_file(self.base).ok)

    def test_unreadable_bytes(self) -> None:
        raw = self.base / "raw.md"
        raw.write_bytes(b"\xff\xfe\x00r\x00a\x00p")
        result = ph.render_file(raw)
        self.assertFalse(result.ok)
        self.assertTrue(result.problem.startswith("读不动或解析不动："), result.problem)
        self.assertIn("utf-8", result.problem)
        self.assertEqual("", result.html)

    def test_an_empty_file_is_not_a_problem(self) -> None:
        # 空文件读得动、也解析得动（零个块）：它没有毛病，只是没内容。
        path = self.base / "empty.md"
        path.write_text("", encoding="utf-8")
        result = ph.render_file(path)
        self.assertTrue(result.ok, result.problem)
        self.assertEqual(0, result.stats["blocks"])
        self.assertNotIn("<p></p>", result.html)

    def test_render_file_never_raises_for_any_text(self) -> None:
        # 这一条给主窗口兜底：预览区那一格坏一下不要紧，主窗口不能跟着崩。
        samples = ["", "\n\n\n", "# ", "$\\frac{a}$", "$a & b$", "```\n未闭合",
                   "| a |\n|---", "![](x.png)", B + "[x]\n", ">>> a\n", "<a href=\">"]
        for index, content in enumerate(samples):
            path = self.base / ("sweep%d.md" % index)
            path.write_text(content, encoding="utf-8")
            result = ph.render_file(path)
            self.assertIsInstance(result.html, str, content)
            self.assertIsInstance(result.problem, str, content)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
