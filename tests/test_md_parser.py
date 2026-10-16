# -*- coding: utf-8 -*-
"""md_parser 的自检。

跑法::

    .venv\\Scripts\\python.exe -m unittest discover -s tests -v

这里盯的全是"错了没人看得出来"的那几条：公式定界符、表格切列、代码块里的假公式、
以及屏蔽哨兵有没有漏进产物。
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import md_parser as mp  # noqa: E402


def first(result: mp.ParseResult):
    return result.blocks[0]


class InlineFormulaTest(unittest.TestCase):
    def test_four_delimiters(self) -> None:
        result = mp.parse(r"行内 $a+b$ 与 \(c+d\) 与 $$e+f$$ 与 \[g+h\]。")
        paragraph = first(result)
        formulas = [node for node in paragraph.inline if isinstance(node, mp.MathRun)]
        self.assertEqual([f.latex for f in formulas], ["a+b", "c+d", "e+f", "g+h"])
        # 口径：夹在文字里的 $$ 算行内；\[..\] 在 TeX 里本来就是显示式
        self.assertEqual([f.display for f in formulas], [False, False, False, True])

    def test_star_inside_formula_is_not_emphasis(self) -> None:
        result = mp.parse(r"见 $\alpha * \beta$ 与 *斜体*。")
        nodes = first(result).inline
        math = next(node for node in nodes if isinstance(node, mp.MathRun))
        self.assertEqual(math.latex, r"\alpha * \beta")
        italic = [node for node in nodes if isinstance(node, mp.TextRun) and node.italic]
        self.assertEqual([node.text for node in italic], ["斜体"])

    def test_money_is_not_a_formula(self) -> None:
        """`$5 和 $6` 是钱，不是公式。判成公式会把整句变成一个转不出来的框。"""
        result = mp.parse("单价 $5 和 $6 元。")
        self.assertEqual(mp.render_blocks(result.blocks).count("math"), 0)
        self.assertEqual(first(result).inline[0].text, "单价 $5 和 $6 元。")

    def test_escaped_dollar_is_literal(self) -> None:
        result = mp.parse(r"成本 \$100 与 $x$。")
        text = "".join(node.text for node in first(result).inline if isinstance(node, mp.TextRun))
        self.assertIn("$100", text)
        self.assertEqual(1, len(result.formulas()))

    def test_inline_code_hides_dollars(self) -> None:
        result = mp.parse("命令 `echo $HOME` 后面跟 $x^2$。")
        nodes = first(result).inline
        self.assertIsInstance(nodes[1], mp.CodeSpan)
        self.assertEqual(nodes[1].text, "echo $HOME")
        self.assertEqual([n.latex for n in result.formulas()], ["x^2"])

    def test_bold_italic_and_escape(self) -> None:
        result = mp.parse("**粗** 与 *斜* 与 ***又粗又斜*** 与 \\\\ 反斜杠。")
        flags = [(node.text, node.bold, node.italic) for node in first(result).inline if isinstance(node, mp.TextRun)]
        self.assertIn(("粗", True, False), flags)
        self.assertIn(("斜", False, True), flags)
        self.assertIn(("又粗又斜", True, True), flags)

    def test_snake_case_underscore_is_not_italic(self) -> None:
        result = mp.parse("字段 material_weight_2026 的名字。")
        self.assertTrue(all(not node.italic for node in first(result).inline if isinstance(node, mp.TextRun)))


class LinkImageTest(unittest.TestCase):
    def test_link_keeps_text_and_drops_brackets(self) -> None:
        """方括号与地址不该原样进 Word：地址进超链接，屏幕上只留那几个字。"""
        result = mp.parse("详见[项目说明](docs/readme.md)最后一页。")
        nodes = first(result).inline
        linked = [node for node in nodes if isinstance(node, mp.TextRun) and node.link]
        self.assertEqual([(node.text, node.link) for node in linked], [("项目说明", "docs/readme.md")])
        self.assertNotIn("[", "".join(node.text for node in nodes if isinstance(node, mp.TextRun)))

    def test_image_becomes_placeholder_node(self) -> None:
        result = mp.parse("![图1 接线示意](shots/wiring.png)")
        node = first(result).inline[0]
        self.assertIsInstance(node, mp.ImageRef)
        self.assertEqual((node.alt, node.url), ("图1 接线示意", "shots/wiring.png"))

    def test_bare_bang_stays_text(self) -> None:
        """`!` 后面不是 `[`](...)` 这一对就还是普通文字，不能被当成图片吃掉半个感叹号。"""
        result = mp.parse("注意!! 这里 $x$ 与 ![不是图片")
        text = "".join(node.text for node in first(result).inline if isinstance(node, mp.TextRun))
        self.assertIn("!!", text)
        self.assertIn("![不是图片", text)
        self.assertEqual(result.stats["images"], 0)

    def test_empty_label_falls_back_to_url(self) -> None:
        result = mp.parse("<https://example.com/a> 见 [](https://example.com/b)")
        linked = [node for node in first(result).inline if isinstance(node, mp.TextRun) and node.link]
        self.assertEqual(
            [(node.text, node.link) for node in linked],
            [("https://example.com/a", "https://example.com/a"), ("https://example.com/b", "https://example.com/b")],
        )
        self.assertEqual(result.stats["links"], 2)
        self.assertNotIn("<", "".join(node.text for node in linked))

    def test_emphasis_inside_link_label_survives(self) -> None:
        result = mp.parse("[**必看**指南](a.md)")
        nodes = first(result).inline
        self.assertEqual([node.text for node in nodes], ["必看", "指南"])
        self.assertTrue(nodes[0].bold)
        self.assertTrue(all(node.link == "a.md" for node in nodes))

    def test_url_with_parenthesis_and_title(self) -> None:
        result = mp.parse('[清单](报表(2025).xlsx "扩大工程量清单")')
        node = first(result).inline[0]
        self.assertEqual((node.text, node.link), ("清单", "报表(2025).xlsx"))

    def test_link_in_table_cell_and_stats(self) -> None:
        result = mp.parse("| 名称 | 依据 |\n|---|---|\n| 甲 | [规范](gb.md) |")
        self.assertEqual(result.stats["links"], 1)
        cell = result.blocks[0].rows[1][1]
        self.assertEqual([(node.text, node.link) for node in cell if isinstance(node, mp.TextRun)], [("规范", "gb.md")])


class BlockMathTest(unittest.TestCase):
    def test_single_line_block_math(self) -> None:
        result = mp.parse("$$E=mc^2$$")
        self.assertIsInstance(first(result), mp.BlockMath)
        self.assertEqual(first(result).latex, "E=mc^2")

    def test_multiline_block_math(self) -> None:
        source = "$$\n\\begin{aligned}\na &= b + c \\\\\n&\\quad \\times 2\n\\end{aligned}\n$$"
        result = mp.parse(source)
        self.assertIsInstance(first(result), mp.BlockMath)
        self.assertIn("\\begin{aligned}", first(result).latex)
        self.assertIn("a &= b + c", first(result).latex)

    def test_tex_display_delimiters_multiline(self) -> None:
        result = mp.parse("\\[\n\\frac{1}{2}\n\\]")
        self.assertIsInstance(first(result), mp.BlockMath)
        self.assertEqual(first(result).latex, "\\frac{1}{2}")

    def test_extra_text_on_the_line_is_not_a_block(self) -> None:
        """口径①：一行里除公式还有别的话 → 按行内，不劈成三段。"""
        result = mp.parse("$$a$$ 且 $$b$$ 成立")
        self.assertIsInstance(first(result), mp.Paragraph)
        self.assertEqual(len([n for n in result.formulas() if not n.display]), 2)


class TableTest(unittest.TestCase):
    SOURCE = "\n".join(
        [
            "| 名称 | 表达式 | 说明 |",
            "|:-----|:------:|-----:|",
            "| 绝对值 | $|x|$ | 竖线在公式里 |",
            "| 条件 | $a \\mid b$ | 反斜线竖线 |",
            "| 和 | $\\sum_{i=1}^{n} i$ | 下标 |",
        ]
    )

    def test_pipes_inside_formula_do_not_split_columns(self) -> None:
        """这一条就是自写解析器的理由：按 | 朴素切分会把 $|x|$ 切成三列。"""
        result = mp.parse(self.SOURCE)
        table = next(b for b in result.blocks if isinstance(b, mp.Table))
        self.assertEqual(table.column_count, 3)
        self.assertEqual(len(table.rows), 4)
        for index, row in enumerate(table.rows[1:], start=1):
            self.assertEqual(len(row), 3, "第 %d 行的列数不对" % index)

    def test_cell_formulas_are_inline_nodes(self) -> None:
        result = mp.parse(self.SOURCE)
        table = next(b for b in result.blocks if isinstance(b, mp.Table))
        second_row = table.rows[2]
        math = next(node for node in second_row[1] if isinstance(node, mp.MathRun))
        self.assertEqual(math.latex, "a \\mid b")

    def test_alignment_columns(self) -> None:
        result = mp.parse(self.SOURCE)
        table = next(b for b in result.blocks if isinstance(b, mp.Table))
        self.assertEqual(table.aligns, ["left", "center", "right"])

    def test_display_formula_inside_cell_flag(self) -> None:
        """口径②：单元格里的显示式公式带 display 标记，builder 据此在格内另起一段居中。"""
        result = mp.parse("| 名称 | 公式 |\n|---|---|\n| 功 | \\[W=Fs\\] |")
        table = next(b for b in result.blocks if isinstance(b, mp.Table))
        cell = table.rows[1][1]
        math = next(node for node in cell if isinstance(node, mp.MathRun))
        self.assertTrue(math.display)

    def test_short_row_is_padded_with_warning(self) -> None:
        result = mp.parse("| a | b | c |\n|---|---|---|\n| 1 |")
        table = next(b for b in result.blocks if isinstance(b, mp.Table))
        self.assertEqual(len(table.rows[1]), 3)
        self.assertTrue(any("补齐" in w for w in result.warnings), result.warnings)

    def test_escaped_pipe_stays_in_cell(self) -> None:
        result = mp.parse("| a | b |\n|---|---|\n| x \\| y | z |")
        table = next(b for b in result.blocks if isinstance(b, mp.Table))
        self.assertEqual(table.column_count, 2)
        self.assertIn("x | y", "".join(n.text for n in table.rows[1][0] if isinstance(n, mp.TextRun)))

    def test_no_leading_pipes(self) -> None:
        result = mp.parse("名称 | 表达式\n--- | ---\nx | $y^2$")
        table = next(b for b in result.blocks if isinstance(b, mp.Table))
        self.assertEqual(table.column_count, 2)
        self.assertEqual(len([n for n in result.formulas()]), 1)

    def test_thematic_break_is_not_a_table(self) -> None:
        result = mp.parse("上文\n\n---\n\n下文")
        self.assertEqual(len(result.blocks), 3)
        self.assertIsInstance(result.blocks[1], mp.ThematicBreak)


class OtherBlocksTest(unittest.TestCase):
    def test_headings_levels_and_math_in_heading(self) -> None:
        result = mp.parse("# 一级 $x^2$\n## 二级 ###")
        self.assertEqual([b.level for b in result.blocks], [1, 2])
        self.assertEqual(len(result.formulas()), 1)

    def test_setext_heading(self) -> None:
        result = mp.parse("标题一\n=====\n标题二\n-----")
        self.assertEqual([b.level for b in result.blocks], [1, 2])

    def test_fenced_code_block_keeps_dollars(self) -> None:
        result = mp.parse("```python\nx = \"$${1}$$\"\n```\n\n正文 $y$")
        code = next(b for b in result.blocks if isinstance(b, mp.CodeBlock))
        self.assertEqual(code.lang, "python")
        self.assertIn("$${1}$$", code.text)
        self.assertEqual([n.latex for n in result.formulas()], ["y"])

    def test_unclosed_fence_raises(self) -> None:
        with self.assertRaises(mp.MarkdownSyntaxError):
            mp.parse("```\n只有一行的代码 $x$")

    def test_nested_list_with_formulas_and_task(self) -> None:
        result = mp.parse("- 甲 $a$\n  - [ ] 乙\n  - [x] 丙\n- 丁\n\n1. 第一\n2. 第二 $b$")
        outer = next(b for b in result.blocks if isinstance(b, mp.ListBlock))
        self.assertTrue(outer.ordered is False)
        self.assertEqual(len(outer.items), 2)
        self.assertEqual(len(outer.items[0].children), 2)
        self.assertTrue(outer.items[0].children[1].inline[0].text.startswith("☑"))
        ordered = next(b for b in result.blocks if isinstance(b, mp.ListBlock) and b.ordered)
        self.assertEqual(len(ordered.items), 2)
        self.assertEqual([f.latex for f in result.formulas()], ["a", "b"])

    def test_blockquote_recurses(self) -> None:
        result = mp.parse("> 引用里的公式 $q$ 与\n> 第二行")
        quote = next(b for b in result.blocks if isinstance(b, mp.QuoteBlock))
        self.assertIsInstance(quote.blocks[0], mp.Paragraph)
        self.assertEqual(len(result.formulas()), 1)

    def test_hard_break_becomes_newline(self) -> None:
        result = mp.parse("第一行  \n第二行")
        self.assertIn("\n", first(result).inline[0].text)

    def test_stats_count_every_where(self) -> None:
        result = mp.parse("# 标题 $a$\n\n段落 $b$ 与 $$c$$\n\n$$\nd=e\n$$\n\n| x | y |\n|---|---|\n| $f$ | g |")
        self.assertEqual(result.stats["tables"], 1)
        # $$c$$ 夹在文字里 → 按行内（口径①），所以"显示式"只有那一条独占跨行的 $$
        self.assertEqual(result.stats["formulas_inline"], 4)
        self.assertEqual(result.stats["formulas_display"], 1)

    def test_bullet_list_and_numbered_list_are_two_lists(self) -> None:
        """第一版把这两张并成了一张：圆点列表后面的 1. 2. 会丢序号、变成第三第四个圆点。"""
        result = mp.parse("- 甲\n- 乙\n\n1. 第一\n2. 第二")
        lists = [b for b in result.blocks if isinstance(b, mp.ListBlock)]
        self.assertEqual(len(lists), 2)
        self.assertEqual([b.ordered for b in lists], [False, True])
        self.assertEqual([len(b.items) for b in lists], [2, 2])


class FileReadTest(unittest.TestCase):
    def test_gbk_file_is_read(self) -> None:
        """中文稿件用 GBK 存很常见，只认 UTF-8 会在第一行就报一句不相干的解码错。"""
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "gbk.md"
            path.write_bytes("# 标题 $x^2$\n正文 你好".encode("gbk"))
            result = mp.parse_file(path)
            self.assertEqual(first(result).level, 1)
            self.assertEqual([n.latex for n in result.formulas()], ["x^2"])

    def test_utf8_bom_is_stripped(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "bom.md"
            path.write_bytes("﻿段落 $x$".encode("utf-8"))
            result = mp.parse_file(path)
            self.assertFalse(first(result).inline[0].text.startswith("\ufeff"))


class InternalMaskTest(unittest.TestCase):
    def test_sentinel_never_leaks(self) -> None:
        """哨兵码位漏进产物 = 解析那一刀切错了，宁可抛错也别往 Word 里写怪字符。"""
        text = "混合 | $a|b$ | `x` | \\| 的用例\n| 一 | 二 |\n|---|---|\n| $p|q$ | `r|s` |"
        result = mp.parse(text)
        dumped = mp.render_blocks(result.blocks)
        for ch in (mp.MASK_OPEN, mp.MASK_CLOSE, mp.PIPE_MASK):
            self.assertNotIn(ch, dumped)


if __name__ == "__main__":
    unittest.main()
