# -*- coding: utf-8 -*-
"""gui/validator.py 的自检：四项各查什么、句子长什么样、按什么次序交回。

跑法（用项目内的 venv）::

    .venv\\Scripts\\python.exe -m unittest discover -s tests -v

这一层不碰 Qt、不碰 Word，只吃 :mod:`md_parser` 的解析结果，所以命令行也能挂上去
（界面上那颗「校验」按钮走的是同一条 :func:`gui.validator.validate_parsed`）。
钉的是三件事：

* **档位**：公式语法与图片路径是"进了 Word 会出错"那一档（``error``），
  表格列数与标题跳级只是提醒（``warn``）——解析器已经当场补齐/合并了列，
  Word 里不会多一列也不会少一列；
* **位置**：行号由解析器标注，读不出一份文件的错误写在 0（念成"整份文件"）；
* **一句短路都不许抛**：``latex_problem`` 在本轮之前会被一个裸的 ``&`` 抛穿，
  那等于一条坏公式把整轮校验带崩、后面几项一个字都没查（见 MathCheckTest）。
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import md_parser as mp  # noqa: E402
from gui import validator  # noqa: E402


def check(text: str, base: Path | None = None) -> list[validator.Finding]:
    """一段 Markdown → 校验结果（不写盘的那一条路）。"""
    return validator.validate_parsed(mp.parse(text), base=base)


def codes(findings):
    return [(item.code, item.line) for item in findings]


class MathCheckTest(unittest.TestCase):
    """四项之一：LaTeX 语法。"""

    def test_display_and_inline_both_get_checked(self) -> None:
        found = check("$\\frac{a}$\n\n$$\\frac{a}$$\n")
        self.assertEqual([("math_syntax", 1), ("math_syntax", 3)], codes(found))
        for item in found:
            self.assertEqual(validator.LEVEL_ERROR, item.level)
            self.assertEqual("bad", item.token)

    def test_the_sentence_comes_from_the_checker_verbatim(self) -> None:
        # 校验念的原因与预览退回原文的那一句是同一个判法给出的，这里钉的是拼接形状。
        found = check("$$\\frac{a}$$\n")
        self.assertEqual([("math_syntax", 1)], codes(found))
        self.assertEqual(r"第 1 行｜公式语法：\frac 只跟了 1 组 {…}，它需要两组（原文：\frac{a}）",
                         found[0].text)

    def test_good_math_is_silent(self) -> None:
        self.assertEqual([], check(r"$$E=mc^2$$ 与 $\alpha$ 与 $\frac{1}{2}$\n"))

    def test_bare_ampersand_does_not_abort_the_round(self) -> None:
        # 这一条是本轮修掉的那处生产缺陷的回归：一个裸的 & 让 lxml 抛 XMLSyntaxError，
        # 而异常从 latex_problem 逃出去就把后面三项（图片、表格、标题）整个带走。
        text = "$a & b$\n\n![x](nope.png)\n\n# a\n\n##### b\n"
        found = check(text, base=Path(tempfile.gettempdir()))
        self.assertEqual({"math_syntax", "image_missing", "heading_skip"},
                         {item.code for item in found})
        self.assertIn("没转义的 &", found[0].message)
        self.assertIn("写成 \\&", found[0].message)

    def test_long_source_is_echoed_truncated(self) -> None:
        # 原文要念出来才好对行，但一行日志放不下六十个字符：截到 48 加省略号。
        latex = "x" * 60 + "{" + r"\alpha"
        found = check("$%s$\n" % latex)
        self.assertEqual(1, len(found))
        self.assertTrue(found[0].text.endswith("x" * validator.LATEX_ECHO_CHARS + "…）"),
                        found[0].text)
        self.assertNotIn("x" * (validator.LATEX_ECHO_CHARS + 1), found[0].text)

    def test_echo_survives_none_and_blank(self) -> None:
        self.assertEqual("", validator._echo(""))
        self.assertEqual("", validator._echo(None))
        self.assertEqual(r"\frac{a}", validator._echo(r"\frac{a}"))


class ImageCheckTest(unittest.TestCase):
    """第二项：本地图片路径在不在。"""

    def setUp(self) -> None:
        self._dir = tempfile.TemporaryDirectory()
        self.base = Path(self._dir.name)
        (self.base / "有.png").write_bytes(b"\x89PNG\r\n")
        (self.base / "My Figure.png").write_bytes(b"\x89PNG\r\n")
        self.addCleanup(self._dir.cleanup)

    def test_existing_local_image_is_silent(self) -> None:
        self.assertEqual([], check("![说明](有.png)\n", base=self.base))

    def test_missing_local_image_is_an_error_and_names_the_place_it_looked(self) -> None:
        found = check("![说明](没图.png)\n", base=self.base)
        self.assertEqual([("image_missing", 1)], codes(found))
        self.assertEqual(validator.LEVEL_ERROR, found[0].level)
        self.assertIn("找不到图片：没图.png", found[0].message)
        # 念的是 resolve() 之后的绝对路径，而 tempfile 交回的是 8.3 短名（ADMINI~1），
        # 直接拼 base 比过去永远比不中——这一格的作用是"告诉你去哪儿查的"，就按长名钉。
        self.assertIn(str((self.base / "没图.png").resolve()), found[0].message)

    def test_relative_path_is_resolved_against_the_markdown_directory(self) -> None:
        # base 给的是这份 .md 所在的目录；按当前工作目录查会出现"双击打得开、校验说没有"。
        self.assertEqual([], check("![x](My%20Figure.png)\n", base=self.base))
        self.assertEqual([], check("![x](My Figure.png#top)\n", base=self.base))
        self.assertEqual([], check("![x](My Figure.png?w=200)\n", base=self.base))
        self.assertEqual([("image_missing", 1)], codes(check("![x](sub/ok.png)\n", base=self.base)))

    def test_remote_and_inline_addresses_are_not_looked_up_on_disk(self) -> None:
        for url in ["http://x.test/a.png", "https://x.test/a.png",
                    "data:image/png;base64,AAAA", "mailto:a@b.test", "//cdn.test/a.png",
                    "", "#anchor", "?w=1"]:
            self.assertEqual([], check("![x](%s)\n" % url, base=self.base), url)

    def test_without_base_it_falls_back_to_the_working_directory(self) -> None:
        self.assertEqual([("image_missing", 1)], codes(check("![x](绝对不在这里的.png)\n")))

    def test_image_inside_a_table_cell_is_checked_too(self) -> None:
        text = "| 图 | 名 |\n|---|---|\n| ![](哪里也没有.png) | 导线 |\n"
        # 行号来自解析器标注的**表格起始行**（表格是块级节点，行内的图不单独标行）：
        # 念 1 而不是 3，人对过去看到的是同一张表，而"哪一格"由 url 自己说清。
        self.assertEqual([("image_missing", 1)], codes(check(text, base=self.base)))


class TableCheckTest(unittest.TestCase):
    """第三项：列数不一致——只当提醒，不当错误。"""

    def test_short_row_is_a_warning_with_the_parsers_own_sentence(self) -> None:
        found = check("| a | b | c |\n|---|---|---|\n| 1 | 2 |\n")
        self.assertEqual([("table_columns", 3)], codes(found))
        self.assertEqual(validator.LEVEL_WARN, found[0].level)
        self.assertEqual("第 3 行｜表格列数：只有 2 列，表头是 3 列，缺的按空格补齐", found[0].text)

    def test_long_row_is_a_warning_too(self) -> None:
        found = check("| a | b |\n|---|---|\n| 1 | 2 | 3 |\n")
        self.assertEqual(validator.LEVEL_WARN, found[0].level)
        self.assertEqual("第 3 行｜表格列数：有 3 列，比表头 2 列多，多出来的已并进最后一格",
                         found[0].text)

    def test_consistent_table_is_silent(self) -> None:
        self.assertEqual([], check("| a | b |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |\n"))


class HeadingCheckTest(unittest.TestCase):
    """第四项：跳级。只看相邻两级，也不要求第一条是 ``#``。"""

    def test_first_heading_need_not_be_level_one(self) -> None:
        # 一份只写章节的说明文件从 ## 起头是常见写法，报它等于报噪音。
        self.assertEqual([], check("## a\n\n### b\n"))

    def test_jumping_two_levels_is_a_warning(self) -> None:
        found = check("## a\n\n##### b\n")
        self.assertEqual([("heading_skip", 3)], codes(found))
        self.assertEqual(validator.LEVEL_WARN, found[0].level)
        self.assertEqual("第 3 行｜标题层级：标题从第 2 级直接跳到第 5 级"
                         "（中间少了第 4 级，Word 的导航窗格会少一层）", found[0].text)

    def test_level_drop_is_legal(self) -> None:
        # 往下走（第 5 级到第 2 级）不算跳级：导航窗格不少层，只是多嵌一层。
        # 这一份里唯一那一条是"1 → 5"往上跳，紧随其后的 5 → 2 一声不吭。
        self.assertEqual([("heading_skip", 3)], codes(check("# a\n\n##### b\n\n## c\n")))
        self.assertEqual([], check("# a\n\n## b\n\n# c\n"))

    def test_heading_with_math_inside_still_counts(self) -> None:
        found = check("# a\n\n##### " + r"$\frac{a}$" + "\n")
        self.assertIn("heading_skip", [item.code for item in found])


class OrderingTest(unittest.TestCase):
    """四项合成一份结果时的次序。"""

    def test_sorted_by_line_with_unknown_line_last(self) -> None:
        text = ("| a | b |\n|---|---|\n| 1 |\n\n"
                "## x\n\n##### y\n\n$$\\frac{a}$$\n")
        found = check(text)
        self.assertEqual(validator.LEVEL_ERROR, found[-1].level)
        self.assertEqual([(3, "table_columns"), (7, "heading_skip"), (9, "math_syntax")],
                         [(item.line, item.code) for item in found])

    def test_the_four_checks_run_even_when_an_earlier_one_found_something(self) -> None:
        text = ("$$\\frac{a}$$\n\n![x](nope.png)\n\n| a | b |\n|---|---|\n| 1 |\n\n"
                "# a\n\n##### b\n")
        self.assertEqual({"math_syntax", "image_missing", "table_columns", "heading_skip"},
                         {item.code for item in check(text)})


class ReportShapeTest(unittest.TestCase):
    """句子、档位标记与概要行。"""

    def test_text_prefix_and_deduplication(self) -> None:
        item = validator.Finding(validator.LEVEL_ERROR, "math_syntax", 3, r"\frac 少了一组")
        self.assertEqual(r"第 3 行｜公式语法：\frac 少了一组", item.text)
        # 解析器那句本来就带着"第 N 行"（它得自己说清是哪一行），拼起来不许念两遍。
        twice = validator.Finding(
            validator.LEVEL_WARN, "table_columns", 3, "第 3 行只有 2 列，表头是 3 列，缺的按空格补齐")
        self.assertEqual("第 3 行｜表格列数：只有 2 列，表头是 3 列，缺的按空格补齐", twice.text)
        # 剥的是那一句行号本身，后面的分隔符归这里拼——所以消息里不该预先带"｜"。
        self.assertEqual("第 3 行｜表格列数：｜这一句照解析器的写法带了分隔符",
                         validator.Finding(validator.LEVEL_WARN, "table_columns", 3,
                                           "第 3 行｜这一句照解析器的写法带了分隔符").text)

    def test_unknown_line_reads_as_the_whole_file(self) -> None:
        item = validator.Finding(validator.LEVEL_ERROR, "file_missing", 0, "已经不在这个位置了")
        self.assertEqual("整份文件｜文件：已经不在这个位置了", item.text)
        self.assertEqual("整份文件", validator.where(0))

    def test_token_is_the_heaviest_level_present(self) -> None:
        self.assertEqual("bad", validator.Finding(validator.LEVEL_ERROR, "math_syntax", 1, "x").token)
        self.assertEqual("warn", validator.Finding(validator.LEVEL_WARN, "heading_skip", 1, "x").token)
        # 认不出的档位往严的一档走：宁可上红也不假装没事。
        self.assertEqual("warn", validator.Finding("info", "heading_skip", 1, "x").token)
        self.assertEqual("warn", validator.Finding("", "heading_skip", 1, "x").token)

    def test_summary_three_branches(self) -> None:
        self.assertEqual("四项都没查出问题", validator.summary_text([]))
        warn = [validator.Finding(validator.LEVEL_WARN, "heading_skip", 1, "x")]
        self.assertEqual("查出 1 处提醒（没有会出错的那一档）", validator.summary_text(warn))
        mixed = warn + [validator.Finding(validator.LEVEL_ERROR, "math_syntax", 2, "x")]
        self.assertEqual("查出 2 处问题，其中 1 处会让 Word 里那份成果出错",
                         validator.summary_text(mixed))

    def test_worst_token_marks_the_whole_round(self) -> None:
        self.assertEqual("ok", validator.worst_token([]))
        self.assertEqual("warn",
                         validator.worst_token([validator.Finding(validator.LEVEL_WARN,
                                                                  "table_columns", 1, "x")]))
        self.assertEqual("bad",
                         validator.worst_token([validator.Finding(validator.LEVEL_WARN,
                                                                  "table_columns", 1, "x"),
                                                validator.Finding(validator.LEVEL_ERROR,
                                                                  "image_missing", 2, "x")]))

    def test_every_code_has_a_chinese_label(self) -> None:
        self.assertEqual({"math_syntax": "公式语法", "image_missing": "图片路径",
                          "table_columns": "表格列数", "heading_skip": "标题层级",
                          "file_missing": "文件", "parse_failed": "解析"},
                         dict(validator.CODE_LABELS))


class WholeFileTest(unittest.TestCase):
    """整份文件那两种读不到解析结果的情况。"""

    def setUp(self) -> None:
        self._dir = tempfile.TemporaryDirectory()
        self.base = Path(self._dir.name)
        self.addCleanup(self._dir.cleanup)

    def test_missing_file_is_one_error_and_does_not_raise(self) -> None:
        found = validator.validate_file(self.base / "nope.md")
        self.assertEqual([("file_missing", 0)], codes(found))
        self.assertIn("已经不在这个位置了", found[0].text)
        self.assertEqual("bad", validator.worst_token(found))

    def test_unreadable_file_becomes_one_error(self) -> None:
        raw = self.base / "raw.md"
        raw.write_bytes(b"\xff\xfe\x00r\x00a\x00p")
        found = validator.validate_file(raw)
        self.assertEqual([("parse_failed", 0)], codes(found))
        self.assertIn("读不动或解析不动", found[0].text)

    def test_a_healthy_file_reports_nothing(self) -> None:
        good = self.base / "good.md"
        good.write_text("# 标题\n\n## 小节\n\n这里有一个 $\\frac{a}{b}$ 与 $$E=mc^2$$。\n\n"
                        "| a | b |\n|---|---|\n| 1 | 2 |\n", encoding="utf-8")
        self.assertEqual([], validator.validate_file(good))


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
