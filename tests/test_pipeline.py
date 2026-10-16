# -*- coding: utf-8 -*-
"""命令行这一层的自检：``python md2docx.py 输入.md 输出.docx`` 那一行。

这一层不该有任何转换逻辑（逻辑在 md_parser / formula_converter / docx_builder 里，各有自己的用例），
它要管的是**人会不会用错、出了错能不能自己看懂**，所以断言集中在四件事：

1. 退出码一档一个原因（脚本与批处理要靠它分诊，不能全回 1）。
2. 拒绝的那句话写着**下一步该干什么**，而且同一件事只说一遍。
3. 命令行那几个可选参数真的落进了 ``word/styles.xml``——参数只收到内存里等于没实现。
4. 合并那一族（``--merge`` / ``--toc`` / ``--cover`` / ``--page-break`` / ``--source-label``）
   的判据读 ``word/document.xml``：封面那三行、``TOC`` 域的代码、分页符、来源标注那几行都得数得出来。
   合并策略本身在 :mod:`gui.merger` 里已有用例，这一层只核"命令行把开关传到了那一步、
   并把这一次做了什么念给了人"，外加钉一条**一个输入且一格开关都没打时根本不调用
   ``merge_entries``**——那一路的产物必须与没有合并这一层时逐字节相同。

和 ``test_docx_builder.py`` 一样，判据读的是**落盘的文件**，不是内存里的对象。
"""

from __future__ import annotations

import contextlib
import io
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import docx_builder as db  # noqa: E402
import main as m  # noqa: E402

try:
    from formula_converter import locate_xslt

    XSLT = locate_xslt()
except Exception:  # pragma: no cover - 取决于机器上有没有 Office
    XSLT = None

needs_office = unittest.skipIf(XSLT is None, "本机找不到 MML2OMML.XSL，转换用例无法真跑")

#: latex2mathml 认不出的一条式子（与 test_docx_builder 用同一款毛病）
BAD_LATEX = r"\foo \bar"
#: 另一条毛病不同的：``\foo \bar2`` 反而转得出来（2 成了 \\bar 的参数），所以两条坏的要各挑一句
BAD_LATEX_2 = r"\sqrt["

GOOD_MD = """# 线路损耗

灵敏度 $S = \\beta I_0$ 一句话。

$$P = UI\\cos\\varphi$$

| 名称 | 表达式 |
| --- | --- |
| 功率 | $P = UI$ |
"""

#: GOOD_MD 里三条公式全都转得成（行内 2 + 显示式 1）——自检那句比的"内存里的条数"就是这个数
EXPECTED_MATH_OK = 3

# --- 合并那几格用的三份稿子（刻意各带一种毛病，让每一格都有东西可核）-------------

#: 一条 ``######``：降级时它撞六级上限，回执要念"有 1 条没法再降"
MERGE_A = "# 甲篇\n\n## 一节\n\n灵敏度 $S = \\beta I_0$ 一句话。\n\n###### 已经到底了\n"
#: 只有一条显示式——合并稿里的公式数就是这一份加上一份行内
MERGE_B = "# 乙篇\n\n正文一段。\n\n$$P = UI\\cos\\varphi$$\n"
MERGE_C = "# 丙篇\n\n最后一句。\n"
#: 三份合起来：行内 1 + 显示式 1
MERGE_MATH_OK = 2


class _Cli(unittest.TestCase):
    """"写一份 .md 到临时目录 → 调 run() → 读回落盘的包"收成几个方法。"""

    maxDiff = None

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="md2docx-cli-"))
        # 临时目录归临时目录：不清就会每跑一条用例在 %TEMP% 里落一个，跑测的人不会知道
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.source = self.tmp / "in.md"
        self.target = self.tmp / "out.docx"
        self.write_source(GOOD_MD)

    def write_source(self, text: str, encoding: str = "utf-8") -> Path:
        self.source.write_bytes(text.encode(encoding))
        return self.source

    def go(self, *argv: str) -> tuple[int, str]:
        """``run(输入, 输出, …)`` 的常用档：前两个参数固定是这两个临时文件。"""
        return m.run([str(self.source), str(self.target), *[str(a) for a in argv]])

    def member(self, name: str) -> str:
        with zipfile.ZipFile(str(self.target)) as package:
            return package.read(name).decode("utf-8")


class HappyPathTest(_Cli):
    @needs_office
    def test_default_run_writes_file_and_says_it_verified(self) -> None:
        code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertTrue(self.target.is_file(), text)
        # 这一句是"公式在 Word 里是原生对象"的证据：数的是文件里的 m:oMath
        self.assertIn("自检：文件里有 3 个 m:oMath（其中显示式 1 个 m:oMathPara），与内存里转成的条数一致", text)
        self.assertIn("输出：%s" % str(self.target.resolve()), text)

    @needs_office
    def test_receipt_counts_match_the_file(self) -> None:
        code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("块 4 个｜表格 1 张｜公式 3 条（转成 3，失败 0）", text)
        self.assertEqual(db.count_omml(self.target), (3, 1), text)

    @needs_office
    def test_parse_counts_line_is_printed_without_verbose(self) -> None:
        """回执里"转成 3"数的是真落笔的那几条；稿子里一共几条得另念一句，

        否则一份只写了 ``$`` 没配对的稿子看起来也是"成功"。
        """
        code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("解析计数：块 4｜标题 1｜行内公式 2｜显示式 1｜代码块 0｜表格 1", text)

    @needs_office
    def test_verbose_swaps_the_counts_line_for_the_block_tree(self) -> None:
        _, plain = self.go()
        code, loud = self.go("-v")
        self.assertEqual(code, m.EXIT_OK, loud)
        self.assertIn("解析计数：", plain)
        self.assertIn("块树统计：", loud)
        self.assertNotIn("解析计数：", loud)

    @needs_office
    def test_missing_output_directory_is_created(self) -> None:
        """命令行指到没建过的子目录是很常见的写法，不该为此回一句"写不出"。"""
        deep = self.tmp / "新建" / "第二层" / "out.docx"
        code, text = m.run([str(self.source), str(deep)])
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertTrue(deep.is_file())

    @needs_office
    def test_running_twice_overwrites_the_same_target(self) -> None:
        """重跑是常态（改完稿子再导一次）：第二次不该因为"文件已存在"而拒绝。"""
        codes = []
        for _ in range(2):
            with contextlib.redirect_stdout(io.StringIO()):
                codes.append(m.main([str(self.source), str(self.target)]))
        self.assertEqual(codes, [m.EXIT_OK, m.EXIT_OK])
        self.assertEqual(db.count_omml(self.target), (3, 1))


class OptionsTest(_Cli):
    @needs_office
    def test_fonts_and_size_reach_styles_xml(self) -> None:
        """--body-font / --body-size / --heading-font 的验收标准是 styles.xml 里那几格。

        中文字体必须落到 ``w:eastAsia``，标题字号是"正文号 + 增量"。
        """
        code, text = self.go("--body-font", "微软雅黑", "--body-size", "10.5", "--heading-font", "楷体")
        self.assertEqual(code, m.EXIT_OK, text)
        styles = self.member("word/styles.xml")
        self.assertIn('w:eastAsia="微软雅黑"', styles)
        self.assertIn('w:eastAsia="楷体"', styles)
        self.assertIn('<w:sz w:val="21"/>', styles)  # 10.5pt = 21 半磅

    @needs_office
    def test_code_and_formula_fonts(self) -> None:
        self.write_source("代码块：\n\n```python\nx = 1\n```\n\n行内 $a+b$。\n")
        code, text = self.go("--code-font", "Courier New", "--formula-font", "Cambria Math")
        self.assertEqual(code, m.EXIT_OK, text)
        document = self.member("word/document.xml")
        self.assertIn('w:ascii="Courier New"', document)
        self.assertIn('w:ascii="Cambria Math"', document)

    def test_body_size_must_be_positive(self) -> None:
        """0 与负数会写进 ``w:sz``，Word 里就是一个开不开的包——argparse 那一档拒掉。"""
        for bad in ("0", "-3", "十二"):
            with self.subTest(value=bad), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as caught:
                    self.go("--body-size", bad)
            self.assertEqual(caught.exception.code, m.EXIT_USAGE)

    def test_missing_positional_argument_is_usage(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as caught:
                m.run([str(self.source)])
        self.assertEqual(caught.exception.code, m.EXIT_USAGE)

    def test_help_lists_the_exit_codes(self) -> None:
        """值班的人手里只有 --help：退出码分档必须写在里面，不然只能重跑一遍才知道差在哪。"""
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            with self.assertRaises(SystemExit):
                m.build_parser().parse_args(["--help"])
        text = buffer.getvalue()
        for code in (m.EXIT_XSLT, m.EXIT_INPUT, m.EXIT_FORMULA, m.EXIT_WRITE):
            self.assertIn(str(code), text)


class InputErrorTest(_Cli):
    def test_missing_input_file(self) -> None:
        code, text = m.run([str(self.tmp / "没有这个文件.md"), str(self.target)])
        self.assertEqual(code, m.EXIT_INPUT)
        self.assertIn("没有这个文件.md", text)
        self.assertFalse(self.target.exists())

    def test_input_and_output_are_the_same_file(self) -> None:
        code, text = m.run([str(self.source), str(self.source)])
        self.assertEqual(code, m.EXIT_USAGE)
        self.assertIn("同一个文件", text)

    def test_forced_encoding_that_does_not_fit(self) -> None:
        """--encoding 是给"自动试探认不出"的稿子用的；指错了就说读不进，不拿乱码出一份文档。"""
        self.write_source("# 标题\n\n正文。\n", "gb18030")
        code, text = self.go("--encoding", "utf-8")
        self.assertEqual(code, m.EXIT_INPUT)
        self.assertIn("读不进", text)
        self.assertFalse(self.target.exists())

    @needs_office
    def test_gb18030_is_detected_and_said_out_loud(self) -> None:
        """Windows 上"另存为 ANSI"的 .md 真常见：自动认出来，并把按哪种编码读的念出来。"""
        self.write_source("# 标题\n\n正文一段。\n", "gb18030")
        code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("gb18030", text)

    @needs_office
    def test_utf8_bom_does_not_leak_into_the_document(self) -> None:
        self.write_source("# 标题\n\n正文一段。\n", "utf-8-sig")
        code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertNotIn("\ufeff", self.member("word/document.xml"))


class StylesheetErrorTest(_Cli):
    def test_xslt_path_that_does_not_exist_says_where_it_looked(self) -> None:
        """指了路径却没有：这是最容易被读成"程序坏了"的一种失败，要念出查过的位置。"""
        ghost = self.tmp / "MML2OMML.XSL"
        code, text = self.go("--xslt-path", str(ghost))
        self.assertEqual(code, m.EXIT_XSLT)
        self.assertIn(str(ghost), text)
        self.assertIn("--xslt-path", text)
        self.assertFalse(self.target.exists())

    def test_the_solution_is_stated_once(self) -> None:
        """回归：异常自己带"解决办法"，命令行这一层再拼一份就会出现两处说法不一样的下一步。

        （真栽过一次：同一段"解决办法"在输出里出现了两遍。）
        """
        code, text = self.go("--xslt-path", str(self.tmp / "无此文件.XSL"))
        self.assertEqual(code, m.EXIT_XSLT)
        self.assertEqual(text.count("解决办法"), 1, text)

    def test_a_file_that_is_not_the_stylesheet_names_that_file(self) -> None:
        """--xslt-path 指到一个存在但不是 XSL 的文件（拷坏了、指错了对象）：

        念的是**这个文件**读不通，而不是让人以为"Office 没装"。
        """
        broken = self.tmp / "MML2OMML.XSL"
        broken.write_text("这只是一段中文，不是 XML", encoding="utf-8")
        code, text = self.go("--xslt-path", str(broken))
        self.assertEqual(code, m.EXIT_XSLT)
        self.assertIn(str(broken), text)
        self.assertFalse(self.target.exists())


class FormulaFailureTest(_Cli):
    @needs_office
    def test_default_keeps_source_text_and_still_writes(self) -> None:
        """几百条公式的稿子不该为一条崩掉：退回原文 + 黄色底纹，退出码仍是成功。"""
        self.write_source("正常 $a+b$，坏掉 $%s$。\n" % BAD_LATEX)
        code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("失败 1", text)
        self.assertIn(BAD_LATEX, text)
        self.assertIn('<w:highlight w:val="yellow"/>', self.member("word/document.xml"))

    @needs_office
    def test_strict_fails_and_writes_nothing(self) -> None:
        """--strict 的意义是"这条稿子还没干净"：既然判失败，就不该留一份半成品在盘上。"""
        self.write_source("坏掉 $%s$。\n" % BAD_LATEX)
        code, text = self.go("--strict")
        self.assertEqual(code, m.EXIT_FORMULA)
        self.assertIn("--strict", text)
        self.assertFalse(self.target.exists())

    @needs_office
    def test_strict_with_a_clean_document_succeeds(self) -> None:
        code, text = self.go("--strict")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertTrue(self.target.is_file())

    @needs_office
    def test_receipt_lists_every_failed_formula_once(self) -> None:
        self.write_source("$%s$ 与 $%s$。\n" % (BAD_LATEX, BAD_LATEX_2))
        code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("转换失败的公式", text)
        self.assertEqual(text.count("  × "), 2, text)


class WriteErrorTest(_Cli):
    @needs_office
    def test_output_path_is_a_directory(self) -> None:
        """指到一个目录（"输出到 out 里"少写了文件名）：单独一档退出码，别和"读不进"混。"""
        code, text = m.run([str(self.source), str(self.tmp)])
        self.assertEqual(code, m.EXIT_WRITE, text)
        self.assertIn("写不出", text)


class VerificationLineTest(_Cli):
    """自检那一句本身的两种说法（不一致时必须比"成功"更响）。"""

    class _Counts:
        def __init__(self, math_ok: int) -> None:
            self.math_ok = math_ok

    @needs_office
    def test_consistent_counts_say_so(self) -> None:
        code, _ = self.go()
        self.assertEqual(code, m.EXIT_OK)
        line = m._verification_line(self.target, self._Counts(EXPECTED_MATH_OK))
        self.assertIn("一致", line)
        self.assertNotIn("不一致", line)

    @needs_office
    def test_mismatch_is_called_out_and_carries_no_console_noise(self) -> None:
        code, _ = self.go()
        self.assertEqual(code, m.EXIT_OK)
        line = m._verification_line(self.target, self._Counts(999))
        self.assertIn("不一致", line)
        self.assertIn("不要直接交这一份", line)
        # 回执是打到控制台给人念的，markdown 的星号到屏幕上就是两粒噪声
        self.assertNotIn("**", line)

    def test_unreadable_package_says_it_could_not_verify(self) -> None:
        """文件写出去了但读不回（被占用、盘坏了）：这一句要说"没核出来"，不能报成功。"""
        line = m._verification_line(self.tmp / "没有这个文件.docx", self._Counts(1))
        self.assertIn("读不回", line)
        self.assertNotIn("一致", line)


class MergeCliTest(_Cli):
    """多输入并成一份，以及 ``--merge`` 那一族开关在 ``word/document.xml`` 里的样子。

    三份稿子各带一种毛病：甲篇有一条 ``######``（降级撞六级上限）、乙篇只有一条显示式、
    丙篇最短（用来验"最后一个文件前面换页、第一个前面不换"）。
    """

    def setUp(self) -> None:
        super().setUp()
        self.files = []
        for name, text in (("a.md", MERGE_A), ("b.md", MERGE_B), ("c.md", MERGE_C)):
            path = self.tmp / name
            path.write_text(text, encoding="utf-8")
            self.files.append(path)

    def merge(self, *argv: str) -> tuple[int, str]:
        """三个输入 → 一份输出；前面的位置参数是开关。"""
        return m.run([str(path) for path in self.files] + [str(self.target), *[str(a) for a in argv]])

    @property
    def document(self) -> str:
        return self.member("word/document.xml")

    def instruction(self) -> str:
        """``w:instrText`` 里那一段域代码——目录收到第几级只有这里看得出来。"""
        start = self.document.index(" TOC")
        return self.document[start:self.document.index("</w:instrText>", start)]

    @needs_office
    def test_three_files_become_one_package(self) -> None:
        code, text = self.merge("--cover", "--title", "投标稿")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("输入（并成一份的顺序）：", text)
        self.assertIn("1. " + str(self.files[0]), text)
        self.assertIn("封面标题：投标稿｜共 3 个文件", text)
        document = self.document
        self.assertIn("投标稿", document)
        self.assertIn("共 3 个文件", document)
        # 封面那行日期是合并策略现算的，用例只认它的样子，不认今天是几号
        self.assertRegex(document, r"\d{4}年\d{1,2}月\d{1,2}日")
        self.assertIn(">目录<", document)
        self.assertEqual(document.count("来源："), 3)
        self.assertIn("来源：a.md", document)

    @needs_office
    def test_source_label_lines_do_not_count_into_the_parse_counts(self) -> None:
        """回执那句"来源标注不计进解析计数"是承诺，不是排版：两行的差就该正好是来源行数。"""
        code, text = self.merge()
        self.assertEqual(code, m.EXIT_OK, text)
        written = int(re.search(r"块 (\d+) 个", text).group(1))
        tree = int(re.search(r"解析计数：块 (\d+)", text).group(1))
        self.assertEqual(written - tree, len(self.files))

    @needs_office
    def test_headings_are_demoted_and_the_toc_reaches_level_four(self) -> None:
        """降过一级，原来的 ``#`` 落在第 2 级；域代码还写 ``1-3`` 就会把它底下的 ``###`` 漏掉。"""
        code, text = self.merge()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertNotIn('w:val="Heading1"', self.document)
        self.assertIn('w:val="Heading2"', self.document)
        self.assertIn('"1-4"', self.instruction())

    @needs_office
    def test_no_demote_keeps_the_first_level_and_the_default_toc_range(self) -> None:
        code, text = self.merge("--no-demote")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("标题降级 关", text)
        self.assertIn('w:val="Heading1"', self.document)
        self.assertIn('"1-3"', self.instruction())

    @needs_office
    def test_demoting_past_level_six_is_said_out_loud(self) -> None:
        """``######`` 没处再降：一声不吭按第 6 级排下去会被读成"降级没生效"。"""
        code, text = self.merge()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("有 1 条已经在第 6 级", text)
        self.assertIn('w:val="Heading6"', self.document)

    @needs_office
    def test_the_toc_can_be_switched_off(self) -> None:
        code, text = self.merge("--no-toc")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("目录 关", text)
        self.assertNotIn("TOC", self.document)
        self.assertNotIn("这一栏是 Word 的目录域", self.document)

    @needs_office
    def test_the_source_label_can_be_switched_off(self) -> None:
        code, text = self.merge("--no-source-label")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("来源标注 关", text)
        self.assertNotIn("来源：", self.document)

    @needs_office
    def test_the_page_break_can_be_switched_off(self) -> None:
        """关掉目录才判这一格——前置页自己那两处分页不归它管（见 :meth:`_write_front_matter`）。"""
        code, text = self.merge("--no-page-break", "--no-toc")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("分页符 关", text)
        self.assertNotIn('w:type="page"', self.document)

    @needs_office
    def test_breaks_sit_between_files_and_after_the_front_pages(self) -> None:
        """三个文件两处换页；封面自己占一页、目录自己占一页，那两处分页**不归**那个开关管。"""
        code, text = self.merge()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertEqual(self.document.count('w:type="page"'), 3)
        code, text = self.merge("--cover")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertEqual(self.document.count('w:type="page"'), 4)

    @needs_office
    def test_the_cover_is_off_by_default(self) -> None:
        """多要一页纸是额外决定，不是"合并"本身含的意思（界面与命令行同一档）。"""
        code, text = self.merge()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("封面 关", text)
        self.assertNotIn("共 3 个文件", self.document)

    @needs_office
    def test_the_title_defaults_to_the_output_stem(self) -> None:
        code, text = self.merge("--cover")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("封面标题：out｜共 3 个文件", text)
        self.assertIn(">out<", self.document)

    @needs_office
    def test_math_still_gets_verified_after_the_merge(self) -> None:
        """并树之后自检那句比的仍是落盘的包——合并不是"公式丢了也报成功"的口子。"""
        code, text = self.merge()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("与内存里转成的条数一致", text)
        self.assertEqual(db.count_omml(self.target), (MERGE_MATH_OK, 1))

    @needs_office
    def test_one_input_with_merge_gets_a_toc_but_no_demote(self) -> None:
        """``--merge`` 让一份稿子也能用目录/封面；那一路的「输入」那句仍念字节与编码（原话一字没改）。"""
        code, text = self.go("--merge", "--cover", "--title", "一份稿")
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertIn("标题降级 关", text)
        self.assertIn(">目录<", self.document)
        self.assertIn("输入：%s（utf-8，" % self.source, text)

    @needs_office
    def test_no_switch_and_one_input_never_enters_the_merge_layer(self) -> None:
        """老路是硬约定：一格开关都没打时产物与回执都该与"没有合并这一层"时一样。"""
        with mock.patch.object(m.mg, "merge_entries", side_effect=AssertionError("不该走合并")):
            code, text = self.go()
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertNotIn("合并说明", text)

    @needs_office
    def test_the_merge_layer_is_used_as_soon_as_there_is_a_second_input(self) -> None:
        with mock.patch.object(m.mg, "merge_entries", wraps=m.mg.merge_entries) as spy:
            code, text = m.run([str(self.files[0]), str(self.files[1]), str(self.target)])
        self.assertEqual(code, m.EXIT_OK, text)
        self.assertEqual(len(spy.call_args[0][0]), 2)
        self.assertTrue(spy.call_args[0][1].demote)

    def test_a_bad_second_file_says_which_one(self) -> None:
        """三个文件里第二个读不进：得念"第 2 个文件"，否则人先去改第一个。"""
        ghost = self.tmp / "没有这个文件.md"
        code, text = m.run([str(self.files[0]), str(ghost), str(self.files[2]), str(self.target)])
        self.assertEqual(code, m.EXIT_INPUT)
        self.assertIn("第 2 个文件", text)
        self.assertIn("没有这个文件.md", text)
        self.assertFalse(self.target.exists())
        # 单独一份时那句前缀不加：脚本里可能已经在按"这个 .md 读不进"那几行判
        single, only = m.run([str(ghost), str(self.target)])
        self.assertEqual(single, m.EXIT_INPUT)
        self.assertNotIn("第 1 个文件", only)

    def test_an_input_that_is_also_the_output_is_usage(self) -> None:
        code, text = m.run([str(self.files[0]), str(self.files[1]), str(self.files[0])])
        self.assertEqual(code, m.EXIT_USAGE)
        self.assertIn("同一个文件", text)

    def test_help_lists_the_merge_switches(self) -> None:
        """合并那几格的默认档（合并时开、单独一份关）写在 --help 里，不然只能读源码。"""
        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            with self.assertRaises(SystemExit):
                m.build_parser().parse_args(["--help"])
        text = buffer.getvalue()
        for flag in ("--merge", "--toc", "--no-toc", "--cover", "--no-cover", "--title",
                     "--page-break", "--no-page-break", "--source-label", "--no-source-label",
                     "--no-demote"):
            self.assertIn(flag, text)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
