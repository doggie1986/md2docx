# -*- coding: utf-8 -*-
"""合并成单一 Word 那六条的自检（降级 / 分页 / 来源行 / 目录域 / 封面 / 复用两个模块）。

判据分三层，各钉各的，因为这三层会各自"自己没错而整体不对"：

1. :mod:`gui.merger` 那一层的**决定**——标题降几级、换页点落在第几块、封面那三行写什么字。
   纯数据，不碰 Word，跑得快也不需要 Office。
2. 落盘文件里的**那一串 XML**——目录域（``fldChar``/``instrText``）、``w:updateFields``、
   分页符、``w:pStyle``。合并这件事只有读回文件才算证据：内存里的声明再对，
   没落进包就等于页面上没有目录。
3. :func:`gui.tasks.convert_paths` 那一路**照面板的勾走**——开关与文件间接错线时，
   前两层各自都对而盘上还是没有目录。

命令行 / 单独输出那两路必须**一个字节都不多写**（第 5 条那个类）。它钉的是
"没合并的稿子不该因为这一轮而改了样子"——那是最容易被人当成"数据变了"的错。
"""

from __future__ import annotations

import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import docx_builder as db  # noqa: E402
import md_parser as mp  # noqa: E402
from gui import merger as mg  # noqa: E402

try:
    from formula_converter import FormulaConverter, locate_xslt

    XSLT = locate_xslt()
except Exception:  # pragma: no cover - 取决于机器上有没有 Office
    XSLT = None

needs_office = unittest.skipIf(XSLT is None, "本机找不到 MML2OMML.XSL，写不出真的 .docx")

MD_A = """# 甲册 说明

## 适用范围

正文一句 $x=1$。

###### 已经到顶的那一级

> ### 引用里的三级
"""

MD_B = """# 乙册 计算

## 负荷

$$P = UI$$
"""

#: 两份稿子各自解析出来的块数（甲 5 + 乙 3）——换了稿子这几条下标断言要跟着改。
BLOCKS_A = 5
BLOCKS_B = 3


def parse(text: str) -> mp.ParseResult:
    return mp.parse(text)


def headings(blocks) -> list:
    """块树里所有标题的层级（引用里的那些一并摊平，按出现顺序）。"""
    out = []
    for block in blocks:
        if isinstance(block, mp.Heading):
            out.append(block.level)
        elif isinstance(block, mp.QuoteBlock):
            out.extend(headings(block.blocks))
    return out


def label_text(block) -> str:
    """一段正文里那几个 run 的文本拼起来（只用于判来源行，不认别的）。"""
    if not isinstance(block, mp.Paragraph):
        return ""
    return "".join(getattr(run, "text", "") for run in block.inline)


ALL_ON = mg.MergeSwitches(demote=True, page_break=True, source_label=True, toc=True, cover=True)


class MergerDecisionTest(unittest.TestCase):
    """第 1 层：这一层只交回"决定"，一个 Word 概念都不碰。"""

    def setUp(self) -> None:
        self.a, self.b = Path("甲.md"), Path("乙.md")
        self.pa, self.pb = parse(MD_A), parse(MD_B)

    def entries(self, *pairs):
        return list(pairs) or [(self.a, self.pa), (self.b, self.pb)]

    def test_headings_move_down_one_level_so_the_book_title_is_free(self) -> None:
        """需求第 2 条：# 降为 ##、## 降为 ###（引用里的那一条也降）。"""
        plan = mg.merge_entries(self.entries(), ALL_ON)
        self.assertEqual(headings(plan.parsed.blocks)[:4], [2, 3, 6, 4])

    def test_a_batch_of_one_file_is_left_alone(self) -> None:
        """只有一份稿子时它自己就是完整的一级结构：降了级反而凭空少一层。"""
        plan = mg.merge_entries(self.entries((self.a, self.pa)), ALL_ON)
        self.assertEqual(headings(plan.parsed.blocks), [1, 2, 6, 3])

    def test_a_heading_already_at_six_stays_and_says_so(self) -> None:
        plan = mg.merge_entries(self.entries(), ALL_ON)
        self.assertEqual(plan.clamped, 1)
        self.assertEqual(len(plan.notes()), 1)
        self.assertIn("第 6 级", plan.notes()[0])
        self.assertIn("%d 条" % plan.clamped, plan.notes()[0])

    def test_the_page_break_lands_on_the_first_block_of_each_next_file(self) -> None:
        """需求第 3 条落成的下标：第二份从"它自己的来源行"起（那一行要在新页顶上）。"""
        plan = mg.merge_entries(self.entries(), ALL_ON)
        self.assertEqual(plan.layout.section_starts, (BLOCKS_A + 1,))

    def test_no_break_before_the_first_file(self) -> None:
        """第一条分组之前换页 = 首页是白的（前置页那两处分页不归这个开关管，见 builder）。"""
        plan = mg.merge_entries(self.entries(), ALL_ON)
        self.assertNotIn(0, plan.layout.section_starts)

    def test_source_labels_are_plain_italic_lines_named_after_the_file(self) -> None:
        """需求第 4 条：默认开着，且**是正文不是标题**——挂了标题就进了目录。"""
        plan = mg.merge_entries(self.entries(), ALL_ON)
        self.assertEqual([label_text(block) for block in plan.parsed.blocks
                          if label_text(block).startswith(mg.SOURCE_LABEL_PREFIX)],
                         ["来源：甲.md", "来源：乙.md"])
        first = plan.parsed.blocks[0]
        self.assertTrue(first.inline[0].italic)
        self.assertEqual(headings([block for block in plan.parsed.blocks
                                   if label_text(block).startswith("来源：")]), [])

    def test_source_label_off_writes_no_line_at_all(self) -> None:
        plan = mg.merge_entries(self.entries(), mg.MergeSwitches(demote=True, page_break=True))
        self.assertNotIn("来源：", [label_text(block) for block in plan.parsed.blocks])

    def test_toc_levels_follow_the_demotion(self) -> None:
        """降过级就收到 1-4：原来的 ### 现在在第 4 级，照默认收到 3 会把它整个漏掉。"""
        self.assertEqual(mg.merge_entries(self.entries(), ALL_ON).layout.toc_levels, "1-4")
        one = mg.merge_entries(self.entries((self.a, self.pa)), ALL_ON)
        self.assertEqual(one.layout.toc_levels, db.TOC_LEVELS_DEFAULT)

    def test_the_cover_says_title_date_and_file_count(self) -> None:
        """需求第 6 条那三格。日期写死中文格式，不跟着 Word 的区域设置换样子。"""
        cover = mg.merge_entries(self.entries(), ALL_ON,
                                 title="20kV 配网材料汇总",
                                 now=datetime(2026, 10, 15)).layout.cover
        self.assertEqual((cover.title, cover.date_text, cover.file_count_text),
                         ("20kV 配网材料汇总", "2026年10月15日", "共 2 个文件"))

    def test_the_cover_title_defaults_to_the_batch_label(self) -> None:
        plan = mg.merge_entries(self.entries(), ALL_ON)
        self.assertEqual(plan.layout.cover.title, "合并稿")

    def test_switches_off_leave_the_tree_exactly_as_the_files_are(self) -> None:
        """全关 → ``layout=None``、块树就是两份稿子首尾相接：这一路不与"没有这一层"分叉。"""
        plan = mg.merge_entries(self.entries(), mg.PLAIN)
        self.assertIsNone(plan.layout)
        self.assertEqual(plan.parsed.blocks, self.pa.blocks + self.pb.blocks)
        self.assertEqual(headings(plan.parsed.blocks), [1, 2, 6, 3, 1, 2])

    def test_warnings_are_concatenated_and_numeric_stats_are_summed(self) -> None:
        """需求第 7 条：复用 md_parser 的那三格纯数据，并树仍然只是串块 + 累加计数。"""
        first = mp.ParseResult([mp.Heading(1, [mp.TextRun("甲")])], ["第一条提醒"],
                               {"blocks": 1, "tables": 2, "encoding": "utf-8"})
        second = mp.ParseResult([mp.Heading(2, [mp.TextRun("乙")])], ["第二条提醒"],
                                {"blocks": 1, "tables": 5, "encoding": "gbk"})
        plan = mg.merge_entries([(Path("a.md"), first), (Path("b.md"), second)], mg.PLAIN)
        self.assertEqual(plan.parsed.warnings, ["第一条提醒", "第二条提醒"])
        self.assertEqual(plan.parsed.stats["tables"], 7)
        # 编码没有"并成一份之后是什么编码"这一说，所以整格不参与相加（逐份仍念在日志那一行）。
        self.assertNotIn("encoding", plan.parsed.stats)

    def test_an_empty_batch_is_an_empty_tree_instead_of_a_crash(self) -> None:
        plan = mg.merge_entries([], ALL_ON)
        self.assertEqual(plan.parsed.blocks, [])
        self.assertIsNone(plan.layout)

    def test_demotion_does_not_touch_the_caller(self) -> None:
        """降级是**新建一棵树**：解析结果被日志与逐份核对共用，就地改会把"甲册的 #"变成 ##。"""
        original = self.pa.blocks[0].level
        mg.merge_entries([(self.a, self.pa), (self.b, self.pb)], ALL_ON)
        self.assertEqual(self.pa.blocks[0].level, original)


class _DocxCase(unittest.TestCase):
    """第 2、3 层共用的一间作坊：两份 .md 落在临时目录，成品读回包里的 XML 再判。"""

    maxDiff = None

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="md2docx-merge-"))
        # 临时目录归临时目录：不清就会每跑一条用例在 %TEMP% 里落一个，跑测的人不会知道
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.a = self.tmp / "甲.md"
        self.b = self.tmp / "乙.md"
        self.a.write_text(MD_A, encoding="utf-8")
        self.b.write_text(MD_B, encoding="utf-8")
        self.out = self.tmp / "out"

    def builder(self, **options) -> db.DocxBuilder:
        return db.DocxBuilder(FormulaConverter(XSLT), db.BuilderOptions(**options))

    def write(self, plan: mg.MergePlan, target: Path) -> db.BuildReport:
        report = self.builder().build(plan.parsed, plan.layout)
        report.save(target)
        return report

    def member(self, target: Path, name: str) -> str:
        with zipfile.ZipFile(str(target)) as package:
            return package.read(name).decode("utf-8")

    def document(self, target: Path) -> str:
        return self.member(target, "word/document.xml")

    def paragraphs(self, target: Path) -> list:
        """一段一个小串。正则必须**自己认得自闭合**：非贪婪配上 "|/>" 会被段里第一个
        <w:i/> 或 <w:spacing/> 截断，那时"这一格里有没有正文"就永远判成没有。"""
        return re.findall(r"<w:p(?:\s[^>]*)?(?:/>|>.*?</w:p>)", self.document(target), re.S)


@needs_office
class BuilderFrontMatterTest(_DocxCase):
    """第 2 层：声明真的落进了包——目录域、自动刷新、分页符、样式归属。"""

    def build_merged(self, **changes) -> tuple[db.BuildReport, Path]:
        switches = mg.MergeSwitches(**{
            "demote": True, "page_break": True, "source_label": True, "toc": True, "cover": True,
            **changes})
        entries = [(self.a, mp.parse_file(self.a)), (self.b, mp.parse_file(self.b))]
        plan = mg.merge_entries(entries, switches, title="合并输出")
        target = self.tmp / "merged.docx"
        return self.write(plan, target), target

    def test_the_toc_is_a_compound_field_not_a_shortcut(self) -> None:
        """需求第 5 条：begin → instrText → separate → 占位 → end 五个 run 在**同一个 w:p** 里。"""
        _, target = self.build_merged()
        xml = self.document(target)
        self.assertEqual(xml.count('w:fldCharType="begin"'), 1)
        self.assertEqual(xml.count('w:fldCharType="separate"'), 1)
        self.assertEqual(xml.count('w:fldCharType="end"'), 1)
        self.assertEqual(len(re.findall(r"<w:fldSimple", xml)), 0)
        # 域代码前后各有一个空格：那是 Word 分词的边界，少了就把 "TOC \o" 读成一个词。
        self.assertIn('<w:instrText xml:space="preserve"> TOC \\o "1-4" \\h \\z \\u </w:instrText>',
                      xml)
        self.assertIn(db.TOC_PLACEHOLDER, xml)

    def test_the_document_asks_word_to_refresh_fields_on_open(self) -> None:
        """页码只有 Word 排版之后才数得出来，所以配套的是这一格而不是写死一个假页码。"""
        _, target = self.build_merged()
        settings = self.member(target, "word/settings.xml")
        self.assertIn('<w:updateFields w:val="true"/>', settings)
        # OOXML 对 settings 的子元素次序有要求：插到 w:compat 之后 Word 会报"文件已损坏"。
        self.assertLess(settings.index("w:updateFields"), settings.index("<w:compat"))

    def test_page_breaks_are_the_front_matter_plus_one_between_the_files(self) -> None:
        """封面后、目录后、两份之间——三处，且第一份之前不多一处（那会顶出一张白页）。"""
        _, target = self.build_merged()
        xml = self.document(target)
        self.assertEqual(xml.count('<w:br w:type="page"/>'), 3)
        paragraphs = self.paragraphs(target)
        break_only = [index for index, paragraph in enumerate(paragraphs)
                      if '<w:br w:type="page"/>' in paragraph and "<w:t" not in paragraph]
        self.assertEqual(len(break_only), 3)

    def test_the_front_pages_come_before_any_heading(self) -> None:
        _, target = self.build_merged()
        paragraphs = self.paragraphs(target)
        body = [index for index, paragraph in enumerate(paragraphs)
                if "来源：甲.md" in paragraph][0]
        cover = [index for index, paragraph in enumerate(paragraphs) if "合并输出" in paragraph][0]
        toc = [index for index, paragraph in enumerate(paragraphs)
               if db.TOC_PLACEHOLDER in paragraph][0]
        self.assertLess(cover, toc)
        self.assertLess(toc, body)

    def test_neither_the_cover_nor_the_toc_title_is_a_heading(self) -> None:
        """挂了 Heading 样式就会被 TOC 域收进去——"共 2 个文件"成一条目录项很难看。"""
        _, target = self.build_merged()
        for paragraph in self.paragraphs(target):
            if ("共 2 个文件" in paragraph or "2026年" in paragraph
                    or ">目录<" in paragraph or db.TOC_PLACEHOLDER in paragraph):
                self.assertNotIn("Heading", paragraph, paragraph[:200])

    def test_demoted_titles_land_on_heading_2_and_3(self) -> None:
        """需求第 2 条在 Word 侧的样子：# → Heading 2，## → Heading 3，**一级整档空着**。"""
        _, target = self.build_merged()
        xml = self.document(target)
        self.assertEqual(xml.count('<w:pStyle w:val="Heading1"/>'), 0)
        self.assertEqual(xml.count('<w:pStyle w:val="Heading2"/>'), 2)
        self.assertEqual(xml.count('<w:pStyle w:val="Heading3"/>'), 2)
        self.assertEqual(xml.count('<w:pStyle w:val="Heading4"/>'), 1)

    def test_the_source_line_is_italic_body_text(self) -> None:
        _, target = self.build_merged()
        hit = [paragraph for paragraph in self.paragraphs(target) if "来源：甲.md" in paragraph]
        self.assertEqual(len(hit), 1)
        self.assertIn("<w:i/>", hit[0])
        self.assertNotIn("Heading", hit[0])

    def test_the_receipt_says_page_numbers_are_word_s_job(self) -> None:
        """回执不能念"目录已生成"：文件里放的是域与占位，页码要 Word 算。"""
        report, _ = self.build_merged()
        self.assertIn(db.TOC_NOTE, report.notes)

    def test_the_breaks_stand_alone_from_the_switch_that_asks_for_them(self) -> None:
        """关了「插入分页符」，前置页那两处分页照旧在——否则目录与正文挤在同一页上。"""
        report, target = self.build_merged(page_break=False)
        xml = self.document(target)
        self.assertEqual(xml.count('<w:br w:type="page"/>'), 2)
        self.assertEqual(xml.count('w:fldCharType="begin"'), 1)
        self.assertEqual(report.blocks, BLOCKS_A + BLOCKS_B + 2)

    def test_toc_off_writes_no_field_and_no_update_flag(self) -> None:
        _, target = self.build_merged(toc=False)
        xml = self.document(target)
        self.assertEqual(xml.count("w:fldChar"), 0)
        self.assertNotIn("w:updateFields", self.member(target, "word/settings.xml"))

    def test_cover_off_writes_no_cover_lines(self) -> None:
        _, target = self.build_merged(cover=False)
        xml = self.document(target)
        self.assertNotIn("共 2 个文件", xml)
        self.assertNotIn("2026年", xml)
        # 目录仍要独占一页：前置页只剩目录那一份分页，加上两份之间那一处。
        self.assertEqual(xml.count('<w:br w:type="page"/>'), 2)

    def test_layout_none_writes_exactly_what_a_single_file_always_wrote(self) -> None:
        """第 3 层最硬的一条：单独输出那一路的 document.xml **逐字节相同**。"""
        parsed = mp.parse_file(self.a)
        merged_way = self.tmp / "plain-way.docx"
        self.write(mg.MergePlan(parsed, None), merged_way)
        direct = self.tmp / "direct.docx"
        self.builder().build(parsed).save(direct)
        self.assertEqual(self.document(merged_way), self.document(direct))
        self.assertEqual(self.member(merged_way, "word/settings.xml"),
                         self.member(direct, "word/settings.xml"))


@needs_office
class TasksWiringTest(_DocxCase):
    """第 3 层：面板上那四个勾真的作用在盘上的文件。"""

    def options_for(self, **changes):
        import gui.tasks as tasks

        base = dict(output_dir=self.out, merged=True, body_font="宋体",
                    body_size=12.0, heading_font="黑体")
        base.update(changes)
        return tasks.Options(**base)

    def run_batch(self, options, lines=None):
        import gui.tasks as tasks

        collected = lines if lines is not None else []
        result = tasks.convert_paths([self.a, self.b], options, collected.append,
                                     lambda percent, text: None)
        return result, collected

    def test_merge_switches_are_the_four_panel_checks_plus_the_demote_that_merging_implies(self) -> None:
        options = self.options_for(with_toc=True, page_break=False,
                                   source_label=True, cover_page=True)
        self.assertEqual(options.merge_switches(),
                         mg.MergeSwitches(demote=True, page_break=False, source_label=True,
                                          toc=True, cover=True))

    def test_separate_output_is_plain_no_matter_how_the_checks_are_ticked(self) -> None:
        """单独输出没有"合并选项"这一说：照勾翻就会出现每个文件都凭空降一级。"""
        options = self.options_for(merged=False)
        self.assertIs(options.merge_switches(), mg.PLAIN)

    def test_a_merged_batch_writes_one_docx_holding_the_field(self) -> None:
        options = self.options_for(merged=True, cover_page=True)
        result, _ = self.run_batch(options)
        written = sorted(path.name for path in self.out.glob("*.docx"))
        self.assertEqual(written, ["合并输出.docx"], result["receipt"])
        xml = self.document(self.out / "合并输出.docx")
        self.assertIn('<w:instrText xml:space="preserve"> TOC \\o "1-4" \\h \\z \\u </w:instrText>',
                      xml)
        self.assertIn("来源：甲.md", xml)
        self.assertIn("共 2 个文件", xml)
        self.assertIn('<w:updateFields w:val="true"/>',
                      self.member(self.out / "合并输出.docx", "word/settings.xml"))

    def test_the_merged_row_counts_the_files_it_swallowed(self) -> None:
        result, _ = self.run_batch(self.options_for())
        row = result["rows"][0]
        self.assertEqual((row[0], row[2]), ("合并输出", "2"))

    def test_the_clamped_note_reaches_the_report_instead_of_staying_in_the_layer(self) -> None:
        result, _ = self.run_batch(self.options_for())
        mentions = [text for _, text in result["problems"] if "第 6 级" in text]
        self.assertEqual(len(mentions), 1, result["problems"])
        self.assertIn("合并时说：", mentions[0])

    def test_the_toc_note_reaches_the_report(self) -> None:
        result, _ = self.run_batch(self.options_for())
        self.assertTrue([text for step, text in result["problems"]
                         if "生成时说：" in text and db.TOC_NOTE in text],
                        result["problems"])

    def test_separate_output_still_writes_per_file_documents_without_front_matter(self) -> None:
        result, _ = self.run_batch(self.options_for(merged=False))
        written = sorted(path.name for path in self.out.glob("*.docx"))
        self.assertEqual(set(written), {"甲.docx", "乙.docx"}, result["receipt"])
        for name in written:
            xml = self.document(self.out / name)
            self.assertEqual(xml.count("w:fldChar"), 0)
            self.assertEqual(xml.count('<w:br w:type="page"/>'), 0)
            self.assertNotIn("来源：", xml)
            self.assertNotIn("共 2 个文件", xml)
            self.assertIn('<w:pStyle w:val="Heading1"/>', xml)
            self.assertNotIn("w:updateFields", self.member(self.out / name, "word/settings.xml"))

    def test_each_group_gets_its_own_cover_named_after_the_group(self) -> None:
        """分组输出是"每组各并一次"，封面标题跟着组名走，不是跟着第一个文件走。"""
        options = self.options_for(merged=False, grouped=True,
                                   groups=[("甲组", [self.a]), ("乙组", [self.b])],
                                   cover_page=True)
        result, _ = self.run_batch(options)
        self.assertEqual(sorted(path.name for path in self.out.glob("*.docx")),
                         ["1-甲组.docx", "2-乙组.docx"], result["receipt"])
        self.assertIn("甲组", self.document(self.out / "1-甲组.docx"))
        self.assertIn("共 1 个文件", self.document(self.out / "1-甲组.docx"))

    def test_group_of_one_is_not_demoted_while_a_group_of_two_is(self) -> None:
        options = self.options_for(merged=False, grouped=True,
                                   groups=[("单文件组", [self.a]),
                                           ("两文件组", [self.a, self.b])])
        self.run_batch(options)
        single = self.document(self.out / "1-单文件组.docx")
        both = self.document(self.out / "2-两文件组.docx")
        self.assertIn('<w:pStyle w:val="Heading1"/>', single)
        self.assertNotIn('<w:pStyle w:val="Heading1"/>', both)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
