# -*- coding: utf-8 -*-
"""合并稿的策略层：把"勾了哪几个开关、这几个文件谁先谁后"翻译成一份纯声明。

边界写死在这里：这一层**认 Markdown 与文件名，不认 Word**——它定标题降几级、
哪一块前面要换页、封面那三行写什么字，然后交回一棵改过的块树加一份
:class:`docx_builder.MergeLayout`；往 .docx 上落笔的仍然只有 builder。
反过来 builder 也不该自己判"第二个文件从哪一块开始"，所以
:data:`docx_builder.MergeLayout.section_starts` 装的是块下标而不是文件名——
这一层负责把文件名换成下标。

五条不是显而易见的口径：

* **只有真合并时才降标题级**（只剩一个文件就不降）：单独一份稿子自己就是完整的一级结构，
  整篇降一级等于凭空让出一层标题，而"这一份为什么没有一级标题"没人答得上来。
* **降到第 6 级就停**，停了几条写在回执里：Markdown 只有六级，``######`` 再降就没处放。
  一声不吭地按第 6 级排下去会被读成"降级没生效"。
* **目录收哪几级跟着降级走**（降过就收 ``1-4``）：合并稿里原来的 ``#`` 落在第 2 级，
  域代码照默认写 ``\\o "1-3"`` 就会把它底下的 ``###``（现在第 4 级）整个漏掉——
  目录少几条是看不出来的错，比多几条糟。
* **来源标注与封面那几行不进目录**：来源行是一段普通斜体文字，封面三行 builder 刻意不挂
  Heading 样式。它们一挂上样式就会被 TOC 域收进去，"共 3 个文件"成了一条目录项很难看。
* **开关全关时交回 ``layout=None``**：那一路的产物与"没有合并这一层"时逐字节相同，
  命令行、单独输出、以及"合并但四个勾一个都不打"都走这一条。
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Optional, Sequence, Tuple

import docx_builder as db
import md_parser as mp

__all__ = ["MergeSwitches", "PLAIN", "MergePlan", "MergeEntry",
           "SOURCE_LABEL_PREFIX", "demote_block", "demote_blocks",
           "source_label", "cover_for", "merge_entries"]

#: 标题降几级。需求写的是"# 降为 ##、## 降为 ###"，所以是一级而不是两级。
HEADING_DEMOTE_STEP = 1

#: Markdown 有六级标题，Word 的 Heading 样式也只到六级（超过就没处放，只能钉在最后一级）。
MAX_HEADING_LEVEL = max(db.HEADING_STYLES)

#: 降过级之后目录该收到第几级：原来最深用到的 ``###`` 现在落在第 4 级。
TOC_LEVELS_AFTER_DEMOTE = "1-4"

#: 来源标注那一行的开头。它是一段正文而不是样式，所以"哪算一行"只有这一处定义。
SOURCE_LABEL_PREFIX = "来源："

#: 降级撞到六级上限时念的那一句（``N`` 是撞上的条数）。
CLAMPED_NOTE = ("标题降级时有 %d 条已经在第 %d 级、没法再降一级，"
                "就按第 %d 级排着（Markdown 只有六级标题）。")

#: 一个待合并的文件：``(文件路径, 解析结果)``。路径是给来源标注和封面用的，
#: 解析结果给块树——**两个一起传**，只给块树就没处念"这几块来自哪一份"。
MergeEntry = Tuple[Path, mp.ParseResult]


@dataclass(frozen=True)
class MergeSwitches:
    """这一次合并勾了哪几样。全部由 :meth:`gui.tasks.Options.merge_switches` 从面板取值翻译过来，
    界面上那颗勾本身不住在这一层（这一层不 import Qt，也不 import 界面）。

    ``demote`` 不是面板上的开关——合并这件事本身就意味着要降一级，
    所以它是"这一份是不是合并稿"的声明，而不是人勾的那一格。
    """

    demote: bool = False
    page_break: bool = False
    source_label: bool = False
    toc: bool = False
    cover: bool = False

    def any_of(self) -> bool:
        """五格有没有一格是开着的——全关时整层都是惰性的，一个字节都不该动。"""
        return self.demote or self.page_break or self.source_label or self.toc or self.cover


#: 什么都不做的那一份：单独输出、命令行那一路用它，保证产物与"没有这一层"时逐字节相同。
PLAIN = MergeSwitches()


@dataclass(frozen=True)
class MergePlan:
    """一次合并的结果：交给 ``builder.build()`` 的那棵块树 + 那份排版声明。

    ``layout=None`` 就是"这份稿子没有前置页、也没有文件之间的换页"，builder 因此
    一个字都不多写。``clamped`` 是降级撞上限的条数，由 :meth:`notes` 换成人话。
    """

    parsed: mp.ParseResult
    layout: Optional[db.MergeLayout] = None
    clamped: int = 0
    notices: Sequence[str] = field(default_factory=tuple)

    def notes(self) -> list[str]:
        """要念在报告里的那几句（降级撞了几条、这一层主动说明的那些）。"""
        out = []
        if self.clamped:
            out.append(CLAMPED_NOTE % (self.clamped, MAX_HEADING_LEVEL, MAX_HEADING_LEVEL))
        out.extend(self.notices)
        return out


def demote_block(block: mp.Block, step: int = HEADING_DEMOTE_STEP) -> tuple[mp.Block, int]:
    """一个块降一级标题，交回 ``(新块, 撞上限的条数)``。

    只往 :class:`md_parser.QuoteBlock` 里递归——那是唯一"块套块"的地方。表格与列表项里装的
    是行内元素（文字/公式/代码/图片），根本没有标题，去那儿找等于白跑一遍。
    """
    if isinstance(block, mp.Heading):
        target = block.level + step
        if target > MAX_HEADING_LEVEL:
            return block, 1
        return replace(block, level=target), 0
    if isinstance(block, mp.QuoteBlock):
        blocks = []
        clamped = 0
        for inner in block.blocks:
            new_inner, hit = demote_block(inner, step)
            blocks.append(new_inner)
            clamped += hit
        if not clamped and all(inner is kept for inner, kept in zip(block.blocks, blocks)):
            return block, 0
        return replace(block, blocks=blocks), clamped
    return block, 0


def demote_blocks(blocks: Sequence[mp.Block]) -> tuple[list[mp.Block], int]:
    """整棵块树降一级，交回新块表与撞上限的条数（不改传进来的那一份）。"""
    out = []
    clamped = 0
    for block in blocks:
        new_block, hit = demote_block(block)
        out.append(new_block)
        clamped += hit
    return out, clamped


def source_label(path: Path) -> mp.Paragraph:
    """``来源：xxx.md`` 那一行：一段斜体正文，**不是标题**（挂了标题就进了目录）。"""
    return mp.Paragraph([mp.TextRun(SOURCE_LABEL_PREFIX + path.name, italic=True)])


def cover_for(title: str, file_count: int, now: Optional[datetime] = None) -> db.CoverPage:
    """封面那三行的文本。日期与文件数在这一层算——builder 只管摆，不认"几个文件"这种概念。

    日期用中文年月日而不是 ``2026-10-15``：封面是给业务方看的，那一行跟着 Word 的
    区域设置走会换样子，写死中文才每张封面都长一个样。
    """
    moment = now or datetime.now()
    return db.CoverPage(
        title=title,
        date_text="%d年%d月%d日" % (moment.year, moment.month, moment.day),
        file_count_text="共 %d 个文件" % file_count,
    )


def merge_entries(entries: Sequence[MergeEntry], switches: MergeSwitches,
                  title: str = "", now: Optional[datetime] = None) -> MergePlan:
    """把 ``[(路径, 解析结果), …]`` 并成一棵块树，顺带定好前置页。

    次序**照传进来的顺序**（左侧列表的勾选顺序就是合并顺序，那一层已经在
    :meth:`Options.output_plan` 里排好了）；这里不重排、也不去重。

    降级只在一个文件以上时做：只有一份稿子时它是完整独立的一级结构，降了级反而少一层。
    """
    if not entries:
        return MergePlan(mp.ParseResult([], [], {}), None, 0)

    demote = switches.demote and len(entries) > 1
    blocks: list = []
    warnings: list[str] = []
    stats: dict = {}
    section_starts: list[int] = []
    clamped = 0

    for position, (path, parsed) in enumerate(entries):
        # 换页点记在"这一节的第一个块"的下标上，且在来源行之前——那一行要在新页的顶上。
        if position and switches.page_break:
            section_starts.append(len(blocks))
        if switches.source_label:
            blocks.append(source_label(path))
        section_blocks = list(parsed.blocks)
        if demote:
            section_blocks, hit = demote_blocks(section_blocks)
            clamped += hit
        blocks.extend(section_blocks)
        warnings.extend(parsed.warnings)
        for key, value in parsed.stats.items():
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                continue
            stats[key] = stats.get(key, 0) + value

    layout = None
    cover = cover_for(title or "合并稿", len(entries), now) if switches.cover else None
    toc_heading = db.TOC_HEADING_DEFAULT if switches.toc else ""
    if cover is not None or toc_heading or section_starts:
        layout = db.MergeLayout(
            cover=cover,
            toc_heading=toc_heading,
            toc_levels=TOC_LEVELS_AFTER_DEMOTE if demote else db.TOC_LEVELS_DEFAULT,
            section_starts=tuple(section_starts),
        )
    return MergePlan(mp.ParseResult(blocks, warnings, stats), layout, clamped)
