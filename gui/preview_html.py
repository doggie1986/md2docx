# -*- coding: utf-8 -*-
"""把 Markdown 近似渲染成 HTML，只给中间那块预览区看。

**这里不生成 .docx**：公式走 :mod:`gui.math_html` 摆成 Qt 富文本画得出来的那几样（上下标、
上下两格的分式、``√(…)``），为的是"看一眼这条公式长什么样、表格有几列、标题是第几级"，
落进 Word 之后是什么样仍然由 :mod:`docx_builder` 说话。

块结构直接借 :mod:`md_parser`（命令行那条链用的就是它），所以预览和真转换
对"这段算不算公式"的判断永远一致——写第二套解析迟早会分叉，那一叉就是
"预览里没有、文件里却有"这种最难查的错。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from html import escape
from pathlib import Path
from typing import Any, Iterable, Sequence

import md_parser as mp
from gui import math_html

__all__ = ["NOTICE", "PreviewResult", "render_file", "render_blocks"]

#: 预览里图片那一格写什么。**预览不画 bitmap**（这一层只管"这份稿子里有什么"），
#: 而 Word 里它是真嵌进去的图片——所以这一句不能照抄 Word 那句占位：那句说的是"没嵌进去"，
#: 抄过来在预览里就成了假话。以前正是共用 :data:`docx_builder.IMAGE_PLACEHOLDER`。
IMAGE_MARK = "【图片：%s】"

#: 这一句摆在预览顶部，说的是"这里和 Word 里差在哪"。差的那几处是量出来的，
#: 不是含糊过去的：分式在块级里是上下两格带一条分数线、在行内只能折成一行，根号没有那一道眉，
#: 而语法体检没过的公式原样退回 LaTeX。
NOTICE = ("这是近似预览：公式里上下标、分式（块级的是上下两格带一条分数线）、根号（写成 √( … )）"
          "都照能画的摆，行内公式里的分式与矩阵只能折成一行读；"
          "转进 Word 之后它们都是可编辑的公式对象。"
          "认不出的写法会退回 LaTeX 原文并涂上警告色——那一处在 Word 里也是原样文本。"
          "图片在这里只写名字，进了 Word 才嵌成真的图。"
          "字号与分页以转换时的样式设置为准。")

#: 对齐写进格子的**三个词**（Markdown 的 `:---` / `:---:` / `---:` 就只能问出这三档）。
#:
#: 走 ``align`` 属性而不是 ``style="text-align:…"``——这一条是量出来的，不是 Style Guide 的口味：
#: Qt 的富文本里，一个元素**带了内联 ``style`` 就不再合并样式表里那一份声明**，于是
#: ``th, td { border: 1px solid …; padding: 3px 7px; }`` 整格没了，一张带对齐的表画出来
#: **一条格线都没有**（实测：同一张表 345 个格线色像素 → 0；逐格加样式则 345→288→201，
#: 掉的正好是加了样式那几格的份额）。``align`` 属性两样都保住（452 个像素与不写对齐时相同，
#: 而墨点确实跑到右半）。``justify`` 不是 ``align`` 认的值，也不可能是 Markdown 问出来的，
#: 所以按"认不出就不写"处理（宁可按默认摆，也不为它单独换回那种会吃掉格线的写法）。
_ALIGN_WORDS = ("left", "center", "right")


@dataclass
class PreviewResult:
    """一次预览的交回：要么有 html，要么有 problem，两者不会同时为空。"""

    html: str = ""
    stats: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    problem: str = ""

    @property
    def ok(self) -> bool:
        return not self.problem


def render_file(path: Path) -> PreviewResult:
    """读一个文件并渲染。读不动 / 解析不动都只交回一句话，不抛——预览区坏了不该把主窗口带崩。"""
    if not path.is_file():
        return PreviewResult(problem=f"文件已经不在了：{path}")
    try:
        parsed = mp.parse_file(path)
    except (OSError, UnicodeDecodeError, ValueError, mp.MarkdownSyntaxError) as exc:
        detail = str(exc) or exc.__class__.__name__
        return PreviewResult(problem=f"读不动或解析不动：{detail}")
    body = _blocks(parsed.blocks)
    return PreviewResult(html=_wrap(path, body), stats=dict(parsed.stats),
                         warnings=list(parsed.warnings))


def render_blocks(blocks: Iterable[Any]) -> str:
    """给一段已经解析好的块渲染 HTML（没有文件也能用，测试走这一条）。"""
    return _blocks(list(blocks))


# ---------------------------------------------------------------- 块级


def _wrap(path: Path, body: str) -> str:
    return (
        "<div class=\"notice\">{notice}</div>"
        "<h3>{name}</h3>"
        "<p class=\"notice\">{source}</p>"
        "{body}"
    ).format(
        notice=escape(NOTICE),
        name=escape(path.name),
        source=escape(str(path)),
        body=body,
    )


def _blocks(blocks: Sequence[Any]) -> str:
    out: list[str] = []
    for block in blocks:
        out.append(_block(block))
    return "".join(out)


def _block(block: Any) -> str:
    if isinstance(block, mp.Heading):
        level = min(max(block.level, 1), 6)
        return "<h%d>%s</h%d>" % (level, _inline(block.inline), level)
    if isinstance(block, mp.Paragraph):
        return "<p>%s</p>" % _inline(block.inline)
    if isinstance(block, mp.BlockMath):
        return "<div class=\"math math-block\">%s</div>" % math_html.render_math(block.latex, display=True)
    if isinstance(block, mp.CodeBlock):
        lang = " data-lang=\"%s\"" % escape(block.lang) if block.lang else ""
        return "<pre%s><code>%s</code></pre>" % (lang, escape(block.text))
    if isinstance(block, mp.Table):
        return _table(block)
    if isinstance(block, mp.QuoteBlock):
        return "<blockquote>%s</blockquote>" % _blocks(block.blocks)
    if isinstance(block, mp.ListBlock):
        return _list(block)
    if isinstance(block, mp.ThematicBreak):
        return "<hr>"
    # 认不出的块：把它的类名念出来而不是悄悄丢掉——丢掉就变成"预览里少了一段"，
    # 而解析器明明是认下它了的。
    return "<p class=\"placeholder\">（预览还不认识这一种块：%s）</p>" % escape(type(block).__name__)


def _table(table: mp.Table) -> str:
    rows: list[str] = []
    aligns = list(table.aligns)
    for index, cells in enumerate(table.rows):
        tag = "th" if (index == 0 and table.header_present) else "td"
        boxes = []
        for column, cell in enumerate(cells):
            attr = _align_attr(aligns[column] if column < len(aligns) else "")
            boxes.append("<%s%s>%s</%s>" % (tag, attr, _inline(cell), tag))
        rows.append("<tr>%s</tr>" % "".join(boxes))
    return "<table>%s</table>" % "".join(rows)


def _align_attr(value: str) -> str:
    word = (value or "").strip().lower()
    if word in _ALIGN_WORDS:
        return " align=\"%s\"" % word
    return ""


def _list(listing: mp.ListBlock) -> str:
    tag = "ol" if listing.ordered else "ul"
    start = " start=\"%d\"" % listing.start if listing.ordered and listing.start not in (0, 1) else ""
    boxes = []
    for item in listing.items:
        inner = _inline(item.inline)
        if item.children:
            # 递归而不是只走一层：解析器认下几层就画几层，否则第三层在预览里凭空消失，
            # 而 Word 那一条是写到第 N 层的（docx_builder._write_list_item 自己递归）。
            # 子列表是 ol 还是 ul 由**孩子自己**那一层说（item.ordered 是父行的记号，
            # 拿它当子层的记号会把圆点演成序号），起始号永远从 1 起——与 Word 同一口径。
            inner += _list(mp.ListBlock(item.children[0].ordered, item.children))
        boxes.append("<li>%s</li>" % inner)
    return "<%s%s>%s</%s>" % (tag, start, "".join(boxes), tag)


# ---------------------------------------------------------------- 行内


def _inline(runs: Iterable[Any]) -> str:
    return "".join(_inline_one(run) for run in runs)


def _inline_one(run: Any) -> str:
    if isinstance(run, mp.MathRun):
        # 行内一律 ``display=False``：Qt 把行内表格提成块级、一句话就被劈开（实测），
        # 所以行内分式在预览里是 ``(a)/(b)``，而 Word 里仍是上下两格——这一句写在 NOTICE 里。
        return "<span class=\"math\">%s</span>" % math_html.render_math(run.latex)
    if isinstance(run, mp.CodeSpan):
        return "<code>%s</code>" % escape(run.text)
    if isinstance(run, mp.ImageRef):
        return "<span class=\"placeholder\">%s</span>" % escape(
            IMAGE_MARK % (run.alt or run.url))
    if isinstance(run, mp.TextRun):
        text = escape(run.text)
        if run.bold:
            text = "<strong>%s</strong>" % text
        if run.italic:
            text = "<em>%s</em>" % text
        if run.link:
            text = "<a href=\"%s\">%s</a>" % (escape(run.link, quote=True), text)
        return text
    return escape(str(getattr(run, "text", run)))
