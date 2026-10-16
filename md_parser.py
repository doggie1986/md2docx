# -*- coding: utf-8 -*-
"""Markdown 解析层：文件文本 → 结构化块列表。

这一层的边界：**只认识 Markdown 与 LaTeX 定界符，不认识 Word、不认识 MathML**。
产出的数据结构（``Block`` / ``Inline``）是它与 docx_builder 之间的合同。

为什么自己写解析器而不用 markdown / mistune —— 两个硬理由：

1. **公式必须一段都不被 Markdown 语法碰**。``$a*b$`` 里的星号会被当成强调，
   ``\\(x_n\\)`` 里的下划线会被当成斜体；先把公式屏蔽掉再解析才是安全的。
2. **表格切列会被公式里的竖线劈错**。``| $|x|$ |`` 这种单元格（绝对值、条件概率的 ``\\mid``）
   在朴素按 ``|`` 切分的解析器里一定会多切出一列。

所以这里的次序是死的：**找公式区间 → 用哨兵屏蔽 → 才切表格列 → 每个单元格再还原成行内片段**。

块级与行内的判定口径（业务方 2026-10-14 确认）：

* 独占一行的 ``$$…$$``（含跨行写法）→ :class:`BlockMath`，居中排版。
* 夹在文字里的 ``$$…$$`` 与 ``$…$`` → **按行内**处理，不劈开那句话。
* ``\\[…\\]`` 本身在 TeX 里就是显示式 → 行内片段但 ``display=True``，
  由 docx_builder 把它单独起一段并居中；``\\(…\\)`` 是行内。
* 表格里出现显示式公式 → 该公式在**那个单元格内**另起一段并居中（``display=True`` 就是这一格的信号）。

行内强调、行内代码、链接与图片也在这切成片段（``TextRun`` / ``MathRun`` / ``CodeSpan`` / ``ImageRef``）：
链接的地址进 ``TextRun.link``（Word 里画成超链接，方括号和地址都不出现在正文上），
图片切成 ``ImageRef``；**这一层仍然不读盘上的任何图片文件**，只是把"去哪找这张图"
（源文件所在目录）盖在节点上，让校验与嵌图两边共用同一条规则（:func:`image_target`）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence, Union
from urllib.parse import unquote

# --------------------------------------------------------------------------
# 哨兵字符：私有区码位，正常 Markdown 文本里不会出现；屏蔽公式/代码段/转义竖线用。
# 它们只在本模块内部活着，出去之前一定会被还原（_unmask 之后还残留就报错，见文件末尾）。
# --------------------------------------------------------------------------
MASK_OPEN = "\uF000"
MASK_CLOSE = "\uF001"
PIPE_MASK = "\uF002"

#: 这些字符前面加反斜杠是"就要这个字符本身"，不是公式定界符的一部分
ESCAPABLE = set("\\`*_{}[]()#+-.!|~<>$")

# --------------------------------------------------------------------------
# 数据模型：块
# --------------------------------------------------------------------------


@dataclass
class TextRun:
    """一段普通文字。``text`` 里的 ``\\n`` 是 Markdown 的硬换行，由 builder 变成 ``w:br``。

    ``link`` 非空就是 ``[文字](地址)`` 里的那段文字，由 builder 画成可点的超链接
    （把方括号和地址原样写进 Word 是噪音，而丢掉地址又是丢信息）。
    """

    text: str
    bold: bool = False
    italic: bool = False
    link: str = ""


@dataclass
class MathRun:
    """一个行内公式区间。``latex`` 是**去掉定界符之后**的原文，一个字符都不加工。"""

    latex: str
    display: bool = False


@dataclass
class CodeSpan:
    """行内代码 ``…``：原样保留，不做公式转换。"""

    text: str


@dataclass
class ImageRef:
    """``![替代文字](地址)``：地址 + 这份 .md 所在的目录，嵌不嵌由下游决定。

    为什么不悄悄丢掉：图片丢了没人知道，而那句占位里写着文件名和"为什么没成"，
    人一眼能找到那张图去哪补。嵌不嵌由 :mod:`docx_builder` 判（这一层不读图片字节）。

    ``base`` 是**这份 .md 自己的目录**，由 :func:`parse_file` 盖上去：合并稿把几个文件
    的块摞进同一棵树，各自的图得按各自文件的目录找——按第一个文件的目录找，第二份稿子的
    图会全数"没找到"。从字符串解析（:func:`parse`）时它是 ``None``，由调用方给一个目录。
    """

    alt: str
    url: str
    base: Optional[Path] = None


Inline = Union[TextRun, MathRun, CodeSpan, ImageRef]


@dataclass
class Heading:
    level: int
    inline: list[Inline]
    line: int = 0


@dataclass
class Paragraph:
    inline: list[Inline]
    line: int = 0


@dataclass
class BlockMath:
    """独占一段的显示式公式。"""

    latex: str
    line: int = 0
    display: bool = True


@dataclass
class CodeBlock:
    """围栏代码块：内容原样，里面的 ``$`` 不当公式。"""

    text: str
    lang: str = ""
    line: int = 0


@dataclass
class ListItem:
    """一个列表条目。``ordered`` 记的是**这一层**是有序还是无序：
    嵌套层的记号跟顶层可以不一样（圆点里套 1. 2.），只存 ListBlock.ordered 的话
    下一层就没人知道该画圆点还是画序号。
    """

    inline: list[Inline]
    children: list["ListItem"] = field(default_factory=list)
    ordered: bool = False


@dataclass
class ListBlock:
    ordered: bool
    items: list[ListItem]
    start: int = 1
    line: int = 0


@dataclass
class QuoteBlock:
    blocks: list["Block"]
    line: int = 0


@dataclass
class ThematicBreak:
    line: int = 0


@dataclass
class Table:
    """Markdown 表格。``rows`` 的每一格是**行内片段列表**，这就是"表格里嵌公式"的落点。

    ``aligns`` 与表头逐列对应，取值 ``left`` / ``center`` / ``right`` / ``none``。
    """

    rows: list[list[list[Inline]]]
    aligns: list[str]
    header_present: bool = True
    line: int = 0

    @property
    def column_count(self) -> int:
        return len(self.aligns) if self.aligns else (len(self.rows[0]) if self.rows else 0)


Block = Union[Heading, Paragraph, BlockMath, CodeBlock, ListBlock, QuoteBlock, ThematicBreak, Table]


# --------------------------------------------------------------------------
# 数据模型：解析结果
# --------------------------------------------------------------------------


@dataclass
class Issue:
    """一条**能定位到行**的问题：校验那一层要按档位上色的结构化记录。

    与 ``warnings`` 的分工：``warnings`` 是给人读的那句中文（一直是、继续是），
    ``Issue`` 是同一件事的可编程形状——``code`` 说"哪一类"、``line`` 说"第几行"。
    只在解析器本来就知道是哪一行的地方登记（目前就是表格列数那两条）：
    给一条没有行号可指的问题硬凑 ``line=0``，会被读成"查过了、在第 0 行"。
    """

    code: str
    line: int
    message: str


@dataclass
class ParseResult:
    """解析回执。``warnings`` 是给人读中文句子，``stats`` 是给报告用的数。"""

    blocks: list[Block]
    warnings: list[str] = field(default_factory=list)
    stats: dict[str, int] = field(default_factory=dict)
    issues: list[Issue] = field(default_factory=list)

    def formulas(self) -> list[MathRun | BlockMath]:
        """按出现顺序把全文的公式都收出来（自检与报告用）。"""
        found: list[MathRun | BlockMath] = []

        def walk_inlines(nodes: Iterable[Inline]) -> None:
            for node in nodes:
                if isinstance(node, MathRun):
                    found.append(node)

        def walk_blocks(blocks: Sequence[Block]) -> None:
            for block in blocks:
                if isinstance(block, BlockMath):
                    found.append(block)
                elif isinstance(block, (Heading, Paragraph)):
                    walk_inlines(block.inline)
                elif isinstance(block, Table):
                    for row in block.rows:
                        for cell in row:
                            walk_inlines(cell)
                elif isinstance(block, ListBlock):
                    _walk_list_items(block.items, walk_inlines)
                elif isinstance(block, QuoteBlock):
                    walk_blocks(block.blocks)

        walk_blocks(self.blocks)
        return found

    def annotated(self, *kinds: type) -> list[tuple[Any, int]]:
        """按出现顺序交回 ``(节点, 所在行的行号)``，只收 ``kinds`` 点名的那几种节点。

        行内片段（公式、图片）自己不带行号——它们是块里的一小段，而校验报错要念得出
        "第几行"，所以由所属块的 ``line`` 补上；``BlockMath`` 自带行号，照它自己那条。
        这块遍历不另起第二套：写第二套"哪些块里有行内片段"迟早会和解析分叉，
        那一叉就是"校验说没有、文件里却有"这种最难查的错（同 :mod:`gui.preview_html` 那条理由）。
        """
        found: list[tuple[Any, int]] = []

        def walk_inlines(nodes: Iterable[Inline], line: int) -> None:
            for node in nodes:
                if isinstance(node, kinds):
                    found.append((node, line))

        def walk_blocks(blocks: Sequence[Block], inherited: int) -> None:
            for block in blocks:
                line = block.line or inherited
                if isinstance(block, kinds):
                    found.append((block, line))
                if isinstance(block, (Heading, Paragraph)):
                    walk_inlines(block.inline, line)
                elif isinstance(block, Table):
                    for row in block.rows:
                        for cell in row:
                            walk_inlines(cell, line)
                elif isinstance(block, ListBlock):
                    _walk_list_items(block.items, walk_inlines, line)
                elif isinstance(block, QuoteBlock):
                    walk_blocks(block.blocks, line)

        walk_blocks(self.blocks, 0)
        return found


def _walk_list_items(items: Sequence[ListItem], visit, *args) -> None:
    for item in items:
        visit(item.inline, *args)
        if item.children:
            _walk_list_items(item.children, visit, *args)


class MarkdownSyntaxError(ValueError):
    """文件结构坏到没法安全继续（目前只有"围栏代码块没闭合"这一种会抛）。"""


# ==========================================================================
# 第一步：找公式与代码段的区间
# ==========================================================================

_MATH_FORMS: tuple[tuple[str, str, bool], ...] = (
    ("$$", "$$", False),
    ("$", "$", False),
    ("\\(", "\\)", False),
    ("\\[", "\\]", True),
)


def _find_close(text: str, closer: str, start: int) -> int:
    """从 ``start`` 找闭合定界符，跳过它前面的反斜杠转义；找不到返回 -1。"""
    index = start
    while True:
        index = text.find(closer, index)
        if index < 0:
            return -1
        if _escaped(text, index):
            index += 1
            continue
        return index


def _escaped(text: str, index: int) -> bool:
    """位置 index 上的字符前面是不是有**奇数个**反斜杠（奇数才是被转义的）。"""
    slashes = 0
    i = index - 1
    while i >= 0 and text[i] == "\\":
        slashes += 1
        i -= 1
    return slashes % 2 == 1


def _code_spans(text: str) -> list[tuple[int, int]]:
    """行内代码 ``…`` 的区间。代码段里面不找公式。"""
    spans: list[tuple[int, int]] = []
    i = 0
    n = len(text)
    while i < n:
        if text[i] != "`" or _escaped(text, i):
            i += 1
            continue
        run = 0
        while i + run < n and text[i + run] == "`":
            run += 1
        close = text.find("`" * run, i + run)
        if close < 0:
            i += run
            continue
        spans.append((i, close + run))
        i = close + run
    return spans


def _inside(spans: Sequence[tuple[int, int]], start: int) -> bool:
    return any(a <= start < b for a, b in spans)


def _math_spans(text: str) -> list[tuple[int, int, str, bool]]:
    """找出不重叠的公式区间 ``(起, 止, latex, display)``。

    四条取舍，都是"宁可不当公式，也不要错当公式"：

    * 空内容（``$`` 紧挨 ``$$``）不当公式。
    * 单 ``$`` 里出现换行不当公式 —— 段落里两个孤立的 ``$`` 更可能是货币符号。
    * ``$5 和 $6`` 这种货币写法不当公式：闭合的 ``$`` 紧跟数字就否掉（pandoc 同一条规则）。
    * 定界符不成对就当没有，原文照留（后面报表里会念"这一段没被认成公式"）。
    """
    code = _code_spans(text)
    spans: list[tuple[int, int, str, bool]] = []
    i = 0
    n = len(text)
    while i < n:
        if _inside(code, i):
            i = next(b for a, b in code if a <= i < b)
            continue
        matched = False
        for opener, closer, display in _MATH_FORMS:
            if not text.startswith(opener, i) or _escaped(text, i):
                continue
            if closer == "$" and text.startswith("$$", i):
                continue  # 单美元不抢双美元的开头，$$ 那一档已经在前面试过
            body_start = i + len(opener)
            end = _find_close(text, closer, body_start)
            if end < 0:
                continue
            if closer == "$" and text.find("\n", body_start, end) >= 0:
                continue
            body = text[body_start:end]
            if not body.strip():
                continue
            if opener == "$" and closer == "$" and _digit_after(text, end):
                continue
            if opener == "$" and body[0].isspace() and body[-1].isspace():
                # `$ x $` 让过（稿件里很常见），但 `$5 和 $6` 已经在上面那条被拦下了
                pass
            spans.append((i, end + len(closer), body, display))
            i = end + len(closer)
            matched = True
            break
        if not matched:
            i += 1
    return spans


def _digit_after(text: str, index: int) -> bool:
    """闭合定界符后面紧跟数字 → 这大概率是"…$5 和 $6"里的货币符号。"""
    tail = text[index + 1 : index + 2]
    return tail.isdigit()


# ==========================================================================
# 第二步：屏蔽 → 切表格列 → 还原
# ==========================================================================


def _token(index: int) -> str:
    return "%s%04d%s" % (MASK_OPEN, index, MASK_CLOSE)


def _mask(text: str) -> tuple[str, list[str]]:
    """把公式与行内代码换成哨兵；返回 ``(屏蔽后的文本, 原片段列表)``。

    切表格列之前必须先过这一手，否则 ``| $|x|$ |`` 会被劈成三列。
    """
    pieces = [(start, end, text[start:end]) for start, end, _latex, _d in _math_spans(text)]
    for start, end in _code_spans(text):
        if not any(start >= s and end <= e for s, e, _ in pieces):
            pieces.append((start, end, text[start:end]))
    pieces.sort(key=lambda item: item[0])

    out: list[str] = []
    saved: list[str] = []
    cursor = 0
    for start, end, source in pieces:
        if start < cursor:  # 区间重叠（代码段套在公式里）就跳过
            continue
        out.append(text[cursor:start])
        out.append(_token(len(saved)))
        saved.append(source)
        cursor = end
    out.append(text[cursor:])
    return "".join(out), saved


def _unmask(text: str, saved: Sequence[str]) -> str:
    """把哨兵换回原片段。"""
    if MASK_OPEN not in text:
        return text
    return re.sub(
        MASK_OPEN + r"(\d{4})" + MASK_CLOSE,
        lambda m: saved[int(m.group(1))],
        text,
    )


def _split_table_row(line: str) -> list[str]:
    """按竖线切一行表格。屏蔽公式、并把转义竖线 ``\\|`` 先藏起来，切完再还原。"""
    masked, saved = _mask(line.strip())
    masked = masked.replace("\\|", PIPE_MASK)
    if masked.startswith("|"):
        masked = masked[1:]
    if masked.endswith("|"):
        masked = masked[:-1]
    cells: list[str] = []
    for raw in masked.split("|"):
        text = raw.replace(PIPE_MASK, "|").strip()
        cells.append(_unmask(text, saved))
    return cells


# ==========================================================================
# 第三步：行内片段
# ==========================================================================

_EMPH_DELIMS: tuple[tuple[str, str, int], ...] = (
    ("***", "***", 3),
    ("___", "___", 3),
    ("**", "**", 2),
    ("__", "__", 2),
    ("*", "*", 1),
    ("_", "_", 1),
)

#: ``[文字](地址 "可选标题")``。地址允许一层圆括号（``a(1).docx``），标题丢掉。
_RE_LINK = re.compile(
    r"\[(?P<label>[^\]]*)\]\(\s*"
    r"(?P<url>[^()\s]*(?:\([^()]*\)[^()\s]*)*)"
    r"(?:[ \t]+[\"'][^\"']*[\"'])?\s*\)"
)
#: ``![替代文字](地址)`` —— 试链接之前必须先试它，否则 ``!`` 会留在文字里
_RE_IMAGE = re.compile(
    r"!\[(?P<alt>[^\]]*)\]\(\s*"
    r"(?P<url>[^()\s]*(?:\([^()]*\)[^()\s]*)*)"
    r"(?:[ \t]+[\"'][^\"']*[\"'])?\s*\)"
)
#: ``<https://…>`` 自动链接：尖括号是 Markdown 的定界符，原样进 Word 就成了噪音
_RE_AUTOLINK = re.compile(r"<(?P<url>[A-Za-z][A-Za-z0-9+.\-]*://[^<>\s]+)>")


def _scan_inline(text: str) -> list[Inline]:
    """把一行/一段文本切成普通文字与公式，公式永远优先切出来。"""
    nodes: list[Inline] = []
    cursor = 0
    for start, end, latex, display in _math_spans(text):
        nodes.extend(_scan_plain(text[cursor:start]))
        nodes.append(MathRun(latex, display))
        cursor = end
    nodes.extend(_scan_plain(text[cursor:]))
    return nodes


def _scan_plain(text: str, bold: bool = False, italic: bool = False) -> list[Inline]:
    """切行内代码、链接/图片、强调（``**粗**`` / ``*斜*``）、转义字符；公式已经在上一层切走了。"""
    nodes: list[Inline] = []
    buf: list[str] = []
    i = 0
    n = len(text)

    def flush() -> None:
        if buf:
            nodes.append(TextRun("".join(buf), bold=bold, italic=italic))
            buf.clear()

    while i < n:
        ch = text[i]

        # 行内代码：里面什么都不认，包括 $
        if ch == "`" and not _escaped(text, i):
            run = 0
            while i + run < n and text[i + run] == "`":
                run += 1
            close = text.find("`" * run, i + run)
            if close >= 0:
                flush()
                nodes.append(CodeSpan(text[i + run : close]))
                i = close + run
                continue

        # 图片：![替代文字](地址)
        if text.startswith("![", i) and not _escaped(text, i):
            image = _RE_IMAGE.match(text, i)
            if image:
                flush()
                nodes.append(ImageRef(image.group("alt"), image.group("url")))
                i = image.end()
                continue

        # 自动链接 <https://…>：尖括号是定界符，原样进 Word 就成了噪音
        if ch == "<" and not _escaped(text, i):
            auto = _RE_AUTOLINK.match(text, i)
            if auto:
                flush()
                nodes.append(TextRun(auto.group("url"), bold=bold, italic=italic, link=auto.group("url")))
                i = auto.end()
                continue

        # 链接：[文字](地址) —— 文字里还可以有强调，所以标签再走一次本函数
        if ch == "[" and not _escaped(text, i):
            link = _RE_LINK.match(text, i)
            if link:
                flush()
                url = link.group("url")
                label = link.group("label")
                inner = _scan_plain(label, bold=bold, italic=italic) if label else []
                if not inner:
                    # ``[](地址)``：没有文字就拿地址当文字，只剩一对空方括号是最容易被人当成"这里坏了"的输出
                    inner = [TextRun(url, bold=bold, italic=italic)]
                for node in inner:
                    if isinstance(node, TextRun) and not node.link:
                        node.link = url
                nodes.extend(inner)
                i = link.end()
                continue

        # 强调
        handled = False
        for opener, closer, weight in _EMPH_DELIMS:
            if not text.startswith(opener, i) or _escaped(text, i):
                continue
            if opener == "_" and _is_word_char(text[i - 1 : i]) and _is_word_char(text[i + len(opener) : i + len(opener) + 1]):
                continue  # 词中间的 snake_case 不判斜体
            end = _find_close(text, closer, i + len(opener))
            if end < 0:
                continue
            inner = text[i + len(opener) : end]
            if not inner.strip():
                continue
            flush()
            node_bold = bold or weight >= 2
            node_italic = italic or weight == 1
            # 三星是粗+斜：拆成两层，内层再走一次本函数
            if weight == 3:
                nodes.extend(_scan_plain(inner, bold=True, italic=True))
            else:
                nodes.extend(_scan_plain(inner, bold=node_bold, italic=node_italic))
            i = end + len(closer)
            handled = True
            break
        if handled:
            continue

        # 转义：\$ \* \\_ 这些就是字面量
        if ch == "\\" and i + 1 < n and text[i + 1] in ESCAPABLE:
            buf.append(text[i + 1])
            i += 2
            continue

        buf.append(ch)
        i += 1

    flush()
    return nodes


def _is_word_char(ch: str) -> bool:
    return bool(ch) and (ch.isalnum() or ch == "_")


def _marker_ordered(marker: str) -> bool:
    """``1.`` / ``1)`` 是有序条目，``-`` / ``*`` / ``+`` 是无序条目。"""
    return marker[-1] in ".)"


# ==========================================================================
# 第四步：块级结构
# ==========================================================================

_RE_ATX = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_RE_FENCE = re.compile(r"^(\s*)(`{3,}|~{3,})\s*([\w.+-]*)\s*$")
_RE_THEMATIC = re.compile(r"^\s{0,3}((\*\s*){3,}|(-\s*){3,}|(_\s*){3,})$")
_RE_SETEXT = re.compile(r"^\s{0,3}(=+|-+)\s*$")
_RE_BLOCKQUOTE = re.compile(r"^\s{0,3}>\s?(.*)$")
_RE_LIST_ITEM = re.compile(r"^(?P<indent>[ \t]*)(?P<marker>[-*+]|\d{1,9}[.)])[ \t]+(?P<rest>.*)$")
_RE_ORDERED_START = re.compile(r"^(\d+)[.)]")
_RE_TASK = re.compile(r"^\[([ xX])\]\s+")
_RE_TABLE_DELIM = re.compile(r"^\s*\|?(?:\s*:?-{2,}:?\s*\|)+[\s:|-]*$")


def normalize_source(text: str) -> str:
    """统一换行、去 BOM、把行首制表位按 4 格展开（代码块内容只动行首，不动中间）。"""
    text = text.lstrip("\ufeff")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = []
    for line in text.split("\n"):
        stripped = line.lstrip(" \t")
        lead = line[: len(line) - len(stripped)].replace("\t", "    ")
        lines.append(lead + stripped)
    return "\n".join(lines)


class _Parser:
    def __init__(self) -> None:
        self.warnings: list[str] = []
        self.issues: list[Issue] = []
        self.stats: dict[str, int] = {}

    # ------------------------------------------------------------ 主循环
    def parse_lines(self, lines: Sequence[str], base_line: int = 1) -> list[Block]:
        blocks: list[Block] = []
        i = 0
        n = len(lines)
        while i < n:
            line = lines[i]
            if not line.strip():
                i += 1
                continue

            # 围栏代码块
            fence = _RE_FENCE.match(line)
            if fence:
                block, i = self._read_code_block(lines, i, fence, base_line)
                blocks.append(block)
                continue

            # 独占一行的显示式公式
            math_block = self._try_block_math(lines, i, base_line)
            if math_block is not None:
                block, i = math_block
                blocks.append(block)
                continue

            # ATX 标题
            atx = _RE_ATX.match(line)
            if atx:
                blocks.append(Heading(len(atx.group(1)), _scan_inline(atx.group(2)), base_line + i))
                i += 1
                continue

            # 分隔线（放在段落后面的 --- 是 setext 二级标题，那种情况在下面处理）
            if _RE_THEMATIC.match(line):
                blocks.append(ThematicBreak(base_line + i))
                i += 1
                continue

            # 表格
            if self._is_table_start(lines, i):
                block, i = self._read_table(lines, i, base_line)
                blocks.append(block)
                continue

            # 引用块
            if _RE_BLOCKQUOTE.match(line):
                block, i = self._read_quote(lines, i, base_line)
                blocks.append(block)
                continue

            # 列表
            if _RE_LIST_ITEM.match(line):
                block, i = self._read_list(lines, i, base_line)
                blocks.append(block)
                continue

            # 普通段落（吃掉后面的 setext 下划线）
            block, i = self._read_paragraph(lines, i, base_line)
            blocks.append(block)
        return blocks

    # ------------------------------------------------------------ 各种块
    def _read_code_block(self, lines, i, fence, base_line):
        indent, marker, lang = fence.group(1), fence.group(2), fence.group(3)
        closer = re.compile(r"^\s*" + re.escape(marker) + r"\s*$")
        body: list[str] = []
        i += 1
        start = base_line + i
        while i < len(lines) and not closer.match(lines[i]):
            # 只按这一层围栏的缩进量往外剥，剥多了会把代码的缩进毁掉
            line = lines[i]
            body.append(line[len(indent) :] if line.startswith(indent) else line)
            i += 1
        if i >= len(lines):
            raise MarkdownSyntaxError(
                "第 %d 行的围栏代码块（%s）没有闭合的那一行，不敢猜后面还有多少算代码" % (base_line + i - len(body) - 1, marker)
            )
        i += 1
        return CodeBlock("\n".join(body), lang, start), i

    def _try_block_math(self, lines, i, base_line):
        """独占一行的 ``$$…$$`` / ``\\[…\\]``（允许跨行）。

        一行里除了公式还有别的话（``$$a$$ 且 $$b$$``）就不算块级，
        交给行内那条路 —— 这是业务方定的口径①。
        """
        stripped = lines[i].strip()
        for opener, closer in (("$$", "$$"), ("\\[", "\\]")):
            if not stripped.startswith(opener):
                continue
            if stripped.endswith(closer) and len(stripped) > len(opener) + len(closer) - 2:
                spans = _math_spans(stripped)
                if len(spans) == 1 and spans[0][0] == 0 and spans[0][1] == len(stripped):
                    body = stripped[len(opener) : len(stripped) - len(closer)]
                    return BlockMath(body.strip(), base_line + i), i + 1
                continue
            end = self._find_math_close(lines, i, closer)
            if end is None:
                continue
            parts: list[str] = []
            first = stripped[len(opener) :]
            if first.strip():
                parts.append(first.strip())
            for j in range(i + 1, end):
                if lines[j].strip():
                    parts.append(lines[j].strip())
            tail = lines[end].strip()
            index = tail.rfind(closer)
            if index > 0:
                parts.append(tail[:index].strip())
            body = "\n".join(parts).strip()
            if body:
                return BlockMath(body, base_line + i), end + 1
        return None

    def _find_math_close(self, lines, start, closer) -> Optional[int]:
        for j in range(start, len(lines)):
            line = lines[j].strip()
            if j == start:
                continue
            if line.endswith(closer) and not _escaped(line, max(0, len(line) - len(closer))):
                return j
        return None

    def _is_table_start(self, lines, i) -> bool:
        """表格的判据：这一行含竖线，下一行是 ``|---|---|`` 那样的分隔行。"""
        if i + 1 >= len(lines):
            return False
        if "|" not in lines[i]:
            return False
        return bool(_RE_TABLE_DELIM.match(lines[i + 1].strip()))

    def _read_table(self, lines, i, base_line):
        start_line = base_line + i
        header = _split_table_row(lines[i])
        aligns = self._read_aligns(lines[i + 1], len(header))
        rows: list[list[list[Inline]]] = [ [_scan_inline(c) for c in header] ]
        i += 2
        while i < len(lines) and lines[i].strip() and "|" in lines[i]:
            cells = _split_table_row(lines[i])
            if len(cells) < len(header):
                sentence = "第 %d 行只有 %d 列，表头是 %d 列，缺的按空格补齐" % (
                    base_line + i, len(cells), len(header))
                self.warnings.append(sentence)
                self.issues.append(Issue("table_columns", base_line + i, sentence))
                cells += [""] * (len(header) - len(cells))
            elif len(cells) > len(header):
                sentence = ("第 %d 行有 %d 列，比表头 %d 列多，多出来的已并进最后一格"
                            % (base_line + i, len(cells), len(header)))
                self.warnings.append(sentence)
                self.issues.append(Issue("table_columns", base_line + i, sentence))
                cells = cells[: len(header) - 1] + [" ".join(cells[len(header) - 1 :])]
            rows.append([_scan_inline(cell) for cell in cells])
            i += 1
        table = Table(rows, aligns, True, start_line)
        return table, i

    def _read_aligns(self, delim_line: str, count: int) -> list[str]:
        marks = _split_table_row(delim_line)
        marks = marks + [""] * (count - len(marks))
        aligns: list[str] = []
        for mark in marks[:count]:
            mark = mark.strip()
            left = mark.startswith(":")
            right = mark.endswith(":")
            if left and right:
                aligns.append("center")
            elif right:
                aligns.append("right")
            elif left:
                aligns.append("left")
            else:
                aligns.append("none")
        return aligns

    def _read_quote(self, lines, i, base_line):
        inner: list[str] = []
        start = base_line + i
        while i < len(lines):
            match = _RE_BLOCKQUOTE.match(lines[i])
            if match:
                inner.append(match.group(1))
                i += 1
                continue
            # 引用块的"惰性延续"：非空且不属别的块的一行，算同一段
            if lines[i].strip() and inner and inner[-1].strip():
                inner.append(lines[i].strip())
                i += 1
                continue
            break
        return QuoteBlock(self.parse_lines(inner, start), start), i

    def _read_list(self, lines, i, base_line):
        start = base_line + i
        first = _RE_LIST_ITEM.match(lines[i])
        assert first is not None
        base_indent = len(first.group("indent"))
        ordered = _marker_ordered(first.group("marker"))
        number = _RE_ORDERED_START.match(first.group("marker"))
        blocks_start = int(number.group(1)) if number else 1
        items, i = self._read_list_items(lines, i, base_indent, base_line, ordered)
        return ListBlock(ordered, items, blocks_start, start), i

    def _read_list_items(self, lines, i, base_indent, base_line, ordered):
        """按缩进吃条目。``ordered`` 是"这一层是有序还是无序"——
        无序表后面紧跟一个有序表是**两张表**，不能因为中间只隔一个空行就并成一张
        （并了以后 Word 里那 1. 2. 会变成第三、第四个圆点，序号全丢）。
        """
        items: list[ListItem] = []
        n = len(lines)
        while i < n:
            match = _RE_LIST_ITEM.match(lines[i])
            if not match:
                if not lines[i].strip():
                    nxt = _RE_LIST_ITEM.match(lines[i + 1]) if i + 1 < n else None
                    if nxt and len(nxt.group("indent")) >= base_indent and _marker_ordered(nxt.group("marker")) == ordered:
                        i += 1  # 松散列表：同类型的下一条还归这一张表
                        continue
                    break
                # 缩进的延续行，接在上一条的正文后面
                if len(lines[i]) - len(lines[i].lstrip()) > base_indent and items:
                    items[-1].inline = _scan_inline(
                        _join_text(_inlines_text(items[-1].inline), lines[i].strip())
                    )
                    i += 1
                    continue
                break
            indent = len(match.group("indent"))
            if indent < base_indent:
                break
            if indent > base_indent:
                # 更深层的列表：整段交给下一层，挂到上一条的 children
                child_items, i = self._read_list_items(
                    lines, i, indent, base_line, _marker_ordered(match.group("marker"))
                )
                if items:
                    items[-1].children.extend(child_items)
                else:
                    items.extend(child_items)
                continue
            rest = _task_glyph(match.group("rest"))
            items.append(ListItem(_scan_inline(rest), ordered=ordered))
            i += 1
        return items, i

    def _read_paragraph(self, lines, i, base_line):
        start = base_line + i
        # 每行连同"这一行结尾是不是硬换行"一起存：行首行尾都 strip 过了，
        # 事后再看结尾那两个空格已经没了（第一版就是栽在这里）
        collected: list[tuple[str, bool]] = []
        while i < len(lines):
            line = lines[i]
            if not line.strip():
                break
            if collected:
                setext = _RE_SETEXT.match(line)
                if setext:
                    level = 1 if setext.group(1).startswith("=") else 2
                    text = _join_lines(collected)
                    return Heading(level, _scan_inline(text), start), i + 1
                if _RE_FENCE.match(line) or _RE_THEMATIC.match(line) or _RE_LIST_ITEM.match(line) \
                        or _RE_BLOCKQUOTE.match(line) or _RE_ATX.match(line) or self._is_table_start(lines, i):
                    break
                if line.strip().startswith("$$") or line.strip().startswith("\\["):
                    break
            collected.append((line.strip(), line.endswith("  ") or line.endswith("\\")))
            i += 1
        if not collected:
            return Paragraph([], start), i + 1
        return Paragraph(_scan_inline(_join_lines(collected)), start), i


def _inlines_text(nodes: Iterable[Inline]) -> str:
    """把行内片段还原成文本（只在合并列表延续行时用，公式原文一并带回去）。"""
    out: list[str] = []
    for node in nodes:
        if isinstance(node, TextRun):
            out.append("[%s](%s)" % (node.text, node.link) if node.link else node.text)
        elif isinstance(node, CodeSpan):
            out.append("`%s`" % node.text)
        elif isinstance(node, ImageRef):
            out.append("![%s](%s)" % (node.alt, node.url))
        else:
            out.append("$%s$" % node.latex)
    return "".join(out)


def _join_text(previous: str, extra: str) -> str:
    return (previous + " " + extra).strip()


def _join_lines(entries: Sequence[tuple[str, bool]]) -> str:
    """段落多行合并：上一行以两个空格或 ``\\`` 结尾 → 硬换行（``\\n``），否则空格相连。"""
    if not entries:
        return ""
    out = entries[0][0]
    for index in range(1, len(entries)):
        out += ("\n" if entries[index - 1][1] else " ") + entries[index][0]
    return out


def _task_glyph(rest: str) -> str:
    """``- [ ]`` / ``- [x]`` 换成方框字符（数据模型里不为待办项加字段）。"""
    match = _RE_TASK.match(rest)
    if not match:
        return rest
    glyph = "☑ " if match.group(1).lower() == "x" else "☐ "
    return glyph + rest[match.end() :]


# ==========================================================================
# 对外接口
# ==========================================================================

#: 试着按这个顺序读；中文 .md 在 Windows 上确实常是 GBK/ANSI 存的
ENCODING_TRIES = ("utf-8-sig", "utf-8", "gb18030")

#: UTF-8 的 BOM。有它才报 utf-8-sig——utf-8-sig 对没有 BOM 的文件照样解得开，
#: 而摘要里那一格"编码"是排查"中文是不是读坏了"的第一条线索，说假话比不说更糟。
BOM_UTF8 = b"\xef\xbb\xbf"


def read_markdown(path: str | Path, encoding: Optional[str] = None) -> tuple[str, str]:
    """读文件，返回 ``(文本, 实际用的编码)``；``encoding`` 给了就只试那一种。

    自动试 utf-8-sig / utf-8 / gb18030：中文稿件用 GBK 保存很常见，
    只认 UTF-8 会在第一行就报一句和 Markdown 毫无关系的解码错误。
    """
    target = Path(path)
    if not target.is_file():
        raise FileNotFoundError("找不到输入文件：%s" % target)
    raw = target.read_bytes()
    if not raw.strip():
        return "", (encoding or "utf-8")
    problems: list[str] = []
    for name in ((encoding,) if encoding else ENCODING_TRIES):
        try:
            text = raw.decode(name)
        except (UnicodeDecodeError, LookupError) as exc:
            reason = getattr(exc, "reason", None) or str(exc)
            problems.append("%s（%s）" % (name, reason))
            continue
        if name == "utf-8-sig" and not raw.startswith(BOM_UTF8):
            name = "utf-8"
        return text, name
    raise UnicodeDecodeError(
        "utf-8", raw, 0, min(len(raw), 64),
        "这个文件按这几种编码都读不出来：%s；用 --encoding 指定一次" % "、".join(problems),
    )


def parse(text: str) -> ParseResult:
    """文本 → :class:`ParseResult`。"""
    parser = _Parser()
    lines = normalize_source(text).split("\n")
    blocks = parser.parse_lines(lines) if text.strip() else []
    _assert_no_mask_left(blocks)
    stats = _count_stats(blocks)
    return ParseResult(blocks, parser.warnings, stats, parser.issues)


def parse_file(path: str | Path, encoding: Optional[str] = None) -> ParseResult:
    """读文件并解析；``encoding`` 给了就只试那一种。"""
    text, used = read_markdown(path, encoding)
    result = parse(text)
    _stamp_image_base(result, Path(path).resolve().parent)
    result.stats["encoding"] = used
    result.stats["encoding_bytes"] = len(text.encode(used, errors="replace"))
    if used != "utf-8":
        # BOM 与 GBK 都值得一句说明：这两种稿子在 Windows 上真常见，
        # 而"读出来的字对不对"是排查转换结果时的第一个问题
        result.warnings.append("文件按 %s 编码读取" % used)
    return result


# --------------------------------------------------------------------------
# 图片地址：一个地址在本机上该去哪儿找 —— 校验与嵌图共用这一条规则
# --------------------------------------------------------------------------

#: 这些开头的地址不是本机路径：本工具不下载、不解码，所以也不查盘。
IMAGE_REMOTE_PREFIXES = ("http://", "https://", "data:", "mailto:", "//")


def image_target(url: str, base: Optional[Path] = None) -> Optional[Path]:
    """``![…](…)`` 里那个地址换成本机上该查的位置；不是本机路径的交回 ``None``。

    先去 ``#锚点`` 与 ``?宽=200`` 这类查询串（它们不是路径的一部分），再百分号解码
    （``My%20Figure.png`` 里那个 ``%20`` 真人敲得出来），最后按 ``base`` 拼绝对路径。
    ``base`` 是这份 .md 所在的目录（:func:`parse_file` 盖在节点上的那一格）；没有就按
    当前工作目录——按 cwd 查会出现"在资源管理器里双击打得开、工具却说找不到"。

    为什么放在这一层：校验（:mod:`gui.validator`）与嵌图（:mod:`docx_builder`）都要回答
    "这张图在哪"。两边各写一套就会出现"校验说找不到、Word 里却嵌进去了"这种对不上的回执。
    """
    raw = (url or "").strip()
    if not raw or raw.startswith(IMAGE_REMOTE_PREFIXES):
        return None
    raw = raw.split("#", 1)[0].split("?", 1)[0]
    if not raw:
        return None
    return ((base if base is not None else Path.cwd()) / unquote(raw)).resolve()


def image_readable(target: Path) -> bool:
    """那个位置是不是一个真打得开的文件。路径坏到系统都不认（非法字符、超长）算"不是"，
    不抛——一句"找不到"比让异常把整轮校验或整份文档带崩有用。"""
    try:
        return target.is_file()
    except OSError:
        return False


def _stamp_image_base(result: ParseResult, directory: Path) -> None:
    """把这份 .md 的目录盖到它自己的每个 :class:`ImageRef` 上（不读图片字节）。"""
    for node, _line in result.annotated(ImageRef):
        node.base = directory


def _assert_no_mask_left(blocks: Sequence[Block]) -> None:
    """哨兵字符绝不能漏出这一层：漏了说明还原那一刀切错了，宁可报错也不要往 Word 里写怪字符。"""
    leftovers: list[str] = []

    def check(text: str) -> None:
        if MASK_OPEN in text or MASK_CLOSE in text or PIPE_MASK in text:
            leftovers.append(text[:40])

    def walk_inlines(nodes: Iterable[Inline]) -> None:
        for node in nodes:
            if isinstance(node, TextRun):
                check(node.text)
                if node.link:
                    check(node.link)
            elif isinstance(node, ImageRef):
                check(node.alt)
                check(node.url)

    def walk(blocks: Sequence[Block]) -> None:
        for block in blocks:
            if isinstance(block, (Heading, Paragraph)):
                walk_inlines(block.inline)
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row:
                        walk_inlines(cell)
            elif isinstance(block, ListBlock):
                _walk_list_items(block.items, walk_inlines)
            elif isinstance(block, QuoteBlock):
                walk(block.blocks)

    walk(blocks)
    if leftovers:
        raise MarkdownSyntaxError("解析内部错误：屏蔽标记没还原干净（%s）" % " / ".join(leftovers[:3]))


def _count_stats(blocks: Sequence[Block]) -> dict[str, int]:
    stats = {
        "blocks": len(blocks),
        "tables": 0,
        "formulas_inline": 0,
        "formulas_display": 0,
        "headings": 0,
        "code_blocks": 0,
        "links": 0,
        "images": 0,
    }

    def count_inlines(nodes: Iterable[Inline]) -> None:
        for node in nodes:
            if isinstance(node, MathRun):
                key = "formulas_display" if node.display else "formulas_inline"
                stats[key] += 1
            elif isinstance(node, ImageRef):
                stats["images"] += 1
            elif isinstance(node, TextRun) and node.link:
                stats["links"] += 1

    def walk(seq: Sequence[Block]) -> None:
        for block in seq:
            if isinstance(block, Heading):
                stats["headings"] += 1
                count_inlines(block.inline)
            elif isinstance(block, Paragraph):
                count_inlines(block.inline)
            elif isinstance(block, BlockMath):
                stats["formulas_display"] += 1
            elif isinstance(block, Table):
                stats["tables"] += 1
                for row in block.rows:
                    for cell in row:
                        count_inlines(cell)
            elif isinstance(block, ListBlock):
                _walk_list_items(block.items, count_inlines)
            elif isinstance(block, QuoteBlock):
                walk(block.blocks)
            elif isinstance(block, CodeBlock):
                stats["code_blocks"] += 1

    walk(blocks)
    return stats


def render_blocks(blocks: Sequence[Block], indent: int = 0) -> str:
    """把解析结果画成一棵文字树（自检与报告用，不参与生成 Word）。"""
    pad = "    " * indent
    out: list[str] = []
    for block in blocks:
        if isinstance(block, Heading):
            out.append("%sHeading(%d) line=%d %s" % (pad, block.level, block.line, _inline_repr(block.inline)))
        elif isinstance(block, Paragraph):
            out.append("%sParagraph line=%d %s" % (pad, block.line, _inline_repr(block.inline)))
        elif isinstance(block, BlockMath):
            out.append("%sBlockMath line=%d latex=%r" % (pad, block.line, block.latex))
        elif isinstance(block, CodeBlock):
            out.append("%sCodeBlock line=%d lang=%r %d 行" % (pad, block.line, block.lang, block.text.count("\n") + 1))
        elif isinstance(block, Table):
            out.append("%sTable line=%d %d行×%d列 aligns=%s" % (pad, block.line, len(block.rows), block.column_count, block.aligns))
            for index, row in enumerate(block.rows):
                out.append("%s  行%d: %s" % (pad, index, " | ".join(_inline_repr(cell) for cell in row)))
        elif isinstance(block, ListBlock):
            out.append("%sList line=%d ordered=%s start=%d" % (pad, block.line, block.ordered, block.start))
            _render_items(block.items, indent + 1, out)
        elif isinstance(block, QuoteBlock):
            out.append("%sQuote line=%d" % (pad, block.line))
            out.append(render_blocks(block.blocks, indent + 1))
        elif isinstance(block, ThematicBreak):
            out.append("%sThematicBreak line=%d" % (pad, block.line))
    return "\n".join(out)


def _render_items(items: Sequence[ListItem], indent: int, out: list[str]) -> None:
    pad = "    " * indent
    for item in items:
        out.append("%s- %s" % (pad, _inline_repr(item.inline)))
        if item.children:
            _render_items(item.children, indent + 1, out)


def _inline_repr(nodes: Iterable[Inline]) -> str:
    parts: list[str] = []
    for node in nodes:
        if isinstance(node, TextRun):
            flag = "".join(f for f, on in (("b", node.bold), ("i", node.italic)) if on)
            if node.link:
                parts.append("link%s(%r->%r)" % ("[" + flag + "]" if flag else "", node.text, node.link))
            else:
                parts.append("text%s(%r)" % ("[" + flag + "]" if flag else "", node.text))
        elif isinstance(node, CodeSpan):
            parts.append("code(%r)" % node.text)
        elif isinstance(node, ImageRef):
            parts.append("image(%r->%r)" % (node.alt, node.url))
        else:
            parts.append("math(%r%s)" % (node.latex, ",display" if node.display else ""))
    return "".join(parts) or "(空)"
