# -*- coding: utf-8 -*-
"""只解析、不写盘的四项检查：公式语法 / 图片路径 / 表格列数 / 标题层级。

校验的存在理由是**在转换之前**把"这份文件里有个东西坏了"念出来，而且每条都要
指得回行号——一句"有公式语法错误"在三十行的文件里等于没说。

四条检查各自借现成的判定，这一层不写第二套：

* 公式语法走 :func:`gui.math_html.latex_problem`（预览里认不出来的写法与这里认不出的
  必须是同一批，否则预览看着好的那条一校验就报坏）。
* 表格列数**已经在解析器里报了**（:mod:`md_parser` 的 ``_read_table``），这里只把它
  交回的结构化 :class:`md_parser.Issue` 换成带档位的记录，不再数一遍列。
* 图片路径由解析器交出 ``ImageRef`` 之后才查盘：解析器刻意不读图片字节（它只管结构），
  所以"文件在不在"是校验这一层的事。而"这个地址该去哪儿找"用的是
  :func:`md_parser.image_target` —— 嵌图那边读的就是这一条，两边不会各判一套。
* 标题层级用的是块上那个 ``level``，不回头重读 ``#``。

档位是**后果**不是类别：会让成果出错的（公式转不出、图找不到）算 ``error``，
只影响读起来顺不顺的（标题跳级、解析器已按空格补齐的列数）算 ``warn``。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import md_parser as mp
from gui import math_html

__all__ = ["CODE_LABELS", "Finding", "summary_text", "validate_file",
           "validate_parsed", "worst_token"]

#: 会让 Word 里那份成果出问题的——日志区按"坏"上色。
LEVEL_ERROR = "error"
#: 不算坏，但看的人会当成坏，或者下一步就该顺手改的——按"提醒"上色。
LEVEL_WARN = "warn"

#: 每一类给人念的那个名字（日志那一行与报告里用的是同一个词，不在两处各拼一遍）。
CODE_LABELS = {
    "math_syntax": "公式语法",
    "image_missing": "图片路径",
    "table_columns": "表格列数",
    "heading_skip": "标题层级",
    "file_missing": "文件",
    "parse_failed": "解析",
}

#: 行号对不上的那几条（整份文件读不进）在日志里顶替"第 N 行"的那一格。
WHERE_UNKNOWN = "整份文件"

#: 坏公式那句后面要带的原文——一条 `\begin{aligned}` 写漏 \end 时，
#: 没有原文就得到处翻文件；整条都贴出来又会把日志区那一行撑满。
LATEX_ECHO_CHARS = 48


@dataclass(frozen=True)
class Finding:
    """一条能定位到行的检查结果。``message`` 是可执行的那句中文，不带行号前缀。"""

    level: str
    code: str
    line: int
    message: str

    @property
    def text(self) -> str:
        """日志区那一整行：``第 12 行｜公式语法：…``。"""
        head = where(self.line)
        message = self.message
        if message.startswith(head):
            # 解析器那句本来就带着行号（它得自己说清是哪一行），拼起来会变成
            # "第 19 行｜表格列数：第 19 行只有 2 列"——同一件事不说两遍。
            message = message[len(head):].lstrip("，,：: ")
        return "%s｜%s：%s" % (head, CODE_LABELS.get(self.code, self.code), message)

    @property
    def token(self) -> str:
        """主题里那格颜色名（主窗口上色只读这一格，不在界面里 ``if level == "error"``）。"""
        return "bad" if self.level == LEVEL_ERROR else "warn"


def where(line: int) -> str:
    return "第 %d 行" % line if line > 0 else WHERE_UNKNOWN


# ------------------------------------------------------------------ 入口


def validate_file(path: Path) -> list[Finding]:
    """读一份文件并查四项。读不动不抛——交回一条 :data:`LEVEL_ERROR` 就算查完了。

    这一条是给"只想看一眼某个文件有没有毛病"的调用方用的（测试与命令行）；
    批量校验走 :func:`validate_parsed`，那边早就解析过一次，不该再读一遍盘。
    """
    if not path.is_file():
        return [Finding(LEVEL_ERROR, "file_missing", 0,
                        "%s 已经不在这个位置了（先把它从左边列表里移除）" % path.name)]
    try:
        parsed = mp.parse_file(path)
    except (OSError, UnicodeDecodeError, ValueError, mp.MarkdownSyntaxError) as exc:
        detail = str(exc) or type(exc).__name__
        return [Finding(LEVEL_ERROR, "parse_failed", 0, "读不动或解析不动：%s" % detail)]
    return validate_parsed(parsed, base=path.parent)


def validate_parsed(parsed: mp.ParseResult,
                    base: Optional[Path] = None) -> list[Finding]:
    """对已经解析好的结果查四项，按行号交回（行号对不上的排最后）。

    图片地址按哪个目录找，**首选节点自己那一格**（:func:`md_parser.parse_file` 盖上去的
    这份 .md 所在的目录），``base`` 是回落——给 :func:`md_parser.parse` 的结果用的。
    合并稿里几个文件的块并成一棵树，各自的相对路径得回各自那一格查，
    拿一个目录管全场会把第二个文件的图判成"找不到"。
    按当前工作目录查会出现"在资源管理器里双击打得开、校验却说找不到"，因为那两边根本不在同一个目录。
    """
    found = list(_mathematical(parsed))
    found.extend(_images(parsed, base if base is not None else Path.cwd()))
    found.extend(_tables(parsed))
    found.extend(_headings(parsed))
    found.sort(key=lambda item: (item.line <= 0, item.line, item.code))
    return found


def summary_text(findings: Iterable[Finding]) -> str:
    """一句"查出几处、其中几处是坏的"，给日志与回执用。"""
    items = list(findings)
    if not items:
        return "四项都没查出问题"
    bad = sum(1 for item in items if item.level == LEVEL_ERROR)
    if bad == 0:
        return "查出 %d 处提醒（没有会出错的那一档）" % len(items)
    return "查出 %d 处问题，其中 %d 处会让 Word 里那份成果出错" % (len(items), bad)


def worst_token(findings: Iterable[Finding]) -> str:
    """一组结果在日志里该上哪一档颜色：有一条坏就是坏，只有提醒就是提醒，全空就是好。

    概要那一行念的是条数、上色却按**最重的那一档**：一条"图找不到"混在四条提醒里时
    这一行还是绿的，看的人会以为这一份查出来的都是小事。
    """
    items = list(findings)
    if not items:
        return "ok"
    if any(item.level == LEVEL_ERROR for item in items):
        return "bad"
    return "warn"


# ------------------------------------------------------------------ 四项


def _mathematical(parsed: mp.ParseResult) -> Iterable[Finding]:
    for node, line in parsed.annotated(mp.MathRun, mp.BlockMath):
        problem = math_html.latex_problem(node.latex)
        if not problem:
            continue
        yield Finding(LEVEL_ERROR, "math_syntax", line,
                      "%s（原文：%s）" % (problem, _echo(node.latex)))


def _images(parsed: mp.ParseResult, directory: Path) -> Iterable[Finding]:
    for node, line in parsed.annotated(mp.ImageRef):
        # "这个地址该去哪儿找"用的是解析器那一条规则（嵌图那边读的是同一条），
        # 节点自带的目录优先：合并稿里每个文件找自己那一格。
        target = mp.image_target(node.url, node.base or directory)
        if target is None:
            continue
        if mp.image_readable(target):
            continue
        yield Finding(LEVEL_ERROR, "image_missing", line,
                      "找不到图片：%s（按这个位置查的：%s）" % (node.url, target))


def _tables(parsed: mp.ParseResult) -> Iterable[Finding]:
    """列数不一致由解析器判（它本来就在那一行上补齐或合并了），这里只换个形状交回。

    算 ``warn`` 不算 ``error``：解析器已经当场修好了，进了 Word 不会多一列也不会少一列，
    真正会出错的是"这一行本来该有四格、你只写了三格"那种数据问题——那句话原样写在
    ``message`` 里，读的人自己判断。
    """
    for issue in parsed.issues:
        if issue.code != "table_columns":
            continue
        yield Finding(LEVEL_WARN, issue.code, issue.line, issue.message)


def _headings(parsed: mp.ParseResult) -> Iterable[Finding]:
    """跳级：``#`` 后面直接来 ``###``。

    刻意**不**要求第一条必须是 ``#``（一份只写章节的说明文件从 ``##`` 起头是常见写法，
    报它等于报噪音），也只看相邻两级——"文档该有几章"是内容问题不是结构错误。
    """
    previous = 0
    for node, line in parsed.annotated(mp.Heading):
        if previous and node.level > previous + 1:
            yield Finding(LEVEL_WARN, "heading_skip", line,
                          "标题从第 %d 级直接跳到第 %d 级（中间少了第 %d 级，Word 的导航窗格会少一层）"
                          % (previous, node.level, node.level - 1))
        previous = node.level


# ------------------------------------------------------------------ 小工具


def _echo(latex: str) -> str:
    source = " ".join((latex or "").split())
    if len(source) <= LATEX_ECHO_CHARS:
        return source
    return source[:LATEX_ECHO_CHARS] + "…"
