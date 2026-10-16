# -*- coding: utf-8 -*-
"""把一条 LaTeX 变成 Qt 富文本画得出来的 HTML，并对它做一次语法体检。

预览区是 :class:`QTextBrowser.QTextBrowser`，不是浏览器内核：这个环境装的是
PySide6-Essentials，没有 QtWebEngine，所以"用 KaTeX / MathJax 渲染"那条路（要一个
能跑 JS 的网页控件）在这里跑不起来；而直接把 MathML 交给 Qt 会被整个摊平
（实测 ``\\frac{a}{b}`` 只剩 ``ab``）。所以这一层自己走一遍 latex2mathml 交回的
MathML 树，只挑 Qt 真认的几样拼：上标 ``<sup>``、下标 ``<sub>``、上下两格的分式表格、
以及 Unicode 字符——希腊字母与运算符 latex2mathml 交回的就是 ``α ∑ ≠``，不需要再映射。

三档"能不能画"是量出来的，不是猜的（:mod:`tests.test_math_html` 钉着形状）：

* ``<sup>`` / ``<sub>`` 会把墨水抬高、压低；
* 组合符号（U+0302 帽子、U+20D7 向量箭头、U+0305 上横线）画得出来，
  而 ``text-decoration:overline`` 一点墨都不出，所以根号写成 ``√( … )``；
* 分式只有摆成块级的两行表格才真上下叠；行内 ``<table>`` 会被 Qt 提成块级、
  把那一行打断，所以**行内分式写成** ``(a)/(b)``。块级那一张的格线**能关**，但要走 CSS：
  HTML 属性 ``border=0`` 与不写一模一样（实测），而 :func:`gui.theme.preview_css` 里按类名写的
  ``.math-frac td { border: none }`` 真被认（同一屏逐像素扫：盖了竖线 0 条，摘掉长出 2 条），
  于是预览里的分式是"上下两格 + 一条分数线"（分数线由权重更高的 ``.math-frac .math-num`` 画，
  实测那一行 58 像素的墨，摘掉那句就一条都没有）——仍然不装作它就是 Word 里那个样子。

:func:`latex_problem` 是**语法体检**，与 :func:`render_math` 同一份判法（预览里坏公式
退回原文、校验里念同一句原因）。为什么不能只靠 latex2mathml 抛不抛异常：它对真的写坏了
是**默默按剩下的算**——实测 ``\\frac{a}`` 交出 ``a``、``\\unknownmacro{x}`` 交出原样的文本、
``{}`` 交出空。所以括号成对、宏要的那几组 ``{…}``、产出里有没有残留的反斜杠、以及那一段
MathML **lxml 认不认**，这四条得自己判。第四条不是防有的公式：正文里带**裸的** ``&`` 时
latex2mathml 交回 ``<mi>&</mi>``，那不是合法 XML（要显示这个符号得写 ``\\&``），
实测 ``$a & b$`` 会让体检本身抛穿——而 :func:`_tree` 一侧早就包了，所以两边都得判同一件事。

这一层**不认识 Office**：只用 latex2mathml，绝不要求 ``MML2OMML.XSL``，
所以"校验"那一路在没装 Word 的机器上照样跑得完（转换那一路才需要样式表）。
"""

from __future__ import annotations

import re
from functools import lru_cache
from html import escape
from typing import Callable, Dict, List, Tuple

import latex2mathml.converter
from lxml import etree

__all__ = ["latex_problem", "render_math"]

#: 这些宏要紧跟两组 ``{…}``，只跟了一组就是敲漏了。只列"少一组会算出残缺公式"的那几个：
#: ``\hat n`` 这种不带花括号是合法写法，所以判的是"跟了但至少少一组"，不是"没带括号"。
_TWO_ARG_MACROS = frozenset({"frac", "dfrac", "tfrac", "binom", "stackrel",
                             "overset", "underset"})

_MACRO_NAME = re.compile(r"\\([A-Za-z]+)")
_STRAY_DOLLAR = re.compile(r"(?<!\\)\$")

#: 整条公式只有间距也算合法（``$a\,b$`` 里那种 ``\,``），不能报成"没有内容"。
_SPACING_ONLY = re.compile(r"^(?:\s|\\,|\\;|\\:|\\!|\\\\|\\quad|\\qquad|\\thinspace)+$")

#: MathML 里 ``mover`` 的那个记号 → 组合 Unicode。左边的记号是实测交回的文本
#: （``\hat``→``^``、``\vec``→``→``、``\overline``→``―``、``\dot``→``˙``）。
_ACCENT_MARKS = {
    "^": "\u0302", "\u02c6": "\u0302", "\u2227": "\u0302",
    "~": "\u0303", "\u02dc": "\u0303", "\u223c": "\u0303",
    "\u2192": "\u20d7", "\u21c0": "\u20d1",
    "\u2015": "\u0305", "\u00af": "\u0305", "\u203e": "\u0305",
    "-": "\u0305", "=": "\u0305",
    ".": "\u0307", "\u02d9": "\u0307",
    "..": "\u0308", '"': "\u0308",
    "\u2218": "\u030a",
    "'": "\u0301", "`": "\u0300", "\u02d8": "\u030c",
    "\u2192\u2192": "\u20d7",
}

#: 这几枚记号要**盖住整个底数**（横线、下划线），所以逐格跟在每个叶子文本后面；
#: 其余的只跟在底数最后面（一个帽子、一个箭头）。
_SPAN_MARKS = frozenset({"\u0305", "\u0332"})

#: 产出文本里剩下的每一个反斜杠都是"这一版没认出来的写法"：实测合法的
#: ``\$ \% \& \# \_ \{ \} \, \; \: \! \\ \quad \alpha … \vec \mathbb \text``
#: 交回来的文本里**一个反斜杠都不留**，所以留着就是坏的（``\3`` 这种接不上名字的也算）。
_RESIDUAL = re.compile(r"\\(?:([A-Za-z]+)|(.))?")


# ---------------------------------------------------------------- 语法体检


def latex_problem(latex: str) -> str:
    """这条 LaTeX 有没有毛病；没问题交回空串。预览与校验共用这一份判法。

    交回的是一句能直接念给人听的原因（写的是"下一步该改哪一格"，不是异常栈）。
    """
    source = (latex or "").strip()
    if not source:
        return "公式是空的（定界符之间没有内容）"

    for check in (_stray_delimiter, _brace_problem, _missing_group):
        sentence = check(source)
        if sentence:
            return sentence

    mathml, failure = _convert(source)
    if failure:
        return "latex2mathml 没能读懂这段 LaTeX（%s）" % failure
    try:
        root = _parse(mathml)
    except etree.XMLSyntaxError as exc:
        return _unparsable(exc)
    leftover = _leftover_macros(root)
    if leftover:
        words = ["一个反斜杠（\\ 后面接不上宏名）" if token == "\\" else token
                 for token in leftover]
        return "有一个宏或转义这一版认不出来：%s（进了 Word 会原样显示成文本）" % "、".join(words)
    if not _visible_text(root).strip() and not _SPACING_ONLY.match(source):
        return "这段公式里没有能显示的内容"
    return ""


def _unparsable(exc: etree.XMLSyntaxError) -> str:
    """MathML 读不通时那一句。

    走到这一格是实测出来的路子：正文里带**裸的** ``&`` 时 latex2mathml 交回 ``<mi>&</mi>``，
    而 XML 里的 ``&`` 必须写成 ``&amp;``，所以那一段根本不是合法 XML。渲染那一侧
    （:func:`_tree`）早就把这一步包成"读不通就退回原文"，体检这一侧原先没包——一条
    ``$a & b$`` 就把异常丢出 :func:`gui.validator._mathematical`，剩下的文件一个都没查。
    两句分开是因为**不能把 "&" 安在别的语法错上**：编出来的原因比不念更坏（人会去改一个
    根本没有 & 的公式）。
    """
    if "Entity" in str(exc):
        return ("公式正文里有一个没转义的 &（LaTeX 里它是表格对齐符，这一版转不成公式；"
                "要显示这个符号写成 \\&）")
    return "这一版读不通这条公式转出的 MathML（进了 Word 会是原样文本）"


def _stray_delimiter(source: str) -> str:
    if _STRAY_DOLLAR.search(source):
        return "公式正文里有一个多余的 $（多半是定界符少配了一个，回去看这一行的 $）"
    return ""


def _brace_problem(source: str) -> str:
    """扫一遍 ``{}``：``\\{`` 那种转义的是字面括号，不算结构。"""
    depth = 0
    extra_close = 0
    index = 0
    while index < len(source):
        char = source[index]
        if char == "\\":
            index += 2
            continue
        if char == "{":
            depth += 1
        elif char == "}":
            if depth:
                depth -= 1
            else:
                extra_close += 1
        index += 1
    if extra_close:
        return "括号不成对：多出来 %d 个 }" % extra_close
    if depth:
        return "括号不成对：有 %d 个 { 没关上" % depth
    return ""


def _missing_group(source: str) -> str:
    """要两组参数的宏只跟了一组（``\\frac{a}``）——latex2mathml 对这一种不报错，
    它交出 ``a``，于是 Word 里进的是一条只剩分子的残缺公式。"""
    for match in _MACRO_NAME.finditer(source):
        name = match.group(1)
        if name not in _TWO_ARG_MACROS:
            continue
        found = _groups_after(source, match.end())
        if 0 < found < 2:
            return "\\%s 只跟了 %d 组 {…}，它需要两组" % (name, found)
    return ""


def _groups_after(source: str, index: int) -> int:
    """从宏名之后数连续跟着的 ``{…}`` 组；``[…]`` 是可选下标（``\\sqrt[3]{x}``），跳过不计数。"""
    count = 0
    length = len(source)
    while index < length:
        while index < length and source[index].isspace():
            index += 1
        if index >= length:
            break
        char = source[index]
        if char == "[":
            closing = source.find("]", index)
            if closing < 0:
                break
            index = closing + 1
            continue
        if char != "{":
            break
        depth = 0
        closed = False
        while index < length:
            if source[index] == "\\":
                index += 2
                continue
            if source[index] == "{":
                depth += 1
            elif source[index] == "}":
                depth -= 1
                if depth == 0:
                    index += 1
                    closed = True
                    break
            index += 1
        if not closed:
            break
        count += 1
    return count


@lru_cache(maxsize=512)
def _convert(source: str) -> Tuple[str, str]:
    """LaTeX → MathML 文本。交回 ``(mathml, 失败原因)``，两者不会同时非空。

    缓存的理由：预览与校验在同一次点击里会把每条公式各问一遍，一份上百条公式的文件
    不该转两回；而 latex2mathml 对同一段输入的结果是确定的。
    """
    try:
        raw = latex2mathml.converter.convert(source)
    except Exception as exc:  # 这一族异常名字杂（NoAvailableTokensError / MissingEndError / ValueError…）
        return "", "%s%s" % (type(exc).__name__, str(exc).strip() and (": " + str(exc).strip()))
    return (raw or "").strip(), ""


def _parse(mathml: str) -> etree._Element:
    return etree.fromstring(mathml.encode("utf-8"))


def _tree(latex: str) -> etree._Element | None:
    """体检过、读得通的公式才走到这一步；读不通交回 None（调用方退回原文）。"""
    mathml, failure = _convert(latex)
    if failure or not mathml:
        return None
    try:
        return _parse(mathml)
    except etree.XMLSyntaxError:
        return None


def _leftover_macros(root: etree._Element) -> List[str]:
    r"""产出文本里还留着反斜杠就是这一版没认出来的写法（实测 ``\unknownmacro{x}`` → ``\unknownmacro``、
    ``a\3`` → ``\3``）。去重后按出现次序交回，每条都带着那个反斜杠，好直接拼进句子。"""
    found: List[str] = []
    for element in _elements(root):
        text = element.text or ""
        if "\\" not in text:
            continue
        for match in _RESIDUAL.finditer(text):
            token = "\\" + (match.group(1) or match.group(2) or "")
            if token not in found:
                found.append(token)
    return found


def _visible_text(root: etree._Element) -> str:
    return "".join((element.text or "") for element in _elements(root))


def _elements(root: etree._Element):
    for element in root.iter():
        if isinstance(element.tag, str):
            yield element


# ---------------------------------------------------------------- 渲染


def render_math(latex: str, display: bool = False) -> str:
    """一条公式 → Qt 认得的 HTML 片段（外层那个 ``span`` / ``div`` 由预览区包）。

    语法坏了就**原样退回 LaTeX** 并挂 ``math-broken`` 类：预览宁可念出原文并让它显眼，
    也不要把一条坏公式画成"看着像公式的空格"——那种空格在 Word 里会照样进一次。
    """
    source = (latex or "").strip()
    if not source:
        return ""
    if latex_problem(source):
        return "<span class=\"math-broken\">%s</span>" % escape(source)
    root = _tree(source)
    if root is None:
        return "<span class=\"math-broken\">%s</span>" % escape(source)
    return _node(root, display)


def _node(element: etree._Element, display: bool, trailing: str = "") -> str:
    tag = _local(element)
    handler = _RENDERERS.get(tag)
    if handler is None:
        return _render_default(element, display, trailing)
    return handler(element, display, trailing)


def _children(element: etree._Element, display: bool, trailing: str = "") -> str:
    return "".join(_node(child, display, trailing) for child in element)


def _leaf(element: etree._Element, display: bool, trailing: str = "") -> str:
    return escape((element.text or "") + trailing)


def _render_default(element: etree._Element, display: bool, trailing: str = "") -> str:
    """没特化的标签：文本 + 子节点，一个字符都不丢。

    丢掉是这里最坏的错——预览少一截而 Word 里却有，就成了"预览骗人"。
    """
    text = element.text or ""
    inner = _children(element, display, trailing)
    if not inner:
        return escape(text + trailing)
    return escape(text) + inner


def _render_frac(element: etree._Element, display: bool, trailing: str) -> str:
    parts = [child for child in element]
    if len(parts) != 2:
        return _children(element, display, trailing)
    # 分子分母内部一律按行内画：块级表格套块级表格在 Qt 里不是没试过，是试出来更难读。
    numerator = _node(parts[0], False, trailing)
    denominator = _node(parts[1], False, trailing)
    if not display:
        return "(%s)/(%s)" % (numerator, denominator)
    return ("<table class=\"math-frac\" cellspacing=\"0\" cellpadding=\"0\">"
            "<tr><td class=\"math-num\">%s</td></tr>"
            "<tr><td class=\"math-den\">%s</td></tr></table>") % (numerator, denominator)


def _render_sup(element: etree._Element, display: bool, trailing: str) -> str:
    parts = list(element)
    if len(parts) < 2:
        return _children(element, display, trailing)
    return _node(parts[0], display, trailing) + "<sup>%s</sup>" % _node(parts[1], False, "")


def _render_sub(element: etree._Element, display: bool, trailing: str) -> str:
    parts = list(element)
    if len(parts) < 2:
        return _children(element, display, trailing)
    return _node(parts[0], display, trailing) + "<sub>%s</sub>" % _node(parts[1], False, "")


def _render_sub_sup(element: etree._Element, display: bool, trailing: str) -> str:
    """``msubsup`` / ``munderover``：底数 + 下那一格 + 上那一格。

    ``\\sum_{i=1}^{n}`` 在 Word 里上下限摆在正上下方，预览里摆在右边一上一下——
    Qt 没有那个位置，而**上下限这两个数都在**才是预览要保住的。
    """
    parts = list(element)
    if len(parts) < 3:
        return _children(element, display, trailing)
    base = _node(parts[0], display, trailing)
    return "%s<sub>%s</sub><sup>%s</sup>" % (base, _node(parts[1], False, ""), _node(parts[2], False, ""))


def _render_over(element: etree._Element, display: bool, trailing: str) -> str:
    parts = list(element)
    if len(parts) < 2:
        return _children(element, display, trailing)
    base = _node(parts[0], display, trailing)
    accent = (_leaf(parts[-1], False, "") or "").strip()
    mark = _ACCENT_MARKS.get(accent)
    if mark is None:
        return base + escape(accent)
    if mark in _SPAN_MARKS:
        # 横线要盖住整串底数，所以作为 trailing 跟在每个叶子文本后面，而不是只跟在末尾。
        return _node(parts[0], display, trailing + mark)
    return base + mark


def _render_sqrt(element: etree._Element, display: bool, trailing: str) -> str:
    """根号写成 ``√( … )``：上横线那一档 Qt 真的一点墨都不出（实测过），
    画不出根号眉就宁可把括号念出来——被开方的范围还得看得见。"""
    return "\u221a(%s)" % _children(element, False, trailing)


def _render_root(element: etree._Element, display: bool, trailing: str) -> str:
    parts = list(element)
    if len(parts) < 2:
        return _render_sqrt(element, display, trailing)
    return "\u221a<sup>%s</sup>(%s)" % (_node(parts[1], False, ""), _node(parts[0], False, trailing))


def _render_table(element: etree._Element, display: bool, trailing: str) -> str:
    """矩阵 / cases / array：块级摆成表格，行内折成 ``1, 2 ; 3, 4``。

    行内不摆表格是实测的结论：Qt 把行内 ``<table>`` 提成块级，一句话就被它劈开。
    """
    rows = [[_node(cell, False, trailing) for cell in row if _local(cell) == "mtd"]
            for row in element if _local(row) == "mtr"]
    rows = [row for row in rows if row]
    if not rows:
        return _children(element, display, trailing)
    if not display:
        return " ; ".join(", ".join(row) for row in rows)
    return "<table class=\"math-matrix\" cellspacing=\"0\" cellpadding=\"0\">%s</table>" % (
        "".join("<tr>%s</tr>" % "".join("<td>%s</td>" % cell for cell in row) for row in rows))


def _render_space(element: etree._Element, display: bool, trailing: str) -> str:
    return "\u00a0"


def _render_empty(element: etree._Element, display: bool, trailing: str) -> str:
    return ""


_RENDERERS: Dict[str, Callable[[etree._Element, bool, str], str]] = {
    "mi": _leaf,
    "mn": _leaf,
    "mo": _leaf,
    "mtext": _leaf,
    "ms": _leaf,
    "mfrac": _render_frac,
    "msup": _render_sup,
    "msub": _render_sub,
    "msubsup": _render_sub_sup,
    "munder": _render_sub,
    "munderover": _render_sub_sup,
    "mover": _render_over,
    "msqrt": _render_sqrt,
    "mroot": _render_root,
    "mtable": _render_table,
    "mspace": _render_space,
    "annotation": _render_empty,
    "annotation-xml": _render_empty,
}


def _local(element: etree._Element) -> str:
    """``{uri}mfrac`` → ``mfrac``。"""
    tag = element.tag
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]
