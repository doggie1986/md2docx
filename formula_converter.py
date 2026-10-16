# -*- coding: utf-8 -*-
"""公式转换层：LaTeX → MathML → OMML。

这一层的边界很窄，只做三件事：

1. ``locate_xslt()``  —— 自动找到 Microsoft 官方的 ``MML2OMML.XSL`` 样式表
   （命令行参数 → 当前目录 → 脚本所在目录 → Office 常见安装路径）。
2. ``FormulaConverter.convert()`` —— 把一段 LaTeX 交给 latex2mathml 转成 MathML，
   再用 XSLT 转成 Word 原生的 OMML（``m:oMath`` / 块级时外包 ``m:oMathPara``）。
3. OMML 的"落笔前体检"（``_polish``）—— 补 ``Cambria Math`` 字体声明、
   给每个 ``m:t`` 补 ``xml:space="preserve"``，这两步不写 Word 会出问题。

这里**不认识** Markdown，也**不认识** python-docx：产出物是一段 lxml 元素树，
由 docx_builder 负责挂进 Word 段落或表格单元格。
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import latex2mathml.converter
from lxml import etree

# --------------------------------------------------------------------------
# 命名空间：OMML 是 Office Math ML，样式表输出用的是这两个固定 URI
# --------------------------------------------------------------------------
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
#: MathML 的命名空间：latex2mathml 的输出带它，XSLT 之后的 OMML 里没有它，
#: 但 docx_builder 复位前缀时要把三个都登记，否则 document.xml 里出现 ns0:。
MATHML_NS = "http://www.w3.org/1998/Math/MathML"

M = "{%s}" % MATH_NS
W = "{%s}" % WORD_NS

#: 样式表文件名（Windows 上大小写不敏感，这里按微软的写法）
XSLT_FILENAME = "MML2OMML.XSL"

#: Office 安装目录里可能出现样式表的相对位置。
#: root/Office16 是即点即用（Office 2016/2019/365）的布局，
#: Office16 / Office15 是老 MSI 安装的布局，两套都试。
OFFICE_SUBDIRS: tuple[str, ...] = (
    r"Microsoft Office\root\Office16",
    r"Microsoft Office\root\Office15",
    r"Microsoft Office\Office16",
    r"Microsoft Office\Office15",
    r"Microsoft Office 15\root\Office15",
)


class XsltNotFoundError(Exception):
    """找不到 MML2OMML.XSL。

    ``searched`` 是已经找过的位置列表，报错时要原样念出来——
    只说"找不到文件"，用户不知道该去哪台机器上拷一份、也不知道该用什么参数指过来。
    """

    def __init__(self, searched: Iterable[Path]) -> None:
        self.searched: list[Path] = list(searched)
        lines = "\n".join("    %d. %s" % (i, p) for i, p in enumerate(self.searched, 1))
        super().__init__(
            "找不到 Microsoft 的公式样式表 {name}。已查找过：\n{lines}\n"
            "解决办法（任选一种）：\n"
            "  · 用 --xslt-path 指定它的完整路径，例如\n"
            "      --xslt-path \"C:\\Program Files\\Microsoft Office\\root\\Office16\\{name}\"\n"
            "  · 把 {name} 复制到当前目录，或复制到本脚本所在目录\n"
            "  · 在装有 Microsoft Office（2016 及以上）的机器上运行本程序".format(
                name=XSLT_FILENAME, lines=lines
            )
        )


class FormulaConversionError(Exception):
    """单条 LaTeX 转换失败（只在 --strict 模式下会中断整个转换）。"""


def _program_files_roots() -> list[Path]:
    """取本机所有 Program Files 目录（64 位与 32 位各一份，去重保序）。"""
    roots: list[Path] = []
    for var in ("PROGRAMFILES", "PROGRAMFILES(X86)", "ProgramW6432", "ProgramFiles6432"):
        value = os.environ.get(var)
        if not value:
            continue
        path = Path(value)
        if path not in roots:
            roots.append(path)
    return roots


def default_search_dirs() -> list[Path]:
    """不指定 --xslt-path 时，按优先级要查的目录。

    顺序就是"离用户越近越优先"：当前目录（用户为这份文件放的那一份）→
    脚本所在目录（跟着项目一起拷走的那一份）→ 系统里的 Office 安装位置。
    """
    dirs: list[Path] = [Path.cwd(), Path(__file__).resolve().parent]
    for root in _program_files_roots():
        for sub in OFFICE_SUBDIRS:
            dirs.append(root / sub)
    return dirs


def locate_xslt(cli_path: Optional[str] = None) -> Path:
    """定位 MML2OMML.XSL，返回存在的第一个路径。

    参数 ``cli_path`` 是 --xslt-path 交来的值，优先级最高，且**不做体检**：
    用户指的路径就算指向一个奇怪位置也照用（那是他的决定），只是文件不存在
    或读不出来时要当场说清，不能拖到 XSLT 编译那一刻才报一句 XML 解析错误。
    """
    candidates: list[Path] = []
    if cli_path:
        candidates.append(Path(cli_path).expanduser())
    else:
        candidates.extend(dir / XSLT_FILENAME for dir in default_search_dirs())

    for path in candidates:
        if path.is_file():
            return path
        # 用户显式指了目录也认（--xslt-path C:\\Office16 而不是完整文件名）
        if cli_path and path.is_dir() and (path / XSLT_FILENAME).is_file():
            return path / XSLT_FILENAME

    if cli_path:
        # 指定了路径却没有：这是最容易被读成"程序坏了"的一种失败，单独说清。
        raise XsltNotFoundError([Path(cli_path).expanduser()])
    raise XsltNotFoundError(candidates)


@dataclass
class ConversionResult:
    """一条公式的转换回执。

    ``omml`` 为 None 就是没转成，此时 ``latex`` 仍是原文、由 docx_builder 决定
    退回成文本还是中断。**失败必须带一句中文原因**（``problem``），
    只报"转换失败"等于没报。
    """

    latex: str
    display: bool
    omml: Optional[etree._Element] = None
    problem: str = ""

    @property
    def ok(self) -> bool:
        return self.omml is not None

    def failure_report(self) -> str:
        """给报告用的一行：原文截断 + 原因。"""
        text = self.latex if len(self.latex) <= 60 else self.latex[:57] + "..."
        return "%s  ←  %s" % (text, self.problem)


class FormulaConverter:
    """把 LaTeX 变成可以塞进 Word 的 OMML 元素。

    XSLT 在构造时编译一次并复用：一份 .md 里可能有上百条公式，
    每条都重新解析一遍样式表是纯浪费。
    """

    def __init__(self, xslt_path: Path | str, formula_font: str = "Cambria Math") -> None:
        self.xslt_path = Path(xslt_path)
        self.formula_font = formula_font
        if not self.xslt_path.is_file():
            raise XsltNotFoundError([self.xslt_path])
        try:
            parsed = etree.parse(str(self.xslt_path))
        except etree.XMLSyntaxError as exc:
            # 拿到的是路径下的一个不是 XML 的文件（多半是拷坏了或指错了对象）
            raise XsltNotFoundError([self.xslt_path]) from exc
        try:
            self._xslt = etree.XSLT(parsed)
        except etree.XSLTParseError as exc:
            raise RuntimeError(
                "样式表 %s 不是可用的 XSLT（%s）。请确认它是 Microsoft 的 %s。"
                % (self.xslt_path, exc, XSLT_FILENAME)
            ) from exc

    # ---------------------------------------------------------------- 对外接口
    def convert(self, latex: str, display: bool = False) -> ConversionResult:
        """转换一条公式，**不抛异常**（除非样式表本身有问题，那在构造时已经拦掉了）。

        默认这一条失败不影响别的公式出表，所以失败以回执形式交回，
        由上层汇总成"第 N 条：原因"。
        """
        source = (latex or "").strip()
        if not source:
            return ConversionResult(latex=source, display=display, problem="公式是空的（定界符之间没内容）")
        try:
            omml = self.convert_or_raise(source, display)
        except FormulaConversionError as exc:
            return ConversionResult(latex=source, display=display, problem=str(exc))
        except Exception as exc:  # latex2mathml 对畸形输入会抛各种内部异常
            return ConversionResult(
                latex=source,
                display=display,
                problem="%s: %s" % (type(exc).__name__, str(exc).strip() or "（无更多信息）"),
            )
        return ConversionResult(latex=source, display=display, omml=omml)

    def convert_or_raise(self, latex: str, display: bool = False) -> etree._Element:
        """转换一条公式，失败就抛 :class:`FormulaConversionError`（--strict 走这一条）。"""
        mathml = self._to_mathml(latex)
        omath = self._xslt_to_omml(mathml, latex)
        return self._polish(omath, display)

    def count_text_nodes(self, omml: etree._Element) -> int:
        """数一下这段 OMML 里有几个 ``m:t``（自检用，不参与排版）。"""
        return len(omml.findall(".//" + M + "t"))

    # ---------------------------------------------------------------- 内部三步
    def _to_mathml(self, latex: str) -> etree._Element:
        """第一步：LaTeX → MathML（latex2mathml），第二步的输入必须是合法 XML。"""
        try:
            raw = latex2mathml.converter.convert(latex)
        except Exception as exc:
            raise FormulaConversionError(
                "latex2mathml 没能读懂这段 LaTeX（%s: %s）" % (type(exc).__name__, str(exc).strip() or "无信息")
            ) from exc
        if not raw or not raw.strip():
            raise FormulaConversionError("latex2mathml 返回了空结果，通常是这段 LaTeX 不成结构")
        try:
            # latex2mathml 的输出根元素是 <math …>，但没有 MathML 的 DTD，直接解析即可
            return etree.fromstring(raw.encode("utf-8"))
        except etree.XMLSyntaxError as exc:
            raise FormulaConversionError("latex2mathml 产出的 MathML 不是合法 XML（%s）" % exc) from exc

    def _xslt_to_omml(self, mathml: etree._Element, latex: str) -> etree._Element:
        """第二步：MathML → OMML（MML2OMML.XSL）。"""
        try:
            result = self._xslt(mathml)
        except etree.XSLTApplyError as exc:
            raise FormulaConversionError("套用微软样式表时出错（%s）" % exc) from exc
        root = result.getroot()
        # XSLT 理论上产出 m:oMath；产出的东西不对时宁可报错，也不要塞一个
        # "看起来是公式其实是一张空框"的元素进 Word。
        if root is None or not isinstance(root.tag, str):
            raise FormulaConversionError("样式表没有产出 OMML 元素")
        if root.tag != M + "oMath":
            raise FormulaConversionError(
                "样式表产出的是 <%s>，不是 m:oMath（样式表版本对不上？）" % _local(root.tag)
            )
        return root

    # ------------------------------------------------------------ 落笔前体检
    def _polish(self, omath: etree._Element, display: bool) -> etree._Element:
        """给 OMML 补两处声明，并按 display 决定要不要包成 ``m:oMathPara``。

        * ``m:r`` 里补 ``w:rPr/w:rFonts`` 的 Cambria Math：样式表并不总是带这个，
          缺了 Word 会用正文字体画公式，上下标位置跟着算错。
        * ``m:t`` 补 ``xml:space="preserve"``：不写 Word 会把公式里的空格吃掉，
          `\beta\alpha` 这种会挤成一坨。
        * 块级公式外包 ``m:oMathPara``：Word 认这是"独占一段的显示式公式"，
          分式、上下限的排版档次不一样（``\sum`` 的上下限会跑到正上方/正下方）。
        """
        for run in omath.iter(M + "r"):
            self._ensure_math_font(run)
        for text in omath.iter(M + "t"):
            text.set("{%s}space" % XML_NS, "preserve")
        if not display:
            return omath
        para = etree.Element(M + "oMathPara")
        para.append(omath)
        return para

    def _ensure_math_font(self, run: etree._Element) -> None:
        """在一个 ``m:r`` 上补 Cambria Math 的 ``w:rPr``（已有就只补缺的字体属性）。"""
        existing = run.find(W + "rPr")
        if existing is None:
            rpr = etree.Element(W + "rPr")
            # OMML 架构里 m:rPr 必须在前、w:rPr 在后，位置摆错 Word 会拒这个包
            first = run[0] if len(run) else None
            index = 1 if first is not None and first.tag == M + "rPr" else 0
            run.insert(index, rpr)
        else:
            rpr = existing
        fonts = rpr.find(W + "rFonts")
        if fonts is None:
            fonts = etree.Element(W + "rFonts")
            rpr.insert(0, fonts)
        for attr in ("ascii", "hAnsi", "cs"):
            if fonts.get(W + attr) is None:
                fonts.set(W + attr, self.formula_font)


def _local(tag: str) -> str:
    """把 ``{uri}oMath`` 形式的限定名收成 ``前缀:oMath`` 给人读。"""
    if not isinstance(tag, str) or not tag.startswith("{"):
        return str(tag)
    uri, _, name = tag[1:].partition("}")
    if uri == MATH_NS:
        return "m:" + name
    if uri == WORD_NS:
        return "w:" + name
    return "{%s}%s" % (uri, name)
