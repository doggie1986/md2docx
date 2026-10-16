# -*- coding: utf-8 -*-
"""命令行入口：把 ``input.md`` 换成公式为原生 OMML 的 ``output.docx``。

这一层只干四件事，逻辑一律不在这里重写：

1. 收参数（``--xslt-path`` / ``--body-font`` / ``--body-size`` / ``--heading-font`` / ``--strict`` …）。
2. 定位 Microsoft 的 ``MML2OMML.XSL``（:func:`formula_converter.locate_xslt`）。
3. 解析 → 转换 → 生成（``md_parser`` + ``FormulaConverter`` + ``docx_builder``）。
4. 把回执念出来，并**读回落盘的文件**数一遍 ``m:oMath``——"公式在 Word 里是原生对象"
   这句话的证据在文件里，不在内存里。

合并那一族（``--merge`` / ``--toc`` / ``--cover`` / ``--title`` / ``--page-break`` /
``--source-label`` / ``--no-demote``）里"谁在前谁在后、标题降几级、哪一块前面换页、封面写什么字"
**全部问 :func:`gui.merger.merge_entries`**——那是界面（``gui/tasks.py``）已经在用的同一份策略层，
命令行再拼一套就会出现"界面上合并对了、脚本里合并不一样"。``gui/merger.py`` 只 import
``md_parser`` 与 ``docx_builder``、不碰 Qt，所以 ``import gui.merger`` 不需要界面在场。

一条不变式：**一个输入且合并那几格开关一个都没打 = 走改动之前那一条路**，产物逐字节相同
（``tests/test_pipeline.py`` 钉这一条）。合并层的开关全关时它交回 ``layout=None``、builder 因此
一个字都不多写，但回执里"输入：xxx（utf-8，N 字节）"那一行读的是原文件的 ``encoding``，
而并树之后那份 stats 只剩数值格（``gui/merger.py`` 的账）——所以这一路压根不进并树那段代码，
别"顺手统一"掉。

退出码分档是有意为之：脚本与批处理要能分清"没装 Office（样式表找不到）"和
"稿子里有坏公式（--strict）"，两类都回 1 的话，值班的人只能重新跑一遍才知道差在哪。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional, Sequence

import docx_builder as db
import md_parser as mp
from formula_converter import (
    FormulaConversionError,
    FormulaConverter,
    XsltNotFoundError,
    locate_xslt,
)
from gui import merger as mg

PROGRAM = "md2docx"
VERSION = "1.0"

# 退出码：一档一个原因，写在 --help 末尾给人查
EXIT_OK = 0
EXIT_USAGE = 2  # argparse 自己那一档（参数写错、少参数）
EXIT_XSLT = 3  # 找不到 MML2OMML.XSL
EXIT_INPUT = 4  # 读不进输入（文件没有 / 编码认不出 / Markdown 结构不成）
EXIT_FORMULA = 5  # --strict 且有条公式转不出来
EXIT_WRITE = 6  # 写不出输出（没权限 / 磁盘满 / Word 正占着这个文件）


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROGRAM,
        description="把带 LaTeX 公式与 Markdown 表格的 .md 转成 .docx，公式为 Word 原生可编辑的 OMML。",
        epilog=(
            "退出码：%d 成功｜%d 参数错｜%d 找不到样式表｜%d 输入读不进｜%d 有坏公式（--strict）｜%d 写不出文件\n"
            "示例：python md2docx.py 报告.md 报告.docx --body-font 宋体 --body-size 12\n"
            "      python md2docx.py 甲.md 乙.md 丙.md 合并.docx --cover --title 投标稿\n"
            "      （合并时目录、分页、来源标注默认都开，不要就写 --no-toc / --no-page-break / "
            "--no-source-label；标题默认降一级，不要就写 --no-demote）\n"
            % (EXIT_OK, EXIT_USAGE, EXIT_XSLT, EXIT_INPUT, EXIT_FORMULA, EXIT_WRITE)
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "input",
        nargs="+",
        metavar="INPUT",
        help="输入的 .md 文件；给两个以上就是把它们并成一份（顺序照你敲的顺序）",
    )
    parser.add_argument("output", help="输出的 .docx 文件（父目录不存在会先建）")
    parser.add_argument(
        "--merge",
        action="store_true",
        help="按合并稿的口径出这一份（一份文件也让目录/封面/来源标注生效）；给两个以上输入时本来就合并",
    )
    parser.add_argument(
        "--toc",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="在开头插一个 Word 的目录域（TOC）。默认：合并时开、单独一份关；"
             "--no-toc 撤掉。页码要 Word 排版之后才数得出来，文件里先放占位",
    )
    parser.add_argument(
        "--cover",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="在开头插一页封面（标题 + 日期 + 共几个文件）。默认关，--cover 打开",
    )
    parser.add_argument(
        "--title",
        default="",
        help="封面标题那一行写什么（不填就用输出文件名的主干）",
    )
    parser.add_argument(
        "--page-break",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="每个文件之前插一个分页符。默认：合并时开、单独一份关；--no-page-break 撤掉",
    )
    parser.add_argument(
        "--source-label",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="每个文件之前标一行「来源：xxx.md」。默认：合并时开、单独一份关；--no-source-label 撤掉",
    )
    parser.add_argument(
        "--no-demote",
        dest="demote",
        action="store_false",
        help="合并时「不要」给每个文件的标题降一级（默认降：# 变 ##，好让整份稿子只有一级标题）",
    )
    parser.add_argument(
        "--xslt-path",
        default=None,
        help="MML2OMML.XSL 的位置（不填就自动找：当前目录 → 本程序目录 → Office 常见安装路径）",
    )
    parser.add_argument(
        "--body-font",
        default=db.BODY_FONT_DEFAULT,
        help="正文字体，中文务必选中文字体（默认 %(default)s）",
    )
    parser.add_argument(
        "--body-size",
        type=_positive_float,
        default=db.BODY_SIZE_DEFAULT,
        help="正文字号（磅），标题按它往上加档（默认 %(default)s）",
    )
    parser.add_argument(
        "--heading-font",
        default=db.HEADING_FONT_DEFAULT,
        help="标题字体（默认 %(default)s）",
    )
    parser.add_argument(
        "--code-font",
        default=db.CODE_FONT_DEFAULT,
        help="代码段与代码块的等宽字体（默认 %(default)s）",
    )
    parser.add_argument(
        "--formula-font",
        default="Cambria Math",
        help="公式的字形声明（OMML 里每个 m:r 都补这一格，默认 %(default)s）",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="有条公式转不出来就当场失败，而不是退回原文 + 黄色底纹",
    )
    parser.add_argument(
        "--encoding",
        default=None,
        help="强制输入文件的编码（默认自动试 utf-8-sig / utf-8 / gb18030）",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="多打几行：块树统计、每一条失败公式的原文",
    )
    parser.add_argument("--version", action="version", version="%(prog)s " + VERSION)
    return parser


def _positive_float(text: str) -> float:
    """字号只接受正数：负数与 0 会写进 ``w:sz`` 变成 Word 拒开的包。"""
    try:
        value = float(text)
    except ValueError:
        raise argparse.ArgumentTypeError("字号得是个数（现在写的是「%s」）" % text)
    if not value > 0:
        raise argparse.ArgumentTypeError("字号得大于 0（现在写的是「%s」）" % text)
    return value


def run(argv: Optional[Sequence[str]] = None) -> tuple[int, str]:
    """跑一遍，回 ``(退出码, 给人读的那一段)``。命令行与测试都走这一个出口。"""
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    sources = [Path(item) for item in args.input]
    target = Path(args.output)

    for source in sources:
        if source.resolve() == target.resolve() and target.is_file():
            return EXIT_USAGE, "输入与输出是同一个文件：%s；.md 不该被 .docx 顶掉" % source

    # 先解析再定位样式表：稿子本身读不进就没必要让人以为问题是 Office 引起的
    entries: list[mg.MergeEntry] = []
    for position, source in enumerate(sources):
        code, message, parsed = _read_one(source, args.encoding)
        if parsed is None:
            if len(sources) > 1:
                message = "第 %d 个文件（%s）：%s" % (position + 1, source, message)
            return code, message
        entries.append((source, parsed))

    switching = _switches(args, len(sources))
    if switching is None:
        parsed, layout = entries[0][1], None
        merge_lines: list[str] = []
    else:
        plan = mg.merge_entries(entries, switching, title=args.title or target.stem)
        parsed, layout = plan.parsed, plan.layout
        merge_lines = _merge_lines(switching, len(sources), plan, args.title or target.stem)

    try:
        xslt = locate_xslt(args.xslt_path)
    except XsltNotFoundError as exc:
        # 不另拼一段"怎么办"：XsltNotFoundError 自己就念出"查过哪几个位置 + 三种解决办法"，
        # 这里再拼一份就会出现"两处说法不一样的下一步"（同一件事两个来源）
        return EXIT_XSLT, str(exc)

    options = db.BuilderOptions(
        body_font=args.body_font,
        body_size=args.body_size,
        heading_font=args.heading_font,
        code_font=args.code_font,
    )
    try:
        converter = FormulaConverter(xslt, formula_font=args.formula_font)
        report = db.DocxBuilder(converter, options, strict=args.strict).build(parsed, layout)
    except FormulaConversionError as exc:
        return EXIT_FORMULA, "--strict：有条公式转不出来，已中止且没有写出文件。\n%s" % exc
    except XsltNotFoundError as exc:
        return EXIT_XSLT, str(exc)
    except OSError as exc:
        return EXIT_XSLT, "样式表 %s 读不出来：%s\n换一份完整的（从 Office 安装目录重新拷一次）" % (xslt, exc)

    try:
        report.save(target)
    except PermissionError as exc:
        return EXIT_WRITE, (
            "写不出 %s：没有权限，或者这个文件正被 Word 打开着（把 Word 关掉再跑一次）。\n%s"
            % (target, exc)
        )
    except OSError as exc:
        return EXIT_WRITE, "写不出 %s：%s" % (target, exc)

    if len(sources) > 1:
        lines = ["输入（并成一份的顺序）："]
        lines.extend(
            "  %d. %s（%s）" % (position + 1, source, _encoding_of(source_parsed))
            for position, (source, source_parsed) in enumerate(entries)
        )
    else:
        origin = entries[0][1]
        lines = [
            "输入：%s（%s，%d 字节）"
            % (sources[0], _encoding_of(origin), origin.stats.get("encoding_bytes", 0))
        ]
    lines.extend(report.lines())
    lines.extend(merge_lines)
    lines.append(_verification_line(target, report))
    lines.append("输出：%s" % target.resolve())
    if args.verbose:
        lines.append("块树统计：%s" % parsed.stats)
    else:
        # 非 -v 时也把"这一份稿子有多少条公式"念一遍：回执里的数只统计真落笔的那些
        lines.append("解析计数：块 %d｜标题 %d｜行内公式 %d｜显示式 %d｜代码块 %d｜表格 %d" % (
            parsed.stats.get("blocks", 0),
            parsed.stats.get("headings", 0),
            parsed.stats.get("formulas_inline", 0),
            parsed.stats.get("formulas_display", 0),
            parsed.stats.get("code_blocks", 0),
            parsed.stats.get("tables", 0),
        ))
    return EXIT_OK, "\n".join(lines)


def _read_one(source: Path, encoding: Optional[str]) -> tuple[int, str, Optional[mp.ParseResult]]:
    """读一个 .md，交回 ``(这一档退出码, 给人读的那句, 解析结果)``。

    读不进时那四句原话一个字都没改：脚本与批处理里可能已经在按关键字判这一行。
    多个输入时"第 N 个文件"那个前缀由调用方拼，这里只管这一个文件本身。
    """
    try:
        return EXIT_OK, "", mp.parse_file(source, encoding)
    except FileNotFoundError as exc:
        return EXIT_INPUT, str(exc), None
    except mp.MarkdownSyntaxError as exc:
        return EXIT_INPUT, "Markdown 结构解析失败：%s" % exc, None
    except UnicodeDecodeError as exc:
        return EXIT_INPUT, "这个 .md 读不进：%s" % exc, None
    except OSError as exc:
        return EXIT_INPUT, "读 %s 失败：%s" % (source, exc), None


def _switches(args: argparse.Namespace, count: int) -> Optional[mg.MergeSwitches]:
    """把人给的格翻成 :class:`gui.merger.MergeSwitches`；**该走老路时交回 None**。

    老路 = 一个输入且合并那几格一格都没打。这时产物与这一轮改动之前逐字节相同，
    回执上"输入：xxx（utf-8，N 字节）"读的那份 ``encoding`` 也还在（并树之后 ``stats``
    只剩数值格，见 :func:`gui.merger.merge_entries`）。

    三态（``None`` = 人没表态）的默认档跟着界面走：**合并时目录 / 分页 / 来源标注默认开，
    单独一份默认关**（界面上那三个勾是"合并选项"，单独输出走 ``gui.merger.PLAIN``）。
    封面两边都默认关——界面与命令行都是"多要一页纸"这个额外决定，不是合并本身含的意思。
    """
    merging = count > 1 or bool(args.merge)

    def decided(value: Optional[bool], default: bool) -> bool:
        return default if value is None else value

    if not merging and not any((args.toc, args.cover, args.page_break, args.source_label)):
        return None
    return mg.MergeSwitches(
        demote=bool(args.demote),
        page_break=decided(args.page_break, merging),
        source_label=decided(args.source_label, merging),
        toc=decided(args.toc, merging),
        cover=decided(args.cover, False),
    )


def _merge_lines(switching: mg.MergeSwitches, count: int, plan: mg.MergePlan,
                 title: str) -> list[str]:
    """把"这一次合并到底做了什么"念出来——开关是人的决定，回执不复述就等于没说。

    ``demote`` 那一格念的是**真降了没有**（``gui.merger`` 的口径是一个文件不降），
    不是开关写了什么：一份稿子念"标题降级 开"会让人以为自己的 ``#`` 没了。
    """
    def word(flag: bool) -> str:
        return "开" if flag else "关"

    lines = [
        "合并说明：标题降级 %s｜分页符 %s｜来源标注 %s｜目录 %s｜封面 %s"
        % (
            word(switching.demote and count > 1),
            word(switching.page_break),
            word(switching.source_label),
            word(switching.toc),
            word(switching.cover),
        )
    ]
    if switching.source_label:
        lines.append("  · 来源标注那 %d 行是正文斜体，不进目录，也不计进下面那行「解析计数」" % count)
    if switching.cover:
        lines.append("  · 封面标题：%s｜共 %d 个文件" % (title, count))
    lines.extend("  · " + text for text in plan.notes())
    return lines


def _encoding_of(parsed: mp.ParseResult) -> str:
    return str(parsed.stats.get("encoding", "utf-8"))


def _verification_line(target: Path, report: db.BuildReport) -> str:
    """读回落盘的文件数原生公式对象——这一句才是"OMML 而不是图片"的证据。

    ``count_omml`` 第一个数是**全部** ``m:oMath``（显示式外包的 ``m:oMathPara`` 里那一个也算），
    所以它比的是 ``report.math_ok``（行内 + 显示式的总条数），不是行内那一条账。
    """
    try:
        math_objects, display_blocks = db.count_omml(target)
    except (OSError, KeyError) as exc:
        return "自检：读不回 %s（%s）；文件已经写出，但公式有没有落进去没核出来" % (target, exc)
    size = target.stat().st_size if target.is_file() else 0
    if math_objects == report.math_ok:
        return "自检：文件里有 %d 个 m:oMath（其中显示式 %d 个 m:oMathPara），与内存里转成的条数一致｜%d 字节" % (
            math_objects, display_blocks, size
        )
    return (
        "自检：不一致——内存里转成 %d 条，文件里数到 %d 个 m:oMath（显示式 %d 个）。"
        "请把这份 .md 反馈给作者，不要直接交这一份 docx" % (report.math_ok, math_objects, display_blocks)
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    code, text = run(argv)
    # 控制台的代码页可能是 GBK，稿子里的字符不一定进得去；宁可打问号也不要为打印崩掉
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except Exception:
                pass
    print(text)
    if code not in (EXIT_OK,):
        print("（退出码 %d）" % code, file=sys.stderr)
    return code


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
