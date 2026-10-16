# -*- coding: utf-8 -*-
"""转换与校验的执行体 + 那一根后台线程。

边界写死在这里：界面只管"什么时候开始、进度念到多少、日志哪一行、报告摆哪几格"，
真要读稿子、转公式、写 .docx 的那一步全在这个模块，而且用的是命令行那一套引擎
（``md_parser.parse_file`` → ``FormulaConverter`` → ``docx_builder.DocxBuilder``）。
**界面上不另算一遍数**：日志念的份数、文件里落的公式条数、报告对话框上的统计读的都是
同一次 ``BuildReport``——"屏幕上对得上、盘上对不上"这一类错得在结构上不成立。

四条不是显而易见的口径：

* **一次跑批只定位、编译一次样式表**：``FormulaConverter`` 在构造时就把 ``MML2OMML.XSL``
  编译好，一批几十个文件复用同一个实例与同一个 ``DocxBuilder``（实测单份 ~60 ms）。
  样式表找不到要**在动手之前**拒——写了一半才发现样式表没了是最难看的失败。
* **合并与分组不是"再拼一次文档"**：``ParseResult`` 是三格纯数据，所以"把几个文件并成一份"
  就是把 ``blocks`` 串起来、``warnings`` 接上、计数键相加，然后交同一条 ``build()``。
  并树这件事、连同标题降级 / 来源标注 / 封面 / 目录域那四个开关怎么落成声明，
  都在 :mod:`gui.merger` 那一层——这一层只负责"照计划逐个读、把计划交回来的那份声明写给 builder"。
  出几份、叫什么名、哪几个文件进哪一份仍然只问 :meth:`Options.output_plan`。
* **坏了不拦整批**：某一个文件读不进就记一条问题、接着走下一个（与命令行"一条公式转不出
  就用原文加黄底标出、不中止"同一策略）。取消是协作式的：手上那一个文件走完才停，
  已经写出去的文件不回滚。
* **统计里的公式数来自落盘那一份**：``report.math_ok`` 是真进文档的条数，解析阶段数出来的
  条数只念在日志那一行——写失败的那一份不计数，否则"转成了 18 条"其实盘上一条都没有。
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional, Sequence

from PySide6.QtCore import QThread, Signal

import docx_builder as db
import md_parser as mp
from formula_converter import FormulaConverter, XsltNotFoundError, locate_xslt
from gui import merger as mg
from gui import validator

__all__ = ["Options", "ConvertWorker", "ACTION_CONVERT", "ACTION_VALIDATE",
           "convert_paths", "validate_paths", "COLUMNS_OF_OUTPUT", "COLUMNS_OF_PROBLEM"]

ACTION_CONVERT = "convert"
ACTION_VALIDATE = "validate"

#: 每读完一个文件之间那一拍。**生产恒 0**：真转换一份稿子只要 60 ms 上下，快得没法在
#: "跑了一半"那一刻点中「取消这一批」，而那条路径必须能测。测试把这一格调大就拦在批次中间。
PAUSE_PER_FILE_SECONDS = 0.0

#: 一份输出里"转不出来的公式"最多摆几行进报告表（行数是计全的，采样只是显示闸门）。
MAX_FAILURE_ROWS = 5

#: Windows 文件名里不让出现的字符（组名直接当文件名用，所以这一关写在出计划的那一处）。
_BAD_NAME_CHARS = '<>:"/\\|?*'

Log = Callable[[str], None]
#: 进度回调：``（百分比 0-100, 给人读那一句）``。总步数由执行体自己算，界面只管摆进度条和状态栏。
Progress = Callable[[int, str], None]
#: 带档位的那一路日志：``（整行文本, 颜色档位）``。它是 :data:`Log` 之外**多出来的一根**而不是
#: 把 :data:`Log` 改了形状——四期定下的那四个信号是主窗口已经连好的线，这一期只加不改。
#: 没传这一根就退化成普通日志，文本一个字都不差（``tests/test_gui_smoke.py`` 里那些只交
#: ``log`` 的调用方因此不用跟着改）。
Logged = Callable[[str, str], None]


@dataclass(frozen=True)
class Options:
    """右侧那块面板的取值，一次跑批之前照抄一份下来（跑一半时人改了面板不该影响正在跑的这批）。

    ``groups`` 是 :meth:`gui.file_panel.FileListWidget.group_plans` 交回的那一份
    ``[(组名, [路径, …]), …]``——**面板负责"谁归哪一组"，这里只负责"一组出哪一份"**，
    所以 ``Options`` 不知道列表控件，``file_panel`` 也不知道输出目录。
    """

    output_dir: Path
    merged: bool
    body_font: str
    body_size: float
    heading_font: str
    with_toc: bool = True
    page_break: bool = True
    source_label: bool = True
    cover_page: bool = False
    grouped: bool = False
    groups: Sequence[Sequence[object]] = ()

    def mode_name(self) -> str:
        """三档输出模式对人说的那三个词。"""
        if self.grouped:
            return "分组输出"
        return "合并输出" if self.merged else "单独输出"

    def summary(self) -> str:
        """一行念清"这批打算怎么出"，日志和状态栏都用它。"""
        merge_flags = [name for name, on in (
            ("生成目录", self.with_toc),
            ("插入分页符", self.page_break),
            ("标注源文件名", self.source_label),
            ("生成封面页", self.cover_page),
        ) if on]
        tail = ""
        if self.merged or self.grouped:
            # 分组输出时这四个开关管的是"每一份"，所以说"每组选项"而不是"合并选项"。
            label = "合并选项" if self.merged else "每组选项"
            tail = "｜%s：%s" % (label, "、".join(merge_flags) if merge_flags else "全没勾")
        return ("%s｜正文 %s %sg｜标题 %s｜输出目录 %s%s" % (
            self.mode_name(),
            self.body_font, _size_text(self.body_size), self.heading_font,
            self.output_dir, tail))

    def merge_switches(self) -> mg.MergeSwitches:
        """面板上那四个勾 + "这是合并稿"这件事，翻成 :mod:`gui.merger` 认识的一份声明。

        **单独输出走** :data:`gui.merger.PLAIN`：那一路一份文件就是完整的一级结构，
        降标题、加来源行都不该有（一个勾都没打时 ``with_toc`` 之类仍是 True，
        照它翻就会出现"单独输出也凭空降了一级"）。
        ``demote`` 不是面板上的开关而是"这一份由几个文件并成"的声明，所以在这里给 True。
        """
        if not (self.merged or self.grouped):
            return mg.PLAIN
        return mg.MergeSwitches(
            demote=True,
            page_break=self.page_break,
            source_label=self.source_label,
            toc=self.with_toc,
            cover=self.cover_page,
        )

    def output_plan(self, paths: Sequence[Path]) -> list:
        """这批文件打算写成哪几份：``[(这一份的叫法, 目标路径, 归这一份的文件), …]``。

        **出几份、每份叫什么名、哪几个文件进哪一份，全在这一个函数里**——主窗口念它、
        执行体照它逐个读逐个写。界面里不另判一套，
        否则"屏幕上念三份、盘上出来两份"这种账根本对不回来。
        """
        pairs = self._grouped_pairs() if self.grouped else []
        if pairs:
            plan = []
            for index, (name, members) in enumerate(pairs, start=1):
                # 序号前缀：两组撞名（或者组名被消毒成同一个词）时至少是两个文件，
                # 而列表里的先后次序本来就是出文件的次序。
                plan.append((name, self.output_dir / ("%d-%s.docx" % (index, _safe_name(name))),
                             members))
            return plan
        if self.merged:
            return [("合并输出", self.target_for(paths), list(paths))]
        return [(path.stem, self.output_dir / (path.stem + ".docx"), [path]) for path in paths]

    def target_for(self, paths: Sequence[Path]) -> Path:
        """这批文件打算写成哪一个 .docx（多份时给第一份的名字）。合并输出只有一份，就叫 ``合并输出.docx``。"""
        if self.grouped:
            pairs = self._grouped_pairs()
            if pairs:
                return self.output_dir / ("1-%s.docx" % _safe_name(pairs[0][0]))
        if self.merged:
            return self.output_dir / "合并输出.docx"
        first = paths[0] if paths else Path("未命名")
        return self.output_dir / (first.stem + ".docx")

    def _grouped_pairs(self) -> list:
        """组名 + 只留盘上真在的、勾上的文件；一个文件都不剩的组不摆（空组由主窗口念）。"""
        pairs = []
        for entry in self.groups or ():
            name, members = (entry[0], entry[1]) if len(entry) == 2 else (str(entry), [])
            kept = [Path(path) for path in members]
            if kept:
                pairs.append((str(name), kept))
        return pairs


def _safe_name(name: str) -> str:
    """组名当文件名用之前把不让用的字符换掉；空着（或全是坏字符）给一个兜底词而不是抛。"""
    cleaned = "".join("_" if char in _BAD_NAME_CHARS else char for char in str(name).strip())
    # 尾部的空格和点资源管理器会悄悄去掉：两份输出就可能撞成一份，所以这里先去掉。
    cleaned = cleaned.rstrip(". ")
    # 整名只剩坏字符时换出来是一串下划线，和"没起名"是同一件事——串下划线念不出哪一组。
    return cleaned if cleaned.strip("_. ") else "未命名"


def _size_text(size: float) -> str:
    """12.0 念成 12，11.5 照念——字号后面跟着 "g"（磅的中文写法），不该出现 "12.0"。"""
    return str(int(size)) if float(size).is_integer() else ("%g" % size)


def _bytes_text(size: int) -> str:
    """文件多大给人读的那一档（报告表里那一格是给人扫的，不是给程序比的）。"""
    if size < 1024:
        return "%d B" % size
    if size < 1024 * 1024:
        return "%.1f KB" % (size / 1024.0)
    return "%.1f MB" % (size / (1024.0 * 1024.0))


# ---------------------------------------------------------------- 真转换

#: 逐份输出那一张表的列名。报告对话框直接画它，界面不自己拼、也不数第几列是什么。
COLUMNS_OF_OUTPUT = ("这一份", "写到哪", "收了几个文件", "公式（条数）", "表格", "图片", "多大", "读回核对")
COLUMNS_OF_PROBLEM = ("哪一步", "说了什么")


@dataclass
class _Ledger:
    """一次跑批攒下的账：几项汇总数 + 逐份清单 + 问题 + 说明，最后 :meth:`payload` 换成一张 dict。

    为什么不直接让界面去数日志行：日志是给人扫一眼的、报告是给人核对的，两处各数一遍
    就会出现"日志说 3 个文件、报告说 2 个"。转换与校验共用这一个账本，
    所以报告对话框那一张表两路同形。
    """

    action: str
    options: Options
    total: int
    outputs: list[list[str]] = field(default_factory=list)
    problems: list[list[str]] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    files_ok: int = 0
    files_failed: int = 0
    math_ok: int = 0
    math_failed: int = 0
    tables: int = 0
    blocks: int = 0
    images: int = 0
    images_embedded: int = 0
    images_skipped: int = 0
    started: float = field(default_factory=time.monotonic)
    stopped_at: Optional[int] = None

    @property
    def cancelled(self) -> bool:
        return self.stopped_at is not None

    @property
    def walked(self) -> int:
        return self.files_ok + self.files_failed

    @property
    def math_total(self) -> int:
        return self.math_ok + self.math_failed

    def stop(self, done: int) -> None:
        self.stopped_at = done

    def payload(self, receipt: str) -> dict:
        """交回报告对话框与状态栏共用的那一份统计（**纯数据**：全是文本与数，没有控件）。"""
        elapsed = time.monotonic() - self.started
        where = ("已取消，停在第 %d 个（后面 %d 个没动）" % (
            self.stopped_at, self.total - self.stopped_at) if self.cancelled
            else "全部走完")
        head = [
            ("成功几个文件", "%d 个" % self.files_ok),
            ("失败几个文件", "%d 个" % self.files_failed),
            ("公式数", "%d 条（真转进 Word %d 条、转不出 %d 条）" % (
                self.math_total, self.math_ok, self.math_failed)),
            ("表格数", "%d 张" % self.tables),
            ("块数", "%d 个" % self.blocks),
        ]
        if self.action == ACTION_CONVERT:
            head.append(("图片", "%d 处（嵌进 %d，没成 %d）" % (
                self.images, self.images_embedded, self.images_skipped)))
            head.append(("写出几份", "%d 份" % len(self.outputs)))
            head.append(("输出目录", str(self.options.output_dir)))
        else:
            head.append(("盘上动了什么", "什么都没动（校验只读，公式数也还只是解析阶段数出来的）"))
        head += [
            ("这一批共几个文件", "%d 个" % self.total),
            ("走到哪一步", where),
            ("输出模式", self.options.mode_name()),
            ("耗时", "%.2f 秒" % elapsed),
        ]
        return {
            "action": self.action,
            "cancelled": self.cancelled,
            "receipt": receipt,
            "title": "转换报告" if self.action == ACTION_CONVERT else "校验报告",
            "head": head,
            "columns": list(COLUMNS_OF_OUTPUT),
            "rows": self.outputs,
            "problem_columns": list(COLUMNS_OF_PROBLEM),
            "problems": self.problems,
            "notes": self.notes,
            "total": self.total,
            "done": self.walked,
            "ok_count": self.files_ok,
            "failed_count": self.files_failed,
            "math_total": self.math_total,
            "math_ok": self.math_ok,
            "math_failed": self.math_failed,
            "tables": self.tables,
            "blocks": self.blocks,
            "outputs": len(self.outputs),
            "elapsed": round(elapsed, 3),
            "output_dir": str(self.options.output_dir),
            "mode": self.options.mode_name(),
            "summary": self.options.summary(),
        }


def _ticker(progress: Progress, steps: int) -> Callable[[str], None]:
    """把"走了几步"换成百分比：进度条要的是 0-100，而总步数只有执行体自己知道。

    步数 = 文件数 + 份数（每份末尾还有"写出去"那一步），所以合并输出时"5 个文件 1 份 docx"
    也走得满 100，而不是读完第 5 个就卡住不动。
    """
    state = {"done": 0}

    def tick(message: str) -> None:
        state["done"] += 1
        percent = int(round(state["done"] * 100.0 / max(steps, 1)))
        progress(min(percent, 100), message)

    return tick


def _issue(log: Log, logged: Optional[Logged], line: str, kind: str) -> None:
    """念一句"出事了"的话：有档位那一根就带上档位，没有就当普通日志念。

    档位只决定日志区那一行涂什么颜色，**不改文本**——同一句话在日志里、在报告表里、
    在测试断言里必须是逐字一样的，多一个颜色不该多一个字。
    """
    if logged is None:
        log(line)
    else:
        logged(line, kind)


def _engine(options: Options, log: Log, output_count: int) -> db.DocxBuilder:
    """定位样式表 + 编译一次 + 装配一个 builder：**一批共用这一套**。

    ``locate_xslt`` 抛的 ``XsltNotFoundError`` 原样往上传——它自己的句子就写了"查过哪几个位置
    + 三种解决办法"（命令行那边同一个先例），这里再拼一份就会出现"两处说法不一样的下一步"。
    找不到样式表要**在读第一个文件之前**拒：写了一半才发现是最难看的失败。
    """
    xslt = locate_xslt()
    builder = db.DocxBuilder(
        FormulaConverter(xslt),
        db.BuilderOptions(body_font=options.body_font, body_size=options.body_size,
                          heading_font=options.heading_font))
    log("样式表 %s｜这一批 %d 份输出共用这一次编译" % (xslt, output_count))
    return builder


def _read_one(path: Path, ledger: _Ledger, log: Log,
              logged: Optional[Logged] = None) -> Optional[mp.ParseResult]:
    """读一个文件：读不进只记一条问题，**不拦后面那些**。

    这里数出来的公式条数只写进日志那一行；统计里的公式数等真写出文件来再加（``_write_one``）。
    """
    if not path.is_file():
        ledger.files_failed += 1
        ledger.problems.append([path.name, "已经不在这个位置了（先把它从左边列表里移除）"])
        _issue(log, logged, "跳过：%s 已经不在了。" % path.name, "bad")
        return None
    try:
        parsed = mp.parse_file(path)
    except (OSError, UnicodeDecodeError, ValueError, mp.MarkdownSyntaxError) as exc:
        ledger.files_failed += 1
        detail = str(exc) or type(exc).__name__
        ledger.problems.append([path.name, "读不进或解析不进：%s" % detail])
        _issue(log, logged, "失败：%s 读不进或解析不进（%s）" % (path.name, detail), "bad")
        return None
    ledger.files_ok += 1
    stats = parsed.stats
    log("读完：%s｜块 %d｜公式 %d（行内 %d / 独立 %d）｜表格 %d｜按 %s 读" % (
        path.name, stats.get("blocks", 0),
        stats.get("formulas_inline", 0) + stats.get("formulas_display", 0),
        stats.get("formulas_inline", 0), stats.get("formulas_display", 0),
        stats.get("tables", 0), stats.get("encoding", "utf-8")))
    for warning in parsed.warnings:
        ledger.problems.append([path.name, "解析提醒：%s" % warning])
    return parsed


def _verify(target: Path, math_ok: int) -> tuple[str, Optional[str]]:
    """读回落盘的文件数 ``m:oMath``——这一格才是"公式真的进了 Word 而不是丢了"的证据。

    内存里的树说得再好听，落盘那一步被丢掉了才是事故（命令行 ``_verification_line`` 同一个先例）。
    """
    try:
        objects, display = db.count_omml(target)
    except (OSError, KeyError) as exc:
        return "读不回（文件已写出，公式条数没核出来）", (
            "%s 写出去了，但读回来数不出 m:oMath（%s）——这一份先别交。" % (target.name, exc))
    if objects == math_ok:
        return "%d 个 m:oMath，与转成的一致" % objects, None
    return ("不一致：内存 %d 条 / 文件 %d 个" % (math_ok, objects),
            "%s 里数到 %d 个 m:oMath，与内存里转成的 %d 条不一致——请把这份 .md 反馈给作者，"
            "不要直接交这一份 .docx。" % (target.name, objects, math_ok))


def _write_one(label: str, target: Path, entries: Sequence[mg.MergeEntry],
               builder: db.DocxBuilder, ledger: _Ledger, log: Log,
               logged: Optional[Logged] = None) -> None:
    """一份输出走到底：并树 → 生成 → 落盘 → 读回核对，并把这一行记进逐份清单。

    并树这件事全在 :func:`gui.merger.merge_entries` 里（谁在前谁在后、标题降不降、
    哪一块前面换页、封面写什么），这一层只把面板那四个勾翻成声明交过去。
    ``label`` 一起当封面标题用：合并那一份叫"合并输出"，分组那份叫组名——
    封面上写的是"这一份叫什么"，不是第一个文件的名字。
    """
    plan = mg.merge_entries(entries, ledger.options.merge_switches(), title=label)
    try:
        report = builder.build(plan.parsed, plan.layout)
    except Exception as exc:  # 引擎内部炸了：只这一份没写成，后面的份照走
        ledger.problems.append([label, "这一份生成时出错：%s（%s）" % (exc, type(exc).__name__)])
        _issue(log, logged, "失败：「%s」生成时出错（%s），%s 没写。"
               % (label, type(exc).__name__, target.name), "bad")
        return
    try:
        report.save(target)
    except PermissionError:
        ledger.problems.append([label, "写不出 %s：没有权限，或者这个文件正被 Word 打开着"
                                    "（把 Word 关掉再写一次）" % target.name])
        _issue(log, logged, "失败：写不出 %s——多半是 Word 还开着这一份" % target.name, "bad")
        return
    except OSError as exc:
        ledger.problems.append([label, "写不出 %s：%s" % (target.name, exc)])
        _issue(log, logged, "失败：写不出 %s（%s）" % (target.name, exc), "bad")
        return
    ledger.math_ok += report.math_ok
    ledger.math_failed += report.math_failed
    ledger.tables += report.tables
    ledger.blocks += report.blocks
    ledger.images += report.images
    ledger.images_embedded += report.images_embedded
    ledger.images_skipped += report.images_skipped
    verify, note = _verify(target, report.math_ok)
    if note:
        ledger.notes.append(note)
    size = target.stat().st_size if target.is_file() else 0
    pictures = "%d 处（嵌进 %d，没成 %d）" % (
        report.images, report.images_embedded, report.images_skipped)
    ledger.outputs.append([
        label, str(target), "%d" % len(entries),
        "%d（转不出 %d）" % (report.math_total, report.math_failed),
        "%d" % report.tables, pictures, _bytes_text(size), verify])
    log("写出：%s｜%d 个文件 → %s｜%s｜公式 %d（转不出 %d）｜表格 %d｜图片 %s｜%s" % (
        label, len(entries), target, _bytes_text(size),
        report.math_total, report.math_failed, report.tables, pictures, verify))
    for item in report.failures[:MAX_FAILURE_ROWS]:
        ledger.problems.append([target.name, "有条公式转不出来（Word 里是原文 + 黄底，可直接改）："
                                             "%s" % item.failure_report()])
    if len(report.failures) > MAX_FAILURE_ROWS:
        ledger.notes.append("「%s」里另有 %d 条转不出的公式没逐条摆出来（条数已计进统计）。" % (
            target.name, len(report.failures) - MAX_FAILURE_ROWS))
    for text in plan.notes():
        ledger.problems.append([target.name, "合并时说：%s" % text])
    for text in report.notes:
        ledger.problems.append([target.name, "生成时说：%s" % text])


def convert_paths(paths: Sequence[Path], options: Options, log: Log, progress: Progress,
                  interrupted: Callable[[], bool] | None = None,
                  logged: Optional[Logged] = None) -> dict:
    """真转换：照 :meth:`Options.output_plan` 逐份读稿 → 转公式 → 写 .docx → 读回核对。

    返回一次跑批的统计（见 :class:`_Ledger`）。**出几份、叫什么名、哪几个文件进哪一份**
    仍然只问 ``output_plan`` 那一份计划——界面念的就是这里写的，两处共用同一个计划才核得回来。
    """
    total = len(paths)
    ledger = _Ledger(ACTION_CONVERT, options, total)
    if total == 0:
        log("没有勾选的文件，什么都没做。")
        return ledger.payload("没有勾选的文件，什么都没写。")
    plan = options.output_plan(paths)
    tick = _ticker(progress, total + len(plan))
    log("开始转换：%d 个文件 → %d 份输出｜%s" % (total, len(plan), options.summary()))
    if len(plan) > 1:
        log("这批要写 %d 份：%s" % (
            len(plan), "；".join("%s → %s" % (label, target.name) for label, target, _ in plan)))
    else:
        log("这批只出 1 份 → %s" % plan[0][1])
    builder = _engine(options, log, len(plan))

    done = 0
    for label, target, members in plan:
        entries: list[mg.MergeEntry] = []
        for path in members:
            if interrupted is not None and interrupted():
                ledger.stop(done)
                log("已取消：%d 个文件已经走完，剩下 %d 个不再动；已经写出去的文件不回滚。" % (
                    done, total - done))
                break
            parsed = _read_one(path, ledger, log, logged)
            if parsed is not None:
                # 路径跟着解析结果一起攒：来源标注那一行要念文件名，而块树里没有文件名。
                entries.append((path, parsed))
            done += 1
            tick("第 %d/%d 个：%s 读完" % (done, total, path.name))
            time.sleep(PAUSE_PER_FILE_SECONDS)
        if ledger.cancelled:
            break
        if not entries:
            # 一份输出里一个可读进来的文件都没有：不摆空表，也不写一份只有空段落的 docx。
            _issue(log, logged, "「%s」这一份没有可读进来的文件，不写 %s。" % (label, target.name),
                   "bad")
            ledger.problems.append([label, "这一份一个文件都没读进来（上面列了各自的原因），所以没写 %s。"
                                    % target.name])
            tick("「%s」没有可写的文件" % label)
            continue
        _write_one(label, target, entries, builder, ledger, log, logged)
        tick("第 %d/%d 个之后：%s 已写出" % (done, total, target.name))

    if ledger.cancelled:
        # 「公式 N 条」在两档回执里必须是同一个意思（解析到几条），所以这里也用 math_total：
        # 拿 math_ok 来拼就会和报告概要那一格差掉"转不出"的那几条，而两句用的是同一个词。
        receipt = ("已取消：%d 个文件走完、写出 %d 份｜公式 %d 条（转不出 %d）｜"
                   "图片 %d 处（嵌进 %d，没成 %d）｜"
                   "已经写出的文件不回滚，剩下的下次接着写。" % (
                       ledger.walked, len(ledger.outputs), ledger.math_total, ledger.math_failed,
                       ledger.images, ledger.images_embedded, ledger.images_skipped))
    else:
        # 图片那三个数进回执与命令行那一份是同一句话（docx_builder.py 的 BuildReport.lines()）：
        # 只念"表格 3 张"会把"有一张没找到"演成没事，而「问题与提醒」里已经写了两句。
        receipt = ("转换完成：%d 个文件成功、%d 个失败｜写出 %d 份｜公式 %d 条（转不出 %d）｜"
                   "表格 %d 张｜图片 %d 处（嵌进 %d，没成 %d）" % (
                       ledger.files_ok, ledger.files_failed, len(ledger.outputs),
                       ledger.math_total, ledger.math_failed, ledger.tables,
                       ledger.images, ledger.images_embedded, ledger.images_skipped))
    log(receipt)
    return ledger.payload(receipt)


def validate_paths(paths: Sequence[Path], options: Options, log: Log, progress: Progress,
                   interrupted: Callable[[], bool] | None = None,
                   logged: Optional[Logged] = None) -> dict:
    """真做检查但**不写盘**：每个文件试解析一次，念出块数 / 公式数 / 表格数，
    再走 :mod:`gui.validator` 那四项（公式语法 / 图片路径 / 表格列数 / 标题层级），
    最后确认样式表找不找得到。统计与转换同形（报告那一张表两路共用）。"""
    total = len(paths)
    ledger = _Ledger(ACTION_VALIDATE, options, total)
    if total == 0:
        log("没有勾选的文件，没什么可校验。")
        return ledger.payload("没有勾选的文件，什么都没查。")
    tick = _ticker(progress, total)
    log("开始校验：%d 个文件（只读，不动盘）" % total)
    ledger.notes.append("校验这一路只读：不写任何 .docx。下面的公式条数是解析阶段数出来的，"
                        "「真进了 Word 几条」要等转换写出来才核得出。")
    xslt_line = _check_xslt()
    done = 0
    for path in paths:
        if interrupted is not None and interrupted():
            ledger.stop(done)
            log("已取消校验：剩下 %d 个文件没看。" % (total - done))
            break
        _validate_one(path, ledger, log, logged)
        done += 1
        tick("第 %d/%d 个：%s" % (done, total, path.name))
        time.sleep(PAUSE_PER_FILE_SECONDS)
    log(xslt_line)
    if ledger.cancelled:
        receipt = "已取消校验：看了 %d 个文件，列出 %d 处问题。" % (done, len(ledger.problems))
    else:
        receipt = "校验完成：%d 个文件，列出 %d 处问题（什么都没写）" % (done, len(ledger.problems))
    log(receipt + "——真的核对 .docx 里的 m:oMath 条数要用「开始转换」。")
    return ledger.payload(receipt)


def _validate_one(path: Path, ledger: _Ledger, log: Log,
                  logged: Optional[Logged] = None) -> None:
    if not path.is_file():
        ledger.files_failed += 1
        ledger.problems.append([path.name, "已经不在这个位置了（先把它从左边列表里移除）"])
        _issue(log, logged, "问题：%s 已经不在了（先把它从列表里移除）。" % path, "bad")
        return
    try:
        parsed = mp.parse_file(path)
    except (OSError, UnicodeDecodeError, ValueError, mp.MarkdownSyntaxError) as exc:
        # 读不动 / 解析不动都算一处问题：校验的存在理由就是把这些当场念出来。
        ledger.files_failed += 1
        detail = str(exc) or type(exc).__name__
        ledger.problems.append([path.name, "读不动或解析不动：%s" % detail])
        _issue(log, logged, "问题：%s 读不动或解析不动：%s" % (path.name, detail), "bad")
        return
    stats = dict(parsed.stats)
    formulas = stats.get("formulas_inline", 0) + stats.get("formulas_display", 0)
    ledger.files_ok += 1
    ledger.tables += stats.get("tables", 0)
    ledger.blocks += stats.get("blocks", 0)
    log("%s｜块 %d｜标题 %d｜表格 %d｜公式 %d（行内 %d / 独立 %d）｜代码块 %d" % (
        path.name, stats.get("blocks", 0), stats.get("headings", 0), stats.get("tables", 0),
        formulas, stats.get("formulas_inline", 0), stats.get("formulas_display", 0),
        stats.get("code_blocks", 0)))
    ledger.outputs.append([path.name, "不写文件（只读）", "1",
                           "解析到 %d 条（还没转）" % formulas, "%d" % stats.get("tables", 0),
                           "—", "校验不写文件，这一行只念解析数"])
    found = validator.validate_parsed(parsed, base=path.parent)
    for item in found:
        ledger.problems.append([path.name, item.text])
        _issue(log, logged, item.text, item.token)
    # 一句"这一份查完了、查出几处"：一项都没中也要念，否则干净的稿子在日志里等于没查过。
    _issue(log, logged, "小结：%s｜%s" % (path.name, validator.summary_text(found)),
           validator.worst_token(found))
    already = {issue.message for issue in parsed.issues}
    for warning in parsed.warnings:
        if warning in already:
            # 列数那一类已由上面那四项念过（句子与解析器那句逐字一样），同一件事不说两遍。
            continue
        ledger.problems.append([path.name, "解析提醒：%s" % warning])


def _check_xslt() -> str:
    """样式表是命令行那条链的定位口径，这里直接借它，不另写一套找文件的逻辑。"""
    try:
        found = locate_xslt()
    except Exception as exc:  # XsltNotFoundError 的句子已经写了"解决办法"，原样念就够
        return "样式表找不到：%s" % str(exc)
    return "样式表找到了：%s" % found


# ---------------------------------------------------------------- 线程壳


class ConvertWorker(QThread):
    """那一根后台线程：跑批在 ``run()`` 里，界面只认识信号。

    五个信号就是主窗口拿到的全部（工作线程里碰不到任何控件，也不回头改那份路径列表）：

    * ``progress(int, str)``——百分比（0-100）加"第 N/M 个：某某.md"，喂进度条与状态栏。
    * ``log(str)``——日志区实时那一行（需求那七条没点它的名，但第 6 条"实时输出日志"走的就是这一根；
      为什么不让主窗口自己写日志：写日志的时机在执行体里，主窗口只知道开始和结束）。
    * ``logged(str, str)``——同一句话**带上一档颜色**（``bad`` / ``warn`` / ``ok``），
      六期第 3 条"错误项高亮"走这一根。为什么是**多一根**而不是把 ``log`` 改成带档位：
      ``log`` 那个形状是四期定下来的、主窗口和一批测试都连着它，改了形状就等于把已验收的
      契约再改一遍；执行体那一路没传这一根时照旧只念 ``log``，文本逐字一样。
    * ``finished(dict)``——这一批走完了（或取消了），交回统计那张 dict。
    * ``error(str)``——**这一批根本没跑成**：样式表找不到，或者引擎里出了没预料到的异常。
      它和 ``finished`` 不是一件事：一批里有个文件读不进照样走 ``finished``，问题列在报告里；
      所以主窗口必须在**这两个**信号上都把忙态收掉，否则出一次错按钮就永远灰着。

    一件必须知道的取舍：**``finished`` 这个名字盖掉了 ``QThread`` 自带的那个 ``finished``**
    （需求点名要这一档参数形状，实测 PySide6 里 ``Signal`` 按实例属性解析，子类声明赢，
    C++ 内置那个从 Python 侧再也拿不到）。后果是 ``worker.finished.connect(worker.deleteLater)``
    这一行到处都在抄的写法在这里**不成立**——它会把统计 dict 交给 ``deleteLater``。
    所以线程壳的销毁要显式收尾：先 ``wait()`` 确认真的退出再 ``deleteLater()``
    （落在 :meth:`gui.main_window.MainWindow._retire_worker`）。

    取消不另开 API：``requestInterruption()`` 是 ``QThread`` 自带的，执行体把
    :meth:`isInterruptionRequested` 当"该停了吗"那个回调，所以点「取消这一批」之后
    手上那一个文件走完才停，已经写出去的文件不回滚。
    """

    progress = Signal(int, str)
    log = Signal(str)
    logged = Signal(str, str)
    finished = Signal(dict)
    error = Signal(str)

    def __init__(self, action: str, paths: Sequence[Path], options: Options,
                 parent=None) -> None:
        super().__init__(parent)
        self._action = action
        self._paths = list(paths)
        self._options = options

    def run(self) -> None:  # noqa: D102 - QThread 的入口
        runner = convert_paths if self._action == ACTION_CONVERT else validate_paths
        try:
            stats = runner(self._paths, self._options, self.log.emit, self.progress.emit,
                           interrupted=self.isInterruptionRequested, logged=self.logged.emit)
        except XsltNotFoundError as exc:
            # 不是"某个文件坏了"而是"这一批没法开始"，而 exc 自己就写了查过哪儿、怎么办。
            self.error.emit(str(exc))
            return
        except Exception as exc:  # 后台炸了不能把界面带崩，也不能不出声
            self.error.emit("这一步中断：%s（%s）" % (exc, type(exc).__name__))
            return
        self.finished.emit(stats)
