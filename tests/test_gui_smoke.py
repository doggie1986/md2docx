# -*- coding: utf-8 -*-
"""gui 那一层的开窗冒烟。

跑法::

    .venv\\Scripts\\python.exe -m unittest discover -s tests -v

**这里不核"好不好看"**——那一条只有真开一次窗口才看得出来。这个文件钉的全是
"接错了界面也不炸、只是悄悄不对"的那几类：工具栏九颗动作的 objectName 与次序、
复选框的勾选语义、右键菜单的条目（单行 / 多选 / 分组线 / 空白四种形状各一版）、
单独转换强制走单独输出、跑批期间的忙态、**四期那条真落盘的那一路**（写出的是 .docx、
进度条走到 100、报告那一张表念的是执行体算的数、取消之后已写的留着没写的不补、
找不到样式表走 error 那一路且忙态照样收）、**报告窗那两句只有数画布才量得到的画法**
（长句子换行而不是裁在窗口边上、取消那一路的墨色真的是警告色）、主题回落、QSettings 往返，
再加三期那两条：**分组归属与"一组出一份 .docx"的命名**、**列表存成 .json 再演回来**
（存读那一路的每一句拒绝、星号、关窗写回、启动恢复），
以及六期那几件**只有真把文件灌进中间那栏才看得见**的：公式摆成了什么形状（块级分式上下两格、
坏公式退回原文而其余部分照旧）、嵌套列表的第三层在不在屏上、校验查出的每一行落在哪一档墨色
（日志区按档数 viewport 上的像素，"档位字符串对"不等于"字被涂成了那个颜色"），
外加切主题时日志重画与预览样式表重贴，
再加八期那几件**只有真把面板接上本机设置才看得见**的：改一格之后那一拍落盘（实时那句不是空话）、
自定义预设的存 / 挑 / 撤与下拉里"演哪一档"是怎么推出来的、被拒的那一句画在面板自己那一行而不弹框、
本机坏掉的预设**不被一次普通落盘顶掉**、一份配置抄本从一扇窗换到另一扇窗（带过去的与没带过去的），
以及「最近打开的工程」那一列的记 / 剪 / 点 / 清。

`OutputPlanTest` 与 `ProjectFileTest` 不开窗：前者只碰 ``Options``（一个纯 dataclass），
后者只碰 ``gui.project_file``（纯标准库）。它们不需要 QApplication，所以起不来 Qt 的机器
照样跑得到——把不需要窗口的用例塞进 ``skipIf`` 里，等于把一半覆盖跟着窗口一起丢掉。

无显示环境（CI）用 offscreen 平台跑；连 Qt 都起不来就整模块 skip——
"跳过"不等于"通过"，看结果的人要能分辨这两件事。
"""

from __future__ import annotations

import gc
import json
import os
import shutil
import sys
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from typing import Sequence

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if sys.platform.startswith("linux") and not os.environ.get("DISPLAY"):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

APP = None
QT_START_ERROR = ""
try:
    from PySide6.QtCore import (QItemSelectionModel, QPoint, QMimeData, QSettings, QThread, Qt,
                                QUrl)
    from PySide6.QtGui import QImage
    from PySide6.QtWidgets import QApplication, QProgressBar, QToolBar

    APP = QApplication.instance() or QApplication(sys.argv[:1])
    APP.setStyle("Fusion")
except Exception as exc:  # pragma: no cover - 起不来 Qt 的机器走到这里
    QT_START_ERROR = "%s：%s" % (type(exc).__name__, exc)

import gui.tasks as tasks  # noqa: E402
import gui.theme as theme  # noqa: E402
from config.settings import CARRY_NOTE, SETTINGS_KEYS, TOP_KEYS  # noqa: E402
from gui import project_file as project_file  # noqa: E402
from gui.file_panel import MARKDOWN_SUFFIXES, UNGROUPED  # noqa: E402
from gui.main_window import (MENU_ACTION_OBJECT_NAMES, TOOLBAR_OBJECT_NAMES, MainWindow,  # noqa: E402
                             file_filter)
from gui.preferences import (KEY_BODY_SIZE, KEY_OUTPUT_DIR, KEY_PROJECT, KEY_RECENT_PROJECTS,  # noqa: E402
                             KEY_SPLITTER, KEY_STYLE_PRESETS, KEY_THEME, MODE_GROUPED,
                             MODE_MERGED, MODE_SEPARATE, Store)
from gui.preview_panel import LOG_KIND_TOKENS  # noqa: E402
from gui.settings_panel import SAVE_DEBOUNCE_MS  # noqa: E402
from gui.tasks import ConvertWorker, Options  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
DEMO = ROOT / "samples" / "demo.md"

#: 单选一行时右键那一张菜单。``文件状态：可用`` 是三期第 5 条那颗状态图标的说法，
#: 它进了菜单就必须跟着菜单一起钉——否则"图标坏了而菜单没跟着坏"这种状态会被读成界面没错。
ROW_MENU_TEXTS = ["单独转换这一条", "这一条不参与转换", "从列表移除",
                  "文件状态：可用", "打开所在目录"]
BLANK_MENU_TEXTS = ["加入 Markdown 文件…", "加入整个文件夹…", "全部参与转换", "全部不参与",
                    "新建分组…", "重新核对文件状态", "清空列表"]
MERGE_CHECK_TEXTS = ["生成目录", "插入分页符", "标注源文件名", "生成封面页"]


def pump(seconds: float = 0.1) -> None:
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        APP.processEvents()
        time.sleep(0.005)


def wait_until(condition, timeout: float = 15.0) -> bool:
    """自己排事件循环等到条件成立。线程的日志与进度信号是排队来的，不 pump 就永远不到。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        APP.processEvents()
        if condition():
            return True
        time.sleep(0.01)
    APP.processEvents()
    return bool(condition())


class UrlDropEvent:
    """只带 ``mimeData()`` / ``acceptProposedAction()`` 的最小事件。

    真造一个 ``QDropEvent`` 要的是屏幕坐标与全局鼠标状态，脚本里没有鼠标；主窗口用的只有
    那两个方法，所以给这两个就够把"拖进来的是本机文件"那一条分支走一遍。
    """

    def __init__(self, paths) -> None:
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(str(path)) for path in paths])
        self._mime = mime
        self.accepted = False

    def mimeData(self) -> QMimeData:
        return self._mime

    def acceptProposedAction(self) -> None:
        self.accepted = True


#: 「说了什么」那一格真会写到的长度（latex2mathml 那句拒绝，实测 837 像素宽）。
LONG_PROBLEM = ("有条公式转不出来（Word 里是原文 + 黄底，可直接改）：\\foo \\bar  ←  "
                "latex2mathml 没能读懂这段 LaTeX（NoAvailableTokensError: 无信息）")


def report_sample(cancelled: bool = False, problem: str = LONG_PROBLEM) -> dict:
    """一张"够报告画满"的 dict。

    下面那两条核的是这一窗**怎么画**，"数从哪来"由跑批那几条用例钉着，所以这里手写一份统计
    而不真跑一批——真跑的话那句坏话长什么样由引擎决定，而这一条要的正是"一句已知很长的话"。
    """
    return {
        "title": "转换报告",
        "cancelled": cancelled,
        "receipt": ("已取消：1 个文件走完、写出 1 份｜公式 1 条" if cancelled
                    else "转换完成：1 个文件成功、0 个失败｜公式 1 条"),
        "head": [["走到哪一步", "已取消，停在第 1 个" if cancelled else "全部走完"]],
        "columns": ["这一份", "写到哪"], "rows": [["demo", r"C:\x\demo.docx"]],
        "problem_columns": ["哪一步", "说了什么"], "problems": [["demo.docx", problem]],
        "notes": [],
    }


def ink_rgb(widget) -> tuple:
    """数出这一格里"最不像底色的那个颜色"。

    底色取出现最多的那一档（不是左上角那一格——文字贴着边起笔时左上角本身就是墨）。
    样式表到底有没有落笔只有看画布量得到：换 ``objectName`` 之后没重算 polish 的话，
    名字是 ``problem``、``text()`` 也对、palette 也不报错，画出来却还是正文色。
    """
    image = widget.grab().toImage().convertToFormat(QImage.Format.Format_RGB32)
    counts = {}
    for y in range(0, image.height(), 2):
        for x in range(0, image.width(), 2):
            pixel = image.pixel(x, y)
            counts[pixel] = counts.get(pixel, 0) + 1
    if len(counts) < 2:
        raise AssertionError("这一格里只有一种颜色： grab 出来是空的，量不到墨色")
    background = max(counts, key=counts.get)
    ink = max((pixel for pixel in counts if pixel != background),
              key=lambda pixel: counts[pixel])
    return ((ink >> 16) & 255, (ink >> 8) & 255, ink & 255)


def near(left: Sequence[int], right: Sequence[int], tolerance: int = 24) -> bool:
    return all(abs(a - b) <= tolerance for a, b in zip(left, right))


#: 认墨色的容差。留 12 是因为"坏"与"提醒"两档只差 43（浅色里 (207,34,46) 与 (164,38,44)），
#: 而抗锯齿把字往底色混出去的那一档差 60 以上——两头都够得着，也不会把邻档认成本档。
INK_TOLERANCE = 12


def ink_pixels(view, colour) -> int:
    """数 viewport 上"正是这一档墨色"的像素有几个（0 = 这一档根本没落笔）。

    取 ``viewport()`` 而不是整块控件：``QPlainTextEdit`` 整块 grab 里含边框与滚动条，实测铬色
    像素 1643 个、比整行正文的墨（176 个）还多，:func:`ink_rgb` 那套"最多的当背景、剩下最多的当墨"
    在日志区量到的就是边框而不是字。日志区要问的是**这一行**到底涂了哪一档，所以按档数像素。
    """
    image = view.viewport().grab().toImage().convertToFormat(QImage.Format.Format_RGB32)
    target = (colour.red(), colour.green(), colour.blue())
    hits = 0
    for y in range(0, image.height(), 2):
        for x in range(0, image.width(), 2):
            pixel = image.pixel(x, y)
            rgb = ((pixel >> 16) & 255, (pixel >> 8) & 255, pixel & 255)
            if all(abs(a - b) <= INK_TOLERANCE for a, b in zip(rgb, target)):
                hits += 1
    return hits


@unittest.skipIf(APP is None, "起不来 Qt（%s）" % (QT_START_ERROR or "未知原因"))
class WindowTest(unittest.TestCase):
    """一测一窗，每扇窗换一个临时 QSettings 档。

    换档是因为 ``QSettings`` 默认写注册表——测试不许把人机器上记的窗口尺寸和主题改掉。
    """

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="md2docx-gui-test-"))
        self.ini = self.tmp / "settings.ini"
        self.out_dir = self.tmp / "out"
        self.second = self.tmp / "second.md"
        self.second.write_text("# 第二份\n\n行内 $a+b$。\n", encoding="utf-8")
        self.settings = QSettings(str(self.ini), QSettings.IniFormat)
        self.win = MainWindow(store=Store(settings=self.settings))
        self.win.show()
        self.win.settings.output_dir.setText(str(self.out_dir))
        pump()
        #: 用例里另开的那些 QSettings（"从盘上再读一遍"用的），清理时统一 sync 掉。
        self.extra_settings = []
        # 临时目录归临时目录：排在 tearDown 之后，且**先 sync 再删**——
        # QSettings 析构时还会把没 sync 的改动回写一遍，而回写会把刚删掉的目录连着建回来，
        # 先删后写就永远删不干净（本机实测：sync 过的那颗析构时不再写，没 sync 的那颗会）。
        self.addCleanup(self._drop_tmp)

    def fresh_settings(self) -> QSettings:
        """另开一颗指向同一个档的 QSettings（读**盘上**那一份用），并登记起来好收尾。"""
        settings = QSettings(str(self.ini), QSettings.IniFormat)
        self.extra_settings.append(settings)
        return settings

    def _drop_tmp(self) -> None:
        # 面板那颗 debounce 计时器（``SAVE_DEBOUNCE_MS`` = 600 ms）多半还排在事件循环里——
        # setUp 里那句 output_dir.setText 就把它排上了。它一拍到就 sync，
        # 而 sync 会把刚删掉的目录连着 settings.ini 建回来（下一拍的落点还是这个档），
        # 所以先把它停下再删；这一窗的用例都在用例体内验过落盘，撤掉的只是收尾那一拍。
        self.win.settings._save_timer.stop()
        for settings in [self.settings] + self.extra_settings:
            settings.sync()    # 先把该写的写完，析构那一刻就没有东西可回写
        gc.collect()           # 测试中途临时造的那些 QSettings 析构时同样会回写，先让它们跑完
        shutil.rmtree(self.tmp, ignore_errors=True)

    def tearDown(self) -> None:
        if self.win.act_cancel.isEnabled():
            self.win.cancel_current()
            wait_until(lambda: self.win.act_convert.isEnabled(), timeout=10)
        self.win.close()
        pump()

    def written_docx(self) -> list:
        return [] if not self.out_dir.exists() else list(self.out_dir.glob("*.docx"))

    def make_md(self, name: str) -> Path:
        path = self.tmp / name
        path.write_text("# %s\n\n行内 $a+b$。\n" % name, encoding="utf-8")
        return path

    def select_rows(self, rows: Sequence[int]) -> None:
        """按行号选中几行。用 ``selectionModel`` 而不是 ``setCurrentRow``：后者每次都是
        ClearAndSelect，多选自多选不出来。"""
        listing = self.win.files.list
        selection = listing.selectionModel()
        for row in rows:
            selection.select(listing.model().index(int(row), 0),
                             QItemSelectionModel.SelectionFlag.Select
                             | QItemSelectionModel.SelectionFlag.Rows)
        pump()

    def menu_texts_at(self, row: int) -> list:
        """摆出右键那张菜单（不 exec）并念它的条目文本。``row=None`` 就是空白处那一张。"""
        listing = self.win.files.list
        if row is None:
            point = QPoint(4, max(listing.height() - 4, 4))
        else:
            point = listing.visualRect(listing.model().index(int(row), 0)).center()
        return [action.text() for action in listing._build_menu(point).actions() if action.text()]

    # ---- 接线 ---------------------------------------------------------------

    def test_the_toolbar_carries_the_nine_named_actions_in_order(self) -> None:
        toolbar = self.win.findChild(QToolBar)
        names = [action.objectName() for action in toolbar.actions() if action.objectName()]
        self.assertEqual(names, list(TOOLBAR_OBJECT_NAMES))
        texts = [action.text() for action in toolbar.actions() if action.text()]
        self.assertEqual(texts, ["添加文件", "添加文件夹", "移除", "上移", "下移",
                                 "清空", "开始转换", "预览", "校验"])

    def test_the_status_bar_holds_text_and_a_progress_bar(self) -> None:
        bars = self.win.findChildren(QProgressBar)
        self.assertTrue(bars)
        self.assertEqual(bars[0].value(), 0)
        self.assertEqual(bars[0].maximum(), 100)
        self.assertIn("就绪", self.win.status_text())

    def test_the_file_filter_comes_from_the_same_suffix_list_as_the_parser(self) -> None:
        for suffix in MARKDOWN_SUFFIXES:
            self.assertIn("*" + suffix, file_filter())

    def test_the_merge_section_is_the_four_checkboxes_named_in_the_request(self) -> None:
        panel = self.win.settings
        self.assertEqual([box.text() for box in (panel.merge_toc, panel.merge_break,
                                                 panel.merge_source, panel.merge_cover)],
                         MERGE_CHECK_TEXTS)
        panel.apply_mode(True)
        self.assertTrue(panel.merge_box.isEnabled())
        self.assertTrue(panel.options().merged)
        panel.merge_cover.setChecked(True)
        self.assertTrue(panel.options().cover_page)
        panel.apply_mode(False)
        self.assertFalse(panel.merge_box.isEnabled())
        self.assertFalse(panel.options().merged)

    # ---- 文件列表 -----------------------------------------------------------

    def test_added_files_are_checked_and_counted(self) -> None:
        self.assertEqual(self.win.files.add_paths([DEMO]), 1)
        self.assertEqual(self.win.files.count(), 1)
        self.assertEqual([path.name for path in self.win.files.checked_paths()], ["demo.md"])
        self.assertIn("共 1 个，勾选 1 个", self.win.files.count_label.text())
        self.win.files.list.set_all_checked(False)
        self.assertIn("共 1 个，勾选 0 个", self.win.files.count_label.text())
        self.assertEqual(self.win.files.checked_paths(), [])

    def test_a_non_markdown_file_is_skipped_out_loud(self) -> None:
        noise = self.tmp / "note.txt"
        noise.write_text("hi", encoding="utf-8")
        self.assertEqual(self.win.files.add_paths([noise]), 0)
        self.assertEqual(self.win.files.count(), 0)
        self.assertIn("不是 Markdown", self.win.log_text())

    def test_the_same_file_is_not_listed_twice(self) -> None:
        self.win.files.add_paths([DEMO])
        self.assertEqual(self.win.files.add_paths([DEMO]), 0)
        self.assertEqual(self.win.files.count(), 1)

    def test_a_folder_adds_the_markdown_files_inside_it(self) -> None:
        self.assertGreaterEqual(self.win.files.add_paths([ROOT / "samples"]), 1)

    def test_dropping_urls_on_the_window_goes_to_the_list(self) -> None:
        event = UrlDropEvent([DEMO])
        self.win.dropEvent(event)
        self.assertTrue(event.accepted)
        self.assertEqual(self.win.files.count(), 1)

    def test_the_first_batch_selects_a_row_so_the_preview_is_not_left_on_the_welcome_page(self) -> None:
        self.win.files.add_paths([DEMO])
        pump()
        self.assertEqual(self.win.files.selected_path().name, "demo.md")
        html = self.win.preview.current_html()
        self.assertIn("demo.md", html)
        # 预览复用命令行那条链的解析器，而公式是**摆出来的**：demo.md 那条 ``\sum`` 的显示式
        # 在这里是上下标 + 一个 Σ 字符，不是带底色的 LaTeX 原文（原文只留给认不出的写法）。
        self.assertIn("Σ", html)
        self.assertIn("class=\"math math-block\"", html)
        self.assertNotIn(r"\sum_", html)
        self.assertIn("预览 demo.md", self.win.status_text())

    def test_the_row_menu_offers_the_named_actions_and_the_blank_area_the_list_actions(self) -> None:
        self.win.files.add_paths([DEMO])
        self.assertEqual(self.menu_texts_at(0), ROW_MENU_TEXTS)
        self.assertEqual(self.menu_texts_at(None), BLANK_MENU_TEXTS)

    def test_the_row_menu_switches_to_the_batch_actions_when_several_rows_are_selected(self) -> None:
        """多选时那张菜单念的是"选中的 N 行"，不是"这一条"——第 4 条要的批量动作就在这一支。"""
        self.win.files.add_paths([DEMO, self.second, self.make_md("third.md")])
        self.select_rows([0, 1, 2])
        self.assertEqual(self.win.files.selected_rows(), [0, 1, 2])
        self.assertEqual(self.menu_texts_at(1),
                         ["单独转换这一条", "选中的 3 行都参与转换", "选中的 3 行都不参与",
                          "移除选中的 3 行", "文件状态：可用", "打开所在目录"])
        self.select_rows([0, 1])
        self.win.files.set_checked_for_rows(self.win.files.selected_rows(), False)
        self.assertEqual(self.win.files.checked_paths(), [])
        self.assertEqual(self.win.files.list.row_kind(2), "file")

    # ---- 分组（需求第 6 条）--------------------------------------------------

    def test_a_group_line_owns_the_files_below_it(self) -> None:
        self.make_md("first.md")
        self.make_md("second.md")
        self.assertTrue(self.win.files.add_group("甲组"))
        self.win.files.add_paths([self.tmp / "first.md", self.tmp / "second.md"])
        self.assertEqual(self.win.files.groups(), ["甲组"])
        self.assertEqual(self.win.files.list.row_kind(0), "group")
        plans = self.win.files.group_plans()
        self.assertEqual([label for label, _ in plans], ["甲组"])
        self.assertEqual([path.name for path in plans[0][1]], ["first.md", "second.md"])
        self.assertEqual(self.win.files.empty_groups(), [])

    def test_files_above_the_first_group_line_are_still_a_group_of_their_own(self) -> None:
        """分组线是"从这一行往下归这一组"，所以线**上面**那几行不会凭空消失——它们归「未分组」。"""
        self.win.files.add_paths([DEMO])
        self.win.files.list.setCurrentRow(0)
        self.assertTrue(self.win.files.add_group("甲组"))
        self.win.files.add_paths([self.second])
        plans = self.win.files.group_plans()
        self.assertEqual([label for label, _ in plans], [UNGROUPED, "甲组"])
        self.assertEqual([path.name for path in plans[0][1]], ["demo.md"])
        self.assertEqual([path.name for path in plans[1][1]], ["second.md"])

    def test_a_group_line_has_no_checkbox_and_batch_checks_skip_it(self) -> None:
        """批量勾/不勾给的是"所有行"，组头必须在里面被跳过：它没有复选框，勾了会把整组当成一条文件。"""
        self.win.files.add_paths([DEMO, self.second])
        self.win.files.list.setCurrentRow(1)
        self.win.files.add_group("甲组")
        rows = list(range(self.win.files.row_count()))
        self.assertEqual(self.win.files.set_checked_for_rows(rows, False), 2)
        self.assertEqual(self.win.files.checked_paths(), [])
        self.assertEqual(self.win.files.groups(), ["甲组"])
        self.assertEqual(self.win.files.list.row_kind(self.win.files.group_rows()[0]), "group")
        self.assertEqual(self.win.files.set_checked_for_rows(rows, True), 2)
        self.assertEqual(len(self.win.files.checked_paths()), 2)

    def test_the_group_line_menu_names_the_group_and_counts_its_members(self) -> None:
        self.win.files.add_group("甲组")
        self.win.files.add_paths([DEMO, self.second])
        self.win.files.list.item(2).setCheckState(Qt.CheckState.Unchecked)
        self.assertEqual(self.menu_texts_at(0),
                         ["改「甲组」的名字…", "「甲组」全部参与（2 个）",
                          "「甲组」全部不参与（2 个）", "去掉这一组的分组线"])

    def test_removing_a_group_line_keeps_its_files(self) -> None:
        self.win.files.add_group("甲组")
        self.win.files.add_paths([DEMO, self.second])
        self.assertEqual(self.win.files.list.remove_group(0), "甲组")
        self.assertEqual(self.win.files.groups(), [])
        self.assertEqual(self.win.files.count(), 2)

    def test_two_groups_cannot_share_a_name(self) -> None:
        """组名就是文件名，两组同名会互相顶掉刚写出的那一份——这一关在建组那一刻就把住。"""
        self.assertTrue(self.win.files.add_group("甲组"))
        self.assertTrue(self.win.files.add_group("乙组"))
        self.assertFalse(self.win.files.add_group("甲组"))
        self.assertFalse(self.win.files.add_group("   "))
        self.assertEqual(self.win.files.groups(), ["甲组", "乙组"])
        self.assertIn("已经有「甲组」这一组", self.win.log_text())
        self.assertIn("组名不能空着", self.win.log_text())

    def test_only_the_group_line_is_painted_with_the_group_background(self) -> None:
        """分组线那条带子必须真的**画在屏幕上**——它是需求第 6 条唯一的视觉凭据。

        两个都会让带子消失、而单元测试读属性读不出来的东西：
        ① 给 QSS 的 ``QListWidget::item`` 加 ``border-radius``，Qt 就改成按样式表自己画背景，
        ``item.setBackground()`` 那一格整个不画（实测同一份 QSS 去掉圆角，那一格立刻演回 group_bg）；
        ② 开隔行底色，它读的是系统调色板的 ``AlternateBase``（这台机器上是绿色 #6cba74，
        QSS 与主题都管不到），落在奇数行的组头就被它盖掉。
        """
        self.win.files.add_paths([DEMO])
        self.win.files.list.setCurrentRow(-1)
        self.assertTrue(self.win.files.add_group("甲组"))
        self.win.files.add_paths([self.second])
        listing = self.win.files.list
        # 选中行有自己的底色（selected），会把行底色顶掉，所以取像素前先把选中清干净。
        listing.clearSelection()
        listing.setCurrentItem(None)
        pump()
        self.assertFalse(listing.alternatingRowColors())
        image = listing.viewport().grab().toImage()
        painted = []
        for row in range(listing.count()):
            rect = listing.visualItemRect(listing.item(row))
            seen = []
            x = rect.left() + 3
            while x < rect.right() - 3:
                seen.append(image.pixelColor(x, rect.center().y()).name())
                x += 7
            # 一行横着采若干点、取众数：按中心那一点取会正好扎在文件名那几个字上。
            painted.append(max(set(seen), key=seen.count))
        self.assertEqual(painted, [theme.color(theme.THEME_LIGHT, token).name()
                                   for token in ("base", "group_bg", "base")])

    def test_open_folder_asks_for_the_row_that_was_clicked(self) -> None:
        self.win.files.add_paths([DEMO])
        self.win.files.list.setCurrentRow(0)
        rect = self.win.files.list.visualRect(self.win.files.list.model().index(0, 0))
        menu = self.win.files.list._build_menu(rect.center())
        seen = []
        original = self.win.files.list.open_folder
        self.win.files.list.open_folder = lambda path: seen.append(path)
        try:
            for action in menu.actions():
                if action.text() == "打开所在目录":
                    action.trigger()
        finally:
            self.win.files.list.open_folder = original
        self.assertEqual([path.name for path in seen], ["demo.md"])

    def test_convert_single_row_forces_separate_output_even_in_merge_mode(self) -> None:
        self.win.settings.apply_mode(True)
        self.win.convert_only(DEMO)
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        log = self.win.log_text()
        self.assertIn("转换这一批：1 个文件｜单独输出", log)
        self.assertNotIn("合并选项", log)
        # 「单独」落在盘上的样子就是只有一份、叫 demo.docx：合并那一份根本不该出现。
        self.assertEqual([path.name for path in self.written_docx()], ["demo.docx"])

    # ---- 跑批：后台线程真读写 -------------------------------------------------

    def test_a_real_batch_writes_the_docx_and_shows_the_report(self) -> None:
        """需求第 3、5、6、7 条在同一批里一次核：文件真写出来、统计进报告、日志实时进中间那一栏。

        报告对话框只画执行体交回的那一份 dict，所以这一条同时钉住
        "屏幕上的数 == 盘上的文件"——两边各数一遍就会出现"报告说 2 份、盘上 1 份"。
        """
        seen_while_running: list[str] = []
        real_log = self.win.log

        def spy(line: str) -> None:
            # 记的是"这一行到窗口的时候，这一批还在不在走"。跑完了才一次性补上，
            # 那一栏就是空的——而屏幕中间一片空白正是需求第 6 条要治的那件事。
            if self.win._running:
                seen_while_running.append(line)
            real_log(line)

        self.win.log = spy
        self.win.files.add_paths([DEMO, self.second])
        self.win.start_convert()
        # 忙态是同步设的：这一拍就能核，不必等事件循环。
        self.assertFalse(self.win.act_convert.isEnabled())
        self.assertTrue(self.win.act_cancel.isEnabled())
        self.assertFalse(self.win.files.isEnabled())
        self.assertFalse(self.win.settings.isEnabled())
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        pump()
        written = sorted(path.name for path in self.written_docx())
        self.assertEqual(written, ["demo.docx", "second.docx"])
        self.assertEqual(self.win.progress.value(), 100)
        self.assertIn("转换完成：2 个文件成功、0 个失败", self.win.status_text())
        # 第 6 条：日志是"实时"来的——跑批那一路的每一行都在这一批还挂着的时候进了中间那一栏，
        # 而不是跑完了再一次性补上（补上的那种，人盯着屏幕的中间什么也看不见）。
        self.assertGreater(len(seen_while_running), 4)
        self.assertTrue(any(line.startswith("读完：") for line in seen_while_running))
        self.assertTrue(any(line.startswith("写出：") for line in seen_while_running))
        self.assertIn("读完：demo.md", self.win.log_text())
        report = self.win._report_dialog()
        self.assertTrue(report.isVisible())
        self.assertEqual(report.windowTitle(), "转换报告")
        self.assertEqual(report.receipt_text(), self.win.status_text())
        # 第 5 条的四个统计：成功 / 失败 / 公式 / 表格，概要那一格与回执念的是同一批数。
        head = dict(pair for pair in report.head_pairs())
        self.assertEqual(head["成功几个文件"], "2 个")
        self.assertEqual(head["失败几个文件"], "0 个")
        self.assertIn("公式数", head)
        self.assertIn("条", head["公式数"])
        self.assertIn("张", head["表格数"])
        self.assertEqual(sorted(row[0] for row in report.output_rows()), ["demo", "second"])
        self.assertEqual(len(report.output_rows()), len(self.written_docx()))
        # 图片这一条在屏幕上要有三个数：只念"嵌进 1"就把"有两张没成"演成没事，
        # 而 demo.docx 里那个媒体文件是这三个数的盘上证据（回执、概要、逐份表三处念的同一套数）。
        self.assertEqual(head["图片"], "3 处（嵌进 1，没成 2）")
        self.assertIn("图片 3 处（嵌进 1，没成 2）", self.win.status_text())
        pictures = dict((row[0], row[5]) for row in report.output_rows())
        self.assertEqual(pictures, {"demo": "3 处（嵌进 1，没成 2）",
                                    "second": "0 处（嵌进 0，没成 0）"})
        with zipfile.ZipFile(self.out_dir / "demo.docx") as archive:
            self.assertIn("word/media/image1.png", archive.namelist())
        sentences = "｜".join(row[1] for row in report.problem_rows())
        self.assertIn("这台机器上没找到这个文件", sentences)
        self.assertIn("这不是本机的文件", sentences)

    def test_the_worker_carries_the_four_signals_named_in_the_request(self) -> None:
        """需求第 2 条点名的那三个信号（外加界面要的 log）都得是 ConvertWorker 自己的信号。

        ``finished`` 这里**盖掉了** ``QThread.finished``：请求要的是"带统计的那一个"，
        所以签名是 ``(dict)`` 而不是 Qt 原本那个无参的。这一条把盖掉这件事钉在测试里——
        有人日后写 ``worker.finished.connect(worker.deleteLater)`` 就会把 dict 交给 deleteLater。
        """
        signals = {}
        for name in ("progress", "log", "finished", "error"):
            signal = getattr(ConvertWorker, name, None)
            self.assertIsNotNone(signal, "ConvertWorker 少了信号 %s" % name)
            signals[name] = signal
        self.assertEqual(signals["progress"].signatures, ("progress(int,QString)",))
        self.assertEqual(signals["log"].signatures, ("log(QString)",))
        self.assertEqual(signals["finished"].signatures, ("finished(QVariantMap)",))
        self.assertEqual(signals["error"].signatures, ("error(QString)",))
        # 盖掉的那一个与 Qt 原本那个不是同一条：Qt 的 finished 不带参数。
        self.assertNotEqual(ConvertWorker.finished, QThread.finished)
        self.assertEqual(QThread.finished.signatures, ("finished()",))

    def test_a_batch_that_cannot_start_frees_the_buttons_and_opens_no_report(self) -> None:
        """样式表找不到 → 走 ``error`` 那一路：忙态照样收、报告不弹、盘上一个新文件都没有。

        这一条钉的是"出错之后按钮永远灰着"那一类：``_on_error`` 与 ``_on_finish`` 都得收忙态，
        少收一处，出一次错就得重启软件才能再跑。
        """
        real = tasks.locate_xslt
        tasks.locate_xslt = lambda *args, **kwargs: (_ for _ in ()).throw(
            tasks.XsltNotFoundError([self.tmp / "假的那一个目录"]))
        try:
            self.win.files.add_paths([DEMO])
            self.win.start_convert()
            self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        finally:
            tasks.locate_xslt = real
        pump()
        self.assertTrue(self.win.act_convert.isEnabled())
        self.assertFalse(self.win.act_cancel.isEnabled())
        self.assertTrue(self.win.files.isEnabled())
        self.assertEqual(self.win.progress.value(), 0)
        self.assertIn("这一批没跑成", self.win.status_text())
        self.assertIn("MML2OMML.XSL", self.win.log_text())
        self.assertIsNone(self.win._report, "没跑成的一批不该弹报告：那一张表里一格数都没有")
        self.assertEqual(self.written_docx(), [])
        self.assertFalse(self.out_dir.exists())

    def test_cancel_keeps_what_is_already_written_and_stops_short_of_the_end(self) -> None:
        tick = tasks.PAUSE_PER_FILE_SECONDS
        tasks.PAUSE_PER_FILE_SECONDS = 0.8
        try:
            self.win.files.add_paths([DEMO, self.second])
            self.win.start_convert()
            self.assertTrue(wait_until(lambda: self.win.progress.value() > 0, timeout=10))
            self.win.cancel_current()
            self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled(), timeout=20))
        finally:
            tasks.PAUSE_PER_FILE_SECONDS = tick
        written = sorted(path.name for path in self.written_docx())
        # 取消只停"还没开始的那几个"：已经写出去的那一份留着（不回滚），后面的不补。
        self.assertEqual(written, ["demo.docx"])
        self.assertLess(self.win.progress.value(), 100)
        self.assertIn("已取消", self.win.status_text())
        self.assertIn("1 个文件走完", self.win.status_text())
        self.assertIn("已经写出去的文件不回滚", self.win.log_text())
        report = self.win._report_dialog()
        self.assertTrue(report.isVisible())
        self.assertEqual(report.receipt_text(), self.win.status_text())
        self.assertIn("已取消", dict(pair for pair in report.head_pairs())["走到哪一步"])
        # 回执里那句"公式 N 条"与概要那一格必须是同一个意思（解析到几条），
        # 两处各数一套就会出现"回执 18、概要 19"而没人知道差的是转不出那一条。
        head = dict(pair for pair in report.head_pairs())
        self.assertIn(head["公式数"].split(" 条")[0], self.win.status_text())

    def test_the_report_window_is_reused_batch_after_batch(self) -> None:
        """两批下来屏幕上只有一扇报告：换内容不换窗口。

        每批开一扇的话，跑第五次转换屏幕上摞着五张表，人对着最早那一张核对刚跑的那一批。
        """
        self.win.files.add_paths([DEMO])
        self.win.start_convert()
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        first = self.win._report
        self.assertEqual(first.receipt_text(), self.win.status_text())
        self.win.files.add_paths([self.second])
        self.win.start_convert()
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        self.assertIs(self.win._report, first)
        self.assertEqual(sorted(path.name for path in self.written_docx()),
                         ["demo.docx", "second.docx"])
        self.assertIn("转换完成：2 个文件成功", self.win.status_text())
        self.assertEqual(first.receipt_text(), self.win.status_text())
        self.assertEqual(len(first.output_rows()), 2)

    # ---- 报告窗的画法（两句只有看画布才量得到）--------------------------------

    def test_a_long_problem_sentence_wraps_instead_of_being_cut_at_the_window_edge(self) -> None:
        """「说了什么」那一格装的是整句话，撑到窗口边读不完就等于这一栏没写。

        截图核出来的那一处生产缺陷：``setWordWrap(False)`` + 最后一格撑满时，
        latex2mathml 那句拒绝在窗口边上被画成"…（NoAvailableTokensError: 无…"。
        这一栏因此换行、行高按内容算；「逐份的账」那一栏相反——那里装的是整条输出路径，
        撑开反而读不完，给它的是横向滚动条。
        """
        report = self.win._report_dialog()
        report.set_stats(report_sample())
        report.show_report()
        pump()
        self.assertTrue(report.problem_table.wordWrap())
        self.assertFalse(report.output_table.wordWrap())

        report.resize(1180, 720)
        pump()
        wide = report.problem_table.rowHeight(0)
        report.resize(620, 720)
        pump()
        narrow = report.problem_table.rowHeight(0)
        # 压窄之后行高必须跟着长：没跟着长就是那句还是一次摆完、被裁在边上。
        self.assertGreater(narrow, wide, "窗口压窄之后行高没变：那句没重算，裁在窗口边上")

        # 行高"跟着长"还得真是"摊成几行"而不是"格子本身变高"：同一格换一条短句就退回一行。
        report.set_stats(report_sample(problem="没问题"))
        pump()
        self.assertGreater(narrow, report.problem_table.rowHeight(0),
                           "短句那一行一样高：那句长的没换行，是被裁掉的")
        report.close()

    def test_the_cancelled_receipt_is_painted_in_the_warning_colour(self) -> None:
        """取消那一路的回执得**真的**是红的，不是"名字改了"就算。

        换 ``objectName`` 之后 Qt 不会重算样式表（这一窗每批复用同一个实例），
        少了 ``unpolish`` + ``polish`` 那两句，画出来还是正文色——而 objectName、
        ``text()``、palette 三样都不报错，只有数画布上的墨色才发现得了。
        """
        report = self.win._report_dialog()
        report.set_stats(report_sample(cancelled=False))
        report.show_report()
        pump()
        plain = ink_rgb(report.receipt)
        report.set_stats(report_sample(cancelled=True))
        pump()
        painted = ink_rgb(report.receipt)

        warn = theme.color(self.win.theme_name(), "warn_fg")
        self.assertNotEqual(plain, painted, "两批墨色一样：样式表没跟着换")
        self.assertTrue(near(painted, (warn.red(), warn.green(), warn.blue())),
                        "取消那一张的墨色 %s 不是警告色 %s" % (painted, warn.name()))
        self.assertFalse(near(plain, (warn.red(), warn.green(), warn.blue())),
                         "没取消的那一张也红了：这一栏就分不出'这批出事了'")
        report.close()

    def test_nothing_is_checked_so_nothing_runs(self) -> None:
        self.win.files.add_paths([DEMO])
        self.win.files.list.set_all_checked(False)
        self.win.start_convert()
        pump()
        self.assertTrue(self.win.act_convert.isEnabled())
        self.assertFalse(self.win.act_cancel.isEnabled())
        self.assertIn("一个文件都没勾上", self.win.status_text())

    def test_validate_reads_the_file_and_names_the_formula_count(self) -> None:
        self.win.files.add_paths([DEMO])
        self.win.start_validate()
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        log = self.win.log_text()
        self.assertIn("公式", log)
        self.assertIn("demo.md", log)
        # 需求第 3 条"错误项高亮"落在**行**上：demo.md 那三处（一处图片路径 + 两处公式语法）
        # 都是"进了 Word 会出错"那一档，小结那一行念的是整轮的档位。
        tiers = dict((text.split("] ", 1)[1], kind) for text, kind in self.win.preview.log_lines())
        bad = [line for line, kind in self.win.preview.log_lines() if kind == "bad"]
        self.assertEqual(3 + 1, len(bad))  # 三处问题 + 那一句小结
        for probe in ["图片路径", "公式语法", "公式语法", "小结："]:
            self.assertTrue(any(probe in line for line in bad), probe)
        self.assertEqual("bad", tiers["小结：demo.md｜查出 3 处问题，其中 3 处会让 Word 里那份成果出错"])
        # 校验那一路与转换那一路的差别就在盘上：同一个窗口、同一套统计，一个都不写。
        self.assertEqual(self.written_docx(), [])
        self.assertFalse(self.out_dir.exists())
        report = self.win._report_dialog()
        self.assertEqual(report.windowTitle(), "校验报告")
        self.assertIn("什么都没动", dict(pair for pair in report.head_pairs())["盘上动了什么"])

    # ---- 分组输出（需求第 6 条的后半：一组一份 .docx）------------------------

    def test_grouped_output_writes_exactly_the_docx_it_named(self) -> None:
        """日志念"这批打算写到 2 份"的那两个名字，就是盘上多出来的那两个文件。

        念名字与写文件读的都是 :meth:`Options.output_plan` 那一份计划，所以这一条钉的是
        "屏幕上说的份数 == 盘上的文件名"——两边各拼一次就会出现"念 1-甲组、写甲组"。
        """
        self.win.files.add_group("甲组")
        self.win.files.add_paths([DEMO])
        self.win.files.add_group("乙组")
        self.win.files.add_paths([self.second])
        self.win.settings.apply_mode_name(MODE_GROUPED)
        self.assertEqual(self.win.settings.options().mode_name(), "分组输出")
        self.win.start_convert()
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        log = self.win.log_text()
        self.assertIn("这批打算写到 2 份", log)
        self.assertIn("甲组 → 1-甲组.docx", log)
        self.assertIn("乙组 → 2-乙组.docx", log)
        self.assertIn("每组选项", log)
        self.assertEqual(sorted(path.name for path in self.written_docx()),
                         ["1-甲组.docx", "2-乙组.docx"])
        rows = dict((row[0], row[1]) for row in self.win._report_dialog().output_rows())
        self.assertEqual(sorted(Path(target).name for target in rows.values()),
                         ["1-甲组.docx", "2-乙组.docx"])

    def test_grouped_mode_without_a_group_line_is_refused_before_any_thread_starts(self) -> None:
        """分组输出却一条分组线都没有：真实原因是"没线"，不是"没勾文件"，而且不许起线程。

        这一关在 ``_start`` 里、在 ``log(...)`` 之前，所以连日志那一行"转换这一批"都不该出现——
        念了那句再去说"其实没动"，人就会以为已经跑过一批。
        """
        self.win.files.add_paths([DEMO])
        self.win.settings.apply_mode_name(MODE_GROUPED)
        before = self.win.log_text()
        self.win.start_convert()
        pump()
        self.assertIn("得先有分组线", self.win.status_text())
        self.assertTrue(self.win.act_convert.isEnabled())
        self.assertFalse(self.win.act_cancel.isEnabled())
        self.assertEqual(self.win.log_text(), before)
        self.assertEqual(self.written_docx(), [])

    def test_an_empty_group_is_announced_so_the_count_is_not_read_as_all_groups_having_output(self) -> None:
        """某一组一个都没勾上 → 那份计划里它就是没有，而日志必须把这一组念出来。"""
        self.win.files.add_group("甲组")
        self.win.files.add_paths([DEMO])
        self.win.files.add_group("空组")
        self.win.settings.apply_mode_name(MODE_GROUPED)
        self.assertEqual(self.win.files.groups(), ["甲组", "空组"])
        self.win.start_convert()
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        log = self.win.log_text()
        self.assertIn("这一组一个文件都没勾上，所以它不出文件：空组", log)
        self.assertIn("这批打算写到 1 份", log)
        # 计划里就没有空的那一组：盘上只多出甲组那一份，"空组.docx" 这种空表不写出去。
        self.assertEqual([path.name for path in self.written_docx()], ["1-甲组.docx"])
        self.assertEqual([row[0] for row in self.win._report_dialog().output_rows()], ["甲组"])

    # ---- 忙态 ---------------------------------------------------------------

    def test_every_named_action_is_greyed_while_a_batch_is_running(self) -> None:
        """跑批期间十八颗具名动作（工具栏九 + 菜单九）一起灰，只留「取消这一批」。

        漏灰任何一颗，人就能在跑一半时改列表或再发一批——日志里的文件名就和盘上写出的对不上了。
        """
        tick = tasks.PAUSE_PER_FILE_SECONDS
        tasks.PAUSE_PER_FILE_SECONDS = 0.5
        try:
            self.win.files.add_paths([DEMO, self.second])
            self.win.start_convert()
            self.assertTrue(self.win._running)
            disabled = [name for name in TOOLBAR_OBJECT_NAMES + MENU_ACTION_OBJECT_NAMES
                        if not getattr(self.win, name).isEnabled()]
            self.assertEqual(sorted(disabled),
                             sorted(TOOLBAR_OBJECT_NAMES + MENU_ACTION_OBJECT_NAMES))
            self.assertTrue(self.win.act_cancel.isEnabled())
            self.assertFalse(self.win.files.isEnabled())
            self.assertFalse(self.win.settings.isEnabled())
            # 「最近打开的工程」那一整棵子菜单也在这一族里：点一行就是换一份工程、重列左边那摞文件。
            self.assertFalse(self.win.recent_menu.isEnabled())
            self.win.cancel_current()
            self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled(), timeout=10))
        finally:
            tasks.PAUSE_PER_FILE_SECONDS = tick
        self.assertTrue(self.win.files.isEnabled())
        self.assertTrue(self.win.settings.isEnabled())
        self.assertTrue(self.win.recent_menu.isEnabled())

    # ---- 主题 ---------------------------------------------------------------

    def test_the_dark_theme_replaces_the_stylesheet_and_the_unknown_name_falls_back(self) -> None:
        self.win.apply_theme(theme.THEME_DARK, announce=False)
        self.assertEqual(self.win.theme_name(), theme.THEME_DARK)
        self.assertIn("#1c2128", APP.styleSheet())
        self.assertTrue(self.win.theme_actions[theme.THEME_DARK].isChecked())
        self.win.apply_theme("没有这一档", announce=False)
        self.assertEqual(self.win.theme_name(), theme.THEME_LIGHT)
        self.assertFalse(self.win.theme_actions[theme.THEME_DARK].isChecked())

    # ---- 预览与校验（六期）----------------------------------------------------

    def preview_of(self, name: str, text: str) -> str:
        """写一份临时 .md、让中间那栏渲染它，交回屏幕上当前的 HTML。

        走 :meth:`MainWindow.show_preview_for` 而不是直接 call :func:`gui.preview_html.render_file`：
        这一栏要钉的是"人选中一行之后**屏幕上出现**的东西"，中间那几层（读盘、灌进 QTextBrowser、
        状态栏那句计数）都在这条路上——那几层接错了，只测渲染函数是发现不了的。
        """
        path = self.tmp / name
        path.write_text(text, encoding="utf-8")
        self.win.show_preview_for(path, force=True)
        pump()
        return self.win.preview.current_html()

    def test_a_display_formula_is_laid_out_and_a_bare_ampersand_costs_only_itself(self) -> None:
        html = self.preview_of("amp.md",
                               "# 标题\n\n$$x=\\frac{1}{2}$$\n\n前面 $a & b$ 后面 $c$ 都还在。\n")
        self.assertIn("<h1>标题</h1>", html)
        self.assertIn("<div class=\"math math-block\">x=<table class=\"math-frac\" "
                      "cellspacing=\"0\" cellpadding=\"0\"><tr><td class=\"math-num\">1</td></tr>"
                      "<tr><td class=\"math-den\">2</td></tr></table></div>", html)
        # 正文里带**裸的** & 时 latex2mathml 交回 <mi>&</mi>，那不是合法 XML。语法体检本身曾在这里
        # 抛穿，而 show_preview_for 没有 try（main_window.py:499）——一处坏就该只坏那一处，
        # 剩下的照旧摆出来。断言"其余部分还在"就是这一条的落点。
        self.assertIn("<span class=\"math\"><span class=\"math-broken\">a &amp; b</span></span>", html)
        self.assertIn("<span class=\"math\">c</span>", html)
        self.assertIn("前面", html)
        self.assertIn("都还在", html)
        self.assertEqual(html.count("&"), html.count("&amp;"))
        self.assertFalse(self.win.preview.browser.document().isEmpty())
        self.assertNotIn("读不动", html)
        self.assertNotIn("读不动", self.win.status_text())
        self.assertIn("预览 amp.md", self.win.status_text())

    def test_a_table_and_a_code_block_come_out_as_those_tags(self) -> None:
        """需求第 1 条后半"表格、标题、代码块都要正确显示"：对齐写在格上、代码块认得出语言。"""
        html = self.preview_of("grid.md",
                               "| 甲 | 乙 |\n|:---|---:|\n| 1 | 2 |\n\n```python\nprint(1)\n```\n")
        self.assertIn("<th align=\"left\">甲</th>", html)
        self.assertIn("<th align=\"right\">乙</th>", html)
        self.assertIn("<td align=\"left\">1</td>", html)
        self.assertIn("<td align=\"right\">2</td>", html)
        self.assertIn("<pre data-lang=\"python\"><code>print(1)</code></pre>", html)

    def test_an_aligned_table_still_draws_its_grid(self) -> None:
        """带对齐的那张表**在屏幕上真有线**——这条只能量像素，样式表文本对不等于线画出来了。

        原先格子写的是 ``style="text-align:left"``，而 Qt 的富文本里一个元素带了内联 ``style``
        就**不再合并**样式表里那一份声明：``th, td { border: 1px solid …; padding: 3px 7px }``
        整格没了，一张有对齐的表就一条格线都没有（量的：345 → 0，逐格加则 345→288→201）。
        改成 ``align`` 属性之后格线与不写对齐时相同（452），而右对齐那一列的墨点真的挪到了右半——
        所以这一条同时钉两件事：**线还在**、**对齐没丢**。
        """
        self.preview_of("grid-ink.md", "| 甲 | 乙 |\n|:---|---:|\n| 1 | 2 |\n")
        view = self.win.preview.browser
        border = theme.color(self.win.theme_name(), "border")
        self.assertGreaterEqual(ink_pixels(view, border), 40,
                                "带对齐的表又画不出格线了（内联 style 那一族的坑）")
        # 对齐真被 Qt 吃进去了（它自己重新序列化文档时把右对齐那两格念了出来）：换成属性写法之后
        # 这一条才必须钉——不然"格线回来了"完全可能只是"对齐没了"。
        serialised = view.document().toHtml()
        self.assertGreaterEqual(
            sum(serialised.count(word) for word in
                ("align=\"right\"", "text-align:right", "text-align: right")), 2,
            "右对齐没进文档：%s" % serialised[serialised.find("<table"):serialised.find("<table") + 200])
        # 公式那两张表反过来：格线该关掉才读得出"这是一条分数线"，同一把尺子量两头。
        self.preview_of("frac-ink.md", "$$\\frac{1}{2}$$\n")
        view = self.win.preview.browser
        self.assertLess(ink_pixels(view, border), 40, "公式表又长出格线了")
        self.assertGreaterEqual(ink_pixels(view, theme.color(self.win.theme_name(), "math_border")), 5,
                                "分数线没了")

    def test_every_level_of_a_nested_list_reaches_the_screen(self) -> None:
        """解析器认下几层就画几层：原先 `_list()` 只走一层，第三层在预览里凭空消失。"""
        html = self.preview_of("nest.md", "- 主干\n  - 分支\n    - 细节\n")
        self.assertIn("<ul><li>主干<ul><li>分支<ul><li>细节</li></ul></li></ul></li></ul>", html)
        self.assertEqual(3, html.count("<li>"))
        # 子层是 ol 还是 ul 由**孩子自己**那一行的记号说：拿父行的记号当子层的，圆点就演成了序号。
        mixed = self.preview_of("mixed.md", "1. 第一步\n   - 注意\n2. 第二步\n")
        self.assertIn("<ol><li>第一步<ul><li>注意</li></ul></li><li>第二步</li></ol>", mixed)

    def test_the_preview_stylesheet_drops_the_grid_but_keeps_the_fraction_bar(self) -> None:
        """公式那两张表的格线只有关得掉才读得出"这是一条分数线"，而 HTML 属性 ``border=0`` 是空操作。

        这里钉的是那份样式表的三件事：**每个令牌都替换成了颜色**（漏一个就原样念出 ``%(text)s``）、
        **两张公式表关掉格线**（``border-collapse`` 之外还要那句 ``border: none``，HTML 属性 ``border=0``
        是实测过的空操作），以及**分数线那条写成 ``.math-frac .math-num`` 而不是裸的 ``.math-num``**——
        靠的是**权重**（0-2-0 对 ``.math-frac td`` 的 0-1-1），不是排次序：所以这里钉的是"那一格的
        选择器带两级"，而不是"它排在后面"（注释里也写着同一个道理，改文件的人按次序去调就会调歪）。
        """
        for name in (theme.THEME_LIGHT, theme.THEME_DARK):
            css = theme.preview_css(name)
            self.assertNotIn("%(", css, "%s 这一版还有没替换的令牌" % name)
            self.assertIn("th, td { border: 1px solid %s; padding: 3px 7px; }"
                          % theme.color(name, "border").name(), css)
            self.assertIn(".math-frac td, .math-matrix td "
                          "{ border: none; padding: 0 5px; text-align: center; }", css)
            self.assertIn(".math-frac .math-num { border-bottom: 1px solid %s; }"
                          % theme.color(name, "math_border").name(), css)
            self.assertNotIn("\n.math-num {", css)
            self.assertIn(".math-broken { background: %s; color: %s; border: 1px solid %s; }"
                          % (theme.color(name, "warn_bg").name(),
                             theme.color(name, "warn_fg").name(),
                             theme.color(name, "warn_fg").name()), css)

    def test_each_validation_tier_is_painted_in_its_own_colour(self) -> None:
        """需求第 3 条"错误项高亮"落在**行**上：坏的那一档有墨，其余四档一格都没有。

        只看 ``log_lines()`` 里那个档位字符串是不够的——颜色是逐行刷进文档的
        （QSS 到不了 ``QTextCharFormat``），字符串对而上色错完全可能。所以一条一条单独重放、
        按档数 viewport 上的像素。
        """
        path = self.tmp / "mix.md"
        path.write_text("# 一级\n\n### 三级（跳了一级）\n\n坏公式 $\\frac{a}$。\n\n"
                        "| 甲 | 乙 | 丙 |\n|---|---|---|\n| 1 | 2 |\n", encoding="utf-8")
        self.win.files.add_paths([path])
        self.win.start_validate()
        self.assertTrue(wait_until(lambda: self.win.act_convert.isEnabled()))
        panel = self.win.preview
        # 先照一张底账：clear_log 会把读的那一列一起清掉（重放那几步要用到原文）。
        snapshot = list(panel.log_lines())
        tokens = sorted(set(LOG_KIND_TOKENS.values()))
        expected = {"公式语法": "bad", "标题层级": "warn", "表格列数": "warn", "小结：": "bad"}
        for probe, kind in expected.items():
            hits = [text for text, found in snapshot if probe in text]
            self.assertEqual(1, len(hits), "%s 那一行没查出来（文件写法变了？）" % probe)
            self.assertEqual(kind, dict(snapshot)[hits[0]])
            painted = LOG_KIND_TOKENS.get(kind, "text")
            panel.clear_log()
            panel.log(hits[0].split("] ", 1)[1], kind)
            pump()
            counts = dict((token, ink_pixels(panel.log_view, theme.color(panel.theme_name(), token)))
                          for token in tokens)
            self.assertGreaterEqual(counts[painted], 40,
                                    "%s 那一行没涂上 %s" % (probe, painted))
            for token, other in counts.items():
                if token != painted:
                    self.assertEqual(0, other, "%s 那一行串档了：%s 也有墨" % (probe, token))

    def test_switching_the_theme_repaints_the_log_and_replies_the_preview_css(self) -> None:
        """日志的颜色是刷进文档的，所以切主题必须**重画**，否则深底上还挂着深字。"""
        panel = self.win.preview
        panel.log("这一条是坏的", "bad")
        pump()
        light = theme.color(theme.THEME_LIGHT, "bad")
        self.assertGreaterEqual(ink_pixels(panel.log_view, light), 40)
        self.win.apply_theme(theme.THEME_DARK, announce=False)
        pump()
        self.assertGreaterEqual(ink_pixels(panel.log_view, theme.color(theme.THEME_DARK, "bad")), 40)
        self.assertEqual(0, ink_pixels(panel.log_view, light))
        self.assertIn("这一条是坏的", panel.log_text())
        self.assertIn("html, body { color: %s; }" % theme.color(theme.THEME_DARK, "text").name(),
                      panel.browser.document().defaultStyleSheet())

    def test_an_unknown_tier_keeps_the_line_and_paints_it_plain(self) -> None:
        """多一个档位名不该把日志区弄崩：底账原样留着那个词，画布按普通一行上色。"""
        panel = self.win.preview
        panel.log("档位是编的一行", "nonsense")
        pump()
        self.assertEqual([("档位是编的一行", "nonsense")],
                         [(text.split("] ", 1)[1], kind) for text, kind in panel.log_lines()])
        counts = dict((token, ink_pixels(panel.log_view, theme.color(panel.theme_name(), token)))
                      for token in sorted(set(LOG_KIND_TOKENS.values())))
        self.assertGreaterEqual(counts["text"], 40)
        for token, hits in counts.items():
            if token != "text":
                self.assertEqual(0, hits, "认不出的档位串到了 %s" % token)

    # ---- 工程文件（需求第 7 条）----------------------------------------------

    def use_project(self, name: str = "工程.json") -> Path:
        """把「打开 / 另存为」那一拍文件框换成一个固定路径。

        真点系统对话框测试做不到（而 ``--check`` 与 CI 上根本没有框），所以换掉的是
        :meth:`MainWindow._ask_project_path` 这一个出口——上面那两格照旧走原路。
        """
        target = self.tmp / name
        self.win._ask_project_path = lambda title, getter: target
        return target

    def test_a_saved_list_comes_back_with_the_same_order_checks_and_groups(self) -> None:
        first = self.make_md("a.md")
        second = self.make_md("b.md")
        self.win.files.add_group("甲组")
        self.win.files.add_paths([first, second])
        self.win.files.set_checked_for_rows([2], False)
        target = self.use_project()
        self.win.save_project_as()
        self.assertTrue(target.is_file())
        self.assertIn("工程文件写好了", self.win.log_text())
        self.win.files.clear_files()
        self.assertEqual(self.win.files.count(), 0)

        self.assertTrue(self.win.open_project(target))
        listing = self.win.files.list
        self.assertEqual([listing.row_kind(row) for row in range(listing.count())],
                         ["group", "file", "file"])
        self.assertEqual(self.win.files.groups(), ["甲组"])
        self.assertEqual([path.name for path in self.win.files.checked_paths()], ["a.md"])
        self.assertEqual(listing.item(2).checkState(), Qt.CheckState.Unchecked)
        self.assertEqual(self.win.project_path(), target)

    def test_a_broken_project_file_leaves_the_list_exactly_as_it_was(self) -> None:
        """读不懂的那一份**不许先动列表**：先读、读通了才交给列表。

        反过来（先清列表再报错）就是把人正在用的那份清单弄没了，而屏幕上只剩一句"读不懂"。
        """
        self.win.files.add_paths([DEMO])
        bad = self.tmp / "坏了.json"
        bad.write_text('{"kind": "别的程序", "version": 1, "entries": []}', encoding="utf-8")
        self.assertFalse(self.win.open_project(bad))
        self.assertEqual([path.name for path in self.win.files.checked_paths()], ["demo.md"])
        self.assertIn("读不懂", self.win.status_text())
        self.assertIn("一个字没动", self.win.log_text())

    def test_the_title_star_tracks_unsaved_edits_and_the_name_appears_after_the_first_save(self) -> None:
        """星号是"这份列表动过而还没存"唯一一处一直看得见的位置；存过之后标题写的是那一份。"""
        self.win.files.add_paths([DEMO])
        self.assertTrue(self.win.is_dirty())
        self.assertTrue(self.win.window_title().endswith("*"))
        self.assertIn("未保存的工程", self.win.window_title())

        target = self.use_project()
        self.win.save_project_as()
        self.assertFalse(self.win.is_dirty())
        self.assertFalse(self.win.window_title().endswith("*"))
        self.assertIn(target.name, self.win.window_title())

        self.win.files.add_group("甲组")
        self.assertTrue(self.win.is_dirty())
        self.assertTrue(self.win.window_title().endswith("*"))

    def test_an_edit_that_changed_nothing_does_not_light_the_title_star(self) -> None:
        """星号只说"列表真的变了"：移不动的那一挪、一行没收的那一次拖入、本来就全没勾的那一次
        批量勾选、敲了同一个名字的改组名，都不许把它点亮。

        亮了就是骗人去点「保存」；而且没点保存直接关窗时那一次"写回人自己定过的那一份"会跟着被
        触发，盘上凭空多出一份内容与之前逐字节相同、时间却新了一截的工程文件。
        """
        first = self.make_md("a.md")
        second = self.make_md("b.md")
        self.win.files.add_group("甲组")
        self.win.files.add_paths([first, second])
        self.win.files.set_all_checked(False)
        target = self.use_project()
        self.win.save_project_as()
        self.assertFalse(self.win.is_dirty())
        listing = self.win.files.list
        self.assertEqual([listing.row_kind(row) for row in range(listing.count())],
                         ["group", "file", "file"])

        listing.clearSelection()
        listing.setCurrentItem(None)
        listing.setCurrentRow(1)
        self.assertEqual(self.win.files.move_selected(True), 0)          # 再往上就越过那条分组线
        self.assertEqual(self.win.files.add_paths([first, second]), 0)   # 两条都已经在列表里
        self.assertEqual(self.win.files.set_all_checked(False), 0)       # 本来就全没勾
        self.assertEqual(self.win.files.set_checked_for_rows([1], False), 0)
        self.assertTrue(listing.rename_group(0, "甲组"))                 # 敲的是同一个名字
        self.assertFalse(self.win.is_dirty())
        self.assertFalse(self.win.window_title().endswith("*"))
        self.assertEqual([listing.row_kind(row) for row in range(listing.count())],
                         ["group", "file", "file"])
        self.assertEqual(self.win.files.groups(), ["甲组"])
        self.assertEqual([listing.item(row).checkState() for row in (1, 2)],
                         [Qt.CheckState.Unchecked, Qt.CheckState.Unchecked])

        # 反向对照：真改一笔就得亮。少了这一句，上面那一段等于把星号整条关了。
        self.assertTrue(listing.rename_group(0, "乙组"))
        self.assertTrue(self.win.is_dirty())

    def test_closing_writes_the_list_back_to_the_place_the_user_chose(self) -> None:
        """动过又没点保存就关窗 → 写回**人自己定过**的那一份，而不是弹窗、也不是另挑地方。"""
        first = self.make_md("a.md")
        target = self.use_project()
        self.win.files.add_paths([first])
        self.win.save_project_as()
        self.assertFalse(self.win.is_dirty())
        second = self.make_md("b.md")
        self.win.files.add_paths([second])
        self.win.close()
        pump()
        counts = project_file.load(target).counts()
        self.assertEqual(counts["files"], 2)

    def test_closing_without_a_chosen_place_does_not_invent_one(self) -> None:
        """没定过落点就不凭空找个目录写 .json：人找不到它、又以为存过了。星号是这一路的提醒。"""
        self.win.files.add_paths([DEMO])
        self.assertIsNone(self.win.project_path())
        self.win.close()
        pump()
        self.assertEqual(list(self.tmp.glob("*.json")), [])

    def test_the_last_project_comes_back_when_the_window_opens_again(self) -> None:
        """需求第 7 条的"下次打开时恢复"：记的是**那一份 .json 的路径**，启动时读回来。"""
        first = self.make_md("a.md")
        target = self.use_project()
        self.win.files.add_paths([first])
        self.win.save_project_as()
        self.win.close()
        pump()

        again = MainWindow(store=Store(settings=self.fresh_settings()))
        again.show()
        pump()
        try:
            self.assertEqual([path.name for path in again.files.checked_paths()], ["a.md"])
            self.assertEqual(again.project_path(), target)
            self.assertFalse(again.is_dirty())
            self.assertTrue(again.window_title().endswith(target.name))
        finally:
            again.close()
            pump()

    def test_a_moved_file_still_comes_back_as_a_row_with_the_gone_icon(self) -> None:
        """稿子被挪走不是"少一行"：那一行照样摆出来，图标是「已经不在了」，日志念得出是哪几条。"""
        staying = self.make_md("stays.md")
        going = self.make_md("goes.md")
        target = self.use_project()
        self.win.files.add_paths([staying, going])
        self.win.save_project_as()
        going.unlink()
        self.win.files.clear_files()

        self.assertTrue(self.win.open_project(target))
        self.assertEqual(self.win.files.count(), 2)
        statuses = [self.win.files.status_at(row) for row in range(self.win.files.row_count())]
        self.assertEqual(statuses, [theme.STATUS_READY, theme.STATUS_MISSING])
        self.assertIn("已经不在了", self.win.log_text())
        self.assertIn("goes.md", self.win.log_text())

    # ---- QSettings ----------------------------------------------------------

    def test_the_window_remembers_theme_and_the_three_column_split(self) -> None:
        # QSplitter 把 setSizes 给的数按比例摊到"当前可用宽"上，所以这里比的是比例；
        # 像素级的复原留给真开窗那一步（CI 上没有窗口管理器，show 时的宽会被夹掉）。
        self.win.resize(1400, 820)
        pump()
        self.win.splitter.setSizes([300, 600, 320])
        pump()
        settled = self.win.splitter_sizes()
        self.win.apply_theme(theme.THEME_DARK, announce=False)
        self.win.close()
        pump()

        on_disk = Store(settings=self.fresh_settings())
        self.assertEqual(on_disk.sizes(KEY_SPLITTER, []), settled)
        self.assertEqual(on_disk.text(KEY_THEME, ""), theme.THEME_DARK)

        again = MainWindow(store=Store(settings=self.fresh_settings()))
        again.show()
        pump()
        again.resize(1400, 820)
        pump()
        try:
            self.assertEqual(again.theme_name(), theme.THEME_DARK)
            for value, wanted in zip(again.splitter_sizes(), settled):
                self.assertAlmostEqual(value, wanted, delta=8)
        finally:
            again.close()


    # ---- 八期：本机设置、样式预设与配置抄本 ------------------------------------

    def store_on_disk(self) -> Store:
        """另开一个 QSettings 档读**盘上**的那一份（不是面板内存里的那一份）。

        "实时保存"那句只有从另一个读者眼里才验得成：读面板自己手里那份 Store 永远是对的，
        哪怕它一个字都没写下去。
        """
        return Store(settings=self.fresh_settings())

    def test_a_panel_edit_lands_in_the_local_store_within_a_beat(self) -> None:
        """"实时"两头都要钉：那一拍没到之前不许落盘（每敲一个字写一遍注册表），
        到了就必须真写下去（不 sync 就还是"关窗时才存"）。
        """
        self.win.settings.body_size.setValue(15.0)
        pump(0.2)
        self.assertNotEqual(self.store_on_disk().number(KEY_BODY_SIZE, 0.0), 15.0)
        pump(SAVE_DEBOUNCE_MS / 1000.0 + 0.5)
        self.assertEqual(self.store_on_disk().number(KEY_BODY_SIZE, 0.0), 15.0)

        moved = str(self.tmp / "另一个输出目录")
        self.win.settings.output_dir.setText(moved)
        pump(SAVE_DEBOUNCE_MS / 1000.0 + 0.5)
        self.assertEqual(self.store_on_disk().text(KEY_OUTPUT_DIR, ""), moved)
        # 一次落盘写的是**整屏**：字号不会因为"这次动的是目录"而被留在上一拍。
        self.assertEqual(self.store_on_disk().number(KEY_BODY_SIZE, 0.0), 15.0)

    def test_a_custom_preset_is_saved_adopted_and_dropped_through_the_panel(self) -> None:
        """存一档 / 点一档 / 撤一档走的全是面板上那三颗控件；下拉里演的是**推出来的**那一档。"""
        self.win.settings.ask_preset_name = lambda label, start: "我的档"
        self.win.settings.body_font.setCurrentText("仿宋")
        self.win.settings.body_size.setValue(11.0)
        self.win.settings.save_preset_button.click()

        self.assertEqual([item.name for item in self.win.settings.custom_presets()], ["我的档"])
        self.assertEqual(self.win.settings.preset.currentText(), "我的档")
        self.assertTrue(self.win.settings.drop_preset_button.isEnabled())
        self.assertEqual(self.win.settings.preset_hint.objectName(), "hint")
        self.assertIn("（自定义的档）", self.win.settings.preset_hint.text())
        self.assertEqual(json.loads(self.store_on_disk().text(KEY_STYLE_PRESETS, "")),
                         [{"body_font": "仿宋", "body_size": 11.0,
                           "heading_font": "黑体", "name": "我的档"}])

        # 点一档就把上面三格换成它那一套；内置那三档撤不掉，所以那颗按钮跟着灰。
        self.assertEqual(self.win.settings.preset.itemText(1), "学术论文")
        self.win.settings.preset.activated[int].emit(1)
        self.assertEqual(self.win.settings.body_font.currentText(), "Times New Roman")
        self.assertEqual(self.win.settings.body_size.value(), 12.0)
        self.assertEqual(self.win.settings.heading_font.currentText(), "Times New Roman")
        self.assertFalse(self.win.settings.drop_preset_button.isEnabled())
        self.assertIn("内置的三档之一", self.win.settings.preset_hint.text())

        # 三格改回那一套 → 下拉里演的名字是推出来的，不是哪一格存过的。
        self.win.settings.body_font.setCurrentText("仿宋")
        self.win.settings.body_size.setValue(11.0)
        self.win.settings.heading_font.setCurrentText("黑体")
        self.assertEqual(self.win.settings.preset.currentText(), "我的档")

        self.win.settings.drop_preset_button.click()
        self.assertEqual(self.win.settings.custom_presets(), [])
        self.assertEqual(json.loads(self.store_on_disk().text(KEY_STYLE_PRESETS, "")), [])
        self.assertFalse(self.win.settings.drop_preset_button.isEnabled())

    def test_a_custom_preset_identical_to_a_builtin_is_the_one_that_shows(self) -> None:
        """两条排版完全一样时演的是**自定义那一条**——撤不掉的档不该占住那颗「撤销」。"""
        self.assertTrue(self.win.settings.save_preset_as("我的技术文档"))
        self.assertEqual(self.win.settings.preset.currentText(), "我的技术文档")
        self.assertTrue(self.win.settings.drop_preset_button.isEnabled())

        self.win.settings.drop_preset_button.click()
        self.assertEqual(self.win.settings.custom_presets(), [])
        self.assertEqual(self.win.settings.preset.currentText(), "技术文档")
        self.assertFalse(self.win.settings.drop_preset_button.isEnabled())

    def test_a_refused_preset_name_is_spoken_in_the_panel_and_opens_no_dialog(self) -> None:
        """这一屏盖在主窗口上：被拒的那一句写进状态栏就等于没说，所以画在预设那一行下面且不弹框。"""
        hint = self.win.settings.preset_hint
        self.win.settings.ask_preset_name = lambda label, start: "技术文档"
        self.win.settings.save_preset_button.click()
        self.assertEqual(hint.objectName(), "problem")
        self.assertIn("内置的「技术文档」已经有这一档", hint.text())
        self.assertEqual(self.win.settings.custom_presets(), [])
        self.assertIsNone(APP.activeModalWidget())

        self.win.settings.ask_preset_name = lambda label, start: "   "
        self.win.settings.save_preset_button.click()
        self.assertIn("预设名字那一格是空的", hint.text())
        self.assertEqual(self.win.settings.custom_presets(), [])
        self.assertIsNone(APP.activeModalWidget())

    def test_a_corrupt_custom_preset_row_is_named_and_survives_an_ordinary_flush(self) -> None:
        """本机那一条读不懂 → 明说、当"一条都没有"，但**一次普通落盘不许把它顶成空**。

        顶掉的后果是"你自己存的那几档没了"，而那正是人最不会因为报错去查的一种损失；
        反过来静默交回空列表又等于骗人"这里从来就没有预设"。两头都要钉。
        """
        broken = '{"名字": "旧档"}'
        # 播种走**窗口自己手里那一份** Store：换一个 QSettings 实例去写，Close 时那一份
        # 把整张表回写一遍就会把这一格顶掉，测的就是"另一个读者"而不是"坏数据留不留得住"。
        self.win._store.put(KEY_STYLE_PRESETS, broken)
        self.win._store.sync()
        self.win.close()
        pump()

        again = MainWindow(store=self.store_on_disk())
        again.show()
        pump()
        try:
            hint = again.settings.preset_hint
            self.assertEqual(hint.objectName(), "problem")
            self.assertIn("本机存的自定义预设不是一个列表", hint.text())
            self.assertEqual(again.settings.custom_presets(), [])

            again.settings.body_size.setValue(13.0)
            pump(SAVE_DEBOUNCE_MS / 1000.0 + 0.5)
            self.assertEqual(self.store_on_disk().text(KEY_STYLE_PRESETS, ""), broken)
            # 坏了那一句一直挂着：改了字号也不许把它演成一句用法提示。
            self.assertIn("不是一个列表", hint.text())

            self.assertTrue(again.settings.save_preset_as("重来的档"))
            self.assertEqual(json.loads(self.store_on_disk().text(KEY_STYLE_PRESETS, "")),
                             [{"body_font": "宋体", "body_size": 13.0,
                               "heading_font": "黑体", "name": "重来的档"}])
            self.assertEqual(hint.objectName(), "hint")
            self.assertIn("重来的档", hint.text())
        finally:
            again.close()
            pump()

    def test_a_config_copy_written_from_the_panel_comes_back_on_a_second_window(self) -> None:
        """导出 / 导入这一对钉的是**带过去哪几格、没带过去哪几格**：排版与合并走，
        输出目录与主题不走——那两格是这台机器这块屏幕、这个目录的事。
        """
        copy = self.tmp / "抄本.json"
        self.win._ask_config_path = lambda title, getter, suggest_name=False: copy
        self.win.settings.body_font.setCurrentText("仿宋")
        self.win.settings.body_size.setValue(11.0)
        self.win.settings.merge_cover.setChecked(True)
        self.assertTrue(self.win.settings.save_preset_as("抄本里的一档"))
        self.win.apply_theme(theme.THEME_DARK, announce=False)

        self.win.export_config()
        self.assertIn("已导出配置", self.win.status_text())
        body = json.loads(copy.read_text(encoding="utf-8"))
        self.assertEqual(body["kind"], "md2docx.config")
        self.assertEqual(sorted(body), sorted(TOP_KEYS))
        self.assertEqual(sorted(body["settings"]), sorted(SETTINGS_KEYS))
        self.assertTrue(body["settings"]["cover_page"])
        self.assertEqual(body["style_presets"], [{"body_font": "仿宋", "body_size": 11.0,
                                                 "heading_font": "黑体", "name": "抄本里的一档"}])
        self.win.close()
        pump()

        again = MainWindow(store=self.store_on_disk())
        again.show()
        pump()
        try:
            other = str(self.tmp / "这边的输出目录")
            again.settings.output_dir.setText(other)
            again.settings.body_font.setCurrentText("楷体")
            again.apply_theme(theme.THEME_LIGHT, announce=False)
            pump(SAVE_DEBOUNCE_MS / 1000.0 + 0.5)

            self.assertTrue(again.import_config(copy))
            values = again.settings.style_values()
            self.assertEqual(values.body_font, "仿宋")
            self.assertEqual(values.body_size, 11.0)
            self.assertTrue(values.cover_page)
            self.assertEqual([item.name for item in again.settings.custom_presets()],
                             ["抄本里的一档"])
            self.assertEqual(again.settings.preset.currentText(), "抄本里的一档")
            # 没带过去的两格：输出目录与主题都还是这边自己定过的那一份。
            self.assertEqual(again.settings.output_dir.text(), other)
            self.assertEqual(again.theme_name(), theme.THEME_LIGHT)
            self.assertIn(CARRY_NOTE, again.log_text())
            # 导完立刻落盘，不等那一拍（"导入配置"不能是"只在屏幕上看看，重开就没了"）。
            self.assertEqual(self.store_on_disk().number(KEY_BODY_SIZE, 0.0), 11.0)
        finally:
            again.close()
            pump()

    def test_an_unreadable_config_copy_changes_nothing_on_the_panel(self) -> None:
        """先读通才动屏幕：读不懂的那一份不许把人在用的那套排版弄没。"""
        bad = self.tmp / "坏了的抄本.json"
        bad.write_text('{"kind": "别的程序", "version": 1, "settings": {}}', encoding="utf-8")
        before = self.win.settings.style_values()
        self.assertTrue(self.win.settings.save_preset_as("还在的档"))
        kept = [item.name for item in self.win.settings.custom_presets()]

        self.assertFalse(self.win.import_config(bad))
        self.assertIn("导入没走通", self.win.status_text())
        self.assertIn("不是 md2docx 的配置文件", self.win.status_text())
        self.assertIn("一个字没动", self.win.log_text())
        self.assertEqual(self.win.settings.style_values(), before)
        self.assertEqual([item.name for item in self.win.settings.custom_presets()], kept)

    def test_cancelling_the_config_dialogs_changes_nothing(self) -> None:
        """两张文件框都撤掉：屏幕不动、盘上一个新文件都不多。
        只有**写**的那一路给一个建议文件名——读的时候替人挑名字等于凭空多一条决定。
        """
        seen = []

        def refuse(title, getter, suggest_name=False):
            seen.append((title, suggest_name))
            return None

        self.win._ask_config_path = refuse
        before = self.win.settings.style_values()
        self.win.export_config()
        self.assertEqual(self.win.status_text(), "没挑地方，这份配置还没抄出去。")
        self.win.import_config()
        self.assertEqual(seen, [("导出配置文件", True), ("导入配置文件", False)])
        self.assertIn("没挑配置文件", self.win.status_text())
        self.assertEqual(self.win.settings.style_values(), before)
        self.assertEqual(list(self.tmp.glob("*.json")), [])

    def test_the_recent_list_remembers_projects_prunes_and_clears(self) -> None:
        """需求第 1 条的"最近打开的文件列表"记的是**工程文件**：那列稿子清单本来就在 .json 里。"""
        first = self.make_md("a.md")
        a = self.use_project("工程A.json")
        self.win.files.add_paths([first])
        self.win.save_project_as()
        b = self.use_project("工程B.json")
        self.win.save_project_as()
        self.win.save_project_as()      # 同一条不占两格，也不因为"又存了一次"重排

        self.assertEqual([Path(text).name for text in self.win.recent_projects()],
                         ["工程B.json", "工程A.json"])
        rows = [action.text() for action in self.win.recent_menu.actions() if action.text()]
        self.assertEqual(rows, self.win.recent_projects() + ["清空这张表"])

        b.unlink()
        self.win._rebuild_recent_menu()
        self.assertEqual(self.win.recent_projects(), [os.path.normpath(str(a))])
        self.assertIn("已经不在这个位置", self.win.log_text())
        self.assertEqual(self.store_on_disk().texts(KEY_RECENT_PROJECTS),
                         [os.path.normpath(str(a))])

        self.win.clear_recent_projects()
        self.assertEqual(self.win.recent_projects(), [])
        items = [action for action in self.win.recent_menu.actions() if action.text()]
        self.assertEqual([action.text() for action in items], ["（还没有打开过工程文件）"])
        self.assertFalse(items[0].isEnabled())
        self.assertTrue(a.is_file())

    def test_a_row_in_the_recent_list_opens_that_project(self) -> None:
        """点一行就是"打开那一份工程"：列里那条路径是入口，不是一句回显。"""
        first = self.make_md("a.md")
        second = self.make_md("b.md")
        a = self.use_project("工程A.json")
        self.win.files.add_paths([first])
        self.win.save_project_as()
        b = self.use_project("工程B.json")
        self.win.files.add_paths([second])
        self.win.save_project_as()
        self.win.files.clear_files()
        self.assertEqual(self.win.files.count(), 0)

        target = self.win.recent_menu.actions()[1].text()
        self.assertEqual(Path(target).name, "工程A.json")
        self.win.recent_menu.actions()[1].trigger()
        pump()

        self.assertEqual([path.name for path in self.win.files.checked_paths()], ["a.md"])
        self.assertEqual(self.win.project_path(), a)
        # 念的是日志不是状态栏：open_project 之后 select_first() 会把状态栏换成"预览 …"那一句。
        self.assertIn("打开工程", self.win.log_text())
        self.assertIn(a.name, self.win.log_text())
        self.assertEqual(Path(self.win.recent_projects()[0]).name, "工程A.json")


def make_options(**changes) -> Options:
    """造一份右侧面板的取值：只给必填那五格，其余按 ``changes`` 盖。

    ``Options`` 是**面板 → 线程唯一的通道**（一个 frozen dataclass），所以"出几份、叫什么名"
    这一整套可以在这里不开窗就钉死——真引擎接线时读的仍然是同一个 :meth:`Options.output_plan`。
    """
    base = dict(output_dir=Path("输出"), merged=False, body_font="宋体",
                body_size=12.0, heading_font="黑体")
    base.update(changes)
    return Options(**base)


class OutputPlanTest(unittest.TestCase):
    """三档输出模式各自"出几份 .docx、叫什么名"——需求第 6 条的那半在这里，另一半在界面上。"""

    def setUp(self) -> None:
        self.a = Path("甲/a.md")
        self.b = Path("甲/b.md")

    def test_separate_output_is_one_docx_per_file_named_after_the_file(self) -> None:
        plan = make_options().output_plan([self.a, self.b])
        self.assertEqual([label for label, _, _ in plan], ["a", "b"])
        self.assertEqual([target.name for _, target, _ in plan], ["a.docx", "b.docx"])
        self.assertEqual([len(members) for _, _, members in plan], [1, 1])

    def test_merged_output_is_one_docx_holding_every_checked_file(self) -> None:
        options = make_options(merged=True)
        plan = options.output_plan([self.a, self.b])
        self.assertEqual(len(plan), 1)
        label, target, members = plan[0]
        self.assertEqual((label, target.name), ("合并输出", "合并输出.docx"))
        self.assertEqual(list(members), [self.a, self.b])
        self.assertEqual(options.mode_name(), "合并输出")

    def test_grouped_output_numbers_the_files_so_two_same_named_groups_cannot_overwrite(self) -> None:
        """组名就是文件名，所以两份都得在：序号前缀是"两组同名互相顶掉"唯一的拦法。"""
        plan = make_options(grouped=True,
                            groups=[("***", [self.a]), ("???", [self.b])]).output_plan([self.a, self.b])
        self.assertEqual([target.name for _, target, _ in plan],
                         ["1-未命名.docx", "2-未命名.docx"])
        self.assertEqual([label for label, _, _ in plan], ["***", "???"])

    def test_a_group_name_is_disinfected_for_the_filesystem_but_never_left_empty(self) -> None:
        plan = make_options(grouped=True, groups=[
            ("甲/组:名", [self.a]), ("结尾. ", [self.b])]).output_plan([self.a, self.b])
        self.assertEqual([target.name for _, target, _ in plan],
                         ["1-甲_组_名.docx", "2-结尾.docx"])

    def test_a_group_with_nothing_in_it_does_not_become_an_empty_docx(self) -> None:
        """空组由界面上那一行念出来（``empty_groups``），不摆进计划：出一份空的 .docx 会被当成成果。"""
        plan = make_options(grouped=True,
                            groups=[("甲组", [self.a]), ("空组", []), ("只剩个名",)]
                            ).output_plan([self.a])
        self.assertEqual([label for label, _, _ in plan], ["甲组"])

    def test_the_target_named_for_a_batch_is_the_first_of_its_files(self) -> None:
        """``target_for`` 答的是"这批写到哪一个文件"——单独/合并/分组三档各有各的答案。"""
        out = Path("输出")
        self.assertEqual(make_options().target_for([self.a, self.b]), out / "a.docx")
        self.assertEqual(make_options(merged=True).target_for([self.a]), out / "合并输出.docx")
        self.assertEqual(make_options(grouped=True,
                                      groups=[("甲组", [self.a])]).target_for([self.a]),
                         out / "1-甲组.docx")
        self.assertEqual(make_options().target_for([]), out / "未命名.docx")

    def test_the_one_line_summary_names_the_mode_and_who_the_four_toggles_apply_to(self) -> None:
        """合并与分组共用那四个开关，念法不一样：分组输出时它们管的是"每一份"。"""
        merged = make_options(merged=True).summary()
        grouped = make_options(grouped=True, groups=[("甲组", [self.a])]).summary()
        separate = make_options().summary()
        self.assertIn("合并选项", merged)
        self.assertIn("每组选项", grouped)
        self.assertNotIn("选项", separate)
        self.assertIn("正文 宋体 12g", separate)
        self.assertIn("11.5g", make_options(body_size=11.5).summary())


class ProjectFileTest(unittest.TestCase):
    """工程文件那一份 .json：存读是同一条判法，认不出来的键一律拒。

    这一类不需要 QApplication（``gui.project_file`` 只用标准库），所以它不跟着开窗的
    ``skipIf`` 走——把不需要窗口的用例塞进 skip 里，等于在无显示机器上把一半覆盖丢掉。
    """

    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="md2docx-proj-test-")).resolve()
        self.inside = self.tmp / "稿子.md"
        self.inside.write_text("# 稿子\n", encoding="utf-8")
        self.target = self.tmp / "工程.json"

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def entries(self) -> list:
        return [{"group": "甲组"},
                {"path": str(self.inside), "checked": True},
                {"path": str(DEMO), "checked": False}]

    def rejected(self, entries, needle: str) -> str:
        """写这一路：交回那句拒绝的话（顺便钉死"抛的是 :class:`ProjectFileError`"）。"""
        with self.assertRaises(project_file.ProjectFileError) as caught:
            project_file.save(entries, self.target)
        self.assertIn(needle, str(caught.exception))
        return str(caught.exception)

    def read_rejects(self, document: str, needle: str) -> str:
        self.target.write_text(document, encoding="utf-8", newline="\n")
        with self.assertRaises(project_file.ProjectFileError) as caught:
            project_file.load(self.target)
        self.assertIn(needle, str(caught.exception))
        return str(caught.exception)

    # ---- 存与读 ------------------------------------------------------------

    def test_a_saved_list_reads_back_with_groups_checks_and_order(self) -> None:
        written = project_file.save(self.entries(), self.target, saved_at="2026-10-15T10:00:00")
        self.assertEqual(Path(written), self.target)
        project = project_file.load(self.target)
        self.assertEqual([entry.get("group") or Path(entry["path"]).name
                          for entry in project.entries],
                         ["甲组", "稿子.md", "demo.md"])
        self.assertEqual(project.entries[1]["checked"], True)
        self.assertEqual(project.entries[2]["checked"], False)
        self.assertEqual(project.counts(), {"files": 2, "groups": 1, "rows": 3})
        self.assertIn("2 个文件｜1 组", project.receipt())

    def test_read_back_paths_are_absolute_whatever_the_archive_wrote(self) -> None:
        """在这个目录里的稿子写成相对（整目录拷走照样打得开），读回来一律是绝对路径。"""
        project_file.save(self.entries(), self.target)
        body = json.loads(self.target.read_text(encoding="utf-8"))
        self.assertFalse(Path(body["entries"][1]["path"]).is_absolute())
        self.assertTrue(Path(body["entries"][2]["path"]).is_absolute())
        project = project_file.load(self.target)
        for entry in project.entries[1:]:
            self.assertTrue(Path(entry["path"]).is_absolute())
        self.assertEqual(Path(project.entries[1]["path"]), self.inside)

    def test_the_same_list_saved_twice_is_the_same_bytes(self) -> None:
        """同一份列表存两次逐字节相同、换行是 ``\\n``、中文不转义——不然"这份改没改"就永远比不出来。"""
        first = project_file.save(self.entries(), self.target, saved_at="2026-10-15T10:00:00")
        text = Path(first).read_bytes()
        second = project_file.save(self.entries(), self.target, saved_at="2026-10-15T10:00:00")
        self.assertEqual(text, Path(second).read_bytes())
        self.assertNotIn(b"\r\n", text)
        self.assertIn("甲组".encode("utf-8"), text)
        self.assertNotIn(rb"\u7532", text)

    def test_a_file_saved_by_notepad_with_a_bom_still_reads(self) -> None:
        self.target.write_bytes(b"\xef\xbb\xbf"
                                + json.dumps({"kind": project_file.KIND, "version": 1,
                                              "entries": []}, ensure_ascii=False).encode("utf-8"))
        self.assertEqual(project_file.load(self.target).entries, [])

    # ---- 拒绝 ------------------------------------------------------------

    def test_writing_refuses_shapes_that_could_never_be_read_back(self) -> None:
        """能写出去的一定能读回来，所以写这一关先把三种坏形状拦下（序号不存：它就是位置）。"""
        self.rejected([["不是字典"]], "不是一项")
        self.rejected([{"group": "  "}], "组名是空的")
        self.rejected([{"group": "甲组", "path": str(self.inside)}], "要么是分组要么是文件")
        self.rejected([{"path": str(self.inside), "序号": 1}], "读不懂的键")
        self.rejected([{"path": str(self.inside), "checked": "是"}], "得是真或假")

    def test_reading_refuses_to_guess_and_names_the_offending_key(self) -> None:
        """认不出来就拒：一份读不懂的档案被"尽力解析"成半条列表，比报错难查得多。"""
        def doc(body: str) -> str:
            return ('{"kind": "%s", "version": 1, "entries": %s}' % (project_file.KIND, body))

        self.read_rejects("[1, 2]", "顶层得是一个对象")
        self.read_rejects('{"kind": "别的东西", "version": 1, "entries": []}', "不是 md2docx 的工程文件")
        self.read_rejects('{"kind": "%s", "version": "1", "entries": []}' % project_file.KIND,
                          "得是一个整数")
        self.read_rejects('{"kind": "%s", "version": 99, "entries": []}' % project_file.KIND,
                          "读不懂第 99 版")
        self.read_rejects('{"kind": "%s", "version": 1}' % project_file.KIND, "没有 entries")
        self.read_rejects('{"kind": "%s", "version": 1, "entries": {}}' % project_file.KIND,
                          "得是一个列表")
        self.read_rejects('{"kind": "%s", "version": 1, "entries": [], "备注": "x"}'
                          % project_file.KIND, "读不懂的键")
        self.read_rejects(doc('[{"notes": "x"}]'), "既不是分组也不是文件")
        self.read_rejects(doc('[{"group": ""}]'), "group 那一格是空的")
        self.read_rejects(doc('[{"path": "a.md", "order": 1}]'), "读不懂的键")
        self.read_rejects(doc('[{"path": "  "}]'), "path 那一格是空的")
        self.read_rejects(doc('[{"path": "a.md", "checked": 1}]'), "得是真或假")

    def test_a_file_that_is_not_json_at_all_is_reported_with_its_position(self) -> None:
        message = self.read_rejects('{"kind": "unfinished', "不是合法的 JSON")
        self.assertIn("第 1 行", message)

    def test_a_missing_file_and_a_non_utf8_file_are_two_different_sentences(self) -> None:
        missing = self.tmp / "没写过.json"
        with self.assertRaises(project_file.ProjectFileError) as caught:
            project_file.load(missing)
        # 那句下一步引的是菜单上真写着的词（「工程另存为…」）——这一条由开窗核对那一步钉住。
        self.assertIn("不在那儿", str(caught.exception))
        self.assertIn("工程另存为", str(caught.exception))
        self.tmp.joinpath("gbk.json").write_bytes("「编码」".encode("gbk"))
        with self.assertRaises(project_file.ProjectFileError) as caught:
            project_file.load(self.tmp / "gbk.json")
        self.assertIn("不是 UTF-8 文本", str(caught.exception))

    # ---- 默认落点 ----------------------------------------------------------

    def test_the_default_place_for_another_project_follows_the_manuscripts(self) -> None:
        """跟着稿子走，不跟着输出目录走：输出目录里全是成品，清单混进去会被当成又一份成果。"""
        self.assertEqual(project_file.default_project_path([]), Path.cwd() / "工程.json")
        sibling = self.tmp / "另一份.md"
        self.assertEqual(project_file.default_project_path([self.inside, sibling]),
                         self.tmp / "工程.json")
        # 跨盘比不出公共父目录（盘根多半没写权限，也比人稿子旁边难找）→ 退回当前目录。
        self.assertEqual(project_file.default_project_path([self.inside, Path("Z:/别机/a.md")]),
                         Path.cwd() / "工程.json")

    def test_the_file_dialog_filter_offers_the_project_suffix_first(self) -> None:
        self.assertIn("*.json", project_file.project_filter())
        self.assertIn("md2docx", project_file.project_filter())


if __name__ == "__main__":
    unittest.main()
