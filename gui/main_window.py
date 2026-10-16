# -*- coding: utf-8 -*-
"""主窗口：三栏 + 工具栏九颗 + 底部状态栏，把三个面板和后台执行体装起来。

这一层只认识**面板交回的东西**（路径、勾选状态、:class:`gui.tasks.Options`），
不知道 Markdown 怎么解析、.docx 怎么写——那两个动作连同后台线程壳都在 :mod:`gui.tasks`，
报告对话框只画它交回的那一份统计（:mod:`gui.report_dialog`）。这里剩下的只有"什么时候开始、
进度念到多少、忙态什么时候收掉"。

几条口径写在最前面，免得日后当成漏写：

* **转换走后台线程**：几十个文件逐个写下去，主线程一动不动就"未响应"，
  进度条也就白摆了。跑批期间只让「取消这一批」亮着。
* **忙态在 ``finished`` 与 ``error`` 两个信号上都要收掉**：``error`` 是"这一批根本没跑成"
  （样式表找不到），只接 ``finished`` 的话出一次错按钮就永远灰着。
* **``worker.finished`` 不是 ``QThread.finished``**：执行体按需求点名要 ``finished(dict)``，
  这个名字盖掉了 ``QThread`` 自带的那个无参 ``finished``，所以
  ``worker.finished.connect(worker.deleteLater)`` 这一行到处都在抄的写法在这里不成立
  （它会把统计 dict 交给 ``deleteLater``）。线程壳的销毁走 :meth:`_retire_worker`。
* **窗口尺寸与三栏宽度各存各的**：``saveState()`` 管的是工具栏/菜单条，
  **不包括 ``QSplitter`` 的分配**，所以三栏宽度得自己存进 QSettings（``window/splitter_sizes``）。
* **拖文件进窗口 anywhere 都收**：文件列表自己当然收，但人往往往中间那块空白处拖，
  所以预览区关掉 acceptDrops，让事件冒到主窗口转交给列表。
* **列表那份清单另有进出口**（:mod:`gui.project_file`）：顺序、勾选、分组线存成 .json，
  而样式与窗口尺寸仍走 QSettings。两本账分开是因为一份是"这一批稿子"、另一份是"这台机器上的偏好"，
  合成一处就会出现"拷走工程文件把别人的输出目录也带走了"。
* **本机设置与配置抄本也是两本账**（八期，:mod:`config.settings`）：前者走 QSettings、改动实时落盘，
  后者是给人拷去另一台机器的抄本——只带排版三格、合并那四个开关和自定义预设，
  输出目录 / 主题 / 窗口尺寸 / 那几张路径表都不带（逐条理由写在 :data:`config.settings.NOT_CARRIED`）。
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Iterable, Optional, Sequence

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QActionGroup
from PySide6.QtWidgets import (QApplication, QFileDialog, QLabel, QMainWindow, QMenu,
                               QMessageBox, QProgressBar, QSplitter, QToolBar, QWidget)

from config.settings import (CARRY_NOTE, SettingsError, config_filter, push_recent,
                             read_config, write_config)
from gui import preview_html, project_file, theme
from gui.file_panel import MARKDOWN_SUFFIXES, FilePanel
from gui.preferences import (KEY_GEOMETRY, KEY_PREVIEW_SPLITTER, KEY_PROJECT,
                             KEY_RECENT_PROJECTS, KEY_SPLITTER, KEY_THEME, KEY_WINDOW_STATE,
                             Store)
from gui.preview_panel import PreviewPanel
from gui.report_dialog import ReportDialog
from gui.settings_panel import SettingsPanel, default_output_dir
from gui.tasks import ACTION_CONVERT, ACTION_VALIDATE, ConvertWorker, Options

__all__ = ["MainWindow", "APP_TITLE", "IDLE_STATUS", "TOOLBAR_OBJECT_NAMES",
           "MENU_ACTION_OBJECT_NAMES"]

APP_TITLE = "Markdown 转 Word 桌面工具"
IDLE_STATUS = "就绪：左边挑文件，右边定样式，「开始转换」把勾上的那批写出去。"

#: 工具栏就这九颗（需求点名的那九个动作），objectName 固定下来给测试和样式用。
TOOLBAR_OBJECT_NAMES = ("act_add_files", "act_add_folder", "act_remove", "act_move_up",
                        "act_move_down", "act_clear", "act_convert", "act_preview",
                        "act_validate")

#: 只在菜单里出现的那几颗（三期新加的是列表与工程那七颗，八期再加导出/导入配置两颗——
#: 工具栏那九颗是需求点名的，不再长）。
#: 跑批期间与那九颗一起灰掉：中途换列表 / 换工程 / 换排版会让日志里的文件名与实际写出的对不上。
MENU_ACTION_OBJECT_NAMES = ("act_new_group", "act_check_selected", "act_uncheck_selected",
                            "act_refresh_status", "act_open_project", "act_save_project",
                            "act_save_project_as", "act_export_config", "act_import_config")

_WINDOW_DEFAULT_SIZE = (1180, 760)
_SPLITTER_WEIGHTS = (1, 3, 1)

#: 「导出配置」那张文件框里建议的名字。抄本是给人来回拷的，所以**不带时间戳**——
#: 每次导都换一个新名字会把那一格变成垃圾桶；什么时候抄的写在文件里的 ``saved_at`` 那一格。
CONFIG_FILE_NAME = "md2docx-config.json"


def file_filter() -> str:
    """文件框那个"只看 Markdown"的过滤串——扩展名只有 :data:`gui.file_panel.MARKDOWN_SUFFIXES` 一份口径。"""
    return "Markdown 文件 (%s);;所有文件 (*)" % " ".join("*" + suffix for suffix in sorted(MARKDOWN_SUFFIXES))


class MainWindow(QMainWindow):
    """那一扇窗。``show()`` 之前把 QSettings 里存的样子复原，``closeEvent`` 里再存回去。"""

    def __init__(self, store: Optional[Store] = None, theme_name: Optional[str] = None,
                 parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._store = store if store is not None else Store()
        self._theme = (theme_name or self._store.text(KEY_THEME, theme.DEFAULT_THEME)).strip().lower()
        if self._theme not in theme.theme_names():
            self._theme = theme.DEFAULT_THEME
        self._worker: Optional[ConvertWorker] = None
        self._running = False
        #: 报告对话框只造一扇、每批换内容（理由见 :mod:`gui.report_dialog`）。
        self._report: Optional[ReportDialog] = None
        #: 上一次演的是哪个文件；选中没换就不重读一遍（``force=True`` 才是「预览」那颗按钮）。
        self._previewed = ""
        #: 现在屏幕上这份列表对应盘上哪一份工程文件；None = 还没存过，「保存工程」就是另存为。
        self._project: Optional[Path] = None
        #: 列表动过而还没写回那份 .json —— 标题上的那颗星号。
        self._dirty = False

        self.setObjectName("mainWindow")
        self.setWindowTitle(APP_TITLE)
        self.resize(*_WINDOW_DEFAULT_SIZE)

        self._build_central()
        self._build_actions()
        self._build_toolbar()
        self._build_menu()
        self._build_status_bar()
        self._wire_panels()

        self.settings.load_from_store()
        self.apply_theme(self._theme, announce=False)
        self._restore_layout()
        self._set_title()
        self.set_status(IDLE_STATUS)
        self.files.select_first()
        # 先摆菜单里那一列（顺手剪掉盘上已经不在了的那几条），再恢复上次那份——
        # 恢复会走 ``_remember_project``，那条自己会重摆一次，所以次序反过来会念两遍同一句剪除。
        self._rebuild_recent_menu()
        self._restore_last_project()

    # ---- 面板与控件 ---------------------------------------------------------

    def _build_central(self) -> None:
        self.files = FilePanel(self)
        self.preview = PreviewPanel(self._theme, self)
        self.settings = SettingsPanel(self._store, self)

        self.splitter = QSplitter(Qt.Orientation.Horizontal, self)
        self.splitter.setObjectName("mainSplitter")
        self.splitter.setChildrenCollapsible(False)
        self.splitter.addWidget(self.files)
        self.splitter.addWidget(self.preview)
        self.splitter.addWidget(self.settings)
        self.splitter.setCollapsible(2, False)
        for index, weight in enumerate(_SPLITTER_WEIGHTS):
            self.splitter.setStretchFactor(index, weight)
        self.splitter.setSizes([280, 620, 300])

        self.setCentralWidget(self.splitter)

    def _build_actions(self) -> None:
        def make(name: str, text: str, slot, shortcut: str = "", tip: str = "") -> QAction:
            action = QAction(text, self)
            action.setObjectName(name)
            if shortcut:
                action.setShortcut(shortcut)
            if tip:
                action.setToolTip(tip)
                action.setStatusTip(tip)
            action.triggered.connect(slot)
            setattr(self, name, action)
            return action

        make("act_add_files", "添加文件", self.pick_files, "Ctrl+O",
             "挑一个或几个 .md 加进左边列表")
        make("act_add_folder", "添加文件夹", self.pick_folder, "Ctrl+Shift+O",
             "整个文件夹翻一遍，里面的 Markdown 都收进来")
        make("act_remove", "移除", self.remove_selected, "Delete",
             "把选中那一行从列表里去掉（盘上的文件不动）")
        make("act_move_up", "上移", lambda: self.move_selected(True), "Alt+Up",
             "选中那一行往上挪一格——合并输出时列表顺序就是章节顺序")
        make("act_move_down", "下移", lambda: self.move_selected(False), "Alt+Down",
             "选中那一行往下挪一格")
        make("act_clear", "清空", self.clear_files, "Ctrl+Shift+Backspace",
             "列表清空（盘上的文件不动）")
        make("act_convert", "开始转换", self.start_convert, "Ctrl+Return",
             "把勾上的那批写成 .docx")
        make("act_preview", "预览", self.preview_selected, "F5",
             "重新演一次选中那个文件的近似预览")
        make("act_validate", "校验", self.start_validate, "Ctrl+Shift+V",
             "只读检查：能不能读、解析出多少公式表格、样式表找不找得到")

        # 三期新加的那几颗：只在菜单里出现，工具栏那九颗是需求点名的，不再长。
        make("act_new_group", "新建分组…", self.new_group, "Ctrl+G",
             "在列表里插一条分组线——分组输出时一组出一份，组名就是文件名")
        make("act_check_selected", "勾上选中的行", lambda: self.check_selected(True), "Ctrl+Space",
             "选中几行就参与几行（Ctrl/Shift 加选、Shift 连选）")
        make("act_uncheck_selected", "取消选中参与", lambda: self.check_selected(False),
             "Ctrl+Shift+Space", "选中几行就从这一批里去掉几行（列表还在）")
        make("act_refresh_status", "重新核对状态", self.refresh_status, "Ctrl+R",
             "再看一遍盘上还在不在：文件被挪走、改名都在这一步念出来（只读，什么都不改）")
        make("act_open_project", "打开工程…", self.open_project, "Ctrl+Shift+E",
             "用一份 .json 换掉左边那列列表（顺序与勾选都按那一份摆）")
        make("act_save_project", "保存工程", self.save_project, "Ctrl+S",
             "把左边那列列表写回工程文件（.json）")
        make("act_save_project_as", "工程另存为…", self.save_project_as, "Ctrl+Shift+S",
             "换一份 .json 存这次这份列表")
        # 八期：配置文件是"这套排版到另一台机器上"，与工程文件（这份列表）各是各的进出口。
        make("act_export_config", "导出配置…", self.export_config, "",
             "把右边那一屏的排版、合并开关与自定义预设抄成一份 .json（不带输出目录与窗口尺寸）")
        make("act_import_config", "导入配置…", self.import_config, "",
             "用一份 .json 换掉右边那一屏（先读通，读不通就一个字都不动）")

        # 取消这一批不在那九颗里：跑批时才亮，平时灰着。
        self.act_cancel = QAction("取消这一批", self)
        self.act_cancel.setObjectName("act_cancel")
        self.act_cancel.setShortcut("Esc")
        self.act_cancel.setToolTip("停下来——已经写完的文件不回滚，剩下的不再动")
        self.act_cancel.triggered.connect(self.cancel_current)
        self.act_cancel.setEnabled(False)

    def _build_toolbar(self) -> None:
        bar = QToolBar("主工具栏", self)
        bar.setObjectName("mainToolBar")
        bar.setMovable(False)
        bar.setFloatable(False)
        bar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        for name in TOOLBAR_OBJECT_NAMES:
            bar.addAction(getattr(self, name))
            if name in ("act_move_down", "act_clear", "act_preview"):
                bar.addSeparator()
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, bar)

    def _build_menu(self) -> None:
        menu_file = self.menuBar().addMenu("文件(&F)")
        for action in (self.act_add_files, self.act_add_folder):
            menu_file.addAction(action)
        menu_file.addSeparator()
        # 工程文件那一组在菜单里，不在工具栏：那九颗是需求点名的"这一批稿子怎么走"，
        # 而 .json 管的是"这份列表怎么留住"，是另一件事。
        for action in (self.act_open_project, self.act_save_project, self.act_save_project_as):
            menu_file.addAction(action)
        # 最近打开的那一列：记的是**工程文件 .json**，不是 .md 稿子——稿子属于哪份清单本来就写在那份
        # .json 里，另记一遍就会出现"菜单说最近打开过 a.md，而它属于哪一批没人知道"。
        self.recent_menu = menu_file.addMenu("最近打开的工程")
        menu_file.addSeparator()
        for action in (self.act_export_config, self.act_import_config):
            menu_file.addAction(action)
        menu_file.addSeparator()
        for action in (self.act_remove, self.act_clear):
            menu_file.addAction(action)
        menu_file.addSeparator()
        menu_file.addAction(self.act_cancel)
        menu_file.addSeparator()
        quit_action = menu_file.addAction("退出")
        quit_action.setShortcut("Ctrl+Q")
        quit_action.triggered.connect(self.close)

        menu_list = self.menuBar().addMenu("列表(&L)")
        for action in (self.act_new_group, self.act_check_selected, self.act_uncheck_selected):
            menu_list.addAction(action)
        menu_list.addSeparator()
        menu_list.addAction(self.act_refresh_status)

        menu_view = self.menuBar().addMenu("视图(&V)")
        self.theme_group = QActionGroup(self)
        self.theme_group.setExclusive(True)
        self.theme_actions: dict[str, QAction] = {}
        for name in theme.theme_names():
            action = QAction("「%s」主题" % theme.theme_label(name), self)
            action.setObjectName("act_theme_%s" % name)
            action.setCheckable(True)
            action.setChecked(name == self._theme)
            action.triggered.connect(lambda checked=True, value=name: self.apply_theme(value))
            self.theme_group.addAction(action)
            menu_view.addAction(action)
            self.theme_actions[name] = action
        menu_view.addSeparator()
        clear_log = menu_view.addAction("清空日志")
        clear_log.triggered.connect(self.preview.clear_log)

        menu_help = self.menuBar().addMenu("帮助(&H)")
        about = menu_help.addAction("关于这一工具")
        about.triggered.connect(self.show_about)

    def _build_status_bar(self) -> None:
        bar = self.statusBar()
        bar.setObjectName("mainStatusBar")
        self.status_label = QLabel(IDLE_STATUS, bar)
        self.status_label.setObjectName("statusText")
        self.status_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.progress = QProgressBar(bar)
        self.progress.setObjectName("statusProgress")
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("%p%")
        self.progress.setFixedWidth(180)
        bar.addWidget(self.status_label, 1)
        bar.addPermanentWidget(self.progress)

    def _wire_panels(self) -> None:
        self.files.noticed.connect(self.log)
        self.files.counts_changed.connect(self._on_counts)
        self.files.selected_changed.connect(self.show_preview_for)
        self.files.add_files_requested.connect(self.pick_files)
        self.files.add_folder_requested.connect(self.pick_folder)
        self.files.clear_requested.connect(self.clear_files)
        self.files.convert_single.connect(self.convert_only)
        self.files.changed.connect(self._on_list_changed)
        self.settings.changed.connect(self._describe_convert)
        self._describe_convert()

    # ---- 工具栏那九颗 --------------------------------------------------------

    def pick_files(self) -> None:
        start = self._last_folder()
        got, _filter = QFileDialog.getOpenFileNames(self, "加入 Markdown 文件", start, file_filter())
        if got:
            self.files.add_paths(Path(path) for path in got)

    def pick_folder(self) -> None:
        start = self._last_folder()
        chosen = QFileDialog.getExistingDirectory(self, "加入整个文件夹", start)
        if chosen:
            self.files.add_paths([Path(chosen)])

    def remove_selected(self) -> None:
        """批量移除：选中几行去掉几行（组头选不中，所以永远不会把分组线单独删掉）。"""
        removed = self.files.remove_selected()
        if not removed:
            self.set_status("先选中左边一行再点「移除」（勾掉复选框是「不参与这一批」，不是移除）。")
            return
        self.log("从列表里去掉了 %d 行：%s（盘上的文件一个字都没动）。" % (
            len(removed), "、".join(path.name for path in removed[:5])
            + ("…" if len(removed) > 5 else "")))

    def move_selected(self, upwards: bool) -> None:
        where = "上" if upwards else "下"
        moved = self.files.move_selected(upwards)
        if not moved:
            self.set_status("%s移不动：要么没选中行，要么这一行已经到顶/到底了，"
                            "要么再挪一格就越过一条分组线（那是换一组，用右键「新建分组…」）。" % where)
        else:
            self.log("选中那 %d 行整块%s挪了一格（列表顺序就是合并时的章节顺序、分组的归属）。"
                     % (moved, where))

    def check_selected(self, checked: bool) -> None:
        """批量启用/禁用：选中几行就改几行的复选框。

        念"改了 N 行"那一句由列表自己发（``noticed``）——那一步只有它知道哪几行本来就是勾着的，
        这里再拼一遍就会出现"面板说改了 3 行、日志说 2 行"。
        """
        word = "勾上" if checked else "取消参与"
        if not self.files.selected_count():
            self.set_status("先选中左边几行（Ctrl 加选、Shift 连选）再点「%s」。"
                            "复选框代表「参与这一批」，所以「%s」不用选中——直接点掉那一格就行。"
                            % (word, word))
            return
        self.files.set_checked_for_selected(checked)

    def refresh_status(self) -> None:
        """把行首那枚状态图标按盘上的现状重核一遍（只读）。"""
        self.files.refresh_status()

    def new_group(self) -> None:
        self.files.prompt_group()

    def clear_files(self) -> None:
        removed = self.files.clear_files()
        self.log("列表清空，去掉了 %d 行（盘上的文件一个字都没动）。" % removed)
        self.preview.show_welcome()

    def preview_selected(self) -> None:
        path = self.files.selected_path()
        if path is None:
            self.set_status("左边还没有选中的文件——先挑一行，或者点「添加文件」。")
            return
        self.show_preview_for(path, force=True)

    def start_convert(self) -> None:
        self._start(ACTION_CONVERT, self.files.checked_paths(), self._options())

    def start_validate(self) -> None:
        self._start(ACTION_VALIDATE, self.files.checked_paths(), self._options())

    def _options(self) -> Options:
        """面板的取值 + 列表交回的分组计划。**分组归属只由列表决定，这里不参与判断。**"""
        return replace(self.settings.options(), groups=self.files.group_plans())

    def convert_only(self, path: Path) -> None:
        """右键「单独转换这一条」：强制走单独输出——勾了合并或分了组还只出一份，那就不是"单独"了。"""
        options = replace(self.settings.options(), merged=False, grouped=False, groups=())
        self._start(ACTION_CONVERT, [Path(path)], options)

    def cancel_current(self) -> None:
        worker = self._worker
        if worker is None:
            return
        worker.requestInterruption()
        self.set_status("正在取消：已经写完的不回滚，剩下的不再动。")
        self.log("收到取消请求，等手上那一个文件走完。")

    # ---- 工程文件（那份列表的进出口） ----------------------------------------

    def open_project(self, path: Optional[Path] = None) -> bool:
        """用一份 .json 换掉左边那列列表；读不懂就**一个字都不动列表**。

        次序是这一条的全部要紧处：先读、读通了才交给列表。反过来（先清列表再报错）
        就把人正在用的那份清单弄没了，而屏幕上只剩一句"读不懂"。
        """
        target = Path(path) if path is not None else self._ask_project_path(
            "打开工程文件", QFileDialog.getOpenFileName)
        if target is None:
            self.set_status("没挑工程文件，左边那份列表照原样留着。")
            return False
        try:
            project = project_file.load(target)
        except project_file.ProjectFileError as exc:
            self.set_status("这份工程文件读不懂：%s" % exc)
            self.log("打开工程文件失败：%s（左边那份列表一个字没动）" % exc, "bad")
            return False
        result = self.files.load_entries(project.entries)
        self._project = Path(project.source)
        self._dirty = False
        self._remember_project(self._project)
        self._set_title()
        count = project.counts()
        self.log("打开工程 %s｜%s" % (self._project, project.receipt()))
        if count["rows"] != result["files"] + result["groups"]:
            # 档案里有 12 条、列表只摆了 10 行——差的那两行必须当场说清为什么没摆（见 skipped）。
            self.log("这份里有 %d 条没摆进列表：%s" % (
                count["rows"] - result["files"] - result["groups"], "；".join(result["skipped"])), "warn")
        if result["missing"]:
            self.log("这 %d 条 .md 已经不在那个位置了（那一行照样摆着，图标是「已经不在了」）：%s" % (
                len(result["missing"]), "、".join(result["missing"][:5])
                + ("…" if len(result["missing"]) > 5 else "")), "warn")
        self.set_status("已打开 %s：%d 行列表%s。" % (
            self._project.name, result["files"] + result["groups"],
            "（%d 条已经不在了）" % len(result["missing"]) if result["missing"] else ""))
        self.files.select_first()
        return True

    def save_project(self) -> None:
        """已有落点就写那一份，没有就走「另存为」（第一次保存本来就得挑地方）。"""
        if self._project is None:
            self.save_project_as()
            return
        self._write_project(self._project)

    def save_project_as(self) -> None:
        target = self._ask_project_path("工程文件另存为", QFileDialog.getSaveFileName)
        if target is None:
            self.set_status("没挑地方，这份列表还没存成工程文件。")
            return
        self._write_project(target)

    def _write_project(self, target: Path) -> bool:
        entries = self.files.entries()
        try:
            written = project_file.save(entries, target)
        except (project_file.ProjectFileError, OSError) as exc:
            # 没权限 / 盘满 / 那是一条目录都算"没存上"，星号必须留着。
            self.set_status("工程文件没写出去：%s" % exc)
            self.log("工程文件没写出去：%s（列表还在，星号也还在）" % exc, "bad")
            return False
        self._project = Path(written)
        self._dirty = False
        self._remember_project(self._project)
        self._set_title()
        groups = sum(1 for entry in entries if "group" in entry)
        self.log("工程文件写好了：%s｜%d 个文件、%d 组（顺序与勾选都按屏幕上这一份）" % (
            written, len(entries) - groups, groups))
        self.set_status("已保存 %s。" % written.name)
        return True

    def _ask_project_path(self, title: str, getter) -> Optional[Path]:
        """文件框那一拍单独列出来，是为了让脚本与测试能把路径直接交给 ``open_project`` /
        ``save_project_as`` 的调用者，而不必真去点一个系统对话框。"""
        start = self._project or project_file.default_project_path(
            [Path(str(entry.get("path"))) for entry in self.files.entries() if "path" in entry])
        chosen, _filter = getter(self, title, str(start), project_file.project_filter())
        return Path(chosen) if chosen else None

    def _restore_last_project(self) -> None:
        """启动那一拍把上次那份 .json 演回来（需求第 7 条的"下次打开时恢复"）。

        只在**列表还是空的**时候演——启动那一刻它必然是空的，所以这一条不是"悄悄换掉人
        正在用的列表"，而人回来时第一眼就看见标题上写着是哪一份。盘上已经没有那一格路径
        就念一句，不去猜"是不是换地方了"。
        """
        stored = self._store.text(KEY_PROJECT, "").strip()
        if not stored or self.files.count():
            return
        path = Path(stored)
        if not path.is_file():
            self.log("上次那份工程文件已经不在这个位置了：%s（用「文件 → 打开工程…」另挑一份）。" % path,
                     "bad")
            return
        self.open_project(path)

    def _on_list_changed(self) -> None:
        """列表动过 → 星号 + 「开始转换」的悬浮说明重算。"""
        self._dirty = True
        self._set_title()
        self._describe_convert()

    def _set_title(self) -> None:
        """标题这一行是"这份列表从哪儿来、动过没有"唯一一处一直看得见的位置。"""
        if self._project is None:
            self.setWindowTitle("%s%s" % (APP_TITLE, "｜未保存的工程 *" if self._dirty else ""))
            return
        self.setWindowTitle("%s — %s%s" % (APP_TITLE, self._project.name, " *" if self._dirty else ""))

    def _remember_project(self, path: Path) -> None:
        """两本账一起记：``project/last_path`` 是启动那一拍要恢复的那一条，
        ``recent/projects`` 是菜单里那一列（需求第 1 条点名的"最近打开的文件列表"）。"""
        self._store.put(KEY_PROJECT, str(path))
        self._store.put(KEY_RECENT_PROJECTS,
                        push_recent(self._store.texts(KEY_RECENT_PROJECTS), [str(path)]))
        self._store.sync()
        self._rebuild_recent_menu()

    def project_path(self) -> Optional[Path]:
        return self._project

    def is_dirty(self) -> bool:
        return self._dirty

    def window_title(self) -> str:
        return self.windowTitle()

    # ---- 最近打开的工程与一份配置抄本（八期）---------------------------------

    def recent_projects(self) -> list[str]:
        """菜单里那一列现在记着哪几条（规范化后的路径，最新在前）。测试与「清空」的回执读这一份。"""
        return self._store.texts(KEY_RECENT_PROJECTS)

    def _rebuild_recent_menu(self) -> None:
        """重摆那一列：**盘上已经不在了的那几条当场剪掉**并写回本机设置。

        剪的理由不是省事：点到一条不存在的 .json 只会听到一句"读不懂"，而那一列存在的全部用处
        就是"上一次那份在哪儿"。剪完必须写回去，否则每次开机都念同一句剪除。
        """
        stored = self._store.texts(KEY_RECENT_PROJECTS)
        kept = [text for text in stored if Path(text).is_file()]
        dropped = [text for text in stored if text not in kept]
        if dropped:
            self._store.put(KEY_RECENT_PROJECTS, kept)
            self._store.sync()
            self.log("「最近打开的工程」里有 %d 条已经不在这个位置了，已从那张表里去掉：%s。" % (
                len(dropped), "、".join(dropped[:3]) + ("…" if len(dropped) > 3 else "")), "warn")
        self.recent_menu.clear()
        if not kept:
            # 空着会被读成"这一栏坏了"，所以给一句"还没有"而不是摆一张空子菜单。
            empty = self.recent_menu.addAction("（还没有打开过工程文件）")
            empty.setEnabled(False)
            return
        for text in kept:
            action = self.recent_menu.addAction(text)
            action.triggered.connect(lambda checked=True, value=text: self.open_project(Path(value)))
        self.recent_menu.addSeparator()
        clear = self.recent_menu.addAction("清空这张表")
        clear.triggered.connect(self.clear_recent_projects)

    def clear_recent_projects(self) -> None:
        """只是这一列不记了：**那几份 .json 盘上一格没动**（同「移除」与「清空」在列表上的分工）。"""
        self._store.put(KEY_RECENT_PROJECTS, [])
        self._store.sync()
        self._rebuild_recent_menu()
        self.set_status("「最近打开的工程」清空了（那几份文件还在盘上，只是这里不记了）。")

    def export_config(self) -> None:
        """把右边那一屏抄成一份 .json（需求第 4 条的一半）。"""
        target = self._ask_config_path("导出配置文件", QFileDialog.getSaveFileName, suggest_name=True)
        if target is None:
            self.set_status("没挑地方，这份配置还没抄出去。")
            return
        try:
            written = write_config(target, self.settings.style_values(),
                                   self.settings.custom_presets())
        except OSError as exc:
            self.set_status("配置文件没写出去：%s" % exc)
            self.log("配置文件没写出去：%s（右边那一屏一个字没动）" % exc, "bad")
            return
        custom = self.settings.custom_presets()
        self.log("配置抄本写好了：%s｜%s｜自定义预设 %d 条。%s" % (
            written, self.settings.preset_summary(), len(custom), CARRY_NOTE))
        self.set_status("已导出配置 %s。" % Path(written).name)

    def import_config(self, path: Optional[Path] = None) -> bool:
        """读一份抄本换掉右边那一屏（需求第 4 条的另一半）：**先读通，读通了才动屏幕**。

        与 :meth:`open_project` 同一次序：反过来（先清屏再报错）就把人正在用的那套排版弄没了，
        而屏幕上只剩一句"读不懂"。自定义预设**整包换掉**而不是逐格合并——一份抄本说的就是
        "这一套预设"，逐格合并永远删不掉一档。
        """
        target = Path(path) if path is not None else self._ask_config_path(
            "导入配置文件", QFileDialog.getOpenFileName)
        if target is None:
            self.set_status("没挑配置文件，右边那一屏照原样。")
            return False
        try:
            values, custom = read_config(target)
        except SettingsError as exc:
            # 句子本身已经点名是哪一份文件、哪一格读不懂，前面不再加"这份配置文件读不懂"——
            # 那会变成"这份配置文件读不懂：这份配置文件不是合法的 JSON"。
            self.set_status("导入没走通：%s" % exc.user_message)
            self.log("导入没走通：%s（右边那一屏一个字没动）" % exc.user_message, "bad")
            return False
        self.settings.apply_style_config(values, custom)
        self.log("配置抄本读进来了：%s｜%s｜自定义预设 %d 条。%s" % (
            target, self.settings.preset_summary(), len(custom), CARRY_NOTE))
        self.set_status("已导入配置 %s。" % target.name)
        return True

    def _ask_config_path(self, title: str, getter, suggest_name: bool = False) -> Optional[Path]:
        """配置那两张文件框，与 :meth:`_ask_project_path` 同一写法（脚本里绕开系统对话框）。

        ``suggest_name`` 只在**写**的那一路给：读的时候替人挑一个文件名等于凭空多一条决定。
        起点用输出目录旁边（人往往刚在那儿取过 .md），``output/directory`` 本身不跟着抄本走。
        """
        start = self._last_folder()
        if suggest_name:
            start = str(Path(start) / CONFIG_FILE_NAME)
        chosen, _filter = getter(self, title, start, config_filter())
        return Path(chosen) if chosen else None

    # ---- 预览 ---------------------------------------------------------------

    def show_preview_for(self, path: Optional[Path], force: bool = False) -> None:
        if path is None:
            self.preview.show_welcome()
            self.status_label.setText(IDLE_STATUS)
            return
        if not force and str(path) == self._previewed:
            return
        result = preview_html.render_file(Path(path))
        self._previewed = str(path)
        if not result.ok:
            self.preview.show_problem(result.problem)
            self.set_status("预览读不到 %s。" % Path(path).name)
            self.log("预览读不到：%s（%s）" % (path, result.problem), "bad")
            return
        self.preview.show_html(result.html)
        stats = result.stats
        self.set_status("预览 %s｜块 %s｜公式 %s｜表格 %s%s" % (
            Path(path).name, stats.get("blocks", 0),
            stats.get("formulas_inline", 0) + stats.get("formulas_display", 0),
            stats.get("tables", 0),
            "｜%d 处提醒" % len(result.warnings) if result.warnings else ""))

    # ---- 跑批 ---------------------------------------------------------------

    def _start(self, action: str, paths: Sequence[Path], options: Options) -> None:
        paths = list(paths)
        if self._worker is not None:
            self.set_status("这一批还在走，先按 Esc 或「取消这一批」。")
            return
        if not paths:
            self.set_status("左边一个文件都没勾上——勾上复选框再点这一个动作。")
            self.log("没有勾选的文件，什么都没做（复选框代表「参与这一批」）。")
            return
        if options.grouped and not self.files.has_groups():
            # 出不了任何东西的真实原因是"没有分组线"，不是"没勾文件"；只念后者会让人去翻复选框。
            self.set_status("「分组输出」得先有分组线：在左边列表空白处右键「新建分组…」，"
                            "或者把模式换成「单独输出」。")
            return
        verb = "转换" if action == ACTION_CONVERT else "校验"
        self.log("%s这一批：%d 个文件｜%s" % (verb, len(paths), options.summary()))
        if action == ACTION_CONVERT:
            plan = options.output_plan(paths)
            self.log("这批打算写到 %d 份：%s" % (
                len(plan), "；".join("%s → %s" % (label, target.name) for label, target, _ in plan)))
            empties = self.files.empty_groups()
            if empties:
                # 分组输出时"某一组一个都没勾"必须当场念：只报"出了 N 份"会让人以为每组都有东西。
                self.log("这一组一个文件都没勾上，所以它不出文件：%s。" % "、".join(empties), "warn")
        worker = ConvertWorker(action, paths, options, self)
        worker.log.connect(self.log)
        worker.logged.connect(self.log)
        worker.progress.connect(self._on_progress)
        worker.finished.connect(self._on_finish)
        worker.error.connect(self._on_error)
        self._worker = worker
        self._running = True
        self._set_running(True)
        # 进度条本来就是 0-100（``setFormat("%p%")``），执行体交回的也是百分比，这里不再换算步数。
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        worker.start()

    def _on_progress(self, percent: int, message: str) -> None:
        self.progress.setValue(max(0, min(int(percent), 100)))
        self.set_status(message)

    def _on_finish(self, stats: dict) -> None:
        """跑完了（或取消了）：收忙态、把回执念在状态栏、弹报告。

        ``stats`` 是 :class:`gui.tasks._Ledger` 交回的那一份纯数据，主窗口一个数都不补。
        """
        stats = stats or {}
        self._retire_worker()
        receipt = str(stats.get("receipt") or "")
        self.set_status(receipt or "这一批走完了。")
        self._report_dialog().set_stats(stats)
        self._report_dialog().show_report()

    def _on_error(self, message: str) -> None:
        """这一批**根本没跑成**（样式表找不到 / 引擎里出了没预料到的异常）。

        与 ``finished`` 分开是因为语义不同：一批里有个文件读不进照样算跑完、问题列在报告里，
        而这一路一份都没出。忙态必须在这里同样收掉，否则出一次错之后按钮永远灰着。
        """
        self._retire_worker()
        self.progress.setValue(0)
        self.log("这一批没跑成：%s" % message, "bad")
        self.set_status("这一批没跑成：%s" % message)

    def _retire_worker(self) -> None:
        """线程壳的收尾：先 ``wait()`` 确认真的退出，再 ``deleteLater()``。

        ``worker.finished.connect(worker.deleteLater)`` 那一行常见写法在这里不成立——
        ``finished`` 已经是带 dict 的那一个信号了（见模块开头那条）。
        """
        worker = self._worker
        self._worker = None
        self._running = False
        self._set_running(False)
        if worker is None:
            return
        if not worker.wait(3000):
            # 3 秒还没退：不 deleteLater，否则对象被销毁时线程还在跑是 Qt 明确不许的那一种崩。
            self.log("后台那一步 3 秒内没停下来，先不管它的回收（已经写出去的文件不回滚）。")
            return
        worker.setParent(None)
        worker.deleteLater()

    def _report_dialog(self) -> ReportDialog:
        if self._report is None:
            self._report = ReportDialog(self)
        return self._report

    def _set_running(self, running: bool) -> None:
        """跑批期间只留「取消这一批」能点：中途改列表会让日志里的文件名与实际写出的对不上。"""
        for name in TOOLBAR_OBJECT_NAMES + MENU_ACTION_OBJECT_NAMES:
            getattr(self, name).setEnabled(not running)
        # 最近那一列点一下就换一份清单，与「打开工程」是同一件事，所以跟着一起灰。
        self.recent_menu.setEnabled(not running)
        self.act_cancel.setEnabled(running)
        self.files.setEnabled(not running)
        self.settings.setEnabled(not running)

    # ---- 状态与日志 ----------------------------------------------------------

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)

    def log(self, line: str, kind: str = "") -> None:
        """日志区那一行。``kind`` 是那一行的**档位**（``bad`` / ``warn`` / ``ok`` / ``info``），
        空着就是正文色——需求第 3 条"错误项高亮"落在这里：颜色跟着这一行走，不另开一栏。

        ``ConvertWorker.log`` 与 ``ConvertWorker.logged`` 两根都连到这一个方法：前者是四期
        定下的形状（只有一个参数，档位为空），后者多带一档。执行体一次只走其中一根，
        所以同一句话不会被念两遍。
        """
        self.preview.log(line, kind)

    def status_text(self) -> str:
        return self.status_label.text()

    def log_text(self) -> str:
        return self.preview.log_text()

    def _on_counts(self, total: int, checked: int) -> None:
        if self._running:
            return
        if total == 0:
            self.set_status("列表是空的：点「添加文件」或把 .md 直接拖进这个窗口。")
        elif checked == 0:
            self.set_status("列表有 %d 个，但一个都没勾上——「开始转换」不做任何事。" % total)
        else:
            self.set_status("列表 %d 个，这一批参与 %d 个。" % (total, checked))

    def _describe_convert(self) -> None:
        """把「开始转换」的悬浮说明跟着面板取值与左边分组换，点之前就知道要出什么。"""
        self.act_convert.setToolTip("把勾上的那批写成 .docx｜" + self._options().summary())

    def show_about(self) -> None:
        QMessageBox.information(
            self, "关于这一工具",
            "Markdown 转 Word 桌面工具\n\n"
            "左栏挑文件、中间看近似预览和日志、右边定输出与样式。\n"
            "公式转 Word 原生可编辑对象这一步与命令行走的是同一套引擎；"
            "「开始转换」在后台线程里逐个文件读写，界面不冻，走完弹一张报告。")

    # ---- 主题 ---------------------------------------------------------------

    def apply_theme(self, name: str, announce: bool = True) -> None:
        self._theme = name if name in theme.theme_names() else theme.DEFAULT_THEME
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(theme.window_qss(self._theme))
        self.preview.set_theme(self._theme)
        self.files.apply_theme(self._theme)
        for key, action in self.theme_actions.items():
            action.setChecked(key == self._theme)
        if announce:
            self.log("界面主题换成了「%s」。" % theme.theme_label(self._theme))

    def theme_name(self) -> str:
        return self._theme

    # ---- 拖拽 ---------------------------------------------------------------

    def dragEnterEvent(self, event) -> None:  # noqa: D102
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:  # noqa: D102
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event) -> None:  # noqa: D102
        if not event.mimeData().hasUrls():
            super().dropEvent(event)
            return
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
        if paths:
            self.files.add_paths(paths)
            event.acceptProposedAction()
        else:
            self.set_status("拖进来的不是本机文件路径，没收。")

    # ---- QSettings ----------------------------------------------------------

    def _restore_layout(self) -> None:
        store = self._store
        geometry = store.blob(KEY_GEOMETRY)
        if geometry is not None:
            self.restoreGeometry(geometry)
        state = store.blob(KEY_WINDOW_STATE)
        if state is not None:
            self.restoreState(state)
        self.splitter.setSizes(store.sizes(KEY_SPLITTER, [280, 620, 300]))
        self.preview.restore_splitter(store.sizes(KEY_PREVIEW_SPLITTER, [560, 200]))
        self._keep_on_screen()

    def _save_layout(self) -> None:
        store = self._store
        store.put(KEY_GEOMETRY, self.saveGeometry())
        store.put(KEY_WINDOW_STATE, self.saveState())
        store.put_sizes(KEY_SPLITTER, self.splitter.sizes())
        store.put_sizes(KEY_PREVIEW_SPLITTER, self.preview.splitter_sizes())
        store.put(KEY_THEME, self._theme)
        self.settings.save_to_store()
        store.sync()

    def _keep_on_screen(self) -> None:
        """存的窗口位置在另一台显示器上时，接回来的窗口会整个落在屏幕外——看不见不等于没开。"""
        box = self.frameGeometry()
        for screen in QApplication.screens():
            if screen.availableGeometry().intersects(box):
                return
        primary = QApplication.primaryScreen()
        if primary is None:
            return
        area = primary.availableGeometry()
        self.move(area.x() + 40, area.y() + 40)

    def _last_folder(self) -> str:
        """文件框的起点：先用输出目录旁边（人往往刚在里面取过 .md），再退回项目根的 samples。"""
        chosen = self.settings.output_dir.text().strip()
        for candidate in (Path(chosen) if chosen else None, Path(default_output_dir()),
                          Path(default_output_dir()).parent / "samples"):
            if candidate is not None and candidate.is_dir():
                return str(candidate)
        return str(Path.cwd())

    def splitter_sizes(self) -> list[int]:
        return list(self.splitter.sizes())

    def restore_splitter(self, sizes: Iterable[int]) -> None:
        values = [int(size) for size in sizes]
        if len(values) == 3 and all(value > 0 for value in values):
            self.splitter.setSizes(values)

    def closeEvent(self, event) -> None:  # noqa: D102
        worker = self._worker
        if worker is not None and worker.isRunning():
            worker.requestInterruption()
            if not worker.wait(5000):
                self.log("后台那一步 5 秒内没停下来，窗口先关（已经写出去的文件不回滚）。")
        # 报告那一扇是窗口级模态：主窗口关了它得跟着关，否则人关掉一扇"挡着主窗口"的窗之后
        # 屏幕上留一张没有主人的报告，而它自己已经没有可回去的地方。
        if self._report is not None:
            self._report.close()
        self._save_last_project()
        self._save_layout()
        super().closeEvent(event)

    def _save_last_project(self) -> None:
        """关窗时那份列表值不值得替人写回去，只看一件事：**落点是这个人自己定过的**。

        定过（打开过或存过）→ 顺手写回那一份，"我挪过的顺序没存"这种事不该靠人记着 Ctrl+S；
        没定过 → **不弹窗也不凭空找个地方写**：弹框会把 ``--check`` 与冒烟测试卡在事件循环里，
        而"退出"那颗按钮点了就该退出；凭空挑一个目录写一份 .json 更难查——人找不到它、
        又以为存过了。那颗星号就是这一路的提醒，它一直挂在标题上。
        """
        if not self._dirty or self._project is None:
            return
        self._write_project(self._project)
