# -*- coding: utf-8 -*-
"""右侧设置面板：输出模式、输出目录、样式（含样式预设）、合并选项。

面板只负责"人填的那一屏"和"把填的收成 :class:`gui.tasks.Options`"，
**它不知道怎么写 .docx**——那两个动作的执行体在 :mod:`gui.tasks`。

默认值一律从 :mod:`docx_builder` 取（``BODY_FONT_DEFAULT`` 那一族），不在界面里另抄一份：
命令行和界面两个入口各写一套默认，就会出"同一个稿子两种排版"这种对不回头的账。

样式预设那一档（:mod:`config.settings`）在这一层只管三件事：
**挑一档就把上面三格换成它那一套**、**把现在这三格存成一档**、**撤掉自定义的那一档**。
下拉里亮着的是**由那三格推出来的**（第一格对上就演那一档的名字，对不上就不演），
所以"现在用的是哪一档"没有第二个存储处——存了就会出现"改了字号、QSettings 里还写着「技术文档」"。
代价只有一句：两条排版**完全一样**的预设屏幕上演的是同一条的名字（自定义的档优先于内置的档，
所以人自己存的那一档不会因为撞上一个内置名字就撤不掉）。

改动**实时**落进本机设置（:data:`SAVE_DEBOUNCE_MS` 那一拍收着连续敲的字，免得每敲一个字
就写一遍注册表），但**回填屏幕时不写**（``_loading`` 那一面旗）：程序自己改控件与人改控件
在 Qt 看来是同一个信号，不分开就变成"一开机就把默认值存成了用户的设置"。
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Sequence

from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDoubleSpinBox, QFileDialog, QGroupBox,
                               QHBoxLayout, QInputDialog, QLabel, QLineEdit, QPushButton,
                               QRadioButton, QVBoxLayout, QWidget)

import docx_builder as db
from config.settings import (SIZE_MAX, SIZE_MIN, SettingsError, StylePreset, StyleValues,
                             add_custom, decode_custom, encode_custom,
                             preset_names, presets, remove_custom)
from gui.preferences import (KEY_BODY_FONT, KEY_BODY_SIZE, KEY_HEADING_FONT, KEY_MERGE_COVER,
                             KEY_MERGE_PAGE_BREAK, KEY_MERGE_SOURCE_LABEL, KEY_MERGE_TOC,
                             KEY_OUTPUT_DIR, KEY_OUTPUT_MODE, KEY_STYLE_PRESETS, MODE_GROUPED,
                             MODE_MERGED, MODE_SEPARATE, Store)
from gui.tasks import Options

__all__ = ["SettingsPanel", "BODY_FONT_CHOICES", "HEADING_FONT_CHOICES", "SAVE_DEBOUNCE_MS",
           "default_output_dir"]

#: 停手多久之后把这一屏写进本机设置。人连敲三次字体名只会写一次，
#: 而关窗那一路 :meth:`gui.main_window.MainWindow._save_layout` 还会再存一遍，不会因为这一拍丢改动。
SAVE_DEBOUNCE_MS = 600

#: 候选是"常见的那几个"，不是白名单：下拉是**可编辑**的，人敲了别的字体名照样收，
#: 由 docx_builder 落进 .docx（Word 那边没有这个字体时自己会回落，这不是本工具该拦的事）。
BODY_FONT_CHOICES = [db.BODY_FONT_DEFAULT, "黑体", "仿宋", "楷体", "等线", "微软雅黑",
                     "Times New Roman", "Arial"]
HEADING_FONT_CHOICES = [db.HEADING_FONT_DEFAULT, "宋体", "微软雅黑", "等线", "Arial"]

#: 字号的范围住在 :mod:`config.settings`（样式预设要判同一件事），这里不抄第二份：
#: 两处各一套就会出现"预设里填得进、旋钮拧不到"那种对不上的档位。

#: 三档模式各自的说明。写在一张表里而不是散在 ``_sync_mode`` 的 if 里：
#: 换一档的措辞只改这一处，而"未分组那一档到底出不出文件"是人点下去之前必须看见的事实。
MODE_HINTS = {
    MODE_SEPARATE: "勾上的每个 .md 各写一个同名 .docx，顺序按列表走。",
    MODE_MERGED: "勾上的 .md 按列表顺序合成一份 .docx，下面的四个开关这时才起作用。",
    MODE_GROUPED: ("左边列表里的每一条分组线各出一份 .docx，文件名就是组名；"
                   "排在第一条分组线之前的那些文件会并成一份「未分组」。"
                   "左边还没有分组线时这一档出不了东西——先在列表里右键「新建分组…」。"),
}


def default_output_dir() -> Path:
    """项目根下的 ``out``（一期命令行那个输出目录），第一次开界面时演这一个。"""
    return Path(__file__).resolve().parent.parent / "out"


class SettingsPanel(QWidget):
    """右侧那一竖栏。``options()`` 交回一次跑批要用的一切；改动实时写进本机设置，关窗那一路再兜一遍。"""

    changed = Signal()

    def __init__(self, store: Optional[Store] = None, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("settingsPanel")
        self._store = store
        #: 自定义预设那几条（内置的三档不归这里，它们在 :mod:`config.settings` 里自己带着）。
        self._custom: List[StylePreset] = []
        #: 程序正在往控件里回填的那一拍：信号照发，但不写盘、也不重算下拉那一档。
        self._loading = False
        #: 那一格读不懂时的句子（本机存的自定义预设坏了）。启动那一路填上，由 :meth:`_sync_hint` 演出来。
        self._preset_problem = ""

        self.mode_separate = QRadioButton("单独输出（每个文件一份 .docx）", self)
        self.mode_merged = QRadioButton("合并输出（勾上的合成一份）", self)
        self.mode_grouped = QRadioButton("分组输出（左边每一组各出一份）", self)
        self.mode_hint = QLabel("", self)
        self.mode_hint.setObjectName("hint")
        self.mode_hint.setWordWrap(True)

        self.output_dir = QLineEdit(self)
        self.output_dir.setPlaceholderText("转换后的 .docx 放哪儿")
        self.output_dir.setClearButtonEnabled(True)
        self.browse_button = QPushButton("浏览…", self)

        self.body_font = QComboBox(self)
        self.body_font.setEditable(True)
        self.body_font.addItems(BODY_FONT_CHOICES)
        self.body_size = QDoubleSpinBox(self)
        self.body_size.setRange(SIZE_MIN, SIZE_MAX)
        self.body_size.setSingleStep(0.5)
        self.body_size.setDecimals(1)
        self.body_size.setSuffix(" pt")
        self.heading_font = QComboBox(self)
        self.heading_font.setEditable(True)
        self.heading_font.addItems(HEADING_FONT_CHOICES)

        # 样式预设那一行：下拉只能点不能敲（档名由「另存为预设…」那条路起，敲出来的名字点不出实体）。
        self.preset = QComboBox(self)
        self.preset.setEditable(False)
        self.save_preset_button = QPushButton("另存为预设…", self)
        self.save_preset_button.setToolTip("把上面那三格存成一档自定义预设")
        self.drop_preset_button = QPushButton("撤销这一档", self)
        self.drop_preset_button.setToolTip("撤掉下拉里亮着的那一档（内置的三档撤不掉）")
        self.preset_hint = QLabel("", self)
        self.preset_hint.setObjectName("hint")
        self.preset_hint.setWordWrap(True)

        self.merge_toc = QCheckBox("生成目录", self)
        self.merge_break = QCheckBox("插入分页符", self)
        self.merge_source = QCheckBox("标注源文件名", self)
        self.merge_cover = QCheckBox("生成封面页", self)

        self._build_layout()
        # 改动落盘收在一拍里：人连敲三个字体名只写一遍注册表，而关窗那一路还会再存一遍，丢不了。
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(SAVE_DEBOUNCE_MS)
        self._save_timer.timeout.connect(self._flush_store)
        self._wire()
        self._load_defaults()

    # ---- 对外：取值与回填 ---------------------------------------------------

    def options(self) -> Options:
        return Options(
            output_dir=Path(self.output_dir.text().strip() or str(default_output_dir())),
            merged=self.mode_merged.isChecked(),
            body_font=self.body_font.currentText().strip() or db.BODY_FONT_DEFAULT,
            body_size=float(self.body_size.value()),
            heading_font=self.heading_font.currentText().strip() or db.HEADING_FONT_DEFAULT,
            with_toc=self.merge_toc.isChecked(),
            page_break=self.merge_break.isChecked(),
            source_label=self.merge_source.isChecked(),
            cover_page=self.merge_cover.isChecked(),
            grouped=self.mode_grouped.isChecked(),
        )

    def apply_mode(self, merged: bool) -> None:
        """老的两档写法：``True`` 走合并、``False`` 走单独（会撤掉分组那一档）。"""
        self.apply_mode_name(MODE_MERGED if merged else MODE_SEPARATE)

    def apply_mode_name(self, mode: str) -> None:
        """按 :mod:`gui.preferences` 里那三个词切档；认不出来回落单独输出。"""
        box = {MODE_MERGED: self.mode_merged, MODE_GROUPED: self.mode_grouped,
               MODE_SEPARATE: self.mode_separate}.get(mode, self.mode_separate)
        box.setChecked(True)
        self._sync_mode()

    def current_mode(self) -> str:
        """那一颗单选亮着，就交回 :mod:`gui.preferences` 里对应的词（存进 QSettings 的就是这个）。"""
        if self.mode_grouped.isChecked():
            return MODE_GROUPED
        if self.mode_merged.isChecked():
            return MODE_MERGED
        return MODE_SEPARATE

    # ---- 对外：样式预设与一份配置文件 ---------------------------------------

    def style_values(self) -> StyleValues:
        """配置文件认的那七格（样式三格 + 合并四格）。

        输出目录、主题、窗口尺寸都不在这里——它们不跟着拷贝走，理由逐条写在
        :data:`config.settings.NOT_CARRIED`，界面这一侧不另解释一遍。
        """
        return StyleValues(
            body_font=self.body_font.currentText().strip() or db.BODY_FONT_DEFAULT,
            body_size=float(self.body_size.value()),
            heading_font=self.heading_font.currentText().strip() or db.HEADING_FONT_DEFAULT,
            with_toc=self.merge_toc.isChecked(),
            page_break=self.merge_break.isChecked(),
            source_label=self.merge_source.isChecked(),
            cover_page=self.merge_cover.isChecked(),
        )

    def custom_presets(self) -> List[StylePreset]:
        """人自己存的那几条。内置的三档不在这里（它们在 :mod:`config.settings` 里自己带着）。"""
        return list(self._custom)

    def apply_style_config(self, values: StyleValues,
                           custom: Sequence[StylePreset] = ()) -> None:
        """把一份配置文件读回来的那七格摆上屏幕，自定义预设**整包换掉**。

        回填那一拍不写盘（``_loading`` 那一面旗），但末尾立刻落一次盘：这一趟改的是设置本身，
        等那一拍就把"导入配置"变成"只在屏幕上看看，重开就没了"。
        """
        self._loading = True
        try:
            self._custom = [item for item in custom if not item.builtin]
            self.body_font.setCurrentText(values.body_font)
            self.body_size.setValue(values.body_size)
            self.heading_font.setCurrentText(values.heading_font)
            self.merge_toc.setChecked(values.with_toc)
            self.merge_break.setChecked(values.page_break)
            self.merge_source.setChecked(values.source_label)
            self.merge_cover.setChecked(values.cover_page)
            self._refresh_presets()
        finally:
            self._loading = False
        self.save_to_store()
        self._write_presets()
        self.changed.emit()

    # ---- 样式预设：点一档 / 存一档 / 撤一档 ---------------------------------

    def save_preset_as(self, name: str) -> bool:
        """把上面那三格存成一档自定义预设。名字空着、撞内置那三档都在 :func:`config.settings.add_custom`
        那一关被拒，句子画在预设那一行的下方——不弹框：这一屏盖在主窗口上，状态栏那句人看不见。"""
        current = self.style_values()
        try:
            self._custom = add_custom(self._custom, StylePreset(
                name=name, body_font=current.body_font, body_size=current.body_size,
                heading_font=current.heading_font))
        except SettingsError as exc:
            self._set_hint(exc.user_message, problem=True)
            return False
        self._refresh_presets()
        self.save_to_store()
        self._write_presets()
        self.changed.emit()
        return True

    def drop_preset(self, name: str) -> bool:
        """撤掉一档自定义预设。内置那三档撤不掉，也是那一关说的。"""
        try:
            self._custom = remove_custom(self._custom, name)
        except SettingsError as exc:
            self._set_hint(exc.user_message, problem=True)
            return False
        self._refresh_presets()
        self.save_to_store()
        self._write_presets()
        self.changed.emit()
        return True

    def prompt_save_preset(self) -> None:
        """「另存为预设…」：先问名字，再走 :meth:`save_preset_as`（判法只在那一关）。"""
        name = self.ask_preset_name("把上面那三格（%s）存成一档预设，给它起个名字："
                                    % self.preset_summary(), "")
        if name is not None:
            self.save_preset_as(name)

    def prompt_drop_preset(self) -> None:
        """「撤销这一档」：撤的是下拉里亮着的哪一档——按钮只在真有自定义档亮着时才亮。"""
        row = self._preset_row()
        if row is None:
            return
        self.drop_preset(row.name)

    def ask_preset_name(self, label: str, start: str) -> Optional[str]:
        """单独列成一个方法是给测试留的口（同 :meth:`gui.file_panel.FilePanel.ask_text`）：
        脚本里换一个假的而不去弹模态窗。取消交回 ``None``，与"名字空着"是两回事。"""
        text, accepted = QInputDialog.getText(self, "样式预设", label, QLineEdit.Normal, start)
        return text if accepted else None

    def preset_summary(self) -> str:
        """那一行把要存的三格念出来——人起名之前就该看见自己存进去的是什么排版。"""
        current = self.style_values()
        return StylePreset(name="", body_font=current.body_font, body_size=current.body_size,
                           heading_font=current.heading_font).describe()

    # ---- 本机设置（QSettings）------------------------------------------------

    def save_to_store(self) -> None:
        """那七格 + 模式 + 目录写进本机设置。

        **自定义预设不在这里**：那一格读不懂时（坏数据）一次普通的实时落盘就会把人存的那几条顶成空，
        而"读不懂"和"没了"是两件事——那几条只在真换档的那一刻写（:meth:`_write_presets`）。
        """
        if self._store is None:
            return
        store = self._store
        store.put(KEY_OUTPUT_MODE, self.current_mode())
        store.put(KEY_OUTPUT_DIR, self.output_dir.text())
        store.put(KEY_BODY_FONT, self.body_font.currentText())
        store.put(KEY_BODY_SIZE, float(self.body_size.value()))
        store.put(KEY_HEADING_FONT, self.heading_font.currentText())
        store.put(KEY_MERGE_TOC, self.merge_toc.isChecked())
        store.put(KEY_MERGE_PAGE_BREAK, self.merge_break.isChecked())
        store.put(KEY_MERGE_SOURCE_LABEL, self.merge_source.isChecked())
        store.put(KEY_MERGE_COVER, self.merge_cover.isChecked())

    def _write_presets(self) -> None:
        """自定义预设整包落盘（点「另存为预设…」「撤销这一档」和导入配置那三条路）。"""
        if self._store is None:
            return
        self._store.put(KEY_STYLE_PRESETS, encode_custom(self._custom))
        self._store.sync()

    def load_from_store(self) -> None:
        """回填上次存的取值。没有存过就用 ``docx_builder`` 的默认，不自己发明第三套默认。"""
        if self._store is None:
            return
        store = self._store
        self._loading = True
        problem = ""
        try:
            # 存的本来就是那三个词之一，按词切档；老档里存的 "separate"/"merged" 照样读得回来。
            self.apply_mode_name(store.text(KEY_OUTPUT_MODE, MODE_SEPARATE))
            self.output_dir.setText(store.text(KEY_OUTPUT_DIR, str(default_output_dir())))
            self.body_font.setCurrentText(store.text(KEY_BODY_FONT, db.BODY_FONT_DEFAULT))
            self.body_size.setValue(store.number(KEY_BODY_SIZE, db.BODY_SIZE_DEFAULT))
            self.heading_font.setCurrentText(store.text(KEY_HEADING_FONT, db.HEADING_FONT_DEFAULT))
            self.merge_toc.setChecked(store.flag(KEY_MERGE_TOC, True))
            self.merge_break.setChecked(store.flag(KEY_MERGE_PAGE_BREAK, True))
            self.merge_source.setChecked(store.flag(KEY_MERGE_SOURCE_LABEL, True))
            self.merge_cover.setChecked(store.flag(KEY_MERGE_COVER, False))
            try:
                self._custom = decode_custom(store.text(KEY_STYLE_PRESETS, ""))
            except SettingsError as exc:
                # 坏的那一条留着不动（不写盘就不会顶掉它），句子由 _sync_preset_choice 一直念着。
                self._custom = []
                problem = exc.user_message
            self._sync_mode()
            self._preset_problem = problem
            self._rebuild_preset_rows()
            self._sync_preset_choice()
        finally:
            self._loading = False
        # 回填那一拍不发 changed，所以这里补一次：外面那颗"转换"按钮的提示语读的是这一屏，
        # 而 _wire_panels() 跑在本方法之前——不补就演着默认档的提示，屏幕上是上次存的档。
        self.changed.emit()

    # ---- 内部 -------------------------------------------------------------

    def _build_layout(self) -> None:
        mode_box = QGroupBox("输出模式", self)
        mode_layout = QVBoxLayout(mode_box)
        mode_layout.addWidget(self.mode_separate)
        mode_layout.addWidget(self.mode_merged)
        mode_layout.addWidget(self.mode_grouped)
        mode_layout.addWidget(self.mode_hint)

        dir_box = QGroupBox("输出目录", self)
        dir_layout = QVBoxLayout(dir_box)
        row = QHBoxLayout()
        row.addWidget(self.output_dir, 1)
        row.addWidget(self.browse_button, 0)
        dir_layout.addLayout(row)
        dir_hint = QLabel("目录不存在会先建；留空按默认目录走。", dir_box)
        dir_hint.setObjectName("hint")
        dir_layout.addWidget(dir_hint)

        style_box = QGroupBox("样式设置", self)
        style_layout = QVBoxLayout(style_box)
        style_layout.addWidget(self._field_row("正文字体", self.body_font))
        style_layout.addWidget(self._field_row("字号", self.body_size))
        style_layout.addWidget(self._field_row("标题字体", self.heading_font))
        preset_row = QHBoxLayout()
        preset_row.addWidget(QLabel("预设", style_box), 0)
        preset_row.addWidget(self.preset, 1)
        preset_row.addWidget(self.save_preset_button, 0)
        preset_row.addWidget(self.drop_preset_button, 0)
        style_layout.addLayout(preset_row)
        style_layout.addWidget(self.preset_hint)
        style_hint = QLabel("这几格与命令行的 --body-font / --body-size / --heading-font 是同一件事。",
                            style_box)
        style_hint.setObjectName("hint")
        style_hint.setWordWrap(True)
        style_layout.addWidget(style_hint)

        merge_box = QGroupBox("合并选项", self)
        self.merge_box = merge_box
        merge_layout = QVBoxLayout(merge_box)
        merge_layout.addWidget(self.merge_toc)
        merge_layout.addWidget(self.merge_break)
        merge_layout.addWidget(self.merge_source)
        merge_layout.addWidget(self.merge_cover)
        merge_hint = QLabel("只有「合并输出」或「分组输出」时用得上——单独输出时这一栏是灰的。", merge_box)
        merge_hint.setObjectName("hint")
        merge_hint.setWordWrap(True)
        merge_layout.addWidget(merge_hint)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(10)
        layout.addWidget(mode_box)
        layout.addWidget(dir_box)
        layout.addWidget(style_box)
        layout.addWidget(merge_box)
        layout.addStretch(1)

    def _field_row(self, label_text: str, widget: QWidget) -> QWidget:
        """一栏一格的小行：左边名称、右边控件。放在一个 QWidget 里，样式和伸缩都好办。"""
        holder = QWidget(self)
        row = QHBoxLayout(holder)
        row.setContentsMargins(0, 0, 0, 0)
        label = QLabel(label_text, holder)
        label.setMinimumWidth(56)
        row.addWidget(label, 0)
        row.addWidget(widget, 1)
        return holder

    def _wire(self) -> None:
        # 每一格的信号名各写各的：QComboBox 没有 textChanged、QDoubleSpinBox 没有 textEdited，
        # 想"统一挑一个信号名"就会在某个控件上抛 AttributeError（而且是点下去才炸）。
        self.mode_separate.toggled.connect(self._sync_mode)
        self.mode_separate.toggled.connect(self._emit_changed)
        self.mode_merged.toggled.connect(self._emit_changed)
        self.mode_grouped.toggled.connect(self._emit_changed)
        self.browse_button.clicked.connect(self._browse)
        for box in (self.merge_toc, self.merge_break, self.merge_source, self.merge_cover):
            box.toggled.connect(self._emit_changed)
        self.output_dir.textChanged.connect(self._emit_changed)
        # 样式那三格走 _on_style_changed：改了排版就要重推"现在亮着哪一档"，
        # 下拉里亮着的是**推出来的**，不是另外存的一格（存了就会出"改了字号、那一行还写着技术文档"）。
        self.body_font.editTextChanged.connect(self._on_style_changed)
        self.heading_font.editTextChanged.connect(self._on_style_changed)
        self.body_size.valueChanged.connect(self._on_style_changed)
        # activated 而不是 currentIndexChanged：后者在程序回填那一档时也会发，
        # 于是"换字号 → 重算下拉 → 又应用一档"自己咬自己。activated 只由人点那一下发。
        # 带的是**行号**：Qt6 把 activated(QString) 撤了（只剩 activated(int) 那一族），
        # 想按名字接就先炸一个 IndexError——而且行号本来就能在下拉那一份表里查到档。
        self.preset.activated[int].connect(self._adopt_preset)
        self.save_preset_button.clicked.connect(self.prompt_save_preset)
        self.drop_preset_button.clicked.connect(self.prompt_drop_preset)

    def _load_defaults(self) -> None:
        """开机那一屏：``docx_builder`` 的默认档。回填这一拍**不写盘**（见 ``_loading``）。"""
        self._loading = True
        try:
            self.mode_separate.setChecked(True)
            self.output_dir.setText(str(default_output_dir()))
            self.body_font.setCurrentText(db.BODY_FONT_DEFAULT)
            self.body_size.setValue(db.BODY_SIZE_DEFAULT)
            self.heading_font.setCurrentText(db.HEADING_FONT_DEFAULT)
            self.merge_toc.setChecked(True)
            self.merge_break.setChecked(True)
            self.merge_source.setChecked(True)
            self.merge_cover.setChecked(False)
            self._sync_mode()
            self._rebuild_preset_rows()
            self._sync_preset_choice()
        finally:
            self._loading = False
        self.changed.emit()

    def _sync_mode(self, *_ignored) -> None:
        """下面那四个开关在"合成一份"和"每组一份"两档都用得上——它们管的是"一份文档里怎么摆"。"""
        merged = self.mode_merged.isChecked()
        grouped = self.mode_grouped.isChecked()
        self.merge_box.setEnabled(merged or grouped)
        self.mode_hint.setText(
            MODE_HINTS[MODE_GROUPED if grouped else (MODE_MERGED if merged else MODE_SEPARATE)])

    def _emit_changed(self, *_ignored) -> None:
        """这一屏变了：通知外面（转换按钮那颗的提示语），并把落盘推到 ``SAVE_DEBOUNCE_MS`` 之后。

        回填那一拍（``_loading``）里信号照发但不算"人改了"：不重排下拉那一档、也不写盘，
        否则"一开机就把默认值存成了用户的设置"，而且刚点中的一档会被自己算成没点中。
        """
        if self._loading:
            return
        self._save_timer.start()
        self.changed.emit()

    def _flush_store(self) -> None:
        """debounce 那一拍到了：写本机设置并**落盘**。不 sync 就还是"关窗时才存"，实时那句白写。"""
        self.save_to_store()
        if self._store is not None:
            self._store.sync()

    def _refresh_presets(self) -> None:
        """那几条自定义预设换了：重画下拉，并撤掉"读不懂"那一句——坏数据已被这一趟写回来了。"""
        self._preset_problem = ""
        self._rebuild_preset_rows()
        self._sync_preset_choice()

    def _on_style_changed(self, *_ignored) -> None:
        self._sync_preset_choice()
        self._emit_changed()

    def _rebuild_preset_rows(self) -> None:
        """下拉里的次序 = :func:`config.settings.presets` 的次序（内置三档在前，自定义按存的次序）。"""
        self.preset.clear()
        self.preset.addItems(preset_names(self._custom))

    def _preset_row(self) -> Optional[StylePreset]:
        """亮着的那一档；没亮着（这三格跟哪一档都不完全一样）交回 ``None``。"""
        index = self.preset.currentIndex()
        if index < 0:
            return None
        return presets(self._custom)[index]

    def _sync_preset_choice(self) -> None:
        """由那三格推出该演哪一档，对不上就整格空着（不塞一条假的「（无预设）」）。"""
        index = self._match_index()
        # 不 blockSignals：这里改的是"演哪一行"，而 activated 只由人点那一下发（setCurrentIndex 不发它）。
        self.preset.setCurrentIndex(index)
        row = self._preset_row()
        self.drop_preset_button.setEnabled(row is not None and not row.builtin)
        if self._preset_problem:
            # 本机那一条读不懂时这一句一直挂着：改了字号也不该把"你存的那几档没了"演成一句用法提示。
            self._set_hint(self._preset_problem, problem=True)
            return
        if row is None:
            self._set_hint("上面那三格（%s）和哪一档都不完全一样——点「另存为预设…」把这一套存成一档。"
                           % self.preset_summary())
            return
        self._set_hint("%s：%s%s" % (row.name, row.describe(),
                                     "（内置的三档之一，撤不掉）" if row.builtin else "（自定义的档）"))

    def _match_index(self) -> int:
        """由那三格推出该演哪一档；**自定义的档优先于**排版一模一样的内置档。

        这一条次序不是美学：「撤销这一档」撤的是下拉里亮着的那一条，而内置那三档撤不掉——
        人自己存的那档恰好与「技术文档」一模一样时，演内置的名字就等于把这一档变成撤不掉的档。
        """
        current = self.style_values()
        builtin_hit = -1
        for index, item in enumerate(presets(self._custom)):
            if not (item.body_font == current.body_font
                    and float(item.body_size) == current.body_size
                    and item.heading_font == current.heading_font):
                continue
            if not item.builtin:
                return index
            if builtin_hit < 0:
                builtin_hit = index
        return builtin_hit

    def _adopt_preset(self, index: int) -> None:
        """点一档就把上面那三格换成它那一套。换完那三格自己会把下拉推回这一档（同一套判法）。"""
        rows = presets(self._custom)
        if not 0 <= index < len(rows):
            # 刚重建过下拉而人点的是旧的那一行（毫秒级的 race，真人手速碰不到，脚本会）：
            # 重算一次让屏幕上回到真相，不假装应用了一档已经不存在的排版。
            self._sync_preset_choice()
            return
        row = rows[index]
        self._loading = True
        try:
            self.body_font.setCurrentText(row.body_font)
            self.body_size.setValue(row.body_size)
            self.heading_font.setCurrentText(row.heading_font)
        finally:
            self._loading = False
        self._sync_preset_choice()
        self._emit_changed()

    def _set_hint(self, text: str, problem: bool = False) -> None:
        """预设那一行下面的说明。坏了那一句得红：它说的是"你存的那几档这一版读不回来"，不是用法提示。"""
        self._preset_problem = text if problem else ""
        name = "problem" if problem else "hint"
        if self.preset_hint.objectName() != name:
            self.preset_hint.setObjectName(name)
            # 换了 objectName 要让样式重算一遍，否则 QLabel#problem 那条规则不会 apply（主题里早就有这一条）。
            style = self.preset_hint.style()
            style.unpolish(self.preset_hint)
            style.polish(self.preset_hint)
        self.preset_hint.setText(text)

    def _browse(self) -> None:
        start = self.output_dir.text().strip() or str(default_output_dir())
        chosen = QFileDialog.getExistingDirectory(self, "选择输出目录", start)
        if chosen:
            self.output_dir.setText(str(Path(chosen)))
