# -*- coding: utf-8 -*-
"""中间那一栏：上半是预览，下半是日志，中间一根竖向 QSplitter。

日志区**只读**（``setReadOnly(True)``）：它是"这一批发生了什么"的凭据，
能编辑就迟早被人随手改掉，改完之后那句"转换过几个文件"不再可信。

预览区是 ``QTextBrowser``，样式走 :func:`gui.theme.preview_css`——QSS 管不到 HTML 正文，
所以切主题时这里要再贴一次 HTML（``self._html`` 留着就是为了这一句：
拿 ``toHtml()`` 回灌会被 QTextDocument 洗成另一份标记，越切越花）。
"""

from __future__ import annotations

import time
from html import escape
from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QTextCharFormat, QTextCursor
from PySide6.QtWidgets import QPlainTextEdit, QSplitter, QTextBrowser, QVBoxLayout, QWidget

from gui import theme

__all__ = ["PreviewPanel", "MAX_LOG_LINES", "WELCOME_HTML"]

#: 日志只往上叠，不封顶会在"加入整个文件夹 + 一路校验"之后吃掉一大堆内存。
MAX_LOG_LINES = 2000

#: 日志那一行的"档位"→ 主题里的颜色令牌。空串是普通一行（转换进度、回执），
#: 校验查出的问题按 ``bad`` / ``warn`` 上色——需求第 3 条"错误项高亮"走的就是这一格。
#: 认不出来的档位照普通一行走而不是抛：多一个档位名不该把日志区弄崩。
LOG_KIND_TOKENS = {
    "": "text",
    "info": "muted",
    "ok": "ok",
    "warn": "warn_fg",
    "bad": "bad",
    "error": "bad",
}

WELCOME_HTML = (
    "<div class=\"notice\">预览是近似的：公式按 Qt 富文本画得出来的那几样摆（上下标、上下两格的分式、"
    "√( … )），认不出的写法退回 LaTeX 原文并涂上警告色；转进 Word 之后它们都是可编辑的公式对象，"
    "字号与分页以右边那块面板里的样式设置为准。</div>"
    "<p>从左边选一个文件就能看近似预览。</p>"
    "<p>把 .md 或整个文件夹从资源管理器拖进这个窗口，列表会自己收下来。</p>"
    "<p>「开始转换」与「校验」的进度念在下面这块日志区，校验查出的问题按轻重上色。</p>"
)


class PreviewPanel(QWidget):
    """主窗口拿到的就是这一块：``show_html`` / ``show_problem`` / ``log`` 三个口。"""

    def __init__(self, theme_name: str = theme.DEFAULT_THEME, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("previewPanel")
        self._theme = theme_name
        self._html = WELCOME_HTML
        #: 日志的**唯一一份底账**：``(整行文本, 档位)``。上色与切主题重画都读它，
        #: 所以控件里那份"只有颜色没有文本"的状态坏了也能原样重建。
        self._lines: list[tuple[str, str]] = []

        self.browser = QTextBrowser(self)
        self.browser.setObjectName("previewBrowser")
        self.browser.setOpenExternalLinks(False)
        # 拖进来的东西只归左边的文件列表：这里不收，事件才会冒到主窗口转交给它。
        # QTextEdit 一家默认 acceptDrops=True（能拖文字进正文），预览区是只读的，
        # 收下只会演成"预览区被改了"，而人本来是想加文件。
        self.browser.setAcceptDrops(False)

        self.log_view = QPlainTextEdit(self)
        self.log_view.setObjectName("previewLog")
        self.log_view.setReadOnly(True)
        self.log_view.setAcceptDrops(False)
        self.log_view.setMaximumBlockCount(MAX_LOG_LINES)
        self.log_view.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)

        self.splitter = QSplitter(Qt.Orientation.Vertical, self)
        self.splitter.setObjectName("previewSplitter")
        self.splitter.addWidget(self.browser)
        self.splitter.addWidget(self.log_view)
        self.splitter.setCollapsible(0, False)
        self.splitter.setCollapsible(1, False)
        self.splitter.setSizes([560, 200])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.splitter)

        self.set_theme(theme_name)
        self.show_welcome()

    # ---- 预览 -------------------------------------------------------------

    def show_html(self, html: str) -> None:
        self._html = html
        self.browser.setHtml(html)
        self.browser.verticalScrollBar().setValue(0)

    def show_problem(self, problem: str) -> None:
        """读不动 / 解析不动时摆一句人话，不清空成白屏——白屏会被读成"这个文件是空的"。"""
        self.show_html("<p class=\"placeholder\">%s</p>" % escape(problem))

    def show_welcome(self) -> None:
        self.show_html(WELCOME_HTML)

    def current_html(self) -> str:
        return self._html

    def splitter_sizes(self) -> list[int]:
        return list(self.splitter.sizes())

    def restore_splitter(self, sizes: list[int]) -> None:
        if len(sizes) == 2 and all(int(size) > 0 for size in sizes):
            self.splitter.setSizes([int(size) for size in sizes])

    # ---- 日志 -------------------------------------------------------------

    def log(self, line: str, kind: str = "") -> None:
        """一行带时刻的日志；回头查"哪一步慢了"就靠这个时刻。

        ``kind`` 是档位（``""`` / ``info`` / ``ok`` / ``warn`` / ``bad``）：校验查出的问题
        要能在几十行进度里一眼看出来，所以颜色跟着**这一行**走，而不是另开一栏。
        """
        entry = "[%s] %s" % (time.strftime("%H:%M:%S"), line)
        self._lines.append((entry, kind or ""))
        if len(self._lines) > MAX_LOG_LINES:
            del self._lines[:len(self._lines) - MAX_LOG_LINES]
        self._append_line(entry, kind or "")
        bar = self.log_view.verticalScrollBar()
        bar.setValue(bar.maximum())

    def _append_line(self, entry: str, kind: str) -> None:
        """把一整行刷进文档。空文档不先插块——否则日志开头多一条空行。"""
        cursor = self.log_view.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        if not self.log_view.document().isEmpty():
            cursor.insertBlock()
        cursor.insertText(entry, self._format(kind))

    def _format(self, kind: str) -> QTextCharFormat:
        fmt = QTextCharFormat()
        fmt.setForeground(theme.color(self._theme, LOG_KIND_TOKENS.get(kind, "text")))
        return fmt

    def clear_log(self) -> None:
        self.log_view.clear()
        self._lines = []

    def log_text(self) -> str:
        return self.log_view.toPlainText()

    def log_lines(self) -> list[tuple[str, str]]:
        """底账那一列 ``((文本, 档位))``——测试与"另存日志"读这一条，不从控件里反推颜色。"""
        return list(self._lines)

    # ---- 主题 -------------------------------------------------------------

    def set_theme(self, name: str) -> None:
        """换预览那份 CSS、重贴一次 HTML，并把日志按新主题**重画上色**。

        日志的颜色是逐行刷进文档的（QSS 到不了 ``QTextCharFormat``），所以切主题必须重画，
        否则深色背景上还挂着深色字——"切了主题日志看不见"这种问题只在切换之后才现形。
        """
        self._theme = name
        self.browser.document().setDefaultStyleSheet(theme.preview_css(name))
        self.browser.setHtml(self._html)
        self._repaint_log()

    def _repaint_log(self) -> None:
        if not self._lines:
            return
        self.log_view.clear()
        cursor = self.log_view.textCursor()
        for index, (entry, kind) in enumerate(self._lines):
            if index:
                cursor.insertBlock()
            cursor.insertText(entry, self._format(kind))
        self.log_view.verticalScrollBar().setValue(self.log_view.verticalScrollBar().maximum())

    def theme_name(self) -> str:
        return self._theme
