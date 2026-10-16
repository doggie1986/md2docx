# -*- coding: utf-8 -*-
"""两套 QSS 主题（浅色 / 深色）、预览区的 HTML 样式，以及列表里那些小图标。

颜色值只出现在这一份文件里：面板代码只说"这一格是提示语"（``QLabel#hint``）、
"这一行是个组头"（:func:`group_colors`）、"这一行的文件已经不在了"
（:func:`status_icon`），不自己拼色号，否则切主题时一定有几处忘了改。

预览区是 HTML，QSS 管不到它的正文，所以 ``preview_css()`` 单独给一份——
两套颜色的口径仍然在这两个字典里，不在别处。

图标是 ``QPainter`` 现画的，不用外部 ``.png``：一是那样一来每加一档状态就要多一个文件，
二是资源图片不带主题色，切深色时那一枚绿色对勾会贴在深底上糊成一团。
"""

from __future__ import annotations

from typing import Dict, Tuple

from PySide6.QtCore import QRect, Qt
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap

__all__ = ["THEME_LIGHT", "THEME_DARK", "THEME_LABELS", "DEFAULT_THEME",
           "STATUS_READY", "STATUS_MISSING", "STATUS_UNREADABLE", "STATUS_UNKNOWN",
           "STATUS_LABELS", "STATUS_ICON_EDGE", "theme_names", "theme_label", "window_qss",
           "preview_css", "color", "group_colors", "status_icon", "status_label"]

THEME_LIGHT = "light"
THEME_DARK = "dark"
DEFAULT_THEME = THEME_LIGHT

THEME_LABELS: Dict[str, str] = {THEME_LIGHT: "浅色", THEME_DARK: "深色"}

#: 列表里每一行右边那枚小圆点的四种档。``STATUS_UNKNOWN`` 是"还没核过"，画**空**图标——
#: 演一枚绿勾再说"没核过"，比什么都不画更容易让人以为文件已经检查过了。
STATUS_READY = "ready"
STATUS_MISSING = "missing"
STATUS_UNREADABLE = "unreadable"
STATUS_UNKNOWN = "unknown"

STATUS_LABELS: Dict[str, str] = {
    STATUS_READY: "可用",
    STATUS_MISSING: "已经不在了",
    STATUS_UNREADABLE: "读不动",
    STATUS_UNKNOWN: "还没核过",
}

#: 图标边长（像素）。列表那边要按同一个数开 ``setIconSize``，否则行高会跟着图片的实际尺寸跳。
STATUS_ICON_EDGE = 16

# 浅色：Qt 自带的 Fusion 偏灰，这里把边框和选中色压深一点，中文小字才不糊。
_LIGHT: Dict[str, str] = {
    "window": "#f4f6f9",
    "base": "#ffffff",
    "text": "#1f2328",
    "muted": "#5c6672",
    "border": "#ccd3db",
    "accent": "#1f6feb",
    "hover": "#eaeef3",
    "selected": "#d3e4ff",
    "selected_text": "#0b2f66",
    "toolbar": "#eceff3",
    "status": "#eceff3",
    "code_bg": "#f2f4f7",
    "code_fg": "#24292f",
    "math_bg": "#fff5e0",
    "math_fg": "#8a5000",
    "math_border": "#f0d8a8",
    "warn_bg": "#ffe9e9",
    "warn_fg": "#a4262c",
    "group_bg": "#e2e9f1",
    "group_fg": "#24313f",
    "ok": "#2da44a",
    "bad": "#cf222e",
}

# 深色：正文不用纯白（#ffffff 在深底上会发晕），强调色比浅色亮一档。
_DARK: Dict[str, str] = {
    "window": "#1c2128",
    "base": "#161b22",
    "text": "#dfe5ec",
    "muted": "#8d98a5",
    "border": "#39424e",
    "accent": "#58a6ff",
    "hover": "#232b34",
    "selected": "#1c3a5e",
    "selected_text": "#dfe5ec",
    "toolbar": "#202730",
    "status": "#202730",
    "code_bg": "#12171d",
    "code_fg": "#dfe5ec",
    "math_bg": "#33291a",
    "math_fg": "#f0c674",
    "math_border": "#5c4a24",
    "warn_bg": "#3a1f22",
    "warn_fg": "#ffb4b4",
    "group_bg": "#242c37",
    "group_fg": "#d8e6fb",
    "ok": "#3fb950",
    "bad": "#ff8a8a",
}

_TOKENS: Dict[str, Dict[str, str]] = {THEME_LIGHT: _LIGHT, THEME_DARK: _DARK}

_WINDOW_QSS = """
QWidget {
    background: %(window)s;
    color: %(text)s;
    font-size: 13px;
}
QMainWindow::separator { background: %(border)s; width: 1px; height: 1px; }

QToolBar {
    background: %(toolbar)s;
    border-bottom: 1px solid %(border)s;
    spacing: 3px;
    padding: 4px;
}
QToolBar::separator { background: %(border)s; width: 1px; margin: 4px 6px; }
QToolButton {
    background: transparent;
    border: 1px solid transparent;
    border-radius: 4px;
    padding: 4px 7px;
}
QToolButton:hover { background: %(hover)s; border-color: %(border)s; }
QToolButton:pressed { background: %(selected)s; }
QToolButton:disabled { color: %(muted)s; }

QMenuBar { background: %(toolbar)s; border-bottom: 1px solid %(border)s; }
QMenuBar::item:selected { background: %(selected)s; color: %(selected_text)s; }
QMenu { background: %(base)s; border: 1px solid %(border)s; }
QMenu::item { padding: 4px 18px; }
QMenu::item:selected { background: %(selected)s; color: %(selected_text)s; }
QMenu::separator { height: 1px; background: %(border)s; margin: 4px 6px; }

QStatusBar { background: %(status)s; border-top: 1px solid %(border)s; }
QStatusBar QLabel { color: %(text)s; }

QSplitter::handle { background: %(border)s; }
QSplitter::handle:horizontal { width: 3px; }
QSplitter::handle:vertical { height: 3px; }
QSplitter::handle:hover { background: %(accent)s; }

QGroupBox {
    border: 1px solid %(border)s;
    border-radius: 5px;
    margin-top: 9px;
    padding: 6px 8px 8px 8px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px;
    color: %(muted)s;
}
QGroupBox:disabled { color: %(muted)s; }
QGroupBox:title:disabled { color: %(muted)s; }

QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QAbstractSpinBox {
    background: %(base)s;
    border: 1px solid %(border)s;
    border-radius: 4px;
    padding: 3px 6px;
    selection-background-color: %(selected)s;
    selection-color: %(selected_text)s;
}
QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus { border-color: %(accent)s; }
QLineEdit:disabled, QComboBox:disabled, QDoubleSpinBox:disabled { background: %(hover)s; color: %(muted)s; }
QComboBox QAbstractItemView {
    background: %(base)s;
    border: 1px solid %(border)s;
    selection-background-color: %(selected)s;
    selection-color: %(selected_text)s;
}
QSpinBox::up-button, QDoubleSpinBox::up-button,
QSpinBox::down-button, QDoubleSpinBox::down-button { width: 14px; background: %(hover)s; }

QPushButton {
    background: %(hover)s;
    border: 1px solid %(border)s;
    border-radius: 4px;
    padding: 4px 10px;
}
QPushButton:hover { border-color: %(accent)s; }
QPushButton:pressed { background: %(selected)s; }
QPushButton:disabled { color: %(muted)s; }

QCheckBox, QRadioButton { spacing: 6px; background: transparent; }
QCheckBox::indicator, QRadioButton::indicator {
    width: 14px; height: 14px;
    border: 1px solid %(border)s;
    background: %(base)s;
}
QCheckBox::indicator { border-radius: 3px; }
QRadioButton::indicator { border-radius: 7px; }
QCheckBox::indicator:checked, QRadioButton::indicator:checked {
    background: %(accent)s;
    border-color: %(accent)s;
}
QCheckBox::indicator:checked:enabled { image: none; }
QCheckBox:disabled, QRadioButton:disabled { color: %(muted)s; }

QListWidget, QTreeWidget, QListView, QTreeView {
    background: %(base)s;
    border: 1px solid %(border)s;
    border-radius: 4px;
    padding: 2px;
    outline: 0;
}
/* 行那一格刻意**不写** border-radius：一旦给 ::item 加了圆角，Qt 就改成按样式表自己画背景，
   item.setBackground() 那一格行底色（分组线靠它）就整个不画了——实测同一条 QSS 去掉圆角，
   组头立刻演回 group_bg，而 padding 留着没影响。选中与悬停那两档是样式表自己给的底色，照旧生效。 */
QListWidget::item, QTreeView::item { padding: 3px 4px; }
QListWidget::item:hover, QTreeView::item:hover { background: %(hover)s; }
QListWidget::item:selected, QTreeView::item:selected {
    background: %(selected)s;
    color: %(selected_text)s;
}

/* 报告对话框那两张表（逐份输出 / 问题清单）。QTableView 不在上面那条 QListWidget 的规则里，
   不写就跟着 Qt 默认样式走——深色主题下会演出一块白底黑框，而表头是浅灰，一眼就看出漏了这一档。 */
QTableView {
    background: %(base)s;
    border: 1px solid %(border)s;
    gridline-color: %(border)s;
    outline: 0;
    selection-background-color: %(selected)s;
    selection-color: %(selected_text)s;
}
QTableView::item { padding: 2px 5px; }
QTableView::item:selected { background: %(selected)s; color: %(selected_text)s; }
QHeaderView { background: %(toolbar)s; }
QHeaderView::section {
    background: %(toolbar)s;
    color: %(text)s;
    padding: 3px 6px;
    border: 0;
    border-right: 1px solid %(border)s;
    border-bottom: 1px solid %(border)s;
    font-weight: bold;
}
QTableCornerButton::section { background: %(toolbar)s; border: 0; border-bottom: 1px solid %(border)s; }

QTextBrowser, QPlainTextEdit {
    background: %(base)s;
    border: 1px solid %(border)s;
    border-radius: 4px;
}
QPlainTextEdit#previewLog { font-family: Consolas, "Courier New", monospace; }

QScrollBar:vertical { background: %(window)s; width: 11px; margin: 0; }
QScrollBar::handle:vertical { background: %(border)s; min-height: 22px; border-radius: 5px; }
QScrollBar::handle:vertical:hover { background: %(muted)s; }
QScrollBar:horizontal { background: %(window)s; height: 11px; margin: 0; }
QScrollBar::handle:horizontal { background: %(border)s; min-width: 22px; border-radius: 5px; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }

QProgressBar {
    background: %(base)s;
    border: 1px solid %(border)s;
    border-radius: 4px;
    min-width: 160px;
    max-height: 15px;
    text-align: center;
    color: %(text)s;
}
QProgressBar::chunk { background: %(accent)s; border-radius: 3px; }

QToolTip { background: %(base)s; color: %(text)s; border: 1px solid %(border)s; }

QLabel#panelTitle { font-weight: bold; color: %(text)s; }
QLabel#hint { color: %(muted)s; }
QLabel#problem { color: %(warn_fg)s; font-weight: bold; }
"""

_PREVIEW_CSS = """
html, body { color: %(text)s; }
body { font-family: "宋体", SimSun, serif; font-size: 12pt; margin: 0; }
h1, h2, h3, h4, h5, h6 { font-family: "黑体", SimHei, sans-serif; margin: 10px 0 4px 0; }
a { color: %(accent)s; }
p { margin: 4px 0; }
table { border-collapse: collapse; margin: 6px 0; }
th, td { border: 1px solid %(border)s; padding: 3px 7px; }
th { background: %(hover)s; }
code, pre { font-family: Consolas, "Courier New", monospace; background: %(code_bg)s; color: %(code_fg)s; }
pre { padding: 6px 8px; border: 1px solid %(border)s; white-space: pre-wrap; }
blockquote { margin: 6px 0; padding-left: 8px; border-left: 3px solid %(border)s; color: %(muted)s; }
hr { border: none; border-top: 1px solid %(border)s; }
.math {
    background: %(math_bg)s;
    color: %(math_fg)s;
    border: 1px solid %(math_border)s;
    padding: 0 3px;
    font-family: "Cambria Math", "Times New Roman", serif;
}
.math-block { display: block; text-align: center; padding: 5px 0; }
/* 公式里的分式与矩阵是 :mod:`gui.math_html` 摆的两行/多行表格。上面那条 ``th, td`` 是给**人读的
   正文表格**准备的（带格线），公式这两张要的只有"分子下面那一条分数线"，所以在这里按类名盖回去。
   实测（把这一屏画进 QImage 逐像素扫）：盖了之后分式那张竖线 0 条、矩阵 0 条，摘掉这句就长出
   2 与 3 条（矩阵中间那条正是两列之间的分隔线），所以 Qt 的富文本真认 ``border: none``。
   分数线那条要写成 ``.math-frac .math-num``：只写 ``.math-num`` 的权重（0-1-0）比
   ``.math-frac td``（0-1-1）低，会被上面那句 ``border: none`` 一起盖掉——线就没了。 */
.math-frac, .math-matrix { border-collapse: collapse; margin: 0; }
.math-frac td, .math-matrix td { border: none; padding: 0 5px; text-align: center; }
.math-frac .math-num { border-bottom: 1px solid %(math_border)s; }
/* 语法体检没过的那条公式：退回 LaTeX 原文并涂警告色，而不是画成一个看着像公式的空格。 */
.math-broken { background: %(warn_bg)s; color: %(warn_fg)s; border: 1px solid %(warn_fg)s; }
.notice { color: %(muted)s; font-size: 10pt; border: 1px dashed %(border)s; padding: 4px 8px; }
.placeholder { color: %(warn_fg)s; background: %(warn_bg)s; }
"""


def theme_names() -> list[str]:
    """主题标识符的先后次序（菜单里就按这个顺序排）。"""
    return [THEME_LIGHT, THEME_DARK]


def theme_label(name: str) -> str:
    return THEME_LABELS.get(name, name)


def _tokens(name: str) -> Dict[str, str]:
    return _TOKENS.get(name, _LIGHT)


def window_qss(name: str) -> str:
    """整个应用的样式表。``name`` 认不出来时按浅色走，不抛——切主题不该把界面弄没。"""
    return _WINDOW_QSS % _tokens(name)


def preview_css(name: str) -> str:
    """预览区那份 HTML 的样式表（QSS 到不了 QTextBrowser 正文，所以另给一份）。"""
    return _PREVIEW_CSS % _tokens(name)


def color(name: str, token: str) -> QColor:
    """按令牌名取一枚颜色。认不出来的令牌交回"这一主题的正文色"而不是抛——
    切主题不该把界面弄崩，而拼错了的令牌名在截图里一眼就看得见。"""
    tokens = _tokens(name)
    return QColor(tokens.get(token) or tokens["text"])


def group_colors(name: str) -> Tuple[QColor, QColor]:
    """组头那一行的底色与字色。走的是**行自己**的背景色而不是 QSS：
    ``QSS`` 的 ``QListWidget::item`` 管不到"某一行"，只有 ``item.setBackground`` 办得到。"""
    return color(name, "group_bg"), color(name, "group_fg")


def status_icon(kind: str, name: str) -> QIcon:
    """一枚 16px 的状态点。同一（档，主题）只画一次——列表几百行，逐行现画会拖慢重排。

    ``STATUS_UNKNOWN`` 交回**真空图标**（不是画一张透明图：透明的图在行里仍占一格宽度，
    看起来像"有个图标但没内容"，而这里的原话是"这一行还没核过"）。
    """
    if kind == STATUS_UNKNOWN:
        return QIcon()
    key = (kind, name if name in _TOKENS else DEFAULT_THEME)
    icon = _ICON_CACHE.get(key)
    if icon is None:
        icon = QIcon(_paint_status(*key))
        _ICON_CACHE[key] = icon
    return icon


def status_label(kind: str) -> str:
    return STATUS_LABELS.get(kind, kind)


def _paint_status(kind: str, name: str) -> QPixmap:
    edge = STATUS_ICON_EDGE
    image = QPixmap(edge, edge)
    image.fill(Qt.GlobalColor.transparent)
    if kind == STATUS_UNKNOWN:
        # 还没核过就什么都不画。演一枚绿勾再说"没核过"，比空白更容易让人以为已经查过了。
        return image
    tokens = _tokens(name)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    inset = 2
    rect = QRect(inset, inset, edge - inset * 2, edge - inset * 2)
    if kind == STATUS_READY:
        painter.setBrush(color(name, "ok"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(rect)
        pen = QPen(color(name, "base"))
        pen.setWidthF(1.9)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        painter.setPen(pen)
        painter.drawLine(5, 8, 7, 11)
        painter.drawLine(7, 11, 11, 5)
    elif kind == STATUS_MISSING:
        pen = QPen(color(name, "bad"))
        pen.setWidthF(1.6)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawEllipse(rect)
        painter.drawLine(5, 5, 11, 11)
    else:
        # 读不动（以及日后多出来、这里不认识的档）都按"有麻烦"演：实心 + 一个叹号。
        painter.setBrush(color(name, "bad"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(rect)
        font = QFont()
        font.setBold(True)
        font.setPixelSize(11)
        painter.setFont(font)
        painter.setPen(QPen(color(name, "base")))
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, "!")
    painter.end()
    return image


_ICON_CACHE: Dict[tuple, QIcon] = {}
