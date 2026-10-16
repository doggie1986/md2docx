# -*- coding: utf-8 -*-
"""QSettings 读写：窗口大小与位置、三栏宽度、主题、设置面板的取值、上次那个工程文件的路径、
自定义样式预设、最近打开的工程都存在这里。

键名集中在这一份文件里（界面代码不写字符串字面量），否则改了键名就变成
"旧设置读不到、界面上看着像没存过"。

这一本账与**配置文件**（:mod:`config.settings` 那份 .json 抄本）是两回事：这里存的是"这台机器上的
偏好"，关机重开还要在；那份是"给人拷走的一套设置"，导入导出走的是文件框。两格形状不一样是因为
可拷贝的那一份必须读得懂，所以 :func:`config.settings.encode_custom` 之类的形状规则住在 Qt 摸不着
的地方，这里只负责把字符串递给它、再把读回来的字符串交回去。

``QSettings`` 存的位置取决于 org/app 名，那两个名字在 :func:`gui.app.main` 里
建 ``QApplication`` **之前**就定好了；这里留成参数只是为了测试能换一个临时的档。
"""

from __future__ import annotations

from typing import Any, Iterable, Optional, Sequence

from PySide6.QtCore import QByteArray, QSettings

__all__ = ["ORG_NAME", "APP_NAME", "Store",
           "KEY_GEOMETRY", "KEY_WINDOW_STATE", "KEY_SPLITTER", "KEY_PREVIEW_SPLITTER",
           "KEY_THEME", "KEY_OUTPUT_MODE", "KEY_OUTPUT_DIR", "KEY_PROJECT",
           "KEY_BODY_FONT", "KEY_BODY_SIZE", "KEY_HEADING_FONT",
           "KEY_MERGE_TOC", "KEY_MERGE_PAGE_BREAK", "KEY_MERGE_SOURCE_LABEL", "KEY_MERGE_COVER",
           "KEY_STYLE_PRESETS", "KEY_RECENT_PROJECTS"]

ORG_NAME = "md2docx"
APP_NAME = "md2docx-gui"

KEY_GEOMETRY = "window/geometry"
KEY_WINDOW_STATE = "window/state"
KEY_SPLITTER = "window/splitter_sizes"
KEY_PREVIEW_SPLITTER = "preview/splitter_sizes"

KEY_THEME = "appearance/theme"

KEY_OUTPUT_MODE = "output/mode"
KEY_OUTPUT_DIR = "output/directory"
KEY_BODY_FONT = "style/body_font"
KEY_BODY_SIZE = "style/body_size"
KEY_HEADING_FONT = "style/heading_font"
KEY_MERGE_TOC = "merge/with_toc"
KEY_MERGE_PAGE_BREAK = "merge/page_break"
KEY_MERGE_SOURCE_LABEL = "merge/source_label"
KEY_MERGE_COVER = "merge/cover_page"

#: 自定义的样式预设存的是**一段文本**（:func:`config.settings.encode_custom` 写、
#: :func:`config.settings.decode_custom` 读）。为什么不像下面那一条那样存成真列表：
#: 那几条预设是有形状的（名字 + 三格），形状规则必须在没有 Qt 的地方也能判——
#: 存成文本就让 :mod:`tests.test_settings` 那一份 Qt-free 的测试能把"读不懂就抛"当场钉住，
#: 而不是只在界面上炸。内置那三档**不在这里**（它们在这一版代码里，改了档就没有对照物）。
KEY_STYLE_PRESETS = "style/custom_presets"

#: 最近打开的工程文件，**一串路径**（QSettings 直接存 list[str]，见 :meth:`Store.texts`）。
#: 记的是工程 .json 而不是 .md 稿子：那列稿子清单本来就存在工程文件里，
#: 再单独记一遍就会出现"菜单里说最近打开过 a.md，而 a.md 属于哪一份清单没人知道"。
KEY_RECENT_PROJECTS = "recent/projects"

#: 上一次打开或存下的那个工程文件的绝对路径。两个用处：软件一启动照它把列表演回来
#: （启动那一刻列表必然是空的，所以这不算"悄悄换掉人正在用的列表"——标题与日志会当场念出是哪一份，
#: 见 gui.main_window.MainWindow._restore_last_project），以及「打开工程…」那个文件框的起点。
KEY_PROJECT = "project/last_path"

MODE_SEPARATE = "separate"
MODE_MERGED = "merged"
MODE_GROUPED = "grouped"
MODE_NAMES = (MODE_SEPARATE, MODE_MERGED, MODE_GROUPED)

_TRUE_WORDS = {"1", "true", "yes", "on"}


class Store:
    """一层薄封装：把 ``QSettings`` 读回来的各种怪形状（str 装的布尔、QVariant 装的列表）归一成 Python 值。"""

    def __init__(self, organization: str = ORG_NAME, application: str = APP_NAME,
                 settings: Optional[QSettings] = None) -> None:
        self._settings = settings if settings is not None else QSettings(organization, application)

    # ---- 读 ---------------------------------------------------------------

    def text(self, key: str, default: str = "") -> str:
        value = self._raw(key)
        return default if value is None else str(value)

    def flag(self, key: str, default: bool = False) -> bool:
        value = self._raw(key)
        if value is None:
            return default
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return bool(value)
        return str(value).strip().lower() in _TRUE_WORDS

    def number(self, key: str, default: float) -> float:
        try:
            return float(self._raw(key))
        except (TypeError, ValueError):
            return default

    def whole(self, key: str, default: int) -> int:
        try:
            return int(float(self._raw(key)))
        except (TypeError, ValueError):
            return default

    def sizes(self, key: str, default: Sequence[int]) -> list[int]:
        """三栏宽度那一类"一串整数"。读不到或形状不对就交回默认，不猜。"""
        value = self._raw(key)
        if value is None:
            return list(default)
        if isinstance(value, (str, bytes, QByteArray)):
            parts = str(value, "utf-8").split(",") if isinstance(value, (bytes, QByteArray)) else str(value).split(",")
        elif isinstance(value, Iterable):
            parts = list(value)
        else:
            return list(default)
        try:
            result = [int(float(part)) for part in parts]
        except (TypeError, ValueError):
            return list(default)
        return result or list(default)

    def texts(self, key: str, default: Sequence[str] = ()) -> list[str]:
        """一串文本（最近打开的那条表）。那一格没有值就交回默认，不猜。

        实测 QSettings 在 IniFormat 与注册表两个后端都能把 ``list[str]`` 原样存回取回
        （中文、带空格、带逗号都不串位），**一条**也照样是列表、**空列表**读回来是 ``[]``，
        所以这一格不需要一对 encode/decode——那是 :data:`KEY_STYLE_PRESETS` 的写法，
        那一条要判形状，这一条只判"是不是文本"。

        列表里混进不是文本的那一格就只扔掉那一格、其余照收（与 :meth:`sizes` 不同：那一串是
        三栏宽度，少一格就把整栏算歪；这一串是一行一条的表，扔掉坏的那行不伤别行）。
        """
        value = self._raw(key)
        if value is None:
            return list(default)
        if isinstance(value, str):
            return [value] if value else []
        if isinstance(value, Iterable):
            return [item for item in value if isinstance(item, str) and item]
        return list(default)

    def blob(self, key: str) -> Optional[QByteArray]:
        value = self._raw(key)
        if isinstance(value, QByteArray) and not value.isEmpty():
            return value
        return None

    # ---- 写 ---------------------------------------------------------------

    def put(self, key: str, value: Any) -> None:
        self._settings.setValue(key, value)

    def put_sizes(self, key: str, sizes: Iterable[int]) -> None:
        self._settings.setValue(key, [int(size) for size in sizes])

    def sync(self) -> None:
        self._settings.sync()

    def clear(self) -> None:
        self._settings.clear()

    # ---- 内部 -------------------------------------------------------------

    def _raw(self, key: str) -> Any:
        value = self._settings.value(key)
        if isinstance(value, QByteArray) and value.isEmpty():
            return None
        return value
