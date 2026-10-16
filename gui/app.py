# -*- coding: utf-8 -*-
"""建 ``QApplication``、按 QSettings 里存的主题起手、把主窗口摆出来。

顺序在这一层是**有关系的**，不是随便排：

1. ``setOrganizationName`` / ``setApplicationName`` 必须在 ``QApplication`` **之前**——
   ``QSettings()`` 不带参数时读的就是这两个名字，晚了第一格的设置就写到另一个位置，
   人会以为"上次调的三栏宽度没存住"。
2. ``setStyle("Fusion")`` 在 :meth:`QApplication.setStyleSheet` 之前：Windows 的原生风格
   不认 QSS 里的 ``border-radius`` 与大部分 ``::sub-control``，切深色时会出"半屏变了、
   半屏还是系统那套灰"。Fusion 是 Qt 自带的，跨机器长得一样。
3. 主题先按库里存的演一次，再建窗口——窗口建好之后才贴 QSS 的话，第一帧是系统色，
   人会看见一闪。
"""

from __future__ import annotations

import sys
from typing import Optional, Sequence

from PySide6.QtWidgets import QApplication

from gui import theme
from gui.main_window import APP_TITLE, MainWindow
from gui.preferences import APP_NAME, KEY_THEME, ORG_NAME, Store

__all__ = ["main", "build_application"]


def build_application(argv: Optional[Sequence[str]] = None) -> QApplication:
    """``QApplication`` 起手那两件：先定 QSettings 的名字，再换 Fusion。测试也走这一条。"""
    QApplication.setOrganizationName(ORG_NAME)
    QApplication.setApplicationName(APP_NAME)
    QApplication.setApplicationDisplayName(APP_TITLE)
    app = QApplication(list(argv) if argv is not None else sys.argv[:1])
    app.setStyle("Fusion")
    return app


def main(argv: Optional[Sequence[str]] = None) -> int:
    """起界面。退出码 0 = 正常关窗，1 = 连窗口都建不起来（话写在 stderr）。"""
    try:
        app = build_application(argv)
    except Exception as exc:  # 没装 PySide6 / 缺 DLL 时要一句人话，不是甩一屏 traceback
        sys.stderr.write("起不来界面：%s（%s）\n先确认依赖装在项目 .venv 里："
                         "pip install -r requirements.txt\n" % (exc, type(exc).__name__))
        return 1

    store = Store()
    name = store.text(KEY_THEME, theme.DEFAULT_THEME).strip().lower()
    if name not in theme.theme_names():
        name = theme.DEFAULT_THEME
    app.setStyleSheet(theme.window_qss(name))

    window = MainWindow(store=store)
    window.show()
    return app.exec()
