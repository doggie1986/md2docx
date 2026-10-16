# -*- coding: utf-8 -*-
"""桌面入口的那颗文件名 —— ``python gui_app.py``。

和 :file:`md2docx.py` 同一个写法：根目录放这颗薄壳，真正的内容在 :mod:`gui.app`。
为什么不能把 ``main.py`` 改名或复用：``main.py`` 是命令行那一条链（参数、退出码、回执），
这一颗要的是 ``QApplication`` 的事件循环，两者唯一的共同点是同一套转换引擎。

``QApplication`` 之前得先把 org/app 两个名字定下来，``QSettings`` 才知道往注册表里
哪一格读写；那一步住在 :func:`gui.app.main` 里，不在这里——薄壳只做"用哪个文件名敲"。
"""

from __future__ import annotations

import sys

from gui.app import main

if __name__ == "__main__":
    sys.exit(main())
