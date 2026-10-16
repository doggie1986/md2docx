# -*- coding: utf-8 -*-
"""Markdown 转 Word 桌面工具的界面层。

这一层只做"人看见的那一屏"：颜色、控件、拖拽、进度条。
**转换与校验都不在这里发生**——读稿子、转公式、写 .docx 那三步连同后台线程壳都集中在
:mod:`gui.tasks`，面板只负责取值与画交回来的那份统计。

预览区是个例外，它确实要读文件：但只借 :mod:`md_parser` 把块结构摆成 HTML
（见 :mod:`gui.preview_html`），既不算公式也不写 .docx，和命令行那条链共用同一份解析口径。
"""

from __future__ import annotations

__all__ = ["__version__"]

__version__ = "0.1.0"
