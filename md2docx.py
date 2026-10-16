# -*- coding: utf-8 -*-
"""``python md2docx.py input.md output.docx`` —— 命令行入口的那颗文件名。

规格里给的就是这一行，所以项目根放这一个薄壳，真正的内容在 :mod:`main`：
参数、退出码、回执都在那儿，这里只负责"用哪个文件名敲"。

为什么是两个文件而不是把 ``main.py`` 改名：

* ``main.py`` 是**能被 import 的模块**（测试与脚本 ``from main import run`` 走它），
  ``md2docx.py`` 是**给人敲的那颗名字**——两者职责不同但代码只有一份。
* Python 把"脚本所在目录"放进 ``sys.path``，所以从任何工作目录敲
  ``python E:\\...\\md2docx.py`` 都能找到同目录下的 ``main`` 与那三个模块。
"""

from __future__ import annotations

import sys

from main import main

if __name__ == "__main__":
    sys.exit(main())
