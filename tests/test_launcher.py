# -*- coding: utf-8 -*-
"""双击启动项 ``打开界面.bat`` 的自检。

这一层不是 Python，所以测不了"跑起来"——跑起来那一步是真按一次双击核对的
（截图与进程清单在临时目录，见说明书 §1）。这里钉的是**改坏它的那几种改法**：

* 有人把它另存成 UTF-8 或 LF —— Windows 控制台默认代码页是 GBK，那样中文提示会
  糊成一屏乱码，而 ``goto`` 的标签在 LF 文件上还可能认错，于是双击之后什么都不发生。
* 有人把解释器换成系统那个 python —— 四个依赖只装在 ``.venv`` 里，用系统的 python
  只会念一句 ``ModuleNotFoundError``，而这正是这份 .bat 存在的理由（它得替人挑对解释器）。
* 有人把"没找到环境"那一句改成"没找到就试系统的 python"—— 那是把一句人话换成
  一个看不见的原因，所以**回落这条路本身要钉住**。

判法全按**字节**读：先看这份文件在 GBK 控制台里念出来是什么，再按那一行行文本判。
"""

from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LAUNCHER = ROOT / "打开界面.bat"


def _text_of(raw: bytes) -> str:
    """按 GBK 读这份 .bat —— 控制台就是按这一套码页解它的。

    ``errors="strict"``：里面一旦出现 GBK 解不出的字节（多半是被另存成了 UTF-8），
    当场抛，让测试红在"编码不对"这一格上，而不是往下比一堆乱码。
    """
    return raw.decode("gbk")


class LauncherTest(unittest.TestCase):
    def test_the_launcher_lives_next_to_the_app(self) -> None:
        self.assertTrue(LAUNCHER.is_file(), "项目根得有 打开界面.bat（和 gui_app.py 同一层）")
        self.assertTrue((ROOT / "gui_app.py").is_file())

    def test_the_bytes_are_gbk_and_the_lines_end_with_crlf(self) -> None:
        raw = LAUNCHER.read_bytes()
        self.assertTrue(raw.startswith(b"@echo off\r\n"), "开头得是 @echo off 且用 CRLF 换行")
        self.assertNotIn(b"\n", raw.replace(b"\r\n", b""))
        self.assertNotIn("\ufffd", _text_of(raw))
        text = _text_of(raw)
        # 这两句是给人看的中文；GBK 解回来还是原话，才说明这份文件没被另存过编码
        self.assertIn("[没能启动]", text)
        self.assertIn("[界面没起来]", text)

    def test_it_runs_the_project_own_interpreter_and_never_a_fallback(self) -> None:
        text = _text_of(LAUNCHER.read_bytes())
        lines = [line.strip() for line in text.splitlines()]
        self.assertIn('".venv\\Scripts\\python.exe" gui_app.py', text)
        # 回落系统 python 就是"看不见的原因"：整份文件里不许出现不带 .venv 的那一条
        for line in lines:
            if line.startswith(("rem ", "@", ":")) or "echo" in line:
                continue
            self.assertNotIn("python gui_app.py", line)
            self.assertNotIn("pythonw", line)

    def test_it_checks_the_venv_before_launching_and_names_the_path(self) -> None:
        text = _text_of(LAUNCHER.read_bytes())
        lines = [line.strip() for line in text.splitlines()]
        check = next(i for i, line in enumerate(lines) if line.startswith("if not exist"))
        launch = next(i for i, line in enumerate(lines) if line.endswith("gui_app.py"))
        self.assertLess(check, launch, "先确认环境在不在，再谈敲不敲得动")
        self.assertIn('"%~dp0.venv\\Scripts\\python.exe"', text)
        self.assertIn('cd /d "%~dp0"', text)

    def test_a_clean_exit_closes_the_window_but_a_failure_keeps_it(self) -> None:
        lines = [line.strip() for line in _text_of(LAUNCHER.read_bytes()).splitlines()]
        launch = next(i for i, line in enumerate(lines) if line.endswith("gui_app.py"))
        self.assertEqual(lines[launch + 1], "if errorlevel 1 goto :failed")
        self.assertEqual(lines[launch + 2], "exit /b 0")
        self.assertEqual(lines[-2:], ["pause", "exit /b 1"])
        for label in (":no_venv", ":failed", ":end"):
            self.assertIn(label, lines)


if __name__ == "__main__":
    unittest.main()
