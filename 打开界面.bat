@echo off
rem md2docx 双击启动：只做三件事 —— 进项目目录、找到本项目自己的 python、敲 gui_app.py。
rem 界面与转换的逻辑全在 Python 那一侧（gui_app.py -> gui/app.py），这一栏不替它做任何决定。
title md2docx 转换界面（这一栏先别关：关了界面就跟着退出）
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto :no_venv
".venv\Scripts\python.exe" gui_app.py
if errorlevel 1 goto :failed
exit /b 0

:no_venv
echo [没能启动] 这一栏所在的文件夹里没有项目自己的虚拟环境，查的是这一格：
echo            "%~dp0.venv\Scripts\python.exe"
echo 这份 .bat 要和整个项目放在一起 —— 单独把它拷到桌面或别处就会念这一句。
echo 想在桌面上留一颗图标：右键这份 .bat → 发送到 → 桌面快捷方式（快捷方式照样找得到项目）。
echo 确实是头一回用、环境还没建，那就在项目目录里敲这两行（敲完再双击）：
echo            python -m venv .venv
echo            .venv\Scripts\pip install -r requirements.txt
goto :end

:failed
echo.
echo [界面没起来] 上面那几行就是原因。多数是四个依赖没装进这个 .venv，补一句就好：.venv\Scripts\pip install -r requirements.txt

:end
pause
exit /b 1
