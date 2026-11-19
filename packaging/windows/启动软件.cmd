@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\pythonw.exe" if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m pip show mddm-document-merger >nul 2>&1
    if not errorlevel 1 goto start_gui
)

echo 正在准备运行环境...
where py >nul 2>&1
if errorlevel 1 goto try_python
py -3.12-64 -m venv .venv >nul 2>&1
if not errorlevel 1 goto install

:try_python
where python >nul 2>&1
if errorlevel 1 goto missing_python
python -c "import sys,struct; sys.exit(not (sys.version_info[:2] == (3,12) and struct.calcsize('P') == 8))" >nul 2>&1
if errorlevel 1 goto missing_python
python -m venv .venv

:install
if errorlevel 1 goto failed
if not exist ".venv\Scripts\python.exe" goto failed
".venv\Scripts\python.exe" -m pip install --no-index --find-links "%~dp0wheelhouse" mddm-document-merger==1.0.0
if errorlevel 1 goto failed

:start_gui
start "" /D "%~dp0" ".venv\Scripts\pythonw.exe" -m mddm.merge_gui
exit /b 0

:missing_python
echo 未找到 64 位 Python 3.12。请先安装 Python 3.12 x64 并启用 Tcl/Tk 和 pip。
pause
exit /b 1

:failed
echo 安装失败。请检查以上错误信息。
pause
exit /b 1
