@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
python "%~dp0build_exe.py" %*
pause