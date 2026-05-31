@echo off
chcp 65001 >nul
setlocal
cd /d %~dp0

:: --- 設定 ---
set "EXE_NAME=MSL集計ソフト"
set "SCRIPT_NAME=MSLdata_check.py"
set "ICON_NAME=logo.ico"
set "NUITKA_CACHE_DIR=%~dp0nuitka_tmp"

echo === STEP 1: CLEANING ===
if exist "%EXE_NAME%.exe" del /f /q "%EXE_NAME%.exe"
if exist "%SCRIPT_NAME%.build" rmdir /s /q "%SCRIPT_NAME%.build"
if exist "%SCRIPT_NAME%.onefile-build" rmdir /s /q "%SCRIPT_NAME%.onefile-build"

echo === STEP 2: NUITKA OPTIMIZED BUILD ===
echo 依存関係の最適化を実施中...

:: コマンドを1行にまとめてエラーを回避します
:: --jobs は整数（4など）を指定します
python -m nuitka --onefile --standalone --enable-plugin=tk-inter --windows-console-mode=disable --windows-icon-from-ico="%ICON_NAME%" --lto=yes --jobs=6 --mingw64 --assume-yes-for-downloads --output-filename="%EXE_NAME%" --remove-output --no-deployment-flag=self-execution --nofollow-import-to=unittest --nofollow-import-to=pydoc --nofollow-import-to=IPython --nofollow-import-to=notebook --nofollow-import-to=numpy.random --nofollow-import-to=matplotlib --nofollow-import-to=PIL "%SCRIPT_NAME%"

echo.
echo === STEP 3: RESULT ===
if exist "%EXE_NAME%.exe" (
    echo [SUCCESS] ビルドが完了しました。
    echo 保存先: %CD%\%EXE_NAME%.exe
) else (
    echo [ERROR] ビルドに失敗しました。
)
pause