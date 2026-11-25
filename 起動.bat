@echo off
setlocal
cd /d "%~dp0"

rem -------------------------------------------------------------
rem  フォルダ目録 Excel 作成ツール 起動スクリプト
rem  .venv があればそれを使い、無ければ py -3 で起動する。
rem  openpyxl が未インストールのときは、初回起動時に
rem  pip install --user openpyxl を自動で実行する。
rem
rem  ※ このファイルは Shift-JIS (CP932) 保存。
rem     UTF-8 で保存すると cmd.exe が行頭を食い違って
rem     コメント行をコマンドとして実行しようとする。
rem -------------------------------------------------------------

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "folder_catalog_excel.py"
    goto :eof
)

if exist ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" "folder_catalog_excel.py"
    goto :eof
)

set "PYTHON_CMD="
where py >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    where python >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD (
    echo Python が見つかりませんでした。
    echo Python 3.10 以降を https://www.python.org/downloads/windows/ からインストールし、
    echo セットアップ時に "Add python.exe to PATH" にチェックを入れてください。
    pause
    exit /b 1
)

%PYTHON_CMD% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo openpyxl が入っていないので、初回インストールを実行します...
    %PYTHON_CMD% -m pip install --user openpyxl
    if errorlevel 1 (
        echo インストールに失敗しました。Python と pip、インターネット接続を確認してください。
        pause
        exit /b 1
    )
)

%PYTHON_CMD% "folder_catalog_excel.py"
if errorlevel 1 (
    echo.
    echo 起動に失敗しました。Python 3 と openpyxl が入っているか確認してください。
    echo （このフォルダに .venv がある場合は、そちらが優先されます）
    pause
)
endlocal
