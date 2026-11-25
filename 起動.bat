@echo off
setlocal
cd /d "%~dp0"

rem -------------------------------------------------------------
rem  フォルダ目録 Excel 作成ツール 起動スクリプト
rem  .venv があればそれを使い、無ければ py -3 で起動する。
rem  ライブラリのインストール手順は要求しない
rem  （openpyxl が入った Python 環境を使う）。
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

py -3 "folder_catalog_excel.py"
if errorlevel 1 (
    echo.
    echo 起動に失敗しました。Python 3 と openpyxl が入っているか確認してください。
    echo （このフォルダに .venv がある場合は、そちらが優先されます）
    pause
)
