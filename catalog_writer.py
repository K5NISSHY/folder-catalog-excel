# -*- coding: utf-8 -*-
"""Excel / CSV への書き出し（GUI 非依存）。

- openpyxl の `write_only=True` でストリーム的に書く（NFR-3）
- 各シートは見出しにフィルタ・先頭行固定・列幅調整を施す（FR-7）
- 既存ファイルは上書きしない。同名があれば自動で連番を付ける（FR-8 / NFR-5）
- CSV は UTF-8（BOM 付き）。シート「一覧」と「フォルダ集計」の 2 種類（FR-16）
"""

from __future__ import annotations

import csv
import os
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from scan_core import (
    ERROR_COLUMNS,
    EXT_COLUMNS,
    FILE_COLUMNS,
    FOLDER_COLUMNS,
    SUMMARY_COLUMNS,
    ScanResult,
)

# 出力ファイル名の連結子（FR-8）
NAME_SUFFIX = "目録"

# シート名（FR-4〜FR-6 / FR-10 / 2.1 の構成）
SHEET_SUMMARY = "サマリー"
SHEET_LIST = "一覧"
SHEET_FOLDERS = "フォルダ集計"
SHEET_EXTS = "拡張子集計"
SHEET_ERRORS = "エラー一覧"

SHEET_TITLES = (SHEET_SUMMARY, SHEET_LIST, SHEET_FOLDERS, SHEET_EXTS, SHEET_ERRORS)

# CSV として出すシート（FR-16）
CSV_SHEETS = (SHEET_LIST, SHEET_FOLDERS)

# CSV のエンコーディング（Excel で開いても文字化けしない BOM 付き UTF-8）
CSV_ENCODING = "utf-8-sig"

_HEADER_FILL = PatternFill("solid", fgColor="305496")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center")

# シートごとの列幅（FR-7）
_COLUMN_WIDTHS = {
    SHEET_SUMMARY: (26, 78),
    SHEET_LIST: (6, 62, 10, 30, 40, 10, 12, 20, 20),
    SHEET_FOLDERS: (6, 62, 10, 14, 18, 14, 18, 14),
    SHEET_EXTS: (6, 14, 10, 16, 20, 62),
    SHEET_ERRORS: (6, 16, 62, 62),
}

# 数値として #,##0 表示にする列（0 始まりの列番号）
_NUMBER_COLUMNS = {
    SHEET_FOLDERS: (3, 4, 5, 6, 7),
    SHEET_EXTS: (2, 3, 4),
}

_INVALID_NAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# ---------------------------------------------------------------------------
# ファイル名（FR-8）
# ---------------------------------------------------------------------------


def sanitize_name(name: str) -> str:
    """フォルダ名をファイル名として安全な形に直す（不正文字は '_' に置換）。"""
    cleaned = _INVALID_NAME_CHARS.sub("_", str(name)).strip().rstrip(". ")
    return cleaned or "対象フォルダ"


def build_base_name(folder_name: str, when: datetime = None) -> str:
    """`<フォルダ名>_目録_<YYYYMMDD_HHMM>` を作る（FR-8）。"""
    when = when or datetime.now()
    return f"{sanitize_name(folder_name)}_{NAME_SUFFIX}_{when:%Y%m%d_%H%M}"


def unique_output_paths(out_dir: str, base_name: str, with_csv: bool = False) -> tuple:
    """既存ファイルと衝突しない出力パスの組を返す（FR-8 / NFR-5）。

    戻り値: (xlsx パス, {シート名: CSV パス})。同名があれば ` (2)`, ` (3)` ... と連番を付ける。
    Excel と CSV で同じ連番を使い、どれかが既にあっても上書きしない。
    """
    wanted = CSV_SHEETS if with_csv else ()
    number = 1
    while True:
        link = "" if number == 1 else f" ({number})"
        xlsx_path = os.path.join(out_dir, f"{base_name}{link}.xlsx")
        csv_paths = {sheet: os.path.join(out_dir, f"{base_name}{link}_{sheet}.csv") for sheet in wanted}
        taken = [p for p in (xlsx_path, *csv_paths.values()) if os.path.exists(p)]
        if not taken:
            return xlsx_path, csv_paths
        number += 1


# ---------------------------------------------------------------------------
# Excel 書き出し
# ---------------------------------------------------------------------------


def _header_cells(ws, titles):
    """見出しセル（太字・白抜き・中央揃え）を作る。"""
    cells = []
    for title in titles:
        cell = WriteOnlyCell(ws, value=title)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
        cells.append(cell)
    return cells


def _write_sheet(wb, title: str, header, rows) -> int:
    """1 シートを書く。見出しにフィルタ・先頭行固定・列幅調整を施す（FR-7）。"""
    ws = wb.create_sheet(title=title)
    ws.freeze_panes = "A2"
    for index, width in enumerate(_COLUMN_WIDTHS.get(title, ()), start=1):
        ws.column_dimensions[get_column_letter(index)].width = width

    number_columns = _NUMBER_COLUMNS.get(title, ())
    ws.append(_header_cells(ws, header))
    if number_columns:
        for row in rows:
            out = []
            for i, value in enumerate(row):
                if i in number_columns and isinstance(value, (int, float)):
                    cell = WriteOnlyCell(ws, value=value)
                    cell.number_format = "#,##0"
                    out.append(cell)
                else:
                    out.append(value)
            ws.append(out)
    else:
        for row in rows:
            ws.append(list(row))

    last_row = len(rows) + 1
    last_col = get_column_letter(len(header))
    ws.auto_filter.ref = f"A1:{last_col}{last_row}"
    return last_row


def write_catalog(result: ScanResult, xlsx_path: str) -> int:
    """目録を 1 つの Excel ワークブックに書き出す（FR-4〜FR-7 / FR-10）。

    既存ファイルがある場合は `FileExistsError`（上書きしないための防御）。
    戻り値は「一覧」シートに書いた行数（見出しを含む）。
    """
    if os.path.exists(xlsx_path):
        raise FileExistsError(f"既存ファイルは上書きしません: {xlsx_path}")

    wb = Workbook(write_only=True)
    _write_sheet(wb, SHEET_SUMMARY, SUMMARY_COLUMNS, result.summary_rows())
    list_rows = _write_sheet(wb, SHEET_LIST, FILE_COLUMNS, result.file_rows)
    _write_sheet(wb, SHEET_FOLDERS, FOLDER_COLUMNS, result.folder_rows)
    _write_sheet(wb, SHEET_EXTS, EXT_COLUMNS, result.ext_rows)
    _write_sheet(wb, SHEET_ERRORS, ERROR_COLUMNS, result.error_rows)
    wb.save(xlsx_path)
    return list_rows


# ---------------------------------------------------------------------------
# CSV 書き出し（FR-16）
# ---------------------------------------------------------------------------


def write_csv(path: str, header, rows) -> int:
    """UTF-8（BOM 付き）で CSV を書く。既存ファイルは上書きしない。"""
    if os.path.exists(path):
        raise FileExistsError(f"既存ファイルは上書きしません: {path}")
    with open(path, "w", encoding=CSV_ENCODING, newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(list(header))
        for row in rows:
            writer.writerow(list(row))
    return len(rows) + 1


def write_csv_catalog(result: ScanResult, csv_paths: dict) -> dict:
    """シート「一覧」と「フォルダ集計」を CSV に書き出す（FR-16）。

    戻り値: {シート名: 書き出した CSV のパス}
    """
    table = {
        SHEET_LIST: (FILE_COLUMNS, result.file_rows),
        SHEET_FOLDERS: (FOLDER_COLUMNS, result.folder_rows),
    }
    written = {}
    for sheet, path in csv_paths.items():
        header, rows = table[sheet]
        write_csv(path, header, rows)
        written[sheet] = path
    return written


def run_catalog(result: ScanResult, out_dir: str, with_csv: bool = False, when: datetime = None) -> dict:
    """出力パスを決めて Excel（必要なら CSV）を書き出す。GUI・検証スクリプト共通の入口。

    戻り値: {'xlsx': パス, 'csv': {シート名: パス}, 'list_rows': 一覧の行数}
    """
    xlsx_path, csv_paths = unique_output_paths(out_dir, build_base_name(result.root_name, when), with_csv)
    list_rows = write_catalog(result, xlsx_path)
    csv_written = write_csv_catalog(result, csv_paths) if csv_paths else {}
    return {"xlsx": xlsx_path, "csv": csv_written, "list_rows": list_rows}
