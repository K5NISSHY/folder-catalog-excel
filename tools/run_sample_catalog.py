# -*- coding: utf-8 -*-
"""サンプルフォルダで実行走査＋Excel/CSV 生成し、**機械判定**を出す（検証用）。

`samples/` を既定の除外設定で走査し、`samples/検証結果/` に Excel と CSV を出す。
そのうえで次の項目を機械的に確かめる（判定出力は 13 行）:

1. 収集ファイル数と「一覧」シート行数が一致
2. シートが 5 枚（サマリー／一覧／フォルダ集計／拡張子集計／エラー一覧）揃う
3. 「一覧」に見出しフィルタ・先頭行固定・列幅調整が付いている
4. 合計サイズがサマリー・一覧・拡張子集計で食い違わない（独立走査の期待値とも一致）
5. 「エラー一覧」シートが見出し付きで存在する
6. CSV が UTF-8（BOM 付き）で、行数が Excel と一致する
7. 再実行すると連番付きの別ファイルになり、既存ファイルは上書きされない
8. 「フォルダ集計」の行数が収集フォルダ数＋1（対象フォルダ自身）と一致
9. 既定の除外（.git / node_modules / desktop.ini / Thumbs.db / 隠し）が効いている
10. 独立走査（os.walk）で数えた期待値と収集件数が一致

実行: `.venv\\Scripts\\python.exe tools\\run_sample_catalog.py`
"""

from __future__ import annotations

import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

from openpyxl import load_workbook  # noqa: E402

import catalog_writer  # noqa: E402
import scan_core  # noqa: E402
from scan_core import DEFAULT_EXCLUDE_NAMES, ScanOptions, bytes_to_kb  # noqa: E402

SAMPLE_ROOT = os.path.join(REPO_ROOT, "samples")
OUT_DIR = os.path.join(SAMPLE_ROOT, "検証結果")

# 前回の実行結果が次の走査に混ざらないよう、出力フォルダ自体を除外する
EXCLUDE_NAMES = tuple(DEFAULT_EXCLUDE_NAMES) + (os.path.basename(OUT_DIR),)

_ATTR_HIDDEN = 0x2
_ATTR_SYSTEM = 0x4


def _keep(name: str, st) -> bool:
    """scan_core と同じ除外規則を、別実装で再現する（突き合わせ用）。"""
    if name.lower() in {n.lower() for n in EXCLUDE_NAMES}:
        return False
    attrs = getattr(st, "st_file_attributes", 0)
    if (attrs & _ATTR_HIDDEN) or name.startswith("."):
        return False
    return not bool(attrs & _ATTR_SYSTEM)


def independent_count(root: str) -> tuple:
    """openpyxl・scan_core とは別の経路（os.walk）で期待値を数える。"""
    files = 0
    folders = 0
    total_bytes = 0
    per_file_kb = 0
    for dirpath, dirnames, filenames in os.walk(root):
        kept_dirs = []
        for name in dirnames:
            try:
                st = os.stat(os.path.join(dirpath, name))
            except OSError:
                continue
            if _keep(name, st):
                kept_dirs.append(name)
        dirnames[:] = kept_dirs
        folders += len(kept_dirs)
        for name in filenames:
            try:
                st = os.stat(os.path.join(dirpath, name))
            except OSError:
                continue
            if not _keep(name, st):
                continue
            files += 1
            total_bytes += st.st_size
            per_file_kb += bytes_to_kb(st.st_size)
    return files, folders, total_bytes, per_file_kb


def main() -> int:
    if not os.path.isdir(SAMPLE_ROOT):
        print("samples/ が見つかりません。先に tools/make_samples.py を実行してください。")
        return 2
    os.makedirs(OUT_DIR, exist_ok=True)

    options = ScanOptions(exclude_names=EXCLUDE_NAMES)
    print(f"走査対象: {SAMPLE_ROOT}")
    result = scan_core.scan(SAMPLE_ROOT, options)
    print(scan_core.preview_text(result))
    print()

    info = catalog_writer.run_catalog(result, OUT_DIR, with_csv=True)
    print(f"Excel: {info['xlsx']}")
    for sheet, path in info["csv"].items():
        print(f"CSV({sheet}): {path}")
    print()

    wb = load_workbook(info["xlsx"])
    ws_list = wb[catalog_writer.SHEET_LIST]
    ws_folders = wb[catalog_writer.SHEET_FOLDERS]
    ws_exts = wb[catalog_writer.SHEET_EXTS]
    ws_errors = wb[catalog_writer.SHEET_ERRORS]
    ws_summary = wb[catalog_writer.SHEET_SUMMARY]

    summary = {row[0].value: row[1].value for row in ws_summary.iter_rows(min_row=2, max_col=2)}
    exp_files, exp_folders, exp_bytes, exp_kb_sum = independent_count(SAMPLE_ROOT)

    list_kb_sum = sum(row[6] for row in ws_list.iter_rows(min_row=2, values_only=True))
    ext_count_sum = sum(row[2] for row in ws_exts.iter_rows(min_row=2, values_only=True))
    ext_kb_sum = sum(row[3] for row in ws_exts.iter_rows(min_row=2, values_only=True))

    checks = []

    def check(name: str, ok: bool, detail: str = "") -> None:
        checks.append((name, bool(ok), detail))

    # 1. 収集件数と一覧行数
    check("収集件数 == 一覧シート行数",
          result.file_count == ws_list.max_row - 1,
          f"scan={result.file_count} / sheet={ws_list.max_row - 1}")

    # 2. シート 5 枚
    check("シート 5 枚が揃う",
          wb.sheetnames == list(catalog_writer.SHEET_TITLES),
          ", ".join(wb.sheetnames))

    # 3. 一覧の書式
    check("一覧: フィルタ・先頭行固定・列幅",
          bool(ws_list.auto_filter.ref) and ws_list.freeze_panes == "A2"
          and ws_list.column_dimensions["B"].width > 0,
          f"filter={ws_list.auto_filter.ref} freeze={ws_list.freeze_panes}")

    # 4. 合計サイズの整合
    check("合計サイズ: サマリー == 独立走査の期待値",
          summary.get("合計サイズ(KB)") == bytes_to_kb(exp_bytes),
          f"summary={summary.get('合計サイズ(KB)')} expected={bytes_to_kb(exp_bytes)}")
    check("合計サイズ: 一覧の KB 合計 == 独立走査の期待値",
          list_kb_sum == exp_kb_sum,
          f"sheet={list_kb_sum} expected={exp_kb_sum}")
    check("拡張子集計: 件数合計 == 収集件数",
          ext_count_sum == result.file_count,
          f"ext={ext_count_sum} files={result.file_count}")
    check("拡張子集計: KB 合計はサマリーと概ね一致（丸め誤差内）",
          abs(ext_kb_sum - int(summary.get("合計サイズ(KB)", 0))) <= len(result.ext_rows),
          f"ext={ext_kb_sum} summary={summary.get('合計サイズ(KB)')} 拡張子種数={len(result.ext_rows)}")

    # 5. エラー一覧
    check("エラー一覧シートが見出し付きで存在",
          ws_errors.max_column == len(scan_core.ERROR_COLUMNS)
          and tuple(c.value for c in ws_errors[1]) == scan_core.ERROR_COLUMNS,
          f"rows={ws_errors.max_row - 1}")

    # 6. CSV
    csv_ok, csv_detail = True, []
    for sheet, path in info["csv"].items():
        with open(path, "rb") as fh:
            bom = fh.read(3) == b"\xef\xbb\xbf"
        with open(path, "r", encoding="utf-8-sig", newline="") as fh:
            lines = sum(1 for _ in fh)
        expect = result.file_count if sheet == catalog_writer.SHEET_LIST else len(result.folder_rows)
        csv_ok = csv_ok and bom and lines == expect + 1
        csv_detail.append(f"{sheet}: BOM={bom} rows={lines - 1}/{expect}")
    check("CSV: UTF-8 BOM 付き・行数が Excel と一致", csv_ok, "; ".join(csv_detail))

    # 7. 上書き防止の連番
    first_bytes = os.path.getsize(info["xlsx"])
    again = catalog_writer.run_catalog(result, OUT_DIR, with_csv=True)
    untouched = os.path.getsize(info["xlsx"]) == first_bytes
    numbered = again["xlsx"].endswith(" (2).xlsx")
    check("再実行で連番付きの別ファイル・既存は上書きされない",
          numbered and untouched and again["xlsx"] != info["xlsx"],
          f"2nd={os.path.basename(again['xlsx'])}")

    # 8. フォルダ集計の行数
    check("フォルダ集計の行数 == 収集フォルダ数 + 1",
          ws_folders.max_row - 1 == result.folder_count + 1,
          f"rows={ws_folders.max_row - 1} folders={result.folder_count}")

    # 9. 既定の除外
    names = {row[4] for row in ws_list.iter_rows(min_row=2, values_only=True)}
    rels = [row[1] for row in ws_list.iter_rows(min_row=2, values_only=True)]
    leaked = [n for n in ("desktop.ini", "Thumbs.db", "config", "index.js", ".hidden.txt") if n in names]
    check("既定の除外（.git/node_modules/隠し等）が効いている",
          not leaked and not any(r.startswith((".git/", "node_modules/")) for r in rels),
          f"混入={leaked or 'なし'}")

    # 10. 独立走査との突き合わせ
    check("独立走査(os.walk)の期待値と一致",
          result.file_count == exp_files and result.folder_count == exp_folders,
          f"scan={result.file_count}/{result.folder_count} walk={exp_files}/{exp_folders}")

    # 後片付け: 検証用に作った 2 回目のファイルを消すと NFR-5（削除しない）に反するので
    # 残す。ただし次回以降の走査からは EXCLUDE_NAMES で除外される。

    width = max(len(name) for name, _ok, _d in checks)
    print("=== 機械判定 ===")
    passed = 0
    for name, ok, detail in checks:
        mark = "PASS" if ok else "FAIL"
        passed += 1 if ok else 0
        print(f"[{mark}] {name.ljust(width)}  {detail}")
    total = len(checks)
    print(f"TOTAL: {passed}/{total} PASS")
    return 0 if passed == total else 1


if __name__ == "__main__":
    raise SystemExit(main())
