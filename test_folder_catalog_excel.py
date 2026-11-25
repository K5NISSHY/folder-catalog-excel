# -*- coding: utf-8 -*-
"""単体テスト（NFR-7）。標準の `unittest` のみを使う（pytest は使わない）。

対象:
- 収集ロジック（scan_core.scan）: 階層・相対パス・サイズ・日時・集計
- 除外判定（隠し／システム／既定の除外名）
- 安全装置（件数上限・中止）とエラー一覧（アクセス拒否・長すぎるパス）
- 出力ファイル名の連番規則（FR-8）と上書き防止（NFR-5）
- Excel の書式（シート構成・フィルタ・先頭行固定・列幅）
- CSV のエンコーディング（UTF-8 BOM 付き）と行数（FR-16）

実行: `.venv\\Scripts\\python.exe -m unittest -v test_folder_catalog_excel`
"""

from __future__ import annotations

import hashlib
import os
import sys
import tempfile
import unittest
from datetime import datetime
from unittest import mock

import catalog_writer
import scan_core
from scan_core import (
    DEFAULT_EXCLUDE_NAMES,
    ERROR_COLUMNS,
    EXT_COLUMNS,
    FILE_COLUMNS,
    FOLDER_COLUMNS,
    NO_EXT_LABEL,
    ROOT_FOLDER_LABEL,
    ScanOptions,
    bytes_to_kb,
)

# 出力ファイル名の組み立ては catalog_writer 側にある
build_base_name = catalog_writer.build_base_name

IS_WINDOWS = sys.platform == "win32"
_HIDDEN_ATTR = 0x2
_SYSTEM_ATTR = 0x4


def set_windows_attrs(path: str, attrs: int) -> bool:
    """Windows のファイル属性を設定する（隠し・システム）。失敗したら False。"""
    if not IS_WINDOWS:
        return False
    try:
        import ctypes

        return bool(ctypes.windll.kernel32.SetFileAttributesW(path, attrs))
    except Exception:  # 環境依存。属性が使えない場合はテスト側でスキップする
        return False


def write_file(path: str, size: int, mtime: float = None) -> None:
    """指定サイズのファイルを作る（中身は 0 バイト埋め）。"""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(b"\0" * size)
    if mtime is not None:
        os.utime(path, (mtime, mtime))


def _read_bytes(path: str) -> bytes:
    with open(path, "rb") as fh:
        return fh.read()


class TreeFixture(unittest.TestCase):
    """日本語名・スペース・多階層の見本ツリーを一時フォルダに作る。"""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory(prefix="fce_test_")
        self.root = self._tmp.name
        self.addCleanup(self._tmp.cleanup)

        self.expected_default_files = 5
        write_file(os.path.join(self.root, "ルート.txt"), 100)
        write_file(os.path.join(self.root, "データ ファイル.xlsx"), 2048)
        write_file(os.path.join(self.root, "サブフォルダA", "ノート.txt"), 512)
        write_file(os.path.join(self.root, "サブフォルダA", "画像.png"), 3000)
        write_file(os.path.join(self.root, "サブフォルダA", "入れ子B", "深いC", "最深ファイル.docx"), 1)

        # 既定で除外されるもの
        write_file(os.path.join(self.root, ".git", "config"), 40)
        write_file(os.path.join(self.root, "desktop.ini"), 13)
        write_file(os.path.join(self.root, ".hidden.txt"), 7)
        self.hidden_file = os.path.join(self.root, "hidden_attr.bin")
        write_file(self.hidden_file, 9)
        self.system_file = os.path.join(self.root, "system_attr.sys")
        write_file(self.system_file, 11)

        # 隠し・システム属性を実際に付ける（属性が使えない環境では件数の期待値を変える）
        self.attrs_set = set_windows_attrs(self.hidden_file, _HIDDEN_ATTR)
        set_windows_attrs(self.system_file, _SYSTEM_ATTR)
        self.expected_default_files = 5 if self.attrs_set else 7

    def scan(self, **kwargs) -> scan_core.ScanResult:
        return scan_core.scan(self.root, **kwargs)


class TestScanCollect(TreeFixture):
    def test_collects_visible_files(self):
        result = self.scan()
        self.assertEqual(result.file_count, self.expected_default_files)
        self.assertEqual(len(result.file_rows), self.expected_default_files)
        self.assertEqual(result.folder_count, 3)
        self.assertEqual(result.max_level, 4)
        self.assertEqual(result.error_count, 0)
        self.assertFalse(result.truncated)
        self.assertFalse(result.cancelled)

    def test_relative_paths_levels_and_parent(self):
        rows = {row[1]: row for row in self.scan().file_rows}
        deep = "サブフォルダA/入れ子B/深いC/最深ファイル.docx"
        self.assertIn(deep, rows)
        row = rows[deep]
        self.assertGreaterEqual(row[0], 1)  # No は 1 始まり
        self.assertEqual(row[2], 4)  # 階層レベル
        self.assertEqual(row[3], "サブフォルダA/入れ子B/深いC")  # 直下のフォルダ
        self.assertEqual(row[4], "最深ファイル.docx")
        self.assertEqual(row[5], ".docx")
        self.assertEqual(rows["ルート.txt"][3], ROOT_FOLDER_LABEL)
        self.assertEqual(rows["データ ファイル.xlsx"][5], ".xlsx")
        # 日時は 'YYYY-MM-DD HH:MM:SS'
        for value in (row[7], row[8]):
            datetime.strptime(value, "%Y-%m-%d %H:%M:%S")

    def test_sizes_and_totals(self):
        result = self.scan()
        self.assertEqual(result.total_bytes, 100 + 2048 + 512 + 3000 + 1)
        rows = {row[1]: row for row in result.file_rows}
        self.assertEqual(rows["ルート.txt"][6], 0)          # 100 バイト → 0 KB
        self.assertEqual(rows["データ ファイル.xlsx"][6], 2)  # 2048 バイト → 2 KB
        self.assertEqual(result.folder_rows[0][5], self.expected_default_files)  # 対象フォルダの配下ファイル数
        self.assertEqual(result.folder_rows[0][6], bytes_to_kb(result.total_bytes))

    def test_no_duplicate_no_numbers(self):
        numbers = [row[0] for row in self.scan().file_rows]
        self.assertEqual(numbers, list(range(1, len(numbers) + 1)))

    def test_japanese_and_space_names_survive(self):
        name = "日本語 の フォルダ/長い名前の ファイル 名 付き.txt"
        write_file(os.path.join(self.root, "日本語 の フォルダ", "長い名前の ファイル 名 付き.txt"), 5)
        rows = {row[1] for row in self.scan().file_rows}
        self.assertIn(name, rows)


class TestScanExclusion(TreeFixture):
    def test_default_names_excluded(self):
        names = {row[4] for row in self.scan().file_rows}
        self.assertNotIn("desktop.ini", names)
        self.assertNotIn("config", names)
        folders = {row[1] for row in self.scan().folder_rows}
        self.assertNotIn(".git", folders)

    def test_default_names_included_when_unchecked(self):
        options = ScanOptions(exclude_default_names=False, exclude_hidden=False, exclude_system=False)
        result = scan_core.scan(self.root, options)
        names = {row[4] for row in result.file_rows}
        self.assertIn("desktop.ini", names)
        self.assertIn("config", names)
        self.assertIn(".git", {row[1] for row in result.folder_rows})

    def test_hidden_by_name_excluded(self):
        names = {row[4] for row in self.scan().file_rows}
        self.assertNotIn(".hidden.txt", names)

    def test_hidden_and_system_attributes(self):
        if not set_windows_attrs(self.hidden_file, _HIDDEN_ATTR):
            self.skipTest("この環境ではファイル属性を設定できません")
        set_windows_attrs(self.system_file, _SYSTEM_ATTR)

        names = {row[4] for row in self.scan().file_rows}
        self.assertNotIn("hidden_attr.bin", names)
        self.assertNotIn("system_attr.sys", names)

        options = ScanOptions(exclude_hidden=False, exclude_system=False)
        names = {row[4] for row in scan_core.scan(self.root, options).file_rows}
        self.assertIn("hidden_attr.bin", names)
        self.assertIn("system_attr.sys", names)

        options = ScanOptions(exclude_hidden=False, exclude_system=True)
        names = {row[4] for row in scan_core.scan(self.root, options).file_rows}
        self.assertIn("hidden_attr.bin", names)
        self.assertNotIn("system_attr.sys", names)

    def test_exclude_names_list_is_case_insensitive(self):
        options = ScanOptions(exclude_names=tuple(DEFAULT_EXCLUDE_NAMES) + ("サブフォルダA",))
        self.assertEqual(scan_core.scan(self.root, options).file_count, 2)


class TestScanSafety(TreeFixture):
    def test_max_files_truncates(self):
        options = ScanOptions(max_files=3)
        result = scan_core.scan(self.root, options)
        self.assertTrue(result.truncated)
        self.assertEqual(result.file_count, 3)
        self.assertEqual(len(result.file_rows), 3)
        self.assertIn("上限到達", [row[1] for row in result.error_rows])

    def test_max_files_zero_means_unlimited(self):
        options = ScanOptions(max_files=0)
        result = scan_core.scan(self.root, options)
        self.assertFalse(result.truncated)
        self.assertEqual(result.file_count, self.expected_default_files)

    def test_cancel_stops_without_crash(self):
        state = {"n": 0}

        def should_cancel():
            state["n"] += 1
            return state["n"] > 5

        result = scan_core.scan(self.root, should_cancel=should_cancel)
        self.assertTrue(result.cancelled)
        self.assertLess(result.file_count, self.expected_default_files)

    def test_progress_callback_reports(self):
        seen = []
        scan_core.scan(self.root, progress=lambda count, folder: seen.append((count, folder)),
                       progress_interval=1)
        self.assertTrue(seen)
        self.assertEqual([c for c, _ in seen], sorted(c for c, _ in seen))
        self.assertTrue(all(isinstance(f, str) for _, f in seen))

    def test_permission_error_is_recorded(self):
        blocked = os.path.join(self.root, "サブフォルダA")
        original = os.scandir

        def fake_scandir(path):
            if os.path.normcase(blocked) in os.path.normcase(str(path)):
                raise PermissionError(5, "アクセスが拒否されました", str(path))
            return original(path)

        with mock.patch("scan_core.os.scandir", side_effect=fake_scandir):
            result = self.scan()
        kinds = [row[1] for row in result.error_rows]
        self.assertIn("アクセス拒否", kinds)
        self.assertEqual(result.error_rows[0][2], "サブフォルダA")
        # 失敗しても他のファイルは収集できている（クラッシュしない）
        self.assertEqual(result.file_count, 2)

    def test_long_path_error_is_recorded(self):
        blocked = os.path.join(self.root, "サブフォルダA")
        original = os.scandir

        def fake_scandir(path):
            if os.path.normcase(blocked) in os.path.normcase(str(path)):
                raise OSError(1, "ファイル名が長すぎます", str(path), 206, None)
            return original(path)

        with mock.patch("scan_core.os.scandir", side_effect=fake_scandir):
            result = self.scan()
        self.assertIn("長すぎるパス", [row[1] for row in result.error_rows])

    def test_extended_path_helper(self):
        self.assertEqual(scan_core.to_extended_path("D:\\a\\b"), "\\\\?\\D:\\a\\b")
        self.assertEqual(scan_core.to_extended_path("\\\\server\\share\\a"), "\\\\?\\UNC\\server\\share\\a")
        self.assertEqual(scan_core.to_extended_path("\\\\?\\D:\\a"), "\\\\?\\D:\\a")
        self.assertTrue(scan_core.is_long_path_error(OSError(1, "x", "y", 206, None)))
        self.assertFalse(scan_core.is_long_path_error(PermissionError(5, "x", "y")))


class TestAggregation(TreeFixture):
    def test_folder_aggregation_direct_and_recursive(self):
        rows = {row[1]: row for row in self.scan().folder_rows}
        a = rows["サブフォルダA"]
        self.assertEqual(a[2], 1)          # 階層レベル
        self.assertEqual(a[3], 2)          # 直下ファイル数
        self.assertEqual(a[5], 3)          # 配下ファイル数（入れ子B/深いC を含む）
        self.assertEqual(a[6], bytes_to_kb(512 + 3000 + 1))
        self.assertEqual(a[7], 4)          # 配下最深階層
        deep = rows["サブフォルダA/入れ子B/深いC"]
        self.assertEqual((deep[3], deep[5], deep[7]), (1, 1, 4))

    def test_ext_aggregation(self):
        rows = {row[1]: row for row in self.scan().ext_rows}
        self.assertEqual(rows[".txt"][2], 2)                 # 件数
        self.assertEqual(rows[".txt"][3], bytes_to_kb(100 + 512))  # 合計サイズ
        self.assertEqual(rows[".png"][4], bytes_to_kb(3000))  # 最大ファイルサイズ
        self.assertEqual(rows[".png"][5], "サブフォルダA/画像.png")
        counts = [row[2] for row in self.scan().ext_rows]
        self.assertEqual(counts, sorted(counts, reverse=True))

    def test_bytes_to_kb_rounding(self):
        self.assertEqual(bytes_to_kb(0), 0)
        self.assertEqual(bytes_to_kb(511), 0)
        self.assertEqual(bytes_to_kb(512), 1)
        self.assertEqual(bytes_to_kb(1024), 1)
        self.assertEqual(bytes_to_kb(1536), 2)

    def test_split_ext(self):
        self.assertEqual(scan_core.split_ext("a.xlsx"), ".xlsx")
        self.assertEqual(scan_core.split_ext("A.DOCX"), ".docx")
        self.assertEqual(scan_core.split_ext("archive.tar.gz"), ".gz")
        self.assertEqual(scan_core.split_ext("README"), NO_EXT_LABEL)
        self.assertEqual(scan_core.split_ext(".gitignore"), NO_EXT_LABEL)

    def test_summary_rows_have_key_items(self):
        keys = {row[0] for row in self.scan().summary_rows()}
        for key in ("対象フォルダ", "収集ファイル数", "収集フォルダ数", "合計サイズ(KB)", "最深階層", "エラー件数"):
            self.assertIn(key, keys)


class TestOutputNaming(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory(prefix="fce_names_")
        self.out = self._tmp.name
        self.addCleanup(self._tmp.cleanup)

    def test_base_name_format(self):
        base = build_base_name("検証用フォルダ", datetime(2026, 11, 24, 12, 0))
        self.assertEqual(base, "検証用フォルダ_目録_20261124_1200")

    def test_invalid_chars_are_replaced(self):
        self.assertEqual(catalog_writer.sanitize_name('a<b>:c"d'), "a_b__c_d")

    def test_numbering_when_exists(self):
        base = build_base_name("x", datetime(2026, 11, 24, 12, 0))
        first, _ = catalog_writer.unique_output_paths(self.out, base)
        self.assertTrue(first.endswith("x_目録_20261124_1200.xlsx"))
        open(first, "wb").close()
        second, _ = catalog_writer.unique_output_paths(self.out, base)
        self.assertTrue(second.endswith("x_目録_20261124_1200 (2).xlsx"))
        open(second, "wb").close()
        third, _ = catalog_writer.unique_output_paths(self.out, base)
        self.assertTrue(third.endswith("x_目録_20261124_1200 (3).xlsx"))

    def test_csv_shares_the_same_number(self):
        base = "x_目録_20261124_1200"
        open(os.path.join(self.out, f"{base}.xlsx"), "wb").close()
        xlsx, csvs = catalog_writer.unique_output_paths(self.out, base, with_csv=True)
        self.assertTrue(xlsx.endswith(f"{base} (2).xlsx"))
        self.assertTrue(csvs[catalog_writer.SHEET_LIST].endswith(f"{base} (2)_一覧.csv"))
        self.assertTrue(csvs[catalog_writer.SHEET_FOLDERS].endswith(f"{base} (2)_フォルダ集計.csv"))

    def test_no_csv_paths_when_off(self):
        _xlsx, csvs = catalog_writer.unique_output_paths(self.out, "x", with_csv=False)
        self.assertEqual(csvs, {})


class TestWriting(TreeFixture):
    def setUp(self) -> None:
        super().setUp()
        self._tmp_out = tempfile.TemporaryDirectory(prefix="fce_out_")
        self.out = self._tmp_out.name
        self.addCleanup(self._tmp_out.cleanup)
        self.result = self.scan()

    def _write(self, with_csv=False):
        return catalog_writer.run_catalog(self.result, self.out, with_csv=with_csv,
                                          when=datetime(2026, 11, 24, 12, 0))

    def test_excel_sheets_and_rows(self):
        info = self._write()
        from openpyxl import load_workbook

        wb = load_workbook(info["xlsx"])
        self.assertEqual(wb.sheetnames, list(catalog_writer.SHEET_TITLES))
        ws = wb[catalog_writer.SHEET_LIST]
        self.assertEqual(ws.max_row, self.result.file_count + 1)
        self.assertEqual(tuple(cell.value for cell in ws[1]), FILE_COLUMNS)
        self.assertEqual(tuple(cell.value for cell in wb[catalog_writer.SHEET_FOLDERS][1]), FOLDER_COLUMNS)
        self.assertEqual(tuple(cell.value for cell in wb[catalog_writer.SHEET_EXTS][1]), EXT_COLUMNS)
        self.assertEqual(tuple(cell.value for cell in wb[catalog_writer.SHEET_ERRORS][1]), ERROR_COLUMNS)
        self.assertEqual(ws.max_row - 1, len(self.result.file_rows))

    def test_excel_formatting(self):
        info = self._write()
        from openpyxl import load_workbook

        wb = load_workbook(info["xlsx"])
        for title in (catalog_writer.SHEET_LIST, catalog_writer.SHEET_FOLDERS,
                      catalog_writer.SHEET_EXTS, catalog_writer.SHEET_ERRORS):
            ws = wb[title]
            self.assertEqual(ws.freeze_panes, "A2", f"{title} の先頭行が固定されていない")
            self.assertTrue(ws.auto_filter.ref, f"{title} にフィルタが付いていない")
            self.assertTrue(ws.column_dimensions["A"].width > 0, f"{title} の列幅が未設定")
        self.assertEqual(wb[catalog_writer.SHEET_LIST].auto_filter.ref,
                         f"A1:I{self.result.file_count + 1}")

    def test_never_overwrites_existing_file(self):
        info = self._write()
        path = info["xlsx"]
        before = hashlib.sha256(_read_bytes(path)).hexdigest()
        with self.assertRaises(FileExistsError):
            catalog_writer.write_catalog(self.result, path)
        self.assertEqual(hashlib.sha256(_read_bytes(path)).hexdigest(), before)
        # 再実行は連番で別ファイルになる
        again = self._write()
        self.assertNotEqual(again["xlsx"], path)
        self.assertTrue(os.path.exists(again["xlsx"]))

    def test_csv_bom_and_rows(self):
        info = self._write(with_csv=True)
        self.assertEqual(set(info["csv"]), set(catalog_writer.CSV_SHEETS))
        for sheet, path in info["csv"].items():
            with open(path, "rb") as fh:
                self.assertEqual(fh.read(3), b"\xef\xbb\xbf", f"{sheet} の CSV に BOM がない")
            with open(path, "r", encoding="utf-8-sig", newline="") as fh:
                lines = [line for line in fh]
            expected = self.result.file_count if sheet == catalog_writer.SHEET_LIST else len(self.result.folder_rows)
            self.assertEqual(len(lines), expected + 1, f"{sheet} の CSV 行数が違う")

    def test_csv_content_roundtrip_japanese(self):
        info = self._write(with_csv=True)
        import csv

        path = info["csv"][catalog_writer.SHEET_LIST]
        with open(path, "r", encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.reader(fh))
        self.assertEqual(rows[0], list(FILE_COLUMNS))
        names = [row[4] for row in rows[1:]]
        self.assertIn("最深ファイル.docx", names)
        self.assertIn("データ ファイル.xlsx", names)

    def test_csv_off_by_default(self):
        info = self._write()
        self.assertEqual(info["csv"], {})
        self.assertEqual([f for f in os.listdir(self.out) if f.endswith(".csv")], [])

    def test_error_sheet_lists_failures(self):
        blocked = os.path.join(self.root, "サブフォルダA")
        original = os.scandir

        def fake_scandir(path):
            if os.path.normcase(blocked) in os.path.normcase(str(path)):
                raise PermissionError(5, "アクセスが拒否されました", str(path))
            return original(path)

        with mock.patch("scan_core.os.scandir", side_effect=fake_scandir):
            result = self.scan()
        info = catalog_writer.run_catalog(result, self.out, when=datetime(2026, 11, 24, 12, 0))
        from openpyxl import load_workbook

        ws = load_workbook(info["xlsx"])[catalog_writer.SHEET_ERRORS]
        self.assertGreaterEqual(ws.max_row, 2)
        self.assertEqual(ws.cell(row=2, column=2).value, "アクセス拒否")


class TestHelpers(unittest.TestCase):
    def test_preview_text_mentions_key_numbers(self):
        with tempfile.TemporaryDirectory(prefix="fce_prev_") as root:
            write_file(os.path.join(root, "a.txt"), 1024)
            result = scan_core.scan(root)
            text = scan_core.preview_text(result)
        for token in ("想定件数", "合計サイズ", "最深階層", "対象フォルダ"):
            self.assertIn(token, text)
        self.assertIn("1 ファイル", text)

    def test_display_path_is_relative_to_root(self):
        root = "D:\\data\\project"
        self.assertEqual(scan_core._display_path("D:\\data\\project\\sub\\a.txt", root), "sub\\a.txt")
        self.assertEqual(scan_core._display_path("D:\\other\\a.txt", root), "D:\\other\\a.txt")


if __name__ == "__main__":
    unittest.main(verbosity=2)
