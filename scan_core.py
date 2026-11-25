# -*- coding: utf-8 -*-
"""走査と収集（GUI 非依存）。

`os.scandir` を自前で再帰（スタックを使った深さ優先）して、ファイル行・
フォルダ集計・拡張子集計・エラー一覧をまとめて返す。

- アクセス拒否・長すぎるパスなどの失敗は例外を投げず、エラー一覧に積む
- 進捗コールバックと中止フラグを受け取る（GUI 側から別スレッドで呼ぶ）
- 行データはタプルを順にためる軽量な形（大量ファイルでも破綻しにくい）
"""

from __future__ import annotations

import os
import stat
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, Optional

# ---------------------------------------------------------------------------
# 列の定義（Excel・CSV の見出しと同じ。ここが唯一の定義元）
# ---------------------------------------------------------------------------

FILE_COLUMNS = (
    "No",
    "相対パス",
    "階層レベル",
    "直下のフォルダ",
    "ファイル名",
    "拡張子",
    "サイズ(KB)",
    "更新日時",
    "作成日時",
)

FOLDER_COLUMNS = (
    "No",
    "フォルダ相対パス",
    "階層レベル",
    "直下ファイル数",
    "直下合計サイズ(KB)",
    "配下ファイル数",
    "配下合計サイズ(KB)",
    "配下最深階層",
)

EXT_COLUMNS = (
    "No",
    "拡張子",
    "件数",
    "合計サイズ(KB)",
    "最大ファイルサイズ(KB)",
    "最大ファイルの相対パス",
)

ERROR_COLUMNS = ("No", "種別", "パス", "理由")

SUMMARY_COLUMNS = ("項目", "値")

# 既定で除外するフォルダ・ファイル名（FR-9）
DEFAULT_EXCLUDE_NAMES = (
    ".git",
    "node_modules",
    "$RECYCLE.BIN",
    "System Volume Information",
    "desktop.ini",
    "Thumbs.db",
)

# 件数上限の既定値（FR-13）
DEFAULT_MAX_FILES = 200_000

# 対象フォルダ直下のファイルを「直下のフォルダ」列に出すときの表記
ROOT_FOLDER_LABEL = "(直下)"
# 対象フォルダ自身をフォルダ集計に出すときの表記
ROOT_REL_LABEL = "(対象フォルダ)"

# 拡張子なしのファイルの表記
NO_EXT_LABEL = "(なし)"

# Windows のエラーコードのうち「パスが長すぎる」系
_LONG_PATH_WINERRORS = (3, 123, 206)

_ATTR_HIDDEN = getattr(stat, "FILE_ATTRIBUTE_HIDDEN", 0x2)
_ATTR_SYSTEM = getattr(stat, "FILE_ATTRIBUTE_SYSTEM", 0x4)
_ATTR_REPARSE = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)


# ---------------------------------------------------------------------------
# 設定・結果
# ---------------------------------------------------------------------------


@dataclass
class ScanOptions:
    """走査の除外設定と安全装置（GUI のチェックボックスに対応する）。"""

    exclude_hidden: bool = True          # 隠しファイル／隠しフォルダ
    exclude_system: bool = True          # システムファイル／フォルダ
    exclude_default_names: bool = True   # DEFAULT_EXCLUDE_NAMES を除外
    exclude_names: tuple = DEFAULT_EXCLUDE_NAMES
    max_files: int = DEFAULT_MAX_FILES   # 件数上限（0 以下で無制限）

    def describe(self) -> list:
        """サマリーシートに出すための説明文（文字列のリスト）。"""
        lines = [
            "隠しファイル/フォルダ: " + ("除外する" if self.exclude_hidden else "除外しない"),
            "システムファイル/フォルダ: " + ("除外する" if self.exclude_system else "除外しない"),
            "既定の除外名: " + ("除外する" if self.exclude_default_names else "除外しない"),
        ]
        if self.exclude_default_names:
            lines.append("既定の除外名: " + ", ".join(self.exclude_names))
        lines.append(
            "件数上限: " + (f"{self.max_files:,} 件" if self.max_files > 0 else "無制限")
        )
        return lines


@dataclass
class ScanResult:
    """走査結果。`catalog_writer` がそのまま Excel/CSV に書き出せる形。"""

    root: str
    root_name: str
    options: ScanOptions
    started_at: datetime
    finished_at: datetime = None

    file_rows: list = field(default_factory=list)     # FILE_COLUMNS と同じ順のタプル
    folder_rows: list = field(default_factory=list)   # FOLDER_COLUMNS と同じ順のタプル
    ext_rows: list = field(default_factory=list)      # EXT_COLUMNS と同じ順のタプル
    error_rows: list = field(default_factory=list)    # ERROR_COLUMNS と同じ順のタプル

    file_count: int = 0
    folder_count: int = 0
    total_bytes: int = 0
    max_level: int = 0
    truncated: bool = False   # 件数上限に到達して打ち切った
    cancelled: bool = False   # 中止ボタンで打ち切った

    @property
    def error_count(self) -> int:
        return len(self.error_rows)

    @property
    def elapsed_seconds(self) -> float:
        if self.finished_at is None:
            return 0.0
        return (self.finished_at - self.started_at).total_seconds()

    def summary_rows(self) -> list:
        """サマリーシート（項目 / 値）の行を作る。"""
        opt = self.options
        rows = [
            ("対象フォルダ", self.root),
            ("対象フォルダ名", self.root_name),
            ("実行開始", fmt_dt(self.started_at)),
            ("実行終了", fmt_dt(self.finished_at)),
            ("所要時間(秒)", f"{self.elapsed_seconds:.2f}"),
            ("収集ファイル数", self.file_count),
            ("収集フォルダ数", self.folder_count),
            ("合計サイズ(KB)", bytes_to_kb(self.total_bytes)),
            ("合計サイズ(MB)", f"{self.total_bytes / (1024 * 1024):.2f}"),
            ("最深階層", self.max_level),
            ("エラー件数", self.error_count),
            ("件数上限に到達", "はい" if self.truncated else "いいえ"),
            ("中止された", "はい" if self.cancelled else "いいえ"),
        ]
        rows.extend((f"除外設定 {i}", t) for i, t in enumerate(opt.describe(), start=1))
        return rows


# ---------------------------------------------------------------------------
# 小さなユーティリティ
# ---------------------------------------------------------------------------


def bytes_to_kb(size_bytes: int) -> int:
    """バイト数を KB に直す（端数は四捨五入、0 バイトは 0）。"""
    return (int(size_bytes) + 512) // 1024


def fmt_dt(value) -> str:
    """タイムスタンプ（または datetime）を 'YYYY-MM-DD HH:MM:SS' の文字列に。"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        dt = value
    else:
        try:
            dt = datetime.fromtimestamp(value)
        except (OverflowError, OSError, ValueError):
            return ""
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def split_ext(name: str) -> str:
    """ファイル名から拡張子（先頭に . を含む小文字）を取り出す。無ければ '(なし)'。"""
    base, dot, ext = name.rpartition(".")
    if not dot or not base:
        return NO_EXT_LABEL
    return "." + ext.lower()


def join_rel(rel_dir: str, name: str) -> str:
    """相対パスをつなぐ（区切りは '/' で統一する）。"""
    return f"{rel_dir}/{name}" if rel_dir else name


def parent_rel(rel_path: str) -> str:
    """相対パスの親（対象フォルダ直下なら ''）。"""
    head, _slash, _tail = rel_path.rpartition("/")
    return head


def is_long_path_error(exc: BaseException) -> bool:
    """パス長が原因の OSError か（Windows のエラーコードで判断）。"""
    winerror = getattr(exc, "winerror", None)
    if winerror in _LONG_PATH_WINERRORS:
        return True
    if isinstance(exc, FileNotFoundError) and len(getattr(exc, "filename", "") or "") > 260:
        return True
    return False


def to_extended_path(path: str) -> str:
    """Windows の拡張パス表記（`\\\\?\\` / `\\\\?\\UNC\\`）に変換する。

    260 文字を超えるパスでも走査できるようにするための下準備。
    対象フォルダの絶対パスから呼び出すこと。
    """
    if path.startswith("\\\\?\\"):
        return path
    if path.startswith("\\\\"):
        return "\\\\?\\UNC\\" + path[2:]
    return "\\\\?\\" + path


def error_kind(exc: BaseException) -> str:
    """例外をエラー一覧の「種別」に分類する。"""
    if isinstance(exc, PermissionError):
        return "アクセス拒否"
    if is_long_path_error(exc):
        return "長すぎるパス"
    if isinstance(exc, FileNotFoundError):
        return "見付からない"
    if isinstance(exc, NotADirectoryError):
        return "フォルダでない"
    if isinstance(exc, OSError):
        return "読み取りエラー"
    return "その他"


def is_hidden_or_system(name: str, st, options: ScanOptions) -> bool:
    """隠し／システム属性の判定。

    - 隠し: Windows の隠し属性、または名前の先頭が「.」（属性が取れない環境の対策）
    - システム: Windows のシステム属性
    """
    attrs = getattr(st, "st_file_attributes", 0) if st is not None else 0
    hidden = bool(attrs & _ATTR_HIDDEN) or name.startswith(".")
    system = bool(attrs & _ATTR_SYSTEM)
    if options.exclude_hidden and hidden:
        return True
    return bool(options.exclude_system and system)


def _cancelled(should_cancel) -> bool:
    """中止フラグ（threading.Event でも callable でもよい）。"""
    if should_cancel is None:
        return False
    is_set = getattr(should_cancel, "is_set", None)
    if callable(is_set):
        return bool(is_set())
    return bool(should_cancel())


# ---------------------------------------------------------------------------
# 本体: 走査
# ---------------------------------------------------------------------------


def scan(
    root: str,
    options: Optional[ScanOptions] = None,
    progress: Optional[Callable[[int, str], None]] = None,
    should_cancel=None,
    progress_interval: int = 200,
    progress_seconds: float = 0.2,
) -> ScanResult:
    """`root` 配下を再帰的に走査して収集する（FR-2 / FR-9 / FR-10 / FR-11 / FR-13）。

    - `progress(処理ファイル数, 現在のフォルダ)` を随時呼ぶ（GUI の進捗表示用）
    - `should_cancel` が真を返したら走査を打ち切る（`result.cancelled = True`）
    - 収集ファイル数が `options.max_files` を超えた時点で打ち切る（`truncated = True`）
    - 失敗したフォルダ・ファイルは例外にせず `error_rows` に積む
    """
    options = options or ScanOptions()
    root = os.path.abspath(root)
    result = ScanResult(
        root=root,
        root_name=os.path.basename(root.rstrip("\\/")) or root,
        options=options,
        started_at=datetime.now(),
    )

    exclude_lookup = {n.lower() for n in options.exclude_names} if options.exclude_default_names else set()

    # フォルダ集計の作業テーブル: 相対パス -> [レベル, 直下数, 直下バイト, 配下数, 配下バイト, 最深階層]
    agg: dict = {"": [0, 0, 0, 0, 0, 0]}
    # 拡張子集計の作業テーブル: 拡張子 -> [件数, 合計バイト, 最大バイト, 最大ファイルの相対パス]
    ext_agg: dict = {}

    stack = [(root, "", 0)]  # (絶対パス, 相対パス, フォルダレベル)
    seen_files = 0
    last_report = time.monotonic()
    stop = False

    while stack and not stop:
        if _cancelled(should_cancel):
            result.cancelled = True
            break
        abs_dir, rel_dir, dir_level = stack.pop()
        try:
            scanner = _scandir(abs_dir)
        except OSError as exc:
            result.error_rows.append((error_kind(exc), _display_path(abs_dir, root), str(exc)))
            continue

        try:
            for entry in scanner:
                if _cancelled(should_cancel):
                    result.cancelled = True
                    stop = True
                    break
                name = entry.name
                if name.lower() in exclude_lookup:
                    continue
                try:
                    is_dir = entry.is_dir(follow_symlinks=False)
                    st = entry.stat(follow_symlinks=False)
                except OSError as exc:
                    result.error_rows.append((error_kind(exc), _display_path(entry.path, root), str(exc)))
                    continue
                if is_hidden_or_system(name, st, options):
                    continue

                rel_path = join_rel(rel_dir, name)
                if is_dir:
                    entry_level = dir_level + 1
                    if rel_path not in agg:
                        agg[rel_path] = [entry_level, 0, 0, 0, 0, 0]
                    if entry_level > result.max_level:
                        result.max_level = entry_level
                    if not _is_reparse(st):
                        stack.append((entry.path, rel_path, entry_level))
                    continue

                # --- ファイル ---
                size = int(getattr(st, "st_size", 0) or 0)
                mtime = fmt_dt(getattr(st, "st_mtime", None))
                ctime = fmt_dt(getattr(st, "st_ctime", None))
                level = dir_level + 1
                seen_files += 1
                result.file_rows.append(
                    (
                        seen_files,
                        rel_path,
                        level,
                        rel_dir or ROOT_FOLDER_LABEL,
                        name,
                        split_ext(name),
                        bytes_to_kb(size),
                        mtime,
                        ctime,
                    )
                )
                result.total_bytes += size
                if level > result.max_level:
                    result.max_level = level

                slot = agg.get(rel_dir)
                if slot is not None:
                    slot[1] += 1
                    slot[2] += size
                    slot[5] = max(slot[5], level)

                ext = split_ext(name)
                slot = ext_agg.get(ext)
                if slot is None:
                    ext_agg[ext] = [1, size, size, rel_path]
                else:
                    slot[0] += 1
                    slot[1] += size
                    if size > slot[2]:
                        slot[2] = size
                        slot[3] = rel_path

                if options.max_files > 0 and seen_files >= options.max_files:
                    result.truncated = True
                    stop = True
                    break

                if seen_files % progress_interval == 0 or (time.monotonic() - last_report) >= progress_seconds:
                    last_report = time.monotonic()
                    if progress is not None:
                        progress(seen_files, abs_dir)
        except OSError as exc:
            # ネットワークドライブなどでは列挙の途中で失敗することがある
            result.error_rows.append((error_kind(exc), _display_path(abs_dir, root), str(exc)))
        finally:
            try:
                scanner.close()
            except OSError:
                pass

    result.file_count = seen_files
    _finish_folders(result, agg)
    _finish_exts(result, ext_agg)
    if result.truncated:
        limit = f"{options.max_files:,}" if options.max_files > 0 else "無制限"
        result.error_rows.append(
            ("上限到達", root, f"収集ファイル数が件数上限（{limit} 件）に達したため走査を打ち切りました")
        )
    # エラー一覧の No を振る（ERROR_COLUMNS = No / 種別 / パス / 理由）
    result.error_rows = [(i,) + tuple(row) for i, row in enumerate(result.error_rows, start=1)]
    result.finished_at = datetime.now()
    return result


def _scandir(path: str):
    """`os.scandir`。パス長が原因で失敗したときは拡張パス表記で一度だけ再試行する。"""
    try:
        return os.scandir(path)
    except OSError as exc:
        if is_long_path_error(exc) and not path.startswith("\\\\?\\"):
            try:
                return os.scandir(to_extended_path(path))
            except OSError:
                raise exc
        raise exc


def _is_reparse(st) -> bool:
    """シンボリックリンクやジャンクション（再解析点）かどうか。ループ防止用に使う。"""
    return bool(getattr(st, "st_file_attributes", 0) & _ATTR_REPARSE)


def _display_path(abs_path: str, root: str) -> str:
    """エラー一覧に出すパス（対象フォルダ配下は相対パスにして読みやすくする）。"""
    cleaned = abs_path
    if cleaned.startswith("\\\\?\\UNC\\"):
        cleaned = "\\\\" + cleaned[9:]
    elif cleaned.startswith("\\\\?\\"):
        cleaned = cleaned[4:]
    try:
        if os.path.commonpath([cleaned, root]) == root and cleaned != root:
            return cleaned[len(root) :].lstrip("\\/")
    except ValueError:
        pass
    return cleaned


def _finish_folders(result: ScanResult, agg: dict) -> None:
    """フォルダ集計を確定する。深い階層から順に親へ配下合計を伝播させる。"""
    # 配下合計は「自身の直下」を出発点にして、子フォルダの配下合計を足し込む
    for slot in agg.values():
        slot[3] = slot[1]
        slot[4] = slot[2]
    for rel in sorted(agg, key=lambda p: (-agg[p][0], p)):
        slot = agg[rel]
        if rel:
            parent = agg.get(parent_rel(rel))
            if parent is not None:
                parent[3] += slot[3]
                parent[4] += slot[4]
                parent[5] = max(parent[5], slot[5])

    rows = []
    for i, rel in enumerate(sorted(agg, key=lambda p: (agg[p][0], p)), start=1):
        level, direct_n, direct_b, total_n, total_b, deepest = agg[rel]
        rows.append(
            (
                i,
                rel or ROOT_REL_LABEL,
                level,
                direct_n,
                bytes_to_kb(direct_b),
                total_n,
                bytes_to_kb(total_b),
                deepest,
            )
        )
    result.folder_rows = rows
    result.folder_count = max(len(rows) - 1, 0)  # 対象フォルダ自身は数えない


def _finish_exts(result: ScanResult, ext_agg: dict) -> None:
    """拡張子集計を確定する（件数の多い順、同数はサイズの大きい順）。"""
    rows = []
    ordered = sorted(ext_agg.items(), key=lambda kv: (-kv[1][0], -kv[1][1], kv[0]))
    for i, (ext, (count, total_b, max_b, max_rel)) in enumerate(ordered, start=1):
        rows.append((i, ext, count, bytes_to_kb(total_b), bytes_to_kb(max_b), max_rel))
    result.ext_rows = rows


def preview_text(result: ScanResult) -> str:
    """事前プレビュー（FR-3）用の文章を作る。GUI と検証スクリプトの両方で使う。"""
    lines = [
        f"対象フォルダ: {result.root}",
        f"想定件数: {result.file_count:,} ファイル",
        f"フォルダ数: {result.folder_count:,}",
        f"合計サイズ: {result.total_bytes:,} バイト（{bytes_to_kb(result.total_bytes):,} KB / "
        f"{result.total_bytes / (1024 * 1024):.2f} MB）",
        f"最深階層: {result.max_level}",
        f"エラー件数: {result.error_count:,}",
    ]
    if result.ext_rows:
        top = ", ".join(f"{r[1]} {r[2]:,}件" for r in result.ext_rows[:8])
        lines.append(f"拡張子（上位）: {top}")
    if result.truncated:
        lines.append(
            f"注意: 件数上限（{result.options.max_files:,} 件）に達したため途中で打ち切りました。"
            "上限を上げて再実行してください。"
        )
    if result.cancelled:
        lines.append("注意: 中止したため途中までしか走査していません。")
    if result.file_count > 50_000:
        lines.append("注意: 件数が多いので Excel の生成に時間がかかります。")
    return "\n".join(lines)
