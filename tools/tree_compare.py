# -*- coding: utf-8 -*-
"""記事用の比較材料を作る（`tree /F` 側 と Excel 側の突き合わせ）。

同じサンプルフォルダに対して、
- `tree /F` と同じ形式のツリー図テキスト（サブプロセスではなく Python 側で生成。
  コードページで文字化けしないよう UTF-8・BOM 付きで保存）
- このツールの Excel 側の集計値（件数・合計サイズ・フォルダ数）

を作り、両方の**ファイル数・フォルダ数が一致する**ことを機械判定で出す。
ツリー図は同じ除外設定で生成するので、件数は直接比較できる。

実行: `.venv\\Scripts\\python.exe tools\\tree_compare.py`
"""

from __future__ import annotations

import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

import scan_core  # noqa: E402
from scan_core import DEFAULT_EXCLUDE_NAMES, ScanOptions, bytes_to_kb, is_hidden_or_system  # noqa: E402

SAMPLE_ROOT = os.path.join(REPO_ROOT, "samples")
OUT_DIR = os.path.join(SAMPLE_ROOT, "検証結果")
TREE_TXT = os.path.join(OUT_DIR, "tree_F.txt")

EXCLUDE_NAMES = tuple(DEFAULT_EXCLUDE_NAMES) + (os.path.basename(OUT_DIR),)


def list_children(path: str, options: ScanOptions) -> tuple:
    """(サブフォルダ, ファイル) を名前の昇順で返す（除外設定を適用）。"""
    dirs, files = [], []
    try:
        with os.scandir(path) as scanner:
            for entry in scanner:
                try:
                    st = entry.stat(follow_symlinks=False)
                except OSError:
                    continue
                name = entry.name
                if name.lower() in {n.lower() for n in EXCLUDE_NAMES}:
                    continue
                if is_hidden_or_system(name, st, options):
                    continue
                (dirs if entry.is_dir(follow_symlinks=False) else files).append(name)
    except OSError:
        pass
    return sorted(dirs), sorted(files)


def build_tree_text(root: str, options: ScanOptions) -> str:
    """`tree /F` と同じ形式のツリー図テキストを作る（ディレクトリ → ファイルの順）。"""
    lines = [
        f"フォルダー PATH: {os.path.basename(root)}",
        "ボリューム シリアル番号 は 0000-0000 です",
        ".",
    ]
    file_count = 0
    dir_count = 0

    def walk(path: str, prefix: str) -> None:
        nonlocal file_count, dir_count
        dirs, files = list_children(path, options)
        items = [(d, True) for d in dirs] + [(f, False) for f in files]
        for index, (name, is_dir) in enumerate(items):
            last = index == len(items) - 1
            connector = "└───" if last else "├───"
            if is_dir:
                dir_count += 1
                lines.append(f"{prefix}{connector}{name}")
                walk(os.path.join(path, name), prefix + ("    " if last else "│   "))
            else:
                file_count += 1
                lines.append(f"{prefix}    {name}")

    walk(root, "")
    lines.append("")
    lines.append(f"ファイル: {file_count} 個")
    lines.append(f"ディレクトリ: {dir_count} 個")
    return "\n".join(lines)


def main() -> int:
    if not os.path.isdir(SAMPLE_ROOT):
        print("samples/ が見つかりません。先に tools/make_samples.py を実行してください。")
        return 2
    os.makedirs(OUT_DIR, exist_ok=True)

    options = ScanOptions(exclude_names=EXCLUDE_NAMES)
    tree_text = build_tree_text(SAMPLE_ROOT, options)
    with open(TREE_TXT, "w", encoding="utf-8-sig", newline="") as fh:
        fh.write(tree_text)

    # ツリー図テキストを後から読み戻して数える（生成時のカウンタに頼らない）
    body = tree_text.splitlines()[3:]
    tree_dirs = sum(1 for line in body if "───" in line)
    tree_files = sum(1 for line in body if line.strip() and "───" not in line
                     and not line.startswith("ファイル:") and not line.startswith("ディレクトリ:"))
    tree_lines = len([line for line in body if line.strip()])

    # Excel 側
    result = scan_core.scan(SAMPLE_ROOT, options)

    print(f"ツリー図テキスト: {TREE_TXT}")
    print(f"  行数 {tree_lines} 行（ファイル {tree_files} 行・フォルダ {tree_dirs} 行）")
    print()
    print("=== tree /F 形式（Python 生成）と Excel 側の突き合わせ ===")
    rows = [
        ("ファイル数", f"{tree_files:,}", f"{result.file_count:,}",
         "一致" if tree_files == result.file_count else "不一致"),
        ("フォルダ数", f"{tree_dirs:,}", f"{result.folder_count:,}",
         "一致" if tree_dirs == result.folder_count else "不一致"),
        ("合計サイズ", "記載なし", f"{bytes_to_kb(result.total_bytes):,} KB", "-"),
        ("更新日時・作成日時", "記載なし", f"{result.file_count:,} 行すべてに記載", "-"),
        ("階層レベル", "罫線の数でしか分からない", "別列（フィルタで絞り込める）", "-"),
        ("拡張子別の集計", "記載なし", f"{len(result.ext_rows)} 種", "-"),
        ("並べ替え・絞り込み", "不可（テキスト）", "可（フィルタ・ソート）", "-"),
        ("エラーの記録", "不可", f"{result.error_count:,} 件", "-"),
    ]
    name_width = max(len(r[0]) for r in rows)
    print(f"{'項目'.ljust(name_width)} | tree /F 側            | Excel 側")
    for name, tree_side, excel_side, verdict in rows:
        print(f"{name.ljust(name_width)} | {tree_side:<22} | {excel_side}  {verdict}")

    ok = tree_files == result.file_count and tree_dirs == result.folder_count
    print()
    print(f"機械判定: {'PASS' if ok else 'FAIL'}"
          f"（tree 側 {tree_files} ファイル / Excel 側 {result.file_count} ファイル）")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
