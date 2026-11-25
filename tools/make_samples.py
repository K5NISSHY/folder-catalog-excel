# -*- coding: utf-8 -*-
"""検証用サンプルフォルダ生成（NFR-8）。

`samples/` 配下に、日本語名・多階層（5 階層）・各種拡張子・サイズ差・更新日の差が
ある**ダミーのみ**のツリーを作る（個人情報・実データは一切使わない）。

- 既存ファイルは上書きしない（既にあるファイルはスキップする）
- 実行すると生成したファイル数・フォルダ数・合計サイズを表示する

実行: `.venv\\Scripts\\python.exe tools\\make_samples.py`
"""

from __future__ import annotations

import os
import sys
import time

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLE_ROOT = os.path.join(REPO_ROOT, "samples")

DAY = 86400

# (相対パス, サイズ[バイト], 何日前に更新したか)
# 既定で除外されるもの（.git / node_modules / desktop.ini / Thumbs.db / 隠し）も入れて、
# 除外が効いているか確かめる材料にする。
def build_spec() -> list:
    spec: list = []

    # --- 直下 -----------------------------------------------------------
    spec += [
        ("README.txt", 1_800, 2),
        ("目録サンプル_説明.txt", 2_400, 5),
        ("メモ.txt", 320, 130),
        ("desktop.ini", 13, 400),                 # 既定で除外
        ("Thumbs.db", 7_168, 400),                # 既定で除外
        (".hidden.txt", 40, 400),                 # 名前の先頭が「.」→ 隠し扱いで除外
        ("隠し属性のファイル.bin", 512, 400),      # 属性で隠し（後で属性を付ける）
    ]

    # --- 2026_上期 ------------------------------------------------------
    spec += [
        ("2026_上期/見積書_第一版.xlsx", 24_576, 190),
        ("2026_上期/見積書_最終版.xlsx", 26_112, 150),
        ("2026_上期/請求書_0425.xlsx", 19_456, 120),
        ("2026_上期/社内メモ.txt", 900, 118),
        ("2026_上期/契約関係/契約書_日本語の長い名前のファイル_最終版_確認済み.pdf", 1_240_576, 96),
        ("2026_上期/契約関係/覚書.docx", 88_064, 95),
        ("2026_上期/契約関係/旧バージョン/契約書_2024.docx", 92_160, 610),
        ("2026_上期/契約関係/旧バージョン/メモ.txt", 240, 610),
        ("2026_上期/契約関係/旧バージョン/深掘り/最深のファイル.xlsx", 15_360, 590),
        ("2026_上期/契約関係/旧バージョン/深掘り/補足資料.png", 344_064, 585),
        ("2026_上期/契約関係/旧バージョン/深掘り/さらに深い/置き忘れ.txt", 128, 700),
    ]
    for i in range(1, 9):
        spec.append((f"2026_上期/画像/写真 {i:02d}.png", 120_000 + i * 15_360, 140 - i * 3))
    for i in range(1, 7):
        spec.append((f"2026_上期/データ/売上_2026-0{i}.csv", 8_192 + i * 1_024, 160 - i * 8))

    # --- 2026_下期 ------------------------------------------------------
    spec += [
        ("2026_下期/見積書_10月.xlsx", 25_600, 45),
        ("2026_下期/請求書_11月.xlsx", 21_504, 12),
        ("2026_下期/議事録_第1回.docx", 45_056, 60),
        ("2026_下期/議事録_第2回.docx", 47_104, 30),
        ("2026_下期/議事録_第3回.docx", 43_008, 9),
    ]
    for i in range(1, 6):
        spec.append((f"2026_下期/資料/スライド_{i:02d}.pdf", 480_000 + i * 22_000, 55 - i * 6))
    for i in range(1, 5):
        spec.append((f"2026_下期/添付/資料一式_その{i}.zip", 2_097_152 + i * 65_536, 50 - i * 5))
    spec += [
        ("2026_下期/下書き/一時ファイル.tmp", 4_096, 20),
        ("2026_下期/下書き/~$一時ファイル.xlsx", 512, 20),
        ("2026_下期/下書き/作業中.docx", 32_768, 18),
    ]

    # --- ドキュメント（4〜5 階層） --------------------------------------
    spec += [
        ("ドキュメント/手順書/手順書_入門.docx", 156_672, 210),
        ("ドキュメント/手順書/手順書_運用.docx", 202_752, 180),
        ("ドキュメント/手順書/図/フロー図 01.png", 262_144, 175),
        ("ドキュメント/手順書/図/フロー図 02.png", 268_288, 172),
        ("ドキュメント/手順書/図/旧版/フロー図_旧.png", 240_000, 520),
        ("ドキュメント/手順書/図/旧版/さらに古い/フロー図_2023.png", 220_000, 900),
        ("ドキュメント/一覧/機器一覧.xlsx", 66_560, 75),
        ("ドキュメント/一覧/アカウント一覧.csv", 12_288, 70),
        ("ドキュメント/一覧/棚卸し_2025.xlsx", 74_752, 330),
        ("ドキュメント/契約/秘密保持契約書.pdf", 640_000, 420),
        ("ドキュメント/契約/注文書_2026.pdf", 380_000, 100),
        ("ドキュメント/契約/契約更新の可否.txt", 1_600, 96),
    ]

    # --- 写真（スペース入り・日本語） ------------------------------------
    for i in range(1, 13):
        spec.append((f"写真/2026 旅行/IMG_{i:03d}.png", 900_000 + i * 40_000, 40 - i))
    for i in range(1, 5):
        spec.append((f"写真/2026 旅行/現像/現像_{i:02d}.png", 1_600_000 + i * 90_000, 38 - i))
    spec += [
        ("写真/2026 旅行/選定メモ.txt", 700, 36),
        ("写真/2025 旅行/IMG_9001.png", 1_100_000, 480),
        ("写真/2025 旅行/IMG_9002.png", 1_150_000, 478),
        ("写真/2025 旅行/アルバム.zip", 3_145_728, 470),
    ]

    # --- 既定で除外されるフォルダ ---------------------------------------
    spec += [
        (".git/config", 230, 300),
        (".git/HEAD", 23, 300),
        (".git/objects/ab/cdef1234", 512, 300),
        ("node_modules/left-pad/index.js", 1_024, 300),
        ("node_modules/left-pad/package.json", 640, 300),
        ("node_modules/.bin/tsc", 2_048, 300),
    ]
    return spec


HIDDEN_ATTR_FILES = ("隠し属性のファイル.bin",)
_HIDDEN = 0x2


def set_hidden_attr(path: str) -> bool:
    """Windows で隠し属性を付ける（失敗してもエラーにしない）。"""
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        return bool(ctypes.windll.kernel32.SetFileAttributesW(path, _HIDDEN))
    except Exception:
        return False


def main() -> int:
    spec = build_spec()
    now = time.time()
    created = 0
    skipped = 0
    hidden_done = 0

    for rel, size, days in spec:
        path = os.path.join(SAMPLE_ROOT, rel.replace("/", os.sep))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if os.path.exists(path):
            skipped += 1
        else:
            with open(path, "wb") as fh:
                fh.write(bytes([(sum(map(ord, rel)) & 0xFF)]) * size)
            created += 1
        stamp = now - days * DAY
        try:
            os.utime(path, (stamp, stamp))
        except OSError:
            pass
        if os.path.basename(rel) in HIDDEN_ATTR_FILES and set_hidden_attr(path):
            hidden_done += 1

    # 実測値を数える（生成したファイルだけでなく、いま在るもの全体）
    files = 0
    folders = 0
    total = 0
    deepest = 0
    for dirpath, dirnames, filenames in os.walk(SAMPLE_ROOT):
        depth = len(os.path.relpath(dirpath, SAMPLE_ROOT).split(os.sep)) if dirpath != SAMPLE_ROOT else 0
        folders += len(dirnames)
        files += len(filenames)
        for name in filenames:
            try:
                total += os.stat(os.path.join(dirpath, name)).st_size
            except OSError:
                pass
        if depth > deepest:
            deepest = depth

    print(f"サンプル生成: {SAMPLE_ROOT}")
    print(f"  新規作成 {created} 件 / 既存スキップ {skipped} 件")
    print(f"  ファイル数: {files}")
    print(f"  フォルダ数: {folders}")
    print(f"  合計サイズ: {total:,} バイト（{total / (1024 * 1024):.2f} MB）")
    print(f"  最深階層（フォルダ）: {deepest}")
    print(f"  隠し属性を付けたファイル: {hidden_done} 件")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
