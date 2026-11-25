# -*- coding: utf-8 -*-
"""フォルダ目録 Excel 作成ツール（Tkinter GUI）。

画面構成は既存ツールと同じ思想: **上に入力、下に「プレビュー」→「実行」**。
走査と Excel 書き出しは別スレッドで動かし、GUI はフリーズしない（FR-11）。
中止ボタン、進捗（処理件数・現在のフォルダ）を備える。

起動: `起動.bat`（venv があればそれを使う）
"""

from __future__ import annotations

import os
import queue
import threading
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import catalog_writer
import scan_core
from scan_core import DEFAULT_EXCLUDE_NAMES, DEFAULT_MAX_FILES, ScanOptions, preview_text

APP_TITLE = "フォルダ目録 Excel 作成ツール"
WINDOW_SIZE = "1040x780"
MIN_SIZE = "900x640"


class CatalogApp:
    """単一画面のメインウィンドウ。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.messages: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.preview_result = None      # プレビューの走査結果（同じ実行中のみ保持・FR-15）
        self.last_paths: dict = {}      # 直近の実行で生成したファイル
        self._out_overridden = False    # 出力先を手で変えたかどうか

        root.title(APP_TITLE)
        root.geometry(WINDOW_SIZE)
        root.minsize(*[int(v) for v in MIN_SIZE.split("x")])

        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Big.TButton", font=("", 10, "bold"))
        style.configure("Head.TLabel", font=("", 10, "bold"))

        self._build_widgets()
        self._set_status("対象フォルダを選んでから「プレビュー」を押してください。")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_queue)

    # ------------------------------------------------------------------
    # 画面
    # ------------------------------------------------------------------

    def _build_widgets(self) -> None:
        pad = {"padx": 8, "pady": 4}
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True, padx=10, pady=10)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        # --- ① 対象・出力先 ------------------------------------------------
        box_in = ttk.LabelFrame(outer, text=" ① 対象フォルダと出力先 ")
        box_in.grid(row=0, column=0, sticky="ew", **pad)
        box_in.columnconfigure(1, weight=1)

        self.target_var = tk.StringVar()
        self.out_var = tk.StringVar()

        ttk.Label(box_in, text="対象フォルダ").grid(row=0, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(box_in, textvariable=self.target_var).grid(row=0, column=1, sticky="ew", padx=4, pady=3)
        ttk.Button(box_in, text="参照…", width=10, command=self._browse_target).grid(row=0, column=2, padx=4, pady=3)

        ttk.Label(box_in, text="出力先フォルダ").grid(row=1, column=0, sticky="w", padx=6, pady=3)
        ttk.Entry(box_in, textvariable=self.out_var).grid(row=1, column=1, sticky="ew", padx=4, pady=3)
        ttk.Button(box_in, text="参照…", width=10, command=self._browse_out).grid(row=1, column=2, padx=4, pady=3)
        ttk.Button(box_in, text="対象の親に設定", width=16, command=self._reset_out_to_parent).grid(
            row=1, column=3, padx=4, pady=3
        )

        self.target_var.trace_add("write", lambda *_: self._on_target_changed())
        self.out_var.trace_add("write", lambda *_: self._on_out_changed())

        # --- ② オプション --------------------------------------------------
        box_opt = ttk.LabelFrame(outer, text=" ② 除外と出力のオプション ")
        box_opt.grid(row=1, column=0, sticky="ew", **pad)
        box_opt.columnconfigure(0, weight=1)

        self.exclude_hidden = tk.BooleanVar(value=True)
        self.exclude_names = tk.BooleanVar(value=True)
        self.with_csv = tk.BooleanVar(value=False)   # CSV は既定オフ（FR-16）

        ttk.Checkbutton(
            box_opt,
            text="隠しファイル／システムファイルを除外する",
            variable=self.exclude_hidden,
        ).grid(row=0, column=0, sticky="w", padx=8, pady=2)
        ttk.Checkbutton(
            box_opt,
            text="既定の除外名を除外する: " + ", ".join(DEFAULT_EXCLUDE_NAMES),
            variable=self.exclude_names,
        ).grid(row=1, column=0, sticky="w", padx=8, pady=2)
        ttk.Checkbutton(
            box_opt,
            text="CSV も書き出す（UTF-8・BOM 付き／シート「一覧」と「フォルダ集計」）",
            variable=self.with_csv,
        ).grid(row=2, column=0, sticky="w", padx=8, pady=2)

        limit_box = ttk.Frame(box_opt)
        limit_box.grid(row=3, column=0, sticky="w", padx=8, pady=2)
        ttk.Label(limit_box, text="件数上限（安全装置・0 で無制限）").pack(side="left")
        self.max_files_var = tk.StringVar(value=str(DEFAULT_MAX_FILES))
        spin = ttk.Spinbox(limit_box, from_=0, to=100_000_000, increment=10_000,
                           width=12, textvariable=self.max_files_var)
        spin.pack(side="left", padx=6)

        # 設定を変えたらプレビュー結果は無効（実行前に見直してもらう）
        for var in (self.exclude_hidden, self.exclude_names, self.with_csv, self.max_files_var):
            var.trace_add("write", lambda *_: self._invalidate_preview())

        # --- ③ プレビュー → 実行 -------------------------------------------
        box_run = ttk.LabelFrame(outer, text=" ③ プレビューしてから実行 ")
        box_run.grid(row=2, column=0, sticky="ew", **pad)
        box_run.columnconfigure(4, weight=1)

        self.btn_preview = ttk.Button(box_run, text="プレビュー", style="Big.TButton",
                                      width=14, command=self.on_preview)
        self.btn_preview.grid(row=0, column=0, padx=6, pady=6)
        self.btn_run = ttk.Button(box_run, text="実行", style="Big.TButton", width=14,
                                  command=self.on_run, state="disabled")
        self.btn_run.grid(row=0, column=1, padx=6, pady=6)
        self.btn_cancel = ttk.Button(box_run, text="中止", width=10, command=self.on_cancel,
                                     state="disabled")
        self.btn_cancel.grid(row=0, column=2, padx=6, pady=6)
        self.btn_open = ttk.Button(box_run, text="生成ファイルを開く", width=18,
                                   command=self.on_open_last, state="disabled")
        self.btn_open.grid(row=0, column=3, padx=6, pady=6)
        self.progress = ttk.Progressbar(box_run, mode="indeterminate", length=220)
        self.progress.grid(row=0, column=4, sticky="ew", padx=8)

        self.status_var = tk.StringVar(value="")
        ttk.Label(outer, textvariable=self.status_var, style="Head.TLabel", anchor="w").grid(
            row=3, column=0, sticky="ew", padx=8
        )

        # --- ④ 結果表示 ----------------------------------------------------
        box_log = ttk.LabelFrame(outer, text=" ④ プレビュー結果・実行ログ ")
        box_log.grid(row=4, column=0, sticky="nsew", **pad)
        box_log.columnconfigure(0, weight=1)
        box_log.rowconfigure(0, weight=1)
        self.log = ScrolledText(box_log, wrap="word", height=18, font=("Meiryo UI", 10),
                                state="disabled")
        self.log.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

    # ------------------------------------------------------------------
    # 入力補助
    # ------------------------------------------------------------------

    def _invalidate_preview(self) -> None:
        """設定が変わったらプレビュー結果を破棄する（実行ボタンも伏せる）。"""
        if self.preview_result is not None:
            self.preview_result = None
            self.btn_run.configure(state="disabled")
            self._set_status("設定を変えました。もう一度「プレビュー」を押してください。")

    def _on_target_changed(self) -> None:
        self._invalidate_preview()
        if not self._out_overridden:
            self._set_parent_as_out()

    def _on_out_changed(self) -> None:
        self._out_overridden = True

    def _reset_out_to_parent(self) -> None:
        self._out_overridden = False
        self._set_parent_as_out()

    def _set_parent_as_out(self) -> None:
        target = self.target_var.get().strip().strip('"')
        if target and os.path.isdir(target):
            self.out_var.set(os.path.dirname(os.path.normpath(target)) or target)

    def _browse_target(self) -> None:
        chosen = filedialog.askdirectory(title="対象フォルダを選択", mustexist=True,
                                         initialdir=self.target_var.get() or None)
        if chosen:
            self.target_var.set(os.path.normpath(chosen))

    def _browse_out(self) -> None:
        chosen = filedialog.askdirectory(title="出力先フォルダを選択",
                                         initialdir=self.out_var.get() or self.target_var.get() or None)
        if chosen:
            self.out_var.set(os.path.normpath(chosen))

    def _read_options(self) -> ScanOptions | None:
        """画面から設定を読み取る。不正なら None（理由をステータスに出す）。"""
        raw = self.max_files_var.get().strip().replace(",", "")
        try:
            limit = int(raw)
        except ValueError:
            self._set_status("件数上限は数字で入力してください。")
            return None
        if limit < 0:
            self._set_status("件数上限は 0 以上の数字を入力してください。")
            return None
        return ScanOptions(
            exclude_hidden=self.exclude_hidden.get(),
            exclude_system=self.exclude_hidden.get(),
            exclude_default_names=self.exclude_names.get(),
            max_files=limit,
        )

    def _validate_paths(self) -> tuple | None:
        target = os.path.normpath(self.target_var.get().strip().strip('"'))
        if not target:
            self._set_status("対象フォルダを入力してください。")
            return None
        if not os.path.exists(target):
            self._set_status(f"対象フォルダが見つかりません: {target}")
            return None
        if not os.path.isdir(target):
            self._set_status(f"対象がフォルダではありません: {target}")
            return None
        out = self.out_var.get().strip().strip('"') or os.path.dirname(target)
        out = os.path.normpath(out)
        if not os.path.isdir(out):
            self._set_status(f"出力先フォルダが見つかりません: {out}")
            return None
        return target, out

    # ------------------------------------------------------------------
    # 実行
    # ------------------------------------------------------------------

    def on_preview(self) -> None:
        if self._busy():
            return
        paths = self._validate_paths()
        options = self._read_options()
        if paths is None or options is None:
            return
        target, _out = paths
        self.preview_result = None
        self.btn_run.configure(state="disabled")
        self._start_worker(mode="preview", target=target, out=_out, options=options)

    def on_run(self) -> None:
        if self._busy():
            return
        paths = self._validate_paths()
        options = self._read_options()
        if paths is None or options is None:
            return
        target, out = paths

        reuse = False
        if self.preview_result is not None:
            answer = messagebox.askyesno(
                "プレビュー結果を使う",
                "プレビュー時の走査結果をそのまま使いますか？\n\n"
                "はい: 再走査せずに Excel を作ります（速い）\n"
                "いいえ: もう一度走査してから Excel を作ります（最新）",
                parent=self.root,
            )
            reuse = answer
        self._start_worker(mode="run", target=target, out=out, options=options, reuse=reuse)

    def on_cancel(self) -> None:
        self.cancel_event.set()
        self._set_status("中止しています…（現在のファイルまでで区切ります）")

    def on_open_last(self) -> None:
        path = self.last_paths.get("xlsx")
        if not path or not os.path.exists(path):
            messagebox.showinfo("開けません", "生成されたファイルが見つかりません。", parent=self.root)
            return
        try:
            os.startfile(path)  # Windows で既定アプリ（Excel）を開く
        except (AttributeError, OSError) as exc:
            messagebox.showwarning("開けません", f"自動で開けませんでした: {exc}\n手動で開いてください。",
                                   parent=self.root)

    def _busy(self) -> bool:
        return self.worker is not None and self.worker.is_alive()

    def _start_worker(self, mode: str, target: str, out: str, options: ScanOptions,
                      reuse: bool = False) -> None:
        self.cancel_event.clear()
        for button in (self.btn_preview, self.btn_run, self.btn_open):
            button.configure(state="disabled")
        self.btn_cancel.configure(state="normal")
        self.progress.start(12)
        self.worker = threading.Thread(
            target=self._worker, name="catalog", daemon=True,
            kwargs={"mode": mode, "target": target, "out": out, "options": options,
                    "reuse": reuse, "with_csv": self.with_csv.get()},
        )
        self.worker.start()

    def _worker(self, mode: str, target: str, out: str, options: ScanOptions,
                reuse: bool, with_csv: bool) -> None:
        """別スレッドで走査＋書き出し（GUI は待たされない）。"""
        try:
            if mode == "run" and reuse and self.preview_result is not None:
                result = self.preview_result
                self.messages.put(("status", f"プレビュー結果を使用: {result.file_count:,} 件"))
            else:
                self.messages.put(("status", "走査中…"))
                result = scan_core.scan(
                    target,
                    options,
                    progress=lambda count, folder: self.messages.put(("progress", count, folder)),
                    should_cancel=self.cancel_event,
                )
            if result.cancelled:
                self.messages.put(("cancelled", result))
                return
            if mode == "preview":
                self.messages.put(("preview", result))
                return

            self.messages.put(("status", "Excel を書き出し中…"))
            info = catalog_writer.run_catalog(result, out, with_csv=with_csv)
            self.messages.put(("done", result, info))
        except Exception:
            self.messages.put(("error", traceback.format_exc()))

    # ------------------------------------------------------------------
    # 結果表示（メインスレッドでしかウィジェットを触らない）
    # ------------------------------------------------------------------

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, *payload = self.messages.get_nowait()
                getattr(self, f"_on_msg_{kind}")(*payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _on_msg_progress(self, count: int, folder: str) -> None:
        self._set_status(f"走査中… {count:,} 件 ／ {folder}")

    def _on_msg_status(self, text: str) -> None:
        self._set_status(text)

    def _on_msg_preview(self, result) -> None:
        self.preview_result = result
        self.progress.stop()
        self._release_buttons()
        self._append("=== プレビュー ===")
        self._append(preview_text(result))
        self._append("")
        if result.error_rows:
            self._append("エラー（先頭 10 件）")
            for row in result.error_rows[:10]:
                self._append(f"  [{row[1]}] {row[2]}")
            self._append("")
        self._set_status(
            f"プレビュー完了: {result.file_count:,} 件 ／ {result.total_bytes:,} バイト ／ "
            f"所要 {result.elapsed_seconds:.2f} 秒。問題なければ「実行」を押してください。"
        )
        self.btn_run.configure(state="normal")

    def _on_msg_done(self, result, info: dict) -> None:
        self.progress.stop()
        self._release_buttons()
        self.last_paths = {"xlsx": info["xlsx"]}
        self._append("=== 実行完了 ===")
        self._append(preview_text(result))
        self._append(f"Excel: {info['xlsx']}")
        for sheet, path in info["csv"].items():
            self._append(f"CSV({sheet}): {path}")
        self._append("")
        self._set_status(
            f"完了: {result.file_count:,} 件を書き出しました（エラー {result.error_count:,} 件）"
        )
        self.btn_open.configure(state="normal")
        if messagebox.askyesno("完了", f"{result.file_count:,} 件の目録を作りました。\n\nExcel を開きますか？",
                               parent=self.root):
            self.on_open_last()

    def _on_msg_cancelled(self, result) -> None:
        self.progress.stop()
        self._release_buttons()
        self._append("=== 中止しました ===")
        self._append(f"ここまでで {result.file_count:,} 件収集しています。実行すると途中までの目録になります。")
        self._append("")
        self._set_status("中止しました。")

    def _on_msg_error(self, detail: str) -> None:
        self.progress.stop()
        self._release_buttons()
        self._append("=== エラーで停止しました ===")
        self._append(detail.strip())
        self._append("")
        self._set_status("エラーで停止しました。ログを確認してください。")
        messagebox.showerror("エラー", "処理中にエラーが発生しました。\nログに詳細を出しました。", parent=self.root)

    def _release_buttons(self) -> None:
        self.btn_preview.configure(state="normal")
        self.btn_cancel.configure(state="disabled")
        if self.preview_result is not None:
            self.btn_run.configure(state="normal")

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _append(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _on_close(self) -> None:
        self.cancel_event.set()
        if self._busy():
            self.worker.join(timeout=2.0)
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    CatalogApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
