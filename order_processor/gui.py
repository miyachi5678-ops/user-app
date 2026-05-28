"""注文処理システム - GUIモード

ダブルクリックで起動するウィンドウ版。
マウス操作のみで使えます。
"""
import tkinter as tk
from tkinter import filedialog, scrolledtext
import threading
import queue
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

from config import DB_PATH
from src.db import initialize_db
from main import cmd_process, cmd_update_master

_VERSION = "1.0"


# ── スレッドセーフな選択ダイアログ ────────────────────────────────────

class GuiAskFunc:
    """
    ワーカースレッドから呼べる選択ダイアログ。

    GUIは必ずメインスレッドで操作する必要があるため、
    キューを使ってスレッド間でやり取りする。
    """

    def __init__(self, root: tk.Tk):
        self._root = root
        self._req: queue.Queue = queue.Queue()
        self._res: queue.Queue = queue.Queue()
        self._poll()

    def _poll(self):
        try:
            label1, label2, context = self._req.get_nowait()
            self._show(label1, label2, context)
        except queue.Empty:
            pass
        self._root.after(100, self._poll)

    def _show(self, label1: str, label2: str, context: str):
        dialog = tk.Toplevel(self._root)
        dialog.title("確認")
        dialog.grab_set()
        dialog.focus_set()
        dialog.resizable(False, False)

        # 背景をやや明るめに
        dialog.configure(bg="#F5F5F5")

        if context:
            tk.Label(
                dialog, text=context,
                wraplength=320, justify="left",
                bg="#F5F5F5", font=("", 10),
                padx=20, pady=12,
            ).pack()

        frame = tk.Frame(dialog, bg="#F5F5F5", pady=8)
        frame.pack(padx=20, pady=(0, 16))

        result = tk.StringVar(value="1")

        def choose(v):
            result.set(v)
            dialog.destroy()

        tk.Button(
            frame, text=label1, width=30, height=2,
            font=("", 10), bg="#2E75B6", fg="white",
            activebackground="#1F5080", cursor="hand2",
            relief="flat",
            command=lambda: choose("1"),
        ).pack(pady=4)

        tk.Button(
            frame, text=label2, width=30, height=2,
            font=("", 10), bg="#F0F0F0",
            activebackground="#DDDDDD", cursor="hand2",
            relief="flat",
            command=lambda: choose("2"),
        ).pack(pady=4)

        # 画面中央に表示
        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        x = self._root.winfo_x() + (self._root.winfo_width()  - w) // 2
        y = self._root.winfo_y() + (self._root.winfo_height() - h) // 2
        dialog.geometry(f"+{x}+{y}")

        self._root.wait_window(dialog)
        self._res.put(result.get())

    def __call__(self, label1: str, label2: str, context: str = "") -> str:
        self._req.put((label1, label2, context))
        return self._res.get()   # ワーカースレッドはここで一時停止して待つ


# ── stdout → テキストウィジェット ────────────────────────────────────

class _TextRedirector:
    def __init__(self, widget: scrolledtext.ScrolledText, root: tk.Tk):
        self._w = widget
        self._root = root

    def write(self, text: str):
        self._root.after(0, self._append, text)

    def _append(self, text: str):
        self._w.config(state="normal")
        self._w.insert("end", text)
        self._w.see("end")
        self._w.config(state="disabled")

    def flush(self):
        pass


# ── メインウィンドウ ──────────────────────────────────────────────────

class App:
    def __init__(self):
        self._root = tk.Tk()
        self._root.title(f"注文処理システム  ver.{_VERSION}")
        self._root.geometry("560x500")
        self._root.resizable(False, True)
        self._root.configure(bg="#FFFFFF")

        self._build_ui()
        self._ask_func    = GuiAskFunc(self._root)
        self._redirector  = _TextRedirector(self._log, self._root)

        initialize_db(DB_PATH)

    # ── UI組み立て ──────────────────────────────────────────────

    def _build_ui(self):
        # ヘッダー
        header = tk.Frame(self._root, bg="#1F4E79", pady=14)
        header.pack(fill="x")
        tk.Label(
            header,
            text=f"注文処理システム  ver.{_VERSION}",
            font=("", 14, "bold"), fg="white", bg="#1F4E79",
        ).pack()

        # ボタンエリア
        btn_area = tk.Frame(self._root, bg="#FFFFFF", pady=18)
        btn_area.pack()

        self._btn_order = tk.Button(
            btn_area,
            text="注文書を処理する",
            font=("", 12), width=26, height=2,
            bg="#2E75B6", fg="white",
            activebackground="#1F5080",
            cursor="hand2", relief="flat",
            command=self._on_process_order,
        )
        self._btn_order.pack(pady=4)
        tk.Label(
            btn_area,
            text="→ クリックしてファイルを選ぶとレポートと仕分けリストを作ります",
            fg="#666666", font=("", 9), bg="#FFFFFF",
        ).pack()

        tk.Frame(btn_area, height=12, bg="#FFFFFF").pack()

        self._btn_master = tk.Button(
            btn_area,
            text="構成一覧を更新する",
            font=("", 12), width=26, height=2,
            bg="#538135", fg="white",
            activebackground="#3A5B24",
            cursor="hand2", relief="flat",
            command=self._on_update_master,
        )
        self._btn_master.pack(pady=4)
        tk.Label(
            btn_area,
            text="→ クリックしてファイルを選ぶと差分を確認しながら更新します",
            fg="#666666", font=("", 9), bg="#FFFFFF",
        ).pack()

        # ログエリア
        log_header = tk.Frame(self._root, bg="#EEEEEE", pady=4)
        log_header.pack(fill="x")
        tk.Label(
            log_header, text="処理状況",
            font=("", 9, "bold"), bg="#EEEEEE", padx=10,
        ).pack(anchor="w")

        self._log = scrolledtext.ScrolledText(
            self._root,
            state="disabled",
            font=("Consolas", 9),
            bg="#FAFAFA", relief="flat",
        )
        self._log.pack(fill="both", expand=True, padx=0, pady=0)

    # ── ボタン操作 ──────────────────────────────────────────────

    def _on_process_order(self):
        path = filedialog.askopenfilename(
            title="注文書 Excel を選んでください",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        self._run(cmd_process, path, ask_func=self._ask_func)

    def _on_update_master(self):
        path = filedialog.askopenfilename(
            title="構成一覧 Excel を選んでください",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        self._run(cmd_update_master, path, ask_func=self._ask_func)

    # ── スレッド実行 ──────────────────────────────────────────────

    def _run(self, func, *args, **kwargs):
        self._set_buttons(False)
        self._clear_log()

        sys.stdout = self._redirector

        def worker():
            try:
                func(*args, **kwargs)
                self._root.after(0, self._on_done)
            except Exception as e:
                self._root.after(0, self._on_error, str(e))
            finally:
                sys.stdout = sys.__stdout__

        threading.Thread(target=worker, daemon=True).start()

    def _on_done(self):
        self._set_buttons(True)
        self._log.config(state="normal")
        self._log.insert("end", "\n── 完了しました ──\n")
        self._log.config(state="disabled")

    def _on_error(self, msg: str):
        self._set_buttons(True)
        self._log.config(state="normal")
        self._log.insert("end", f"\n⚠ エラー: {msg}\n")
        self._log.config(state="disabled")

    def _set_buttons(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self._btn_order.config(state=state)
        self._btn_master.config(state=state)

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    # ── 起動 ──────────────────────────────────────────────────

    def run(self):
        self._root.mainloop()


if __name__ == "__main__":
    App().run()
