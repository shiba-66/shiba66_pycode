# -*- coding: utf-8 -*-
"""
ドラッグ&ドロップ / ファイル選択ダイアログで1ファイルを選ばせる
汎用ウィンドウ。他のスクリプトから import して使う。

使用例:
    from dnd_filepicker import pick_file

    def validate(path):
        if not path.lower().endswith(".xlsx"):
            return "xlsxファイルを選択してください。"
        return None  # 問題なければNoneを返す

    path = pick_file(
        filetypes=[("Excel files", "*.xlsx")],
        title="見積書ファイル選択",
        prompt="ここに見積書(.xlsx)を\nドラッグ&ドロップ",
        validate=validate,
    )
    if path:
        ...
"""
import tkinter as tk
from tkinter import filedialog, messagebox

from tkinterdnd2 import DND_FILES, TkinterDnD


def pick_file(
    filetypes=(("All files", "*.*"),),
    title="ファイル選択",
    prompt="ここにファイルを\nドラッグ&ドロップ",
    validate=None,
    geometry="420x220",
):
    """ウィンドウを開き、ドラッグ&ドロップまたは「ファイルを選択...」ボタンで
    選ばれたファイルパスを返す。ウィンドウを閉じた場合はNoneを返す。

    validate: callable(path) -> エラーメッセージ(str) または None。
        ファイルの種類チェックなどは呼び出し側でこの関数を実装する。
        Noneを返せば受理、文字列を返せばエラー表示して選び直しになる。
    """
    selected = {"path": None}

    root = TkinterDnD.Tk()
    root.title(title)
    root.geometry(geometry)
    root.resizable(False, False)

    def accept(path):
        path = path.strip("{}")
        if validate is not None:
            error = validate(path)
            if error:
                messagebox.showerror("エラー", error)
                return
        selected["path"] = path
        root.destroy()

    def on_drop(event):
        paths = root.tk.splitlist(event.data)
        if paths:
            accept(paths[0])

    def browse():
        path = filedialog.askopenfilename(title=title, filetypes=list(filetypes))
        if path:
            accept(path)

    drop_area = tk.Label(
        root,
        text=prompt,
        relief="ridge",
        bd=2,
        bg="#f5f5f5",
    )
    drop_area.pack(padx=20, pady=(20, 10), fill="both", expand=True)
    drop_area.drop_target_register(DND_FILES)
    drop_area.dnd_bind("<<Drop>>", on_drop)

    tk.Label(root, text="または").pack()
    tk.Button(root, text="ファイルを選択...", command=browse).pack(pady=(5, 20))

    root.mainloop()
    return selected["path"]


if __name__ == "__main__":
    result = pick_file()
    print("selected:", result)
