import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from common import convert_any_email

class ProgressDialog:
    def __init__(self, files):
        self.files = files

        # --- ウィンドウ ---
        self.win = tk.Tk()
        self.win.title("変換中…")
        self.win.geometry("420x180")
        self.win.configure(bg="white")
        self.win.resizable(False, False)

        # 閉じるボタン無効（処理中に閉じられないように）
        self.win.protocol("WM_DELETE_WINDOW", lambda: None)

        # --- Apple風スタイル ---
        style = ttk.Style()
        style.theme_use("default")

        # 小さめの Apple Primary ボタン（完了ボタン用）
        style.configure(
            "ApplePrimarySmall.TButton",
            font=("Meiryo", 10),
            padding=4,
            foreground="white",
            background="#007AFF",
            borderwidth=0
        )
        style.map(
            "ApplePrimarySmall.TButton",
            background=[("active", "#005FCC")]
        )

        # Apple風の細い進捗バー
        style.configure(
            "Apple.Horizontal.TProgressbar",
            troughcolor="white",
            bordercolor="white",
            background="#4A90E2",
            lightcolor="#4A90E2",
            darkcolor="#4A90E2",
            thickness=4
        )

        # --- ラベル ---
        tk.Label(
            self.win,
            text="メールを変換しています…",
            font=("Meiryo", 12),
            bg="white"
        ).pack(pady=(20, 10))

        # --- 進捗バー ---
        self.progress = ttk.Progressbar(
            self.win,
            orient="horizontal",
            mode="determinate",
            style="Apple.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", padx=40, pady=(0, 20))

        self.progress["value"] = 0
        self.progress["maximum"] = 100

        # --- ボタン（最初は処理中で押せない） ---
        self.button = ttk.Button(
            self.win,
            text="処理中…",
            width=10,
            state="disabled"
        )
        self.button.pack()

        # --- 変換処理を開始 ---
        self.win.after(100, self.run_conversion)
        self.win.mainloop()

    # ----------------------------------------
    # 変換処理（D&D専用）
    # ----------------------------------------
    def run_conversion(self):
        total = len(self.files)
        self.progress["value"] = 0
        self.progress["maximum"] = total

        for index, p in enumerate(self.files, start=1):
            try:
                out_dir = os.path.dirname(p)
                convert_any_email(
                    p,
                    output_dir=out_dir,
                    save_external_images=False
                )
            except Exception as e:
                # D&D時は messagebox を出さずログだけ
                print(f"Error converting {p}: {e}")

            # 進捗更新
            self.progress["value"] = index
            self.win.update_idletasks()

        # 完了時に満タンにする
        self.progress["value"] = total
        self.win.update_idletasks()

        # 完了ボタンに切り替え
        self.button.config(
            text="完了",
            state="normal",
            style="ApplePrimarySmall.TButton",
            command=self.win.destroy
        )

class EmlConverterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("EML / MSG → HTML 変換ツール")
        self.root.geometry("620x650")
        self.root.configure(bg="white")

        self.selected_files = []
        self.output_dir = None

        # ------------------------------
        # スタイル設定
        # ------------------------------
        style = ttk.Style()
        style.theme_use("default")

        # プライマリーボタン
        style.configure(
            "ApplePrimary.TButton",
            font=("Meiryo", 12),
            padding=10,
            foreground="white",
            background="#007AFF",
            borderwidth=0
        )
        style.map(
            "ApplePrimary.TButton",
            background=[("active", "#005FCC")]
        )

        # 細い進捗バー
        style.configure(
            "Apple.Horizontal.TProgressbar",
            troughcolor="white",
            bordercolor="white",
            background="#4A90E2",
            lightcolor="#4A90E2",
            darkcolor="#4A90E2",
            thickness=4
        )

        # ------------------------------
        # 上部余白（説明文は削除）
        # ------------------------------
        tk.Frame(self.root, height=20, bg="white").pack()

        # ------------------------------
        # ファイル選択セクション
        # ------------------------------
        section_top = tk.Frame(self.root, bg="#F7F7F7")
        section_top.pack(fill="x", padx=30, pady=(0, 20))

        top_btn_frame = tk.Frame(section_top, bg="#F7F7F7")
        top_btn_frame.pack(pady=15)

        ttk.Button(
            top_btn_frame,
            text="ファイルを選択",
            width=18,
            command=self.select_files
        ).grid(row=0, column=0, padx=10)

        ttk.Button(
            top_btn_frame,
            text="出力先フォルダ",
            width=18,
            command=self.select_output_folder
        ).grid(row=0, column=1, padx=10)

        self.output_label = tk.Label(
            section_top,
            text="出力先：元ファイルと同じフォルダ（デフォルト）",
            font=("Meiryo", 10),
            bg="#F7F7F7",
            anchor="w",
            padx=10
        )
        self.output_label.pack(fill="x", padx=20, pady=(5, 15))

        # 外部画像保存チェック（デフォルト OFF）
        self.save_images_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            section_top,
            text="外部画像を保存する（http/https）",
            variable=self.save_images_var
        ).pack(pady=(0, 10))

        # ------------------------------
        # リストセクション
        # ------------------------------
        section_list = tk.Frame(self.root, bg="white")
        section_list.pack(fill="both", expand=True, padx=30, pady=(0, 20))

        list_frame = tk.Frame(section_list, bg="white")
        list_frame.pack()

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_frame,
            width=70,
            height=12,  # ← 高さを調整
            selectmode="extended",
            yscrollcommand=scrollbar.set,
            relief="flat",
            highlightthickness=1,
            highlightcolor="#DDD",
            highlightbackground="#DDD",
            font=("Meiryo", 10)
        )
        self.listbox.pack(side="left")

        scrollbar.config(command=self.listbox.yview)
        self.listbox.bind("<Delete>", self.delete_selected_event)
        self.listbox.bind("<<ListboxSelect>>", lambda e: self.update_select_toggle_button())

        # Listbox 下の操作ボタン（すべて選択 / 選択解除）
        select_toggle_frame = tk.Frame(section_list, bg="white")
        select_toggle_frame.pack(anchor="w", pady=(5, 0))
        
        self.select_toggle_btn = ttk.Button(
            select_toggle_frame,
            text="すべて選択",
            width=12,
            command=self.toggle_select_all
        )
        self.select_toggle_btn.pack(padx=5)

        # ------------------------------
        # 操作セクション
        # ------------------------------
        section_bottom = tk.Frame(self.root, bg="#F7F7F7")
        section_bottom.pack(fill="x", padx=30, pady=(0, 20))

        bottom_frame = tk.Frame(section_bottom, bg="#F7F7F7")
        bottom_frame.pack(fill="x", pady=15)

        # 左
        left_column = tk.Frame(bottom_frame, bg="#F7F7F7")
        left_column.pack(side="left")

        ttk.Button(
            left_column,
            text="選択項目を削除",
            width=15,
            command=self.delete_selected
        ).pack(pady=4)

        ttk.Button(
            left_column,
            text="クリア",
            width=12,
            command=self.clear_listbox
        ).pack(pady=4)

        # 中央（主ボタン）
        center_frame = tk.Frame(bottom_frame, bg="#F7F7F7")
        center_frame.pack(side="left", expand=True)

        ttk.Button(
            center_frame,
            text="変換する",
            width=18,
            style="ApplePrimary.TButton",
            command=self.convert_files
        ).pack(pady=5)

        # 右
        ttk.Button(
            bottom_frame,
            text="閉じる",
            width=12,
            command=self.on_close
        ).pack(side="right")

        # ------------------------------
        # 細い進捗バー
        # ------------------------------
        progress_frame = tk.Frame(self.root, bg="white")
        progress_frame.pack(fill="x", pady=(0, 25))

        self.progress = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            style="Apple.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", padx=80)

        # 起動時は空にする
        self.progress["value"] = 0
        self.progress["maximum"] = 100

    # ------------------------------
    # GUI メソッド
    # ------------------------------

    def select_output_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir = path
            self.output_label.config(text=f"出力先：{path}")
        else:
            self.output_dir = None
            self.output_label.config(text="出力先：元ファイルと同じフォルダ（デフォルト）")

    def select_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[
                ("Email files", "*.eml *.msg"),
                ("EML files", "*.eml"),
                ("MSG files", "*.msg"),
                ("All files", "*.*")
            ]
        )
        if not paths:
            return

        for p in paths:
            if p not in self.selected_files:
                self.selected_files.append(p)
                self.listbox.insert(tk.END, p)

    def delete_selected_event(self, event):
        self.delete_selected()

    def delete_selected(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "削除する項目を選択してください。")
            return

        for index in reversed(selection):
            del self.selected_files[index]
            self.listbox.delete(index)

    def clear_listbox(self):
        self.listbox.delete(0, tk.END)
        self.selected_files.clear()

    def toggle_select_all(self):
        # 何も選択されていない → すべて選択
        if len(self.listbox.curselection()) == 0:
            self.listbox.select_set(0, tk.END)
        else:
            # 1つ以上選択されている → 選択解除
            self.listbox.select_clear(0, tk.END)
    
        # ボタン表示を更新
        self.update_select_toggle_button()

    def update_select_toggle_button(self):
        if len(self.listbox.curselection()) == 0:
            self.select_toggle_btn.config(text="すべて選択")
        else:
            self.select_toggle_btn.config(text="選択解除")

    def on_close(self):
        self.root.destroy()

    # ------------------------------
    # 進捗バー付き変換処理
    # ------------------------------
    def convert_files(self):
        if not self.selected_files:
            messagebox.showwarning("警告", "ファイルが選択されていません。")
            return

        overwrite = messagebox.askyesno(
            "確認",
            f"{len(self.selected_files)} 件のファイルを変換します。\n"
            "既存の HTML がある場合は上書きされます。\n\n続行しますか？"
        )
        if not overwrite:
            return

        total = len(self.selected_files)
        self.progress["value"] = 0
        self.progress["maximum"] = total
        self.root.update_idletasks()

        for index, p in enumerate(self.selected_files, start=1):
            try:
                convert_any_email(
                    p,
                    output_dir=self.output_dir,
                    save_external_images=self.save_images_var.get()
                )
            except Exception as e:
                messagebox.showerror(
                    "エラー",
                    f"{os.path.basename(p)} の変換中にエラーが発生しました:\n{e}"
                )

            self.progress["value"] = index
            self.root.update_idletasks()

        messagebox.showinfo("完了", f"{total} 件のメールを変換しました。")

    def run(self):
        self.root.mainloop()

# ------------------------------
# D&D 実行時の処理
# ------------------------------
if len(sys.argv) > 1:
    targets = [
        p for p in sys.argv[1:]
        if os.path.isfile(p) and (p.lower().endswith(".eml") or p.lower().endswith(".msg"))
    ]

    if targets:
        ProgressDialog(targets)

    sys.exit(0)


# ------------------------------
# 通常起動（GUI）
# ------------------------------
if __name__ == "__main__":
    gui = EmlConverterGUI()
    gui.run()