import os
import re
import sys
from email import policy
from email.parser import BytesParser
from html import escape as html_escape
import urllib.parse
import urllib.request
import tkinter as tk
from tkinter import filedialog, messagebox

# ------------------------------
# EML 読み込み
# ------------------------------
def read_eml(path: str):
    with open(path, "rb") as f:
        return f.read()

# ------------------------------
# charset に応じてデコード
# ------------------------------
def decode_bytes(data: bytes, charset_hint: str | None):
    candidates = []
    if charset_hint:
        candidates.append(charset_hint.lower())
    candidates += ["utf-8", "iso-2022-jp", "cp932"]

    tried = set()
    for cs in candidates:
        if cs in tried:
            continue
        tried.add(cs)
        try:
            return data.decode(cs)
        except Exception:
            pass

    return data.decode("utf-8", errors="replace")

# ------------------------------
# HTML / TEXT の本文を選ぶ
# ------------------------------
def pick_best_part(msg):
    html_part = None
    text_part = None

    if msg.is_multipart():
        for part in msg.walk():
            disp = (part.get_content_disposition() or "").lower()
            if disp == "attachment":
                continue
            ctype = part.get_content_type()
            if ctype == "text/html" and html_part is None:
                html_part = part
            elif ctype == "text/plain" and text_part is None:
                text_part = part
    else:
        ctype = msg.get_content_type()
        if ctype == "text/html":
            html_part = msg
        elif ctype == "text/plain":
            text_part = msg

    return html_part or text_part

# ------------------------------
# 添付ファイル保存
# ------------------------------
def extract_attachments(msg, base_folder: str):
    attachments = []

    for part in msg.walk():
        disp = (part.get_content_disposition() or "").lower()
        if disp != "attachment":
            continue
        filename = part.get_filename()
        if not filename:
            continue
        payload = part.get_payload(decode=True) or b""
        attachments.append((filename, payload))

    if not attachments:
        return 0

    os.makedirs(base_folder, exist_ok=True)

    for filename, payload in attachments:
        out_path = os.path.join(base_folder, filename)
        with open(out_path, "wb") as f:
            f.write(payload)

    return len(attachments)

# ------------------------------
# URL / メールアドレス自動リンク化
# ------------------------------
def autolink_plus(text: str) -> str:
    text = html_escape(text)

    def link(pattern, repl, s):
        return re.sub(pattern, repl, s, flags=re.IGNORECASE)

    text = link(r"(https?://[^\s<>]+)", r'<a href="\1">\1</a>', text)
    text = link(r"(www\.[^\s<>]+)", r'<a href="http://\1">\1</a>', text)
    text = link(r"([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})",
                r'<a href="mailto:\1">\1</a>', text)

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\n", "<br>\n")

    return text

# ------------------------------
# 外部画像保存（http/https）
# ------------------------------
def ensure_unique(path: str) -> str:
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    i = 1
    while True:
        new = f"{root}({i}){ext}"
        if not os.path.exists(new):
            return new
        i += 1

def download_external_images(body_html: str, out_dir: str, subfolder: str) -> str:
    img_dir = os.path.join(out_dir, subfolder)
    os.makedirs(img_dir, exist_ok=True)

    urls = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', body_html)

    for src in urls:
        if not (src.startswith("http://") or src.startswith("https://")):
            continue

        parsed = urllib.parse.urlparse(src)
        filename = os.path.basename(parsed.path)
        if not filename:
            continue

        save_path = ensure_unique(os.path.join(img_dir, filename))

        try:
            urllib.request.urlretrieve(src, save_path)
        except Exception:
            continue

        local_ref = f"{subfolder}/{os.path.basename(save_path)}"
        body_html = body_html.replace(src, local_ref)

    return body_html

# ------------------------------
# EML → HTML 変換（v1.2）
# ------------------------------
def eml_to_html(eml_path: str, output_dir: str | None = None, save_external_images: bool = False):
    raw = read_eml(eml_path)
    msg = BytesParser(policy=policy.default).parsebytes(raw)

    folder = output_dir if output_dir else os.path.dirname(eml_path)
    base = os.path.splitext(os.path.basename(eml_path))[0]
    html_out = os.path.join(folder, base + ".html")
    attach_folder = os.path.join(folder, base + "_files")

    extract_attachments(msg, attach_folder)

    part = pick_best_part(msg)
    if part is None:
        body_html = "<html><body><pre>本文が見つかりませんでした。</pre></body></html>"
    else:
        payload = part.get_payload(decode=True) or b""
        charset = part.get_content_charset()
        text = decode_bytes(payload, charset)

        if part.get_content_type() == "text/html":
            body_html = text

            if "<meta" in body_html.lower():
                body_html = re.sub(
                    r'<meta[^>]*charset=[^>]*>',
                    '<meta charset="UTF-8">',
                    body_html,
                    flags=re.IGNORECASE
                )
            else:
                if "<head>" in body_html.lower():
                    body_html = re.sub(
                        r'(?i)<head>',
                        '<head>\n<meta charset="UTF-8">',
                        body_html
                    )
                else:
                    body_html = "<head><meta charset=\"UTF-8\"></head>\n" + body_html

        else:
            body_html = (
                "<html><head><meta charset=\"UTF-8\"></head>"
                "<body>" + autolink_plus(text) + "</body></html>"
            )

    # ★ 外部画像保存（GUI 任意 / D&D 常に有効）
    if save_external_images:
        body_html = download_external_images(body_html, folder, base + "_files")

    with open(html_out, "w", encoding="utf-8", newline="\n") as f:
        f.write(body_html)

    return html_out

# ------------------------------
# GUI
# ------------------------------
class EmlConverterGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("EML → HTML 変換ツール")
        self.root.geometry("600x600")

        self.selected_files = []
        self.output_dir = None

        label = tk.Label(
            self.root,
            text="EML ファイルを選択して一覧に追加できます。\n"
                 "GUI 上でのドラッグ＆ドロップには対応していません。",
            font=("Meiryo", 11)
        )
        label.pack(pady=10)

        # ★ 外部画像保存チェックボックス
        self.save_images_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            self.root,
            text="外部画像を保存する（http/https）",
            variable=self.save_images_var,
            font=("Meiryo", 10)
        ).pack()

        # --- 以下 GUI は v1.1 のまま ---
        top_btn_frame = tk.Frame(self.root)
        top_btn_frame.pack(pady=5)

        tk.Button(
            top_btn_frame,
            text="ファイルを選択",
            font=("Meiryo", 10),
            width=15,
            command=self.select_files
        ).grid(row=0, column=0, padx=5)

        tk.Button(
            top_btn_frame,
            text="出力先フォルダ",
            font=("Meiryo", 10),
            width=15,
            command=self.select_output_folder
        ).grid(row=0, column=1, padx=5)

        self.output_label = tk.Label(
            self.root,
            text="出力先：元ファイルと同じフォルダ（デフォルト）",
            font=("Meiryo", 10),
            relief="groove",
            anchor="w",
            padx=10
        )
        self.output_label.pack(fill="x", padx=20, pady=5)

        select_frame = tk.Frame(self.root)
        select_frame.pack(pady=5)

        tk.Button(
            select_frame,
            text="すべて選択",
            font=("Meiryo", 10),
            width=15,
            command=self.select_all
        ).grid(row=0, column=0, padx=5)
        
        tk.Button(
            select_frame,
            text="選択を解除",
            font=("Meiryo", 10),
            width=15,
            command=self.clear_selection
        ).grid(row=0, column=1, padx=5)
        
        list_frame = tk.Frame(self.root)
        list_frame.pack(pady=10)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_frame,
            width=75,
            height=14,
            selectmode="extended",
            yscrollcommand=scrollbar.set
        )
        self.listbox.pack(side="left")

        scrollbar.config(command=self.listbox.yview)

        self.listbox.bind("<Delete>", self.delete_selected_event)

        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=20, pady=10)
        
        # 左側の縦並びフレーム
        left_column = tk.Frame(bottom_frame)
        left_column.pack(side="left")
        
        # 選択項目を削除
        tk.Button(
            left_column,
            text="選択項目を削除",
            font=("Meiryo", 10),
            width=18,
            command=self.delete_selected
        ).pack(pady=2)
        
        # ▼ 追加：クリアボタン（Listbox を空にする）
        tk.Button(
            left_column,
            text="クリア",
            font=("Meiryo", 9),
            width=12,
            command=self.clear_listbox
        ).pack(pady=2)
        
        # 右側：変換する
        tk.Button(
            bottom_frame,
            text="変換する",
            font=("Meiryo", 12),
            width=18,
            command=self.convert_files
        ).pack(side="right")

    # --- GUI のメソッドは v1.1 のまま（省略せず全部残す） ---
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
            filetypes=[("EML files", "*.eml"), ("All files", "*.*")]
        )
        if not paths:
            return

        for p in paths:
            if p not in self.selected_files:
                self.selected_files.append(p)
                self.listbox.insert(tk.END, p)

    def select_all(self):
        self.listbox.select_set(0, tk.END)

    def clear_selection(self):
        self.listbox.select_clear(0, tk.END)

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

        for p in self.selected_files:
            if self.output_dir:
                eml_to_html(
                    p,
                    output_dir=self.output_dir,
                    save_external_images=self.save_images_var.get()
                )
            else:
                eml_to_html(
                    p,
                    save_external_images=self.save_images_var.get()
                )

        messagebox.showinfo("完了", f"{len(self.selected_files)} 件の EML を変換しました。")

    def run(self):
        self.root.mainloop()

# ------------------------------
# exe に D&D された場合（常に画像保存 ON）
# ------------------------------
if len(sys.argv) > 1:
    targets = [p for p in sys.argv[1:] if os.path.isfile(p) and p.lower().endswith(".eml")]

    if targets:
        root = tk.Tk()
        root.withdraw()

        for p in targets:
            eml_to_html(p, save_external_images=True)

        messagebox.showinfo("完了", f"{len(targets)} 件の EML を変換しました。")

    sys.exit(0)

# GUI 起動
if __name__ == "__main__":
    gui = EmlConverterGUI()
    gui.run()