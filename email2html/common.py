import os
import re
import urllib.request
import urllib.parse


# ============================================================
# 文字コード処理
# ============================================================
def decode_bytes(data: bytes, charset: str | None) -> str:
    if charset:
        try:
            return data.decode(charset, errors="replace")
        except Exception:
            pass

    for enc in ("utf-8", "cp932", "shift_jis", "iso-2022-jp", "euc-jp"):
        try:
            return data.decode(enc)
        except Exception:
            continue

    return data.decode("utf-8", errors="replace")


# ============================================================
# ファイル名の重複回避
# ============================================================
def ensure_unique(path: str) -> str:
    if not os.path.exists(path):
        return path

    base, ext = os.path.splitext(path)
    i = 1
    while True:
        new = f"{base}({i}){ext}"
        if not os.path.exists(new):
            return new
        i += 1


# ============================================================
# URL 自動リンク
# ============================================================
def autolink_plus(text: str) -> str:
    url_pattern = r"(https?://[^\s<>]+)"
    return re.sub(url_pattern, r'<a href="\1">\1</a>', text)


# ============================================================
# HTML の <meta charset> を保証
# ============================================================
def ensure_meta_charset(html: str) -> str:
    if "<meta charset" in html.lower():
        return html

    return html.replace(
        "<head>",
        '<head><meta charset="UTF-8">'
    )


# ============================================================
# Outlook HTML の補正
# ============================================================
def normalize_html(html: str) -> str:
    lower = html.lower()

    # MSO コメントを先頭に残す
    prefix = ""
    m = re.match(r'^(<!--\[if.*?endif\]-->\s*)', html, flags=re.I | re.S)
    if m:
        prefix = m.group(1)
        html = html[len(prefix):]
        lower = html.lower()

    # <html> が無ければ追加
    if "<html" not in lower:
        html = "<html>" + html + "</html>"
        lower = html.lower()

    # <head> が無ければ追加
    if "<head" not in lower:
        html = html.replace("<html>", "<html><head></head>")
        lower = html.lower()

    # <body> が無ければ追加
    if "<body" not in lower:
        html = html.replace("</head>", "</head><body>")
        if "</body>" not in lower:
            html += "</body>"

    return prefix + html


# ============================================================
# cid:画像 → 添付フォルダのパスに置換（MSG 用）
# ============================================================
def replace_cid_images(html: str, attachments: list, subfolder: str) -> str:
    for att in attachments:
        fname = att["filename"]
        base, ext = os.path.splitext(fname)

        # cid:image001.png または cid:image001.png@xxxx
        pattern = rf'cid:{re.escape(base)}[^"\']*'

        local_path = f"{subfolder}/{fname}"

        html = re.sub(pattern, local_path, html, flags=re.I)

    return html


# ============================================================
# 外部画像ダウンロード（日本語ファイル名対応）
# ============================================================
def download_external_images(html: str, folder: str, subfolder: str):
    save_dir = folder
    os.makedirs(save_dir, exist_ok=True)

    urls = re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', html)

    for src in urls:
        if not (src.startswith("http://") or src.startswith("https://")):
            continue

        parsed = urllib.parse.urlparse(src)
        filename = urllib.parse.unquote(os.path.basename(parsed.path))

        if not filename:
            continue

        save_path = ensure_unique(os.path.join(save_dir, filename))

        try:
            urllib.request.urlretrieve(src, save_path)
        except Exception:
            continue

        # HTML から見た相対パス
        local_ref = f"{subfolder}/{filename}"
        html = html.replace(src, local_ref)

    return html


# ============================================================
# MSG/EML 共通の HTML 組み立て（v2.0）
# ============================================================
def build_html_from_msg(msg_data: dict,
                        save_external_images: bool,
                        attach_folder: str,
                        base_name: str) -> str:

    html = msg_data.get("body_html") or ""

    if isinstance(html, bytes):
        html = decode_bytes(html, None)

    html = html.strip()

    if not html:
        text = msg_data.get("body_text", "") or ""
        html = f"<html><body><pre>{autolink_plus(text)}</pre></body></html>"
    else:
        html = normalize_html(html)

    html = ensure_meta_charset(html)

    # ★ attach_folder は呼び出し側で必ず決定済み
    os.makedirs(attach_folder, exist_ok=True)
    subfolder = os.path.basename(attach_folder)

    # cid 置換（MSG のみ）
    html = replace_cid_images(html, msg_data.get("attachments", []), subfolder)

    # 外部画像保存（常に attach_folder に保存）
    if save_external_images:
        html = download_external_images(html, attach_folder, subfolder)

    try:
        if os.path.isdir(attach_folder) and not os.listdir(attach_folder):
            os.rmdir(attach_folder)
    except Exception:
        pass

    return html


# ============================================================
# EML / MSG 自動判別
# ============================================================
def convert_any_email(path: str, output_dir: str | None, save_external_images: bool):
    ext = os.path.splitext(path)[1].lower()

    if output_dir:
        out_dir = output_dir
    else:
        out_dir = os.path.dirname(path)

    if ext == ".eml":
        from eml_converter import eml_to_html
        return eml_to_html(path, out_dir, save_external_images)

    elif ext == ".msg":
        from msg_converter import msg_to_html
        return msg_to_html(path, out_dir, save_external_images)

    else:
        raise ValueError("EML または MSG ファイルではありません。")