import os
from email import policy
from email.parser import BytesParser

from common import (
    decode_bytes,
    autolink_plus,
    normalize_html,
    ensure_meta_charset,
    build_html_from_msg,
)


# ============================================================
# EML 読み込み
# ============================================================
def read_eml(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


# ============================================================
# 添付ファイル抽出（0 件ならフォルダを作らない）
# ============================================================
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

    # 添付が 0 件ならフォルダを作らない
    if not attachments:
        return 0

    os.makedirs(base_folder, exist_ok=True)

    for filename, payload in attachments:
        out_path = os.path.join(base_folder, filename)
        with open(out_path, "wb") as f:
            f.write(payload)

    return len(attachments)


# ============================================================
# 最適な本文パートを選択
# ============================================================
def pick_best_part(msg):
    html_part = None
    text_part = None

    for part in msg.walk():
        ctype = part.get_content_type()

        if ctype == "text/html":
            html_part = part
        elif ctype == "text/plain":
            text_part = part

    return html_part or text_part


# ============================================================
# EML → HTML（v2.0 完全版）
# ============================================================
def eml_to_html(eml_path: str, output_dir: str | None = None, save_external_images: bool = False):
    raw = read_eml(eml_path)
    msg = BytesParser(policy=policy.default).parsebytes(raw)

    # 出力先フォルダ
    folder = output_dir if output_dir else os.path.dirname(eml_path)
    base = os.path.splitext(os.path.basename(eml_path))[0]

    html_out = os.path.join(folder, base + ".html")
    attach_folder = os.path.join(folder, base + "_files")

    # 添付ファイル保存（0 件ならフォルダを作らない）
    attach_count = extract_attachments(msg, attach_folder)

    # 添付が 0 件ならフォルダ削除（存在していれば）
    if attach_count == 0 and os.path.exists(attach_folder):
        try:
            os.rmdir(attach_folder)
        except OSError:
            pass

    # 本文抽出
    part = pick_best_part(msg)
    if part is None:
        body_html = "<html><body><pre>本文が見つかりませんでした。</pre></body></html>"
    else:
        payload = part.get_payload(decode=True) or b""
        charset = part.get_content_charset()
        text = decode_bytes(payload, charset)

        if part.get_content_type() == "text/html":
            # HTML パート
            body_html = normalize_html(text)
            body_html = ensure_meta_charset(body_html)
        else:
            # TEXT パート → <pre> で改行保持
            body_html = (
                "<html><head><meta charset=\"UTF-8\"></head>"
                "<body><pre>" + autolink_plus(text) + "</pre></body></html>"
            )

    # ★ 外部画像保存のために attach_folder を必ず渡す
    #    （添付が無くてもフォルダは build_html_from_msg 内で作られる）
    final_html = build_html_from_msg(
        {
            "body_html": body_html,
            "body_text": "",
            "attachments": [],  # EML の添付は cid 参照しないので空でOK
        },
        save_external_images,
        attach_folder,   # ← None にしない
        base
    )

    # HTML 保存
    with open(html_out, "w", encoding="utf-8", newline="\n") as f:
        f.write(final_html)

    return html_out