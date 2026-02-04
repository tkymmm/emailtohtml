import os
import extract_msg

from common import (
    decode_bytes,
    build_html_from_msg,
)


# ============================================================
# MSG を解析（extract_msg 版）
# ============================================================
def parse_msg_via_library(msg_path: str) -> dict:
    """
    extract_msg を使って MSG を解析し、
    HTML 本文・TEXT 本文・添付ファイル（OLE画像含む）を取得する。
    """
    msg = extract_msg.Message(msg_path)

    subject = msg.subject or ""
    body_html = msg.htmlBody or ""
    body_text = msg.body or ""

    attachments = []
    for att in msg.attachments:
        filename = att.longFilename or att.shortFilename or att.filename
        if not filename:
            continue

        data = att.data
        if not data:
            continue

        attachments.append({
            "filename": filename,
            "data": data
        })

    return {
        "subject": subject,
        "body_html": body_html,
        "body_text": body_text,
        "attachments": attachments,
    }


# ============================================================
# MSG 添付ファイル保存（OLE画像含む）
# ============================================================
def save_msg_attachments(msg_data: dict, attach_folder: str):
    os.makedirs(attach_folder, exist_ok=True)

    for att in msg_data.get("attachments", []):
        filename = att["filename"]
        data = att["data"]

        out_path = os.path.join(attach_folder, filename)

        # 同名ファイルがあれば連番を付ける
        base, ext = os.path.splitext(out_path)
        i = 1
        while os.path.exists(out_path):
            out_path = f"{base}({i}){ext}"
            i += 1

        with open(out_path, "wb") as f:
            f.write(data)


# ============================================================
# MSG → HTML（v2.0 完全版）
# ============================================================
def msg_to_html(msg_path: str, output_dir: str, save_external_images: bool):
    # 1. MSG を解析
    msg_data = parse_msg_via_library(msg_path)

    # 2. 出力ファイル名のベース
    base_name = os.path.splitext(os.path.basename(msg_path))[0]

    # 3. 添付フォルダ（外部画像保存のために常に作る）
    attach_folder = os.path.join(output_dir, base_name + "_files")
    os.makedirs(attach_folder, exist_ok=True)

    # 4. 添付ファイル保存（添付がある場合のみ）
    attachments = msg_data.get("attachments", [])
    if len(attachments) > 0:
        save_msg_attachments(msg_data, attach_folder)

    # 5. HTML 本文を構築（cid 置換・外部画像保存などは common 側）
    html = build_html_from_msg(
        msg_data,
        save_external_images,
        attach_folder,   # ← None にしない（外部画像保存のため）
        base_name
    )

    # 6. HTML の保存
    html_out = os.path.join(output_dir, base_name + ".html")
    with open(html_out, "w", encoding="utf-8", newline="\n") as f:
        f.write(html)

    return html_out