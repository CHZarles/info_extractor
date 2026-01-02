import os
import sqlite3
import tempfile
from io import BytesIO
from typing import List, Dict, Any

import pandas as pd
from cnocr import CnOcr
from flask import Flask, jsonify, request, send_file, send_from_directory

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "data.db")
CHANNEL_XLSX = os.path.join(BASE_DIR, "渠道明细.xlsx")

app = Flask(__name__, static_folder="static", static_url_path="")
# 指定检测模型权重，识别模型沿用默认
ocr_model = CnOcr(
    det_model_fp=os.path.join(BASE_DIR, "ppocr", "ch_PP-OCRv5_det", "ch_PP-OCRv5_det_infer.onnx"),
)


# ---------- DB helpers ----------
def get_db_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                日期 TEXT,
                微信名 TEXT,
                渠道 TEXT,
                备注意向 TEXT
            )
            """
        )
        conn.commit()


# ---------- OCR helpers ----------
_channel_cache: List[str] = []


def load_channels() -> List[str]:
    global _channel_cache
    if _channel_cache:
        return _channel_cache

    if not os.path.exists(CHANNEL_XLSX):
        _channel_cache = []
        return _channel_cache

    df = pd.read_excel(CHANNEL_XLSX, usecols=[0])
    values = df.iloc[:, 0].dropna().tolist()
    # values.append("二门诊")
    _channel_cache = values
    return _channel_cache


def parse_text_lines(text_lines: List[str], channels: List[str]) -> List[Dict[str, str]]:
    parsed: List[Dict[str, str]] = []
    for line in text_lines:
        item: Dict[str, str] = {}
        split_index = 8
        # for i, char in enumerate(line):
        #     if i >= 8:
        #         split_index = i
        #         break
        #     if not (char.isdigit() or char == "."):
        #         split_index = i
        #         break
        num_part = line[:split_index]
        chinese_part = line[split_index:]
        item["日期"] = num_part
        vaild = False
        for split_word in channels:
            if split_word in chinese_part:
                vaild = True
                parts = chinese_part.split(split_word)
                if len(parts) == 2:
                    part1 = parts[0]
                    part2 = split_word
                    part3 = parts[1]
                    item["微信名"] = part1
                    item["渠道"] = part2
                    item["备注/意向"] = part3
                    break
        if not vaild:
            continue
        parsed.append(item)
    return parsed


def run_ocr_on_image_bytes(image_bytes: bytes, channels: List[str]) -> List[Dict[str, str]]:
    with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
        tmp.write(image_bytes)
        tmp_path = tmp.name

    try:
        out = ocr_model.ocr(tmp_path)
        filtered = [item for item in out if item.get("score", 0) >= 0.2]
        filtered = [item for item in filtered if item.get("text", "") and item["text"][0].isdigit()]
        # filtered = [item for item in filtered if item.get("text", "") and "\u4e00" <= item["text"][-1] <= "\u9fa5"]
        text_lines = [item["text"] for item in filtered]
        return parse_text_lines(text_lines, channels)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


# ---------- Routes ----------
@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/ocr", methods=["POST"])
def api_ocr():
    if "images" not in request.files:
        return jsonify({"error": "No images provided"}), 400

    files = request.files.getlist("images")
    channels = load_channels()
    all_rows: List[Dict[str, Any]] = []

    for file in files:
        image_bytes = file.read()
        parsed_rows = run_ocr_on_image_bytes(image_bytes, channels)
        all_rows.extend(parsed_rows)

    return jsonify({"rows": all_rows})


@app.route("/api/channels", methods=["GET", "POST"])
def api_channels():
    if request.method == "GET":
        exists = os.path.exists(CHANNEL_XLSX)
        return jsonify({"exists": exists, "path": CHANNEL_XLSX if exists else ""})

    # POST upload
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400
    # basic extension guard
    if not file.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx allowed"}), 400

    file.save(CHANNEL_XLSX)
    # reset cache
    global _channel_cache
    _channel_cache = []
    return jsonify({"status": "ok", "path": CHANNEL_XLSX})


@app.route("/api/contacts", methods=["GET"])
def api_list_contacts():
    with get_db_conn() as conn:
        rows = conn.execute(
            "SELECT id, 日期, 微信名, 渠道, 备注意向 FROM contacts ORDER BY id DESC"
        ).fetchall()
    data = [
        {
            "id": row["id"],
            "日期": row["日期"] or "",
            "微信名": row["微信名"] or "",
            "渠道": row["渠道"] or "",
            "备注/意向": row["备注意向"] or "",
        }
        for row in rows
    ]
    return jsonify({"rows": data})


@app.route("/api/contacts/bulk", methods=["POST"])
def api_save_contacts():
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows", [])
    if not isinstance(rows, list):
        return jsonify({"error": "Invalid payload"}), 400

    with get_db_conn() as conn:
        cursor = conn.cursor()
        saved_rows: List[Dict[str, Any]] = []
        for row in rows:
            row = row or {}
            row_id = row.get("id")
            date = row.get("日期", "")
            name = row.get("微信名", "")
            channel = row.get("渠道", "")
            note = row.get("备注/意向", "")

            if row_id:
                cursor.execute(
                    "UPDATE contacts SET 日期=?, 微信名=?, 渠道=?, 备注意向=? WHERE id=?",
                    (date, name, channel, note, row_id),
                )
                saved_id = row_id
            else:
                cursor.execute(
                    "INSERT INTO contacts (日期, 微信名, 渠道, 备注意向) VALUES (?, ?, ?, ?)",
                    (date, name, channel, note),
                )
                saved_id = cursor.lastrowid
            saved_rows.append(
                {"id": saved_id, "日期": date, "微信名": name, "渠道": channel, "备注/意向": note}
            )
        conn.commit()

    return jsonify({"rows": saved_rows})


@app.route("/api/contacts/<int:contact_id>", methods=["DELETE"])
def api_delete_contact(contact_id: int):
    with get_db_conn() as conn:
        conn.execute("DELETE FROM contacts WHERE id=?", (contact_id,))
        conn.commit()
    return jsonify({"status": "ok"})


@app.route("/api/export", methods=["GET", "POST"])
def api_export():
    if request.method == "POST":
        payload = request.get_json(silent=True) or {}
        rows = payload.get("rows", [])
        if not isinstance(rows, list):
            return jsonify({"error": "Invalid rows"}), 400
        df = pd.DataFrame(rows)
        # Ensure column order / missing columns are handled
        for col in ["日期", "微信名", "渠道", "备注/意向"]:
            if col not in df.columns:
                df[col] = ""
        df = df[["日期", "微信名", "渠道", "备注/意向"]]
    else:
        with get_db_conn() as conn:
            df = pd.read_sql_query(
                "SELECT 日期, 微信名, 渠道, 备注意向 as 备注/意向 FROM contacts ORDER BY id DESC",
                conn,
            )

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="contacts.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)
