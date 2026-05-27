from __future__ import annotations
from flask import Flask, render_template, request, send_file, jsonify
from io import BytesIO
from py_backend.main import process_sox_files

app = Flask(__name__)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/process")
def process():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    uploaded = []
    for f in files:
        name = f.filename or "file.xlsx"
        if not name.lower().endswith(".xlsx"):
            return jsonify({"error": "Only .xlsx files are allowed"}), 400
        uploaded.append((name, f.read()))

    blob, filename = process_sox_files(uploaded)
    return send_file(BytesIO(blob), as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
