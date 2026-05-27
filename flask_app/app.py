from __future__ import annotations

import io

from flask import Flask, jsonify, render_template, request, send_file

from sox_processor import InputFile, classify_error, overrides_from_json, process_sox_files


app = Flask(__name__)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/process")
def process():
    uploaded = request.files.getlist("files")
    files: list[InputFile] = []

    for storage in uploaded:
        if not storage.filename:
            continue
        files.append(InputFile(name=storage.filename, data=storage.read()))

    try:
        result = process_sox_files(files, overrides_from_json(request.form.get("manualOverrides")))
    except Exception as exc:
        return jsonify({"error": classify_error(exc)}), 400

    return send_file(
        io.BytesIO(result.data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=result.filename,
    )


if __name__ == "__main__":
    app.run(debug=True)
