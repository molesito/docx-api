import os
import io
import requests
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Inches

app = Flask(__name__)

@app.get("/health")
def health():
    return jsonify({"ok": True})

@app.post("/docx")
def make_docx():
    data = request.get_json(silent=True) or {}
    text = data.get("text", "Hola mundo desde Flask + Render")
    filename = data.get("filename", "output.docx")

    doc = Document()
    doc.add_paragraph(text)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
