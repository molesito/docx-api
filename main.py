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
    """
    Espera un JSON como este:
    {
      "filename": "output.docx",
      "content": [
        {"type": "heading", "text": "Título", "level": 1},
        {"type": "paragraph", "text": "Texto normal"},
        {"type": "list", "ordered": false, "items": ["Uno","Dos"]},
        {"type": "table", "rows": [["Col1","Col2"],["A","B"]]},
        {"type": "image", "url": "https://.../imagen.png", "width_in": 2.5}
      ]
    }
    """
    data = request.get_json(silent=True) or {}
    filename = data.get("filename", "output.docx")
    content = data.get("content", [])

    doc = Document()

    for item in content:
        t = (item.get("type") or "").lower()

        if t == "heading":
            doc.add_heading(item.get("text", ""), level=int(item.get("level", 1)))

        elif t == "paragraph":
            doc.add_paragraph(item.get("text", ""))

        elif t == "list":
            style = "List Number" if item.get("ordered") else "List Bullet"
            for it in item.get("items", []):
                doc.add_paragraph(str(it), style=style)

        elif t == "table":
            rows = item.get("rows", [])
            if rows:
                cols = max(len(r) for r in rows)
                table = doc.add_table(rows=len(rows), cols=cols)
                table.style = "Table Grid"
                for i, row in enumerate(rows):
                    for j, val in enumerate(row):
                        table.cell(i, j).text = str(val)

        elif t == "image":
            try:
                url = item.get("url")
                width_in = item.get("width_in")
                r = requests.get(url, timeout=15)
                r.raise_for_status()
                stream = io.BytesIO(r.content)
                if width_in:
                    doc.add_picture(stream, width=Inches(float(width_in)))
                else:
                    doc.add_picture(stream)
            except Exception as e:
                doc.add_paragraph(f"[No se pudo insertar la imagen: {e}]")

        elif t == "page_break":
            doc.add_page_break()

        else:
            doc.add_paragraph(str(item))

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
