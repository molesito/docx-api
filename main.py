import os
import io
import re
import requests
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Inches

app = Flask(__name__)

@app.get("/health")
def health():
    return jsonify({"ok": True})

# ------------ Utilidades Markdown -> DOCX ----------------

HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)$')
UL_RE = re.compile(r'^\s*[-*+]\s+(.*)$')
OL_RE = re.compile(r'^\s*\d+\.\s+(.*)$')
IMG_RE = re.compile(r'!\[[^\]]*\]\(([^)]+)\)')  # ![alt](url)
TABLE_ROW_RE = re.compile(r'^\s*\|(.+)\|\s*$')
TABLE_ALIGN_RE = re.compile(r'^\s*\|?\s*(:?-{3,}:?\s*\|)+\s*(:?-{3,}:?)\s*\|?\s*$')

def add_paragraph(doc, text):
    if not text.strip():
        doc.add_paragraph("")  # línea en blanco
        return
    # Si hay imágenes inline en el párrafo, colócalas antes/después (Word no soporta inline igual que MD)
    imgs = IMG_RE.findall(text)
    clean = IMG_RE.sub("", text).strip()
    if clean:
        doc.add_paragraph(clean)
    for url in imgs:
        try:
            r = requests.get(url, timeout=12)
            r.raise_for_status()
            stream = io.BytesIO(r.content)
            doc.add_picture(stream, width=Inches(4))
        except Exception as e:
            doc.add_paragraph(f"[Imagen no insertada: {url} ({e})]")

def flush_list(doc, buf, ordered):
    if not buf:
        return
    style = "List Number" if ordered else "List Bullet"
    for item in buf:
        doc.add_paragraph(item, style=style)
    buf.clear()

def flush_table(doc, rows):
    if not rows:
        return
    # Ignora fila de alineación tipo | :-- | --- | :--: |
    filtered = []
    for r in rows:
        if TABLE_ALIGN_RE.match(r):
            continue
        filtered.append(r)
    if not filtered:
        return
    matrix = []
    for r in filtered:
        m = TABLE_ROW_RE.match(r)
        if not m:
            continue
        cells = [c.strip() for c in m.group(1).split("|")]
        matrix.append(cells)
    if not matrix:
        return
    cols = max(len(r) for r in matrix)
    table = doc.add_table(rows=len(matrix), cols=cols)
    table.style = "Table Grid"
    for i, row in enumerate(matrix):
        for j in range(cols):
            table.cell(i, j).text = row[j] if j < len(row) else ""

def markdown_to_doc(md_text, filename="output.docx"):
    doc = Document()

    lines = md_text.splitlines()
    ul_buf, ol_buf, tbl_buf = [], [], []
    in_table = False
    para_buf = []

    def flush_para():
        if para_buf:
            add_paragraph(doc, " ".join(para_buf).strip())
            para_buf.clear()

    for raw in lines:
        line = raw.rstrip("\n")

        # tablas (líneas que empiezan/terminan con |)
        if TABLE_ROW_RE.match(line):
            in_table = True
            tbl_buf.append(line)
            continue
        else:
            if in_table:
                # cierra tabla al salir del bloque
                flush_para()
                flush_list(doc, ul_buf, ordered=False)
                flush_list(doc, ol_buf, ordered=True)
                flush_table(doc, tbl_buf)
                tbl_buf = []
                in_table = False

        # encabezados
        m = HEADING_RE.match(line)
        if m:
            flush_para()
            flush_list(doc, ul_buf, ordered=False)
            flush_list(doc, ol_buf, ordered=True)
            level = len(m.group(1))
            text = m.group(2).strip()
            level = min(max(level, 1), 9)
            doc.add_heading(text, level=level)
            continue

        # listas
        m_ul = UL_RE.match(line)
        m_ol = OL_RE.match(line)
        if m_ul:
            flush_para()
            flush_list(doc, ol_buf, ordered=True)
            ul_buf.append(m_ul.group(1).strip())
            continue
        if m_ol:
            flush_para()
            flush_list(doc, ul_buf, ordered=False)
            ol_buf.append(m_ol.group(1).strip())
            continue

        # separación de párrafos
        if not line.strip():
            flush_para()
            flush_list(doc, ul_buf, ordered=False)
            flush_list(doc, ol_buf, ordered=True)
            continue

        # texto normal (acumula para párrafo)
        para_buf.append(line.strip())

    # flush final
    flush_para()
    flush_list(doc, ul_buf, ordered=False)
    flush_list(doc, ol_buf, ordered=True)
    if tbl_buf:
        flush_table(doc, tbl_buf)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, filename

# -------------------- Endpoint principal --------------------

@app.post("/docx")
def make_docx():
    """
    Acepta:
      - {"markdown": "...", "filename": "informe.docx"}
      - o {"text": "...", "filename": "..."} (compatibilidad)
    """
    data = request.get_json(silent=True) or {}
    filename = data.get("filename", "output.docx")

    if "markdown" in data and data["markdown"]:
        buf, fname = markdown_to_doc(data["markdown"], filename)
    else:
        # Compatibilidad con tu versión inicial: mete 'text' como un único párrafo
        text = data.get("text", "")
        doc = Document()
        add_paragraph(doc, text)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        fname = filename

    return send_file(
        buf,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
