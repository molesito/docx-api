import os
import io
import re
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import RGBColor

app = Flask(__name__)

@app.get("/health")
def health():
    return jsonify({"ok": True})

# ------------ Utilidades Markdown -> DOCX (sin imágenes) ----------------

HEADING_RE = re.compile(r'^(#{1,6})\s+(.*)$')
UL_RE      = re.compile(r'^\s*[-*+]\s+(.*)$')
OL_RE      = re.compile(r'^\s*\d+\.\s+(.*)$')
TABLE_ROW_RE   = re.compile(r'^\s*\|(.+)\|\s*$')
TABLE_ALIGN_RE = re.compile(r'^\s*\|?\s*(:?-{3,}:?\s*\|)+\s*(:?-{3,}:?)\s*\|?\s*$')

def force_styles_black(doc: Document):
    """Pone el color de fuente a negro en estilos usados (encabezados, normal, listas)."""
    target_styles = ["Normal", "List Paragraph", "List Bullet", "List Number"]
    target_styles += [f"Heading {i}" for i in range(1, 10)]
    for name in target_styles:
        try:
            style = doc.styles[name]
            if style and style.font:
                style.font.color.rgb = RGBColor(0, 0, 0)
        except KeyError:
            # Si el estilo no existe en esta plantilla, seguimos
            pass

def add_paragraph(doc, text):
    if not text.strip():
        doc.add_paragraph("")  # línea en blanco
    else:
        p = doc.add_paragraph(text)

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
    filtered = [r for r in rows if not TABLE_ALIGN_RE.match(r)]
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
    force_styles_black(doc)

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
            h = doc.add_heading(text, level=level)
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

    # Asegura estilos negros tras crear contenido (por si Word reintroduce colores por tema)
    force_styles_black(doc)

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

    if data.get("markdown"):
        buf, fname = markdown_to_doc(data["markdown"], filename)
    else:
        # compat: texto plano en un párrafo
        doc = Document()
        force_styles_black(doc)
        add_paragraph(doc, data.get("text", ""))
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

