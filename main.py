import os
import io
import re
import base64
import requests
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Inches, RGBColor

app = Flask(__name__)

@app.get("/health")
def health():
    return jsonify({"ok": True})

# ----------------- Regex y utilidades -----------------

HEADING_RE     = re.compile(r'^(#{1,6})\s+(.*)$')
UL_RE          = re.compile(r'^\s*[-*+]\s+(.*)$')
OL_RE          = re.compile(r'^\s*\d+\.\s+(.*)$')
TABLE_ROW_RE   = re.compile(r'^\s*\|(.+)\|\s*$')
TABLE_ALIGN_RE = re.compile(r'^\s*\|?\s*(:?-{3,}:?\s*\|)+\s*(:?-{3,}:?)\s*\|?\s*$')
IMG_RE         = re.compile(r'!\[[^\]]*\]\(([^)]+)\)')  # ![alt](target)

def force_styles_black(doc: Document):
    """Fuerza color negro en estilos habituales (títulos, normal, listas)."""
    target_styles = ["Normal", "List Paragraph", "List Bullet", "List Number"]
    target_styles += [f"Heading {i}" for i in range(1, 10)]
    for name in target_styles:
        try:
            style = doc.styles[name]
            if style and style.font:
                style.font.color.rgb = RGBColor(0, 0, 0)
        except KeyError:
            pass

def add_paragraph(doc: Document, text: str, images_map: dict):
    """
    Inserta el párrafo y, si hay ![](...), coloca imágenes.
    targets soportados:
      - http/https → se descargan
      - 'img-X.jpeg' → se busca en images_map
    """
    images_map = images_map or {}

    targets = IMG_RE.findall(text)
    clean_text = IMG_RE.sub("", text).strip()

    if not clean_text and not targets:
        doc.add_paragraph("")
    elif clean_text:
        doc.add_paragraph(clean_text)

    for target in targets:
        try:
            if target.lower().startswith(("http://", "https://")):
                r = requests.get(target, timeout=12)
                r.raise_for_status()
                stream = io.BytesIO(r.content)
                doc.add_picture(stream, width=Inches(4))
            else:
                if target not in images_map:
                    raise ValueError(f"imagen '{target}' no recibida")
                stream = io.BytesIO(images_map[target])
                doc.add_picture(stream, width=Inches(4))
        except Exception as e:
            doc.add_paragraph(f"[Imagen no insertada: {target} ({e})]")

def flush_list(doc: Document, buf: list, ordered: bool, images_map: dict):
    if not buf:
        return
    style = "List Number" if ordered else "List Bullet"
    for item in buf:
        # Renderizamos como párrafo de lista; si hay imágenes, se insertan después
        p = doc.add_paragraph("", style=style)
        p.add_run(item)
        # Si el ítem tuviera ![](...), renderízalo aparte:
        tmp = Document()
        add_paragraph(tmp, item, images_map)
        for extra in tmp.paragraphs[1:]:
            doc.add_paragraph(extra.text)
    buf.clear()

def flush_table(doc: Document, rows: list):
    if not rows:
        return
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

def markdown_to_doc(md_text: str, images_map: dict = None):
    doc = Document()
    force_styles_black(doc)

    lines = md_text.splitlines()
    ul_buf, ol_buf, tbl_buf = [], [], []
    in_table = False
    para_buf = []

    def flush_para():
        if para_buf:
            add_paragraph(doc, " ".join(para_buf).strip(), images_map)
            para_buf.clear()

    for raw in lines:
        line = raw.rstrip("\n")

        # tablas
        if TABLE_ROW_RE.match(line):
            in_table = True
            tbl_buf.append(line)
            continue
        else:
            if in_table:
                flush_para()
                flush_list(doc, ul_buf, ordered=False, images_map=images_map)
                flush_list(doc, ol_buf, ordered=True, images_map=images_map)
                flush_table(doc, tbl_buf)
                tbl_buf = []
                in_table = False

        # encabezados
        m = HEADING_RE.match(line)
        if m:
            flush_para()
            flush_list(doc, ul_buf, ordered=False, images_map=images_map)
            flush_list(doc, ol_buf, ordered=True, images_map=images_map)
            level = min(max(len(m.group(1)), 1), 9)
            text = m.group(2).strip()
            doc.add_heading(text, level=level)
            continue

        # listas
        m_ul = UL_RE.match(line)
        m_ol = OL_RE.match(line)
        if m_ul:
            flush_para()
            flush_list(doc, ol_buf, ordered=True, images_map=images_map)
            ul_buf.append(m_ul.group(1).strip())
            continue
        if m_ol:
            flush_para()
            flush_list(doc, ul_buf, ordered=False, images_map=images_map)
            ol_buf.append(m_ol.group(1).strip())
            continue

        # separación
        if not line.strip():
            flush_para()
            flush_list(doc, ul_buf, ordered=False, images_map=images_map)
            flush_list(doc, ol_buf, ordered=True, images_map=images_map)
            continue

        # texto normal
        para_buf.append(line.strip())

    flush_para()
    flush_list(doc, ul_buf, ordered=False, images_map=images_map)
    flush_list(doc, ol_buf, ordered=True, images_map=images_map)
    if tbl_buf:
        flush_table(doc, tbl_buf)

    force_styles_black(doc)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

# ----------------- Endpoint -----------------

@app.post("/docx")
def make_docx():
    """
    multipart/form-data:
      - markdown (texto)
      - filename (opcional)
      - data0, data1, data2 ... (binarios). Se mapearán a:
          dataN -> img-N.jpeg
      - También acepta uploads con filename 'img-N.jpeg'.

    application/json (opcional):
      { "filename": "...", "markdown": "...", "images_map": { "img-0.jpeg": <bytes base64> } }
    """
    # Modo multipart: recomendado para n8n con data0..dataN
    if request.content_type and request.content_type.startswith("multipart/form-data"):
        md_text = request.form.get("markdown", "")
        filename = request.form.get("filename", "output.docx")

        images_map = {}
        for fieldname, storage in request.files.items():
            content = storage.read()
            if not content:
                continue

            # 1) Si viene como dataN -> mapear a img-N.jpeg
            m = re.match(r"data(\d+)$", fieldname)
            if m:
                key = f"img-{m.group(1)}.jpeg"
                images_map[key] = content

            # 2) Si el filename ya es 'img-N.jpeg', también guardarlo
            if storage.filename:
                images_map[storage.filename] = content

            # 3) Por compatibilidad, también guarda el nombre del campo tal cual
            images_map[fieldname] = content

        buf = markdown_to_doc(md_text, images_map)
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # Modo JSON (opcional)
    data = request.get_json(silent=True) or {}
    filename = data.get("filename", "output.docx")
    md_text  = data.get("markdown", data.get("text", ""))
    images_map_b64 = data.get("images_map") or {}

    # Si vienen en base64 por JSON, decodifica
    images_map = {}
    for k, v in images_map_b64.items():
        try:
            if isinstance(v, (bytes, bytearray)):
                images_map[k] = v
            else:
                if isinstance(v, str) and v.startswith("data:"):
                    v = v.split(",", 1)[1]
                images_map[k] = base64.b64decode(v)
        except Exception:
            pass

    buf = markdown_to_doc(md_text, images_map)
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

