import os
import re
import tempfile
from uuid import uuid4
from typing import Optional, List, Tuple

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

from docx import Document
from docx.shared import RGBColor, Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


app = Flask(__name__)

# CORS amplo para ambientes dev/prod
DEFAULT_ORIGINS = [
    "https://designer-contratos-v0-01.vercel.app",
    "http://localhost:3000",
    "*",
]
CORS(app, resources={r"/*": {"origins": os.environ.get("ALLOWED_ORIGINS", ",".join(DEFAULT_ORIGINS)).split(",")}},
     supports_credentials=True,
     allow_headers=["*"],
     methods=["*"])

# Diretório de arquivos gerados (persistência temporária)
FILES_DIR = os.environ.get("FILES_DIR", os.path.join(tempfile.gettempdir(), "designer_contratos"))
os.makedirs(FILES_DIR, exist_ok=True)


def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    color = hex_color.lstrip('#')
    return tuple(int(color[i:i+2], 16) for i in (0, 2, 4))


def shade_cell(cell, hex_color: str):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), hex_color.lstrip('#'))
    cell._element.get_or_add_tcPr().append(shading_elm)


def create_formatted_table(doc: Document, data: List[Tuple[str, str]], primary_color: str):
    r, g, b = hex_to_rgb(primary_color)
    table = doc.add_table(rows=len(data), cols=2)
    table.style = 'Light Grid Accent 1'

    for i, (key, value) in enumerate(data):
        row = table.rows[i]

        # Coluna da chave
        key_cell = row.cells[0]
        key_para = key_cell.paragraphs[0]
        key_run = key_para.add_run(key)
        key_run.font.bold = True
        key_run.font.size = Pt(10)
        key_run.font.color.rgb = RGBColor(255, 255, 255)
        shade_cell(key_cell, primary_color)

        # Coluna do valor
        value_cell = row.cells[1]
        value_para = value_cell.paragraphs[0]
        value_run = value_para.add_run(value)
        value_run.font.size = Pt(10)

        # Zebra na coluna do valor
        if i % 2 == 1:
            shade_cell(value_cell, 'F3F4F6')


def build_document(text_content: str, primary_color: str, logo_path: Optional[str]) -> Document:
    doc = Document()

    # Página A4 e margens
    section = doc.sections[0]
    section.page_height = Cm(29.7)
    section.page_width = Cm(21)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2)

    r_primary, g_primary, b_primary = hex_to_rgb(primary_color)

    # Cabeçalho com logo e separador
    if logo_path and os.path.exists(logo_path):
        header = section.header
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = header_para.add_run()
        run.add_picture(logo_path, width=Inches(1.8))

        sep = header.add_paragraph()
        sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sep_run = sep.add_run('_' * 80)
        sep_run.font.color.rgb = RGBColor(r_primary, g_primary, b_primary)
        sep_run.font.size = Pt(6)

    lines = text_content.split('\n')
    table_data: List[Tuple[str, str]] = []
    in_table = False

    for line in lines:
        line = line.strip()

        # Linha vazia: fecha tabela se aberta e adiciona espaçamento
        if not line:
            if in_table and table_data:
                create_formatted_table(doc, table_data, primary_color)
                table_data = []
                in_table = False
            doc.add_paragraph()
            continue

        # Detecção de tabela simples: "CHAVE: Valor"
        if ':' in line:
            parts = line.split(':', 1)
            if len(parts) == 2:
                table_data.append((parts[0].strip(), parts[1].strip()))
                in_table = True
                continue

        # Se estava em tabela mas linha não é de tabela, cria a tabela
        if in_table and table_data:
            create_formatted_table(doc, table_data, primary_color)
            table_data = []
            in_table = False

        # Título em maiúsculas, começa com CONTRATO/CLÁUSULA ou numeração
        is_title = (
            line.isupper() or
            line.startswith('CONTRATO') or
            line.startswith('CLÁUSULA') or
            re.match(r'^\d+\.', line) is not None
        )

        # Item de lista
        is_list_item = line.startswith(('- ', '• ', '* '))

        if is_title:
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(line)
            run.font.size = Pt(14)
            run.font.bold = True
            run.font.color.rgb = RGBColor(r_primary, g_primary, b_primary)
            run.font.name = 'Arial'
        elif is_list_item:
            para = doc.add_paragraph(line)
            for run in para.runs:
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(50, 50, 50)
                run.font.name = 'Arial'
        else:
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            run = para.add_run(line)
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(50, 50, 50)
            run.font.name = 'Arial'

    # Última tabela pendente
    if table_data:
        create_formatted_table(doc, table_data, primary_color)

    # Rodapé
    footer = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_para.add_run('Documento gerado automaticamente')
    footer_run.font.size = Pt(8)
    footer_run.font.color.rgb = RGBColor(150, 150, 150)

    return doc


def make_public_url(filename: str) -> str:
    proto = request.headers.get('X-Forwarded-Proto', 'https')
    host = request.headers.get('Host', request.host)
    return f"{proto}://{host}/files/{filename}"


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.get("/files/<path:filename>")
def files(filename: str):
    return send_from_directory(FILES_DIR, filename, as_attachment=False)


@app.post("/api/process")
def process():
    text = request.form.get("text", "")
    primary_color = request.form.get("primaryColor", "#4F46E5")
    logo = request.files.get("logo")

    if not text.strip():
        return jsonify({"error": "Texto não fornecido"}), 400

    logo_path = None
    if logo:
        ext = os.path.splitext(logo.filename)[1] or ".png"
        logo_name = f"logo_{uuid4().hex}{ext}"
        logo_path = os.path.join(FILES_DIR, logo_name)
        logo.save(logo_path)

    # Geração do DOCX
    doc = build_document(text, primary_color, logo_path)
    docx_name = f"contrato_{uuid4().hex}.docx"
    docx_path = os.path.join(FILES_DIR, docx_name)
    doc.save(docx_path)

    # PDF opcional: em ambientes Linux (Render) a docx2pdf não funciona nativamente
    # Retornamos o DOCX e, se quiser, apontamos o PDF para o mesmo arquivo.
    docx_url = make_public_url(docx_name)
    pdf_url = docx_url

    return jsonify({"docxUrl": docx_url, "pdfUrl": pdf_url})


# Aliases para compatibilidade com o frontend que tenta rotas alternativas
@app.post("/process")
def process_alias():
    return process()

@app.post("/api/process_text")
def process_text():
    return process()

@app.post("/process_text")
def process_text_alias():
    return process()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
