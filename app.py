import os
import tempfile
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}})

UPLOAD_FOLDER = tempfile.gettempdir()

def apply_formatting(doc, primary_color, logo_path=None):
    """
    Aplica formatação visual ao documento
    """
    # Converte cor hex para RGB
    color = primary_color.lstrip('#')
    r, g, b = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
    
    # 1. Adiciona logo no cabeçalho (se fornecido)
    if logo_path and os.path.exists(logo_path):
        section = doc.sections[0]
        header = section.header
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        run = paragraph.add_run()
        run.add_picture(logo_path, width=Inches(1.5))
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 2. Formata todos os parágrafos
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # Aplica cor primária aos títulos (texto em negrito)
            if run.bold:
                run.font.color.rgb = RGBColor(r, g, b)
            # Melhora legibilidade do texto normal
            else:
                run.font.color.rgb = RGBColor(50, 50, 50)
            
            # Ajusta tamanho da fonte
            if run.font.size and run.font.size < Pt(10):
                run.font.size = Pt(11)
    
    # 3. Adiciona bordas/sombreamento às tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                # Cabeçalhos de tabela com cor primária
                if row == table.rows[0]:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(255, 255, 255)
                            run.font.bold = True
                    # Fundo colorido (requer manipulação XML)
                    cell._element.get_or_add_tcPr().append(
                        parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
                    )
    
    return doc

@app.route('/api/process', methods=['POST'])
def process_contract():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        primary_color = request.form.get('primaryColor', '#2563eb')
        
        # Salva arquivo temporariamente
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(input_path)
        
        # Processa logo se enviado
        logo_path = None
        if 'logo' in request.files:
            logo = request.files['logo']
            logo_path = os.path.join(UPLOAD_FOLDER, f"logo_{logo.filename}")
            logo.save(logo_path)
        
        # Abre e processa documento
        doc = Document(input_path)
        doc = apply_formatting(doc, primary_color, logo_path)
        
        # Salva documento formatado
        output_filename = f"FORMATADO_{file.filename}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        doc.save(output_path)
        
        # Limpa arquivos temporários de entrada
        os.remove(input_path)
        if logo_path and os.path.exists(logo_path):
            os.remove(logo_path)
        
        base_url = request.url_root.rstrip('/')
        
        return jsonify({
            'docxUrl': f"{base_url}/api/download?file={output_filename}",
            'message': 'Documento formatado com sucesso!'
        })
        
    except Exception as e:
        print(f"Erro: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/download', methods=['GET'])
def download_file():
    filename = request.args.get('file')
    if not filename:
        return "Arquivo não especificado", 400
    
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    
    try:
        response = send_file(file_path, as_attachment=True)
        # Limpa arquivo após download
        @response.call_on_close
        def cleanup():
            try:
                os.remove(file_path)
            except:
                pass
        return response
    except FileNotFoundError:
        return "Arquivo não encontrado", 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
