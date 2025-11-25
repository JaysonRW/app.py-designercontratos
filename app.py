import os
import tempfile
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
# Certifique-se de instalar: pip install flask flask-cors python-docx

app = Flask(__name__)
# Permite que seu frontend na Vercel converse com este backend
CORS(app, resources={r"/api/*": {"origins": "*"}})

# Em produção, usamos diretórios temporários
UPLOAD_FOLDER = tempfile.gettempdir()

@app.route('/api/process', methods=['POST'])
def process_contract():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['file']
        # primary_color = request.form.get('primaryColor') # Usar na lógica real
        
        # Salva temporariamente
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(input_path)

        # --- LÓGICA DE PROCESSAMENTO (SIMPLIFICADA) ---
        # Aqui entra seu código real com python-docx
        # from docx import Document
        # doc = Document(input_path)
        # ... edições ...
        # doc.save(output_path)
        
        # Simulando saída (apenas renomeando)
        output_filename = f"FORMATADO_{file.filename}"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        
        # Copia simples para exemplo (substitua pela lógica do docx)
        import shutil
        shutil.copy(input_path, output_path)
        # -----------------------------------------------

        # Retorna URL completa para download
        # Em produção, request.host_url será algo como "https://seu-backend.onrender.com/"
        base_url = request.url_root.rstrip('/')
        
        return jsonify({
            'docxUrl': f"{base_url}/api/download?file={output_filename}",
            'pdfUrl': f"{base_url}/api/download?file={output_filename}" # Placeholder
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
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        return "Arquivo expirou ou não encontrado.", 404

if __name__ == '__main__':
    # Pega a porta do ambiente (necessário para Render/Heroku)
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
