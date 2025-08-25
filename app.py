from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import requests  # Importante!

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# === Função para enviar para a planilha Google ===
def envia_para_planilha(dados):
    url = 'https://script.google.com/macros/s/AKfycbwoLU5XFzqErQ04gHK5-Juvr2g34yPj0qEnW8OYnqvM0z9F0oVd44yOrbvdMJiWSOwW/exec'  # <-- Cole aqui sua URL do Apps Script
    resp = requests.post(url, json=dados)
    print(resp.text)  # Só para debug. Vai aparecer "OK" se der certo.

# === Rota de upload/importação ===
@app.route('/importar', methods=['POST'])
def importar():
    if 'pdfs' not in request.files:
        return jsonify({'success': False, 'error': 'Nenhum arquivo enviado'})
    files = request.files.getlist('pdfs')
    resultados = []
    for file in files:
        caminho = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(caminho)

        # --- Aqui entra sua lógica de leitura do PDF e extração de dados ---
        # Por enquanto, vamos simular dados para teste:
        dados = [
            "Exemplo Instalação", "2025-07-14", "2025-08-01", "100,00", "22", "2025-07-14",
            "2025-06-14", "2025-07-14", "", "TESTE", "3", "6"
        ]
        # Envia para planilha:
        envia_para_planilha(dados)
        resultados.append(f"Arquivo {file.filename} processado e enviado para planilha.")
    return jsonify({'success': True, 'detalhes': resultados})

if __name__ == '__main__':
    app.run(debug=True)
