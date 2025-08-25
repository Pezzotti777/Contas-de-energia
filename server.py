from flask import Flask, request, jsonify
import requests
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

app = Flask(__name__)

# URL do seu Apps Script publicado como “App da Web”
WEBAPP_URL = 'https://script.google.com/macros/s/SEU_ID_AQUI/exec'

# -------------------------------------------------------
# 1) FUNÇÃO: envia JSON ao Apps Script para inserir na planilha
# -------------------------------------------------------
def append_to_sheet(payload: dict):
    resp = requests.post(WEBAPP_URL, json=payload)
    resp.raise_for_status()
    return resp.json()  # { status: 'ok' } ou { status: 'erro', message: ... }

# -------------------------------------------------------
# 2) FUNÇÃO: autentica e faz upload do PDF ao Google Drive
# -------------------------------------------------------
def upload_pdf_to_drive(local_pdf_path: str, title: str = None):
    gauth = GoogleAuth()
    # executa o fluxo de OAuth no navegador na primeira vez
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)

    file_drive = drive.CreateFile({
        'title': title or local_pdf_path,
        # opcional: para enviar a uma pasta específica, coloque 'parents': [{'id': 'PASTA_ID'}]
    })
    file_drive.SetContentFile(local_pdf_path)
    file_drive.Upload()
    return file_drive['id']

# -------------------------------------------------------
# ROTA: recebe PDF + JSON de dados e processa ambos uploads
# -------------------------------------------------------
@app.route('/extract', methods=['POST'])
def extract():
    try:
        # 1) JSON de campos extraídos (cliente, emissao, vencimento, valor, etc.)
        data = request.get_json()

        # 2) PDF a ser salvo localmente (campo 'pdf_base64' ou 'pdf_path')
        #    Aqui assumimos que o PDF já foi salvo em disco por outro código:
        pdf_path = data.get('pdf_path', 'saida.pdf')

        # envia dados à planilha
        sheet_resp = append_to_sheet(data)

        # faz upload do PDF ao Drive
        drive_file_id = upload_pdf_to_drive(pdf_path, title=f"conta_{data.get('cliente')}.pdf")

        return jsonify({
            'status': 'ok',
            'sheet': sheet_resp,
            'drive_file_id': drive_file_id
        }), 200

    except requests.HTTPError as e:
        return jsonify(status='erro-sheet', detail=str(e)), 500
    except Exception as e:
        return jsonify(status='erro-geral', detail=str(e)), 500

if __name__ == '__main__':
    app.run(port=3000, debug=True)