from flask import Flask, render_template, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pdfplumber
import os
import re
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
CREDENCIAL = 'client_secret.json'
PLANILHA_URL = 'https://docs.google.com/spreadsheets/d/170LPTCD-_9Dk6oOt6D2SGNr7eQQaDFS8h4SuVW92N1c/edit?usp=sharing'
ABA = 'CONTAS'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENCIAL, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url(PLANILHA_URL)
worksheet = sheet.worksheet(ABA)
headers = worksheet.row_values(1)

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def extrair_texto(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join([page.extract_text() or "" for page in pdf.pages])

def limpar_valor(val):
    return val.replace('.', '').replace(',', '.').strip()

def extrair_dados_por_regex(texto):
    resultados = {h: "" for h in headers}

    resultados['instalacao'] = re.search(r"N\u00ba DA INSTALA\u00c7\u00c3O\s*(\d+)", texto).group(1) if re.search(r"N\u00ba DA INSTALA\u00c7\u00c3O\s*(\d+)", texto) else "0"
    resultados['fatDataVcto'] = re.search(r"Vencimento\s*(\d{2}/\d{2}/\d{4})", texto).group(1) if re.search(r"Vencimento\s*(\d{2}/\d{2}/\d{4})", texto) else "0"
    resultados['fatDataEmissao'] = re.search(r"Data de emiss\u00e3o:\s*(\d{2}/\d{2}/\d{4})", texto).group(1) if re.search(r"Data de emiss\u00e3o:\s*(\d{2}/\d{2}/\d{4})", texto) else "0"
    resultados['NOTAFISCAL'] = re.search(r"NOTA FISCAL N\u00ba\s*(\d+)", texto).group(1) if re.search(r"NOTA FISCAL N\u00ba\s*(\d+)", texto) else "0"
    resultados['fatCodigoBarras'] = re.search(r"(\d{11}-\d\s+){3}\d{11}-\d", texto).group(0) if re.search(r"(\d{11}-\d\s+){3}\d{11}-\d", texto) else "0"

    leitura = re.search(r"Datas de Leitura.*?(\d{2}/\d{2})\s+(\d{2}/\d{2})\s+(\d+)\s+(\d{2}/\d{2})", texto)
    if leitura:
        resultados['fatDataLeituraAnterior'] = leitura.group(1)
        resultados['fatDataLeituraAtual'] = leitura.group(2)
        resultados['fatNDias'] = leitura.group(3)
        resultados['fatDataLeituraProxima'] = leitura.group(4)

    resultados['fatValorFatura'] = re.search(r"Valor a pagar.*?R\$\s*([\d.,]+)", texto).group(1) if re.search(r"Valor a pagar.*?R\$\s*([\d.,]+)", texto) else "0"

    # Novos padrões ajustados ao layout real
    campos_valores = {
        'CN': r"Energia El\u00e9trica kWh\s+\d+\s+[\d,.]+\s+([\d,.]+)",
        'CQ': r"Energia compensada GD I kWh\s+\d+\s+[\d,.]+\s+(-?[\d,.]+)",
        'CV': r"Energia SCEE ISENTA kWh\s+\d+\s+[\d,.]+\s+([\d,.]+)",
        'CT': r"Contrib Ilum Publica Municipal\s+([\d,.]+)",
        'DJ1': r"Multa.*?\s+([\d.,]+)",
        'DJ2': r"Juros.*?\s+([\d.,]+)",
        'IRPJ': r"IRPJ\s+(-?[\d.,]+)",
        'CSLL': r"CSLL\s+(-?[\d.,]+)",
        'PIS': r"PIS.*?(-?[\d.,]+)",
        'COFINS': r"COFINS.*?(-?[\d.,]+)"
    }

    for campo, padrao in campos_valores.items():
        match = re.search(padrao, texto)
        if match:
            resultados[campo] = match.group(1)

    dj1 = limpar_valor(resultados.get('DJ1', '0'))
    dj2 = limpar_valor(resultados.get('DJ2', '0'))
    try:
        resultados['DJ'] = f"{float(dj1) + float(dj2):.2f}".replace('.', ',')
    except:
        resultados['DJ'] = '0'

    desconto = re.search(r"Aplicado desconto de\s+([\d.,]+)\s*%", texto)
    if desconto:
        resultados['fatDescontoFio'] = desconto.group(1).replace('.', ',')

    endereco = re.search(r"\n(.*?)\n(.*?)\n(\d{5}-\d{3}.*?)\n", texto)
    if endereco:
        rua, bairro, cidade = endereco.groups()
        resultados['ENDERECO'] = f"{rua}, {bairro}, {cidade}"

    resultados['cadTarifaCod'] = "3"
    resultados['cadSubGrupoCod'] = "6"
    resultados['fatDataCadastro'] = datetime.now().strftime("%d/%m/%Y")
    resultados['fatDataReferencia'] = datetime.now().replace(day=1).strftime("%d/%m/%Y")

    for h in headers:
        if not resultados[h].strip():
            resultados[h] = "0"

    return [resultados[h] for h in headers]

@app.route('/', methods=['GET', 'POST'])
def index():
    msg = ''
    if request.method == 'POST':
        arquivos = request.files.getlist('pdfs')
        mensagens = []

        if not arquivos or all(f.filename == '' for f in arquivos):
            msg = "Nenhum arquivo selecionado."
            return render_template('index.html', msg=msg)

        for pdf_file in arquivos:
            if not pdf_file.filename.endswith(".pdf"):
                mensagens.append(f"[ERRO] {pdf_file.filename}: N\u00e3o \u00e9 PDF.")
                continue

            save_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
            pdf_file.save(save_path)

            try:
                texto = extrair_texto(save_path)
                linha = extrair_dados_por_regex(texto)
                worksheet.append_row(linha)
                mensagens.append(f"[OK] {pdf_file.filename} processado com sucesso.")
            except Exception as e:
                mensagens.append(f"[ERRO] {pdf_file.filename}: {str(e)}")
            finally:
                os.remove(save_path)

        msg = "\n".join(mensagens)

    return render_template('index.html', msg=msg)

if __name__ == '__main__':
    app.run(debug=True)