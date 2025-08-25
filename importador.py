from flask import Flask, render_template, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pdfplumber
import os
import re
from datetime import datetime

# ——— DETECÇÃO DE TIPO DE CONTA ———
def detectar_tipo_conta_inicial(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        texto = pdf.pages[0].extract_text() or ""
        texto_upper = texto.upper()

        if "THS VERDE A4" in texto_upper or ("GRUPO A" in texto_upper and "THS" in texto_upper and "VERDE" in texto_upper):
            return "THS_VERDE_A4"
        elif "A4 VERDE" in texto_upper or ("GRUPO A" in texto_upper and "TUSD" in texto_upper):
            return "A4_VERDE"
        elif "B3" in texto_upper or "SUBGRUPO B3" in texto_upper or "GRUPO B" in texto_upper:
            return "B3"
        else:
            return "DESCONHECIDO"

def detectar_multa_ou_padrao(page, resultados=None) -> dict:
    """
    Verifica se existem as palavras 'multa', 'juros' ou 'correção' dentro da área de energia (x0 <= 305),
    e ajusta o deslocamento vertical conforme a quantidade:
      - 1 termo: -10
      - 2 termos:  0 (padrão)
      - 3 termos: +10

    Também verifica se o campo CT (fatDescCsllValRetImposto) tem o mesmo valor que DJ (fatMultasDiversas),
    e se for um imposto retido, considera que DJ não existe e o zera — somente para contas B3 Convencional.
    """
    try:
        palavras = page.extract_words()
        termos_detectados = set()
        termos_alvo = ['multa', 'juros', 'correção']

        for p in palavras:
            texto = p['text'].strip().lower()
            x0 = float(p['x0'])
            if x0 <= 305:
                for termo in termos_alvo:
                    if termo in texto:
                        termos_detectados.add(termo)
                        print(f"[DEBUG] Palavra '{termo}' detectada dentro da área de energia (x={x0:.2f}) ✔️")

        n_termos = len(termos_detectados)
        deslocamento = 0
        if n_termos == 1:
            deslocamento = -10
        elif n_termos == 3:
            deslocamento = 10

        # Coordenadas padrão
        coordenadas = {
            'DG': (305.55, 290.57, 338.65, 300.19),
            'CT': (305.15, 300.18, 338.66, 309.79),
            'CQ': (305.15, 309.77, 338.66, 319.39),
            'CN': (305.15, 319.38, 338.66, 328.99),
            'CV': (305.15, 328.98, 338.66, 338.59),
            'DC': (305.55, 348.18, 338.65, 357.79),
        }

        if deslocamento != 0:
            print(f"[DEBUG] Aplicando deslocamento de {deslocamento:+}px em Y nas coordenadas devido a {n_termos} termo(s) encontrado(s)...")
            for k in coordenadas:
                x0, y0, x1, y1 = coordenadas[k]
                coordenadas[k] = (x0, y0 + deslocamento, x1, y1 + deslocamento)
                print(f"[DEBUG] {k}: y0={y0:.2f} → {y0 + deslocamento:.2f}, y1={y1:.2f} → {y1 + deslocamento:.2f}")

        # Limpando DJ apenas para contas B3 Convencional
        if resultados and resultados.get("cadSubGrupoCod") == "6":
            ct_val = resultados.get("CT", "").replace(".", "").replace(",", ".").strip()
            dj_val = resultados.get("DJ", "").replace(".", "").replace(",", ".").strip()
            if ct_val and dj_val and ct_val == dj_val:
                for imposto in ["CSLL", "PIS", "COFINS", "IRPJ"]:
                    imposto_val = resultados.get(imposto, "").replace(".", "").replace(",", ".").strip()
                    if imposto_val == ct_val and ct_val != "0":
                        print(f"[INFO] [B3] CT = DJ = imposto retido ({imposto}) → limpando DJ")
                        resultados["DJ"] = "0"
                        resultados["DJ1"] = "0"
                        resultados["DJ2"] = "0"
                        break

        return coordenadas

    except Exception as e:
        print(f"[ERRO] na detecção ou ajuste de coordenadas por multa/juros/correção: {e}")
        return {}


def extrair_fatDescontoFio(texto: str) -> str:
    """
    Extrai o valor do desconto em porcentagem após o trecho 'Aplicado desconto de' para preencher fatDescontoFio.
    Exemplo: 'Aplicado desconto de 49,62 %' → retorna '49,62'
    """
    padrao = r"Aplicado desconto de\s+([\d.,]+)\s*%"
    match = re.search(padrao, texto)
    if match:
        return match.group(1).replace(",", ".")  # ou mantenha vírgula se preferir
    return ""


# ——— PARSER TUSD A4 VERDE (MÓDULO ATUALIZADO) ———
def extrair_por_regras_a4_verde(pdf_path: str) -> list:
    resultados = {h: "" for h in headers}

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        print(f"[DEBUG] A4 Verde → {total} pág.")
        print(">>> CHAVES A4:", list(COORDENADAS_A4.keys()))

        # ——— Texto completo para regex ———
        texto_completo = "\n".join([page.extract_text() or "" for page in pdf.pages])
        texto0 = pdf.pages[0].extract_text() or ""

        # ——— Regex: Desconto em % ———
        resultados["fatDescontoFio"] = extrair_fatDescontoFio(texto_completo)

        # ——— Detectar “livre” para preencher fatDescontoFioKWh (DK) ———
        try:
            page_livre = pdf.pages[0]
            bbox_livre = (296.4, 181.23, 360.0, 193.59)
            texto_livre = page_livre.within_bbox(bbox_livre).extract_text() or ""
            if "livre" in texto_livre.lower():
                resultados["fatDescontoFioKWh"] = "46,45"
            else:
                resultados["fatDescontoFioKWh"] = "0"
            print(f"[DEBUG] fatDescontoFioKWh: '{texto_livre.strip()}' → {resultados['fatDescontoFioKWh']}")
        except Exception as e:
            print(f"[ERRO] ao verificar fatDescontoFioKWh: {e}")
            resultados["fatDescontoFioKWh"] = "0"

        # ——— Preencher DJ1 e DJ2 ———
        valores_temporarios = {}
        for dj_tag in ["DJ1", "DJ2"]:
            if dj_tag in COORDENADAS_A4:
                pg, x0, y0, x1, y1 = COORDENADAS_A4[dj_tag]
                if pg <= total:
                    texto = extrair_na_bbox(pdf.pages[pg - 1], x0, y0, x1, y1)
                    valores_temporarios[dj_tag] = texto.strip()
                    print(f"[DEBUG] {dj_tag}: '{texto.strip()}'")

        # ——— Loop por headers ———
        for idx, header_name in enumerate(headers, start=1):
            letra = get_column_letter(idx)

            # CASO: fatMultasDiversas = DJ1 + DJ2
            if header_name == "fatMultasDiversas":
                valor1 = valores_temporarios.get("DJ1", "0").replace(",", ".")
                valor2 = valores_temporarios.get("DJ2", "0").replace(",", ".")
                try:
                    soma = float(valor1) + float(valor2)
                    resultados[header_name] = "{:.2f}".format(soma).replace(".", ",")
                    print(f"[A4] fatMultasDiversas (DJ1 + DJ2): {valor1} + {valor2} = {resultados[header_name]}")
                except Exception as e:
                    print(f"[A4] fatMultasDiversas erro: {e}")
                    resultados[header_name] = "0"
                continue

            # Coordenadas gerais
            if letra in COORDENADAS_A4:
                pg, x0, y0, x1, y1 = COORDENADAS_A4[letra]
                if pg > total:
                    print(f"[ERRO] Página {pg} não existe para {header_name}")
                    continue
                page = pdf.pages[pg - 1]
                raw = extrair_na_bbox(page, x0, y0, x1, y1)
                clean = limpar_valor(header_name, raw)
                resultados[header_name] = clean
                valores_temporarios[letra] = clean
                print(f"[A4] {header_name} (letra {letra}): raw='{raw}' → clean='{clean}' coords=({pg}, {x0}, {y0}, {x1}, {y1})")

        # ——— Regex: endereço, impostos e nota fiscal ———
        resultados["ENDERECO"] = extrair_endereco_completo(texto0)
        resultados.update(extrair_impostos_retidos_por_regex(texto0))
        resultados["NOTAFISCAL"] = extrair_numero_nota_fiscal(texto0)

        # ——— Datas auxiliares ———
        resultados["fatDataCadastro"]   = datetime.now().strftime("%d/%m/%Y")
        resultados["fatDataReferencia"] = datetime.now().replace(day=1).strftime("%d/%m/%Y")

        # ——— Códigos fixos A4 Verde ———
        resultados["cadTarifaCod"] = "1"
        resultados["cadSubGrupoCod"] = "5"
        # ——— Preencher concCod com 22 se for da CEMIG ———
        if "cemig" in texto0.lower():
            resultados["concCod"] = "22"
            print("[DEBUG] concCod = 22 (Detectado CEMIG)")
        else:
            resultados["concCod"] = "0"  # Ou outro valor padrão, se desejar

        # ——— Zerar campos não aplicáveis ———
        campos_zero = [
            "fatConFPontaInjetadoValorReais",
            "fatConPontaInjetadoUsina",
            "fatConPontaInjetadoUsinaSaldoAcumulado",
            "fatConFPontaInjetadoUsina",
            "fatConFPontaInjetadoUsinaSaldoAcumulado",
            "fatDemandasDevolucaoPtaValorReais",
            "fatValBandeira"
        ]
        for campo in campos_zero:
            resultados[campo] = "0"

    # ——— Código de barras (BA) ———
    if "fatCodigoBarras" in resultados:
        cb = re.search(r"\d{11}-\d\s+\d{11}-\d\s+\d{11}-\d\s+\d{11}-\d", texto_completo)
        if cb:
            resultados["fatCodigoBarras"] = cb.group(0)
            print(f"[A4] Código de Barras encontrado: {resultados['fatCodigoBarras']}")
        else:
            print("[A4] Código de Barras não encontrado")

    # ——— Substituir campos vazios por "0" ———
    for h in headers:
        if not resultados[h].strip():
            resultados[h] = "0"

    return [resultados[h] for h in headers]


# ——— CONFIGURAÇÃO GOOGLE SHEETS ———
UPLOAD_FOLDER = 'uploads'
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/170LPTCD-_9Dk6oOt6D2SGNr7eQQaDFS8h4SuVW92N1c/edit?usp=sharing"
ABA = "CONTAS"
CREDENCIAL = "client_secret.json"

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENCIAL, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url(PLANILHA_URL)
worksheet = sheet.worksheet(ABA)
headers = worksheet.row_values(1)

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# ——— COORDENADAS B3 ———
COORDENADAS_B3 = {
    'A':  (142.5, 144.3, 225.9, 164.95),
    'B':  (380.0, 91.0, 460.0, 100.0),
    'C':  (353.45, 750.55, 403.49, 764.32),
    'D':  (463.9, 750.55, 507.26, 764.32),
    'G':  (413.9, 190.92, 436.45, 203.29),
    'H':  (456.45, 190.92, 479.0, 203.29),
    'I':  (303.6, 270.98, 338.66, 280.59),
    'J':  (355.0, 181.23, 390.0, 193.59),
    'K':  (296.4, 181.23, 360.0, 193.59),
    'Z':  (540.05, 453.18, 565.38, 462.79),
    'AM': (475.8, 678.9, 524.77, 689.89),
    'AN': (478.05, 686.35, 524.79, 697.34),
    'AO': (478.05, 693.8, 524.79, 704.79),
    'AP': (311.8, 678.9, 363.0, 689.89),
    'AR': (15.0, 89.0, 85.0, 100.0),
    'AU': (145.6, 117.0, 222.75, 127.0),
    'CA': (301.25, 261.38, 338.68, 270.99),
    'CC': (416.0, 529.09, 440.0, 538.43),
    'CE': (340.1, 529.09, 366.0, 538.43),
    'CL': (407.15, 693.8, 453.89, 704.79),
    'CO': (409.4, 693.8, 453.91, 704.79),
    'CR': (407.15, 678.9, 453.89, 689.89),
}

# ——— MAPEAMENTO POR NOME DO HEADER PARA TUSD A4 VERDE ———
# ——— COORDENADAS PARA A4 VERDE ———
# chave = letra da coluna, valor = (página, x0, y0, x1, y1)
COORDENADAS_A4 = {
    'A': (1, 142.5, 144.3, 225.9, 164.95),
    'AC': (1, 200.5, 272.9, 221.0, 280.0),
    'AD': (1, 199.6, 263.4, 221.0, 270.4),
    'AG': (3, 269.7, 347.6, 277.5, 354.6),
    'AH': (3, 268.8, 295.2, 280.4, 302.2),
    'AJ': (1, 213.3, 292.2, 221.0, 299.2),
    'AK': (1, 209.4, 282.6, 221.0, 289.6),
    'AM': (1, 475.8, 678.9, 524.77, 689.89),
    'AN': (1, 478.05, 686.35, 524.79, 697.34),
    'AO': (1, 478.05, 693.8, 524.79, 704.79),
    'AP': (1, 311.8, 678.9, 363.0, 689.89),
    'AR': (1, 15.0, 89.0, 85.0, 100.0),
    'AU': (1, 145.6, 117.0, 222.75, 127.0),
    'B': (1, 380.0, 91.0, 460.0, 100.0),
    'BN': (1, 297.2, 244.2, 324.5, 251.2),
    'BP': (1, 390.9, 253.8, 412.3, 260.8),
    'BR': (1, 307.0, 301.8, 324.5, 308.8),
    'BS': (1, 297.2, 272.9, 324.5, 279.9),
    'BT': (1, 297.2, 263.4, 324.5, 270.4),
    'BW': (1, 307.0, 292.2, 324.5, 299.2),
    'BX': (1, 303.1, 282.6, 324.5, 289.6),
    'C': (1, 353.45, 750.55, 403.49, 764.32),
    'CA': (1, 301.25, 261.38, 338.68, 270.99),
    'CC': (1, 416.0, 529.09, 440.0, 538.43),
    'CE': (1, 340.1, 529.09, 366.0, 538.43),
    'CG': (1, 300.7, 359.4, 324.5, 366.4),
    'CL': (1, 407.15, 686.35, 453.91, 697.34),
    'CN': (2, 297.2, 263.4, 340.0, 270.4),
    'CQ': (2, 297.2, 253.8, 340.0, 260.8),
    'CT': (2, 297.2, 244.2, 340.0, 251.2),
    'CV': (2, 297.2, 272.9, 340.0, 279.9),
    'CO': (1, 409.4, 693.8, 453.91, 704.79),
    'CR': (1, 407.15, 678.9, 453.89, 689.89),
    'D': (1, 463.9, 750.55, 507.26, 764.32),
    'DC': (1, 305.55, 338.18, 338.65, 347.79),
    'DG': (1, 305.55, 280.57, 338.65, 290.19),
    'DJ1': (1, 310.9, 330.6, 324.5, 337.6),
    'DJ2': (1, 310.9, 340.2, 324.5, 347.2),
    'DL': (1, 300.7, 369.0, 324.5, 376.0),
    'DP': (1, 307.0, 311.4, 324.5, 318.4),
    'DQ': (1, 307.0, 321.0, 324.5, 328.0),
    'DR': (1, 199.6, 311.4, 221.0, 318.4),
    'DS': (1, 199.6, 321.0, 221.0, 328.0),
    'G': (1, 413.9, 190.92, 436.45, 203.29),
    'H': (1, 456.45, 190.92, 479.0, 203.29),
    'I': (1, 307.0, 349.8, 324.5, 356.8),
    'J': (1, 355.0, 181.23, 390.0, 193.59),
    'K': (1, 296.4, 181.23, 360.0, 193.59),
    'M': (3, 420.0, 255.9, 427.7, 262.9),
    'N': (3, 269.7, 308.3, 277.5, 315.3),
    'O': (3, 269.7, 255.9, 277.5, 262.9),
    'R': (1, 217.1, 301.8, 221.0, 308.8),
    'T': (3, 270.7, 282.1, 274.6, 289.1),
    'V': (1, 213.3, 244.2, 221.0, 251.2),
    'X': (1, 217.1, 253.8, 221.0, 260.8),
    'Y': (3, 229.3, 598.55, 246.81, 605.55),
    'Z': (3, 227.4, 532.6, 248.8, 539.6),
}

# ——— FUNÇÕES AUXILIARES ———
def get_column_letter(n: int) -> str:
    result = ""
    while n > 0:
        n, r = divmod(n-1, 26)
        result = chr(65 + r) + result
    return result

def extrair_na_bbox(page, x0, y0, x1, y1, margem=1.5) -> str:
    top, bottom = min(y0, y1), max(y0, y1)
    rec = page.within_bbox((x0 - margem, top - margem, x1 + margem, bottom + margem))
    txt = rec.extract_text()
    print(f"[DEBUG] bbox=({x0}, {y0}, {x1}, {y1}) → '{txt.strip() if txt else ''}'")
    return txt.strip() if txt else ""

def diagnosticar_vazios_na_pagina(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            texto = page.extract_text() or ""

            def extrair_palavras_corrigidas(pagina):
                palavras_originais = pagina.extract_words()
                palavras_corrigidas = []
                skip_next = False

                for i, palavra in enumerate(palavras_originais):
                    if skip_next:
                        skip_next = False
                        continue

                    texto = palavra["text"]
                    if i + 1 < len(palavras_originais):
                        proxima = palavras_originais[i + 1]["text"]
                        if texto.endswith(".") and proxima[0].isdigit():
                            # Junta palavras separadas incorretamente como "1." + "736,72"
                            texto += proxima
                            skip_next = True

                    nova_palavra = palavra.copy()
                    nova_palavra["text"] = texto
                    palavras_corrigidas.append(nova_palavra)

                return palavras_corrigidas

            print(f"[Página {i}] Total de palavras detectadas: {len(palavras)}")
            if not texto.strip():
                print("⚠️ Nada extraído com extract_text() — suspeita de imagem.")


def visualizar_bbox(pdf_path, pagina, x0, y0, x1, y1):
    import matplotlib.pyplot as plt
    import pdfplumber

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[pagina - 1]
        im = page.to_image(resolution=150)
        im.draw_rect((x0, y0, x1, y1), stroke="red", fill=None)
        im.annotate((x0, y0, x1, y1), "bbox", stroke="red")
        im.save("debug_output.png")
        print("Salvo em debug_output.png")


# ——— FUNÇÃO DE LIMPEZA PARA VALORES MONETÁRIOS ———
def limpar_valor(campo: str, valor: str) -> str:
    valor = valor.strip()
    if campo == "Instalação":
        return ''.join(filter(str.isdigit, valor))

    if campo == "fatValorFatura":
        valor = valor.replace('R$', '').replace('$', '').strip()
        valor = re.sub(r"[^\d,\.]", "", valor)
        valor = valor.replace('.', '').replace(',', '.')
        try:
            return "{:.2f}".format(float(valor)).replace('.', ',')
        except:
            return "0"

    return valor

def extrair_por_conteudo(texto: str) -> dict:
    resultados = {}
    padrao = r"SALDO ATUAL DE GERAÇÃO:\s*([\d.,]+)\s+kWh\s+FP/\u00danico,\s*([\d.,]+)\s+kWh\s+ponta"
    match = re.search(padrao, texto)
    if match:
        cc = match.group(1).replace(' ', '').replace('kWh', '').strip()
        ce = match.group(2).replace(' ', '').replace('kWh', '').strip()
        if re.match(r'^[\d.,]+$', ce) and re.match(r'^[\d.,]+$', cc):
            resultados['fatConPontaInjetadoUsinaSaldoAcumulado'] = ce
            resultados['fatConFPontaInjetadoUsinaSaldoAcumulado'] = cc
    return resultados

def extrair_impostos_retidos_por_regex(texto: str) -> dict:
    return {
        'fatDescPisPercRetImposto': "0,65",
        'fatDescCofinsPercRetImposto': "3,00",
        'fatDescCsllPercRetImposto': "1,00",
        'fatDescIrpjPercRetImposto': "1,20"
    }

# ——— EXTRAÇÃO DE ENDEREÇO (ROBUSTA) ———
def extrair_endereco_completo(texto: str) -> str:
    endereco = re.search(r"\n(.*?)\n(.*?)\n(\d{5}-\d{3}.*?)\n", texto)
    if endereco:
        rua, bairro, cidade = endereco.groups()
        return f"{rua}, {bairro}, {cidade}"
    return "0"


def extrair_numero_nota_fiscal(texto: str) -> str:
    match = re.search(r'NOTA FISCAL Nº\s*(\d+)', texto)
    return match.group(1) if match else ""


def extrair_por_regras(pdf_path: str) -> list:
    resultados = {h: "" for h in headers}

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        texto_completo = page.extract_text() or ""

        # Detecta e aplica coordenadas de multa, se necessário
        COORDENADAS_MULTA = detectar_multa_ou_padrao(page)

        # AQ = DG + DJ se existir saldo + compensação
        if "Saldo para o próximo mês" in texto_completo and "Compensação FIC mensal" in texto_completo:
            try:
                x0_dg, y0_dg, x1_dg, y1_dg = COORDENADAS_MULTA["DG"]
                x0_dj, y0_dj, x1_dj, y1_dj = COORDENADAS_MULTA["DJ"]

                val_dg = extrair_na_bbox(page, x0_dg, y0_dg, x1_dg, y1_dg).replace('.', '').replace(',', '.')
                val_dj = extrair_na_bbox(page, x0_dj, y0_dj, x1_dj, y1_dj).replace('.', '').replace(',', '.')

                soma = float(val_dg) + float(val_dj)
                resultados["AQ"] = "{:.2f}".format(soma).replace('.', ',')
                resultados["DG"] = "0"
                resultados["DJ"] = "0"
                print(f"[B3] AQ = DG({val_dg}) + DJ({val_dj}) = {resultados['AQ']}")
            except Exception as e:
                print(f"[ERRO] ao calcular AQ = DG + DJ: {e}")

        # Validação do subgrupo
        try:
            subgrupo = extrair_na_bbox(page, 355.0, 181.23, 390.0, 193.59).strip()
        except:
            subgrupo = ""
        if subgrupo != "B3":
            raise ValueError(f"Subgrupo inválido: '{subgrupo}'. Apenas B3 conv. são aceitas.")

        mapa_letra_para_header = {get_column_letter(i): h for i, h in enumerate(headers, start=1)}

        # Extração campo a campo
        for idx, campo in enumerate(headers, start=1):
            if campo == "ENDERECO":
                continue

            letra = get_column_letter(idx)

            if letra in COORDENADAS_B3:
                x0, y0, x1, y1 = COORDENADAS_B3[letra]
            elif letra in COORDENADAS_MULTA:
                x0, y0, x1, y1 = COORDENADAS_MULTA[letra]
                print(f"[DEBUG] {campo} com coordenada ajustada (multa) → ({x0}, {y0}, {x1}, {y1})")
            else:
                continue  # pula se não estiver em nenhuma

            raw = extrair_na_bbox(page, x0, y0, x1, y1)
            resultados[campo] = limpar_valor(campo, raw)

        # Campos dinâmicos
        resultados.update(extrair_por_conteudo(texto_completo))
        resultados.update(extrair_impostos_retidos_por_regex(texto_completo))
        resultados["NOTAFISCAL"] = extrair_numero_nota_fiscal(texto_completo)

        # Concessionária
        if "CEMIG" in texto_completo.upper():
            resultados["concCod"] = "22"

        # Datas
        resultados["fatDataCadastro"] = datetime.now().strftime("%d/%m/%Y")
        resultados["fatDataReferencia"] = datetime.now().replace(day=1).strftime("%d/%m/%Y")

        # Injetado AY, AZ, CD
        try:
            valor_injetado = extrair_na_bbox(page, 205.65, 261.38, 230.98, 270.99)
            for c in ["fatConFPontaInjetadoRegistrado", "fatConFPontaInjetadoFaturado", "fatConFPontaInjetadoUsina"]:
                if c in resultados:
                    resultados[c] = limpar_valor(c, valor_injetado)
        except:
            pass

        # fatConFPontaIndValorReais (BT) = soma de duas regiões
        try:
            v1 = extrair_na_bbox(page, 301.65, 242.18, 338.67, 251.79).replace('.', '').replace(',', '.')
            v2 = extrair_na_bbox(page, 301.65, 251.77, 338.67, 261.39).replace('.', '').replace(',', '.')
            soma = float(v1) + float(v2)
            resultados["fatConFPontaIndValorReais"] = "{:.2f}".format(soma).replace('.', ',')
        except:
            pass

        # Código de barras (BA)
        if "fatCodigoBarras" in resultados:
            cb = re.search(r"\d{11}-\d\s+\d{11}-\d\s+\d{11}-\d\s+\d{11}-\d", texto_completo)
            if cb:
                resultados["fatCodigoBarras"] = cb.group(0)

    # Pós-processamento
    for campo in resultados:
        if resultados[campo] == "":
            resultados[campo] = "0"

    # Zera DJ se for igual a DG
    dg_val = resultados.get("DG", "").replace('.', '').replace(',', '.').strip()
    dj_val = resultados.get("DJ", "").replace('.', '').replace(',', '.').strip()
    if dg_val and dj_val and dg_val == dj_val:
        print(f"[INFO] DJ = DG ({dg_val}) → limpando DJ")
        resultados["DJ"] = "0"

    # Zera DJ se a palavra 'correção' não estiver no texto
    if "correção" not in texto_completo.lower():
        print("[INFO] Palavra 'correção' não encontrada → limpando DJ")
        resultados["DJ"] = "0"

    resultados["fatConFPontaIndFaturado"] = resultados.get("fatConFPontaIndRegistrado", "0")

    return [resultados[h] for h in headers]

# ———ROTA FLASK COM SUPORTE A B3 E A4 VERDE ———
@app.route('/', methods=['GET', 'POST'])
def index():
    msg = ''
    if request.method == 'POST':
        arquivos = request.files.getlist('pdfs')
        mensagens = []

        if not arquivos or all(f.filename == '' for f in arquivos):
            msg = "Nenhum arquivo foi selecionado."
            return render_template('index.html', msg=msg)

        for pdf_file in arquivos:
            if not pdf_file.filename.endswith(".pdf"):
                mensagens.append(f"[ERRO] {pdf_file.filename}: Arquivo não é PDF.")
                continue

            save_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
            pdf_file.save(save_path)

            try:
                tipo_detectado = detectar_tipo_conta_inicial(save_path)

                if tipo_detectado == "B3":
                    linha = extrair_por_regras(save_path)
                    tarifa_cod = "3"
                    subgrupo_cod = "6"
                elif tipo_detectado == "A4_VERDE":
                    linha = extrair_por_regras_a4_verde(save_path)
                    tarifa_cod = "1"
                    subgrupo_cod = "5"
                elif tipo_detectado == "THS_VERDE_A4":
                    linha = extrair_por_regras_ths_verde_a4(save_path)
                    tarifa_cod = "2"
                    subgrupo_cod = "7"
                else:
                    mensagens.append(f"[ERRO] {pdf_file.filename}: Tipo de conta não suportado.")
                    continue

                # Atualiza os campos se existirem no header
                if "cadTarifaCod" in headers:
                    idx_tarifa = headers.index("cadTarifaCod")
                    linha[idx_tarifa] = tarifa_cod
                if "cadSubGrupoCod" in headers:
                    idx_subgrupo = headers.index("cadSubGrupoCod")
                    linha[idx_subgrupo] = subgrupo_cod

                worksheet.append_row(linha)
                mensagens.append(f"[OK] {pdf_file.filename} ({tipo_detectado}) processado com sucesso.")
            except Exception as e:
                mensagens.append(f"[ERRO] {pdf_file.filename}: {str(e)}")
            finally:
                os.remove(save_path)

        msg = "\n".join(mensagens)

    return render_template('index.html', msg=msg)



if __name__ == "__main__":
    app.run(debug=True)
