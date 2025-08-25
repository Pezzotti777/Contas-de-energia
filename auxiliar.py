import pdfplumber
import os

# Caminho do PDF
CAMINHO_PDF = r"C:\Users\enzo.pezzotti\Desktop\Contas de energia\Contas de energia\Recebido\BB_931224.pdf"

def extrair_texto_e_posicoes(pdf_path):
    if not os.path.exists(pdf_path):
        print(f"[ERRO] Arquivo não encontrado: {pdf_path}")
        return

    with pdfplumber.open(pdf_path) as pdf:
        total_paginas = len(pdf.pages)
        print(f"[INFO] PDF possui {total_paginas} páginas")

        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            print(f"\n[DEBUG] Página {numero_pagina} --------------------------")

            palavras = pagina.extract_words()
            if not palavras:
                print("[AVISO] Nenhuma palavra encontrada nesta página.")
                continue

            for palavra in palavras:
                texto = palavra.get("text", "").strip()
                x0 = round(palavra["x0"], 2)
                y0 = round(palavra["top"], 2)
                x1 = round(palavra["x1"], 2)
                y1 = round(palavra["bottom"], 2)
                print(f"[P{numero_pagina}] '{texto}' → ({x0}, {y0}, {x1}, {y1})")

if __name__ == "__main__":
    extrair_texto_e_posicoes(CAMINHO_PDF)
