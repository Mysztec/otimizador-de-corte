import pdfplumber
import re

def analisar_pdf(caminho_pdf):

    folhas = []

    with pdfplumber.open(caminho_pdf) as pdf:

        for numero_pagina, pagina in enumerate(pdf.pages):

            texto = pagina.extract_text()

            if not texto:
                folhas.append({
                    "pagina": numero_pagina,
                    "maior_area": 0
                })
                continue

            linhas = texto.split("\n")

            maior_area = 0

            for linha in linhas:

                linha = linha.strip()

                # IGNORAR linhas técnicas
                if any(palavra in linha for palavra in [
                    "Dimensão:",
                    "Cliente:",
                    "Projeto:",
                    "Chapa",
                    "Descrição:",
                    "Material:",
                    "Fornecedor:",
                    "Aproveitamento:",
                    "Configurações",
                    "Data/Hora"
                ]):
                    continue

                # Regex para pegar apenas peças tipo: AK 20 x 209,5
                padrao = r'^[A-Z]{1,3}\s+(\d+[,.]?\d*)\s*x\s*(\d+[,.]?\d*)'

                match = re.search(padrao, linha)

                if match:
                    largura = float(match.group(1).replace(",", "."))
                    altura = float(match.group(2).replace(",", "."))

                    area = largura * altura

                    if area > maior_area:
                        maior_area = area

            folhas.append({
                "pagina": numero_pagina,
                "maior_area": maior_area
            })

    return folhas