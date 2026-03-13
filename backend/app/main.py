from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List
import shutil
import os
import zipfile
import pdfplumber
import re

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfgen import canvas

from app.services.pdf_analyzer import analisar_pdf
from app.services.pdf_reorder import reorganizar_pdf

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app.mount(
    "/static",
    StaticFiles(directory=os.path.join(FRONTEND_DIR, "static")),
    name="static"
)

templates = Jinja2Templates(
    directory=os.path.join(FRONTEND_DIR, "templates")
)

# =========================================================
# NUMERAÇÃO DE PÁGINA
# =========================================================

class NumeracaoCanvas(canvas.Canvas):

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):

        num_pages = len(self._saved_page_states)

        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)

        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):

        self.setFont("Helvetica", 9)

        texto = f"Página {self._pageNumber} de {page_count}"

        self.drawRightString(self._pagesize[0] - 30, 15, texto)


# =========================================================
# FUNÇÕES
# =========================================================

def extrair_grupo_codigo(linha):

    try:

        descricao = str(linha[0])

        match = re.search(r'Cod\.:\s*([A-Za-z0-9]+)', descricao)

        if not match:
            return 0

        codigo = match.group(1)

        if "B" in codigo:
            return int(codigo[0])

        return int(codigo[-1])

    except:
        return 0


def obter_celula_segura(linha, indice):

    return linha[indice] if indice < len(linha) else ''


def extrair_cabecalho_limpo(pdf_path):

    cliente = "Não identificado"
    projeto = "Não identificado"
    data_hora = "Não identificada"

    with pdfplumber.open(pdf_path) as pdf:

        texto = pdf.pages[0].extract_text() or ""

        texto_limpo = re.sub(r'\s+', ' ', texto)

        match_cliente = re.search(
            r'Cliente:\s*(.*?)(?=Projeto:)',
            texto_limpo,
            re.IGNORECASE
        )

        if match_cliente:
            cliente = match_cliente.group(1).strip()

        match_data = re.search(
            r'Data/Hora:\s*(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})',
            texto_limpo,
            re.IGNORECASE
        )

        if match_data:
            data_hora = match_data.group(1).strip()

        match_projeto = re.search(
            r'Projeto:\s*(.*?)(?=Chapas\b|Sobras\b|Descrição\b|$)',
            texto_limpo,
            re.IGNORECASE
        )

        if match_projeto:
            projeto = re.sub(r'\s+', ' ', match_projeto.group(1)).strip()

    return cliente, projeto, data_hora


def criar_desenhar_cabecalho(cliente, projeto, data_hora):

    def desenhar_cabecalho(canvas, doc):

        canvas.saveState()

        canvas.setFont('Helvetica-Bold', 16)

        canvas.drawCentredString(
            doc.pagesize[0] / 2.0,
            doc.pagesize[1] - 35,
            "Lista de Compras"
        )

        canvas.setFont('Helvetica-Bold', 10)

        canvas.drawString(
            doc.leftMargin,
            doc.pagesize[1] - 60,
            "Cliente:"
        )

        canvas.setFont('Helvetica', 10)

        canvas.drawString(
            doc.leftMargin + 45,
            doc.pagesize[1] - 60,
            cliente
        )

        canvas.setFont('Helvetica-Bold', 10)

        canvas.drawRightString(
            doc.pagesize[0] - doc.rightMargin - 120,
            doc.pagesize[1] - 60,
            "Data/Hora:"
        )

        canvas.setFont('Helvetica', 10)

        canvas.drawRightString(
            doc.pagesize[0] - doc.rightMargin,
            doc.pagesize[1] - 60,
            data_hora
        )

        canvas.setFont('Helvetica-Bold', 10)

        canvas.drawString(
            doc.leftMargin,
            doc.pagesize[1] - 75,
            "Projeto:"
        )

        canvas.setFont('Helvetica', 10)

        canvas.drawString(
            doc.leftMargin + 45,
            doc.pagesize[1] - 75,
            projeto
        )

        canvas.line(
            doc.leftMargin,
            doc.pagesize[1] - 85,
            doc.pagesize[0] - doc.rightMargin,
            doc.pagesize[1] - 85
        )

        canvas.restoreState()

    return desenhar_cabecalho


# =========================================================
# ORGANIZAR SOBRAS
# =========================================================

def organizar_pdf_sobras(caminho_entrada, caminho_saida):

    cabecalho_mestre = None
    dados_mestre = []

    indices_desejados = [0, 1, 2, 3, 6, 7, 8]

    cliente, projeto, data_hora = extrair_cabecalho_limpo(caminho_entrada)

    with pdfplumber.open(caminho_entrada) as pdf:

        for pagina in pdf.pages:

            tabelas = pagina.extract_tables()

            for tabela in tabelas:

                tabela_limpa = []

                for linha in tabela:

                    linha_limpa = [
                        str(c).replace('\n', ' ').strip() if c else ''
                        for c in linha
                    ]

                    if any(linha_limpa):
                        tabela_limpa.append(linha_limpa)

                if not tabela_limpa:
                    continue

                cabecalho_tabela_str = str(tabela_limpa[0]).lower()

                if 'descrição' in cabecalho_tabela_str:

                    if cabecalho_mestre is None:

                        cabecalho_original = tabela_limpa[0]

                        cabecalho_mestre = [
                            obter_celula_segura(cabecalho_original, i)
                            for i in indices_desejados
                        ]

                    for linha in tabela_limpa[1:]:

                        linha_filtrada = [
                            obter_celula_segura(linha, i)
                            for i in indices_desejados
                        ]

                        dados_mestre.append(linha_filtrada)

    # 244x122 primeiro

    primeiro_item = None

    for item in dados_mestre:

        largura = str(item[2])
        altura = str(item[3])

        if "244" in largura and "122" in altura:

            primeiro_item = item
            break

    if primeiro_item:
        dados_mestre.remove(primeiro_item)

    dados_organizados = sorted(
        dados_mestre,
        key=extrair_grupo_codigo
    )

    if primeiro_item:
        dados_organizados.insert(0, primeiro_item)

    largura_pagina, altura_pagina = A4

    margem_esq = 30
    margem_dir = 30

    largura_util = largura_pagina - margem_esq - margem_dir

    larguras_secundarias = [50, 50, 60, 40, 40, 40]

    largura_descricao = largura_util - sum(larguras_secundarias)

    larguras_colunas = [largura_descricao] + larguras_secundarias

    doc = SimpleDocTemplate(
        caminho_saida,
        pagesize=A4,
        rightMargin=margem_dir,
        leftMargin=margem_esq,
        topMargin=95,
        bottomMargin=40
    )

    elementos = []

    estilo_descricao = ParagraphStyle(
        name='Descricao',
        fontName='Helvetica',
        fontSize=8,
        alignment=TA_LEFT
    )

    tabela_final = [cabecalho_mestre]

    for linha in dados_organizados:

        texto = linha[0]

        texto = re.sub(
            r'(Cod\.:\s*)([A-Za-z0-9]+)',
            r'\1<b><font size="12">\2</font></b>',
            texto
        )

        linha[0] = Paragraph(texto, estilo_descricao)

        tabela_final.append(linha)

    tabela_pdf = Table(
        tabela_final,
        colWidths=larguras_colunas,
        repeatRows=1
    )

    tabela_pdf.setStyle(TableStyle([

        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),

        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),

        ('FONTSIZE', (0, 0), (-1, 0), 9),

        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),

        ('FONTSIZE', (0, 1), (-1, -1), 8),

        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),

    ]))

    elementos.append(tabela_pdf)

    rotina_cabecalho = criar_desenhar_cabecalho(
        cliente,
        projeto,
        data_hora
    )

    doc.build(
        elementos,
        onFirstPage=rotina_cabecalho,
        onLaterPages=rotina_cabecalho,
        canvasmaker=NumeracaoCanvas
    )


# =========================================================
# ROTAS
# =========================================================

@app.get("/")
def home(request: Request):

    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )


# CUTPRO
@app.post("/organizar_corte")
async def organizar_corte(arquivos: List[UploadFile] = File(...)):

    if not arquivos or not arquivos[0].filename:
        return {"erro": "Nenhum arquivo enviado"}

    arquivos_processados = []

    for file in arquivos:

        caminho_upload = os.path.join(
            UPLOAD_DIR,
            file.filename
        )

        with open(caminho_upload, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        folhas = analisar_pdf(caminho_upload)

        folhas_ordenadas = sorted(
            folhas,
            key=lambda f: f["maior_area"],
            reverse=True
        )

        nome_base, extensao = os.path.splitext(file.filename)

        nome_download = f"{nome_base} organizado{extensao}"

        caminho_saida = os.path.join(
            OUTPUT_DIR,
            nome_download
        )

        reorganizar_pdf(
            caminho_upload,
            caminho_saida,
            folhas_ordenadas
        )

        arquivos_processados.append(
            (caminho_saida, nome_download)
        )

    if len(arquivos_processados) == 1:

        return FileResponse(
            arquivos_processados[0][0],
            filename=arquivos_processados[0][1],
            media_type="application/pdf"
        )

    else:

        zip_filename = "Lotes organizados.zip"

        zip_filepath = os.path.join(
            OUTPUT_DIR,
            zip_filename
        )

        with zipfile.ZipFile(
            zip_filepath,
            'w',
            zipfile.ZIP_DEFLATED
        ) as zipf:

            for caminho, nome in arquivos_processados:
                zipf.write(caminho, arcname=nome)

        return FileResponse(
            zip_filepath,
            filename=zip_filename,
            media_type="application/zip"
        )


# SOBRAS
@app.post("/organizar_sobras")
async def processar_sobras(arquivos: List[UploadFile] = File(...)):

    if not arquivos or not arquivos[0].filename:
        return {"erro": "Nenhum arquivo enviado"}

    arquivos_processados = []

    for arquivo_pdf in arquivos:

        nome_base, extensao = os.path.splitext(arquivo_pdf.filename)

        nome_download = f"{nome_base} organizado{extensao}"

        caminho_entrada = os.path.join(
            UPLOAD_DIR,
            "entrada_sobras_" + arquivo_pdf.filename
        )

        caminho_saida = os.path.join(
            OUTPUT_DIR,
            nome_download
        )

        with open(caminho_entrada, "wb") as buffer:
            shutil.copyfileobj(arquivo_pdf.file, buffer)

        organizar_pdf_sobras(
            caminho_entrada,
            caminho_saida
        )

        arquivos_processados.append(
            (caminho_saida, nome_download)
        )

    if len(arquivos_processados) == 1:

        return FileResponse(
            arquivos_processados[0][0],
            filename=arquivos_processados[0][1],
            media_type="application/pdf"
        )

    else:

        zip_filename = "Lotes organizados.zip"

        zip_filepath = os.path.join(
            OUTPUT_DIR,
            zip_filename
        )

        with zipfile.ZipFile(
            zip_filepath,
            'w',
            zipfile.ZIP_DEFLATED
        ) as zipf:

            for caminho, nome in arquivos_processados:
                zipf.write(caminho, arcname=nome)

        return FileResponse(
            zip_filepath,
            filename=zip_filename,
            media_type="application/zip"
        )