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
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# Importações dos seus serviços originais do FastAPI
from app.services.pdf_analyzer import analisar_pdf
from app.services.pdf_reorder import reorganizar_pdf

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")

# Mude a variável OUTPUT_DIR abaixo se quiser que o Python salve automaticamente
# numa pasta específica do seu computador.
# Exemplo: OUTPUT_DIR = r"C:\Users\Lenovo\OneDrive\Desktop\Meus PDFs Prontos"
# Se deixar como está, ele salva na pasta "outputs" do projeto.
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

# =====================================================================
# FUNÇÕES DO ORGANIZADOR DE SOBRAS
# =====================================================================

def extrair_ultimo_digito(linha):
    try:
        descricao = str(linha[0])
        match = re.search(r'Cod\.:\s*([A-Za-z0-9]+)', descricao)
        if match:
            codigo = match.group(1)
            return codigo[-1]
        return '0'
    except:
        return '0'

def obter_celula_segura(linha, indice):
    return linha[indice] if indice < len(linha) else ''

def extrair_cabecalho_limpo(pdf_path):
    cliente = "Não identificado"
    projeto = "Não identificado"
    data_hora = "Não identificada"

    with pdfplumber.open(pdf_path) as pdf:
        texto = pdf.pages[0].extract_text() or ""
        
        texto_limpo = re.sub(r'\s+', ' ', texto)

        match_cliente = re.search(r'Cliente:\s*(.*?)(?=Projeto:)', texto_limpo, re.IGNORECASE)
        if match_cliente:
            cliente = match_cliente.group(1).strip()

        match_data = re.search(r'Data/Hora:\s*(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})', texto_limpo, re.IGNORECASE)
        str_data_completa = ""
        if match_data:
            data_hora = match_data.group(1).strip()
            str_data_completa = match_data.group(0)

        match_projeto = re.search(r'Projeto:\s*(.*?)(?=Chapas\b|Sobras\b|Descrição\b|$)', texto_limpo, re.IGNORECASE)
        if match_projeto:
            projeto_bruto = match_projeto.group(1)
            
            if str_data_completa:
                projeto_bruto = projeto_bruto.replace(str_data_completa, '')
            
            projeto_bruto = re.sub(r'Data/Hora:\s*', '', projeto_bruto, flags=re.IGNORECASE)
            projeto = re.sub(r'\s+', ' ', projeto_bruto).strip()

    return cliente, projeto, data_hora

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
                    linha_limpa = [str(celula).replace('\n', ' ').strip() if celula else '' for celula in linha]
                    if any(linha_limpa):
                        tabela_limpa.append(linha_limpa)
                        
                if not tabela_limpa:
                    continue
                
                if cabecalho_mestre is None:
                    cabecalho_original = tabela_limpa[0]
                    cabecalho_mestre = [obter_celula_segura(cabecalho_original, i) for i in indices_desejados]
                
                inicio_dados = 1 if 'Descrição' in tabela_limpa[0][0] else 0
                
                for linha in tabela_limpa[inicio_dados:]:
                    linha_filtrada = [obter_celula_segura(linha, i) for i in indices_desejados]
                    dados_mestre.append(linha_filtrada)

    dados_organizados = sorted(dados_mestre, key=extrair_ultimo_digito)

    largura_pagina, altura_pagina = landscape(A4)
    margem_esq = 30
    margem_dir = 30
    largura_util = largura_pagina - margem_esq - margem_dir
    
    larguras_secundarias = [60, 60, 70, 45, 45, 45]
    largura_descricao = largura_util - sum(larguras_secundarias)
    larguras_colunas = [largura_descricao] + larguras_secundarias

    doc = SimpleDocTemplate(caminho_saida, pagesize=landscape(A4), rightMargin=margem_dir, leftMargin=margem_esq, topMargin=30, bottomMargin=30)
    elementos = []
    
    estilo_titulo = ParagraphStyle(name='Titulo', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, spaceAfter=15)
    estilo_info = ParagraphStyle(name='Info', fontName='Helvetica', fontSize=10, alignment=TA_LEFT)
    estilo_info_dir = ParagraphStyle(name='InfoDir', fontName='Helvetica', fontSize=10, alignment=TA_RIGHT)
    
    elementos.append(Paragraph("Lista de Compras", estilo_titulo))
    
    p_cliente = Paragraph(f"<b>Cliente:</b> {cliente}", estilo_info)
    p_data = Paragraph(f"<b>Data/Hora:</b> {data_hora}", estilo_info_dir)
    p_projeto = Paragraph(f"<b>Projeto:</b> {projeto}", estilo_info)

    tabela_cabecalho = Table([
        [p_cliente, p_data],
        [p_projeto, ""]
    ], colWidths=[largura_util/2, largura_util/2])
    
    tabela_cabecalho.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('SPAN', (0, 1), (1, 1)),
    ]))

    elementos.append(tabela_cabecalho)
    elementos.append(Spacer(1, 15))

    estilo_descricao = ParagraphStyle(name='EstiloDescricao', fontName='Helvetica', fontSize=8, alignment=TA_LEFT)
    
    tabela_final = [cabecalho_mestre]
    for linha in dados_organizados:
        texto_descricao = linha[0]
        texto_destacado = re.sub(r'(Cod\.:\s*)([A-Za-z0-9]+)', r'\1<b><font size="14">\2</font></b>', texto_descricao)
        linha[0] = Paragraph(texto_destacado, estilo_descricao)
        tabela_final.append(linha)

    tabela_pdf = Table(tabela_final, colWidths=larguras_colunas, repeatRows=1) 
    tabela_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('TOPPADDING', (0, 0), (-1, 0), 10),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
    ]))
    
    elementos.append(tabela_pdf)
    doc.build(elementos)

# =====================================================================
# ROTAS E ENDPOINTS
# =====================================================================

@app.get("/")
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# Rota 1: Otimizador de Corte
@app.post("/organizar_corte")
async def organizar_corte(arquivos: List[UploadFile] = File(...)):
    if not arquivos or not arquivos[0].filename:
        return {"erro": "Nenhum arquivo enviado"}
    
    arquivos_processados = []

    for file in arquivos:
        caminho_upload = os.path.join(UPLOAD_DIR, file.filename)
        with open(caminho_upload, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        folhas = analisar_pdf(caminho_upload)
        folhas_ordenadas = sorted(folhas, key=lambda f: f["maior_area"], reverse=True)

        # Regra do novo nome: "NomeOriginal organizado.pdf"
        nome_base, extensao = os.path.splitext(file.filename)
        nome_download = f"{nome_base} organizado{extensao}"
        
        caminho_saida = os.path.join(OUTPUT_DIR, nome_download)
        reorganizar_pdf(caminho_upload, caminho_saida, folhas_ordenadas)
        
        arquivos_processados.append((caminho_saida, nome_download))

    if len(arquivos_processados) == 1:
        return FileResponse(arquivos_processados[0][0], filename=arquivos_processados[0][1], media_type="application/pdf")
    else:
        zip_filename = "Lotes organizados.zip"
        zip_filepath = os.path.join(OUTPUT_DIR, zip_filename)
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for caminho, nome in arquivos_processados:
                zipf.write(caminho, arcname=nome)
        return FileResponse(zip_filepath, filename=zip_filename, media_type="application/zip")


# Rota 2: Organizador de Sobras
@app.post("/organizar_sobras")
async def processar_sobras(arquivos: List[UploadFile] = File(...)):
    if not arquivos or not arquivos[0].filename:
        return {"erro": "Nenhum arquivo enviado"}
    
    arquivos_processados = []

    for arquivo_pdf in arquivos:
        # Regra do novo nome: "NomeOriginal organizado.pdf"
        nome_base, extensao = os.path.splitext(arquivo_pdf.filename)
        nome_download = f"{nome_base} organizado{extensao}"

        caminho_entrada = os.path.join(UPLOAD_DIR, "entrada_sobras_" + arquivo_pdf.filename)
        caminho_saida = os.path.join(OUTPUT_DIR, nome_download)

        with open(caminho_entrada, "wb") as buffer:
            shutil.copyfileobj(arquivo_pdf.file, buffer)

        organizar_pdf_sobras(caminho_entrada, caminho_saida)
        arquivos_processados.append((caminho_saida, nome_download))

    if len(arquivos_processados) == 1:
        return FileResponse(arquivos_processados[0][0], filename=arquivos_processados[0][1], media_type="application/pdf")
    else:
        zip_filename = "Lotes organizados.zip"
        zip_filepath = os.path.join(OUTPUT_DIR, zip_filename)
        with zipfile.ZipFile(zip_filepath, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for caminho, nome in arquivos_processados:
                zipf.write(caminho, arcname=nome)
        return FileResponse(zip_filepath, filename=zip_filename, media_type="application/zip")