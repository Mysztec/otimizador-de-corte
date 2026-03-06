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
from reportlab.pdfgen import canvas

# Importações dos seus serviços originais do FastAPI
from app.services.pdf_analyzer import analisar_pdf
from app.services.pdf_reorder import reorganizar_pdf

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")

# Mude a variável OUTPUT_DIR abaixo se quiser que o Python salve automaticamente
# numa pasta específica do seu computador.
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
# CLASSE PARA NUMERAÇÃO DE PÁGINAS (X de Y)
# =====================================================================
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

        # Extração Cliente
        match_cliente = re.search(r'Cliente:\s*(.*?)(?=Projeto:)', texto_limpo, re.IGNORECASE)
        if match_cliente:
            cliente = match_cliente.group(1).strip()

        # Extração Data
        match_data = re.search(r'Data/Hora:\s*(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})', texto_limpo, re.IGNORECASE)
        str_data_completa = ""
        if match_data:
            data_hora = match_data.group(1).strip()
            str_data_completa = match_data.group(0)

        # Extração Projeto
        match_projeto = re.search(r'Projeto:\s*(.*?)(?=Chapas\b|Sobras\b|Descrição\b|$)', texto_limpo, re.IGNORECASE)
        if match_projeto:
            projeto_bruto = match_projeto.group(1)
            
            if str_data_completa:
                projeto_bruto = projeto_bruto.replace(str_data_completa, '')
            
            projeto_bruto = re.sub(r'Data/Hora:\s*', '', projeto_bruto, flags=re.IGNORECASE)
            projeto = re.sub(r'\s+', ' ', projeto_bruto).strip()

    return cliente, projeto, data_hora

def criar_desenhar_cabecalho(cliente, projeto, data_hora):
    def desenhar_cabecalho(canvas, doc):
        canvas.saveState()
        
        # Título Central
        canvas.setFont('Helvetica-Bold', 16)
        canvas.drawCentredString(doc.pagesize[0] / 2.0, doc.pagesize[1] - 35, "Lista de Compras")
        
        # Informações
        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawString(doc.leftMargin, doc.pagesize[1] - 60, f"Cliente: ")
        canvas.setFont('Helvetica', 10)
        canvas.drawString(doc.leftMargin + 45, doc.pagesize[1] - 60, cliente)

        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin - 120, doc.pagesize[1] - 60, "Data/Hora: ")
        canvas.setFont('Helvetica', 10)
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, doc.pagesize[1] - 60, data_hora)

        canvas.setFont('Helvetica-Bold', 10)
        canvas.drawString(doc.leftMargin, doc.pagesize[1] - 75, f"Projeto: ")
        canvas.setFont('Helvetica', 10)
        canvas.drawString(doc.leftMargin + 45, doc.pagesize[1] - 75, projeto)

        # Linha divisória subiu para a posição 85
        canvas.line(doc.leftMargin, doc.pagesize[1] - 85, doc.pagesize[0] - doc.rightMargin, doc.pagesize[1] - 85)
        
        canvas.restoreState()
    return desenhar_cabecalho

def organizar_pdf_sobras(caminho_entrada, caminho_saida):
    cabecalho_mestre = None
    dados_mestre = []
    dados_chapas = []
    
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
                
                cabecalho_tabela_str = str(tabela_limpa[0]).lower()
                
                # IDENTIFICAR TABELA DE CHAPAS
                if 'chapa' in cabecalho_tabela_str or 'material' in cabecalho_tabela_str or 'matéria' in cabecalho_tabela_str:
                    if not dados_chapas:
                        dados_chapas.append(tabela_limpa[0])
                    dados_chapas.extend(tabela_limpa[1:])
                    continue
                
                # IDENTIFICAR TABELA DE SOBRAS
                if 'descrição' in cabecalho_tabela_str:
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

    # Top margin alterado para 95
    doc = SimpleDocTemplate(caminho_saida, pagesize=landscape(A4), rightMargin=margem_dir, leftMargin=margem_esq, topMargin=95, bottomMargin=40)
    elementos = []
    
    estilo_descricao = ParagraphStyle(name='EstiloDescricao', fontName='Helvetica', fontSize=8, alignment=TA_LEFT)
    estilo_titulo_secao = ParagraphStyle(name='Subtitulo', fontName='Helvetica-Bold', fontSize=12, alignment=TA_LEFT, spaceAfter=10, spaceBefore=15)
    estilo_linha_chapa = ParagraphStyle(name='LinhaChapa', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, spaceAfter=20, spaceBefore=0)
    
    # --- RENDERIZAR LINHA DE CHAPAS (NOVO FORMATO) ---
    if len(dados_chapas) > 1:
        cabecalho_chapas = [str(c).lower().strip() for c in dados_chapas[0]]
        
        idx_mat = 0
        idx_larg = -1
        idx_alt = -1
        idx_qtd = -1
        
        # Procura as colunas corretas dinamicamente
        for i, c in enumerate(cabecalho_chapas):
            if 'chapa' in c or 'material' in c or 'desc' in c:
                idx_mat = i
            elif 'comp' in c:
                idx_larg = i
            elif 'larg' in c:
                if idx_larg == -1: idx_larg = i
                else: idx_alt = i
            elif 'alt' in c:
                idx_alt = i
            elif 'qtd' in c or 'quant' in c or 'chapas' in c:
                idx_qtd = i
                
        # Se não achou pelos nomes, tenta os índices padrão
        if idx_larg == -1 and len(cabecalho_chapas) > 1: idx_larg = 1
        if idx_alt == -1 and len(cabecalho_chapas) > 2: idx_alt = 2
        if idx_qtd == -1: idx_qtd = len(cabecalho_chapas) - 1

        # Formata a frase para cada chapa encontrada
        for linha in dados_chapas[1:]:
            mat = linha[idx_mat] if idx_mat < len(linha) else "Não identificado"
            larg = linha[idx_larg] if (idx_larg != -1 and idx_larg < len(linha)) else ""
            alt = linha[idx_alt] if (idx_alt != -1 and idx_alt < len(linha)) else ""
            qtd = linha[idx_qtd] if (idx_qtd != -1 and idx_qtd < len(linha)) else ""
            
            # Limpa lixo de formatação
            mat = mat.replace('\n', ' ').strip()
            larg = larg.replace('\n', '').strip()
            alt = alt.replace('\n', '').strip()
            qtd = qtd.replace('\n', '').strip()
            
            dimensoes = ""
            if larg and alt:
                dimensoes = f", largura e altura: {larg} x {alt}"
            elif larg:
                dimensoes = f", largura e altura: {larg}"
                
            texto_chapa = f"descrição do material: {mat}{dimensoes} - quantidade: {qtd}"
            elementos.append(Paragraph(texto_chapa, estilo_linha_chapa))

    # --- RENDERIZAR TABELA DE SOBRAS ---
    if dados_organizados:
        
        tabela_final = [cabecalho_mestre] if cabecalho_mestre else []
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
    
    rotina_cabecalho = criar_desenhar_cabecalho(cliente, projeto, data_hora)
    doc.build(elementos, onFirstPage=rotina_cabecalho, onLaterPages=rotina_cabecalho, canvasmaker=NumeracaoCanvas)


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