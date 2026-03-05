from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import shutil
import os

from app.services.pdf_analyzer import analisar_pdf
from app.services.pdf_reorder import reorganizar_pdf

app = FastAPI()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")

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

@app.get("/")
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/organizar")
async def organizar(file: UploadFile = File(...)):

    caminho_upload = os.path.join(UPLOAD_DIR, file.filename)

    with open(caminho_upload, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    folhas = analisar_pdf(caminho_upload)

    # Ordena pela maior área encontrada em cada página
    folhas_ordenadas = sorted(
        folhas,
        key=lambda f: f["maior_area"],
        reverse=True
    )

    caminho_saida = os.path.join(OUTPUT_DIR, "ordenado_" + file.filename)

    reorganizar_pdf(caminho_upload, caminho_saida, folhas_ordenadas)

    return FileResponse(
        caminho_saida,
        filename="ordenado_" + file.filename,
        media_type="application/pdf"
    )