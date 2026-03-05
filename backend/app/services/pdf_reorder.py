from pypdf import PdfReader, PdfWriter

def reorganizar_pdf(caminho_entrada, caminho_saida, folhas_ordenadas):

    reader = PdfReader(caminho_entrada)
    writer = PdfWriter()

    for folha in folhas_ordenadas:
        writer.add_page(reader.pages[folha["pagina"]])

    with open(caminho_saida, "wb") as f:
        writer.write(f)