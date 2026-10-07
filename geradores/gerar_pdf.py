from pathlib import Path
import tempfile
from geradores.gerar_docx import gerar_docx
from geradores.office_helper import atualizar_documento, converter_pdf

def gerar_pdf(dados, caminho_pdf, caminho_docx_existente=None):
    caminho_pdf=Path(caminho_pdf); caminho_pdf.parent.mkdir(parents=True,exist_ok=True)
    if caminho_docx_existente:
        caminho_docx=Path(caminho_docx_existente)
        if not caminho_docx.exists(): gerar_docx(dados,caminho_docx,atualizar_word=False)
        atualizar_documento(caminho_docx)
        pdf,_=converter_pdf(caminho_docx,caminho_pdf)
        return pdf
    with tempfile.TemporaryDirectory(prefix="bi_coger_") as td:
        caminho_docx=Path(td)/"boletim_temporario.docx"
        gerar_docx(dados,caminho_docx,atualizar_word=False)
        atualizar_documento(caminho_docx)
        pdf,_=converter_pdf(caminho_docx,caminho_pdf)
        return pdf
