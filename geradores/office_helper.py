from pathlib import Path
import shutil
import subprocess
import tempfile

WORD_PROGID = "Word.Application"
LIBREOFFICE_CANDIDATOS = [
    r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
]

def detectar_word():
    try:
        import win32com.client
        app = win32com.client.DispatchEx(WORD_PROGID)
        try:
            app.Visible = False
            app.DisplayAlerts = 0
        finally:
            app.Quit()
        return True, None
    except Exception as e:
        return False, str(e)

def localizar_libreoffice():
    p = shutil.which("soffice") or shutil.which("soffice.exe")
    if p:
        return p
    for c in LIBREOFFICE_CANDIDATOS:
        if Path(c).exists():
            return c
    return None

def detectar_suite():
    ok, erro = detectar_word()
    if ok:
        return {"tipo":"word","nome":"Microsoft Word","caminho":None,"erro_word":None}
    lo = localizar_libreoffice()
    if lo:
        return {"tipo":"libreoffice","nome":"LibreOffice","caminho":lo,"erro_word":erro}
    return {"tipo":"nenhum","nome":None,"caminho":None,"erro_word":erro}

def _run_lo(soffice, args, timeout=120):
    cmd=[str(soffice),"--headless","--nologo","--nofirststartwizard"]+list(args)
    p=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=timeout,creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0))
    if p.returncode!=0:
        raise RuntimeError("LibreOffice retornou erro.\n\n"+(p.stderr.strip() or p.stdout.strip()))
    return p


def _forcar_cor_preta_word(doc):
    """
    O Word costuma aplicar azul em hyperlinks/TOC. Força o documento inteiro para preto.
    """
    try:
        # wdColorBlack = 0
        doc.Content.Font.Color = 0
    except Exception:
        pass

    try:
        for i in range(1, doc.Hyperlinks.Count + 1):
            rng = doc.Hyperlinks(i).Range
            rng.Font.Color = 0
            rng.Font.Underline = 0
    except Exception:
        pass

    try:
        for i in range(1, doc.TablesOfContents.Count + 1):
            rng = doc.TablesOfContents(i).Range
            rng.Font.Color = 0
    except Exception:
        pass

def atualizar_com_word(caminho_docx):
    import win32com.client
    caminho_docx=Path(caminho_docx).resolve()
    app=win32com.client.DispatchEx(WORD_PROGID); app.Visible=False; app.DisplayAlerts=0
    doc=None
    try:
        doc=app.Documents.Open(str(caminho_docx))
        try: doc.Fields.Update()
        except Exception: pass
        try:
            for i in range(1,doc.TablesOfContents.Count+1): doc.TablesOfContents(i).Update()
        except Exception: pass
        try: doc.Repaginate()
        except Exception: pass
        doc.Save()
    finally:
        if doc is not None:
            try: doc.Close(SaveChanges=True)
            except Exception: pass
        try: app.Quit()
        except Exception: pass
    return caminho_docx

def atualizar_com_libreoffice(caminho_docx, soffice=None):
    caminho_docx=Path(caminho_docx).resolve(); soffice=soffice or localizar_libreoffice()
    if not soffice: raise FileNotFoundError("LibreOffice não foi encontrado.")
    # Abrir e salvar via conversão DOCX em diretório temporário força repaginação do Writer.
    with tempfile.TemporaryDirectory(prefix="bi_coger_lo_") as td:
        td=Path(td)
        _run_lo(soffice,["--convert-to","docx","--outdir",str(td),str(caminho_docx)])
        cand=td/caminho_docx.name
        if not cand.exists():
            xs=list(td.glob('*.docx'))
            if not xs: raise RuntimeError("LibreOffice não gerou o DOCX atualizado.")
            cand=xs[0]
        shutil.copy2(cand,caminho_docx)
    return caminho_docx

def atualizar_documento(caminho_docx):
    s=detectar_suite()
    if s['tipo']=='word':
        atualizar_com_word(caminho_docx); return 'word'
    if s['tipo']=='libreoffice':
        atualizar_com_libreoffice(caminho_docx,s['caminho']); return 'libreoffice'
    return 'nenhum'

def converter_pdf_com_word(caminho_docx,caminho_pdf):
    import win32com.client
    caminho_docx=Path(caminho_docx).resolve(); caminho_pdf=Path(caminho_pdf).resolve()
    app=win32com.client.DispatchEx(WORD_PROGID); app.Visible=False; app.DisplayAlerts=0
    doc=None
    try:
        doc=app.Documents.Open(str(caminho_docx))
        try: doc.Fields.Update()
        except Exception: pass
        try:
            for i in range(1,doc.TablesOfContents.Count+1): doc.TablesOfContents(i).Update()
        except Exception: pass
        try: doc.Repaginate()
        except Exception: pass
        doc.ExportAsFixedFormat(str(caminho_pdf),17)
    finally:
        if doc is not None:
            try: doc.Close(SaveChanges=True)
            except Exception: pass
        try: app.Quit()
        except Exception: pass
    if not caminho_pdf.exists(): raise RuntimeError("O Word não gerou o PDF.")
    return caminho_pdf

def converter_pdf_com_libreoffice(caminho_docx,caminho_pdf,soffice=None):
    caminho_docx=Path(caminho_docx).resolve(); caminho_pdf=Path(caminho_pdf).resolve(); soffice=soffice or localizar_libreoffice()
    if not soffice: raise FileNotFoundError("LibreOffice não foi encontrado.")
    caminho_pdf.parent.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="bi_coger_pdf_") as td:
        td=Path(td)
        _run_lo(soffice,["--convert-to","pdf","--outdir",str(td),str(caminho_docx)])
        cand=td/(caminho_docx.stem+'.pdf')
        if not cand.exists():
            xs=list(td.glob('*.pdf'))
            if not xs: raise RuntimeError("LibreOffice não gerou o PDF.")
            cand=xs[0]
        shutil.copy2(cand,caminho_pdf)
    return caminho_pdf

def converter_pdf(caminho_docx,caminho_pdf):
    s=detectar_suite()
    if s['tipo']=='word': return converter_pdf_com_word(caminho_docx,caminho_pdf),'word'
    if s['tipo']=='libreoffice': return converter_pdf_com_libreoffice(caminho_docx,caminho_pdf,s['caminho']),'libreoffice'
    raise RuntimeError("Não foi encontrado Microsoft Word nem LibreOffice neste computador.\n\nO DOCX pode ser gerado normalmente, mas para criar PDF é necessário ter Microsoft Word ou LibreOffice instalado.")
