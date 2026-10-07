from pathlib import Path


def atualizar_documento_word(caminho_docx):
    """Atualiza campos/índice e repagina o DOCX usando o Microsoft Word."""
    caminho_docx = Path(caminho_docx).resolve()
    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "O módulo pywin32 não está instalado. Execute instalar.bat novamente."
        ) from e

    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(str(caminho_docx))
        doc.Repaginate()

        # Atualiza primeiro o índice (TOC) e depois os demais campos.
        try:
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
        except Exception:
            pass

        try:
            doc.Fields.Update()
        except Exception:
            pass

        # Atualiza campos também em cabeçalhos/rodapés e outras stories.
        try:
            for story_type in range(1, 18):
                try:
                    rng = doc.StoryRanges(story_type)
                    while rng is not None:
                        try:
                            rng.Fields.Update()
                        except Exception:
                            pass
                        try:
                            rng = rng.NextStoryRange
                        except Exception:
                            rng = None
                except Exception:
                    continue
        except Exception:
            pass

        doc.Repaginate()
        doc.Save()
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def exportar_pdf_word(caminho_docx, caminho_pdf):
    """Atualiza o documento e exporta PDF pelo próprio Microsoft Word."""
    caminho_docx = Path(caminho_docx).resolve()
    caminho_pdf = Path(caminho_pdf).resolve()

    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError(
            "O módulo pywin32 não está instalado. Execute instalar.bat novamente."
        ) from e

    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(str(caminho_docx))
        doc.Repaginate()

        try:
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
        except Exception:
            pass
        try:
            doc.Fields.Update()
        except Exception:
            pass

        doc.Repaginate()
        doc.Save()

        # 17 = wdExportFormatPDF
        doc.ExportAsFixedFormat(str(caminho_pdf), 17)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
