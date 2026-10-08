from pathlib import Path
from geradores.office_helper import atualizar_documento
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


BASE_DIR = Path(__file__).resolve().parent.parent
LOGO_PATH = BASE_DIR / "assets" / "brasao_coger.png"

MESES = {
    1: "jan.", 2: "fev.", 3: "mar.", 4: "abr.", 5: "maio", 6: "jun.",
    7: "jul.", 8: "ago.", 9: "set.", 10: "out.", 11: "nov.", 12: "dez.",
}


def _parse_data(txt):
    try:
        d, m, a = [int(x) for x in txt.split("/")]
        return d, m, a
    except Exception:
        return None


def _fmt_data_inicio(txt):
    """
    Padrão do BI:
    01 out.
    """
    data = _parse_data(txt)
    if not data:
        return txt
    d, m, a = data
    return f"{d:02d} {MESES[m]}"


def _fmt_data_final(txt):
    """
    Padrão do BI:
    07 out. 26
    """
    data = _parse_data(txt)
    if not data:
        return txt
    d, m, a = data
    return f"{d:02d} {MESES[m]} {str(a)[-2:]}"


def _fmt_periodo(data_inicio, data_fim):
    """
    Exemplo final:
    01 out. à 07 out. 26
    """
    return f"{_fmt_data_inicio(data_inicio)} à {_fmt_data_final(data_fim)}"


def _fonte(run, tamanho=10, negrito=False, italico=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.italic = italico
    # Documento oficial: sempre preto. Destaque somente por negrito/sublinhado.
    run.font.color.rgb = RGBColor(0, 0, 0)

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:eastAsia"), "Times New Roman")


def _set_cell_text(cell, text, bold=False, size=9, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    if align is not None:
        p.alignment = align
    r = p.add_run(str(text or ""))
    _fonte(r, size, bold)


def _set_cell_margins(cell, top=70, start=70, bottom=70, end=70):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in [("top", top), ("start", start), ("bottom", bottom), ("end", end)]:
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def _page_border(section):
    sectPr = section._sectPr
    pgBorders = sectPr.find(qn("w:pgBorders"))
    if pgBorders is None:
        pgBorders = OxmlElement("w:pgBorders")
        pgBorders.set(qn("w:offsetFrom"), "page")
        sectPr.append(pgBorders)

    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "double")
        el.set(qn("w:sz"), "12")
        el.set(qn("w:space"), "10")
        el.set(qn("w:color"), "000000")
        pgBorders.append(el)


def _page_number_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    _fonte(run, 8)


def _set_table_borders(table, bottom=False):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        el = borders.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            borders.append(el)
        if edge == "bottom" and bottom:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "nil")


def _configurar_estilos(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10)

    # Parte do boletim: aparece no nível principal do índice.
    h1 = doc.styles["Heading 1"]
    h1.font.name = "Times New Roman"
    h1.font.size = Pt(10.5)
    h1.font.bold = True
    h1.font.underline = True
    h1.paragraph_format.space_before = Pt(5)
    h1.paragraph_format.space_after = Pt(5)
    h1.paragraph_format.keep_with_next = True

    # Título numerado: 1., 2., 3...
    h2 = doc.styles["Heading 2"]
    h2.font.name = "Times New Roman"
    h2.font.size = Pt(10)
    h2.font.bold = True
    h2.paragraph_format.left_indent = Cm(0)
    h2.paragraph_format.space_before = Pt(4)
    h2.paragraph_format.space_after = Pt(2)
    h2.paragraph_format.keep_with_next = True

    # Subtítulo: a., b., c...
    h3 = doc.styles["Heading 3"]
    h3.font.name = "Times New Roman"
    h3.font.size = Pt(10)
    h3.font.bold = True
    h3.paragraph_format.left_indent = Cm(0)
    h3.paragraph_format.space_before = Pt(3)
    h3.paragraph_format.space_after = Pt(2)
    h3.paragraph_format.keep_with_next = True



def _configurar_estilos_pretos(doc):
    """
    Remove cores automáticas do Word. Títulos, subtítulos e índice ficam pretos,
    usando apenas negrito/sublinhado para destaque.
    """
    estilos_alvo = [
        "Normal",
        "Title",
        "Subtitle",
        "Heading 1",
        "Heading 2",
        "Heading 3",
        "TOC 1",
        "TOC 2",
        "TOC 3",
        "Hyperlink",
        "FollowedHyperlink",
    ]

    for nome in estilos_alvo:
        try:
            estilo = doc.styles[nome]
            estilo.font.name = "Times New Roman"
            estilo.font.color.rgb = RGBColor(0, 0, 0)
        except Exception:
            pass

def _configurar_documento(doc, dados):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.header_distance = Cm(0.7)
    section.footer_distance = Cm(0.7)
    section.different_first_page_header_footer = True

    _page_border(section)
    _configurar_estilos(doc)

    periodo = _fmt_periodo(dados["data_inicio"], dados["data_fim"])

    # Cabeçalho das folhas seguintes: texto à esquerda e 'fl. X' no extremo direito.
    header = section.header
    p0 = header.paragraphs[0]
    p0._element.getparent().remove(p0._element)

    table = header.add_table(rows=1, cols=2, width=Cm(18.0))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(16.1)
    table.columns[1].width = Cm(1.9)
    _set_table_borders(table, bottom=True)

    left = table.cell(0, 0)
    right = table.cell(0, 1)
    left.width = Cm(16.1)
    right.width = Cm(1.9)

    _set_cell_text(
        left,
        f"Boletim- Interno nº {dados['numero']} de {periodo} - COGER",
        size=8,
        align=WD_ALIGN_PARAGRAPH.LEFT,
    )

    right.text = ""
    p = right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("fl. ")
    _fonte(r, 8)
    _page_number_field(p)

    first = section.first_page_header
    first.paragraphs[0].text = ""


def _paragrafo_central(doc, texto="", size=10, bold=False, underline=False, before=0, after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(texto)
    _fonte(r, size, bold)
    r.underline = underline
    return p


def _paragrafo_heading(doc, texto, nivel, central=False):
    style_name = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}[nivel]
    p = doc.add_paragraph(style=style_name)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if central else WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(texto)
    _fonte(r, 10.5 if nivel == 1 else 10, True)
    if nivel == 1:
        r.underline = True
    return p


def _paragrafo_texto(doc, texto, justify=True, first_indent=False, after=3):
    p = doc.add_paragraph()
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1
    if first_indent:
        p.paragraph_format.first_line_indent = Cm(0.7)
    r = p.add_run(texto)
    _fonte(r, 10)
    return p



def _add_foto(doc, caminho, largura_max_cm=15.5):
    """
    Insere uma foto centralizada no BI, mantendo a proporção.
    A largura máxima evita ultrapassar as margens do documento.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("[Foto não encontrada]")
        _fonte(r, 9, italico=True)
        return

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)

    run = p.add_run()
    run.add_picture(str(caminho), width=Cm(largura_max_cm))


def _add_tabela(doc, dados):
    if not dados:
        return
    cols = max(len(r) for r in dados)
    table = doc.add_table(rows=len(dados), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = True

    for i, row in enumerate(dados):
        row = list(row) + [""] * (cols - len(row))
        for j, valor in enumerate(row):
            cell = table.cell(i, j)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_margins(cell)
            _set_cell_text(cell, valor, bold=(i == 0), size=8.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def _capa(doc, dados):
    periodo = _fmt_periodo(dados["data_inicio"], dados["data_fim"])

    _paragrafo_central(doc, "POLÍCIA MILITAR DO PARANÁ", 16, True, after=0)
    _paragrafo_central(doc, "CORREGEDORIA-GERAL", 12, True, after=5)

    if LOGO_PATH.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(7)
        r = p.add_run()
        r.add_picture(str(LOGO_PATH), width=Cm(3.0))

    _paragrafo_central(doc, "BOLETIM-INTERNO", 14, True, before=3, after=8)
    _paragrafo_central(doc, f"Nº {dados['numero']}", 13, True, after=14)
    _paragrafo_central(doc, f"{dados['local']}, de {periodo}.", 10, True, after=18)
    _paragrafo_central(
        doc,
        "Para conhecimento e devida execução, torno público o seguinte:",
        10,
        True,
        after=12,
    )


def _adicionar_toc(doc):
    """Insere um índice Word real: pontilhado + páginas atualizadas pelo Word."""
    _paragrafo_central(doc, "ÍNDICE", 12, True, after=8)

    p = doc.add_paragraph()
    run = p.add_run()

    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '

    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")

    placeholder = OxmlElement("w:t")
    placeholder.text = "Índice será atualizado automaticamente."

    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")

    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(placeholder)
    run._r.append(fld_end)



def _forcar_todo_texto_preto(doc):
    """
    Garante preto em todos os textos existentes antes de salvar.
    """
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.color.rgb = RGBColor(0, 0, 0)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.color.rgb = RGBColor(0, 0, 0)

    for section in doc.sections:
        for area in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
        ):
            for p in area.paragraphs:
                for run in p.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)

def gerar_docx(dados, caminho, atualizar_word=True):
    doc = Document()
    _configurar_documento(doc, dados)
    _capa(doc, dados)

    for parte in dados["partes"]:
        _paragrafo_heading(doc, parte["titulo"], 1, central=True)

        itens = parte.get("itens", [])
        if not itens:
            # Entra no índice, como no BI de referência.
            _paragrafo_heading(doc, "SEM ALTERAÇÃO", 2, central=True)
            continue

        numero_titulo = 0
        letra_subtitulo = 0

        for item in itens:
            nivel = item.get("nivel", "subtitulo")
            titulo = item.get("titulo", "").strip()

            if nivel == "titulo":
                numero_titulo += 1
                letra_subtitulo = 0
                if titulo:
                    _paragrafo_heading(doc, f"{numero_titulo}. {titulo}", 2)
            else:
                letra_subtitulo += 1
                if titulo:
                    letra = chr(96 + letra_subtitulo)
                    _paragrafo_heading(doc, f"{letra}. {titulo}", 3)

            if item.get("tipo") == "tabela" and item.get("tabela"):
                _add_tabela(doc, item["tabela"])
            elif item.get("texto"):
                for linha in str(item["texto"]).splitlines():
                    linha = linha.rstrip()
                    if linha:
                        _paragrafo_texto(doc, linha, justify=True, first_indent=False)
                    else:
                        doc.add_paragraph()

            for foto in item.get("fotos") or []:
                _add_foto(doc, foto)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    r = p.add_run("-" * 112)
    _fonte(r, 8)

    if dados.get("assina_nome") or dados.get("assina_funcao"):
        p = doc.add_paragraph()
        r = p.add_run("ASSINA:")
        _fonte(r, 10, True)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if dados.get("assina_nome"):
            r = p.add_run(dados["assina_nome"] + "\n")
            _fonte(r, 10)
        if dados.get("assina_funcao"):
            r = p.add_run(dados["assina_funcao"])
            _fonte(r, 10, True)

    if dados.get("confere_nome") or dados.get("confere_funcao"):
        p = doc.add_paragraph()
        r = p.add_run("CONFERE:")
        _fonte(r, 10, True)

        p = doc.add_paragraph()
        if dados.get("confere_nome"):
            r = p.add_run(dados["confere_nome"] + "\n")
            _fonte(r, 10)
        if dados.get("confere_funcao"):
            r = p.add_run(dados["confere_funcao"])
            _fonte(r, 10, True)

    doc.add_page_break()
    _adicionar_toc(doc)

    caminho = Path(caminho)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    _forcar_todo_texto_preto(doc)
    doc.save(str(caminho))
    if atualizar_word:
        atualizar_documento(caminho)


    return caminho
