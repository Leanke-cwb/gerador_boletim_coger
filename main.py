
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import sqlite3
import json
from datetime import datetime
from datetime import datetime
import os
import sys

import customtkinter as ctk
from PIL import Image
from tkcalendar import Calendar

from geradores.gerar_docx import gerar_docx
from geradores.gerar_pdf import gerar_pdf
from geradores.office_helper import detectar_suite
from ocr.processar_imagem import processar_imagem, localizar_tesseract

# ============================================================
# CAMINHOS INTERNOS DO PROGRAMA
# ============================================================

if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
    # Quando estiver rodando como executável criado pelo PyInstaller
    RESOURCE_DIR = Path(sys._MEIPASS)
else:
    # Quando estiver rodando diretamente pelo Python
    RESOURCE_DIR = Path(__file__).resolve().parent

ASSETS_DIR = RESOURCE_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "brasao_coger.png"


# ============================================================
# CAMINHOS GRAVÁVEIS DO USUÁRIO
# ============================================================

LOCAL_APP_DATA = Path(
    os.environ.get(
        "LOCALAPPDATA",
        str(Path.home() / "AppData" / "Local")
    )
)

USER_DATA_DIR = LOCAL_APP_DATA / "Gerador Boletim COGER"
DATABASE_DIR = USER_DATA_DIR / "database"

DOCUMENTOS_DIR = Path.home() / "Documents"
SAIDA_DIR = DOCUMENTOS_DIR / "Gerador Boletim COGER" / "Boletins"

# Cria somente pastas em locais onde o usuário tem permissão de gravação.
DATABASE_DIR.mkdir(parents=True, exist_ok=True)
SAIDA_DIR.mkdir(parents=True, exist_ok=True)

DB_PATH = DATABASE_DIR / "boletim.db"

PARTES = [
    "1ª PARTE - SERVIÇOS DIÁRIOS",
    "2ª PARTE - INSTRUÇÃO",
    "3ª PARTE - ASSUNTOS GERAIS E ADMINISTRATIVOS",
    "4ª PARTE - JUSTIÇA E DISCIPLINA",
]

PARTES_CURTAS = ["1ª Parte", "2ª Parte", "3ª Parte", "4ª Parte"]

NIVEIS = [
    "Título numerado (1, 2, 3...)",
    "Subtítulo por letra (a, b, c...)",
]

AZUL = "#183A5A"
AZUL_HOVER = "#244D73"
AZUL_CLARO = "#E9F0F6"
VERDE = "#2E6E51"
VERDE_HOVER = "#245A42"
VERMELHO = "#A04444"
VERMELHO_HOVER = "#873636"
CINZA = "#6C7782"


def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS configuracao (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            assina_nome TEXT,
            assina_funcao TEXT,
            confere_nome TEXT,
            confere_funcao TEXT
        )
    """)
    cur.execute("""
        INSERT OR IGNORE INTO configuracao
        (id, assina_nome, assina_funcao, confere_nome, confere_funcao)
        VALUES (1, '', '', '', '')
    """)

    # Rascunho único do boletim em edição.
    # Fica salvo no perfil do usuário, fora da pasta do programa.
    cur.execute("""
        CREATE TABLE IF NOT EXISTS rascunho_boletim (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            dados_json TEXT NOT NULL,
            atualizado_em TEXT NOT NULL
        )
    """)

    # Histórico permanente dos boletins.
    cur.execute("""
        CREATE TABLE IF NOT EXISTS boletins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero TEXT NOT NULL,
            local TEXT,
            data_inicio TEXT,
            data_fim TEXT,
            dados_json TEXT NOT NULL,
            criado_em TEXT NOT NULL,
            atualizado_em TEXT NOT NULL
        )
    """)
    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_boletins_numero
        ON boletins(numero)
    """)
    cur.execute("""
        CREATE INDEX IF NOT EXISTS idx_boletins_periodo
        ON boletins(data_inicio, data_fim)
    """)

    con.commit()
    con.close()


def carregar_config():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        SELECT assina_nome, assina_funcao, confere_nome, confere_funcao
        FROM configuracao WHERE id=1
    """)
    row = cur.fetchone()
    con.close()
    return row or ("", "", "", "")


def salvar_config(dados):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        UPDATE configuracao
        SET assina_nome=?, assina_funcao=?, confere_nome=?, confere_funcao=?
        WHERE id=1
    """, dados)
    con.commit()
    con.close()



def salvar_rascunho_db(dados):
    payload = json.dumps(dados, ensure_ascii=False)
    atualizado_em = datetime.now().isoformat(timespec="seconds")

    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        INSERT INTO rascunho_boletim (id, dados_json, atualizado_em)
        VALUES (1, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            dados_json=excluded.dados_json,
            atualizado_em=excluded.atualizado_em
    """, (payload, atualizado_em))
    con.commit()
    con.close()
    return atualizado_em


def carregar_rascunho_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        SELECT dados_json, atualizado_em
        FROM rascunho_boletim
        WHERE id=1
    """)
    row = cur.fetchone()
    con.close()

    if not row:
        return None, None

    try:
        return json.loads(row[0]), row[1]
    except Exception:
        return None, row[1]


def excluir_rascunho_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("DELETE FROM rascunho_boletim WHERE id=1")
    con.commit()
    con.close()



def salvar_boletim_db(dados, boletim_id=None):
    """
    Insere um novo boletim ou atualiza um boletim já aberto.
    Retorna (id, atualizado_em).
    """
    agora = datetime.now().isoformat(timespec="seconds")
    payload = json.dumps(dados, ensure_ascii=False)

    numero = (dados.get("numero") or "").strip()
    local = (dados.get("local") or "").strip()
    data_inicio = (dados.get("data_inicio") or "").strip()
    data_fim = (dados.get("data_fim") or "").strip()

    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()

    if boletim_id:
        cur.execute("""
            UPDATE boletins
            SET numero=?, local=?, data_inicio=?, data_fim=?,
                dados_json=?, atualizado_em=?
            WHERE id=?
        """, (
            numero, local, data_inicio, data_fim,
            payload, agora, boletim_id
        ))
        if cur.rowcount == 0:
            boletim_id = None

    if not boletim_id:
        cur.execute("""
            INSERT INTO boletins
            (numero, local, data_inicio, data_fim, dados_json, criado_em, atualizado_em)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            numero, local, data_inicio, data_fim,
            payload, agora, agora
        ))
        boletim_id = cur.lastrowid

    con.commit()
    con.close()
    return boletim_id, agora


def listar_boletins_db(filtro=""):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()

    filtro = (filtro or "").strip()
    if filtro:
        termo = f"%{filtro}%"
        cur.execute("""
            SELECT id, numero, local, data_inicio, data_fim, criado_em, atualizado_em
            FROM boletins
            WHERE numero LIKE ?
               OR local LIKE ?
               OR data_inicio LIKE ?
               OR data_fim LIKE ?
            ORDER BY atualizado_em DESC, id DESC
        """, (termo, termo, termo, termo))
    else:
        cur.execute("""
            SELECT id, numero, local, data_inicio, data_fim, criado_em, atualizado_em
            FROM boletins
            ORDER BY atualizado_em DESC, id DESC
        """)

    rows = cur.fetchall()
    con.close()
    return rows


def carregar_boletim_db(boletim_id):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        SELECT dados_json
        FROM boletins
        WHERE id=?
    """, (boletim_id,))
    row = cur.fetchone()
    con.close()

    if not row:
        return None

    try:
        return json.loads(row[0])
    except Exception:
        return None


def excluir_boletim_db(boletim_id):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("DELETE FROM boletins WHERE id=?", (boletim_id,))
    con.commit()
    con.close()


def duplicar_boletim_db(boletim_id):
    dados = carregar_boletim_db(boletim_id)
    if not dados:
        return None

    # A cópia recebe um novo registro e pode ser alterada independentemente.
    dados = json.loads(json.dumps(dados, ensure_ascii=False))
    numero = (dados.get("numero") or "").strip()
    dados["numero"] = numero

    novo_id, _ = salvar_boletim_db(dados, boletim_id=None)
    return novo_id


def dados_documento_de_editor(dados_editor):
    """
    Converte o formato usado no editor/rascunho para o formato usado
    pelos geradores DOCX/PDF.
    """
    assina_nome, assina_funcao, confere_nome, confere_funcao = carregar_config()
    partes_editor = dados_editor.get("partes") or []

    partes = []
    for i in range(4):
        bloco = partes_editor[i] if i < len(partes_editor) else {}
        itens = bloco.get("itens") or []
        partes.append({
            "titulo": PARTES[i],
            "itens": itens,
        })

    return {
        "numero": (dados_editor.get("numero") or "").strip(),
        "local": (dados_editor.get("local") or "").strip() or "Curitiba",
        "data_inicio": (dados_editor.get("data_inicio") or "").strip(),
        "data_fim": (dados_editor.get("data_fim") or "").strip(),
        "partes": partes,
        "assina_nome": assina_nome,
        "assina_funcao": assina_funcao,
        "confere_nome": confere_nome,
        "confere_funcao": confere_funcao,
    }


def rotular_itens(itens):
    numero_titulo = 0
    letra_subtitulo = 0
    resultado = []
    for item in itens:
        if item.get("nivel", "subtitulo") == "titulo":
            numero_titulo += 1
            letra_subtitulo = 0
            rotulo = f"{numero_titulo}."
        else:
            letra_subtitulo += 1
            rotulo = f"{chr(96 + letra_subtitulo)}."
        resultado.append((rotulo, item))
    return resultado



class CalendarDateField(ctk.CTkFrame):
    """
    Campo de data com calendário visual.
    O valor retornado permanece em DD/MM/AAAA para manter compatibilidade
    com a formatação do boletim.
    """
    def __init__(self, master, placeholder="Selecione a data", on_change=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.on_change = on_change
        self.grid_columnconfigure(0, weight=1)

        self.entry = ctk.CTkEntry(
            self,
            height=36,
            corner_radius=9,
            placeholder_text=placeholder,
        )
        self.entry.grid(row=0, column=0, sticky="ew")
        self.entry.bind("<KeyRelease>", self._changed)

        self.btn = ctk.CTkButton(
            self,
            text="📅",
            width=42,
            height=36,
            corner_radius=9,
            fg_color=AZUL,
            hover_color=AZUL_HOVER,
            command=self.abrir_calendario,
        )
        self.btn.grid(row=0, column=1, padx=(6, 0))

    def _changed(self, _=None):
        if self.on_change:
            self.on_change()

    def get(self):
        return self.entry.get()

    def delete(self, first, last=None):
        return self.entry.delete(first, last)

    def insert(self, index, text):
        return self.entry.insert(index, text)

    def focus_set(self):
        return self.entry.focus_set()

    def bind(self, sequence=None, command=None, add=None):
        # Mantém compatibilidade com chamadas existentes feitas no campo.
        if sequence:
            return self.entry.bind(sequence, command, add)
        return super().bind(sequence, command, add)

    def abrir_calendario(self):
        win = ctk.CTkToplevel(self)
        win.title("Selecionar data")
        win.resizable(False, False)
        win.transient(self.winfo_toplevel())
        win.grab_set()

        hoje = datetime.now()
        dia, mes, ano = hoje.day, hoje.month, hoje.year

        atual = self.get().strip()
        try:
            d, m, a = [int(x) for x in atual.split("/")]
            dia, mes, ano = d, m, a
        except Exception:
            pass

        container = ctk.CTkFrame(win, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=16, pady=16)

        cal = Calendar(
            container,
            selectmode="day",
            year=ano,
            month=mes,
            day=dia,
            date_pattern="dd/mm/yyyy",
            locale="pt_BR",
            showweeknumbers=False,
        )
        cal.pack(fill="both", expand=True)

        botoes = ctk.CTkFrame(container, fg_color="transparent")
        botoes.pack(fill="x", pady=(12, 0))

        def selecionar():
            valor = cal.get_date()
            self.entry.delete(0, "end")
            self.entry.insert(0, valor)
            if self.on_change:
                self.on_change()
            win.destroy()

        ctk.CTkButton(
            botoes,
            text="Cancelar",
            width=100,
            fg_color=("gray87", "gray26"),
            hover_color=("gray79", "gray32"),
            text_color=("gray20", "gray90"),
            command=win.destroy,
        ).pack(side="right")

        ctk.CTkButton(
            botoes,
            text="Selecionar",
            width=110,
            fg_color=AZUL,
            hover_color=AZUL_HOVER,
            command=selecionar,
        ).pack(side="right", padx=(0, 8))


class ModernDialog(ctk.CTkToplevel):
    def __init__(self, master, title, size="760x420"):
        super().__init__(master)
        self.title(title)
        self.geometry(size)
        self.minsize(640, 360)
        self.transient(master)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)


class ParteView(ctk.CTkFrame):
    def __init__(self, master, titulo, on_change=None):
        super().__init__(master, fg_color="transparent")
        self.titulo = titulo
        self.itens = []
        self.on_change = on_change
        self.editando_indice = None

        self.grid_columnconfigure((0, 1), weight=1, uniform="cols")
        self.grid_rowconfigure(1, weight=1)

        # Cabeçalho da seção
        cab = ctk.CTkFrame(self, fg_color="transparent")
        cab.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        cab.grid_columnconfigure(0, weight=1)

        self.lbl_titulo_secao = ctk.CTkLabel(
            cab,
            text=titulo,
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=AZUL,
        )
        self.lbl_titulo_secao.grid(row=0, column=0, sticky="w")

        self.lbl_qtd = ctk.CTkLabel(
            cab,
            text="0 publicações",
            fg_color=AZUL_CLARO,
            text_color=AZUL,
            corner_radius=12,
            padx=12,
            pady=5,
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.lbl_qtd.grid(row=0, column=1, sticky="e")

        # Card editor
        editor = ctk.CTkFrame(self, corner_radius=14)
        editor.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        editor.grid_columnconfigure(0, weight=1)
        editor.grid_rowconfigure(8, weight=1)

        ctk.CTkLabel(
            editor,
            text="Nova publicação",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=18, pady=(18, 2))

        ctk.CTkLabel(
            editor,
            text="Escolha o nível, informe o título e escreva o conteúdo.",
            text_color=("gray45", "gray70"),
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, sticky="w", padx=18, pady=(0, 14))

        ctk.CTkLabel(editor, text="Nível do item", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=2, column=0, sticky="w", padx=18, pady=(0, 5)
        )

        self.cmb_nivel = ctk.CTkComboBox(
            editor,
            values=NIVEIS,
            state="readonly",
            height=36,
            corner_radius=9,
            button_color=AZUL,
            button_hover_color=AZUL_HOVER,
        )
        self.cmb_nivel.set(NIVEIS[1])
        self.cmb_nivel.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 12))

        ctk.CTkLabel(editor, text="Título / subtítulo", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=4, column=0, sticky="w", padx=18, pady=(0, 5)
        )

        self.ent_titulo = ctk.CTkEntry(
            editor,
            height=38,
            corner_radius=9,
            placeholder_text="Ex.: ORDEM DE MOVIMENTAÇÃO",
        )
        self.ent_titulo.grid(row=5, column=0, sticky="ew", padx=18, pady=(0, 12))

        linha = ctk.CTkFrame(editor, fg_color="transparent")
        linha.grid(row=6, column=0, sticky="ew", padx=18)
        linha.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            linha, text="Texto / conteúdo", font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        self.lbl_chars = ctk.CTkLabel(
            linha, text="0 caracteres", text_color=("gray50", "gray65"), font=ctk.CTkFont(size=11)
        )
        self.lbl_chars.grid(row=0, column=1, sticky="e")

        self.txt = ctk.CTkTextbox(
            editor,
            corner_radius=10,
            border_width=1,
            border_color=("gray78", "gray35"),
            font=("Segoe UI", 13),
            wrap="word",
        )
        self.txt.grid(row=8, column=0, sticky="nsew", padx=18, pady=(6, 12))
        self.txt.bind("<KeyRelease>", self._atualizar_contador)

        acoes = ctk.CTkFrame(editor, fg_color="transparent")
        acoes.grid(row=9, column=0, sticky="ew", padx=18, pady=(0, 10))
        acoes.grid_columnconfigure(0, weight=1)

        self.btn_adicionar = ctk.CTkButton(
            acoes,
            text="Adicionar publicação",
            height=40,
            corner_radius=10,
            fg_color=AZUL,
            hover_color=AZUL_HOVER,
            command=self.adicionar_ou_salvar,
        )
        self.btn_adicionar.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(
            acoes,
            text="Limpar",
            width=88,
            height=40,
            corner_radius=10,
            fg_color=("gray88", "gray26"),
            hover_color=("gray80", "gray32"),
            text_color=("gray20", "gray90"),
            command=self.limpar_campos,
        ).grid(row=0, column=1, padx=(6, 0))

        ctk.CTkButton(
            editor,
            text="Importar print / OCR",
            height=38,
            corner_radius=10,
            fg_color="transparent",
            border_width=1,
            border_color=AZUL,
            text_color=AZUL,
            hover_color=AZUL_CLARO,
            command=self.importar_print,
        ).grid(row=10, column=0, sticky="ew", padx=18, pady=(0, 18))

        # Card lista
        lista_card = ctk.CTkFrame(self, corner_radius=14)
        lista_card.grid(row=1, column=1, sticky="nsew", padx=(8, 0))
        lista_card.grid_columnconfigure(0, weight=1)
        lista_card.grid_rowconfigure(3, weight=1)

        cab_lista = ctk.CTkFrame(lista_card, fg_color="transparent")
        cab_lista.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 0))
        cab_lista.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            cab_lista,
            text="Publicações desta parte",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            cab_lista,
            text="duplo clique para editar",
            text_color=("gray50", "gray65"),
            font=ctk.CTkFont(size=11),
        ).grid(row=0, column=1, sticky="e")

        ctk.CTkLabel(
            lista_card,
            text="A ordem abaixo será usada no boletim e no índice.",
            text_color=("gray45", "gray70"),
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, sticky="w", padx=18, pady=(3, 10))

        # Treeview dentro do CustomTkinter
        tree_box = ctk.CTkFrame(lista_card, corner_radius=10, fg_color=("white", "#202428"))
        tree_box.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 12))
        tree_box.grid_columnconfigure(0, weight=1)
        tree_box.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_box,
            columns=("ordem", "nivel", "titulo"),
            show="headings",
            selectmode="browse",
            style="Modern.Treeview",
        )
        self.tree.heading("ordem", text="Índice")
        self.tree.heading("nivel", text="Tipo")
        self.tree.heading("titulo", text="Título / subtítulo")
        self.tree.column("ordem", width=68, anchor="center", stretch=False)
        self.tree.column("nivel", width=104, anchor="center", stretch=False)
        self.tree.column("titulo", width=380, anchor="w")

        scroll = ctk.CTkScrollbar(tree_box, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        scroll.grid(row=0, column=1, sticky="ns", padx=(4, 8), pady=8)

        self.tree.bind("<Double-1>", lambda e: self.editar_selecionado())
        self.tree.bind("<Return>", lambda e: self.editar_selecionado())
        self.tree.bind("<Delete>", lambda e: self.remover())

        toolbar = ctk.CTkFrame(lista_card, fg_color="transparent")
        toolbar.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 18))
        toolbar.grid_columnconfigure(3, weight=1)

        for col, (texto, comando) in enumerate([
            ("Editar", self.editar_selecionado),
            ("Subir", lambda: self.mover(-1)),
            ("Descer", lambda: self.mover(1)),
        ]):
            ctk.CTkButton(
                toolbar,
                text=texto,
                width=84,
                height=34,
                corner_radius=9,
                fg_color=("gray88", "gray26"),
                hover_color=("gray80", "gray32"),
                text_color=("gray20", "gray90"),
                command=comando,
            ).grid(row=0, column=col, padx=(0, 6))

        ctk.CTkButton(
            toolbar,
            text="Remover",
            width=92,
            height=34,
            corner_radius=9,
            fg_color=VERMELHO,
            hover_color=VERMELHO_HOVER,
            command=self.remover,
        ).grid(row=0, column=4, sticky="e")

    def _atualizar_contador(self, _=None):
        n = len(self.txt.get("1.0", "end-1c"))
        self.lbl_chars.configure(text=f"{n} caracteres")

    def _nivel_atual(self):
        return "titulo" if self.cmb_nivel.get() == NIVEIS[0] else "subtitulo"

    def _notificar(self):
        qtd = len(self.itens)
        self.lbl_qtd.configure(text=f"{qtd} publicação" if qtd == 1 else f"{qtd} publicações")
        if self.on_change:
            self.on_change()

    def _indice_selecionado(self):
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    def _atualizar_lista(self, selecionar=None):
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        for idx, (rotulo, item) in enumerate(rotular_itens(self.itens)):
            nome = item.get("titulo", "").strip() or "Conteúdo sem título"
            nivel = "Título" if item.get("nivel") == "titulo" else "Subtítulo"
            if item.get("tipo") == "tabela":
                nome += "  [Tabela OCR]"
            self.tree.insert("", "end", iid=str(idx), values=(rotulo, nivel, nome))

        if selecionar is not None and 0 <= selecionar < len(self.itens):
            iid = str(selecionar)
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self.tree.see(iid)

        self._notificar()

    def adicionar_ou_salvar(self):
        titulo = self.ent_titulo.get().strip()
        texto = self.txt.get("1.0", "end").strip()

        if not titulo and not texto:
            messagebox.showwarning("Atenção", "Informe um título/subtítulo, um texto ou ambos.")
            return

        item = {
            "tipo": "texto",
            "nivel": self._nivel_atual(),
            "titulo": titulo,
            "texto": texto,
        }

        if self.editando_indice is None:
            self.itens.append(item)
            idx = len(self.itens) - 1
        else:
            anterior = self.itens[self.editando_indice]
            if anterior.get("tipo") == "tabela":
                item["tipo"] = "tabela"
                item["tabela"] = anterior.get("tabela", [])
                item["imagem_origem"] = anterior.get("imagem_origem")
            self.itens[self.editando_indice] = item
            idx = self.editando_indice

        self._atualizar_lista(idx)
        self.limpar_campos()
        try:
            self.winfo_toplevel().salvar_rascunho(manual=False)
        except Exception:
            pass

    def editar_selecionado(self):
        idx = self._indice_selecionado()
        if idx is None or not (0 <= idx < len(self.itens)):
            return

        item = self.itens[idx]
        self.editando_indice = idx
        self.cmb_nivel.set(NIVEIS[0] if item.get("nivel") == "titulo" else NIVEIS[1])

        self.ent_titulo.delete(0, "end")
        self.ent_titulo.insert(0, item.get("titulo", ""))

        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", item.get("texto", ""))
        self._atualizar_contador()

        self.btn_adicionar.configure(text="Salvar alteração")
        self.ent_titulo.focus_set()

        if item.get("tipo") == "tabela":
            if messagebox.askyesno(
                "Editar tabela",
                "Este item contém uma tabela OCR.\n\nDeseja abrir também o editor da tabela?"
            ):
                self.mostrar_revisao_tabela(item)

    def importar_print(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar print ou imagem",
            filetypes=[
                ("Imagens", "*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if not caminho:
            return

        try:
            resultado = processar_imagem(caminho)
        except FileNotFoundError as e:
            messagebox.showerror(
                "Tesseract não encontrado",
                str(e) + "\n\nInstale o Tesseract OCR e o idioma Português.",
            )
            return
        except Exception as e:
            messagebox.showerror("Erro no OCR", str(e))
            return

        titulo = self.ent_titulo.get().strip() or "Conteúdo importado por OCR"
        nivel = self._nivel_atual()

        if resultado["tipo"] == "tabela":
            tabela = resultado.get("tabela", [])
            if not tabela:
                messagebox.showwarning("OCR", "A imagem parece conter tabela, mas nenhuma célula foi reconhecida.")
                return
            item = {
                "tipo": "tabela",
                "nivel": nivel,
                "titulo": titulo,
                "texto": "",
                "tabela": tabela,
                "imagem_origem": caminho,
            }
            self.itens.append(item)
            self._atualizar_lista(len(self.itens) - 1)
            self.mostrar_revisao_tabela(item)
        else:
            texto = resultado.get("texto", "").strip()
            if not texto:
                messagebox.showwarning("OCR", "Nenhum texto foi reconhecido na imagem.")
                return
            item = {
                "tipo": "texto",
                "nivel": nivel,
                "titulo": titulo,
                "texto": texto,
                "imagem_origem": caminho,
            }
            self.itens.append(item)
            self._atualizar_lista(len(self.itens) - 1)
            self.mostrar_revisao_texto(item)

        self.limpar_campos()

    def mostrar_revisao_texto(self, item):
        win = ModernDialog(self.winfo_toplevel(), "Revisar texto reconhecido", "900x650")

        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.grid(row=0, column=0, sticky="nsew", padx=22, pady=22)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            frame, text="Revisão do OCR", font=ctk.CTkFont(size=20, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            frame,
            text="Confira especialmente nomes, CPF, números e protocolos antes de salvar.",
            text_color=("gray45", "gray70"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 12))

        txt = ctk.CTkTextbox(
            frame,
            corner_radius=12,
            border_width=1,
            border_color=("gray78", "gray35"),
            font=("Segoe UI", 13),
        )
        txt.grid(row=2, column=0, sticky="nsew")
        txt.insert("1.0", item["texto"])

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=3, column=0, sticky="e", pady=(14, 0))

        def salvar():
            item["texto"] = txt.get("1.0", "end").strip()
            self._atualizar_lista()
            win.destroy()

        ctk.CTkButton(
            botoes,
            text="Cancelar",
            width=100,
            fg_color=("gray88", "gray26"),
            hover_color=("gray80", "gray32"),
            text_color=("gray20", "gray90"),
            command=win.destroy,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            botoes, text="Salvar revisão", width=130, fg_color=AZUL, hover_color=AZUL_HOVER, command=salvar
        ).pack(side="left")

    def mostrar_revisao_tabela(self, item):
        win = ModernDialog(self.winfo_toplevel(), "Revisar tabela reconhecida", "1100x700")

        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.grid(row=0, column=0, sticky="nsew", padx=22, pady=22)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            frame, text="Revisão da tabela", font=ctk.CTkFont(size=20, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            frame,
            text="Dê duplo clique em qualquer célula para corrigir o conteúdo.",
            text_color=("gray45", "gray70"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 12))

        dados = item.get("tabela", [])
        max_cols = max((len(r) for r in dados), default=1)
        cols = [f"c{i}" for i in range(max_cols)]

        box = ctk.CTkFrame(frame, corner_radius=12)
        box.grid(row=2, column=0, sticky="nsew")
        box.grid_columnconfigure(0, weight=1)
        box.grid_rowconfigure(0, weight=1)

        tree = ttk.Treeview(box, columns=cols, show="headings", style="Modern.Treeview")
        for i, c in enumerate(cols):
            tree.heading(c, text=f"Coluna {i + 1}")
            tree.column(c, width=160, anchor="w")

        for row in dados:
            values = list(row) + [""] * (max_cols - len(row))
            tree.insert("", "end", values=values)

        ysb = ctk.CTkScrollbar(box, command=tree.yview)
        xsb = ctk.CTkScrollbar(box, orientation="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=(10, 0))
        ysb.grid(row=0, column=1, sticky="ns", padx=(5, 10), pady=(10, 0))
        xsb.grid(row=1, column=0, sticky="ew", padx=(10, 0), pady=(5, 10))

        def editar(event):
            iid = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if not iid or not col:
                return
            col_idx = int(col.replace("#", "")) - 1
            bbox = tree.bbox(iid, col)
            if not bbox:
                return
            x, y, w, h = bbox
            valor = tree.item(iid, "values")[col_idx]

            entry = ttk.Entry(tree)
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, valor)
            entry.focus_set()
            entry.selection_range(0, "end")

            def salvar_edicao(_=None):
                vals = list(tree.item(iid, "values"))
                vals[col_idx] = entry.get()
                tree.item(iid, values=vals)
                entry.destroy()

            entry.bind("<Return>", salvar_edicao)
            entry.bind("<Escape>", lambda e: entry.destroy())
            entry.bind("<FocusOut>", salvar_edicao)

        tree.bind("<Double-1>", editar)

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=3, column=0, sticky="e", pady=(14, 0))

        def salvar():
            item["tabela"] = [list(tree.item(iid, "values")) for iid in tree.get_children()]
            self._atualizar_lista()
            win.destroy()

        ctk.CTkButton(
            botoes,
            text="Cancelar",
            width=100,
            fg_color=("gray88", "gray26"),
            hover_color=("gray80", "gray32"),
            text_color=("gray20", "gray90"),
            command=win.destroy,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            botoes, text="Salvar tabela", width=130, fg_color=AZUL, hover_color=AZUL_HOVER, command=salvar
        ).pack(side="left")

    def mover(self, direcao):
        idx = self._indice_selecionado()
        if idx is None:
            return
        novo = idx + direcao
        if novo < 0 or novo >= len(self.itens):
            return
        self.itens[idx], self.itens[novo] = self.itens[novo], self.itens[idx]
        self._atualizar_lista(novo)
        try:
            self.winfo_toplevel().salvar_rascunho(manual=False)
        except Exception:
            pass

    def remover(self):
        idx = self._indice_selecionado()
        if idx is None:
            return
        nome = self.itens[idx].get("titulo", "") or "item selecionado"
        if not messagebox.askyesno("Remover publicação", f"Deseja remover:\n\n{nome}?"):
            return
        self.itens.pop(idx)
        self._atualizar_lista(min(idx, len(self.itens) - 1))
        self.limpar_campos()
        try:
            self.winfo_toplevel().salvar_rascunho(manual=False)
        except Exception:
            pass

    def limpar_campos(self):
        self.editando_indice = None
        self.cmb_nivel.set(NIVEIS[1])
        self.ent_titulo.delete(0, "end")
        self.txt.delete("1.0", "end")
        self.btn_adicionar.configure(text="Adicionar publicação")
        self._atualizar_contador()

    def exportar_rascunho(self):
        """
        Salva tanto as publicações já adicionadas quanto o texto que ainda está
        sendo digitado no formulário e não foi adicionado.
        """
        return {
            "itens": self.itens,
            "editor": {
                "nivel": self._nivel_atual(),
                "titulo": self.ent_titulo.get(),
                "texto": self.txt.get("1.0", "end-1c"),
                "editando_indice": self.editando_indice,
            },
        }

    def restaurar_rascunho(self, dados):
        dados = dados or {}
        self.itens = list(dados.get("itens") or [])
        self._atualizar_lista()

        editor = dados.get("editor") or {}
        nivel = editor.get("nivel", "subtitulo")
        self.cmb_nivel.set(NIVEIS[0] if nivel == "titulo" else NIVEIS[1])

        self.ent_titulo.delete(0, "end")
        self.ent_titulo.insert(0, editor.get("titulo", ""))

        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", editor.get("texto", ""))

        idx = editor.get("editando_indice")
        if isinstance(idx, int) and 0 <= idx < len(self.itens):
            self.editando_indice = idx
            self.btn_adicionar.configure(text="Salvar alteração")
        else:
            self.editando_indice = None
            self.btn_adicionar.configure(text="Adicionar publicação")

        self._atualizar_contador()


    def obter_itens(self):
        return list(self.itens)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("blue")

        self.title("Gerador de Boletim Interno - COGER")
        self.geometry("1360x860")
        self.minsize(1120, 740)

        self.frames_partes = []
        self.nav_buttons = []
        self.parte_atual = 0
        # None = boletim ainda não salvo no histórico.
        # Inteiro = registro atualmente aberto para edição.
        self.boletim_id_atual = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._configurar_treeview()
        self._montar_sidebar()
        self._montar_area_principal()

        self.bind("<Control-g>", lambda e: self.gerar("ambos"))
        self.bind("<F6>", lambda e: self.verificar_ocr())

        self.mostrar_parte(0)
        self._atualizar_resumo()

        # Recupera automaticamente o último trabalho salvo.
        self._restaurar_rascunho_inicial()

        # Salvamento automático periódico. Também salva ao fechar a janela.
        self._autosave_job = None
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)
        self._agendar_autosalvamento()

    def _configurar_treeview(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Modern.Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground="#26323D",
            rowheight=32,
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.configure(
            "Modern.Treeview.Heading",
            background="#EEF2F5",
            foreground="#293642",
            font=("Segoe UI Semibold", 9),
            relief="flat",
            padding=8,
        )
        style.map(
            "Modern.Treeview",
            background=[("selected", "#DCEAF5")],
            foreground=[("selected", "#1E2933")],
        )
        style.map(
            "Modern.Treeview.Heading",
            background=[("active", "#E3EAF0")],
        )

    def _montar_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=238, corner_radius=0, fg_color=AZUL)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_columnconfigure(0, weight=1)
        self.sidebar.grid_rowconfigure(9, weight=1)

        logo_box = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_box.grid(row=0, column=0, sticky="ew", padx=18, pady=(22, 8))

        if LOGO_PATH.exists():
            try:
                img = Image.open(LOGO_PATH)
                self.logo_ctk = ctk.CTkImage(light_image=img, dark_image=img, size=(58, 70))
                ctk.CTkLabel(logo_box, text="", image=self.logo_ctk).pack(side="left", padx=(0, 10))
            except Exception:
                pass

        tit = ctk.CTkFrame(logo_box, fg_color="transparent")
        tit.pack(side="left", fill="x")
        ctk.CTkLabel(
            tit,
            text="COGER",
            text_color="white",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            tit,
            text="Boletim Interno",
            text_color="#D8E4EE",
            font=ctk.CTkFont(size=11),
        ).pack(anchor="w")

        ctk.CTkLabel(
            self.sidebar,
            text="PARTES DO BOLETIM",
            text_color="#AFC2D2",
            font=ctk.CTkFont(size=10, weight="bold"),
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(18, 8))

        for i, nome in enumerate(PARTES_CURTAS):
            btn = ctk.CTkButton(
                self.sidebar,
                text=f"{i+1:02d}   {nome}",
                anchor="w",
                height=42,
                corner_radius=10,
                fg_color="transparent",
                hover_color=AZUL_HOVER,
                text_color="white",
                font=ctk.CTkFont(size=13, weight="bold"),
                command=lambda idx=i: self.mostrar_parte(idx),
            )
            btn.grid(row=2+i, column=0, sticky="ew", padx=14, pady=3)
            self.nav_buttons.append(btn)

        ctk.CTkFrame(self.sidebar, height=1, fg_color="#345873").grid(
            row=6, column=0, sticky="ew", padx=18, pady=(18, 12)
        )

        ctk.CTkButton(
            self.sidebar,
            text="Assinaturas",
            anchor="w",
            height=38,
            corner_radius=9,
            fg_color="transparent",
            hover_color=AZUL_HOVER,
            text_color="white",
            command=self.configuracoes,
        ).grid(row=7, column=0, sticky="ew", padx=14, pady=2)

        ctk.CTkButton(
            self.sidebar,
            text="Verificar Word / LibreOffice",
            anchor="w",
            height=38,
            corner_radius=9,
            fg_color="transparent",
            hover_color=AZUL_HOVER,
            text_color="white",
            command=self.verificar_office,
        ).grid(row=8, column=0, sticky="ew", padx=14, pady=2)

        ctk.CTkButton(
            self.sidebar,
            text="Verificar OCR",
            anchor="w",
            height=38,
            corner_radius=9,
            fg_color="transparent",
            hover_color=AZUL_HOVER,
            text_color="white",
            command=self.verificar_ocr,
        ).grid(row=9, column=0, sticky="ew", padx=14, pady=2)

        ctk.CTkButton(
            self.sidebar,
            text="Salvar boletim",
            anchor="w",
            height=38,
            corner_radius=9,
            fg_color="transparent",
            hover_color=AZUL_HOVER,
            text_color="white",
            command=lambda: self.salvar_boletim(manual=True),
        ).grid(row=10, column=0, sticky="ew", padx=14, pady=(10, 2))

        ctk.CTkButton(
            self.sidebar,
            text="Boletins salvos",
            anchor="w",
            height=38,
            corner_radius=9,
            fg_color="transparent",
            hover_color=AZUL_HOVER,
            text_color="white",
            command=self.abrir_historico,
        ).grid(row=11, column=0, sticky="ew", padx=14, pady=2)

        ctk.CTkButton(
            self.sidebar,
            text="Salvar rascunho agora",
            anchor="w",
            height=38,
            corner_radius=9,
            fg_color="transparent",
            hover_color=AZUL_HOVER,
            text_color="white",
            command=lambda: self.salvar_rascunho(manual=True),
        ).grid(row=12, column=0, sticky="ew", padx=14, pady=2)

        ctk.CTkButton(
            self.sidebar,
            text="Novo boletim",
            anchor="w",
            height=38,
            corner_radius=9,
            fg_color="transparent",
            hover_color=AZUL_HOVER,
            text_color="white",
            command=self.novo_boletim,
        ).grid(row=13, column=0, sticky="ew", padx=14, pady=2)

        rodape = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        rodape.grid(row=14, column=0, sticky="ew", padx=16, pady=14)
        rodape.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            rodape,
            text="Aparência",
            text_color="#D8E4EE",
            font=ctk.CTkFont(size=11),
        ).grid(row=0, column=0, sticky="w")

        self.switch_tema = ctk.CTkSwitch(
            rodape,
            text="Modo escuro",
            text_color="white",
            progress_color="#6EA4CC",
            button_color="white",
            button_hover_color="#E6EEF4",
            command=self.alternar_tema,
        )
        self.switch_tema.grid(row=1, column=0, sticky="w", pady=(5, 0))

    def _montar_area_principal(self):
        self.main = ctk.CTkFrame(self, fg_color=("gray96", "#15191D"), corner_radius=0)
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(2, weight=1)

        # Topbar
        top = ctk.CTkFrame(self.main, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 12))
        top.grid_columnconfigure(0, weight=1)

        self.lbl_top_titulo = ctk.CTkLabel(
            top,
            text="1ª Parte",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        self.lbl_top_titulo.grid(row=0, column=0, sticky="w")

        self.lbl_status_top = ctk.CTkLabel(
            top,
            text="Novo boletim",
            corner_radius=10,
            fg_color=(AZUL_CLARO, "#213448"),
            text_color=(AZUL, "#BFD7EA"),
            font=ctk.CTkFont(size=11, weight="bold"),
            padx=12,
            pady=5,
        )
        self.lbl_status_top.grid(row=0, column=1, sticky="e")

        # Dados gerais
        dados = ctk.CTkFrame(self.main, corner_radius=14)
        dados.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 14))
        dados.grid_columnconfigure((1, 3, 5, 7), weight=1)

        ctk.CTkLabel(
            dados, text="Dados do boletim", font=ctk.CTkFont(size=15, weight="bold")
        ).grid(row=0, column=0, columnspan=8, sticky="w", padx=18, pady=(15, 10))

        labels = ["Número", "Local", "Data inicial", "Data final"]
        for c, txt in zip((0, 2, 4, 6), labels):
            ctk.CTkLabel(
                dados, text=txt, font=ctk.CTkFont(size=11, weight="bold")
            ).grid(row=1, column=c, sticky="w", padx=(18 if c == 0 else 8, 6), pady=(0, 14))

        self.ent_numero = ctk.CTkEntry(dados, height=36, corner_radius=9, placeholder_text="038")
        self.ent_numero.grid(row=1, column=1, sticky="ew", padx=(0, 12), pady=(0, 14))

        self.ent_local = ctk.CTkEntry(dados, height=36, corner_radius=9)
        self.ent_local.insert(0, "Curitiba")
        self.ent_local.grid(row=1, column=3, sticky="ew", padx=(0, 12), pady=(0, 14))

        self.ent_inicio = CalendarDateField(
            dados,
            placeholder="Selecione a data",
            on_change=self._atualizar_resumo,
        )
        self.ent_inicio.grid(row=1, column=5, sticky="ew", padx=(0, 12), pady=(0, 14))

        self.ent_fim = CalendarDateField(
            dados,
            placeholder="Selecione a data",
            on_change=self._atualizar_resumo,
        )
        self.ent_fim.grid(row=1, column=7, sticky="ew", padx=(0, 18), pady=(0, 14))

        for ent in (self.ent_numero, self.ent_local):
            ent.bind("<KeyRelease>", lambda e: self._atualizar_resumo())

        # Container das partes
        self.conteudo = ctk.CTkFrame(self.main, fg_color="transparent")
        self.conteudo.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 14))
        self.conteudo.grid_columnconfigure(0, weight=1)
        self.conteudo.grid_rowconfigure(0, weight=1)

        for parte in PARTES:
            view = ParteView(self.conteudo, parte, on_change=self._atualizar_resumo)
            view.grid(row=0, column=0, sticky="nsew")
            self.frames_partes.append(view)

        # Barra inferior de geração
        bottom = ctk.CTkFrame(self.main, corner_radius=14)
        bottom.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 20))
        bottom.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(
            bottom,
            text="Pronto para editar.",
            text_color=("gray45", "gray70"),
            font=ctk.CTkFont(size=11),
        )
        self.lbl_status.grid(row=0, column=0, sticky="w", padx=18, pady=14)

        btns = ctk.CTkFrame(bottom, fg_color="transparent")
        btns.grid(row=0, column=1, sticky="e", padx=14, pady=10)

        ctk.CTkButton(
            btns,
            text="DOCX",
            width=96,
            height=38,
            corner_radius=10,
            fg_color=("gray87", "gray26"),
            hover_color=("gray79", "gray32"),
            text_color=("gray20", "gray90"),
            command=lambda: self.gerar("docx"),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btns,
            text="PDF",
            width=96,
            height=38,
            corner_radius=10,
            fg_color=("gray87", "gray26"),
            hover_color=("gray79", "gray32"),
            text_color=("gray20", "gray90"),
            command=lambda: self.gerar("pdf"),
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btns,
            text="Gerar DOCX + PDF",
            width=166,
            height=40,
            corner_radius=10,
            fg_color=VERDE,
            hover_color=VERDE_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda: self.gerar("ambos"),
        ).pack(side="left")

    def _dados_rascunho(self):
        return {
            "versao": 1,
            "numero": self.ent_numero.get().strip(),
            "local": self.ent_local.get().strip(),
            "data_inicio": self.ent_inicio.get().strip(),
            "data_fim": self.ent_fim.get().strip(),
            "parte_atual": self.parte_atual,
            "partes": [frame.exportar_rascunho() for frame in self.frames_partes],
        }

    def salvar_rascunho(self, manual=False):
        try:
            atualizado = salvar_rascunho_db(self._dados_rascunho())
            hora = datetime.fromisoformat(atualizado).strftime("%H:%M:%S")
            if hasattr(self, "lbl_status"):
                self.lbl_status.configure(text=f"Rascunho salvo automaticamente às {hora}.")
            if manual:
                messagebox.showinfo(
                    "Rascunho salvo",
                    "O boletim em edição foi salvo com sucesso.\n\n"
                    "Se o programa fechar ou travar, ele será restaurado na próxima abertura."
                )
            return True
        except Exception as e:
            if manual:
                messagebox.showerror("Erro ao salvar rascunho", str(e))
            return False

    def _agendar_autosalvamento(self):
        # A cada 3 segundos. É um JSON pequeno gravado no SQLite local.
        try:
            self.salvar_rascunho(manual=False)
        finally:
            self._autosave_job = self.after(3000, self._agendar_autosalvamento)

    def _restaurar_rascunho_inicial(self):
        dados, atualizado_em = carregar_rascunho_db()
        if not dados:
            return

        try:
            # O rascunho é uma proteção temporária e não assume automaticamente
            # que está vinculado a um registro do histórico.
            self.boletim_id_atual = None

            self.ent_numero.delete(0, "end")
            self.ent_numero.insert(0, dados.get("numero", ""))

            self.ent_local.delete(0, "end")
            self.ent_local.insert(0, dados.get("local", "") or "Curitiba")

            self.ent_inicio.delete(0, "end")
            self.ent_inicio.insert(0, dados.get("data_inicio", ""))

            self.ent_fim.delete(0, "end")
            self.ent_fim.insert(0, dados.get("data_fim", ""))

            partes = dados.get("partes") or []
            for i, frame in enumerate(self.frames_partes):
                if i < len(partes):
                    frame.restaurar_rascunho(partes[i])

            parte = dados.get("parte_atual", 0)
            if not isinstance(parte, int) or parte < 0 or parte >= len(self.frames_partes):
                parte = 0
            self.mostrar_parte(parte)
            self._atualizar_resumo()

            if atualizado_em:
                try:
                    dt = datetime.fromisoformat(atualizado_em)
                    quando = dt.strftime("%d/%m/%Y às %H:%M:%S")
                except Exception:
                    quando = atualizado_em
                self.lbl_status.configure(text=f"Rascunho restaurado — último salvamento: {quando}.")
            else:
                self.lbl_status.configure(text="Rascunho anterior restaurado automaticamente.")
        except Exception as e:
            messagebox.showwarning(
                "Rascunho",
                "Foi encontrado um rascunho, mas não foi possível restaurá-lo completamente.\n\n"
                f"Detalhes: {e}"
            )

    def _aplicar_dados_editor(self, dados, boletim_id=None):
        """
        Carrega no editor um rascunho ou um boletim salvo.
        """
        dados = dados or {}
        self.boletim_id_atual = boletim_id

        self.ent_numero.delete(0, "end")
        self.ent_numero.insert(0, dados.get("numero", ""))

        self.ent_local.delete(0, "end")
        self.ent_local.insert(0, dados.get("local", "") or "Curitiba")

        self.ent_inicio.delete(0, "end")
        self.ent_inicio.insert(0, dados.get("data_inicio", ""))

        self.ent_fim.delete(0, "end")
        self.ent_fim.insert(0, dados.get("data_fim", ""))

        partes = dados.get("partes") or []
        for i, frame in enumerate(self.frames_partes):
            if i < len(partes):
                frame.restaurar_rascunho(partes[i])
            else:
                frame.restaurar_rascunho({})

        parte = dados.get("parte_atual", 0)
        if not isinstance(parte, int) or parte < 0 or parte >= len(self.frames_partes):
            parte = 0

        self.mostrar_parte(parte)
        self._atualizar_resumo()

    def salvar_boletim(self, manual=True):
        numero = self.ent_numero.get().strip()
        if not numero:
            if manual:
                messagebox.showwarning(
                    "Número do boletim",
                    "Informe o número do boletim antes de salvá-lo no histórico."
                )
                self.ent_numero.focus_set()
            return False

        try:
            dados = self._dados_rascunho()
            self.boletim_id_atual, atualizado = salvar_boletim_db(
                dados, self.boletim_id_atual
            )

            # Mantém o rascunho sincronizado com a versão salva.
            salvar_rascunho_db(dados)

            hora = datetime.fromisoformat(atualizado).strftime("%d/%m/%Y às %H:%M:%S")
            self.lbl_status.configure(
                text=f"BI nº {numero} salvo no histórico em {hora}."
            )

            if manual:
                messagebox.showinfo(
                    "Boletim salvo",
                    f"O BI nº {numero} foi salvo no histórico.\n\n"
                    "Você poderá reabri-lo e continuar a edição quando quiser."
                )
            return True
        except Exception as e:
            if manual:
                messagebox.showerror("Erro ao salvar boletim", str(e))
            return False

    def abrir_boletim_salvo(self, boletim_id, janela=None):
        # Protege o trabalho atual antes de trocar de boletim.
        self.salvar_rascunho(manual=False)

        dados = carregar_boletim_db(boletim_id)
        if not dados:
            messagebox.showerror("Boletim", "Não foi possível carregar o boletim selecionado.")
            return

        self._aplicar_dados_editor(dados, boletim_id=boletim_id)
        salvar_rascunho_db(self._dados_rascunho())

        if janela is not None:
            try:
                janela.destroy()
            except Exception:
                pass

        self.lbl_status.configure(
            text=f"BI nº {self.ent_numero.get().strip()} aberto para edição."
        )

    def _gerar_boletim_salvo(self, boletim_id):
        dados_editor = carregar_boletim_db(boletim_id)
        if not dados_editor:
            messagebox.showerror("Boletim", "Não foi possível carregar o boletim selecionado.")
            return

        dados = dados_documento_de_editor(dados_editor)

        if not dados["numero"] or not dados["data_inicio"] or not dados["data_fim"]:
            messagebox.showwarning(
                "Dados incompletos",
                "Esse boletim não possui número ou período completo."
            )
            return

        nome_base = f"BI_{dados['numero']}"
        caminho_docx = SAIDA_DIR / f"{nome_base}.docx"
        caminho_pdf = SAIDA_DIR / f"{nome_base}.pdf"

        try:
            gerar_docx(dados, caminho_docx, atualizar_word=True)
            gerar_pdf(dados, caminho_pdf, caminho_docx_existente=caminho_docx)
            messagebox.showinfo(
                "Boletim gerado",
                "DOCX e PDF foram gerados novamente:\n\n"
                f"{caminho_docx}\n{caminho_pdf}"
            )
        except Exception as e:
            messagebox.showerror("Erro ao gerar documento", str(e))

    def abrir_historico(self):
        win = ModernDialog(self, "Boletins salvos", "1040x650")
        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.grid(row=0, column=0, sticky="nsew", padx=22, pady=22)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(3, weight=1)

        ctk.CTkLabel(
            frame,
            text="Boletins salvos",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            frame,
            text="Abra um boletim antigo para continuar a edição ou gere novamente o DOCX/PDF.",
            text_color=("gray45", "gray70"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 12))

        busca_box = ctk.CTkFrame(frame, fg_color="transparent")
        busca_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        busca_box.grid_columnconfigure(0, weight=1)

        ent_busca = ctk.CTkEntry(
            busca_box,
            height=38,
            corner_radius=9,
            placeholder_text="Buscar por número, local ou data...",
        )
        ent_busca.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        tabela_box = ctk.CTkFrame(frame, corner_radius=12)
        tabela_box.grid(row=3, column=0, sticky="nsew")
        tabela_box.grid_rowconfigure(0, weight=1)
        tabela_box.grid_columnconfigure(0, weight=1)

        tree = ttk.Treeview(
            tabela_box,
            columns=("numero", "periodo", "local", "atualizado"),
            show="headings",
            style="Modern.Treeview",
            selectmode="browse",
        )
        tree.heading("numero", text="BI")
        tree.heading("periodo", text="Período")
        tree.heading("local", text="Local")
        tree.heading("atualizado", text="Última alteração")

        tree.column("numero", width=90, anchor="center")
        tree.column("periodo", width=240)
        tree.column("local", width=170)
        tree.column("atualizado", width=180)

        scroll = ttk.Scrollbar(tabela_box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
        scroll.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=10)

        def formatar_data_iso(valor):
            try:
                return datetime.fromisoformat(valor).strftime("%d/%m/%Y %H:%M")
            except Exception:
                return valor or ""

        def atualizar_lista(*_):
            for item in tree.get_children():
                tree.delete(item)

            for row in listar_boletins_db(ent_busca.get()):
                bid, numero, local, ini, fim, criado, atualizado = row
                periodo = f"{ini or '...'} a {fim or '...'}"
                tree.insert(
                    "",
                    "end",
                    iid=str(bid),
                    values=(
                        numero,
                        periodo,
                        local or "",
                        formatar_data_iso(atualizado),
                    ),
                )

        def selecionado():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Boletins salvos", "Selecione um boletim.")
                return None
            return int(sel[0])

        def abrir():
            bid = selecionado()
            if bid is not None:
                self.abrir_boletim_salvo(bid, janela=win)

        def duplicar():
            bid = selecionado()
            if bid is None:
                return
            novo_id = duplicar_boletim_db(bid)
            if novo_id:
                atualizar_lista()
                tree.selection_set(str(novo_id))
                tree.focus(str(novo_id))
                messagebox.showinfo(
                    "Boletim duplicado",
                    "Foi criada uma cópia independente do boletim selecionado."
                )

        def excluir():
            bid = selecionado()
            if bid is None:
                return

            vals = tree.item(str(bid), "values")
            numero = vals[0] if vals else ""

            if not messagebox.askyesno(
                "Excluir boletim",
                f"Deseja excluir o BI nº {numero} do histórico?\n\n"
                "Essa ação não apaga DOCX/PDF já gerados."
            ):
                return

            excluir_boletim_db(bid)

            if self.boletim_id_atual == bid:
                self.boletim_id_atual = None

            atualizar_lista()

        def gerar():
            bid = selecionado()
            if bid is not None:
                self._gerar_boletim_salvo(bid)

        ent_busca.bind("<KeyRelease>", atualizar_lista)
        tree.bind("<Double-1>", lambda e: abrir())

        ctk.CTkButton(
            busca_box,
            text="Buscar",
            width=90,
            command=atualizar_lista,
        ).grid(row=0, column=1)

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=4, column=0, sticky="ew", pady=(12, 0))

        ctk.CTkButton(
            botoes,
            text="Abrir / Editar",
            width=120,
            fg_color=AZUL,
            hover_color=AZUL_HOVER,
            command=abrir,
        ).pack(side="left")

        ctk.CTkButton(
            botoes,
            text="Duplicar",
            width=100,
            command=duplicar,
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            botoes,
            text="Gerar DOCX + PDF",
            width=145,
            command=gerar,
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            botoes,
            text="Excluir",
            width=90,
            fg_color="#A83A3A",
            hover_color="#8D2F2F",
            command=excluir,
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            botoes,
            text="Fechar",
            width=90,
            fg_color=("gray87", "gray26"),
            hover_color=("gray79", "gray32"),
            text_color=("gray20", "gray90"),
            command=win.destroy,
        ).pack(side="right")

        atualizar_lista()
        ent_busca.focus_set()


    def novo_boletim(self):
        if not messagebox.askyesno(
            "Novo boletim",
            "Deseja iniciar um novo boletim?\n\n"
            "O rascunho atual será apagado."
        ):
            return

        excluir_rascunho_db()
        self.boletim_id_atual = None

        self.ent_numero.delete(0, "end")
        self.ent_local.delete(0, "end")
        self.ent_local.insert(0, "Curitiba")
        self.ent_inicio.delete(0, "end")
        self.ent_fim.delete(0, "end")

        for frame in self.frames_partes:
            frame.itens = []
            frame._atualizar_lista()
            frame.limpar_campos()

        self.mostrar_parte(0)
        self._atualizar_resumo()
        self.lbl_status.configure(text="Novo boletim iniciado.")

        # Salva imediatamente o estado limpo.
        self.salvar_rascunho(manual=False)

    def _ao_fechar(self):
        try:
            if self._autosave_job is not None:
                self.after_cancel(self._autosave_job)
        except Exception:
            pass

        # Última tentativa de persistir inclusive o texto ainda não adicionado.
        self.salvar_rascunho(manual=False)
        self.destroy()


    def alternar_tema(self):
        modo = "Dark" if self.switch_tema.get() else "Light"
        ctk.set_appearance_mode(modo)

        style = ttk.Style()
        if modo == "Dark":
            style.configure(
                "Modern.Treeview",
                background="#202428",
                fieldbackground="#202428",
                foreground="#E8EDF1",
            )
            style.configure(
                "Modern.Treeview.Heading",
                background="#2A3035",
                foreground="#E8EDF1",
            )
            style.map(
                "Modern.Treeview",
                background=[("selected", "#294860")],
                foreground=[("selected", "#FFFFFF")],
            )
        else:
            self._configurar_treeview()

    def mostrar_parte(self, idx):
        self.parte_atual = idx
        self.frames_partes[idx].tkraise()
        self.lbl_top_titulo.configure(text=PARTES_CURTAS[idx])

        for i, btn in enumerate(self.nav_buttons):
            btn.configure(
                fg_color="#2A5273" if i == idx else "transparent",
                text_color="white",
            )

    def _atualizar_resumo(self):
        total = sum(len(f.itens) for f in self.frames_partes)
        numero = self.ent_numero.get().strip() if hasattr(self, "ent_numero") else ""
        ini = self.ent_inicio.get().strip() if hasattr(self, "ent_inicio") else ""
        fim = self.ent_fim.get().strip() if hasattr(self, "ent_fim") else ""

        texto = f"{total} publicação" if total == 1 else f"{total} publicações"
        if numero:
            texto = f"BI nº {numero}  •  {texto}"
        if ini or fim:
            texto += f"  •  {ini or '...'} a {fim or '...'}"
        self.lbl_status_top.configure(text=texto)

    def verificar_office(self):
        suite = detectar_suite()
        if suite["tipo"] == "word":
            messagebox.showinfo("Suíte de documentos", "Microsoft Word encontrado.\n\nO programa usará o Word para atualizar índice, paginação e gerar o PDF.")
            return
        if suite["tipo"] == "libreoffice":
            messagebox.showinfo("Suíte de documentos", "Microsoft Word não foi encontrado, mas o LibreOffice está disponível.\n\nO programa usará o LibreOffice para atualizar o documento e gerar o PDF.\n\nPode haver pequenas diferenças de layout/paginação em relação ao Word.")
            return
        messagebox.showwarning("Suíte de documentos não encontrada", "Não foi encontrado Microsoft Word nem LibreOffice.\n\nO DOCX ainda pode ser gerado, mas PDF e atualização automática do índice/paginação não estarão disponíveis.")

    def verificar_ocr(self):
        caminho = localizar_tesseract()
        if caminho:
            self.lbl_status.configure(text="OCR disponível e pronto para uso.")
            messagebox.showinfo("OCR disponível", f"Tesseract encontrado em:\n{caminho}")
        else:
            self.lbl_status.configure(text="OCR não localizado neste computador.")
            messagebox.showwarning(
                "OCR não encontrado",
                "Tesseract OCR não foi encontrado.\n\nInstale o Tesseract OCR para usar a importação de prints.",
            )

    def validar(self):
        if not self.ent_numero.get().strip():
            messagebox.showwarning("Dados incompletos", "Informe o número do boletim.")
            self.ent_numero.focus_set()
            return False
        if not self.ent_inicio.get().strip():
            messagebox.showwarning("Dados incompletos", "Informe a data inicial.")
            self.ent_inicio.focus_set()
            return False
        if not self.ent_fim.get().strip():
            messagebox.showwarning("Dados incompletos", "Informe a data final.")
            self.ent_fim.focus_set()
            return False
        return True

    def dados_boletim(self):
        assina_nome, assina_funcao, confere_nome, confere_funcao = carregar_config()
        return {
            "numero": self.ent_numero.get().strip(),
            "local": self.ent_local.get().strip() or "Curitiba",
            "data_inicio": self.ent_inicio.get().strip(),
            "data_fim": self.ent_fim.get().strip(),
            "partes": [
                {"titulo": PARTES[i], "itens": self.frames_partes[i].obter_itens()}
                for i in range(4)
            ],
            "assina_nome": assina_nome,
            "assina_funcao": assina_funcao,
            "confere_nome": confere_nome,
            "confere_funcao": confere_funcao,
        }

    def gerar(self, tipo):
        if not self.validar():
            return

        dados = self.dados_boletim()
        nome_base = f"BI_{dados['numero']}"
        caminho_docx = SAIDA_DIR / f"{nome_base}.docx"
        caminho_pdf = SAIDA_DIR / f"{nome_base}.pdf"

        self.lbl_status.configure(text="Gerando documento... o Word pode abrir rapidamente em segundo plano.")
        self.update_idletasks()

        try:
            arquivos = []

            if tipo == "docx":
                gerar_docx(dados, caminho_docx, atualizar_word=True)
                arquivos.append(str(caminho_docx))
            elif tipo == "pdf":
                gerar_pdf(dados, caminho_pdf)
                arquivos.append(str(caminho_pdf))
            else:
                gerar_docx(dados, caminho_docx, atualizar_word=True)
                gerar_pdf(dados, caminho_pdf, caminho_docx_existente=caminho_docx)
                arquivos.extend([str(caminho_docx), str(caminho_pdf)])

            # Documento gerado com sucesso: mantém também uma versão editável no histórico.
            try:
                self.boletim_id_atual, _ = salvar_boletim_db(
                    self._dados_rascunho(), self.boletim_id_atual
                )
            except Exception:
                pass

            self.lbl_status.configure(text="Documento gerado com sucesso.")
            messagebox.showinfo(
                "Boletim gerado",
                "Arquivo(s) criado(s) com sucesso:\n\n" + "\n".join(arquivos),
            )
        except Exception as e:
            self.lbl_status.configure(text="Ocorreu um erro durante a geração.")
            messagebox.showerror("Erro ao gerar documento", str(e))

    def configuracoes(self):
        win = ModernDialog(self, "Assinaturas do boletim", "720x430")
        frame = ctk.CTkFrame(win, fg_color="transparent")
        frame.grid(row=0, column=0, sticky="nsew", padx=24, pady=24)
        frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame, text="Assinaturas", font=ctk.CTkFont(size=22, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            frame,
            text="Esses dados serão aplicados automaticamente no final do boletim.",
            text_color=("gray45", "gray70"),
        ).grid(row=1, column=0, sticky="w", pady=(2, 16))

        card = ctk.CTkFrame(frame, corner_radius=14)
        card.grid(row=2, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        vals = carregar_config()
        campos = [
            ("ASSINA — Nome / posto", vals[0]),
            ("ASSINA — Função", vals[1]),
            ("CONFERE — Nome / posto", vals[2]),
            ("CONFERE — Função", vals[3]),
        ]

        entries = []
        for i, (label, valor) in enumerate(campos):
            ctk.CTkLabel(
                card, text=label, font=ctk.CTkFont(size=11, weight="bold")
            ).grid(row=i*2, column=0, sticky="w", padx=18, pady=(14 if i == 0 else 9, 4))
            e = ctk.CTkEntry(card, height=36, corner_radius=9)
            e.insert(0, valor or "")
            e.grid(row=i*2+1, column=0, sticky="ew", padx=18, pady=(0, 4 if i < 3 else 16))
            entries.append(e)

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=3, column=0, sticky="e", pady=(16, 0))

        def salvar():
            salvar_config(tuple(e.get().strip() for e in entries))
            self.lbl_status.configure(text="Configurações de assinatura atualizadas.")
            win.destroy()

        ctk.CTkButton(
            botoes,
            text="Cancelar",
            width=100,
            fg_color=("gray87", "gray26"),
            hover_color=("gray79", "gray32"),
            text_color=("gray20", "gray90"),
            command=win.destroy,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            botoes,
            text="Salvar",
            width=110,
            fg_color=AZUL,
            hover_color=AZUL_HOVER,
            command=salvar,
        ).pack(side="left")


if __name__ == "__main__":
    init_db()
    App().mainloop()
