# -*- coding: utf-8 -*-
"""
InvenTI - Controle de Inventário de TI (v5.0)

Dependências:  pip install customtkinter openpyxl reportlab
"""
import hashlib
import hmac
import os
import secrets
import sqlite3
import subprocess
import sys
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape

import customtkinter as ctk
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ==============================================================================
# TEMA / CONSTANTES
# ==============================================================================
APP_NOME = "InvenTI"
FONTE = "Segoe UI"
ITERACOES_HASH = 200_000

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Cores no formato (modo claro, modo escuro) — o customtkinter troca sozinho
COR = {
    "bg": ("#F3F4F8", "#0F1117"),
    "sidebar": ("#FFFFFF", "#151821"),
    "card": ("#FFFFFF", "#1A1D29"),
    "borda": ("#E3E6EE", "#272C3D"),
    "texto": ("#111827", "#F3F4F6"),
    "texto2": ("#6B7280", "#9CA3AF"),
    "entrada": ("#F5F6FA", "#232838"),
    "entrada_borda": ("#D5D9E6", "#333A50"),
    "accent": "#6366F1",
    "accent_h": "#4F46E5",
}

ESTILOS_BOTAO = {
    "primario": dict(fg_color=COR["accent"], hover_color=COR["accent_h"], text_color="#FFFFFF"),
    "neutro": dict(fg_color=("#E8EAF3", "#252A3B"), hover_color=("#DCE0EE", "#30364D"), text_color=COR["texto"]),
    "perigo": dict(fg_color=("#FDE8E8", "#3B1F26"), hover_color=("#FBD2D2", "#4D2630"), text_color=("#DC2626", "#F87171")),
    "aviso": dict(fg_color=("#FEF3C7", "#3A2F14"), hover_color=("#FDE68A", "#4B3B18"), text_color=("#B45309", "#FBBF24")),
    "ok": dict(fg_color=("#D1FAE5", "#123528"), hover_color=("#A7F3D0", "#184632"), text_color=("#047857", "#34D399")),
}

STATUS_LISTA = ["Ativo", "Em Manutenção", "Desativado"]
ICONE_STATUS = {"Ativo": "●", "Em Manutenção": "◐", "Desativado": "○"}

# (chave, título, coluna no banco, largura, alinhamento, estica?)
COLUNAS = [
    ("id", "ID", "id", 50, "center", False),
    ("marca", "Marca", "marca", 110, "w", False),
    ("modelo", "Modelo", "modelo", 160, "w", True),
    ("setor", "Setor", "setor", 120, "w", False),
    ("proc", "Processador", "processador", 150, "w", True),
    ("ram", "RAM", "ram", 70, "center", False),
    ("gpu", "Placa de vídeo", "gpu", 130, "w", False),
    ("status", "Status", "status", 140, "w", False),
    ("obs", "Observações", "obs", 180, "w", True),
]


# ------------------------------------------------------------------------------
# Opções padrão das listas de seleção (você pode editar à vontade).
# Tudo que for cadastrado no sistema também passa a aparecer nas listas.
# ------------------------------------------------------------------------------
MODELOS_POR_MARCA = {
    "Dell": ["Latitude 3420", "Latitude 5420", "Latitude 5430", "Inspiron 15 3520",
             "Vostro 3510", "OptiPlex 3080", "OptiPlex 7090"],
    "HP": ["240 G8", "ProBook 440 G8", "ProBook 450 G9", "EliteBook 840 G8",
           "ProDesk 400 G7", "Pavilion 15"],
    "Lenovo": ["ThinkPad E14", "ThinkPad L14", "ThinkBook 15", "IdeaPad 3",
               "V14 G2", "ThinkCentre M70s"],
    "Acer": ["Aspire 3", "Aspire 5", "TravelMate P2", "Veriton X"],
    "Asus": ["VivoBook 15", "ExpertBook B1", "TUF Gaming F15", "ROG Strix"],
    "Apple": ["MacBook Air M1", "MacBook Air M2", "MacBook Pro 14", "iMac 24"],
    "Samsung": ["Galaxy Book", "Galaxy Book2", "Odyssey"],
    "Positivo": ["Motion", "Master", "Vision"],
}
SETORES_PADRAO = [
    "TI", "Administrativo", "Financeiro", "RH", "Comercial", "Compras",
    "Logística", "Produção", "Recepção", "Diretoria", "Suporte",
]
PROCESSADORES_PADRAO = (
    [f"Core {fam} {ger}ª Gen" for fam in ("i3", "i5", "i7") for ger in range(8, 14)]
    + ["Core i9", "Pentium", "Celeron", "Ryzen 3", "Ryzen 5", "Ryzen 7", "Ryzen 9",
       "Apple M1", "Apple M2", "Apple M3"]
)
RAM_PADRAO = ["4GB", "8GB", "12GB", "16GB", "24GB", "32GB", "64GB"]
GPU_PADRAO = [
    "Integrada", "Intel UHD Graphics", "Intel Iris Xe", "AMD Radeon Vega",
    "NVIDIA GeForce MX450", "NVIDIA GeForce GTX 1650", "NVIDIA GeForce RTX 3050",
    "NVIDIA GeForce RTX 3060", "NVIDIA GeForce RTX 4060", "AMD Radeon RX 6600",
]


def botao(master, texto, comando, tipo="neutro", **kw):
    kw.setdefault("height", 38)
    kw.setdefault("corner_radius", 10)
    kw.setdefault("font", (FONTE, 13, "bold"))
    return ctk.CTkButton(master, text=texto, command=comando, **ESTILOS_BOTAO[tipo], **kw)


def entrada(master, **kw):
    kw.setdefault("height", 40)
    kw.setdefault("corner_radius", 10)
    kw.setdefault("font", (FONTE, 13))
    return ctk.CTkEntry(
        master,
        fg_color=COR["entrada"],
        border_color=COR["entrada_borda"],
        text_color=COR["texto"],
        **kw,
    )

def aplicar_estilo_tabela():
    escuro = ctk.get_appearance_mode() == "Dark"
    bg = "#1A1D29" if escuro else "#FFFFFF"
    fg = "#F3F4F6" if escuro else "#111827"
    cab = "#222639" if escuro else "#EEF0F7"

    estilo = ttk.Style()
    estilo.theme_use("clam")
    estilo.layout("Moderno.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    estilo.configure(
        "Moderno.Treeview",
        background=bg,
        fieldbackground=bg,
        foreground=fg,
        rowheight=38,
        borderwidth=0,
        font=(FONTE, 10),
    )
    estilo.configure(
        "Moderno.Treeview.Heading",
        background=cab,
        foreground=fg,
        relief="flat",
        borderwidth=0,
        font=(FONTE, 10, "bold"),
        padding=(10, 9),
    )
    estilo.map(
        "Moderno.Treeview",
        background=[("selected", COR["accent"])],
        foreground=[("selected", "#FFFFFF")],
    )
    estilo.map("Moderno.Treeview.Heading", background=[("active", cab)])

# ==============================================================================
# BANCO DE DADOS
# ==============================================================================
def pasta_app():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def conectar_banco():
    return sqlite3.connect(os.path.join(pasta_app(), "inventario_ti.db"))


def gerar_hash(senha):
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), salt, ITERACOES_HASH)
    return f"pbkdf2${salt.hex()}${dk.hex()}"


def verificar_senha(senha, armazenado):
    try:
        _, salt_hex, hash_hex = armazenado.split("$")
        dk = hashlib.pbkdf2_hmac(
            "sha256", senha.encode("utf-8"), bytes.fromhex(salt_hex), ITERACOES_HASH
        )
        return hmac.compare_digest(dk.hex(), hash_hex)
    except ValueError:
        return False


def autenticar(usuario, senha):
    con = conectar_banco()
    linha = con.execute(
        "SELECT senha, perfil FROM usuarios WHERE usuario = ?", (usuario,)
    ).fetchone()
    con.close()
    if linha and verificar_senha(senha, linha[0]):
        return linha[1]
    return None


def inicializar_banco():
    con = conectar_banco()
    cur = con.cursor()

    cur.execute(
        """CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            senha TEXT NOT NULL,
            perfil TEXT NOT NULL
        )"""
    )
    cur.execute(
        """CREATE TABLE IF NOT EXISTS equipamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            marca TEXT NOT NULL,
            modelo TEXT NOT NULL,
            setor TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Ativo'
        )"""
    )

    # Migração de colunas novas
    cur.execute("PRAGMA table_info(equipamentos)")
    existentes = [c[1] for c in cur.fetchall()]
    for coluna in ("processador", "ram", "gpu", "obs"):
        if coluna not in existentes:
            cur.execute(f"ALTER TABLE equipamentos ADD COLUMN {coluna} TEXT")

    # Migração: senhas antigas em texto puro passam a ser guardadas com hash
    cur.execute("SELECT id, senha FROM usuarios")
    for uid, senha in cur.fetchall():
        if not senha.startswith("pbkdf2$"):
            cur.execute("UPDATE usuarios SET senha = ? WHERE id = ?", (gerar_hash(senha), uid))

    # Primeiro uso: cria o administrador padrão (troque a senha em "Alterar senha")
    cur.execute("SELECT COUNT(*) FROM usuarios")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO usuarios (usuario, senha, perfil) VALUES (?, ?, ?)",
            ("admin", gerar_hash("admin123"), "Admin"),
        )

    con.commit()
    con.close()

# ==============================================================================
# DIÁLOGOS
# ==============================================================================
def _mesclar(padrao, extras):
    """Padrão primeiro; depois o que já foi cadastrado (sem repetir)."""
    vistos = {v.lower() for v in padrao}
    resultado = list(padrao)
    for v in sorted(extras, key=str.lower):
        if v.lower() not in vistos:
            vistos.add(v.lower())
            resultado.append(v)
    return resultado


def opcoes_lista(coluna, padrao):
    con = conectar_banco()
    linhas = con.execute(
        f"SELECT DISTINCT TRIM({coluna}) FROM equipamentos "
        f"WHERE {coluna} IS NOT NULL AND TRIM({coluna}) != ''"
    ).fetchall()
    con.close()
    return _mesclar(padrao, [l[0] for l in linhas])


def opcoes_modelos(marca):
    padrao = []
    for m, modelos in MODELOS_POR_MARCA.items():
        if m.lower() == marca.strip().lower():
            padrao = modelos
    con = conectar_banco()
    linhas = con.execute(
        "SELECT DISTINCT TRIM(modelo) FROM equipamentos "
        "WHERE modelo != '' AND marca = ? COLLATE NOCASE",
        (marca.strip(),),
    ).fetchall()
    con.close()
    return _mesclar(padrao, [l[0] for l in linhas])


class Dialogo(ctk.CTkToplevel):
    """Janela modal centralizada sobre a janela principal."""

    def __init__(self, master, titulo, largura, altura):
        super().__init__(master)
        self.title(titulo)
        self.configure(fg_color=COR["bg"])
        self.resizable(False, False)
        raiz = master.winfo_toplevel()
        x = raiz.winfo_rootx() + (raiz.winfo_width() - largura) // 2
        y = raiz.winfo_rooty() + (raiz.winfo_height() - altura) // 2
        self.geometry(f"{largura}x{altura}+{max(x, 0)}+{max(y, 0)}")
        self.transient(raiz)
        self.bind("<Escape>", lambda _e: self.destroy())
        self.after(150, self._focar)

    def _focar(self):
        try:
            self.grab_set()
            self.focus_force()
        except Exception:
            pass

    def titulo(self, texto, sub=None):
        ctk.CTkLabel(
            self, text=texto, font=(FONTE, 22, "bold"), text_color=COR["texto"]
        ).pack(anchor="w", padx=28, pady=(24, 2))
        if sub:
            ctk.CTkLabel(
                self, text=sub, font=(FONTE, 12), text_color=COR["texto2"]
            ).pack(anchor="w", padx=28, pady=(0, 14))


class FormEquipamento(Dialogo):
    def __init__(self, master, ao_salvar, id_=None, dados=None):
        editar = id_ is not None
        super().__init__(master, "Editar equipamento" if editar else "Novo equipamento", 620, 690)
        self.ao_salvar = ao_salvar
        self.id_ = id_
        dados = dados or {}

        self.titulo(
            "Editar equipamento" if editar else "Novo equipamento",
            "Escolha na lista ou digite para adicionar um valor novo. * = obrigatório.",
        )

        corpo = ctk.CTkFrame(
            self, fg_color=COR["card"], corner_radius=14,
            border_width=1, border_color=COR["borda"],
        )
        corpo.pack(fill="both", expand=True, padx=24)
        corpo.grid_columnconfigure((0, 1), weight=1, uniform="c")
        self.campos = {}

        def rotulo(texto, linha, coluna, span=1):
            ctk.CTkLabel(
                corpo, text=texto, font=(FONTE, 12, "bold"),
                text_color=COR["texto2"], anchor="w",
            ).grid(row=linha, column=coluna, columnspan=span, sticky="w", padx=16, pady=(14, 2))

        def campo(chave, texto, linha, coluna, placeholder, span=1):
            rotulo(texto, linha * 2, coluna, span)
            e = entrada(corpo, placeholder_text=placeholder)
            e.grid(row=linha * 2 + 1, column=coluna, columnspan=span, sticky="ew", padx=16)
            if dados.get(chave):
                e.insert(0, dados[chave])
            self.campos[chave] = e

        def combo(chave, texto, linha, coluna, valores, comando=None):
            rotulo(texto, linha * 2, coluna)
            cb = ctk.CTkComboBox(
                corpo, values=valores, height=40, corner_radius=10, command=comando,
                font=(FONTE, 13), dropdown_font=(FONTE, 13),
                fg_color=COR["entrada"], border_color=COR["entrada_borda"],
                button_color=COR["accent"], button_hover_color=COR["accent_h"],
                text_color=COR["texto"],
            )
            cb.set(dados.get(chave) or "")
            cb.grid(row=linha * 2 + 1, column=coluna, sticky="ew", padx=16)
            self.campos[chave] = cb
            return cb

        def atualizar_modelos(_valor=None, limpar=False):
            novos = opcoes_modelos(self.campos["marca"].get())
            self.campos["modelo"].configure(values=novos)
            if limpar and self.campos["modelo"].get() not in novos:
                self.campos["modelo"].set("")

        marcas = _mesclar(list(MODELOS_POR_MARCA), opcoes_lista("marca", []))
        combo("marca", "Marca *", 0, 0, marcas, lambda v: atualizar_modelos(v, limpar=True))
        combo("modelo", "Modelo *", 0, 1, opcoes_modelos(dados.get("marca", "")))
        self.campos["marca"].bind("<KeyRelease>", lambda _e: atualizar_modelos())
        combo("setor", "Setor *", 1, 0, opcoes_lista("setor", SETORES_PADRAO))

        rotulo("Status", 2, 1)
        self.cb_status = ctk.CTkComboBox(
            corpo, values=STATUS_LISTA, height=40, corner_radius=10, state="readonly",
            font=(FONTE, 13), fg_color=COR["entrada"], border_color=COR["entrada_borda"],
            button_color=COR["accent"], button_hover_color=COR["accent_h"],
            text_color=COR["texto"],
        )
        self.cb_status.set(dados.get("status") or "Ativo")
        self.cb_status.grid(row=3, column=1, sticky="ew", padx=16)

        combo("processador", "Processador", 2, 0, opcoes_lista("processador", PROCESSADORES_PADRAO))
        combo("ram", "Memória RAM", 3, 0, opcoes_lista("ram", RAM_PADRAO))
        combo("gpu", "Placa de vídeo", 3, 1, opcoes_lista("gpu", GPU_PADRAO))

        rotulo("Observações", 8, 0, 2)
        self.txt_obs = ctk.CTkTextbox(
            corpo, height=80, corner_radius=10, border_width=2, font=(FONTE, 13),
            fg_color=COR["entrada"], border_color=COR["entrada_borda"],
            text_color=COR["texto"],
        )
        self.txt_obs.grid(row=9, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 16))
        if dados.get("obs"):
            self.txt_obs.insert("1.0", dados["obs"])

        self.lbl_erro = ctk.CTkLabel(self, text="", font=(FONTE, 12), text_color="#EF4444")
        self.lbl_erro.pack(anchor="w", padx=28, pady=(8, 0))

        rodape = ctk.CTkFrame(self, fg_color="transparent")
        rodape.pack(fill="x", padx=24, pady=(4, 20))
        botao(rodape, "Cancelar", self.destroy, "neutro", width=110).pack(side="right", padx=(8, 0))
        botao(rodape, "Salvar", self.salvar, "primario", width=130).pack(side="right")
        self.after(200, self.campos["marca"].focus)

    def salvar(self):
        d = {k: e.get().strip() for k, e in self.campos.items()}
        d["obs"] = self.txt_obs.get("1.0", "end").strip()
        d["status"] = self.cb_status.get()
        if not (d["marca"] and d["modelo"] and d["setor"]):
            self.lbl_erro.configure(text="Marca, Modelo e Setor são obrigatórios.")
            return
        self.ao_salvar(d, self.id_)
        self.destroy()


class DialogoSenha(Dialogo):
    def __init__(self, master, usuario):
        super().__init__(master, "Alterar senha", 420, 470)
        self.usuario = usuario
        self.titulo("Alterar senha", f"Usuário: {usuario}")

        self.ent_atual = self._campo("Senha atual")
        self.ent_nova = self._campo("Nova senha (mínimo 6 caracteres)")
        self.ent_conf = self._campo("Confirmar nova senha")

        self.lbl_msg = ctk.CTkLabel(self, text="", font=(FONTE, 12), text_color="#EF4444")
        self.lbl_msg.pack(anchor="w", padx=28, pady=(10, 0))
        botao(self, "Salvar nova senha", self.salvar, "primario").pack(fill="x", padx=28, pady=(8, 0))
        self.ent_conf.bind("<Return>", lambda _e: self.salvar())

    def _campo(self, texto):
        ctk.CTkLabel(
            self, text=texto, font=(FONTE, 12, "bold"), text_color=COR["texto2"]
        ).pack(anchor="w", padx=28, pady=(8, 2))
        e = entrada(self, show="•")
        e.pack(fill="x", padx=28)
        return e

    def salvar(self):
        atual, nova, conf = self.ent_atual.get(), self.ent_nova.get(), self.ent_conf.get()
        if autenticar(self.usuario, atual) is None:
            self.lbl_msg.configure(text="A senha atual está incorreta.")
        elif len(nova) < 6:
            self.lbl_msg.configure(text="A nova senha precisa ter ao menos 6 caracteres.")
        elif nova != conf:
            self.lbl_msg.configure(text="A confirmação não confere com a nova senha.")
        else:
            con = conectar_banco()
            con.execute("UPDATE usuarios SET senha = ? WHERE usuario = ?", (gerar_hash(nova), self.usuario))
            con.commit()
            con.close()
            messagebox.showinfo("Senha alterada", "Sua senha foi atualizada com sucesso.", parent=self)
            self.destroy()


class DialogoUsuarios(Dialogo):
    def __init__(self, master, usuario_atual):
        super().__init__(master, "Gerenciar usuários", 480, 620)
        self.usuario_atual = usuario_atual
        self.titulo("Usuários", "Quem pode acessar o sistema.")

        self.lista = ctk.CTkScrollableFrame(
            self, fg_color=COR["card"], corner_radius=14, height=230,
            border_width=1, border_color=COR["borda"],
        )
        self.lista.pack(fill="x", padx=24, pady=(0, 14))

        ctk.CTkLabel(
            self, text="Novo usuário", font=(FONTE, 14, "bold"), text_color=COR["texto"]
        ).pack(anchor="w", padx=28)
        self.ent_user = entrada(self, placeholder_text="Nome de usuário")
        self.ent_user.pack(fill="x", padx=28, pady=(8, 6))
        self.ent_pass = entrada(self, placeholder_text="Senha (mínimo 6 caracteres)", show="•")
        self.ent_pass.pack(fill="x", padx=28, pady=(0, 6))
        self.cb_perfil = ctk.CTkComboBox(
            self, values=["Usuário", "Admin"], height=40, corner_radius=10, state="readonly",
            fg_color=COR["entrada"], border_color=COR["entrada_borda"],
            button_color=COR["accent"], button_hover_color=COR["accent_h"],
            text_color=COR["texto"],
        )
        self.cb_perfil.set("Usuário")
        self.cb_perfil.pack(fill="x", padx=28, pady=(0, 6))
        self.lbl_msg = ctk.CTkLabel(self, text="", font=(FONTE, 12), text_color="#EF4444")
        self.lbl_msg.pack(anchor="w", padx=28)
        botao(self, "Adicionar usuário", self.adicionar, "primario").pack(fill="x", padx=28, pady=(4, 0))
        self.atualizar()

    def atualizar(self):
        for w in self.lista.winfo_children():
            w.destroy()
        con = conectar_banco()
        linhas = con.execute("SELECT usuario, perfil FROM usuarios ORDER BY usuario").fetchall()
        con.close()
        for usuario, perfil in linhas:
            linha = ctk.CTkFrame(self.lista, fg_color="transparent")
            linha.pack(fill="x", pady=3)
            ctk.CTkLabel(
                linha, text=usuario, font=(FONTE, 13, "bold"), text_color=COR["texto"]
            ).pack(side="left", padx=(8, 10))
            ctk.CTkLabel(
                linha, text=perfil, font=(FONTE, 12), text_color=COR["texto2"]
            ).pack(side="left")
            if usuario != self.usuario_atual:
                botao(
                    linha, "Remover", lambda u=usuario: self.remover(u), "perigo",
                    width=84, height=28, font=(FONTE, 12, "bold"),
                ).pack(side="right", padx=6)

    def adicionar(self):
        usuario, senha = self.ent_user.get().strip(), self.ent_pass.get()
        if not usuario or len(senha) < 6:
            self.lbl_msg.configure(text="Informe o usuário e uma senha de ao menos 6 caracteres.")
            return
        try:
            con = conectar_banco()
            con.execute(
                "INSERT INTO usuarios (usuario, senha, perfil) VALUES (?, ?, ?)",
                (usuario, gerar_hash(senha), self.cb_perfil.get()),
            )
            con.commit()
            con.close()
        except sqlite3.IntegrityError:
            self.lbl_msg.configure(text="Esse nome de usuário já existe.")
            return
        self.lbl_msg.configure(text="")
        self.ent_user.delete(0, "end")
        self.ent_pass.delete(0, "end")
        self.atualizar()

    def remover(self, usuario):
        if messagebox.askyesno("Remover usuário", f"Remover o usuário “{usuario}”?", parent=self):
            con = conectar_banco()
            con.execute("DELETE FROM usuarios WHERE usuario = ?", (usuario,))
            con.commit()
            con.close()
            self.atualizar()


# ==============================================================================
# TELA DE LOGIN
# ==============================================================================
class LoginFrame(ctk.CTkFrame):
    def __init__(self, master, ao_entrar):
        super().__init__(master, fg_color=COR["bg"], corner_radius=0)
        self.ao_entrar = ao_entrar
        self.grid_columnconfigure(0, weight=5, uniform="l")
        self.grid_columnconfigure(1, weight=6, uniform="l")
        self.grid_rowconfigure(0, weight=1)

        # ---- Painel da marca (esquerda)
        marca = ctk.CTkFrame(self, fg_color=COR["accent"], corner_radius=0)
        marca.grid(row=0, column=0, sticky="nsew")
        centro = ctk.CTkFrame(marca, fg_color="transparent")
        centro.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(centro, text="⬢", font=(FONTE, 54), text_color="#FFFFFF").pack(anchor="w")
        ctk.CTkLabel(centro, text=APP_NOME, font=(FONTE, 34, "bold"), text_color="#FFFFFF").pack(anchor="w")
        ctk.CTkLabel(
            centro, text="Controle de inventário de TI", font=(FONTE, 14), text_color="#E0E7FF"
        ).pack(anchor="w", pady=(0, 28))
        for item in (
            "Todos os equipamentos em um só lugar",
            "Busca e filtros instantâneos",
            "Relatórios em PDF e Excel",
        ):
            ctk.CTkLabel(
                centro, text=f"✓  {item}", font=(FONTE, 13), text_color="#EEF2FF"
            ).pack(anchor="w", pady=3)

        # ---- Formulário (direita)
        direita = ctk.CTkFrame(self, fg_color="transparent")
        direita.grid(row=0, column=1, sticky="nsew")
        form = ctk.CTkFrame(direita, fg_color="transparent", width=330)
        form.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            form, text="Bem-vindo de volta", font=(FONTE, 26, "bold"), text_color=COR["texto"]
        ).pack(anchor="w")
        ctk.CTkLabel(
            form, text="Entre com suas credenciais para continuar.",
            font=(FONTE, 13), text_color=COR["texto2"],
        ).pack(anchor="w", pady=(2, 24))

        ctk.CTkLabel(
            form, text="Usuário", font=(FONTE, 12, "bold"), text_color=COR["texto2"]
        ).pack(anchor="w", pady=(0, 4))
        self.ent_usuario = entrada(form, placeholder_text="Seu usuário", width=330, height=44)
        self.ent_usuario.pack()

        ctk.CTkLabel(
            form, text="Senha", font=(FONTE, 12, "bold"), text_color=COR["texto2"]
        ).pack(anchor="w", pady=(14, 4))
        self.ent_senha = entrada(form, placeholder_text="Sua senha", show="•", width=330, height=44)
        self.ent_senha.pack()

        self.var_mostrar = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            form, text="Mostrar senha", variable=self.var_mostrar, command=self._alternar_senha,
            font=(FONTE, 12), text_color=COR["texto2"], fg_color=COR["accent"],
            hover_color=COR["accent_h"], checkbox_width=18, checkbox_height=18,
        ).pack(anchor="w", pady=(10, 0))

        self.lbl_erro = ctk.CTkLabel(form, text="", font=(FONTE, 12), text_color="#EF4444")
        self.lbl_erro.pack(anchor="w", pady=(8, 0))

        botao(form, "Entrar", self.entrar, "primario", width=330, height=46).pack(pady=(4, 0))

        for e in (self.ent_usuario, self.ent_senha):
            e.bind("<Return>", self.entrar)
        self.after(200, self.ent_usuario.focus)

    def _alternar_senha(self):
        self.ent_senha.configure(show="" if self.var_mostrar.get() else "•")

    def entrar(self, _evento=None):
        usuario = self.ent_usuario.get().strip()
        senha = self.ent_senha.get()
        if not usuario or not senha:
            self.lbl_erro.configure(text="Informe usuário e senha.")
            return
        perfil = autenticar(usuario, senha)
        if perfil:
            self.ao_entrar(usuario, perfil)
        else:
            self.lbl_erro.configure(text="Usuário ou senha inválidos.")
            self.ent_senha.delete(0, "end")
            self.ent_senha.focus()

# ==============================================================================
# TELA PRINCIPAL
# ==============================================================================
class MainFrame(ctk.CTkFrame):
    def __init__(self, master, usuario, perfil, ao_sair):
        super().__init__(master, fg_color=COR["bg"], corner_radius=0)
        self.usuario = usuario
        self.perfil = perfil
        self.ao_sair = ao_sair
        self.eh_admin = perfil == "Admin"
        self.filtro = "Todos"
        self.ordem = ("id", False)  # (chave da coluna, crescente?)
        self.dados_atuais = []
        self._toast_id = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        aplicar_estilo_tabela()
        self._montar_sidebar()
        self._montar_conteudo()
        self.carregar()

    # ---------------------------------------------------------------- sidebar
    def _secao(self, master, texto):
        ctk.CTkLabel(
            master, text=texto, font=(FONTE, 10, "bold"), text_color=COR["texto2"]
        ).pack(anchor="w", padx=24, pady=(18, 4))

    def _montar_sidebar(self):
        sb = ctk.CTkFrame(self, width=240, corner_radius=0, fg_color=COR["sidebar"])
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.pack_propagate(False)

        topo = ctk.CTkFrame(sb, fg_color="transparent")
        topo.pack(fill="x", padx=22, pady=(26, 8))
        ctk.CTkLabel(topo, text="⬢", font=(FONTE, 28), text_color=COR["accent"]).pack(side="left")
        caixa = ctk.CTkFrame(topo, fg_color="transparent")
        caixa.pack(side="left", padx=(10, 0))
        ctk.CTkLabel(
            caixa, text=APP_NOME, font=(FONTE, 18, "bold"), text_color=COR["texto"]
        ).pack(anchor="w")
        ctk.CTkLabel(
            caixa, text="Inventário de TI", font=(FONTE, 11), text_color=COR["texto2"]
        ).pack(anchor="w")

        self._secao(sb, "AÇÕES")
        botao(sb, "＋   Novo equipamento", self.novo, "primario", anchor="w").pack(fill="x", padx=18, pady=3)

        self._secao(sb, "RELATÓRIOS")
        botao(sb, "▤   Exportar Excel", self.exportar_excel, "neutro", anchor="w").pack(fill="x", padx=18, pady=3)
        botao(sb, "▣   Exportar PDF", self.exportar_pdf, "neutro", anchor="w").pack(fill="x", padx=18, pady=3)

        if self.eh_admin:
            self._secao(sb, "ADMINISTRAÇÃO")
            botao(
                sb, "☺   Usuários", lambda: DialogoUsuarios(self, self.usuario), "neutro", anchor="w"
            ).pack(fill="x", padx=18, pady=3)

        # Rodapé: usuário, tema e sair
        rodape = ctk.CTkFrame(sb, fg_color="transparent")
        rodape.pack(side="bottom", fill="x", padx=18, pady=18)

        cartao = ctk.CTkFrame(rodape, fg_color=("#F3F4F8", "#1D2130"), corner_radius=12)
        cartao.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            cartao, text=self.usuario[:1].upper(), width=38, height=38, corner_radius=19,
            fg_color=COR["accent"], text_color="#FFFFFF", font=(FONTE, 15, "bold"),
        ).pack(side="left", padx=10, pady=10)
        info = ctk.CTkFrame(cartao, fg_color="transparent")
        info.pack(side="left")
        ctk.CTkLabel(
            info, text=self.usuario, font=(FONTE, 13, "bold"), text_color=COR["texto"]
        ).pack(anchor="w")
        ctk.CTkLabel(
            info, text=self.perfil, font=(FONTE, 11), text_color=COR["texto2"]
        ).pack(anchor="w")

        self.sw_tema = ctk.CTkSwitch(
            rodape, text="Modo escuro", command=self.alternar_tema, font=(FONTE, 12),
            text_color=COR["texto2"], progress_color=COR["accent"],
        )
        if ctk.get_appearance_mode() == "Dark":
            self.sw_tema.select()
        self.sw_tema.pack(anchor="w", padx=6, pady=(0, 10))

        botao(rodape, "Alterar senha", lambda: DialogoSenha(self, self.usuario), "neutro", height=34).pack(fill="x", pady=3)
        botao(rodape, "Sair", self.ao_sair, "perigo", height=34).pack(fill="x", pady=3)

    # --------------------------------------------------------------- conteúdo
    def _card_stat(self, master, titulo, cor, coluna):
        card = ctk.CTkFrame(
            master, corner_radius=14, fg_color=COR["card"],
            border_width=1, border_color=COR["borda"],
        )
        card.grid(row=0, column=coluna, sticky="ew", padx=(0 if coluna == 0 else 6, 0 if coluna == 3 else 6))
        ctk.CTkFrame(card, width=4, corner_radius=2, fg_color=cor).pack(
            side="left", fill="y", padx=(14, 0), pady=16
        )
        caixa = ctk.CTkFrame(card, fg_color="transparent")
        caixa.pack(side="left", padx=14, pady=12)
        num = ctk.CTkLabel(caixa, text="0", font=(FONTE, 28, "bold"), text_color=COR["texto"])
        num.pack(anchor="w")
        ctk.CTkLabel(caixa, text=titulo, font=(FONTE, 12), text_color=COR["texto2"]).pack(anchor="w")
        return num

    def _montar_conteudo(self):
        c = ctk.CTkFrame(self, fg_color="transparent")
        c.grid(row=0, column=1, sticky="nsew", padx=28, pady=24)
        c.grid_columnconfigure(0, weight=1)
        c.grid_rowconfigure(3, weight=1)

        # Cabeçalho
        cab = ctk.CTkFrame(c, fg_color="transparent")
        cab.grid(row=0, column=0, sticky="ew")
        cab.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            cab, text="Equipamentos", font=(FONTE, 28, "bold"), text_color=COR["texto"], anchor="w"
        ).grid(row=0, column=0, sticky="w")
        self.lbl_sub = ctk.CTkLabel(
            cab, text="", font=(FONTE, 13), text_color=COR["texto2"], anchor="w"
        )
        self.lbl_sub.grid(row=1, column=0, sticky="w")
        self.lbl_toast = ctk.CTkLabel(
            cab, text="", corner_radius=8, text_color="#FFFFFF", font=(FONTE, 12, "bold"), height=32
        )
        self.lbl_toast.grid(row=0, column=1, rowspan=2, sticky="e")
        self.lbl_toast.grid_remove()

        # Indicadores
        stats = ctk.CTkFrame(c, fg_color="transparent")
        stats.grid(row=1, column=0, sticky="ew", pady=(18, 14))
        stats.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="s")
        self.stats = {
            "total": self._card_stat(stats, "Total de equipamentos", COR["accent"], 0),
            "Ativo": self._card_stat(stats, "Ativos", "#10B981", 1),
            "Em Manutenção": self._card_stat(stats, "Em manutenção", "#F59E0B", 2),
            "Desativado": self._card_stat(stats, "Desativados", "#EF4444", 3),
        }

        # Barra de busca + filtros
        barra = ctk.CTkFrame(c, fg_color="transparent")
        barra.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        barra.grid_columnconfigure(0, weight=1)
        self.ent_busca = entrada(
            barra, placeholder_text="🔍  Buscar por marca, modelo, setor, processador, ID…", height=40
        )
        self.ent_busca.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.ent_busca.bind("<KeyRelease>", lambda _e: self.carregar())

        chips = ctk.CTkFrame(barra, fg_color="transparent")
        chips.grid(row=0, column=1)
        self.chips = {}
        for nome in ["Todos"] + STATUS_LISTA:
            b = botao(chips, nome, lambda n=nome: self.filtrar_status(n), "neutro", height=40, width=96)
            b.pack(side="left", padx=(0, 6))
            self.chips[nome] = b
        self._pintar_chips()

        # Tabela
        caixa = ctk.CTkFrame(
            c, fg_color=COR["card"], corner_radius=14, border_width=1, border_color=COR["borda"]
        )
        caixa.grid(row=3, column=0, sticky="nsew")
        caixa.grid_columnconfigure(0, weight=1)
        caixa.grid_rowconfigure(0, weight=1)

        self.tabela = ttk.Treeview(
            caixa, columns=[col[0] for col in COLUNAS], show="headings",
            style="Moderno.Treeview", selectmode="extended",
        )
        for chave, titulo, _sql, largura, ancora, estica in COLUNAS:
            self.tabela.column(chave, width=largura, minwidth=40, anchor=ancora, stretch=estica)
        self.tabela.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        sb = ctk.CTkScrollbar(caixa, command=self.tabela.yview)
        sb.grid(row=0, column=1, sticky="ns", padx=(4, 8), pady=12)
        self.tabela.configure(yscrollcommand=sb.set)

        self.lbl_vazio = ctk.CTkLabel(
            caixa, text="Nenhum equipamento encontrado.", font=(FONTE, 14), text_color=COR["texto2"]
        )

        self.tabela.bind("<Double-1>", self._duplo_clique)
        self.tabela.bind("<Return>", self.editar)
        self.tabela.bind("<Delete>", lambda _e: self.excluir())
        self.aplicar_tema_tabela()

        # Ações sobre a seleção
        acoes = ctk.CTkFrame(c, fg_color="transparent")
        acoes.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        botao(acoes, "✎  Editar", self.editar, "neutro", width=110).pack(side="left", padx=(0, 8))
        botao(acoes, "Pôr em manutenção", lambda: self.alterar_status("Em Manutenção"), "aviso").pack(side="left", padx=(0, 8))
        botao(acoes, "Ativar", lambda: self.alterar_status("Ativo"), "ok", width=90).pack(side="left", padx=(0, 8))
        botao(acoes, "Desativar", lambda: self.alterar_status("Desativado"), "neutro", width=100).pack(side="left")
        self.btn_excluir = botao(acoes, "Excluir", self.excluir, "perigo", width=100)
        self.btn_excluir.pack(side="right")
        if not self.eh_admin:
            self.btn_excluir.configure(state="disabled")
        ctk.CTkLabel(
            acoes, text="Dica: duplo clique edita · Ctrl/Shift seleciona vários",
            font=(FONTE, 11), text_color=COR["texto2"],
        ).pack(side="right", padx=16)

    # ------------------------------------------------------------ utilidades
    def aplicar_tema_tabela(self):
        aplicar_estilo_tabela()
        escuro = ctk.get_appearance_mode() == "Dark"
        self.tabela.tag_configure("par", background="#1A1D29" if escuro else "#FFFFFF")
        self.tabela.tag_configure("impar", background="#1F2333" if escuro else "#F7F8FC")
        self.tabela.tag_configure("manut", foreground="#F59E0B" if escuro else "#B45309")
        self.tabela.tag_configure("desat", foreground="#6B7280" if escuro else "#9CA3AF")

    def alternar_tema(self):
        ctk.set_appearance_mode("Dark" if self.sw_tema.get() else "Light")
        self.aplicar_tema_tabela()

    def toast(self, msg, tipo="ok"):
        cor = {"ok": "#10B981", "erro": "#EF4444", "info": COR["accent"]}[tipo]
        self.lbl_toast.configure(text=f"   {msg}   ", fg_color=cor)
        self.lbl_toast.grid()
        if self._toast_id:
            self.after_cancel(self._toast_id)
        self._toast_id = self.after(3200, self._esconder_toast)

    def _esconder_toast(self):
        try:
            self.lbl_toast.grid_remove()
        except Exception:
            pass

    def _pintar_chips(self):
        for nome, b in self.chips.items():
            b.configure(**ESTILOS_BOTAO["primario" if nome == self.filtro else "neutro"])

    def filtrar_status(self, nome):
        self.filtro = nome
        self._pintar_chips()
        self.carregar()

    def ordenar(self, chave):
        atual, crescente = self.ordem
        self.ordem = (chave, not crescente) if atual == chave else (chave, True)
        self.carregar()

    def _cabecalhos(self):
        for chave, titulo, *_ in COLUNAS:
            seta = ""
            if self.ordem[0] == chave:
                seta = "  ▲" if self.ordem[1] else "  ▼"
            self.tabela.heading(chave, text=titulo + seta, anchor="w", command=lambda k=chave: self.ordenar(k))

    def ids_selecionados(self):
        return [int(i) for i in self.tabela.selection()]

    def _duplo_clique(self, evento):
        if self.tabela.identify_region(evento.x, evento.y) == "cell":
            self.editar()

    # ------------------------------------------------------------------ dados
    def carregar(self):
        termo = self.ent_busca.get().strip()
        colunas_sql = [col[2] for col in COLUNAS]
        sql = f"SELECT {', '.join(colunas_sql)} FROM equipamentos"
        condicoes, params = [], []
        if self.filtro != "Todos":
            condicoes.append("status = ?")
            params.append(self.filtro)
        if termo:
            campos = ["CAST(id AS TEXT)", "marca", "modelo", "setor", "processador", "ram", "gpu", "obs"]
            condicoes.append("(" + " OR ".join(f"{c} LIKE ?" for c in campos) + ")")
            params += [f"%{termo}%"] * len(campos)
        if condicoes:
            sql += " WHERE " + " AND ".join(condicoes)

        chave, crescente = self.ordem
        col_sql = next(col[2] for col in COLUNAS if col[0] == chave)
        sql += f" ORDER BY {col_sql}{'' if col_sql == 'id' else ' COLLATE NOCASE'} {'ASC' if crescente else 'DESC'}"

        con = conectar_banco()
        linhas = con.execute(sql, params).fetchall()
        contagem = dict(con.execute("SELECT status, COUNT(*) FROM equipamentos GROUP BY status").fetchall())
        con.close()

        anteriores = set(self.tabela.selection())
        self.tabela.delete(*self.tabela.get_children())
        self.dados_atuais = [tuple("" if v is None else v for v in l) for l in linhas]

        for n, linha in enumerate(self.dados_atuais):
            valores = list(linha)
            status = valores[7]
            valores[7] = f"{ICONE_STATUS.get(status, '•')}  {status}"
            tags = ["par" if n % 2 == 0 else "impar"]
            if status == "Em Manutenção":
                tags.append("manut")
            elif status == "Desativado":
                tags.append("desat")
            self.tabela.insert("", "end", iid=str(linha[0]), values=valores, tags=tags)

        manter = [i for i in anteriores if self.tabela.exists(i)]
        if manter:
            self.tabela.selection_set(manter)

        if self.dados_atuais:
            self.lbl_vazio.place_forget()
        else:
            self.lbl_vazio.place(relx=0.5, rely=0.5, anchor="center")
            self.lbl_vazio.lift()

        total = sum(contagem.values())
        self.stats["total"].configure(text=str(total))
        for st in STATUS_LISTA:
            self.stats[st].configure(text=str(contagem.get(st, 0)))
        self.lbl_sub.configure(text=f"Exibindo {len(self.dados_atuais)} de {total} equipamentos")
        self._cabecalhos()

    def novo(self):
        FormEquipamento(self, self.salvar_equipamento)

    def editar(self, _evento=None):
        ids = self.ids_selecionados()
        if len(ids) != 1:
            self.toast("Selecione um equipamento para editar", "info")
            return
        con = conectar_banco()
        linha = con.execute(
            "SELECT marca, modelo, setor, processador, ram, gpu, obs, status FROM equipamentos WHERE id = ?",
            (ids[0],),
        ).fetchone()
        con.close()
        if not linha:
            return
        chaves = ["marca", "modelo", "setor", "processador", "ram", "gpu", "obs", "status"]
        FormEquipamento(self, self.salvar_equipamento, ids[0], dict(zip(chaves, ("" if v is None else v for v in linha))))
        return "break"

    def salvar_equipamento(self, d, id_=None):
        valores = (d["marca"], d["modelo"], d["setor"], d["processador"], d["ram"], d["gpu"], d["obs"], d["status"])
        con = conectar_banco()
        if id_ is None:
            con.execute(
                "INSERT INTO equipamentos (marca, modelo, setor, processador, ram, gpu, obs, status) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                valores,
            )
        else:
            con.execute(
                "UPDATE equipamentos SET marca=?, modelo=?, setor=?, processador=?, ram=?, gpu=?, obs=?, status=? "
                "WHERE id=?",
                valores + (id_,),
            )
        con.commit()
        con.close()
        self.carregar()
        self.toast("Equipamento cadastrado" if id_ is None else "Alterações salvas")

    def alterar_status(self, novo_status):
        ids = self.ids_selecionados()
        if not ids:
            self.toast("Selecione ao menos um equipamento", "info")
            return
        con = conectar_banco()
        con.executemany("UPDATE equipamentos SET status = ? WHERE id = ?", [(novo_status, i) for i in ids])
        con.commit()
        con.close()
        self.carregar()
        self.toast(f"{len(ids)} item(ns) → {novo_status}")

    def excluir(self):
        if not self.eh_admin:
            self.toast("Apenas administradores podem excluir", "erro")
            return
        ids = self.ids_selecionados()
        if not ids:
            self.toast("Selecione ao menos um equipamento", "info")
            return
        if messagebox.askyesno("Confirmação", f"Excluir permanentemente {len(ids)} equipamento(s)?"):
            con = conectar_banco()
            con.executemany("DELETE FROM equipamentos WHERE id = ?", [(i,) for i in ids])
            con.commit()
            con.close()
            self.carregar()
            self.toast(f"{len(ids)} equipamento(s) excluído(s)")

    # ------------------------------------------------------------ exportações
    def _pedir_caminho(self, extensao, descricao):
        if not self.dados_atuais:
            self.toast("Não há dados para exportar", "info")
            return None
        return filedialog.asksaveasfilename(
            defaultextension=extensao,
            filetypes=[(descricao, f"*{extensao}")],
            initialfile=f"inventario_{datetime.now():%Y%m%d_%H%M}{extensao}",
            initialdir=pasta_app(),
        )

    def _concluir_exportacao(self, caminho):
        if messagebox.askyesno("Exportação concluída", f"Arquivo salvo em:\n{caminho}\n\nDeseja abri-lo agora?"):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(caminho)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", caminho])
                else:
                    subprocess.Popen(["xdg-open", caminho])
            except Exception:
                pass

    def exportar_excel(self):
        caminho = self._pedir_caminho(".xlsx", "Planilha Excel")
        if not caminho:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Inventário"
        ws.append(["ID", "Marca", "Modelo", "Setor", "Processador", "RAM", "Placa de vídeo", "Status", "Observações"])

        fill_cab = PatternFill("solid", fgColor="4F46E5")
        centro = Alignment(horizontal="center", vertical="center")
        esquerda = Alignment(vertical="center", wrap_text=True)
        borda = Border(bottom=Side(style="thin", color="E5E7EB"))
        for cel in ws[1]:
            cel.fill = fill_cab
            cel.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            cel.alignment = centro
        ws.row_dimensions[1].height = 26

        cores = {"Ativo": "047857", "Em Manutenção": "B45309", "Desativado": "B91C1C"}
        for i, linha in enumerate(self.dados_atuais, start=2):
            ws.append(list(linha))
            for j in range(1, 10):
                cel = ws.cell(i, j)
                cel.border = borda
                cel.alignment = centro if j in (1, 6, 8) else esquerda
                if i % 2 == 1:
                    cel.fill = PatternFill("solid", fgColor="F5F6FA")
            ws.cell(i, 8).font = Font(bold=True, color=cores.get(linha[7], "111827"))

        for col in ws.columns:
            largura = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(largura + 3, 10), 45)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        try:
            wb.save(caminho)
        except PermissionError:
            messagebox.showerror("Erro", "Não foi possível salvar. Feche o arquivo se ele estiver aberto.")
            return
        self._concluir_exportacao(caminho)

    def exportar_pdf(self):
        caminho = self._pedir_caminho(".pdf", "Documento PDF")
        if not caminho:
            return

        estilos = getSampleStyleSheet()
        celula = ParagraphStyle("celula", parent=estilos["BodyText"], fontName="Helvetica", fontSize=8.5, leading=11)
        cabecalho = ParagraphStyle("cabecalho", parent=celula, fontName="Helvetica-Bold", textColor=colors.white)
        titulo = ParagraphStyle("titulo", parent=estilos["Title"], fontName="Helvetica-Bold",
                                fontSize=18, textColor=colors.HexColor("#1E1B4B"), alignment=0)
        sub = ParagraphStyle("sub", parent=celula, fontSize=9, textColor=colors.HexColor("#6B7280"))

        pagina = landscape(A4)
        doc = SimpleDocTemplate(caminho, pagesize=pagina, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=36)

        titulos = ["ID", "Marca", "Modelo", "Setor", "Processador", "RAM", "Placa de vídeo", "Status", "Observações"]
        tabela = [[Paragraph(t, cabecalho) for t in titulos]]
        for linha in self.dados_atuais:
            tabela.append([Paragraph(escape(str(v)), celula) for v in linha])

        t = Table(tabela, colWidths=[28, 70, 95, 80, 110, 45, 100, 85, 169], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4F46E5")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F6FA")]),
            ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5E7EB")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))

        resumo = f"Emitido em {datetime.now():%d/%m/%Y às %H:%M} · {len(self.dados_atuais)} equipamento(s)"
        if self.filtro != "Todos" or self.ent_busca.get().strip():
            resumo += " · lista filtrada"

        def rodape(canvas, _doc):
            canvas.saveState()
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(colors.HexColor("#6B7280"))
            canvas.drawString(30, 18, f"{APP_NOME} — Relatório de inventário de TI")
            canvas.drawRightString(pagina[0] - 30, 18, f"Página {_doc.page}")
            canvas.restoreState()

        try:
            doc.build(
                [Paragraph("Relatório de Inventário de TI", titulo), Paragraph(resumo, sub), Spacer(1, 14), t],
                onFirstPage=rodape, onLaterPages=rodape,
            )
        except PermissionError:
            messagebox.showerror("Erro", "Não foi possível salvar. Feche o arquivo se ele estiver aberto.")
            return
        self._concluir_exportacao(caminho)


# ==============================================================================
# APLICAÇÃO
# ==============================================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NOME} — Controle de Inventário de TI")
        self.frame_atual = None
        self.mostrar_login()

    def _centralizar(self, largura, altura):
        x = (self.winfo_screenwidth() - largura) // 2
        y = max((self.winfo_screenheight() - altura) // 2 - 20, 0)
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    def _trocar(self, frame):
        if self.frame_atual is not None:
            self.frame_atual.destroy()
        self.frame_atual = frame
        frame.pack(fill="both", expand=True)

    def mostrar_login(self):
        self._trocar(LoginFrame(self, self._apos_login))
        self.resizable(False, False)
        self.minsize(0, 0)
        self._centralizar(920, 580)

    def _apos_login(self, usuario, perfil):
        self.after(30, lambda: self.mostrar_principal(usuario, perfil))

    def mostrar_principal(self, usuario, perfil):
        self._trocar(MainFrame(self, usuario, perfil, self.mostrar_login))
        self.resizable(True, True)
        self.minsize(1120, 680)
        self._centralizar(1320, 780)


if __name__ == "__main__":
    inicializar_banco()
    App().mainloop()
