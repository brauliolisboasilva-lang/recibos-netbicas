# -*- coding: utf-8 -*-
"""
Emissao de Recibos - NetBicas FIBRA
------------------------------------
Aplicativo local (roda no computador, sem internet) para gerar recibos
em PDF com numeracao crescente automatica, usando a logo e os dados
cadastrais da NetBicas Telecom.

Como funciona:
- Cada recibo emitido fica salvo em um historico (arquivo dados/recibos.json)
  e o numero do proximo recibo eh sugerido automaticamente.
- Ao clicar em "Salvar recibo (PDF)", o programa pede onde salvar o
  arquivo PDF no seu computador (pasta padrao: Documentos/Recibos NetBicas).
- E possivel reabrir/recarregar um recibo anterior a partir do historico.

Requisitos (ja incluidos no instalador): Python 3, reportlab, Pillow.
"""

import json
import os
import sys
import webbrowser
from datetime import date
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


# ---------------------------------------------------------------------------
# Caminhos e dados fixos da empresa
# ---------------------------------------------------------------------------

def base_dir() -> Path:
    """Pasta onde o app esta instalado (funciona tanto rodando o .py quanto
    um executavel empacotado com PyInstaller)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = base_dir()
ASSETS_DIR = APP_DIR / "assets"
DATA_DIR = APP_DIR / "dados"
DATA_FILE = DATA_DIR / "recibos.json"
LOGO_PATH = ASSETS_DIR / "logo.png"

DEFAULT_OUTPUT_DIR = Path.home() / "Documents" / "Recibos NetBicas"
if not (Path.home() / "Documents").exists():
    # em alguns sistemas em portugues a pasta se chama "Documentos"
    DEFAULT_OUTPUT_DIR = Path.home() / "Documentos" / "Recibos NetBicas"

EMPRESA = {
    "nome": "NETBICAS TELECOM",
    "cnpj": "38.468.300/0001-72",
    "endereco": "Rua Rio de Janeiro, nº 636 — Bairro Tereza Cristina",
    "cidade_cep": "São Joaquim de Bicas — MG — CEP 32920-000",
    "telefone": "(31) 3111-3538",
    "local": "São Joaquim de Bicas — MG",
}

NAVY = HexColor("#0a2d52")
BLUE = HexColor("#0d6fb8")
MUTED = HexColor("#64748b")
LINE = HexColor("#98a3b5")
INK = HexColor("#1c2330")

FORMAS_PAGAMENTO = [
    "Dinheiro",
    "PIX",
    "Transferência bancária",
    "Cartão de débito",
    "Cartão de crédito",
    "Boleto",
]


# ---------------------------------------------------------------------------
# Numero por extenso (igual ao usado no recibo original)
# ---------------------------------------------------------------------------

UNI = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
DEZ_ESP = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis",
           "dezessete", "dezoito", "dezenove"]
DEZ = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta",
       "setenta", "oitenta", "noventa"]
CEM = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos",
       "seiscentos", "setecentos", "oitocentos", "novecentos"]


def _trezentos(n: int) -> str:
    if n == 0:
        return ""
    if n == 100:
        return "cem"
    c, resto = divmod(n, 100)
    partes = []
    if c > 0:
        partes.append(CEM[c])
    if resto > 0:
        if resto < 10:
            partes.append(UNI[resto])
        elif resto < 20:
            partes.append(DEZ_ESP[resto - 10])
        else:
            d, u = divmod(resto, 10)
            partes.append(f"{DEZ[d]} e {UNI[u]}" if u > 0 else DEZ[d])
    return " e ".join(partes)


def _inteiro_extenso(n: int) -> str:
    if n == 0:
        return "zero"
    grupos = []
    t = n
    while t > 0:
        grupos.append(t % 1000)
        t //= 1000
    sing = ["", "mil", "milhão", "bilhão"]
    plur = ["", "mil", "milhões", "bilhões"]
    partes = []
    for i in range(len(grupos) - 1, -1, -1):
        g = grupos[i]
        if g == 0:
            continue
        texto = _trezentos(g)
        if i == 1:
            texto = "mil" if g == 1 else f"{texto} mil"
        elif i > 1:
            texto += " " + (sing[i] if g == 1 else plur[i])
        partes.append(texto)
    if len(partes) <= 1:
        return "".join(partes)
    return ", ".join(partes[:-1]) + " e " + partes[-1]


def valor_extenso(valor: float) -> str:
    valor = round(valor * 100) / 100
    reais = int(valor)
    centavos = round((valor - reais) * 100)
    if reais == 0 and centavos == 0:
        return "zero reais"
    texto = ""
    if reais > 0:
        texto += _inteiro_extenso(reais) + (" real" if reais == 1 else " reais")
    if centavos > 0:
        if reais > 0:
            texto += " e "
        texto += _inteiro_extenso(centavos) + (" centavo" if centavos == 1 else " centavos")
    return texto[0].upper() + texto[1:]


MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
          "agosto", "setembro", "outubro", "novembro", "dezembro"]


def data_extensa(iso: str) -> str:
    if not iso:
        return "—"
    y, m, d = (int(x) for x in iso.split("-"))
    return f"{d} de {MESES[m - 1]} de {y}"


def formata_brl(v: float) -> str:
    s = f"{v:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def parse_valor(texto: str) -> float:
    if not texto:
        return 0.0
    limpo = "".join(c for c in texto if c.isdigit() or c in ",.-")
    limpo = limpo.replace(".", "").replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------------------
# Persistencia (numero atual + historico de recibos)
# ---------------------------------------------------------------------------

class Armazenamento:
    def __init__(self, arquivo: Path):
        self.arquivo = arquivo
        self.arquivo.parent.mkdir(parents=True, exist_ok=True)
        self.dados = {"contador": 1, "recibos": {}}
        self._carregar()

    def _carregar(self):
        if self.arquivo.exists():
            try:
                with open(self.arquivo, "r", encoding="utf-8") as f:
                    self.dados = json.load(f)
            except Exception:
                pass
        self.dados.setdefault("contador", 1)
        self.dados.setdefault("recibos", {})

    def salvar(self):
        with open(self.arquivo, "w", encoding="utf-8") as f:
            json.dump(self.dados, f, ensure_ascii=False, indent=2)

    def proximo_numero(self) -> int:
        return int(self.dados.get("contador", 1))

    def registrar(self, registro: dict):
        numero = int(registro["numero"])
        self.dados["recibos"][str(numero)] = registro
        if numero >= int(self.dados.get("contador", 1)):
            self.dados["contador"] = numero + 1
        self.salvar()

    def excluir(self, numero: int):
        self.dados["recibos"].pop(str(numero), None)
        self.salvar()

    def listar(self):
        itens = list(self.dados["recibos"].values())
        itens.sort(key=lambda r: r["numero"], reverse=True)
        return itens


# ---------------------------------------------------------------------------
# Geracao do PDF
# ---------------------------------------------------------------------------

def gerar_pdf(caminho: Path, registro: dict):
    largura, altura = A4
    margem = 20 * mm
    c = canvas.Canvas(str(caminho), pagesize=A4)

    x0 = margem
    x1 = largura - margem
    y = altura - margem

    # ---- Cabecalho: logo + dados da empresa ----
    if LOGO_PATH.exists():
        img = ImageReader(str(LOGO_PATH))
        iw, ih = img.getSize()
        alvo_h = 14 * mm
        alvo_w = alvo_h * iw / ih
        c.drawImage(img, x0, y - alvo_h, width=alvo_w, height=alvo_h,
                    mask="auto")

    c.setFont("Helvetica-Bold", 11)
    c.setFillColor(NAVY)
    c.drawRightString(x1, y - 3 * mm, EMPRESA["nome"])
    c.setFont("Helvetica", 8.5)
    c.setFillColor(MUTED)
    linhas_empresa = [
        f"CNPJ {EMPRESA['cnpj']}",
        EMPRESA["endereco"],
        EMPRESA["cidade_cep"],
        EMPRESA["telefone"],
    ]
    yy = y - 7.5 * mm
    for linha in linhas_empresa:
        c.drawRightString(x1, yy, linha)
        yy -= 3.6 * mm

    y -= 21 * mm
    c.setStrokeColor(NAVY)
    c.setLineWidth(1.4)
    c.line(x0, y, x1, y)

    # ---- Caixa RECIBO / VALOR ----
    y -= 10 * mm
    caixa_h = 20 * mm
    caixa_top = y + 4 * mm
    caixa_bottom = caixa_top - caixa_h
    largura_direita = 40 * mm
    divisor_x = x1 - largura_direita

    c.setStrokeColor(NAVY)
    c.setLineWidth(1.1)
    c.rect(x0, caixa_bottom, x1 - x0, caixa_h, stroke=1, fill=0)
    c.line(divisor_x, caixa_bottom, divisor_x, caixa_top)

    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x0 + 6 * mm, caixa_top - 8 * mm, "RECIBO")
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(BLUE)
    c.drawString(x0 + 6 * mm, caixa_top - 13.5 * mm,
                 f"Nº {int(registro['numero']):04d}")

    c.setFillColor(HexColor("#f6f9fc"))
    c.rect(divisor_x, caixa_bottom, largura_direita, caixa_h, stroke=0, fill=1)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(divisor_x + largura_direita / 2, caixa_top - 6 * mm, "VALOR")
    c.setFillColor(NAVY)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(divisor_x + largura_direita / 2, caixa_top - 13 * mm,
                         formata_brl(registro["valor"]))

    y = caixa_bottom - 10 * mm

    # ---- Corpo do texto ----
    c.setFillColor(INK)
    c.setFont("Helvetica", 10.5)
    texto = (
        f"Recebi de {EMPRESA['nome']}, CNPJ {EMPRESA['cnpj']}, a importância de "
        f"{formata_brl(registro['valor'])} ({valor_extenso(registro['valor'])}), "
        f"referente a {registro.get('descricao') or '____________________'}."
    )
    y = _texto_justificado(c, texto, x0, y, x1 - x0, 13.5)

    # ---- Caixa Recebedor ----
    y -= 6 * mm
    caixa2_h = 14 * mm
    c.setFillColor(HexColor("#f6f9fc"))
    c.setStrokeColor(HexColor("#d8dee8"))
    c.roundRect(x0, y - caixa2_h, x1 - x0, caixa2_h, 3, stroke=1, fill=1)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9)
    doc_tipo = registro.get("tipo_documento", "CPF")
    c.drawString(x0 + 5 * mm, y - 5.5 * mm,
                 f"Recebedor: {registro.get('recebedor') or '—'}")
    c.drawString(x0 + 5 * mm, y - 10.5 * mm,
                 f"{doc_tipo}: {registro.get('documento') or '—'}   ·   "
                 f"Telefone: {registro.get('telefone') or '—'}")
    y -= caixa2_h + 6 * mm

    # ---- Forma de pagamento ----
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8.5)
    c.drawString(x0, y, f"Forma de pagamento: {registro.get('pagamento', '')}")
    y -= 18 * mm

    # ---- Rodape: local/data + assinatura ----
    c.setFillColor(INK)
    c.setFont("Helvetica", 9.5)
    c.drawString(x0, y, f"{EMPRESA['local']}, {data_extensa(registro.get('data'))}.")

    assinatura_w = 65 * mm
    ax = x1 - assinatura_w
    c.setStrokeColor(INK)
    c.setLineWidth(0.8)
    c.line(ax, y + 1 * mm, x1, y + 1 * mm)
    c.setFont("Helvetica", 9)
    c.drawCentredString(ax + assinatura_w / 2, y - 4 * mm,
                         registro.get("recebedor") or "—")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(ax + assinatura_w / 2, y - 8 * mm, "assinatura do recebedor")

    c.showPage()
    c.save()


def _texto_justificado(c: canvas.Canvas, texto: str, x: float, y: float,
                        largura: float, leading: float) -> float:
    """Quebra o texto em linhas simples (sem justificacao completa, apenas
    alinhado a esquerda) respeitando a largura disponivel."""
    palavras = texto.split()
    linha = ""
    for palavra in palavras:
        teste = (linha + " " + palavra).strip()
        if c.stringWidth(teste, "Helvetica", 10.5) > largura:
            c.drawString(x, y, linha)
            y -= leading
            linha = palavra
        else:
            linha = teste
    if linha:
        c.drawString(x, y, linha)
        y -= leading
    return y


# ---------------------------------------------------------------------------
# Interface grafica
# ---------------------------------------------------------------------------

class RecibosApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Emissão de Recibos — NetBicas FIBRA")
        self.geometry("980x640")
        self.minsize(860, 560)
        self.configure(bg="#eef1f6")

        self.storage = Armazenamento(DATA_FILE)

        self._montar_estilo()
        self._montar_layout()
        self._novo_recibo(mostrar_status=False)
        self._atualizar_historico()

    # -- estilo -------------------------------------------------------
    def _montar_estilo(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TLabel", background="#eef1f6", foreground="#1c2330",
                         font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI Semibold", 13),
                         foreground="#0a2d52")
        style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground="#64748b")
        style.configure("Section.TLabel", font=("Segoe UI Semibold", 10),
                         foreground="#0a2d52")
        style.configure("TButton", font=("Segoe UI Semibold", 10), padding=8)
        style.configure("Primary.TButton", background="#0a2d52", foreground="white")
        style.map("Primary.TButton", background=[("active", "#0d6fb8")])
        style.configure("Card.TFrame", background="white")
        style.configure("Bg.TFrame", background="#eef1f6")

    # -- layout ---------------------------------------------------------
    def _montar_layout(self):
        topo = ttk.Frame(self, style="Bg.TFrame", padding=(18, 14, 18, 6))
        topo.pack(fill="x")
        ttk.Label(topo, text="Emissão de recibos", style="Header.TLabel").pack(anchor="w")
        ttk.Label(topo, text="NetBicas FIBRA — pagamentos a terceiros, em segundos",
                   style="Sub.TLabel").pack(anchor="w")

        corpo = ttk.Frame(self, style="Bg.TFrame", padding=(18, 6, 18, 18))
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(0, weight=0, minsize=360)
        corpo.columnconfigure(1, weight=1)
        corpo.rowconfigure(0, weight=1)

        # ---- painel esquerdo (formulario) ----
        painel = ttk.Frame(corpo, style="Card.TFrame", padding=16)
        painel.grid(row=0, column=0, sticky="nsew", padx=(0, 14))

        ttk.Label(painel, text="DADOS DO PAGAMENTO", style="Section.TLabel",
                   background="white").grid(row=0, column=0, columnspan=2,
                                             sticky="w", pady=(0, 10))

        self.vars = {
            "numero": tk.StringVar(),
            "data": tk.StringVar(value=date.today().isoformat()),
            "recebedor": tk.StringVar(),
            "tipo_documento": tk.StringVar(value="CPF"),
            "documento": tk.StringVar(),
            "telefone": tk.StringVar(),
            "valor": tk.StringVar(),
            "pagamento": tk.StringVar(value="Dinheiro"),
        }

        r = 1
        r = self._campo(painel, r, "Nº do recibo", self.vars["numero"], width=10, col=0)
        self._campo(painel, r - 1, "Data (AAAA-MM-DD)", self.vars["data"], width=14, col=1)

        r = self._campo(painel, r, "Nome do recebedor", self.vars["recebedor"], colspan=2)

        frame_doc = ttk.Frame(painel, style="Card.TFrame")
        frame_doc.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        frame_doc.columnconfigure(1, weight=1)
        ttk.Label(frame_doc, text="Tipo", background="white").grid(row=0, column=0, sticky="w")
        combo_tipo = ttk.Combobox(frame_doc, values=["CPF", "CNPJ"], width=8,
                                    textvariable=self.vars["tipo_documento"], state="readonly")
        combo_tipo.grid(row=1, column=0, sticky="w", padx=(0, 8))
        ttk.Label(frame_doc, text="Número do documento", background="white").grid(row=0, column=1, sticky="w")
        ttk.Entry(frame_doc, textvariable=self.vars["documento"]).grid(row=1, column=1, sticky="ew")
        r += 1

        r = self._campo(painel, r, "Telefone do recebedor", self.vars["telefone"], colspan=2)
        r = self._campo(painel, r, "Valor (R$)", self.vars["valor"], colspan=2)

        ttk.Label(painel, text="Referente a", background="white").grid(
            row=r, column=0, columnspan=2, sticky="w")
        r += 1
        self.txt_descricao = tk.Text(painel, height=3, font=("Segoe UI", 10),
                                       wrap="word", relief="solid", borderwidth=1)
        self.txt_descricao.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(2, 10))
        r += 1

        ttk.Label(painel, text="Forma de pagamento", background="white").grid(
            row=r, column=0, columnspan=2, sticky="w")
        r += 1
        ttk.Combobox(painel, values=FORMAS_PAGAMENTO, state="readonly",
                      textvariable=self.vars["pagamento"]).grid(
            row=r, column=0, columnspan=2, sticky="ew", pady=(2, 14))
        r += 1

        btn_salvar = ttk.Button(painel, text="Salvar recibo (PDF)",
                                  style="Primary.TButton", command=self._salvar_recibo)
        btn_salvar.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        r += 1
        ttk.Button(painel, text="Limpar campos (novo recibo)",
                    command=self._novo_recibo).grid(row=r, column=0, columnspan=2,
                                                       sticky="ew")
        r += 1

        self.lbl_status = ttk.Label(painel, text="", background="white",
                                      foreground="#c0392b", wraplength=320)
        self.lbl_status.grid(row=r, column=0, columnspan=2, sticky="w", pady=(8, 0))
        r += 1

        ttk.Label(painel, text="ÚLTIMOS RECIBOS EMITIDOS", style="Section.TLabel",
                   background="white").grid(row=r, column=0, columnspan=2,
                                             sticky="w", pady=(16, 6))
        r += 1

        hist_frame = ttk.Frame(painel, style="Card.TFrame")
        hist_frame.grid(row=r, column=0, columnspan=2, sticky="nsew")
        painel.rowconfigure(r, weight=1)

        colunas = ("numero", "recebedor", "valor")
        self.tree = ttk.Treeview(hist_frame, columns=colunas, show="headings", height=8)
        self.tree.heading("numero", text="Nº")
        self.tree.heading("recebedor", text="Recebedor")
        self.tree.heading("valor", text="Valor")
        self.tree.column("numero", width=50, anchor="center")
        self.tree.column("recebedor", width=150)
        self.tree.column("valor", width=90, anchor="e")
        self.tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(hist_frame, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Double-1>", self._carregar_do_historico)

        botoes_hist = ttk.Frame(painel, style="Card.TFrame")
        botoes_hist.grid(row=r + 1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(botoes_hist, text="Carregar / reimprimir",
                    command=self._carregar_do_historico).pack(side="left")
        ttk.Button(botoes_hist, text="Excluir do histórico",
                    command=self._excluir_do_historico).pack(side="left", padx=6)

        # ---- painel direito (pre-visualizacao textual) ----
        preview_wrap = ttk.Frame(corpo, style="Card.TFrame", padding=20)
        preview_wrap.grid(row=0, column=1, sticky="nsew")

        ttk.Label(preview_wrap, text="Pré-visualização", style="Section.TLabel",
                   background="white").pack(anchor="w", pady=(0, 8))

        self.preview_texto = tk.Text(preview_wrap, wrap="word", font=("Georgia", 11),
                                       relief="flat", background="#fbfcfe",
                                       height=28, padx=16, pady=16, state="disabled")
        self.preview_texto.pack(fill="both", expand=True)

        for var in self.vars.values():
            var.trace_add("write", lambda *a: self._atualizar_preview())
        self.txt_descricao.bind("<KeyRelease>", lambda e: self._atualizar_preview())

    def _campo(self, parent, row, label, var, width=None, col=0, colspan=1):
        if colspan == 2:
            ttk.Label(parent, text=label, background="white").grid(
                row=row, column=0, columnspan=2, sticky="w")
            entry = ttk.Entry(parent, textvariable=var, width=width)
            entry.grid(row=row + 1, column=0, columnspan=2, sticky="ew", pady=(2, 10))
            return row + 2
        else:
            ttk.Label(parent, text=label, background="white").grid(
                row=row, column=col, sticky="w")
            entry = ttk.Entry(parent, textvariable=var, width=width)
            entry.grid(row=row + 1, column=col, sticky="ew", padx=(0, 8) if col == 0 else 0,
                        pady=(2, 10))
            return row + 2

    # -- logica -----------------------------------------------------------
    def _registro_atual(self) -> dict:
        return {
            "numero": self.vars["numero"].get().strip(),
            "data": self.vars["data"].get().strip(),
            "recebedor": self.vars["recebedor"].get().strip(),
            "tipo_documento": self.vars["tipo_documento"].get(),
            "documento": self.vars["documento"].get().strip(),
            "telefone": self.vars["telefone"].get().strip(),
            "valor": parse_valor(self.vars["valor"].get()),
            "descricao": self.txt_descricao.get("1.0", "end").strip(),
            "pagamento": self.vars["pagamento"].get(),
        }

    def _atualizar_preview(self):
        try:
            reg = self._registro_atual()
        except Exception:
            return
        numero = reg["numero"] or "0001"
        try:
            numero_fmt = f"{int(numero):04d}"
        except ValueError:
            numero_fmt = numero
        linhas = []
        linhas.append(f"{EMPRESA['nome']}")
        linhas.append(f"CNPJ {EMPRESA['cnpj']} — {EMPRESA['endereco']}")
        linhas.append(f"{EMPRESA['cidade_cep']} — {EMPRESA['telefone']}")
        linhas.append("")
        linhas.append(f"RECIBO Nº {numero_fmt}                         "
                       f"Valor: {formata_brl(reg['valor'])}")
        linhas.append("-" * 70)
        linhas.append("")
        linhas.append(
            f"Recebi de {EMPRESA['nome']}, CNPJ {EMPRESA['cnpj']}, a importância de "
            f"{formata_brl(reg['valor'])} "
            f"({valor_extenso(reg['valor']) if reg['valor'] > 0 else '—'}), referente a "
            f"{reg['descricao'] or '_____________________'}."
        )
        linhas.append("")
        linhas.append(f"Recebedor: {reg['recebedor'] or '—'}")
        linhas.append(f"{reg['tipo_documento']}: {reg['documento'] or '—'}    "
                       f"Telefone: {reg['telefone'] or '—'}")
        linhas.append("")
        linhas.append(f"Forma de pagamento: {reg['pagamento']}")
        linhas.append("")
        linhas.append(f"{EMPRESA['local']}, {data_extensa(reg['data'])}.")
        linhas.append("")
        linhas.append("")
        linhas.append("_____________________________")
        linhas.append(f"{reg['recebedor'] or '—'}")
        linhas.append("assinatura do recebedor")

        self.preview_texto.configure(state="normal")
        self.preview_texto.delete("1.0", "end")
        self.preview_texto.insert("1.0", "\n".join(linhas))
        self.preview_texto.configure(state="disabled")

    def _novo_recibo(self, mostrar_status=True):
        self.vars["numero"].set(str(self.storage.proximo_numero()))
        self.vars["data"].set(date.today().isoformat())
        self.vars["recebedor"].set("")
        self.vars["tipo_documento"].set("CPF")
        self.vars["documento"].set("")
        self.vars["telefone"].set("")
        self.vars["valor"].set("")
        self.txt_descricao.delete("1.0", "end")
        self.vars["pagamento"].set("Dinheiro")
        self.lbl_status.configure(text="")
        self._atualizar_preview()

    def _validar(self, reg: dict) -> str:
        if not reg["numero"] or not reg["numero"].isdigit() or int(reg["numero"]) < 1:
            return "Informe um número de recibo válido."
        if not reg["recebedor"]:
            return "Informe o nome do recebedor."
        if not reg["valor"] or reg["valor"] <= 0:
            return "Informe um valor válido."
        return ""

    def _salvar_recibo(self):
        reg = self._registro_atual()
        erro = self._validar(reg)
        if erro:
            self.lbl_status.configure(text=erro, foreground="#c0392b")
            return
        reg["numero"] = int(reg["numero"])

        DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        nome_sugerido = f"Recibo {reg['numero']:04d} - {reg['recebedor']}.pdf"
        nome_sugerido = "".join(c for c in nome_sugerido if c not in '\\/:*?"<>|')
        caminho = filedialog.asksaveasfilename(
            title="Salvar recibo como...",
            initialdir=str(DEFAULT_OUTPUT_DIR),
            initialfile=nome_sugerido,
            defaultextension=".pdf",
            filetypes=[("Arquivo PDF", "*.pdf")],
        )
        if not caminho:
            return

        try:
            gerar_pdf(Path(caminho), reg)
        except Exception as exc:
            messagebox.showerror("Erro ao gerar PDF", str(exc))
            return

        self.storage.registrar(reg)
        self._atualizar_historico()
        self.lbl_status.configure(
            text=f"Recibo nº {reg['numero']:04d} salvo com sucesso.",
            foreground="#0a2d52")

        if messagebox.askyesno("Recibo salvo",
                                 "Recibo salvo com sucesso!\n\nDeseja abrir o PDF agora?"):
            self._abrir_arquivo(caminho)

        self._novo_recibo(mostrar_status=False)

    def _abrir_arquivo(self, caminho: str):
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f'open "{caminho}"')
            else:
                webbrowser.open(caminho)
        except Exception:
            pass

    def _atualizar_historico(self):
        self.tree.delete(*self.tree.get_children())
        for item in self.storage.listar():
            self.tree.insert("", "end", iid=str(item["numero"]), values=(
                f"{item['numero']:04d}",
                item.get("recebedor", ""),
                formata_brl(item.get("valor", 0)),
            ))

    def _carregar_do_historico(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        numero = sel[0]
        reg = self.storage.dados["recibos"].get(numero)
        if not reg:
            return
        self.vars["numero"].set(str(reg["numero"]))
        self.vars["data"].set(reg.get("data", ""))
        self.vars["recebedor"].set(reg.get("recebedor", ""))
        self.vars["tipo_documento"].set(reg.get("tipo_documento", "CPF"))
        self.vars["documento"].set(reg.get("documento", ""))
        self.vars["telefone"].set(reg.get("telefone", ""))
        self.vars["valor"].set(str(reg.get("valor", "")))
        self.txt_descricao.delete("1.0", "end")
        self.txt_descricao.insert("1.0", reg.get("descricao", ""))
        self.vars["pagamento"].set(reg.get("pagamento", "Dinheiro"))
        self._atualizar_preview()
        self.lbl_status.configure(
            text=f"Recibo nº {int(numero):04d} carregado. Você pode reimprimir "
                  f"clicando em 'Salvar recibo (PDF)'.",
            foreground="#0a2d52")

    def _excluir_do_historico(self):
        sel = self.tree.selection()
        if not sel:
            return
        numero = int(sel[0])
        if messagebox.askyesno("Excluir recibo",
                                 f"Remover o recibo nº {numero:04d} do histórico?\n"
                                 "(o PDF já salvo no computador não será apagado)"):
            self.storage.excluir(numero)
            self._atualizar_historico()


if __name__ == "__main__":
    app = RecibosApp()
    app.mainloop()
