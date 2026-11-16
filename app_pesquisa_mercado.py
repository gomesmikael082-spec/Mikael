import os
import io
import re
import json
import math
import queue
import logging
import threading
import copy
import statistics
from datetime import datetime
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pesquisa_mercado")
os.makedirs(CONFIG_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config_variaveis.json")

# --- ORDEM DO SCHEMA OPENXML (WORD) ---
SUC_TCPR  = ["w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark"]
SUC_TBLPR = ["w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"]
SUC_CMAR  = ["w:tblLook"]
SUC_PPR   = ["w:shd", "w:tabs", "w:suppressAutoHyphens", "w:spacing", "w:ind", "w:jc", "w:rPr"]


def inserir_ordenado(pai, el, sucessores):
    """Insere elemento XML respeitando a sequência estrita de tags do schema WordprocessingML."""
    for tag in sucessores:
        achado = pai.find(qn(tag))
        if achado is not None:
            achado.addprevious(el)
            return
    pai.append(el)


# --- TABELA 02: NOTA AGRONÔMICA (Adaptado de Kozma 1984 x IN RFB 1877/2019) ---
TABELA_NOTA_AGRONOMICA = {
    "I":   {"Ótima": "1,00", "Muito Boa": "0,95", "Boa": "0,90", "Desfavorável": "0,80", "Má": "0,75", "Péssima": "0,70"},
    "II":  {"Ótima": "0,80", "Muito Boa": "0,76", "Boa": "0,72", "Desfavorável": "0,64", "Má": "0,60", "Péssima": "0,56"},
    "III": {"Ótima": "0,64", "Muito Boa": "0,61", "Boa": "0,58", "Desfavorável": "0,51", "Má": "0,48", "Péssima": "0,45"},
    "IV":  {"Ótima": "0,51", "Muito Boa": "0,48", "Boa": "0,46", "Desfavorável": "0,41", "Má": "0,38", "Péssima": "0,36"},
    "V":   {"Ótima": "0,41", "Muito Boa": "0,39", "Boa": "0,37", "Desfavorável": "0,33", "Má": "0,31", "Péssima": "0,29"},
    "VI":  {"Ótima": "0,20", "Muito Boa": "0,19", "Boa": "0,18", "Desfavorável": "0,16", "Má": "0,15", "Péssima": "0,14"},
}

OPCOES_APTIDAO = [
    "I - Lavoura (Aptidão Boa)",
    "II - Lavoura (Aptidão Regular)",
    "III - Lavoura (Aptidão Restrita)",
    "IV - Pastagem Plantada",
    "V - Silvicultura ou Pastagem Natural",
    "VI - Preservação de Fauna/Flora"
]

OPCOES_SITUACAO = [
    "Ótima",
    "Muito Boa",
    "Boa",
    "Desfavorável",
    "Má",
    "Péssima"
]

# --- PERFIS PRÉ-CADASTRADOS DE VARIÁVEIS ---
DEFAULT_VARS_PADRAO = [
    {"nome": "Via", "tipo": "codigo", "opcoes": ["1 - Local", "2 - Coletora", "3 - Arterial"]},
    {"nome": "Uso", "tipo": "codigo", "opcoes": ["1 - Residencial", "2 - Comercial", "3 - Misto"]},
    {"nome": "Testada (m)", "tipo": "numero", "opcoes": []},
    {"nome": "PGV (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "PIB (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "Condição", "tipo": "codigo", "opcoes": ["1 - Oferta", "2 - Transação"]}
]

DEFAULT_VARS_RURAL = [
    {"nome": "Aptidão", "tipo": "codigo", "opcoes": OPCOES_APTIDAO},
    {"nome": "Situação", "tipo": "codigo", "opcoes": OPCOES_SITUACAO},
    {"nome": "Nota Agronômica", "tipo": "numero", "opcoes": []},
    {"nome": "Benfeitoria", "tipo": "codigo", "opcoes": ["1 - Sem Benfeitorias", "2 - Com Benfeitorias", "3 - Produtiva"]}
]

DEFAULT_VARS_COPASA = [
    {"nome": "Frente", "tipo": "texto", "opcoes": ["Não informado"]},
    {"nome": "Via de acesso", "tipo": "codigo", "opcoes": ["pavimentada", "não pavimentada"]},
    {"nome": "Informações", "tipo": "texto", "opcoes": []}
]

CATALOGO_VARIAVEIS_USUAIS = [
    {"nome": "Via", "tipo": "codigo", "opcoes": ["1 - Local", "2 - Coletora", "3 - Arterial"]},
    {"nome": "Uso", "tipo": "codigo", "opcoes": ["1 - Residencial", "2 - Comercial", "3 - Misto"]},
    {"nome": "Testada (m)", "tipo": "numero", "opcoes": []},
    {"nome": "PGV (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "PIB (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "Condição", "tipo": "codigo", "opcoes": ["1 - Oferta", "2 - Transação"]},
    {"nome": "Distância ao Centro (m)", "tipo": "numero", "opcoes": []},
    {"nome": "Topografia", "tipo": "codigo", "opcoes": ["1 - Plano", "2 - Aclive/Declive", "3 - Irregular"]},
    {"nome": "Padrão Construtivo", "tipo": "codigo", "opcoes": ["1 - Baixo", "2 - Normal", "3 - Alto"]},
    {"nome": "Aptidão", "tipo": "codigo", "opcoes": OPCOES_APTIDAO},
    {"nome": "Situação", "tipo": "codigo", "opcoes": OPCOES_SITUACAO},
    {"nome": "Nota Agronômica", "tipo": "numero", "opcoes": []},
    {"nome": "Benfeitoria", "tipo": "codigo", "opcoes": ["1 - Sem Benfeitorias", "2 - Com Benfeitorias", "3 - Produtiva"]},
    {"nome": "Frente", "tipo": "texto", "opcoes": ["Não informado"]},
    {"nome": "Via de acesso", "tipo": "codigo", "opcoes": ["pavimentada", "não pavimentada"]},
    {"nome": "Informações", "tipo": "texto", "opcoes": []}
]


def data_hoje_br():
    return datetime.now().strftime("%d/%m/%Y")


def normalizar_chave_var(nome):
    """Normaliza nome de coluna/variável para comparação exata."""
    s = str(nome).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def calcular_nota_agronomica(aptidao_str, situacao_str):
    """Cruza automaticamente a Aptidão (I a VI) e a Situação/Acesso conforme a Tabela 02."""
    if not aptidao_str or not situacao_str:
        return ""
    m = re.match(r"^\s*(VI|IV|V|III|II|I)\b", str(aptidao_str).strip(), re.IGNORECASE)
    if not m:
        return ""
    romano = m.group(1).upper()
    sit = str(situacao_str).strip().lower()
    mapa_sit = TABELA_NOTA_AGRONOMICA.get(romano, {})
    for k, v in mapa_sit.items():
        if k.lower() == sit:
            return v
    for k in sorted(mapa_sit.keys(), key=len, reverse=True):
        if k.lower() in sit:
            return mapa_sit[k]
    return ""


def converter_latlon_para_utm(texto_coord):
    """
    Extrai Latitude e Longitude (graus decimais, referência SIRGAS 2000 / WGS84)
    e converte para Zona UTM, Coord. E (m) e Coord. S (m).
    """
    if not texto_coord:
        return None
    nums = re.findall(r"[-+]?\d+\.\d+", str(texto_coord))
    if len(nums) < 2:
        return None
    lat, lon = float(nums[0]), float(nums[1])
    if not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return None

    a = 6378137.0
    f = 1 / 298.257222101
    k0 = 0.9996
    e2 = 2 * f - f * f
    ep2 = e2 / (1 - e2)

    zona_num = int((lon + 180) / 6) + 1
    lon0 = math.radians((zona_num - 1) * 6 - 180 + 3)
    lat_rad = math.radians(lat)
    lon_rad = math.radians(lon)

    N = a / math.sqrt(1 - e2 * math.sin(lat_rad) ** 2)
    T = math.tan(lat_rad) ** 2
    C = ep2 * math.cos(lat_rad) ** 2
    A = math.cos(lat_rad) * (lon_rad - lon0)

    M = a * (
        (1 - e2 / 4 - 3 * e2 ** 2 / 64 - 5 * e2 ** 3 / 256) * lat_rad
        - (3 * e2 / 8 + 3 * e2 ** 2 / 32 + 45 * e2 ** 3 / 1024) * math.sin(2 * lat_rad)
        + (15 * e2 ** 2 / 256 + 45 * e2 ** 3 / 1024) * math.sin(4 * lat_rad)
        - (35 * e2 ** 3 / 3072) * math.sin(6 * lat_rad)
    )

    easting = k0 * N * (
        A + (1 - T + C) * A ** 3 / 6
        + (5 - 18 * T + T ** 2 + 72 * C - 58 * ep2) * A ** 5 / 120
    ) + 500000.0

    northing = k0 * (
        M + N * math.tan(lat_rad) * (
            A ** 2 / 2 + (5 - T + 9 * C + 4 * C ** 2) * A ** 4 / 24
            + (61 - 58 * T + T ** 2 + 600 * C - 330 * ep2) * A ** 6 / 720
        )
    )
    if lat < 0:
        northing += 10000000.0

    bandas = "CDEFGHJKLMNPQRSTUVWX"
    idx_banda = int((lat + 80) / 8)
    letra_zona = bandas[max(0, min(len(bandas) - 1, idx_banda))] if -80 <= lat <= 84 else ("S" if lat < 0 else "N")

    return f"{zona_num} {letra_zona}", f"{easting:.2f}", f"{northing:.2f}"


def converter_para_float(texto):
    """Converte strings ou números para float sem corromper decimais já numéricos."""
    if texto is None:
        return 0.0
    if isinstance(texto, (int, float)):
        val_f = float(texto)
        return 0.0 if math.isnan(val_f) else val_f
    s = str(texto).replace("R$", "").replace("m²", "").replace("ha", "").strip()
    if not s or s.lower() == "nan":
        return 0.0
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        partes = s.split(".")
        if len(partes) > 2:
            s = s.replace(".", "")
        elif len(partes) == 2 and len(partes[1]) == 3 and len(partes[0]) <= 3:
            s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def formatar_moeda_br(valor):
    try:
        return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def formatar_numero_br(valor, casas=2):
    try:
        fmt = f"{{:,.{casas}f}}"
        return fmt.format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def limpar_sufixo_coord(coord_str):
    s = str(coord_str).strip()
    s = re.sub(r'\s*m\s*[ESes]\s*$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s*[ESes]\s*$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s*m\s*$', '', s, flags=re.IGNORECASE)
    return s.strip()


def normalizar_url_gdrive(url):
    if not url:
        return ""
    url = url.strip()
    match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url) or re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if match:
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
    return url


def normalizar_dado(dado):
    d_id = dado.get("dado_id") if dado.get("dado_id") is not None else dado.get("D.", 1)
    utilizado = dado.get("utilizado") if "utilizado" in dado else dado.get("Utilizado", True)
    if isinstance(utilizado, str):
        utilizado = utilizado.strip().lower() not in ("false", "0", "não", "nao", "☐")

    informante = dado.get("informante") or dado.get("Informante") or ""
    telefone = dado.get("telefone") or dado.get("Telefone") or ""
    endereco = dado.get("endereco") or dado.get("Endereço") or ""
    bairro = dado.get("bairro") or dado.get("Bairro") or ""
    municipio = dado.get("municipio") or dado.get("Município") or ""
    
    unidade = dado.get("unidade") or dado.get("Unidade") or "m²"
    
    valor_total = dado.get("valor_total")
    if valor_total is None:
        valor_total = dado.get("Valor Total (R$)", 0.0)
    valor_total = converter_para_float(valor_total)
    
    area_terreno = dado.get("area_terreno")
    if area_terreno is None:
        area_terreno = dado.get(f"Área Terreno ({unidade})") or dado.get("Área Terreno (ha)") or dado.get("Área Terreno (m²)") or 0.0
    area_terreno = converter_para_float(area_terreno)
    
    area_construida = dado.get("area_construida")
    if area_construida is None:
        area_construida = dado.get("Área Construída (m²)", 0.0)
    area_construida = converter_para_float(area_construida)
    
    unitario = dado.get("unitario")
    if unitario is None:
        unitario = dado.get(f"Unitário (R$/{unidade})") or dado.get("Unitário (R$/ha)") or dado.get("Unitário (R$/m²)")
        if unitario is None and area_terreno > 0:
            unitario = valor_total / area_terreno
        else:
            unitario = converter_para_float(unitario)
    else:
        unitario = converter_para_float(unitario)

    zona_utm = dado.get("zona_utm") or dado.get("Zona UTM") or ""
    coord_e = dado.get("coord_e") or dado.get("Coord. E (m)") or ""
    coord_s = dado.get("coord_s") or dado.get("Coord. S (m)") or ""
    data = dado.get("data") or dado.get("Data") or data_hoje_br()
    link = dado.get("link") or dado.get("Link") or ""
    foto1 = dado.get("foto1") or dado.get("Foto1") or ""
    foto2 = dado.get("foto2") or dado.get("Foto2") or ""
    localizacao = dado.get("localizacao") or dado.get("Localização") or ("Rural" if unidade == "ha" else "Urbana")

    vars_extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras") or {}

    return {
        "dado_id": int(d_id),
        "utilizado": bool(utilizado),
        "informante": str(informante).strip(),
        "telefone": str(telefone).strip(),
        "endereco": str(endereco).strip(),
        "bairro": str(bairro).strip(),
        "municipio": str(municipio).strip(),
        "valor_total": float(valor_total),
        "area_terreno": float(area_terreno),
        "area_construida": float(area_construida),
        "unitario": float(unitario) if unitario else 0.0,
        "unidade": str(unidade).strip(),
        "localizacao": str(localizacao).strip(),
        "zona_utm": str(zona_utm).strip(),
        "coord_e": limpar_sufixo_coord(str(coord_e)),
        "coord_s": limpar_sufixo_coord(str(coord_s)),
        "data": str(data).strip(),
        "link": str(link).strip(),
        "foto1": str(foto1).strip(),
        "foto2": str(foto2).strip(),
        "variaveis_extras": vars_extras
    }


class AppPesquisaMercado:
    def __init__(self, root):
        self.root = root
        self.root.title("ENPROL - Sistema de Pesquisa de Mercado v6.1 (Padrão SISDEA Pro)")
        self.root.geometry("1220x860")
        self.root.minsize(1050, 740)

        self._configurar_estilos()

        self.modelo_ativo = "PADRAO"
        self.dados_pesquisas = []
        self.item_em_edicao = None
        self.variaveis_config = self._carregar_config_variaveis()
        self.celula_ativa = {"dado_id": None, "campo": None, "rotulo": ""}
        self.fila_eventos_ui = queue.Queue()

        self.container_principal = ttk.Frame(self.root)
        self.container_principal.pack(fill="both", expand=True)

        self.root.after(100, self._processar_fila_ui)
        self._exibir_tela_hub()

    def _processar_fila_ui(self):
        """Processa eventos enviados por threads secundárias de forma thread-safe na main thread."""
        try:
            while True:
                acao, payload = self.fila_eventos_ui.get_nowait()
                if acao == "progresso":
                    self._atualizar_progresso(payload)
                elif acao == "concluido":
                    self._concluir_exportacao()
                elif acao == "erro":
                    self._falha_exportacao(payload)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._processar_fila_ui)

    def _configurar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        self.root.configure(bg="#f4f6f9")
        style.configure("TFrame", background="#f4f6f9")
        style.configure("TLabelframe", background="#f4f6f9", font=("Segoe UI", 9, "bold"))
        style.configure("TLabelframe.Label", background="#f4f6f9", foreground="#1f3c5b", font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background="#f4f6f9", font=("Segoe UI", 9))
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#e1e7ef", foreground="#1f3c5b")

        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background="#1f3c5b", foreground="white")
        style.map("Primary.TButton", background=[("active", "#2c5480")])

        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), background="#2b7a78", foreground="white")
        style.map("Accent.TButton", background=[("active", "#3a9895")])

        style.configure("Hub.TButton", font=("Segoe UI", 11, "bold"), padding=10, background="#1f3c5b", foreground="white")
        style.map("Hub.TButton", background=[("active", "#2d5784")])

    def _limpar_container(self):
        for widget in self.container_principal.winfo_children():
            widget.destroy()
        for attr in ("tree", "frame_vars", "lbl_estatisticas", "progress_bar"):
            if hasattr(self, attr):
                delattr(self, attr)

    def _carregar_config_variaveis(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.warning(f"Erro ao carregar configurações: {e}")
        return copy.deepcopy(DEFAULT_VARS_PADRAO)

    def _salvar_config_variaveis(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.variaveis_config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("Erro ao salvar variáveis", str(e))

    # --- HUB INICIAL ---
    def _exibir_tela_hub(self):
        self._limpar_container()
        self.root.config(menu="")

        frame_hub = ttk.Frame(self.container_principal, padding=40)
        frame_hub.place(relx=0.5, rely=0.5, anchor="center")

        lbl_logo = ttk.Label(frame_hub, text="SISTEMA DE PESQUISA DE MERCADO v6.1", font=("Segoe UI", 18, "bold"), foreground="#1f3c5b")
        lbl_logo.pack(pady=(0, 5))

        lbl_sub = ttk.Label(frame_hub, text="Integração Direta SISDEA • Automação Agronômica • Relatórios Word", font=("Segoe UI", 10), foreground="#555555")
        lbl_sub.pack(pady=(0, 30))

        btn_novo = ttk.Button(frame_hub, text="📄  Criar Novo Arquivo / Projeto", style="Hub.TButton", width=40, command=self._popup_escolha_modelo)
        btn_novo.pack(pady=7)

        btn_abrir = ttk.Button(frame_hub, text="📂  Carregar Projeto Existente (.json)", style="Hub.TButton", width=40, command=self._abrir_projeto_hub)
        btn_abrir.pack(pady=7)

        btn_importar = ttk.Button(frame_hub, text="📥  Importar Dados de Planilha (.xlsx)", style="Hub.TButton", width=40, command=self._importar_planilha_hub)
        btn_importar.pack(pady=7)

        btn_sair = ttk.Button(frame_hub, text="🚪  Sair do Sistema", width=40, command=self.root.quit)
        btn_sair.pack(pady=(20, 0))

    def _popup_escolha_modelo(self, callback_pos_confirmacao=None):
        janela = tk.Toplevel(self.root)
        janela.title("Selecionar Modelo e Perfil Inicial")
        janela.geometry("480x310")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Selecione o modelo de ficha e perfil de variáveis:", font=("Segoe UI", 11, "bold")).pack(pady=(18, 12))

        var_mod = tk.StringVar(value="PADRAO_URBANO")
        ttk.Radiobutton(janela, text="Modelo 1 - Padrão Urbano (Via, Uso, Testada, PGV, PIB)", value="PADRAO_URBANO", variable=var_mod).pack(anchor="w", padx=35, pady=5)
        ttk.Radiobutton(janela, text="Modelo 1 - Padrão Rural (Aptidão, Situação, Nota Agronômica, Benfeitoria)", value="PADRAO_RURAL", variable=var_mod).pack(anchor="w", padx=35, pady=5)
        ttk.Radiobutton(janela, text="Modelo 2 - Contrato COPASA (Ficha Técnica Oficial)", value="COPASA", variable=var_mod).pack(anchor="w", padx=35, pady=5)

        def confirmar():
            escolha = var_mod.get()
            if escolha == "COPASA":
                self.modelo_ativo = "COPASA"
                self.variaveis_config = copy.deepcopy(DEFAULT_VARS_COPASA)
                self._unidade_padrao_inicial = "m²"
            elif escolha == "PADRAO_RURAL":
                self.modelo_ativo = "PADRAO"
                self.variaveis_config = copy.deepcopy(DEFAULT_VARS_RURAL)
                self._unidade_padrao_inicial = "ha"
            else:
                self.modelo_ativo = "PADRAO"
                self.variaveis_config = copy.deepcopy(DEFAULT_VARS_PADRAO)
                self._unidade_padrao_inicial = "m²"

            self._salvar_config_variaveis()
            janela.destroy()
            self.dados_pesquisas = []
            self._exibir_tela_trabalho()
            if callback_pos_confirmacao:
                callback_pos_confirmacao()

        ttk.Button(janela, text="Avançar", style="Primary.TButton", command=confirmar).pack(pady=(22, 0))

    def _abrir_projeto_hub(self):
        if self._executar_abertura_json():
            self._exibir_tela_trabalho()

    def _importar_planilha_hub(self):
        self._popup_escolha_modelo(callback_pos_confirmacao=self._importar_planilha_excel)

    # --- TELA DE TRABALHO ---
    def _exibir_tela_trabalho(self):
        self._limpar_container()
        self._criar_menu_superior()
        self._criar_layout_operacional()
        self._atualizar_interface_variaveis()
        self._configurar_colunas_grid()
        self._recarregar_grid()

    def _criar_menu_superior(self):
        menubar = tk.Menu(self.root)
        menu_arquivo = tk.Menu(menubar, tearoff=0)
        menu_arquivo.add_command(label="🏠 Voltar ao Hub Inicial", command=self._exibir_tela_hub)
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Novo Projeto...", command=self._popup_escolha_modelo)
        menu_arquivo.add_command(label="Abrir Projeto (.json)...", command=self._abrir_projeto_menu)
        menu_arquivo.add_command(label="Salvar Projeto (.json)", command=self._salvar_projeto)
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="📥 Importar Planilha Excel...", command=self._importar_planilha_excel)
        menu_arquivo.add_command(label="📋 Colar Linhas do Excel (Ctrl+V)", command=self._colar_dados_clipboard)
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Sair", command=self.root.quit)
        menubar.add_cascade(label="Arquivo", menu=menu_arquivo)

        menu_config = tk.Menu(menubar, tearoff=0)
        menu_config.add_command(label="⚙ Gerenciar Variáveis e Perfis...", command=self._janela_config_variaveis)
        menu_config.add_separator()
        menu_config.add_command(label="Aplicar Perfil Urbano Padrão", command=lambda: self._aplicar_perfil_rapido("URBANO"))
        menu_config.add_command(label="Aplicar Perfil Rural (Nota Agronômica)", command=lambda: self._aplicar_perfil_rapido("RURAL"))
        menu_config.add_command(label="Aplicar Perfil COPASA", command=lambda: self._aplicar_perfil_rapido("COPASA"))
        menu_config.add_separator()
        menu_config.add_command(label="📊 Ver Tabela 02 - Nota Agronômica (Referência)", command=self._mostrar_tabela_nota_agronomica)
        menubar.add_cascade(label="Variáveis e Perfis", menu=menu_config)

        self.root.config(menu=menubar)

    def _aplicar_perfil_rapido(self, tipo):
        if tipo == "URBANO":
            self.variaveis_config = copy.deepcopy(DEFAULT_VARS_PADRAO)
            self.var_unidade.set("m²")
        elif tipo == "RURAL":
            self.variaveis_config = copy.deepcopy(DEFAULT_VARS_RURAL)
            self.var_unidade.set("ha")
        elif tipo == "COPASA":
            self.modelo_ativo = "COPASA"
            self.variaveis_config = copy.deepcopy(DEFAULT_VARS_COPASA)
            self.var_unidade.set("m²")
        self._salvar_config_variaveis()
        self._atualizar_interface_variaveis()
        self._configurar_colunas_grid()
        self._recarregar_grid()

    def _criar_layout_operacional(self):
        frame_topo_info = ttk.Frame(self.container_principal)
        frame_topo_info.pack(fill="x", padx=12, pady=(4, 0))

        mod_label = "Modelo 1: Padrão / Memória de Cálculo" if self.modelo_ativo == "PADRAO" else "Modelo 2: Contrato COPASA"
        ttk.Label(frame_topo_info, text=f"Modo Ativo: {mod_label}", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b").pack(side="left")

        ttk.Button(frame_topo_info, text="🌐 Lat/Long ou Link Maps ➔ UTM", command=self._janela_conversor_utm).pack(side="right", padx=4)
        ttk.Button(frame_topo_info, text="📊 Tabela Nota Agronômica", command=self._mostrar_tabela_nota_agronomica).pack(side="right", padx=4)

        frame_form = ttk.LabelFrame(self.container_principal, text=" Cadastro e Edição do Dado de Mercado ", padding=6)
        frame_form.pack(fill="x", padx=10, pady=3)

        ttk.Label(frame_form, text="Informante:").grid(row=0, column=0, sticky="w")
        self.txt_informante = ttk.Entry(frame_form, width=22)
        self.txt_informante.grid(row=0, column=1, padx=4, pady=2, sticky="w")

        ttk.Label(frame_form, text="Telefone:").grid(row=0, column=2, sticky="w")
        self.txt_telefone = ttk.Entry(frame_form, width=20)
        self.txt_telefone.grid(row=0, column=3, padx=4, pady=2, sticky="w")

        ttk.Label(frame_form, text="Logradouro/Endereço:").grid(row=0, column=4, sticky="w")
        self.txt_endereco = ttk.Entry(frame_form, width=28)
        self.txt_endereco.grid(row=0, column=5, padx=4, pady=2, sticky="w")

        ttk.Label(frame_form, text="Bairro:").grid(row=1, column=0, sticky="w")
        self.txt_bairro = ttk.Entry(frame_form, width=22)
        self.txt_bairro.grid(row=1, column=1, padx=4, pady=2, sticky="w")

        ttk.Label(frame_form, text="Município:").grid(row=1, column=2, sticky="w")
        self.txt_municipio = ttk.Entry(frame_form, width=20)
        self.txt_municipio.grid(row=1, column=3, padx=4, pady=2, sticky="w")

        ttk.Label(frame_form, text="Valor Total (R$):").grid(row=1, column=4, sticky="w")
        self.txt_valor = ttk.Entry(frame_form, width=20)
        self.txt_valor.grid(row=1, column=5, padx=4, pady=2, sticky="w")

        ttk.Label(frame_form, text="Área Terreno:").grid(row=2, column=0, sticky="w")
        frame_area = ttk.Frame(frame_form)
        frame_area.grid(row=2, column=1, sticky="w", padx=4, pady=2)
        self.txt_area = ttk.Entry(frame_area, width=13)
        self.txt_area.pack(side="left")
        un_ini = getattr(self, "_unidade_padrao_inicial", "ha" if self.modelo_ativo == "PADRAO" else "m²")
        self.var_unidade = tk.StringVar(value=un_ini)
        self.cb_unidade = ttk.Combobox(frame_area, textvariable=self.var_unidade, values=["m²", "ha"], width=5, state="readonly")
        self.cb_unidade.pack(side="left", padx=2)

        ttk.Label(frame_form, text="Área Const. (m²):").grid(row=2, column=2, sticky="w")
        self.txt_area_const = ttk.Entry(frame_form, width=20)
        self.txt_area_const.grid(row=2, column=3, padx=4, pady=2, sticky="w")

        self.var_utilizado = tk.BooleanVar(value=True)
        self.chk_utilizado = ttk.Checkbutton(
            frame_form,
            text="☑ Utilizado no Modelo (Desmarcar = 'Não utilizado no modelo')",
            variable=self.var_utilizado
        )
        self.chk_utilizado.grid(row=2, column=4, columnspan=2, sticky="w", padx=4, pady=2)

        frame_coords = ttk.Frame(frame_form)
        frame_coords.grid(row=3, column=0, columnspan=6, sticky="w", pady=2)

        ttk.Label(frame_coords, text="Zona UTM:", foreground="#b22222", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 2))
        self.txt_zona = ttk.Entry(frame_coords, width=7)
        self.txt_zona.pack(side="left", padx=(0, 8))

        ttk.Label(frame_coords, text="Coord. E (m):").pack(side="left", padx=(0, 2))
        self.txt_coord_e = ttk.Entry(frame_coords, width=15)
        self.txt_coord_e.pack(side="left", padx=(0, 8))

        ttk.Label(frame_coords, text="Coord. S (m):").pack(side="left", padx=(0, 2))
        self.txt_coord_s = ttk.Entry(frame_coords, width=15)
        self.txt_coord_s.pack(side="left", padx=(0, 8))

        ttk.Label(frame_coords, text="Data:").pack(side="left", padx=(0, 2))
        self.txt_data = ttk.Entry(frame_coords, width=11)
        self.txt_data.insert(0, data_hoje_br())
        self.txt_data.pack(side="left", padx=(0, 10))

        ttk.Label(frame_coords, text="Link:").pack(side="left", padx=(0, 2))
        self.txt_link = ttk.Entry(frame_coords, width=36)
        self.txt_link.pack(side="left")

        frame_fotos = ttk.Frame(frame_form)
        frame_fotos.grid(row=4, column=0, columnspan=6, sticky="w", pady=2)

        ttk.Label(frame_fotos, text="Foto 1 (Imóvel/Drone):").pack(side="left")
        self.txt_foto1 = ttk.Entry(frame_fotos, width=34)
        self.txt_foto1.pack(side="left", padx=4)
        ttk.Button(frame_fotos, text="Buscar", width=7, command=lambda: self._buscar_arquivo_foto(self.txt_foto1)).pack(side="left", padx=(0, 12))

        ttk.Label(frame_fotos, text="Foto 2 (Print Anúncio):").pack(side="left")
        self.txt_foto2 = ttk.Entry(frame_fotos, width=34)
        self.txt_foto2.pack(side="left", padx=4)
        ttk.Button(frame_fotos, text="Buscar", width=7, command=lambda: self._buscar_arquivo_foto(self.txt_foto2)).pack(side="left")

        self.frame_vars = ttk.LabelFrame(self.container_principal, text=" Variáveis da Avaliação (Nota Agronômica automática ao escolher Aptidão e Situação) ", padding=6)
        self.frame_vars.pack(fill="x", padx=10, pady=2)
        self.widgets_dinamicos = {}

        frame_btn_cad = ttk.Frame(self.container_principal, padding=2)
        frame_btn_cad.pack(fill="x", padx=10)

        self.btn_salvar_dado = ttk.Button(frame_btn_cad, text="➕ Adicionar à Grade", style="Primary.TButton", command=self._adicionar_ou_salvar_dado)
        self.btn_salvar_dado.pack(side="left", padx=3)

        ttk.Button(frame_btn_cad, text="🧹 Limpar Campos", command=self._limpar_formulario).pack(side="left", padx=3)

        self.btn_cancelar_edicao = ttk.Button(frame_btn_cad, text="✖ Cancelar Edição", command=self._limpar_formulario, state="disabled")
        self.btn_cancelar_edicao.pack(side="left", padx=3)

        ttk.Button(frame_btn_cad, text="📋 Colar do Excel (Ctrl+V)", style="Accent.TButton", command=self._colar_dados_clipboard).pack(side="left", padx=10)
        ttk.Button(frame_btn_cad, text="⚙ Configurar Variáveis / Perfis", command=self._janela_config_variaveis).pack(side="right", padx=3)

        # --- BARRA DE FÓRMULA / EDIÇÃO RÁPIDA ESTILO SISDEA ---
        frame_barra_sisdea = ttk.Frame(self.container_principal, padding=(10, 3))
        frame_barra_sisdea.pack(fill="x")

        self.lbl_celula_sisdea = tk.Label(
            frame_barra_sisdea,
            text="Campo / D.",
            width=22,
            anchor="w",
            bg="#e1e7ef",
            fg="#1f3c5b",
            font=("Segoe UI", 9, "bold"),
            relief="groove",
            padx=6
        )
        self.lbl_celula_sisdea.pack(side="left")

        tk.Label(frame_barra_sisdea, text=" ➔ ", bg="#f4f6f9", fg="#2b7a78", font=("Segoe UI", 11, "bold")).pack(side="left")

        self.txt_barra_sisdea = ttk.Entry(frame_barra_sisdea, font=("Segoe UI", 9))
        self.txt_barra_sisdea.pack(side="left", fill="x", expand=True, padx=4)
        self.txt_barra_sisdea.bind("<Return>", lambda event: self._confirmar_edicao_barra_sisdea())

        ttk.Button(frame_barra_sisdea, text="✔ Aplicar na Célula (Enter)", command=self._confirmar_edicao_barra_sisdea).pack(side="left", padx=3)

        # --- GRADE ESTILO PLANILHA SISDEA ---
        frame_tabela = ttk.LabelFrame(
            self.container_principal,
            text=" Grade de Dados (Clique na coluna '☑' para Ativar/Desativar • Clique simples edita na barra • Duplo clique carrega no formulário) ",
            padding=6
        )
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=2)

        frame_tree_container = ttk.Frame(frame_tabela)
        frame_tree_container.pack(fill="both", expand=True)

        self.scroll_y = ttk.Scrollbar(frame_tree_container, orient="vertical")
        self.scroll_x = ttk.Scrollbar(frame_tree_container, orient="horizontal")

        self.tree = ttk.Treeview(
            frame_tree_container,
            show="headings",
            height=9,
            yscrollcommand=self.scroll_y.set,
            xscrollcommand=self.scroll_x.set,
            selectmode="extended"
        )
        self.scroll_y.config(command=self.tree.yview)
        self.scroll_x.config(command=self.tree.xview)

        self.scroll_y.pack(side="right", fill="y")
        self.scroll_x.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.tag_configure("inativo", foreground="#888888", background="#f2f2f2")
        self.tree.tag_configure("ativo_par", background="#ffffff")
        self.tree.tag_configure("ativo_impar", background="#f7FAFC")

        self.tree.bind("<ButtonRelease-1>", self._ao_clicar_celula_grid)
        self.tree.bind("<Double-1>", lambda event: self._carregar_para_edicao())
        self.root.bind("<Control-v>", lambda event: self._colar_dados_clipboard_atalho(event))

        frame_botoes_grid = ttk.Frame(frame_tabela)
        frame_botoes_grid.pack(fill="x", pady=3)
        ttk.Button(frame_botoes_grid, text="✏ Carregar no Formulário", command=self._carregar_para_edicao).pack(side="left", padx=3)
        ttk.Button(frame_botoes_grid, text="☑ Alternar Utilizado/Não Utilizado", command=self._alternar_utilizado_selecionados).pack(side="left", padx=3)
        ttk.Button(frame_botoes_grid, text="⚡ Edição em Lote", command=self._janela_edicao_lote).pack(side="left", padx=3)
        ttk.Button(frame_botoes_grid, text="🗑 Excluir Selecionado(s)", command=self._excluir_dado).pack(side="left", padx=3)
        ttk.Button(frame_botoes_grid, text="📥 Importar Planilha (.xlsx)", command=self._importar_planilha_excel).pack(side="right", padx=3)

        self.frame_stats = tk.Frame(self.container_principal, bg="#e1e7ef", bd=1, relief="solid")
        self.frame_stats.pack(fill="x", padx=10, pady=2)
        self.lbl_estatisticas = tk.Label(
            self.frame_stats,
            text="📊 Resumo Estatístico (Amostras Ativas): 0 dados | Média: R$ 0,00 | Mín: R$ 0,00 | Máx: R$ 0,00 | CV: 0,00%",
            bg="#e1e7ef",
            fg="#1f3c5b",
            font=("Segoe UI", 9, "bold"),
            pady=3
        )
        self.lbl_estatisticas.pack(anchor="w", padx=10)

        self.progress_bar = ttk.Progressbar(self.container_principal, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=12, pady=2)
        self.progress_bar.pack_forget()

        frame_acoes = ttk.Frame(self.container_principal, padding=4)
        frame_acoes.pack(fill="x", padx=10, pady=3)

        ttk.Button(frame_acoes, text="📊 Exportar Planilha Excel (SISDEA)", style="Accent.TButton", command=self._exportar_excel).pack(side="left", padx=6, expand=True, fill="x")
        ttk.Button(frame_acoes, text="📄 Gerar Relatório Word (Fichas Prontas)", style="Primary.TButton", command=self._iniciar_exportacao_word_thread).pack(side="right", padx=6, expand=True, fill="x")

    def _configurar_colunas_grid(self):
        if not hasattr(self, "tree"):
            return
        colunas_base = [
            ("dado_id", "D.", 38, "center"),
            ("utilizado", "☑", 38, "center"),
            ("informante", "Informante", 130, "w"),
            ("telefone", "Telefone", 105, "w"),
            ("endereco", "Endereço", 160, "w"),
            ("bairro", "Bairro", 110, "w"),
            ("municipio", "Município", 110, "w"),
            ("valor_total", "* Valor Total", 105, "e"),
            ("area_terreno", "Área", 85, "e"),
            ("unidade", "Un.", 42, "center"),
        ]
        colunas_vars = []
        for v_cfg in self.variaveis_config:
            nome_v = v_cfg["nome"]
            colunas_vars.append((f"var_{nome_v}", nome_v, 95, "center"))

        colunas_fim = [
            ("unitario", "Unitário (R$/un)", 110, "e"),
            ("status", "Status", 65, "center")
        ]

        todas_colunas = colunas_base + colunas_vars + colunas_fim
        self.mapa_colunas_grid = todas_colunas

        ids_cols = [c[0] for c in todas_colunas]
        self.tree["columns"] = ids_cols

        for col_id, titulo, larg, alinhamento in todas_colunas:
            self.tree.heading(col_id, text=titulo)
            self.tree.column(col_id, width=larg, minwidth=36, anchor=alinhamento, stretch=False)

    def _atualizar_interface_variaveis(self):
        if not hasattr(self, "frame_vars"):
            return
        for w in self.frame_vars.winfo_children():
            w.destroy()
        self.widgets_dinamicos.clear()

        for idx, var in enumerate(self.variaveis_config):
            col = (idx % 3) * 2
            row = idx // 3

            lbl = ttk.Label(self.frame_vars, text=f"{var['nome']}:")
            lbl.grid(row=row, column=col, sticky="w", padx=4, pady=2)

            if var["tipo"] == "codigo" and var.get("opcoes"):
                cb = ttk.Combobox(self.frame_vars, values=var["opcoes"], width=23)
                if var["opcoes"]:
                    cb.set(var["opcoes"][0])
                cb.grid(row=row, column=col + 1, sticky="w", padx=4, pady=2)
                cb.bind("<<ComboboxSelected>>", lambda e: self._verificar_calculo_nota_agronomica())
                cb.bind("<KeyRelease>", lambda e: self._verificar_calculo_nota_agronomica())
                self.widgets_dinamicos[var["nome"]] = cb
            else:
                ent = ttk.Entry(self.frame_vars, width=25)
                ent.grid(row=row, column=col + 1, sticky="w", padx=4, pady=2)
                self.widgets_dinamicos[var["nome"]] = ent

        self._verificar_calculo_nota_agronomica()

    def _verificar_calculo_nota_agronomica(self):
        w_apt = self.widgets_dinamicos.get("Aptidão")
        w_sit = self.widgets_dinamicos.get("Situação") or self.widgets_dinamicos.get("Acesso")
        w_nota = self.widgets_dinamicos.get("Nota Agronômica")

        if w_apt and w_sit and w_nota:
            nota = calcular_nota_agronomica(w_apt.get(), w_sit.get())
            if nota:
                w_nota.delete(0, tk.END)
                w_nota.insert(0, nota)

    def _mostrar_tabela_nota_agronomica(self):
        janela = tk.Toplevel(self.root)
        janela.title("Tabela 02 – Nota Agronômica (Kozma 1984 x IN RFB 1877/2019)")
        janela.geometry("620x280")
        janela.transient(self.root)

        ttk.Label(
            janela,
            text="Tabela 02 – Cruzamento Automático de Aptidão x Situação/Acesso",
            font=("Segoe UI", 10, "bold"),
            foreground="#1f3c5b"
        ).pack(pady=8)

        cols = ("situacao", "peso", "I", "II", "III", "IV", "V", "VI")
        t_ref = ttk.Treeview(janela, columns=cols, show="headings", height=7)
        t_ref.heading("situacao", text="Situação / Acesso")
        t_ref.heading("peso", text="Peso")
        for rom, p_apt in [("I", "1,00"), ("II", "0,80"), ("III", "0,64"), ("IV", "0,51"), ("V", "0,41"), ("VI", "0,20")]:
            t_ref.heading(rom, text=f"{rom} ({p_apt})")
            t_ref.column(rom, width=65, anchor="center")

        t_ref.column("situacao", width=120, anchor="w")
        t_ref.column("peso", width=55, anchor="center")
        t_ref.pack(fill="both", expand=True, padx=10, pady=4)

        pesos_sit = [("Ótima", "1,00"), ("Muito Boa", "0,95"), ("Boa", "0,90"), ("Desfavorável", "0,80"), ("Má", "0,75"), ("Péssima", "0,70")]
        for sit, p_s in pesos_sit:
            t_ref.insert("", "end", values=(
                sit, p_s,
                TABELA_NOTA_AGRONOMICA["I"][sit],
                TABELA_NOTA_AGRONOMICA["II"][sit],
                TABELA_NOTA_AGRONOMICA["III"][sit],
                TABELA_NOTA_AGRONOMICA["IV"][sit],
                TABELA_NOTA_AGRONOMICA["V"][sit],
                TABELA_NOTA_AGRONOMICA["VI"][sit]
            ))

    def _janela_conversor_utm(self):
        texto = simpledialog.askstring(
            "Conversor Lat/Long ou Google Maps para UTM",
            "Cole abaixo o par de Coordenadas (ex: -16.6868, -49.2648)\nou o link completo do Google Maps:",
            parent=self.root
        )
        if not texto:
            return
        res = converter_latlon_para_utm(texto)
        if not res:
            messagebox.showwarning("Formato não reconhecido", "Não foi possível localizar Latitude e Longitude decimais no texto informado.")
            return
        zona, e_val, s_val = res
        self.txt_zona.delete(0, tk.END)
        self.txt_zona.insert(0, zona)
        self.txt_coord_e.delete(0, tk.END)
        self.txt_coord_e.insert(0, e_val)
        self.txt_coord_s.delete(0, tk.END)
        self.txt_coord_s.insert(0, s_val)
        messagebox.showinfo("Coordenadas Convertidas", f"Conversão concluída (SIRGAS 2000 / WGS84):\nZona: {zona}\nCoord. E: {e_val} m\nCoord. S: {s_val} m")

    def _janela_config_variaveis(self):
        janela = tk.Toplevel(self.root)
        janela.title("Gerenciador de Variáveis e Perfis Pré-Cadastrados")
        janela.geometry("760x540")
        janela.transient(self.root)
        janela.grab_set()

        frame_perfis = ttk.LabelFrame(janela, text=" 1. Carregar Perfil Pré-Cadastrado com 1 Clique ", padding=8)
        frame_perfis.pack(fill="x", padx=10, pady=6)

        entradas_vars = []

        def recarregar_linhas_editor(lista_cfg):
            for widget in frame_lista.winfo_children():
                widget.destroy()
            entradas_vars.clear()

            total_slots = max(10, len(lista_cfg) + 3)
            ttk.Label(frame_lista, text="Ativar", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, padx=2)
            ttk.Label(frame_lista, text="Nome da Variável", font=("Segoe UI", 8, "bold")).grid(row=0, column=1, padx=3, sticky="w")
            ttk.Label(frame_lista, text="Tipo", font=("Segoe UI", 8, "bold")).grid(row=0, column=2, padx=3)
            ttk.Label(frame_lista, text="Opções (separadas por vírgula)", font=("Segoe UI", 8, "bold")).grid(row=0, column=3, padx=3, sticky="w")

            for i in range(total_slots):
                ativa = i < len(lista_cfg)
                cfg = lista_cfg[i] if ativa else {"nome": "", "tipo": "texto", "opcoes": []}

                var_ativa = tk.BooleanVar(value=ativa)
                chk = ttk.Checkbutton(frame_lista, text=f"Var {i+1}", variable=var_ativa)
                chk.grid(row=i+1, column=0, padx=2, pady=2)

                e_nome = ttk.Entry(frame_lista, width=20)
                e_nome.insert(0, cfg.get("nome", ""))
                e_nome.grid(row=i+1, column=1, padx=3, pady=2)

                cb_tipo = ttk.Combobox(frame_lista, values=["codigo", "numero", "texto"], width=9, state="readonly")
                cb_tipo.set(cfg.get("tipo", "texto"))
                cb_tipo.grid(row=i+1, column=2, padx=3, pady=2)

                e_opcoes = ttk.Entry(frame_lista, width=42)
                e_opcoes.insert(0, ", ".join(cfg.get("opcoes", [])))
                e_opcoes.grid(row=i+1, column=3, padx=3, pady=2)

                entradas_vars.append((var_ativa, e_nome, cb_tipo, e_opcoes))

        ttk.Button(frame_perfis, text="🏙 Perfil Urbano (Via, Uso, Testada, PGV, PIB, Condição)", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_PADRAO))).pack(side="left", padx=4)
        ttk.Button(frame_perfis, text="🌾 Perfil Rural (Aptidão, Situação, Nota Agronômica, Benfeitoria)", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_RURAL))).pack(side="left", padx=4)
        ttk.Button(frame_perfis, text="💧 Perfil COPASA", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_COPASA))).pack(side="left", padx=4)

        frame_cat = ttk.Frame(janela, padding=(10, 2))
        frame_cat.pack(fill="x")
        ttk.Label(frame_cat, text="Adicionar variável usual do catálogo:").pack(side="left")
        nomes_cat = [v["nome"] for v in CATALOGO_VARIAVEIS_USUAIS]
        cb_cat = ttk.Combobox(frame_cat, values=nomes_cat, width=25, state="readonly")
        cb_cat.pack(side="left", padx=6)

        def adicionar_do_catalogo():
            escolha = cb_cat.get()
            if not escolha:
                return
            item_cat = next((v for v in CATALOGO_VARIAVEIS_USUAIS if v["nome"] == escolha), None)
            if not item_cat:
                return
            atuais = []
            for var_ativa, e_nome, cb_tipo, e_opcoes in entradas_vars:
                if var_ativa.get() and e_nome.get().strip():
                    opts = [op.strip() for op in e_opcoes.get().split(",") if op.strip()]
                    atuais.append({"nome": e_nome.get().strip(), "tipo": cb_tipo.get(), "opcoes": opts})
            if not any(normalizar_chave_var(x["nome"]) == normalizar_chave_var(item_cat["nome"]) for x in atuais):
                atuais.append(copy.deepcopy(item_cat))
            recarregar_linhas_editor(atuais)

        ttk.Button(frame_cat, text="➕ Incluir Variável", command=adicionar_do_catalogo).pack(side="left")

        frame_lista = ttk.LabelFrame(janela, text=" 2. Variáveis Ativas no Projeto ", padding=8)
        frame_lista.pack(fill="both", expand=True, padx=10, pady=6)

        recarregar_linhas_editor(self.variaveis_config)

        def salvar():
            novas = []
            for var_ativa, e_nome, cb_tipo, e_opcoes in entradas_vars:
                if var_ativa.get() and e_nome.get().strip():
                    opts = [op.strip() for op in e_opcoes.get().split(",") if op.strip()]
                    novas.append({
                        "nome": e_nome.get().strip(),
                        "tipo": cb_tipo.get(),
                        "opcoes": opts
                    })
            self.variaveis_config = novas
            self._salvar_config_variaveis()
            self._atualizar_interface_variaveis()
            self._configurar_colunas_grid()
            self._recarregar_grid()
            janela.destroy()
            messagebox.showinfo("Sucesso", "Variáveis e colunas da grade atualizadas com sucesso!")

        ttk.Button(janela, text="💾 Salvar e Atualizar Grade", style="Primary.TButton", command=salvar).pack(pady=8)

    def _buscar_arquivo_foto(self, entry_widget):
        caminho = filedialog.askopenfilename(
            filetypes=[
                ("Imagens", "*.png *.jpg *.jpeg *.webp *.bmp"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if caminho:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, caminho)

    def _limpar_formulario(self):
        self.item_em_edicao = None
        self.var_utilizado.set(True)
        self.txt_informante.delete(0, tk.END)
        self.txt_telefone.delete(0, tk.END)
        self.txt_endereco.delete(0, tk.END)
        self.txt_bairro.delete(0, tk.END)
        self.txt_municipio.delete(0, tk.END)
        self.txt_valor.delete(0, tk.END)
        self.txt_area.delete(0, tk.END)
        self.txt_area_const.delete(0, tk.END)
        self.txt_zona.delete(0, tk.END)
        self.txt_coord_e.delete(0, tk.END)
        self.txt_coord_s.delete(0, tk.END)
        self.txt_data.delete(0, tk.END)
        self.txt_data.insert(0, data_hoje_br())
        self.txt_link.delete(0, tk.END)
        self.txt_foto1.delete(0, tk.END)
        self.txt_foto2.delete(0, tk.END)

        for _, widget in self.widgets_dinamicos.items():
            if isinstance(widget, ttk.Combobox):
                vals = widget.cget("values")
                if vals:
                    widget.set(vals[0])
            else:
                widget.delete(0, tk.END)

        self._verificar_calculo_nota_agronomica()
        self.btn_salvar_dado.config(text="➕ Adicionar à Grade")
        self.btn_cancelar_edicao.config(state="disabled")

    def _adicionar_ou_salvar_dado(self):
        try:
            self._verificar_calculo_nota_agronomica()
            valor_total = converter_para_float(self.txt_valor.get())
            area_num = converter_para_float(self.txt_area.get())
            area_const = converter_para_float(self.txt_area_const.get())
            unidade = self.var_unidade.get()
            unitario = valor_total / area_num if area_num > 0 else 0.0

            dado_num = self.item_em_edicao.get("dado_id", self.item_em_edicao.get("D.", 1)) if self.item_em_edicao else (len(self.dados_pesquisas) + 1)

            coord_e = limpar_sufixo_coord(self.txt_coord_e.get())
            coord_s = limpar_sufixo_coord(self.txt_coord_s.get())

            registro = {
                "dado_id": int(dado_num),
                "utilizado": bool(self.var_utilizado.get()),
                "informante": self.txt_informante.get().strip(),
                "telefone": self.txt_telefone.get().strip(),
                "endereco": self.txt_endereco.get().strip(),
                "bairro": self.txt_bairro.get().strip(),
                "municipio": self.txt_municipio.get().strip(),
                "valor_total": valor_total,
                "area_terreno": area_num,
                "area_construida": area_const,
                "unitario": unitario,
                "unidade": unidade,
                "localizacao": "Rural" if unidade == "ha" else "Urbana",
                "zona_utm": self.txt_zona.get().strip(),
                "coord_e": coord_e,
                "coord_s": coord_s,
                "data": self.txt_data.get().strip(),
                "link": self.txt_link.get().strip(),
                "foto1": self.txt_foto1.get().strip(),
                "foto2": self.txt_foto2.get().strip(),
                "variaveis_extras": {}
            }

            for nome, widget in self.widgets_dinamicos.items():
                val = widget.get().strip()
                registro["variaveis_extras"][nome] = val

            if self.item_em_edicao:
                idx = next((i for i, d in enumerate(self.dados_pesquisas) if (d.get("dado_id") or d.get("D.")) == dado_num), None)
                if idx is not None:
                    self.dados_pesquisas[idx] = registro
            else:
                self.dados_pesquisas.append(registro)

            self._recarregar_grid()
            self._limpar_formulario()
        except Exception as e:
            messagebox.showerror("Erro de Preenchimento", f"Verifique os dados numéricos: {e}")

    def _recarregar_grid(self):
        if not hasattr(self, "tree"):
            return
        for item in self.tree.get_children():
            self.tree.delete(item)

        unitarios_ativos = []
        un_ref = self.var_unidade.get() if hasattr(self, "var_unidade") else "m²"

        for idx, dado in enumerate(self.dados_pesquisas):
            d_id = dado.get("dado_id") or dado.get("D.", idx + 1)
            utilizado = dado.get("utilizado", True)
            un = dado.get("unidade") or dado.get("Unidade", "m²")
            un_ref = un
            casas = 4 if un == "ha" else 2

            v_total = converter_para_float(dado.get("valor_total") if dado.get("valor_total") is not None else dado.get("Valor Total (R$)", 0.0))
            a_total = converter_para_float(dado.get("area_terreno") if dado.get("area_terreno") is not None else (dado.get(f"Área Terreno ({un})") or 0.0))
            u_unit = v_total / a_total if a_total > 0 else 0.0
            dado["unitario"] = u_unit

            if utilizado and u_unit > 0:
                unitarios_ativos.append(u_unit)

            tem_foto = bool(dado.get("foto1") or dado.get("foto2") or dado.get("Foto1") or dado.get("Foto2"))
            tem_coord = bool(dado.get("coord_e") or dado.get("Coord. E (m)"))
            status_ic = f"{'📷' if tem_foto else '⚪'}{'📍' if tem_coord else '⚪'}"

            valores_linha = [
                d_id,
                "☑" if utilizado else "☐",
                dado.get("informante") or dado.get("Informante", ""),
                dado.get("telefone") or dado.get("Telefone", ""),
                dado.get("endereco") or dado.get("Endereço", ""),
                dado.get("bairro") or dado.get("Bairro", ""),
                dado.get("municipio") or dado.get("Município", ""),
                formatar_moeda_br(v_total),
                formatar_numero_br(a_total, casas),
                un
            ]

            extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras", {})
            for v_cfg in self.variaveis_config:
                nome_v = v_cfg["nome"]
                valores_linha.append(extras.get(nome_v, dado.get(nome_v, "")))

            valores_linha.append(formatar_moeda_br(u_unit))
            valores_linha.append(status_ic)

            tag = "inativo" if not utilizado else ("ativo_par" if idx % 2 == 0 else "ativo_impar")
            self.tree.insert("", "end", values=valores_linha, tags=(tag,))

        self._atualizar_estatisticas_rodape(unitarios_ativos, len(self.dados_pesquisas), un_ref)

    def _atualizar_estatisticas_rodape(self, unitarios_ativos, total_geral, unidade):
        if not hasattr(self, "lbl_estatisticas"):
            return
        n_ativos = len(unitarios_ativos)
        if n_ativos == 0:
            self.lbl_estatisticas.config(
                text=f"📊 Resumo Estatístico: 0 ativos de {total_geral} cadastrados | Média: R$ 0,00/{unidade} | Mín: R$ 0,00 | Máx: R$ 0,00 | CV: 0,00%"
            )
            return
        media = statistics.mean(unitarios_ativos)
        minimo = min(unitarios_ativos)
        maximo = max(unitarios_ativos)
        desvio = statistics.stdev(unitarios_ativos) if n_ativos > 1 else 0.0
        cv = (desvio / media * 100.0) if media > 0 else 0.0

        self.lbl_estatisticas.config(
            text=(
                f"📊 Resumo Estatístico (Amostras Ativas: {n_ativos}/{total_geral})  •  "
                f"Média: R$ {formatar_moeda_br(media)}/{unidade}  •  "
                f"Mín: R$ {formatar_moeda_br(minimo)}  •  "
                f"Máx: R$ {formatar_moeda_br(maximo)}  •  "
                f"CV: {formatar_numero_br(cv, 2)}%   (📷=Foto | 📍=Coord)"
            )
        )

    def _ao_clicar_celula_grid(self, event):
        regiao = self.tree.identify("region", event.x, event.y)
        if regiao != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id_str = self.tree.identify_column(event.x)
        if not row_id or not col_id_str:
            return

        col_idx = int(col_id_str.replace("#", "")) - 1
        if col_idx < 0 or col_idx >= len(self.mapa_colunas_grid):
            return

        valores = self.tree.item(row_id, "values")
        if not valores:
            return
        dado_id = int(valores[0])
        dado = next((d for d in self.dados_pesquisas if (d.get("dado_id") or d.get("D.")) == dado_id), None)
        if not dado:
            return

        chave_col, titulo_col, _, _ = self.mapa_colunas_grid[col_idx]

        if chave_col == "utilizado":
            dado["utilizado"] = not dado.get("utilizado", True)
            self._recarregar_grid()
            return

        if chave_col in ("dado_id", "unitario", "status"):
            return

        valor_atual = valores[col_idx]
        self.celula_ativa = {"dado_id": dado_id, "campo": chave_col, "rotulo": titulo_col}
        self.lbl_celula_sisdea.config(text=f"{titulo_col} / {dado_id}")
        self.txt_barra_sisdea.delete(0, tk.END)
        self.txt_barra_sisdea.insert(0, str(valor_atual))
        self.txt_barra_sisdea.focus_set()
        self.txt_barra_sisdea.select_range(0, tk.END)

    def _confirmar_edicao_barra_sisdea(self):
        d_id = self.celula_ativa.get("dado_id")
        campo = self.celula_ativa.get("campo")
        if not d_id or not campo:
            return

        dado = next((d for d in self.dados_pesquisas if (d.get("dado_id") or d.get("D.")) == d_id), None)
        if not dado:
            return

        novo_val = self.txt_barra_sisdea.get().strip()

        if campo in ("valor_total", "area_terreno"):
            dado[campo] = converter_para_float(novo_val)
            a_t = converter_para_float(dado.get("area_terreno", 0.0))
            v_t = converter_para_float(dado.get("valor_total", 0.0))
            dado["unitario"] = v_t / a_t if a_t > 0 else 0.0
        elif campo == "unidade":
            dado["unidade"] = "ha" if "ha" in novo_val.lower() else "m²"
            dado["localizacao"] = "Rural" if dado["unidade"] == "ha" else "Urbana"
        elif campo.startswith("var_"):
            nome_var = campo[4:]
            if "variaveis_extras" not in dado:
                dado["variaveis_extras"] = {}
            dado["variaveis_extras"][nome_var] = novo_val

            if nome_var in ("Aptidão", "Situação", "Acesso"):
                apt = dado["variaveis_extras"].get("Aptidão", "")
                sit = dado["variaveis_extras"].get("Situação") or dado["variaveis_extras"].get("Acesso", "")
                nota_auto = calcular_nota_agronomica(apt, sit)
                if nota_auto:
                    dado["variaveis_extras"]["Nota Agronômica"] = nota_auto
        else:
            dado[campo] = novo_val

        self._recarregar_grid()

    def _alternar_utilizado_selecionados(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showwarning("Aviso", "Selecione uma ou mais linhas na grade.")
            return
        for item in sels:
            d_id = int(self.tree.item(item, "values")[0])
            dado = next((d for d in self.dados_pesquisas if (d.get("dado_id") or d.get("D.")) == d_id), None)
            if dado:
                dado["utilizado"] = not dado.get("utilizado", True)
        self._recarregar_grid()

    def _janela_edicao_lote(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showwarning("Edição em Lote", "Selecione as pesquisas na grade (use Shift ou Ctrl) para editar em lote.")
            return

        janela = tk.Toplevel(self.root)
        janela.title(f"Edição em Lote ({len(sels)} pesquisas selecionadas)")
        janela.geometry("420x250")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text=f"Aplicar valor a {len(sels)} pesquisas selecionadas:", font=("Segoe UI", 10, "bold")).pack(pady=10)

        campos_opcoes = ["Município", "Bairro", "Data", "Zona UTM", "Unidade (m²/ha)", "Status (Utilizado / Não utilizado)"]
        for v_cfg in self.variaveis_config:
            campos_opcoes.append(f"Variável: {v_cfg['nome']}")

        ttk.Label(janela, text="Campo a alterar:").pack(anchor="w", padx=30)
        cb_campo = ttk.Combobox(janela, values=campos_opcoes, width=38, state="readonly")
        cb_campo.set(campos_opcoes[0])
        cb_campo.pack(padx=30, pady=4)

        ttk.Label(janela, text="Novo valor:").pack(anchor="w", padx=30)
        e_val = ttk.Entry(janela, width=40)
        e_val.pack(padx=30, pady=4)

        def aplicar():
            escolha = cb_campo.get()
            val = e_val.get().strip()
            ids_sel = [int(self.tree.item(it, "values")[0]) for it in sels]

            for dado in self.dados_pesquisas:
                if (dado.get("dado_id") or dado.get("D.")) in ids_sel:
                    if escolha == "Município":
                        dado["municipio"] = val
                    elif escolha == "Bairro":
                        dado["bairro"] = val
                    elif escolha == "Data":
                        dado["data"] = val
                    elif escolha == "Zona UTM":
                        dado["zona_utm"] = val
                    elif escolha == "Unidade (m²/ha)":
                        dado["unidade"] = "ha" if "ha" in val.lower() else "m²"
                        dado["localizacao"] = "Rural" if dado["unidade"] == "ha" else "Urbana"
                    elif escolha == "Status (Utilizado / Não utilizado)":
                        dado["utilizado"] = val.lower() not in ("não", "nao", "0", "false")
                    elif escolha.startswith("Variável: "):
                        n_var = escolha.replace("Variável: ", "")
                        dado.setdefault("variaveis_extras", {})[n_var] = val
                        if n_var in ("Aptidão", "Situação", "Acesso"):
                            apt = dado["variaveis_extras"].get("Aptidão", "")
                            sit = dado["variaveis_extras"].get("Situação") or dado["variaveis_extras"].get("Acesso", "")
                            n_auto = calcular_nota_agronomica(apt, sit)
                            if n_auto:
                                dado["variaveis_extras"]["Nota Agronômica"] = n_auto

            self._recarregar_grid()
            janela.destroy()
            messagebox.showinfo("Sucesso", "Edição em lote aplicada!")

        ttk.Button(janela, text="✔ Aplicar a Todos Selecionados", style="Primary.TButton", command=aplicar).pack(pady=15)

    # --- COLAR DIRETO DO EXCEL (CTRL+V ESTILO SISDEA) ---
    def _colar_dados_clipboard_atalho(self, event):
        if not hasattr(self, "tree"):
            return
        foco = self.root.focus_get()
        if isinstance(foco, (ttk.Entry, tk.Entry)):
            return
        self._colar_dados_clipboard()

    def _colar_dados_clipboard(self):
        if not hasattr(self, "tree"):
            return
        try:
            texto_clip = self.root.clipboard_get()
        except Exception:
            messagebox.showwarning("Área de Transferência Vazia", "Copie as células/linhas no Excel primeiro (Ctrl+C) antes de colar.")
            return

        linhas_raw = [l for l in texto_clip.splitlines() if l.strip()]
        if not linhas_raw:
            return

        linhas_split = [l.split("\t") for l in linhas_raw]

        primeira = [c.strip().lower() for c in linhas_split[0]]
        tem_cabecalho = any(k in " ".join(primeira) for k in ["informante", "endereço", "endereco", "logradouro", "valor total", "município"])

        if tem_cabecalho and len(linhas_split) > 1:
            n = len(linhas_split[0])
            corpo = [(l + [""] * n)[:n] for l in linhas_split[1:]]
            df = pd.DataFrame(corpo, columns=linhas_split[0])
            self._processar_dataframe_importado(df, origem="Área de Transferência (Ctrl+V)")
        else:
            novos = []
            un_atual = self.var_unidade.get() if hasattr(self, "var_unidade") else "m²"
            for cols in linhas_split:
                if len(cols) < 2:
                    continue
                offset = 1 if (cols[0].strip().isdigit() and len(cols) >= 7) else 0
                get_c = lambda idx: cols[idx + offset].strip() if (idx + offset) < len(cols) else ""

                inf = get_c(0)
                tel = get_c(1)
                end = get_c(2)
                bai = get_c(3)
                mun = get_c(4)
                v_tot = converter_para_float(get_c(5))
                a_ter = converter_para_float(get_c(6))
                u_calc = v_tot / a_ter if a_ter > 0 else 0.0

                item_id = len(self.dados_pesquisas) + len(novos) + 1
                reg = {
                    "dado_id": item_id,
                    "utilizado": True,
                    "informante": inf,
                    "telefone": tel,
                    "endereco": end,
                    "bairro": bai,
                    "municipio": mun,
                    "valor_total": v_tot,
                    "area_terreno": a_ter,
                    "area_construida": 0.0,
                    "unitario": u_calc,
                    "unidade": un_atual,
                    "localizacao": "Rural" if un_atual == "ha" else "Urbana",
                    "zona_utm": "",
                    "coord_e": "",
                    "coord_s": "",
                    "data": self.txt_data.get().strip() or data_hoje_br(),
                    "link": "",
                    "foto1": "",
                    "foto2": "",
                    "variaveis_extras": {}
                }

                for idx_v, v_cfg in enumerate(self.variaveis_config):
                    val_v = get_c(7 + idx_v)
                    if val_v:
                        reg["variaveis_extras"][v_cfg["nome"]] = val_v

                apt = reg["variaveis_extras"].get("Aptidão", "")
                sit = reg["variaveis_extras"].get("Situação") or reg["variaveis_extras"].get("Acesso", "")
                nota_auto = calcular_nota_agronomica(apt, sit)
                if nota_auto and not reg["variaveis_extras"].get("Nota Agronômica"):
                    reg["variaveis_extras"]["Nota Agronômica"] = nota_auto

                novos.append(reg)

            if novos:
                self.dados_pesquisas.extend(novos)
                self._recarregar_grid()
                messagebox.showinfo("Colagem Concluída", f"{len(novos)} pesquisas coladas diretamente na grade!")

    def _carregar_para_edicao(self):
        sel = self.tree.selection()
        if not sel:
            return

        item_val = self.tree.item(sel[0])["values"]
        dado_id = int(item_val[0])
        dado = next((d for d in self.dados_pesquisas if (d.get("dado_id") or d.get("D.")) == dado_id), None)
        if not dado:
            return

        self.item_em_edicao = dado
        self.btn_salvar_dado.config(text=f"💾 Atualizar Dado {dado_id}")
        self.btn_cancelar_edicao.config(state="normal")

        self.var_utilizado.set(dado.get("utilizado", True))

        self.txt_informante.delete(0, tk.END)
        self.txt_informante.insert(0, dado.get("informante") or dado.get("Informante", ""))

        self.txt_telefone.delete(0, tk.END)
        self.txt_telefone.insert(0, dado.get("telefone") or dado.get("Telefone", ""))

        self.txt_endereco.delete(0, tk.END)
        self.txt_endereco.insert(0, dado.get("endereco") or dado.get("Endereço", ""))

        self.txt_bairro.delete(0, tk.END)
        self.txt_bairro.insert(0, dado.get("bairro") or dado.get("Bairro", ""))

        self.txt_municipio.delete(0, tk.END)
        self.txt_municipio.insert(0, dado.get("municipio") or dado.get("Município", ""))

        v_total = dado.get("valor_total") if dado.get("valor_total") is not None else dado.get("Valor Total (R$)", 0.0)
        self.txt_valor.delete(0, tk.END)
        self.txt_valor.insert(0, formatar_moeda_br(v_total))

        un = dado.get("unidade") or dado.get("Unidade", "m²")
        self.var_unidade.set(un)

        a_total = dado.get("area_terreno") if dado.get("area_terreno") is not None else (dado.get(f"Área Terreno ({un})") or 0.0)
        casas = 4 if un == "ha" else 2
        self.txt_area.delete(0, tk.END)
        self.txt_area.insert(0, formatar_numero_br(a_total, casas))

        a_const = dado.get("area_construida") if dado.get("area_construida") is not None else dado.get("Área Construída (m²)", 0.0)
        self.txt_area_const.delete(0, tk.END)
        if a_const > 0:
            self.txt_area_const.insert(0, formatar_numero_br(a_const, 2))

        self.txt_zona.delete(0, tk.END)
        self.txt_zona.insert(0, dado.get("zona_utm") or dado.get("Zona UTM", ""))

        self.txt_coord_e.delete(0, tk.END)
        self.txt_coord_e.insert(0, dado.get("coord_e") or dado.get("Coord. E (m)", ""))

        self.txt_coord_s.delete(0, tk.END)
        self.txt_coord_s.insert(0, dado.get("coord_s") or dado.get("Coord. S (m)", ""))

        self.txt_data.delete(0, tk.END)
        self.txt_data.insert(0, dado.get("data") or dado.get("Data", ""))

        self.txt_link.delete(0, tk.END)
        self.txt_link.insert(0, dado.get("link") or dado.get("Link", ""))

        self.txt_foto1.delete(0, tk.END)
        self.txt_foto1.insert(0, dado.get("foto1") or dado.get("Foto1", ""))

        self.txt_foto2.delete(0, tk.END)
        self.txt_foto2.insert(0, dado.get("foto2") or dado.get("Foto2", ""))

        extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras", {})
        for nome, widget in self.widgets_dinamicos.items():
            val = extras.get(nome, "")
            if isinstance(widget, ttk.Combobox):
                widget.set(val)
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(val))

    def _excluir_dado(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showwarning("Aviso", "Selecione uma ou mais pesquisas para excluir.")
            return

        ids_excluir = [int(self.tree.item(s, "values")[0]) for s in sels]
        if messagebox.askyesno("Confirmar Exclusão", f"Deseja excluir {len(ids_excluir)} pesquisa(s) selecionada(s)?"):
            self.dados_pesquisas = [d for d in self.dados_pesquisas if (d.get("dado_id") or d.get("D.")) not in ids_excluir]
            for i, d in enumerate(self.dados_pesquisas):
                d["dado_id"] = i + 1
            self._recarregar_grid()
            self._limpar_formulario()

    def _salvar_projeto(self):
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Projeto de Pesquisa", "*.json")])
        if not caminho:
            return
        dados_salvar = {
            "modelo": self.modelo_ativo,
            "variaveis_config": self.variaveis_config,
            "pesquisas": self.dados_pesquisas
        }
        try:
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(dados_salvar, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("Sucesso", "Projeto salvo com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar arquivo: {e}")

    def _executar_abertura_json(self):
        caminho = filedialog.askopenfilename(filetypes=[("Projeto de Pesquisa", "*.json")])
        if not caminho:
            return False
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                conteudo = json.load(f)
            self.modelo_ativo = conteudo.get("modelo", "PADRAO")
            self.variaveis_config = conteudo.get("variaveis_config") or conteudo.get("config_variaveis") or copy.deepcopy(DEFAULT_VARS_PADRAO)
            
            raw_pesquisas = conteudo.get("pesquisas", [])
            self.dados_pesquisas = [normalizar_dado(d) for d in raw_pesquisas]

            if hasattr(self, "tree"):
                self._atualizar_interface_variaveis()
                self._configurar_colunas_grid()
                self._recarregar_grid()
                self._limpar_formulario()
                messagebox.showinfo("Sucesso", f"{len(self.dados_pesquisas)} pesquisas carregadas com sucesso!")

            return True
        except Exception as e:
            messagebox.showerror("Erro ao Abrir", f"Arquivo corrompido ou formato inválido: {e}")
            return False

    def _abrir_projeto_menu(self):
        self._executar_abertura_json()

    def _importar_planilha_excel(self):
        caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
        if not caminho:
            return
        try:
            xls = pd.ExcelFile(caminho)
            sheet_name = xls.sheet_names[0]

            df_scan = pd.read_excel(caminho, sheet_name=sheet_name, header=None)
            header_row = 0
            for idx, row in df_scan.head(10).iterrows():
                vals = [str(v).strip().lower() for v in row.dropna().values]
                if any(k in vals for k in ["informante", "endereço", "endereco", "valor total (r$)", "área terreno (m²)"]):
                    header_row = idx
                    break

            df = pd.read_excel(caminho, sheet_name=sheet_name, header=header_row)
            self._processar_dataframe_importado(df, origem="Planilha Excel")
        except Exception as e:
            messagebox.showerror("Erro na Importação", f"Falha ao ler a planilha: {e}")

    def _processar_dataframe_importado(self, df, origem="Planilha"):
        mapeamento = {}
        unidade_detectada_coluna = None

        for col in df.columns:
            c_clean = str(col).strip().lower()
            if "informante" in c_clean:
                mapeamento["informante"] = col
            elif "telefone" in c_clean or "contato" in c_clean:
                mapeamento["telefone"] = col
            elif "endereço" in c_clean or "endereco" in c_clean or "logradouro" in c_clean:
                mapeamento["endereco"] = col
            elif "bairro" in c_clean:
                mapeamento["bairro"] = col
            elif "município" in c_clean or "municipio" in c_clean:
                mapeamento["municipio"] = col
            elif "valor total" in c_clean or "oferta" in c_clean:
                mapeamento["valor_total"] = col
            elif "área construída" in c_clean or "area construida" in c_clean or "constru" in c_clean:
                mapeamento["area_construida"] = col
            elif "área terreno" in c_clean or "area terreno" in c_clean or c_clean.startswith(("area", "área")):
                mapeamento["area_terreno"] = col
                if "(ha)" in c_clean:
                    unidade_detectada_coluna = "ha"
                elif "(m²)" in c_clean or "(m2)" in c_clean:
                    unidade_detectada_coluna = "m²"
            elif "coord. e" in c_clean or "coord e" in c_clean or "e (m)" in c_clean:
                mapeamento["coord_e"] = col
            elif "coord. s" in c_clean or "coord s" in c_clean or "s (m)" in c_clean:
                mapeamento["coord_s"] = col
            elif "zona" in c_clean:
                mapeamento["zona_utm"] = col
            elif "data" in c_clean:
                mapeamento["data"] = col
            elif "link" in c_clean:
                mapeamento["link"] = col
            elif "unidade" in c_clean or "unid" in c_clean:
                mapeamento["unidade"] = col
            elif "utilizado" in c_clean or "ativo" in c_clean:
                mapeamento["utilizado"] = col

        novos_dados = []
        for _, row in df.iterrows():
            if pd.isna(row.get(mapeamento.get("informante", ""))) and pd.isna(row.get(mapeamento.get("endereco", ""))):
                continue

            v_total = converter_para_float(row.get(mapeamento.get("valor_total", ""), 0.0))
            a_terr = converter_para_float(row.get(mapeamento.get("area_terreno", ""), 0.0))
            a_const = converter_para_float(row.get(mapeamento.get("area_construida", ""), 0.0))

            if "unidade" in mapeamento and not pd.isna(row.get(mapeamento["unidade"])):
                u_cand = str(row.get(mapeamento["unidade"])).strip().lower()
                unidade = "ha" if "ha" in u_cand else "m²"
            elif unidade_detectada_coluna:
                unidade = unidade_detectada_coluna
            else:
                unidade = self.var_unidade.get() if hasattr(self, "var_unidade") else "m²"

            utilizado_flag = True
            if "utilizado" in mapeamento and not pd.isna(row.get(mapeamento["utilizado"])):
                val_u = str(row.get(mapeamento["utilizado"])).strip().lower()
                utilizado_flag = val_u not in ("não", "nao", "false", "0", "☐")

            u_calc = v_total / a_terr if a_terr > 0 else 0.0
            item_id = len(self.dados_pesquisas) + len(novos_dados) + 1

            registro = {
                "dado_id": item_id,
                "utilizado": utilizado_flag,
                "informante": str(row.get(mapeamento.get("informante", ""), "")).replace("nan", "").strip(),
                "telefone": str(row.get(mapeamento.get("telefone", ""), "")).replace("nan", "").strip(),
                "endereco": str(row.get(mapeamento.get("endereco", ""), "")).replace("nan", "").strip(),
                "bairro": str(row.get(mapeamento.get("bairro", ""), "")).replace("nan", "").strip(),
                "municipio": str(row.get(mapeamento.get("municipio", ""), "")).replace("nan", "").strip(),
                "valor_total": v_total,
                "area_terreno": a_terr,
                "area_construida": a_const,
                "unitario": u_calc,
                "unidade": unidade,
                "localizacao": "Rural" if unidade == "ha" else "Urbana",
                "zona_utm": str(row.get(mapeamento.get("zona_utm", ""), "")).replace("nan", "").strip(),
                "coord_e": limpar_sufixo_coord(str(row.get(mapeamento.get("coord_e", ""), "")).replace("nan", "")),
                "coord_s": limpar_sufixo_coord(str(row.get(mapeamento.get("coord_s", ""), "")).replace("nan", "")),
                "data": str(row.get(mapeamento.get("data", data_hoje_br()))).replace("nan", "").strip(),
                "link": str(row.get(mapeamento.get("link", ""), "")).replace("nan", "").strip(),
                "foto1": "",
                "foto2": "",
                "variaveis_extras": {}
            }

            # Casamento exato (normalizado) entre variáveis configuradas e colunas da planilha
            for v_cfg in self.variaveis_config:
                chave_alvo = normalizar_chave_var(v_cfg["nome"])
                for col_planilha in df.columns:
                    if normalizar_chave_var(col_planilha) == chave_alvo:
                        val_cel = row.get(col_planilha, "")
                        registro["variaveis_extras"][v_cfg["nome"]] = "" if pd.isna(val_cel) else str(val_cel).strip()
                        break

            apt = registro["variaveis_extras"].get("Aptidão", "")
            sit = registro["variaveis_extras"].get("Situação") or registro["variaveis_extras"].get("Acesso", "")
            nota_auto = calcular_nota_agronomica(apt, sit)
            if nota_auto and not registro["variaveis_extras"].get("Nota Agronômica"):
                registro["variaveis_extras"]["Nota Agronômica"] = nota_auto

            novos_dados.append(registro)

        if novos_dados:
            self.dados_pesquisas.extend(novos_dados)
            self._recarregar_grid()
            messagebox.showinfo("Importação Concluída", f"{len(novos_dados)} pesquisas importadas de {origem}!")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado válido encontrado para importação.")

    def _preparar_imagem_compativel(self, stream_ou_caminho):
        """Converte formatos como .webp para PNG em memória, garantindo compatibilidade total com python-docx."""
        try:
            with Image.open(stream_ou_caminho) as img:
                formato = (img.format or "").upper()
                if formato in ("JPEG", "JPG", "PNG", "BMP", "GIF", "TIFF"):
                    if isinstance(stream_ou_caminho, io.BytesIO):
                        stream_ou_caminho.seek(0)
                    return stream_ou_caminho
                buffer_png = io.BytesIO()
                img.convert("RGB").save(buffer_png, format="PNG")
                buffer_png.seek(0)
                return buffer_png
        except Exception as e:
            logging.warning(f"Falha ao validar/converter imagem com Pillow: {e}")
            return None

    def _baixar_imagem(self, caminho_ou_url):
        if not caminho_ou_url:
            return None
        caminho_ou_url = normalizar_url_gdrive(caminho_ou_url)
        try:
            if caminho_ou_url.startswith("http://") or caminho_ou_url.startswith("https://"):
                resp = requests.get(caminho_ou_url, timeout=12)
                content_type = resp.headers.get("Content-Type", "").lower()
                if resp.status_code == 200 and content_type.startswith("image/"):
                    return self._preparar_imagem_compativel(io.BytesIO(resp.content))
                logging.warning(f"URL não retornou uma imagem pública válida (Content-Type: {content_type}): {caminho_ou_url}")
            elif os.path.exists(caminho_ou_url):
                return self._preparar_imagem_compativel(caminho_ou_url)
        except Exception as e:
            logging.warning(f"Erro ao carregar imagem '{caminho_ou_url}': {e}")
            return None
        return None

    def _exportar_excel(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportar.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilha Excel", "*.xlsx")])
        if not caminho:
            return

        linhas_export = []
        for d in self.dados_pesquisas:
            d_id = d.get("dado_id") or d.get("D.", 1)
            utilizado = d.get("utilizado", True)
            un = d.get("unidade") or d.get("Unidade", "m²")
            v_total = d.get("valor_total") if d.get("valor_total") is not None else d.get("Valor Total (R$)", 0.0)
            a_total = d.get("area_terreno") if d.get("area_terreno") is not None else (d.get(f"Área Terreno ({un})") or 0.0)
            a_const = d.get("area_construida") if d.get("area_construida") is not None else d.get("Área Construída (m²)", 0.0)
            u_unit = d.get("unitario") if d.get("unitario") is not None else (d.get(f"Unitário (R$/{un})") or 0.0)

            linha = {
                "D.": d_id,
                "Utilizado no Modelo": "Sim" if utilizado else "Não utilizado no modelo",
                "Informante": d.get("informante") or d.get("Informante", ""),
                "Telefone": d.get("telefone") or d.get("Telefone", ""),
                "Endereço": d.get("endereco") or d.get("Endereço", ""),
                "Bairro": d.get("bairro") or d.get("Bairro", ""),
                "Município": d.get("municipio") or d.get("Município", ""),
                "Valor Total (R$)": v_total,
                f"Área Terreno ({un})": a_total,
                "Área Construída (m²)": a_const,
            }
            extras = d.get("variaveis_extras") or d.get("VariaveisExtras", {})
            for v_cfg in self.variaveis_config:
                k = v_cfg["nome"]
                linha[k] = extras.get(k, "")

            linha[f"Unitário (R$/{un})"] = u_unit
            linha["Zona UTM"] = d.get("zona_utm") or d.get("Zona UTM", "")
            linha["Coord. E (m)"] = d.get("coord_e") or d.get("Coord. E (m)", "")
            linha["Coord. S (m)"] = d.get("coord_s") or d.get("Coord. S (m)", "")
            linha["Localização"] = d.get("localizacao") or d.get("Localização", "Urbana")
            linha["Data"] = d.get("data") or d.get("Data", "")
            linha["Link"] = d.get("link") or d.get("Link", "")
            linhas_export.append(linha)

        df = pd.DataFrame(linhas_export)
        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Ficha de Pesquisa", index=False)

        messagebox.showinfo("Sucesso", "Planilha para SISDEA exportada com sucesso!")

    def _iniciar_exportacao_word_thread(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportar.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        if hasattr(self, "progress_bar"):
            self.progress_bar.pack(fill="x", padx=12, pady=4)
            self.progress_bar["value"] = 0

        snapshot_dados = copy.deepcopy(self.dados_pesquisas)
        snapshot_vars = copy.deepcopy(self.variaveis_config)
        modelo_snap = self.modelo_ativo

        thread = threading.Thread(
            target=self._processar_geracao_word,
            args=(caminho, snapshot_dados, snapshot_vars, modelo_snap),
            daemon=True
        )
        thread.start()

    def _processar_geracao_word(self, caminho, dados_lista, variaveis_lista, modelo_selecionado):
        try:
            doc = Document()
            total_dados = len(dados_lista)

            caminho_logo_enprol = os.path.join(BASE_DIR, "logo_enprol.png")
            caminho_logo_copasa = os.path.join(BASE_DIR, "logo_copasa.png")

            section = doc.sections[0]
            if modelo_selecionado == "COPASA":
                section.top_margin = Inches(0.39375)
                section.bottom_margin = Inches(0.39375)
                section.left_margin = Inches(1.18125)
                section.right_margin = Inches(1.18125)

                sectPr = section._sectPr
                for child in list(sectPr):
                    if child.tag.endswith("pgBorders"):
                        sectPr.remove(child)
                pgBrd = parse_xml(
                    f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
                    f'<w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/>'
                    f'<w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/>'
                    f'<w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/>'
                    f'<w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/>'
                    f'</w:pgBorders>'
                )
                pgMar = sectPr.find(qn("w:pgMar"))
                if pgMar is not None:
                    pgMar.addnext(pgBrd)
                else:
                    sectPr.append(pgBrd)

                hdr = section.header
                t_logos = hdr.add_table(rows=1, cols=2, width=Inches(5.90))
                t_logos.alignment = WD_TABLE_ALIGNMENT.CENTER
                c_l1, c_l2 = t_logos.cell(0, 0), t_logos.cell(0, 1)
                c_l1.width = Inches(2.95)
                c_l2.width = Inches(2.95)

                p_l1 = c_l1.paragraphs[0]
                p_l1.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if os.path.exists(caminho_logo_enprol):
                    try:
                        p_l1.add_run().add_picture(caminho_logo_enprol, height=Inches(0.42))
                    except Exception:
                        r = p_l1.add_run("ENPROL")
                        r.bold = True
                        r.font.name = "Arial"
                        r.font.size = Pt(13)
                        r.font.color.rgb = RGBColor(28, 93, 153)
                else:
                    r = p_l1.add_run("ENPROL")
                    r.bold = True
                    r.font.name = "Arial"
                    r.font.size = Pt(13)
                    r.font.color.rgb = RGBColor(28, 93, 153)

                p_l2 = c_l2.paragraphs[0]
                p_l2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if os.path.exists(caminho_logo_copasa):
                    try:
                        p_l2.add_run().add_picture(caminho_logo_copasa, height=Inches(0.42))
                    except Exception:
                        r2 = p_l2.add_run("copasa")
                        r2.bold = True
                        r2.font.name = "Arial"
                        r2.font.size = Pt(13)
                        r2.font.color.rgb = RGBColor(50, 110, 180)
                else:
                    r2 = p_l2.add_run("copasa")
                    r2.bold = True
                    r2.font.name = "Arial"
                    r2.font.size = Pt(13)
                    r2.font.color.rgb = RGBColor(50, 110, 180)

                t_faixa = hdr.add_table(rows=1, cols=3, width=Inches(5.90))
                t_faixa.alignment = WD_TABLE_ALIGNMENT.CENTER
                t_faixa.autofit = False

                tblPr_f = t_faixa._tbl.tblPr
                tblCellMar = parse_xml(
                    f'<w:tblCellMar {nsdecls("w")}>'
                    f'<w:top w:w="0" w:type="dxa"/>'
                    f'<w:left w:w="108" w:type="dxa"/>'
                    f'<w:bottom w:w="0" w:type="dxa"/>'
                    f'<w:right w:w="108" w:type="dxa"/>'
                    f'</w:tblCellMar>'
                )
                inserir_ordenado(tblPr_f, tblCellMar, SUC_CMAR)

                c_f0, c_f1, c_f2 = t_faixa.cell(0, 0), t_faixa.cell(0, 1), t_faixa.cell(0, 2)
                c_f0.width = Inches(1.77)
                c_f1.width = Inches(2.16)
                c_f2.width = Inches(1.97)

                for c_f in (c_f0, c_f1, c_f2):
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="4472C4"/>')
                    inserir_ordenado(c_f._tc.get_or_add_tcPr(), shd, SUC_TCPR[1:])

                pf0 = c_f0.paragraphs[0]
                pf0.paragraph_format.space_before = Pt(0)
                pf0.paragraph_format.space_after = Pt(0)
                pf0.paragraph_format.line_spacing = 1.0
                rf0 = pf0.add_run("CONTRATO: COPASA I")
                rf0.bold = True
                rf0.font.name = "Arial"
                rf0.font.size = Pt(10)
                rf0.font.color.rgb = RGBColor(255, 255, 255)

                pf1 = c_f1.paragraphs[0]
                pf1.paragraph_format.space_before = Pt(0)
                pf1.paragraph_format.space_after = Pt(0)
                pf1.paragraph_format.line_spacing = 1.0

                pf2 = c_f2.paragraphs[0]
                pf2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                pf2.paragraph_format.space_before = Pt(0)
                pf2.paragraph_format.space_after = Pt(0)
                pf2.paragraph_format.line_spacing = 1.0
                rf2 = pf2.add_run("PESQUISA DE MERCADO")
                rf2.bold = True
                rf2.font.name = "Arial"
                rf2.font.size = Pt(10)
                rf2.font.color.rgb = RGBColor(255, 255, 255)

                p_sp_hdr = hdr.add_paragraph()
                p_sp_hdr.paragraph_format.space_before = Pt(0)
                p_sp_hdr.paragraph_format.space_after = Pt(4)

            else:
                section.top_margin = Inches(0.35)
                section.bottom_margin = Inches(0.35)
                section.left_margin = Inches(0.55)
                section.right_margin = Inches(0.55)

            for i in range(0, total_dados, 2):
                if i > 0:
                    doc.add_page_break()

                if modelo_selecionado == "PADRAO":
                    if i == 0:
                        p_tit = doc.add_paragraph()
                        p_tit.paragraph_format.space_before = Pt(0)
                        p_tit.paragraph_format.space_after = Pt(2)
                        r_tit = p_tit.add_run("PESQUISA DE MERCADO - MEMÓRIA DE CÁLCULO")
                        r_tit.bold = True
                        r_tit.font.name = "Arial"
                        r_tit.font.size = Pt(10)

                        p_bar = doc.add_paragraph()
                        p_bar.paragraph_format.space_before = Pt(0)
                        p_bar.paragraph_format.space_after = Pt(4)
                        pBrd = parse_xml(
                            f'<w:pBdr {nsdecls("w")}>'
                            f'<w:bottom w:val="single" w:sz="12" w:space="1" w:color="245D8C"/>'
                            f'</w:pBdr>'
                        )
                        inserir_ordenado(p_bar._p.get_or_add_pPr(), pBrd, SUC_PPR)
                    else:
                        p_sp = doc.add_paragraph()
                        p_sp.paragraph_format.space_before = Pt(0)
                        p_sp.paragraph_format.space_after = Pt(4)

                tabela = doc.add_table(rows=0, cols=2)
                tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
                tabela.autofit = False

                if modelo_selecionado == "PADRAO":
                    tblPr_p = tabela._tbl.tblPr
                    borders_p = parse_xml(
                        f'<w:tblBorders {nsdecls("w")}>'
                        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        f'<w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        f'<w:insideV w:val="none"/>'
                        f'</w:tblBorders>'
                    )
                    inserir_ordenado(tblPr_p, borders_p, SUC_TBLPR)
                else:
                    tabela.style = 'Table Grid'

                lote = dados_lista[i:i+2]
                for dado in lote:
                    d_id = dado.get("dado_id") or dado.get("D.", 1)
                    utilizado = dado.get("utilizado", True)
                    sufixo_nao_utilizado = "" if utilizado else " (Não utilizado no modelo)"
                    un = dado.get("unidade") or dado.get("Unidade", "m²")

                    col0_w = Inches(2.95) if modelo_selecionado == "COPASA" else Inches(3.70)
                    col1_w = Inches(3.42) if modelo_selecionado == "COPASA" else Inches(3.70)

                    r_dados = tabela.add_row()
                    c_dados, c_foto = r_dados.cells[0], r_dados.cells[1]
                    c_dados.width = col0_w
                    c_foto.width = col1_w
                    c_dados.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    c_foto.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    p_dados = c_dados.paragraphs[0]
                    p_dados.paragraph_format.line_spacing = 1.08
                    p_dados.paragraph_format.space_before = Pt(2)
                    p_dados.paragraph_format.space_after = Pt(0)

                    def add_f_line(p, label, val):
                        r1 = p.add_run(label)
                        r1.bold = True
                        r1.font.name = "Arial"
                        r1.font.size = Pt(10)
                        r2 = p.add_run(f" {val}\n")
                        r2.font.name = "Arial"
                        r2.font.size = Pt(10)

                    if modelo_selecionado == "PADRAO":
                        p_dados.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_top = p_dados.add_run(f"Pesquisa – {d_id:02d}{sufixo_nao_utilizado}\n\n")
                        r_top.bold = True
                        r_top.font.name = "Arial"
                        r_top.font.size = Pt(10)

                        p_corpo = c_dados.add_paragraph()
                        p_corpo.paragraph_format.line_spacing = 1.08
                        p_corpo.paragraph_format.space_before = Pt(0)
                        p_corpo.paragraph_format.space_after = Pt(0)

                        add_f_line(p_corpo, "Logradouro:", dado.get("endereco") or dado.get("Endereço", ""))
                        bairro_val = dado.get("bairro") or dado.get("Bairro", "")
                        if bairro_val:
                            add_f_line(p_corpo, "Bairro:", bairro_val)
                        add_f_line(p_corpo, "Município:", dado.get("municipio") or dado.get("Município", ""))
                        tel_val = dado.get("telefone") or dado.get("Telefone", "")
                        inf_val = dado.get("informante") or dado.get("Informante", "")
                        add_f_line(p_corpo, "Contato:", f"{tel_val} - {inf_val}")
                        add_f_line(p_corpo, "Link:", dado.get("link") or dado.get("Link", ""))
                        p_corpo.add_run("\n")

                        a_terr = dado.get("area_terreno") if dado.get("area_terreno") is not None else (dado.get(f"Área Terreno ({un})") or 0.0)
                        casas = 4 if un == "ha" else 2
                        add_f_line(p_corpo, "Área Terreno:", f"{formatar_numero_br(a_terr, casas)} {un}")

                        a_const = dado.get("area_construida") if dado.get("area_construida") is not None else dado.get("Área Construída (m²)", 0.0)
                        add_f_line(p_corpo, "Área Construída:", f"{formatar_numero_br(a_const, 2)} m²")

                        v_tot = dado.get("valor_total") if dado.get("valor_total") is not None else dado.get("Valor Total (R$)", 0.0)
                        add_f_line(p_corpo, "Valor da Oferta:", f"R$ {formatar_moeda_br(v_tot)}")

                        u_val = dado.get("unitario") if dado.get("unitario") is not None else (dado.get(f"Unitário (R$/{un})") or 0.0)
                        add_f_line(p_corpo, f"Valor Unitário/{un}:", f"R$ {formatar_moeda_br(u_val)}/{un}")
                        p_corpo.add_run("\n")

                        extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras", {})
                        for v_cfg in variaveis_lista:
                            v_nome = v_cfg["nome"]
                            v_val = extras.get(v_nome, dado.get(v_nome, ""))
                            add_f_line(p_corpo, f"{v_nome}:", v_val)

                        zona_val = dado.get("zona_utm") or dado.get("Zona UTM", "")
                        zona_str = f"{str(zona_val).strip()} " if zona_val else ""
                        coord_e = limpar_sufixo_coord(dado.get("coord_e") or dado.get("Coord. E (m)", ""))
                        coord_s = limpar_sufixo_coord(dado.get("coord_s") or dado.get("Coord. S (m)", ""))
                        coord_texto = f"{zona_str}{coord_e} m E / {coord_s} m S" if coord_e or coord_s else ""

                        add_f_line(p_corpo, "Coordenadas UTM:", coord_texto)
                        loc_val = dado.get("localizacao") or dado.get("Localização", "Urbana")
                        add_f_line(p_corpo, "Localização:", loc_val)
                        p_corpo.add_run("\n")
                        add_f_line(p_corpo, "Data:", dado.get("data") or dado.get("Data", data_hoje_br()))

                        inserir_ordenado(
                            c_dados._tc.get_or_add_tcPr(),
                            parse_xml(f'<w:tcBorders {nsdecls("w")}><w:right w:val="nil"/></w:tcBorders>'),
                            SUC_TCPR
                        )
                        inserir_ordenado(
                            c_foto._tc.get_or_add_tcPr(),
                            parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="nil"/></w:tcBorders>'),
                            SUC_TCPR
                        )

                    else:
                        add_f_line(p_dados, "Logradouro:", dado.get("endereco") or dado.get("Endereço", ""))
                        add_f_line(p_dados, "Bairro:", dado.get("bairro") or dado.get("Bairro", ""))
                        add_f_line(p_dados, "Município:", dado.get("municipio") or dado.get("Município", ""))
                        p_dados.add_run("\n")

                        extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras", {})
                        info_val = extras.get("Informações", dado.get("Informações", ""))
                        add_f_line(p_dados, "Informações:", info_val)
                        p_dados.add_run("\n")

                        a_terr = dado.get("area_terreno") if dado.get("area_terreno") is not None else (dado.get("Área Terreno (m²)") or 0.0)
                        add_f_line(p_dados, "Área Terreno:", f"{formatar_numero_br(a_terr, 2)} m²")

                        a_const = dado.get("area_construida") if dado.get("area_construida") is not None else dado.get("Área Construída (m²)", 0.0)
                        add_f_line(p_dados, "Área Construída:", f"{formatar_numero_br(a_const, 2)} m²")

                        add_f_line(p_dados, "Frente:", extras.get("Frente", dado.get("Frente", "Não informado")))
                        p_dados.add_run("\n")
                        add_f_line(p_dados, "Via de acesso:", extras.get("Via de acesso", dado.get("Via de acesso", "pavimentada")))

                        zona_val = dado.get("zona_utm") or dado.get("Zona UTM", "")
                        zona_str = f"{str(zona_val).strip()} " if zona_val else ""
                        coord_e = limpar_sufixo_coord(dado.get("coord_e") or dado.get("Coord. E (m)", ""))
                        coord_s = limpar_sufixo_coord(dado.get("coord_s") or dado.get("Coord. S (m)", ""))
                        coord_texto = f"{zona_str}{coord_e} m E / {coord_s} m S" if coord_e or coord_s else ""
                        add_f_line(p_dados, "Coordenadas UTM:", coord_texto)
                        p_dados.add_run("\n")

                        u_val = dado.get("unitario") if dado.get("unitario") is not None else (dado.get("Unitário (R$/m²)") or 0.0)
                        add_f_line(p_dados, "Valor Unitário:", f"R$/m² {formatar_numero_br(u_val, 2)}")

                        v_tot = dado.get("valor_total") if dado.get("valor_total") is not None else dado.get("Valor Total (R$)", 0.0)
                        add_f_line(p_dados, "Valor Total:", f"R$ {formatar_moeda_br(v_tot)}")
                        p_dados.add_run("\n")
                        add_f_line(p_dados, "Data:", dado.get("data") or dado.get("Data", data_hoje_br()))

                    img1 = self._baixar_imagem(dado.get("foto1") or dado.get("Foto1"))
                    img2 = self._baixar_imagem(dado.get("foto2") or dado.get("Foto2"))

                    p_foto = c_foto.paragraphs[0]
                    p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    largura_foto = Inches(3.30) if modelo_selecionado == "COPASA" else Inches(2.75)

                    if img1 and img2:
                        try:
                            p_foto.add_run().add_picture(img1, width=largura_foto)
                        except Exception:
                            p_foto.add_run("[ Erro na Foto 1 ]\n")

                        p_foto2 = c_foto.add_paragraph()
                        p_foto2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foto2.paragraph_format.space_before = Pt(2)
                        try:
                            p_foto2.add_run().add_picture(img2, width=largura_foto)
                        except Exception:
                            p_foto2.add_run("[ Erro na Foto 2 ]")

                    elif img1 or img2:
                        img_unica = img1 if img1 else img2
                        try:
                            p_foto.add_run().add_picture(img_unica, width=largura_foto)
                        except Exception:
                            p_foto.add_run("[ Erro ao carregar imagem ]")
                    else:
                        r_vazio = p_foto.add_run("[ Sem fotos anexadas ]")
                        r_vazio.font.name = "Arial"
                        r_vazio.font.size = Pt(10)

                    if modelo_selecionado == "COPASA":
                        inserir_ordenado(
                            c_dados._tc.get_or_add_tcPr(),
                            parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'),
                            SUC_TCPR
                        )
                        inserir_ordenado(
                            c_foto._tc.get_or_add_tcPr(),
                            parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="nil"/><w:bottom w:val="nil"/></w:tcBorders>'),
                            SUC_TCPR
                        )

                        r_lbl = tabela.add_row()
                        c_m = r_lbl.cells[0].merge(r_lbl.cells[1])
                        c_m.width = Inches(6.37)

                        inserir_ordenado(
                            c_m._tc.get_or_add_tcPr(),
                            parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="nil"/></w:tcBorders>'),
                            SUC_TCPR
                        )

                        p_lbl = c_m.paragraphs[0]
                        p_lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_lbl.paragraph_format.space_before = Pt(4)
                        p_lbl.paragraph_format.space_after = Pt(4)
                        r_num = p_lbl.add_run(f"Pesquisa {d_id:02d}{sufixo_nao_utilizado}")
                        r_num.font.name = "Arial"
                        r_num.font.size = Pt(10)

                    progresso = int(((i + 1) / total_dados) * 100)
                    self.fila_eventos_ui.put(("progresso", progresso))

            doc.save(caminho)
            self.fila_eventos_ui.put(("concluido", None))
        except Exception as e:
            self.fila_eventos_ui.put(("erro", e))

    def _atualizar_progresso(self, valor):
        if hasattr(self, "progress_bar"):
            self.progress_bar["value"] = valor

    def _concluir_exportacao(self):
        if hasattr(self, "progress_bar"):
            self.progress_bar.pack_forget()
        messagebox.showinfo("Sucesso", "Fichas de Pesquisa no Word geradas com sucesso!")

    def _falha_exportacao(self, err):
        if hasattr(self, "progress_bar"):
            self.progress_bar.pack_forget()
        messagebox.showerror("Erro na Exportação", f"Ocorreu um erro ao gerar o documento Word: {err}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppPesquisaMercado(root)
    root.mainloop()
