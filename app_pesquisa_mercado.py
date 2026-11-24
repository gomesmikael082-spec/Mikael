import os
import io
import re
import sys
import json
import math
import queue
import shutil
import logging
import threading
import copy
import statistics
from datetime import datetime
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from PIL import Image, ImageOps, ImageTk, ImageGrab
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

try:
    from tksheet import Sheet
    TKSHEET_DISPONIVEL = True
except ImportError:
    Sheet = None
    TKSHEET_DISPONIVEL = False

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pesquisa_mercado")
os.makedirs(CONFIG_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config_variaveis.json")
PERFIS_FILE = os.path.join(CONFIG_DIR, "perfis_usuario.json")
AUTOSAVE_FILE = os.path.join(CONFIG_DIR, "autosave_projeto.json")
ANEXOS_TEMP_DIR = os.path.join(CONFIG_DIR, "anexos_temp")
os.makedirs(ANEXOS_TEMP_DIR, exist_ok=True)

# Compatibilidade Pillow < 9.1 e >= 9.1
LANCZOS_FILTER = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.BICUBIC)

# --- ORDEM DO SCHEMA OPENXML (WORD) ---
SUC_TCPR  = ["w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark"]
SUC_TBLPR = ["w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"]
SUC_CMAR  = ["w:tblLook"]
SUC_PPR   = ["w:shd", "w:tabs", "w:suppressAutoHyphens", "w:spacing", "w:ind", "w:jc", "w:rPr"]


def inserir_ordenado(pai, el, sucessores):
    for tag in sucessores:
        achado = pai.find(qn(tag))
        if achado is not None:
            achado.addprevious(el)
            return
    pai.append(el)


# --- TABELA 02: NOTA AGRONÔMICA (Kozma 1984 x IN RFB 1877/2019) ---
TABELA_NOTA_AGRONOMICA = {
    "I":   {"Ótima": "1,00", "Muito Boa": "0,95", "Boa": "0,90", "Desfavorável": "0,80", "Má": "0,75", "Péssima": "0,70"},
    "II":  {"Ótima": "0,80", "Muito Boa": "0,76", "Boa": "0,72", "Desfavorável": "0,64", "Má": "0,60", "Péssima": "0,56"},
    "III": {"Ótima": "0,64", "Muito Boa": "0,61", "Boa": "0,58", "Desfavorável": "0,51", "Má": "0,48", "Péssima": "0,45"},
    "IV":  {"Ótima": "0,51", "Muito Boa": "0,48", "Boa": "0,46", "Desfavorável": "0,41", "Má": "0,38", "Péssima": "0,36"},
    "V":   {"Ótima": "0,41", "Muito Boa": "0,39", "Boa": "0,37", "Desfavorável": "0,33", "Má": "0,31", "Péssima": "0,29"},
    "VI":  {"Ótima": "0,20", "Muito Boa": "0,19", "Boa": "0,18", "Desfavorável": "0,16", "Má": "0,15", "Péssima": "0,14"},
}

OPCOES_APTIDAO = [
    "",
    "I - Lavoura (Aptidão Boa)",
    "II - Lavoura (Aptidão Regular)",
    "III - Lavoura (Aptidão Restrita)",
    "IV - Pastagem Plantada",
    "V - Silvicultura ou Pastagem Natural",
    "VI - Preservação de Fauna/Flora"
]

OPCOES_SITUACAO = [
    "",
    "Ótima",
    "Muito Boa",
    "Boa",
    "Desfavorável",
    "Má",
    "Péssima"
]

DEFAULT_VARS_PADRAO = [
    {"nome": "Via", "tipo": "codigo", "opcoes": ["", "1 - Local", "2 - Coletora", "3 - Arterial"]},
    {"nome": "Uso", "tipo": "codigo", "opcoes": ["", "1 - Residencial", "2 - Comercial", "3 - Misto"]},
    {"nome": "Testada (m)", "tipo": "numero", "opcoes": []},
    {"nome": "PGV (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "PIB (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "Condição", "tipo": "codigo", "opcoes": ["", "1 - Oferta", "2 - Transação"]}
]

DEFAULT_VARS_RURAL = [
    {"nome": "Aptidão Agrícola", "tipo": "codigo", "opcoes": OPCOES_APTIDAO},
    {"nome": "Situação", "tipo": "codigo", "opcoes": OPCOES_SITUACAO},
    {"nome": "Nota Agronômica", "tipo": "numero", "opcoes": []},
    {"nome": "Benfeitoria", "tipo": "codigo", "opcoes": ["", "1 - Sem Benfeitorias", "2 - Com Benfeitorias", "3 - Produtiva"]}
]

DEFAULT_VARS_COPASA = [
    {"nome": "Frente", "tipo": "texto", "opcoes": ["Não informado"]},
    {"nome": "Via de acesso", "tipo": "codigo", "opcoes": ["", "pavimentada", "não pavimentada"]},
    {"nome": "Informações", "tipo": "texto", "opcoes": []}
]

CATALOGO_VARIAVEIS_USUAIS = [
    {"nome": "Via", "tipo": "codigo", "opcoes": ["", "1 - Local", "2 - Coletora", "3 - Arterial"]},
    {"nome": "Uso", "tipo": "codigo", "opcoes": ["", "1 - Residencial", "2 - Comercial", "3 - Misto"]},
    {"nome": "Testada (m)", "tipo": "numero", "opcoes": []},
    {"nome": "PGV (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "PIB (R$)", "tipo": "numero", "opcoes": []},
    {"nome": "Condição", "tipo": "codigo", "opcoes": ["", "1 - Oferta", "2 - Transação"]},
    {"nome": "Distância ao Centro (m)", "tipo": "numero", "opcoes": []},
    {"nome": "Topografia", "tipo": "codigo", "opcoes": ["", "1 - Plano", "2 - Aclive/Declive", "3 - Irregular"]},
    {"nome": "Padrão Construtivo", "tipo": "codigo", "opcoes": ["", "1 - Baixo", "2 - Normal", "3 - Alto"]},
    {"nome": "Aptidão Agrícola", "tipo": "codigo", "opcoes": OPCOES_APTIDAO},
    {"nome": "Situação", "tipo": "codigo", "opcoes": OPCOES_SITUACAO},
    {"nome": "Nota Agronômica", "tipo": "numero", "opcoes": []},
    {"nome": "Benfeitoria", "tipo": "codigo", "opcoes": ["", "1 - Sem Benfeitorias", "2 - Com Benfeitorias", "3 - Produtiva"]},
    {"nome": "Frente", "tipo": "texto", "opcoes": ["Não informado"]},
    {"nome": "Via de acesso", "tipo": "codigo", "opcoes": ["", "pavimentada", "não pavimentada"]},
    {"nome": "Informações", "tipo": "texto", "opcoes": []}
]

DEFAULT_CAMPOS_OPCIONAIS = {
    "telefone": True,
    "bairro": True,
    "area_construida": True,
    "link": True,
    "localizacao": True
}


def data_hoje_br():
    return datetime.now().strftime("%d/%m/%Y")


def normalizar_chave_var(nome):
    s = str(nome).strip().lower()
    return re.sub(r"\s+", " ", s)


def sanitizar_texto_celula(txt):
    if txt is None:
        return ""
    try:
        if pd.isna(txt):
            return ""
    except Exception:
        pass
    if isinstance(txt, (pd.Timestamp, datetime)):
        return txt.strftime("%d/%m/%Y")
    s = str(txt).strip()
    if s.lower() in ("nan", "nat", "none"):
        return ""
    m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})\s+00:00:00$", s)
    if m_iso:
        return f"{m_iso.group(3)}/{m_iso.group(2)}/{m_iso.group(1)}"
    return re.sub(r'[\r\n\t]+', ' ', s).strip()


def sanitizar_telefone(txt):
    s = sanitizar_texto_celula(txt)
    if re.match(r"^\d+\.0$", s):
        return s[:-2]
    return s


def calcular_nota_agronomica_flexivel(variaveis_extras):
    apt = ""
    sit = ""
    for k, v in variaveis_extras.items():
        k_low = normalizar_chave_var(k)
        if k_low.startswith("aptid"):
            apt = str(v)
        elif k_low.startswith("situa") or k_low.startswith("acess"):
            sit = str(v)
    if not apt or not sit:
        return ""
    m = re.match(r"^\s*(VI|IV|V|III|II|I)\b", apt.strip(), re.IGNORECASE)
    if not m:
        return ""
    romano = m.group(1).upper()
    sit_clean = sit.strip().lower()
    mapa_sit = TABELA_NOTA_AGRONOMICA.get(romano, {})
    for k, v in mapa_sit.items():
        if k.lower() == sit_clean:
            return v
    for k in sorted(mapa_sit.keys(), key=len, reverse=True):
        if k.lower() in sit_clean:
            return mapa_sit[k]
    return ""


def aplicar_nota_agronomica_no_dado(dado, variaveis_config):
    extras = dado.setdefault("variaveis_extras", {})
    nota_auto = calcular_nota_agronomica_flexivel(extras)
    if not nota_auto:
        return
    nome_chave_nota = None
    for v_cfg in variaveis_config:
        if normalizar_chave_var(v_cfg["nome"]).startswith("nota agro"):
            nome_chave_nota = v_cfg["nome"]
            break
    if not nome_chave_nota:
        for k in extras:
            if normalizar_chave_var(k).startswith("nota agro"):
                nome_chave_nota = k
                break
    if nome_chave_nota:
        extras[nome_chave_nota] = nota_auto


def _dms_para_decimal(graus, minutos, segundos, hemisferio):
    dec = float(graus) + float(minutos) / 60.0 + float(segundos) / 3600.0
    if hemisferio.upper() in ("S", "W", "O"):
        dec = -abs(dec)
    return dec


def converter_latlon_para_utm(texto_coord):
    if not texto_coord:
        return None
    txt = str(texto_coord).strip()
    lat, lon = None, None

    dms_matches = re.findall(
        r"(\d{1,3})\s*[°º]\s*(\d{1,2})\s*['’′]\s*(\d{1,2}(?:[.,]\d+)?)\s*(?:[\"”″])?\s*([NSWEOnwseo])",
        txt
    )
    if len(dms_matches) >= 2:
        g1, m1, s1, h1 = dms_matches[0]
        g2, m2, s2, h2 = dms_matches[1]
        lat = _dms_para_decimal(g1, m1, s1.replace(",", "."), h1)
        lon = _dms_para_decimal(g2, m2, s2.replace(",", "."), h2)
    else:
        pares = re.findall(r"[-+]?\d+[.,]\d+", txt)
        if len(pares) >= 2:
            lat = float(pares[0].replace(",", "."))
            lon = float(pares[1].replace(",", "."))

    if lat is None or lon is None or not (-90 <= lat <= 90 and -180 <= lon <= 180):
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


def converter_para_float(texto, permitir_milhar_ponto=True):
    if texto is None:
        return 0.0
    if isinstance(texto, (int, float)):
        val_f = float(texto)
        return 0.0 if math.isnan(val_f) else val_f
    s = str(texto).replace("R$", "").replace("m²", "").replace("ha", "").strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return 0.0
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        partes = s.split(".")
        if len(partes) > 2:
            s = s.replace(".", "")
        elif permitir_milhar_ponto and len(partes) == 2 and len(partes[1]) == 3 and len(partes[0]) <= 3 and int(partes[0] or 0) > 0:
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


def redimensionar_e_corrigir_imagem(caminho_ou_stream, max_largura=1200, max_altura=850):
    try:
        if isinstance(caminho_ou_stream, io.BytesIO):
            img = Image.open(caminho_ou_stream)
        else:
            if not os.path.exists(caminho_ou_stream):
                return None
            img = Image.open(caminho_ou_stream)
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.thumbnail((max_largura, max_altura), LANCZOS_FILTER)
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=88)
        buffer.seek(0)
        return buffer
    except Exception as e:
        logging.warning(f"Erro ao processar imagem: {e}")
        return None


def normalizar_url_gdrive(url):
    if not url:
        return ""
    url = url.strip()
    match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url) or re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if match:
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
    return url


def normalizar_dado(dado, variaveis_config=None):
    d_id = dado.get("dado_id") if dado.get("dado_id") is not None else dado.get("D.", 1)
    utilizado = dado.get("utilizado") if "utilizado" in dado else dado.get("Utilizado", True)
    if isinstance(utilizado, str):
        utilizado = utilizado.strip().lower() not in ("false", "0", "não", "nao", "☐", "não utilizado no modelo")

    informante = sanitizar_texto_celula(dado.get("informante") or dado.get("Informante") or "")
    telefone = sanitizar_telefone(dado.get("telefone") or dado.get("Telefone") or "")
    endereco = sanitizar_texto_celula(dado.get("endereco") or dado.get("Endereço") or "")
    bairro = sanitizar_texto_celula(dado.get("bairro") or dado.get("Bairro") or "")
    municipio = sanitizar_texto_celula(dado.get("municipio") or dado.get("Município") or "")
    unidade = sanitizar_texto_celula(dado.get("unidade") or dado.get("Unidade") or "m²") or "m²"

    valor_total = dado.get("valor_total")
    if valor_total is None:
        valor_total = dado.get("Valor Total (R$)", 0.0)
    valor_total = converter_para_float(valor_total)

    area_terreno = dado.get("area_terreno")
    if area_terreno is None:
        area_terreno = dado.get(f"Área Terreno ({unidade})") or dado.get("Área Terreno (ha)") or dado.get("Área Terreno (m²)") or 0.0
    area_terreno = converter_para_float(area_terreno, permitir_milhar_ponto=(unidade != "ha"))

    area_construida = dado.get("area_construida")
    if area_construida is None:
        area_construida = dado.get("Área Construída (m²)", 0.0)
    area_construida = converter_para_float(area_construida)

    unitario = valor_total / area_terreno if area_terreno > 0 else 0.0

    zona_utm = sanitizar_texto_celula(dado.get("zona_utm") or dado.get("Zona UTM") or "")
    coord_e = limpar_sufixo_coord(sanitizar_texto_celula(dado.get("coord_e") or dado.get("Coord. E (m)") or ""))
    coord_s = limpar_sufixo_coord(sanitizar_texto_celula(dado.get("coord_s") or dado.get("Coord. S (m)") or ""))
    data = sanitizar_texto_celula(dado.get("data") or dado.get("Data") or data_hoje_br()) or data_hoje_br()
    link = sanitizar_texto_celula(dado.get("link") or dado.get("Link") or "")
    foto1 = sanitizar_texto_celula(dado.get("foto1") or dado.get("Foto1") or "")
    foto2 = sanitizar_texto_celula(dado.get("foto2") or dado.get("Foto2") or "")
    localizacao = sanitizar_texto_celula(dado.get("localizacao") or dado.get("Localização") or ("Rural" if unidade == "ha" else "Urbana"))

    vars_extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras") or {}
    vars_sanitizadas = {k: sanitizar_texto_celula(v) for k, v in vars_extras.items()}

    reg = {
        "dado_id": int(d_id),
        "utilizado": bool(utilizado),
        "informante": informante,
        "telefone": telefone,
        "endereco": endereco,
        "bairro": bairro,
        "municipio": municipio,
        "valor_total": float(valor_total),
        "area_terreno": float(area_terreno),
        "area_construida": float(area_construida),
        "unitario": float(unitario),
        "unidade": unidade,
        "localizacao": localizacao,
        "zona_utm": zona_utm,
        "coord_e": coord_e,
        "coord_s": coord_s,
        "data": data,
        "link": link,
        "foto1": foto1,
        "foto2": foto2,
        "variaveis_extras": vars_sanitizadas
    }
    if variaveis_config:
        aplicar_nota_agronomica_no_dado(reg, variaveis_config)
    return reg


class AppPesquisaMercado:
    def __init__(self, root):
        self.root = root
        self.root.title("ENPROL - Sistema de Pesquisa de Mercado v6.2 (Planilha SISDEA)")
        self.root.geometry("1300x840")
        self.root.minsize(1060, 700)

        self._configurar_estilos()

        self.modelo_ativo = "PADRAO"
        self.unidade_padrao_projeto = "m²"
        self.dados_pesquisas = []
        self.variaveis_config = self._carregar_config_variaveis()
        self.campos_opcionais = copy.deepcopy(DEFAULT_CAMPOS_OPCIONAIS)
        self.perfis_usuario = self._carregar_perfis_usuario()
        self.caminho_projeto_atual = None
        self.alteracoes_pendentes = False

        self.celula_ativa = {"row": None, "col": None, "chave": None}
        self.painel_lateral_visivel = False
        self._img_thumb1 = None
        self._img_thumb2 = None
        self.fila_eventos_ui = queue.Queue()

        self.container_principal = ttk.Frame(self.root)
        self.container_principal.pack(fill="both", expand=True)

        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)
        self.root.after(100, self._processar_fila_ui)

        if not TKSHEET_DISPONIVEL:
            messagebox.showerror(
                "Biblioteca tksheet não encontrada",
                "A Versão 6.2 utiliza a biblioteca 'tksheet' para a grade estilo planilha SISDEA.\n\n"
                "Abra o Prompt de Comando (CMD) e execute:\n"
                "pip install tksheet\n\n"
                "Depois abra o programa novamente."
            )
            self.root.destroy()
            return

        self._exibir_tela_hub()

    def _processar_fila_ui(self):
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
        style.configure("TLabelframe.Label", background="#f4f6f9", foreground="#1f3c5b", font=("Segoe UI", 9, "bold"))
        style.configure("TLabel", background="#f4f6f9", font=("Segoe UI", 9))

        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), background="#1f3c5b", foreground="white")
        style.map("Primary.TButton", background=[("active", "#2c5480")])

        style.configure("Accent.TButton", font=("Segoe UI", 9, "bold"), background="#2b7a78", foreground="white")
        style.map("Accent.TButton", background=[("active", "#3a9895")])

        style.configure("Hub.TButton", font=("Segoe UI", 11, "bold"), padding=10, background="#1f3c5b", foreground="white")
        style.map("Hub.TButton", background=[("active", "#2d5784")])

    def _limpar_container(self):
        for widget in self.container_principal.winfo_children():
            widget.destroy()
        self.painel_lateral_visivel = False
        self.celula_ativa = {"row": None, "col": None, "chave": None}
        for attr in ("sheet", "lbl_estatisticas", "progress_bar", "lbl_info_mod", "lbl_fx", "txt_fx"):
            if hasattr(self, attr):
                delattr(self, attr)

    def _carregar_config_variaveis(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return copy.deepcopy(DEFAULT_VARS_PADRAO)

    def _salvar_config_variaveis(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.variaveis_config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("Erro ao salvar variáveis", str(e))

    def _carregar_perfis_usuario(self):
        if os.path.exists(PERFIS_FILE):
            try:
                with open(PERFIS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _salvar_perfis_usuario(self):
        try:
            with open(PERFIS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.perfis_usuario, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("Erro ao salvar perfis", str(e))

    def _autosave(self):
        if not self.dados_pesquisas:
            return
        try:
            dados_salvar = {
                "modelo": self.modelo_ativo,
                "unidade_padrao": self.unidade_padrao_projeto,
                "campos_opcionais": self.campos_opcionais,
                "variaveis_config": self.variaveis_config,
                "pesquisas": self.dados_pesquisas
            }
            with open(AUTOSAVE_FILE, "w", encoding="utf-8") as f:
                json.dump(dados_salvar, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def _ao_fechar_janela(self):
        if self.dados_pesquisas and self.alteracoes_pendentes:
            self._autosave()
            resp = messagebox.askyesnocancel("Sair do Sistema", "Existem alterações não salvas. Deseja salvar o projeto antes de sair?")
            if resp is None:
                return
            if resp is True:
                self._salvar_projeto()
        self.root.destroy()

    # --- HUB INICIAL ---
    def _exibir_tela_hub(self):
        self._limpar_container()
        self.root.config(menu="")

        frame_hub = ttk.Frame(self.container_principal, padding=40)
        frame_hub.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(frame_hub, text="SISTEMA DE PESQUISA DE MERCADO v6.2", font=("Segoe UI", 18, "bold"), foreground="#1f3c5b").pack(pady=(0, 5))
        ttk.Label(frame_hub, text="Grade Planilha SISDEA (tksheet) • Perfis Persistentes • Automação Agronômica", font=("Segoe UI", 10), foreground="#555555").pack(pady=(0, 30))

        ttk.Button(frame_hub, text="📄  Criar Novo Arquivo / Projeto", style="Hub.TButton", width=42, command=self._popup_escolha_modelo).pack(pady=7)
        ttk.Button(frame_hub, text="📂  Carregar Projeto Existente (.json)", style="Hub.TButton", width=42, command=self._abrir_projeto_hub).pack(pady=7)
        ttk.Button(frame_hub, text="📥  Importar Dados de Planilha (.xlsx)", style="Hub.TButton", width=42, command=self._importar_planilha_hub).pack(pady=7)

        if os.path.exists(AUTOSAVE_FILE):
            ttk.Button(frame_hub, text="🔄  Recuperar Sessão Anterior (Autosave)", style="Accent.TButton", width=42, command=self._carregar_autosave).pack(pady=7)

        ttk.Button(frame_hub, text="🚪  Sair do Sistema", width=42, command=self._ao_fechar_janela).pack(pady=(20, 0))

    def _carregar_autosave(self):
        try:
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f:
                conteudo = json.load(f)
            self._aplicar_conteudo_json(conteudo)
            self._exibir_tela_trabalho()
            messagebox.showinfo("Recuperado", f"{len(self.dados_pesquisas)} pesquisas recuperadas do autosave!")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível recuperar o autosave: {e}")

    def _popup_escolha_modelo(self, callback_pos_confirmacao=None):
        if self.dados_pesquisas and self.alteracoes_pendentes:
            if not messagebox.askyesno("Novo Projeto", "Criar um novo projeto descartará os dados atuais da tela. Deseja continuar?"):
                return

        janela = tk.Toplevel(self.root)
        janela.title("Selecionar Modelo e Perfil Inicial")
        janela.geometry("500x360")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Selecione o modelo de ficha e perfil de variáveis:", font=("Segoe UI", 11, "bold")).pack(pady=(16, 10))

        var_mod = tk.StringVar(value="PADRAO_URBANO")
        ttk.Radiobutton(janela, text="Modelo 1 - Padrão Urbano (Via, Uso, Testada, PGV, PIB)", value="PADRAO_URBANO", variable=var_mod).pack(anchor="w", padx=35, pady=4)
        ttk.Radiobutton(janela, text="Modelo 1 - Padrão Rural (Aptidão, Situação, Nota Agronômica, Benfeitoria)", value="PADRAO_RURAL", variable=var_mod).pack(anchor="w", padx=35, pady=4)
        ttk.Radiobutton(janela, text="Modelo 2 - Contrato COPASA (Ficha Técnica Oficial)", value="COPASA", variable=var_mod).pack(anchor="w", padx=35, pady=4)

        cb_perfil_custom = None
        if self.perfis_usuario:
            ttk.Separator(janela, orient="horizontal").pack(fill="x", padx=30, pady=10)
            ttk.Radiobutton(janela, text="Usar Perfil Personalizado Salvo:", value="CUSTOM", variable=var_mod).pack(anchor="w", padx=35, pady=2)
            cb_perfil_custom = ttk.Combobox(janela, values=list(self.perfis_usuario.keys()), width=38, state="readonly")
            cb_perfil_custom.current(0)
            cb_perfil_custom.pack(anchor="w", padx=55, pady=4)

        def confirmar():
            escolha = var_mod.get()
            self.campos_opcionais = copy.deepcopy(DEFAULT_CAMPOS_OPCIONAIS)
            if escolha == "COPASA":
                self.modelo_ativo = "COPASA"
                self.variaveis_config = copy.deepcopy(DEFAULT_VARS_COPASA)
                self.unidade_padrao_projeto = "m²"
            elif escolha == "PADRAO_RURAL":
                self.modelo_ativo = "PADRAO"
                self.variaveis_config = copy.deepcopy(DEFAULT_VARS_RURAL)
                self.unidade_padrao_projeto = "ha"
            elif escolha == "CUSTOM" and cb_perfil_custom:
                nome_p = cb_perfil_custom.get()
                p = self.perfis_usuario[nome_p]
                self.modelo_ativo = p.get("modelo", "PADRAO")
                self.variaveis_config = copy.deepcopy(p.get("variaveis_config", DEFAULT_VARS_PADRAO))
                self.unidade_padrao_projeto = p.get("unidade_padrao", "m²")
                self.campos_opcionais = copy.deepcopy(p.get("campos_opcionais", DEFAULT_CAMPOS_OPCIONAIS))
            else:
                self.modelo_ativo = "PADRAO"
                self.variaveis_config = copy.deepcopy(DEFAULT_VARS_PADRAO)
                self.unidade_padrao_projeto = "m²"

            self._salvar_config_variaveis()
            self.caminho_projeto_atual = None
            self.alteracoes_pendentes = False
            self.dados_pesquisas = []
            janela.destroy()
            self._exibir_tela_trabalho()
            if callback_pos_confirmacao:
                callback_pos_confirmacao()

        ttk.Button(janela, text="Avançar", style="Primary.TButton", command=confirmar).pack(pady=(18, 0))

    def _abrir_projeto_hub(self):
        if self._executar_abertura_json():
            self._exibir_tela_trabalho()

    def _importar_planilha_hub(self):
        self._popup_escolha_modelo(callback_pos_confirmacao=self._importar_planilha_excel)

    # --- TELA DE TRABALHO (PLANILHA TKSHEET) ---
    def _exibir_tela_trabalho(self):
        self._limpar_container()
        self._criar_menu_superior()
        self._criar_layout_planilha()
        self._reconstruir_planilha()

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
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Sair", command=self._ao_fechar_janela)
        menubar.add_cascade(label="Arquivo", menu=menu_arquivo)

        menu_config = tk.Menu(menubar, tearoff=0)
        menu_config.add_command(label="⚙ Gerenciar Variáveis e Perfis...", command=self._janela_config_variaveis)
        menu_config.add_command(label="👁 Campos Opcionais da Ficha (Bairro, Área Const...)", command=self._janela_campos_opcionais)
        menu_config.add_command(label="💾 Salvar Configuração Atual como Novo Perfil...", command=self._salvar_como_novo_perfil)
        menu_config.add_separator()
        menu_config.add_command(label="Aplicar Perfil Urbano Padrão", command=lambda: self._aplicar_perfil_rapido("URBANO"))
        menu_config.add_command(label="Aplicar Perfil Rural (Nota Agronômica)", command=lambda: self._aplicar_perfil_rapido("RURAL"))
        menu_config.add_command(label="Aplicar Perfil COPASA", command=lambda: self._aplicar_perfil_rapido("COPASA"))
        menu_config.add_separator()
        menu_config.add_command(label="📊 Ver Tabela 02 - Nota Agronômica", command=self._mostrar_tabela_nota_agronomica)
        menubar.add_cascade(label="Variáveis e Perfis", menu=menu_config)

        self.root.config(menu=menubar)

    def _criar_layout_planilha(self):
        # 1. Toolbar Compacta Superior
        toolbar = ttk.Frame(self.container_principal, padding=(8, 4))
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="➕ Nova Linha", style="Primary.TButton", command=self._adicionar_nova_linha).pack(side="left", padx=2)
        ttk.Button(toolbar, text="📥 Importar Excel", command=self._importar_planilha_excel).pack(side="left", padx=2)
        ttk.Button(toolbar, text="💾 Salvar", command=self._salvar_projeto).pack(side="left", padx=2)
        ttk.Button(toolbar, text="⚙ Variáveis", command=self._janela_config_variaveis).pack(side="left", padx=2)
        ttk.Button(toolbar, text="👁 Campos", command=self._janela_campos_opcionais).pack(side="left", padx=2)
        ttk.Button(toolbar, text="🌐 Lat/Long ➔ UTM", command=self._janela_conversor_utm).pack(side="left", padx=2)

        ttk.Button(toolbar, text="📄 Gerar Word", style="Primary.TButton", command=self._iniciar_exportacao_word_thread).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📊 Exportar Excel", style="Accent.TButton", command=self._exportar_excel).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📷 Mídia / Fotos", command=self._alternar_painel_lateral).pack(side="right", padx=6)

        mod_label = "Modelo 1: Padrão" if self.modelo_ativo == "PADRAO" else "Modelo 2: COPASA"
        self.lbl_info_mod = ttk.Label(toolbar, text=f"  |  {mod_label} ({self.unidade_padrao_projeto})", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b")
        self.lbl_info_mod.pack(side="left", padx=6)

        # 2. Barra de Fórmula (fx) Estilo SISDEA
        barra_fx = ttk.Frame(self.container_principal, padding=(8, 2))
        barra_fx.pack(fill="x")

        self.lbl_fx = tk.Label(barra_fx, text="fx  [ Célula ]", width=24, anchor="w", bg="#e1e7ef", fg="#1f3c5b", font=("Segoe UI", 9, "bold"), relief="groove", padx=6)
        self.lbl_fx.pack(side="left")

        self.txt_fx = ttk.Entry(barra_fx, font=("Segoe UI", 9))
        self.txt_fx.pack(side="left", fill="x", expand=True, padx=4)
        self.txt_fx.bind("<Return>", lambda e: self._aplicar_barra_fx())

        # 3. Área Central: Planilha tksheet + Painel Lateral Recolhível
        self.pane_principal = tk.PanedWindow(self.container_principal, orient="horizontal", sashrelief="raised", bg="#d0d7de")
        self.pane_principal.pack(fill="both", expand=True, padx=8, pady=2)

        self.frame_sheet = ttk.Frame(self.pane_principal)
        self.pane_principal.add(self.frame_sheet, stretch="always")

        self.sheet = Sheet(
            self.frame_sheet,
            data=[],
            headers=[],
            font=("Segoe UI", 9, "normal"),
            header_font=("Segoe UI", 9, "bold"),
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            empty_horizontal=0,
            empty_vertical=0
        )
        self.sheet.enable_bindings(
            "single_select", "row_select", "column_select", "drag_select",
            "column_width_resize", "double_click_column_resize", "row_height_resize",
            "arrowkeys", "right_click_popup_menu", "rc_select",
            "copy", "cut", "paste", "delete", "undo", "edit_cell"
        )
        self.sheet.pack(fill="both", expand=True)

        self.sheet.popup_menu_add_command("➕ Inserir Nova Linha Abaixo", self._adicionar_nova_linha)
        self.sheet.popup_menu_add_command("📋 Duplicar Linha Selecionada", self._duplicar_linha_selecionada)
        self.sheet.popup_menu_add_command("☑ Alternar Utilizado / Não Utilizado", self._alternar_utilizado_selecionado)
        self.sheet.popup_menu_add_command("⬇ Preencher Para Baixo (Ctrl+D)", self._preencher_para_baixo)
        self.sheet.popup_menu_add_command("🗑 Excluir Linha(s) Selecionada(s)", self._excluir_linhas_selecionadas)

        self.sheet.extra_bindings("cell_select", self._ao_selecionar_celula)
        self.sheet.extra_bindings("end_edit_cell", self._ao_terminar_edicao_celula)
        self.sheet.extra_bindings("end_paste", self._ao_colar_planilha)
        self.sheet.extra_bindings("end_delete", self._ao_deletar_conteudo)
        self.root.bind("<Control-d>", lambda e: self._preencher_para_baixo())

        self.frame_painel_lateral = ttk.LabelFrame(self.pane_principal, text=" Mídia, Fotos & Coordenadas ", padding=8)
        self._construir_painel_lateral()

        # 4. Rodapé Estatístico Único + Progresso
        self.progress_bar = ttk.Progressbar(self.container_principal, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=8, pady=1)
        self.progress_bar.pack_forget()

        frame_rodape = tk.Frame(self.container_principal, bg="#e1e7ef", bd=1, relief="solid")
        frame_rodape.pack(fill="x", padx=8, pady=(2, 6))
        self.lbl_estatisticas = tk.Label(
            frame_rodape,
            text="📊 0 amostras ativas",
            bg="#e1e7ef",
            fg="#1f3c5b",
            font=("Segoe UI", 9, "bold"),
            pady=4
        )
        self.lbl_estatisticas.pack(anchor="w", padx=10)

    def _obter_esquema_colunas(self):
        cols = [
            ("dado_id", "D.", 42),
            ("utilizado", "☑", 42),
            ("informante", "Informante", 140),
        ]
        if self.campos_opcionais.get("telefone", True):
            cols.append(("telefone", "Telefone", 110))
        cols.append(("endereco", "Endereço / Logradouro", 180))
        if self.campos_opcionais.get("bairro", True):
            cols.append(("bairro", "Bairro", 115))
        cols.extend([
            ("municipio", "Município", 115),
            ("valor_total", "Valor Total (R$)", 115),
            ("area_terreno", "Área Terreno", 95),
            ("unidade", "Unid.", 52),
        ])
        if self.campos_opcionais.get("area_construida", True):
            cols.append(("area_construida", "Área Const. (m²)", 105))

        for v_cfg in self.variaveis_config:
            cols.append((f"var_{v_cfg['nome']}", v_cfg["nome"], 115))

        cols.extend([
            ("unitario", "Unitário (R$/un)", 115),
            ("zona_utm", "Zona UTM", 75),
            ("coord_e", "Coord. E (m)", 100),
            ("coord_s", "Coord. S (m)", 100),
            ("data", "Data", 88),
        ])
        if self.campos_opcionais.get("link", True):
            cols.append(("link", "Link", 140))
        cols.append(("status", "Mídia", 65))
        return cols

    def _montar_linha_sheet(self, dado):
        un = dado.get("unidade", self.unidade_padrao_projeto)
        casas = 4 if un == "ha" else 2
        v_tot = converter_para_float(dado.get("valor_total", 0.0))
        a_ter = converter_para_float(dado.get("area_terreno", 0.0), permitir_milhar_ponto=(un != "ha"))
        u_val = v_tot / a_ter if a_ter > 0 else 0.0
        dado["unitario"] = u_val

        tem_foto = bool(dado.get("foto1") or dado.get("foto2"))
        tem_coord = bool(dado.get("coord_e") and dado.get("coord_s"))
        status_txt = f"{'📷' if tem_foto else '—'} {'📍' if tem_coord else '—'}"

        extras = dado.get("variaveis_extras", {})
        linha = []
        for chave, _, _ in self.esquema_colunas:
            if chave == "dado_id":
                linha.append(dado["dado_id"])
            elif chave == "utilizado":
                linha.append(bool(dado.get("utilizado", True)))
            elif chave == "valor_total":
                linha.append(formatar_moeda_br(v_tot) if v_tot else "")
            elif chave == "area_terreno":
                linha.append(formatar_numero_br(a_ter, casas) if a_ter else "")
            elif chave == "area_construida":
                ac = converter_para_float(dado.get("area_construida", 0.0))
                linha.append(formatar_numero_br(ac, 2) if ac else "")
            elif chave == "unitario":
                linha.append(formatar_moeda_br(u_val) if u_val else "")
            elif chave == "status":
                linha.append(status_txt)
            elif chave.startswith("var_"):
                linha.append(extras.get(chave[4:], ""))
            else:
                linha.append(dado.get(chave, ""))
        return linha

    def _reconstruir_planilha(self):
        if not hasattr(self, "sheet"):
            return
        self.esquema_colunas = self._obter_esquema_colunas()
        headers = [c[1] for c in self.esquema_colunas]
        matriz = [self._montar_linha_sheet(d) for d in self.dados_pesquisas]

        self.sheet.set_sheet_data(matriz, reset_col_positions=True, reset_row_positions=True, redraw=False)
        self.sheet.headers(headers)

        for c_idx, (chave, _, larg) in enumerate(self.esquema_colunas):
            self.sheet.column_width(column=c_idx, width=larg)
            if chave == "utilizado" and matriz:
                self.sheet.checkbox(self.sheet.span(f"B1:B{len(matriz)}"), checked=True)
                for r_idx, d in enumerate(self.dados_pesquisas):
                    self.sheet.set_cell_data(r_idx, c_idx, bool(d.get("utilizado", True)), redraw=False)
            elif chave == "unidade" and matriz:
                self.sheet.dropdown(self.sheet.span(0, c_idx, len(matriz), c_idx + 1), values=["m²", "ha"], set_value=None)
            elif chave in ("dado_id", "unitario", "status"):
                self.sheet.readonly_columns(columns=[c_idx], readonly=True)
                self.sheet.highlight_columns(columns=[c_idx], bg="#eef2f6", fg="#334e68")
            elif chave.startswith("var_"):
                nome_v = chave[4:]
                v_cfg = next((v for v in self.variaveis_config if v["nome"] == nome_v), None)
                if v_cfg and v_cfg.get("tipo") == "codigo" and v_cfg.get("opcoes") and matriz:
                    opcoes_limpas = [""] + [o for o in v_cfg["opcoes"] if o != ""]
                    self.sheet.dropdown(self.sheet.span(0, c_idx, len(matriz), c_idx + 1), values=opcoes_limpas, set_value=None)

        self._atualizar_estatisticas_e_destaques()
        self.sheet.redraw()

    def _atualizar_estatisticas_e_destaques(self):
        if not hasattr(self, "lbl_estatisticas"):
            return
        unitarios_ativos = []
        unidades_vistas = set()
        for r_idx, d in enumerate(self.dados_pesquisas):
            ativo = bool(d.get("utilizado", True))
            u_val = d.get("unitario", 0.0)
            if ativo and u_val > 0:
                unitarios_ativos.append(u_val)
                unidades_vistas.add(d.get("unidade", self.unidade_padrao_projeto))

        total_geral = len(self.dados_pesquisas)
        n_ativos = len(unitarios_ativos)

        if total_geral > 0:
            self.sheet.dehighlight_rows(rows=list(range(total_geral)), redraw=False)

        if n_ativos == 0:
            self.lbl_estatisticas.config(text=f"📊 Ativos: 0/{total_geral} | Média: R$ 0,00 | Mediana: R$ 0,00 | CV: 0,00%", fg="#1f3c5b")
            return

        if len(unidades_vistas) > 1:
            self.lbl_estatisticas.config(
                text=f"⚠ ALERTA: Mistura de unidades ({', '.join(unidades_vistas)}) nas amostras ativas! Padronize para calcular estatísticas.",
                fg="#b22222"
            )
            return

        un_str = next(iter(unidades_vistas))
        media = statistics.mean(unitarios_ativos)
        mediana = statistics.median(unitarios_ativos)
        minimo = min(unitarios_ativos)
        maximo = max(unitarios_ativos)
        desvio = statistics.stdev(unitarios_ativos) if n_ativos > 1 else 0.0
        cv = (desvio / media * 100.0) if media > 0 else 0.0

        outliers = 0
        for r_idx, d in enumerate(self.dados_pesquisas):
            if not d.get("utilizado", True):
                self.sheet.highlight_rows(rows=[r_idx], bg="#eeeeee", fg="#888888", redraw=False)
            else:
                u_val = d.get("unitario", 0.0)
                if u_val > 0 and abs(u_val - media) > 0.30 * media:
                    outliers += 1
                    self.sheet.highlight_rows(rows=[r_idx], bg="#fdecea", fg="#9b1c1c", redraw=False)

        alerta_out = f"  •  ⚠ {outliers} fora de ±30%" if outliers else ""
        self.lbl_estatisticas.config(
            text=(
                f"📊 Ativos: {n_ativos}/{total_geral}  •  Média: R$ {formatar_moeda_br(media)}/{un_str}  •  "
                f"Mediana: R$ {formatar_moeda_br(mediana)}  •  DP: R$ {formatar_moeda_br(desvio)}  •  "
                f"Ampl.: R$ {formatar_moeda_br(minimo)} a {formatar_moeda_br(maximo)}  •  "
                f"CV: {formatar_numero_br(cv, 2)}%{alerta_out}"
            ),
            fg="#1f3c5b"
        )

    # --- EVENTOS DA PLANILHA ---
    def _ao_selecionar_celula(self, event=None):
        sel = self.sheet.get_currently_selected()
        if not sel:
            return
        r, c = sel.row, sel.column
        if r is None or c is None or r >= len(self.dados_pesquisas) or c >= len(self.esquema_colunas):
            return
        chave, titulo, _ = self.esquema_colunas[c]
        self.celula_ativa = {"row": r, "col": c, "chave": chave}
        val = self.sheet.get_cell_data(r, c)
        self.lbl_fx.config(text=f"fx  [{titulo} · D.{self.dados_pesquisas[r]['dado_id']}]")
        self.txt_fx.delete(0, tk.END)
        self.txt_fx.insert(0, "" if val is None else str(val))
        if self.painel_lateral_visivel:
            self._atualizar_painel_lateral(r)

    def _aplicar_barra_fx(self):
        r = self.celula_ativa.get("row")
        c = self.celula_ativa.get("col")
        if r is None or c is None:
            return
        novo_val = self.txt_fx.get()
        self._gravar_alteracao_celula(r, c, novo_val)
        self.sheet.redraw()

    def _ao_terminar_edicao_celula(self, event=None):
        if not event:
            return
        r = getattr(event, "row", None)
        c = getattr(event, "column", None)
        val = getattr(event, "value", None)
        if r is None or c is None:
            return
        self._gravar_alteracao_celula(r, c, val)
        self.sheet.redraw()

    def _gravar_alteracao_celula(self, r, c, novo_val):
        if r >= len(self.dados_pesquisas) or c >= len(self.esquema_colunas):
            return
        chave = self.esquema_colunas[c][0]
        if chave in ("dado_id", "unitario", "status"):
            return

        dado = self.dados_pesquisas[r]
        self.alteracoes_pendentes = True

        if chave == "utilizado":
            dado["utilizado"] = bool(novo_val) if isinstance(novo_val, bool) else str(novo_val).strip().lower() not in ("false", "0", "não", "nao", "")
        elif chave == "valor_total":
            dado["valor_total"] = converter_para_float(novo_val)
        elif chave == "area_terreno":
            un = dado.get("unidade", self.unidade_padrao_projeto)
            dado["area_terreno"] = converter_para_float(novo_val, permitir_milhar_ponto=(un != "ha"))
        elif chave == "area_construida":
            dado["area_construida"] = converter_para_float(novo_val)
        elif chave == "unidade":
            un_str = "ha" if "ha" in str(novo_val).lower() else "m²"
            dado["unidade"] = un_str
            dado["localizacao"] = "Rural" if un_str == "ha" else "Urbana"
        elif chave == "telefone":
            dado["telefone"] = sanitizar_telefone(novo_val)
        elif chave in ("coord_e", "coord_s"):
            dado[chave] = limpar_sufixo_coord(sanitizar_texto_celula(novo_val))
        elif chave.startswith("var_"):
            nome_v = chave[4:]
            dado.setdefault("variaveis_extras", {})[nome_v] = sanitizar_texto_celula(novo_val)
            aplicar_nota_agronomica_no_dado(dado, self.variaveis_config)
        else:
            dado[chave] = sanitizar_texto_celula(novo_val)

        nova_linha = self._montar_linha_sheet(dado)
        for col_i, val_i in enumerate(nova_linha):
            self.sheet.set_cell_data(r, col_i, val_i, redraw=False)

        self._atualizar_estatisticas_e_destaques()
        self._autosave()

    def _ao_colar_planilha(self, event=None):
        matriz = self.sheet.get_sheet_data()
        while len(self.dados_pesquisas) < len(matriz):
            self._criar_registro_vazio(herdar_anterior=True)

        for r_idx, linha in enumerate(matriz):
            dado = self.dados_pesquisas[r_idx]
            dado["dado_id"] = r_idx + 1
            for c_idx, (chave, _, _) in enumerate(self.esquema_colunas):
                if c_idx >= len(linha) or chave in ("dado_id", "unitario", "status"):
                    continue
                val = linha[c_idx]
                if chave == "utilizado":
                    dado["utilizado"] = bool(val) if isinstance(val, bool) else str(val).strip().lower() not in ("false", "0", "não", "nao", "")
                elif chave == "valor_total":
                    dado["valor_total"] = converter_para_float(val)
                elif chave == "area_terreno":
                    un = dado.get("unidade", self.unidade_padrao_projeto)
                    dado["area_terreno"] = converter_para_float(val, permitir_milhar_ponto=(un != "ha"))
                elif chave == "area_construida":
                    dado["area_construida"] = converter_para_float(val)
                elif chave == "telefone":
                    dado["telefone"] = sanitizar_telefone(val)
                elif chave.startswith("var_"):
                    dado.setdefault("variaveis_extras", {})[chave[4:]] = sanitizar_texto_celula(val)
                else:
                    dado[chave] = sanitizar_texto_celula(val)
            aplicar_nota_agronomica_no_dado(dado, self.variaveis_config)

        self.alteracoes_pendentes = True
        self._reconstruir_planilha()
        self._autosave()

    def _ao_deletar_conteudo(self, event=None):
        self._ao_colar_planilha()

    def _criar_registro_vazio(self, herdar_anterior=True):
        item_id = len(self.dados_pesquisas) + 1
        ant = self.dados_pesquisas[-1] if (herdar_anterior and self.dados_pesquisas) else {}
        un = ant.get("unidade", self.unidade_padrao_projeto)
        reg = {
            "dado_id": item_id,
            "utilizado": True,
            "informante": "",
            "telefone": "",
            "endereco": "",
            "bairro": "",
            "municipio": ant.get("municipio", ""),
            "valor_total": 0.0,
            "area_terreno": 0.0,
            "area_construida": 0.0,
            "unitario": 0.0,
            "unidade": un,
            "localizacao": "Rural" if un == "ha" else "Urbana",
            "zona_utm": ant.get("zona_utm", ""),
            "coord_e": "",
            "coord_s": "",
            "data": ant.get("data", data_hoje_br()),
            "link": "",
            "foto1": "",
            "foto2": "",
            "variaveis_extras": {}
        }
        self.dados_pesquisas.append(reg)
        return reg

    def _adicionar_nova_linha(self):
        self._criar_registro_vazio(herdar_anterior=True)
        self.alteracoes_pendentes = True
        self._reconstruir_planilha()
        self._autosave()

    def _duplicar_linha_selecionada(self):
        sel = self.sheet.get_currently_selected()
        if not sel or sel.row is None or sel.row >= len(self.dados_pesquisas):
            return
        copia = copy.deepcopy(self.dados_pesquisas[sel.row])
        copia["dado_id"] = len(self.dados_pesquisas) + 1
        self.dados_pesquisas.append(copia)
        self.alteracoes_pendentes = True
        self._reconstruir_planilha()
        self._autosave()

    def _alternar_utilizado_selecionado(self):
        linhas = self.sheet.get_selected_rows(get_cells_as_rows=True)
        if not linhas:
            sel = self.sheet.get_currently_selected()
            if sel and sel.row is not None:
                linhas = {sel.row}
        for r in linhas:
            if r < len(self.dados_pesquisas):
                self.dados_pesquisas[r]["utilizado"] = not self.dados_pesquisas[r].get("utilizado", True)
        self.alteracoes_pendentes = True
        self._reconstruir_planilha()
        self._autosave()

    def _preencher_para_baixo(self):
        caixas = self.sheet.get_all_selection_boxes()
        if not caixas:
            return
        for box in caixas:
            r1, c1, r2, c2 = box.from_r, box.from_c, box.upto_r, box.upto_c
            if r2 - r1 <= 1:
                continue
            for c in range(c1, c2):
                val_topo = self.sheet.get_cell_data(r1, c)
                for r in range(r1 + 1, r2):
                    self._gravar_alteracao_celula(r, c, val_topo)
        self.sheet.redraw()

    def _excluir_linhas_selecionadas(self):
        linhas = sorted(self.sheet.get_selected_rows(get_cells_as_rows=True), reverse=True)
        if not linhas:
            sel = self.sheet.get_currently_selected()
            if sel and sel.row is not None:
                linhas = [sel.row]
        if not linhas:
            return
        if not messagebox.askyesno("Excluir", f"Deseja excluir {len(linhas)} linha(s) selecionada(s)?"):
            return
        for r in linhas:
            if r < len(self.dados_pesquisas):
                self.dados_pesquisas.pop(r)
        for i, d in enumerate(self.dados_pesquisas):
            d["dado_id"] = i + 1
        self.celula_ativa = {"row": None, "col": None, "chave": None}
        self.alteracoes_pendentes = True
        self._reconstruir_planilha()
        self._autosave()

    # --- PAINEL LATERAL DE MÍDIA (COM MINIATURAS E CTRL+V DE PRINT) ---
    def _construir_painel_lateral(self):
        for w in self.frame_painel_lateral.winfo_children():
            w.destroy()

        self.lbl_lat_titulo = ttk.Label(self.frame_painel_lateral, text="Selecione uma linha na grade", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b")
        self.lbl_lat_titulo.pack(anchor="w", pady=(0, 8))

        ttk.Label(self.frame_painel_lateral, text="Foto 1 (Imóvel / Drone):", font=("Segoe UI", 8, "bold")).pack(anchor="w")
        f1 = ttk.Frame(self.frame_painel_lateral)
        f1.pack(fill="x", pady=2)
        self.lat_foto1 = ttk.Entry(f1, width=22)
        self.lat_foto1.pack(side="left", fill="x", expand=True)
        ttk.Button(f1, text="...", width=3, command=lambda: self._escolher_foto_lateral("foto1")).pack(side="left", padx=2)
        ttk.Button(f1, text="📋 Print", width=7, command=lambda: self._colar_print_clipboard("foto1")).pack(side="left")
        self.lbl_thumb1 = tk.Label(self.frame_painel_lateral, text="[ Sem Foto 1 ]", bg="#e9ecef", width=32, height=6, relief="groove")
        self.lbl_thumb1.pack(fill="x", pady=(2, 8))

        ttk.Label(self.frame_painel_lateral, text="Foto 2 (Print Anúncio):", font=("Segoe UI", 8, "bold")).pack(anchor="w")
        f2 = ttk.Frame(self.frame_painel_lateral)
        f2.pack(fill="x", pady=2)
        self.lat_foto2 = ttk.Entry(f2, width=22)
        self.lat_foto2.pack(side="left", fill="x", expand=True)
        ttk.Button(f2, text="...", width=3, command=lambda: self._escolher_foto_lateral("foto2")).pack(side="left", padx=2)
        ttk.Button(f2, text="📋 Print", width=7, command=lambda: self._colar_print_clipboard("foto2")).pack(side="left")
        self.lbl_thumb2 = tk.Label(self.frame_painel_lateral, text="[ Sem Foto 2 ]", bg="#e9ecef", width=32, height=6, relief="groove")
        self.lbl_thumb2.pack(fill="x", pady=(2, 10))

        ttk.Button(self.frame_painel_lateral, text="💾 Aplicar Fotos na Linha", style="Primary.TButton", command=self._salvar_fotos_lateral).pack(fill="x", pady=3)
        ttk.Button(self.frame_painel_lateral, text="✖ Ocultar Painel", command=self._alternar_painel_lateral).pack(fill="x", pady=2)

    def _alternar_painel_lateral(self):
        if self.painel_lateral_visivel:
            self.pane_principal.forget(self.frame_painel_lateral)
            self.painel_lateral_visivel = False
        else:
            self.pane_principal.add(self.frame_painel_lateral, stretch="never")
            self.pane_principal.sash_place(0, 940, 0)
            self.painel_lateral_visivel = True
            r = self.celula_ativa.get("row")
            if r is not None and r < len(self.dados_pesquisas):
                self._atualizar_painel_lateral(r)

    def _atualizar_painel_lateral(self, r):
        if not self.painel_lateral_visivel or r >= len(self.dados_pesquisas):
            return
        dado = self.dados_pesquisas[r]
        self.lbl_lat_titulo.config(text=f"Pesquisa D.{dado['dado_id']} — {dado.get('municipio', '')}")
        self.lat_foto1.delete(0, tk.END)
        self.lat_foto1.insert(0, dado.get("foto1", ""))
        self.lat_foto2.delete(0, tk.END)
        self.lat_foto2.insert(0, dado.get("foto2", ""))
        self._renderizar_miniatura(dado.get("foto1", ""), self.lbl_thumb1, 1)
        self._renderizar_miniatura(dado.get("foto2", ""), self.lbl_thumb2, 2)

    def _renderizar_miniatura(self, caminho, label_widget, slot):
        caminho_real = self._resolver_caminho_foto(caminho)
        if caminho_real and os.path.exists(caminho_real):
            try:
                img = Image.open(caminho_real)
                img = ImageOps.exif_transpose(img)
                img.thumbnail((240, 110), LANCZOS_FILTER)
                tk_img = ImageTk.PhotoImage(img)
                if slot == 1:
                    self._img_thumb1 = tk_img
                else:
                    self._img_thumb2 = tk_img
                label_widget.config(image=tk_img, text="", width=240, height=110)
                return
            except Exception:
                pass
        label_widget.config(image="", text=f"[ Sem Foto {slot} ]", width=32, height=6)

    def _escolher_foto_lateral(self, campo):
        caminho = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg *.webp *.bmp"), ("Todos", "*.*")])
        if not caminho:
            return
        entry = self.lat_foto1 if campo == "foto1" else self.lat_foto2
        entry.delete(0, tk.END)
        entry.insert(0, caminho)
        self._salvar_fotos_lateral()

    def _colar_print_clipboard(self, campo):
        r = self.celula_ativa.get("row")
        if r is None or r >= len(self.dados_pesquisas):
            messagebox.showwarning("Aviso", "Clique em uma linha da grade antes de colar o print.")
            return
        try:
            img = ImageGrab.grabclipboard()
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler a área de transferência: {e}")
            return
        if not isinstance(img, Image.Image):
            messagebox.showwarning("Sem Imagem", "Nenhuma imagem encontrada na área de transferência. Use Win+Shift+S para capturar a tela primeiro.")
            return

        pasta_dest = self._obter_pasta_anexos_projeto()
        os.makedirs(pasta_dest, exist_ok=True)
        nome_arq = f"D{self.dados_pesquisas[r]['dado_id']:02d}_{campo}_{datetime.now().strftime('%H%M%S')}.png"
        caminho_salvo = os.path.join(pasta_dest, nome_arq)
        img.save(caminho_salvo, "PNG")

        entry = self.lat_foto1 if campo == "foto1" else self.lat_foto2
        entry.delete(0, tk.END)
        entry.insert(0, caminho_salvo)
        self._salvar_fotos_lateral()

    def _salvar_fotos_lateral(self):
        r = self.celula_ativa.get("row")
        if r is None or r >= len(self.dados_pesquisas):
            return
        dado = self.dados_pesquisas[r]
        dado["foto1"] = self.lat_foto1.get().strip()
        dado["foto2"] = self.lat_foto2.get().strip()
        self.alteracoes_pendentes = True
        self._atualizar_painel_lateral(r)
        nova_linha = self._montar_linha_sheet(dado)
        for col_i, val_i in enumerate(nova_linha):
            self.sheet.set_cell_data(r, col_i, val_i, redraw=False)
        self.sheet.redraw()
        self._autosave()

    def _obter_pasta_anexos_projeto(self):
        if self.caminho_projeto_atual:
            base = os.path.splitext(self.caminho_projeto_atual)[0]
            return f"{base}_anexos"
        return ANEXOS_TEMP_DIR

    def _resolver_caminho_foto(self, caminho):
        if not caminho:
            return ""
        if os.path.exists(caminho):
            return caminho
        if self.caminho_projeto_atual:
            rel = os.path.join(os.path.dirname(self.caminho_projeto_atual), caminho)
            if os.path.exists(rel):
                return rel
        return caminho

    # --- CONVERSOR UTM E PERFIS ---
    def _janela_conversor_utm(self):
        r = self.celula_ativa.get("row")
        texto = simpledialog.askstring(
            "Conversor Lat/Long, DMS ou Google Maps ➔ UTM",
            "Cole abaixo coordenadas decimais (ex: -16,6868, -49,2648),\nDMS (16°41'12\"S 49°15'53\"W) ou link do Google Maps:",
            parent=self.root
        )
        if not texto:
            return
        res = converter_latlon_para_utm(texto)
        if not res:
            messagebox.showwarning("Formato não reconhecido", "Não foi possível extrair Latitude e Longitude do texto informado.")
            return
        zona, e_val, s_val = res
        if r is not None and r < len(self.dados_pesquisas):
            dado = self.dados_pesquisas[r]
            dado["zona_utm"] = zona
            dado["coord_e"] = e_val
            dado["coord_s"] = s_val
            self.alteracoes_pendentes = True
            nova_linha = self._montar_linha_sheet(dado)
            for col_i, val_i in enumerate(nova_linha):
                self.sheet.set_cell_data(r, col_i, val_i, redraw=False)
            self.sheet.redraw()
            self._autosave()
            messagebox.showinfo("Aplicado na Linha", f"UTM gravado na Pesquisa D.{dado['dado_id']}:\nZona: {zona} | E: {e_val} m | S: {s_val} m")
        else:
            messagebox.showinfo("Resultado UTM", f"Zona: {zona}\nCoord. E: {e_val} m\nCoord. S: {s_val} m")

    def _janela_campos_opcionais(self):
        janela = tk.Toplevel(self.root)
        janela.title("Campos Opcionais da Ficha")
        janela.geometry("360x260")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Marque os campos que deseja exibir na grade e no Word:", font=("Segoe UI", 9, "bold")).pack(pady=10, padx=12, anchor="w")
        rotulos = {
            "telefone": "Telefone / Contato",
            "bairro": "Bairro",
            "area_construida": "Área Construída (m²)",
            "link": "Link do Anúncio",
            "localizacao": "Localização (Urbana/Rural)"
        }
        vars_tk = {}
        for chave, rot in rotulos.items():
            v = tk.BooleanVar(value=self.campos_opcionais.get(chave, True))
            ttk.Checkbutton(janela, text=rot, variable=v).pack(anchor="w", padx=25, pady=3)
            vars_tk[chave] = v

        def aplicar():
            for chave, v in vars_tk.items():
                self.campos_opcionais[chave] = v.get()
            self.alteracoes_pendentes = True
            self._reconstruir_planilha()
            janela.destroy()

        ttk.Button(janela, text="✔ Aplicar", style="Primary.TButton", command=aplicar).pack(pady=14)

    def _salvar_como_novo_perfil(self):
        nome = simpledialog.askstring("Salvar Perfil", "Digite um nome para este perfil (ex: Rural Goianésia):", parent=self.root)
        if not nome or not nome.strip():
            return
        self.perfis_usuario[nome.strip()] = {
            "modelo": self.modelo_ativo,
            "unidade_padrao": self.unidade_padrao_projeto,
            "campos_opcionais": copy.deepcopy(self.campos_opcionais),
            "variaveis_config": copy.deepcopy(self.variaveis_config)
        }
        self._salvar_perfis_usuario()
        messagebox.showinfo("Perfil Salvo", f"Perfil '{nome.strip()}' salvo permanentemente!")

    def _aplicar_perfil_rapido(self, tipo):
        if tipo == "URBANO":
            self.modelo_ativo = "PADRAO"
            self.variaveis_config = copy.deepcopy(DEFAULT_VARS_PADRAO)
            self.unidade_padrao_projeto = "m²"
        elif tipo == "RURAL":
            self.modelo_ativo = "PADRAO"
            self.variaveis_config = copy.deepcopy(DEFAULT_VARS_RURAL)
            self.unidade_padrao_projeto = "ha"
        elif tipo == "COPASA":
            self.modelo_ativo = "COPASA"
            self.variaveis_config = copy.deepcopy(DEFAULT_VARS_COPASA)
            self.unidade_padrao_projeto = "m²"
        self._salvar_config_variaveis()
        if hasattr(self, "lbl_info_mod"):
            mod_label = "Modelo 1: Padrão" if self.modelo_ativo == "PADRAO" else "Modelo 2: COPASA"
            self.lbl_info_mod.config(text=f"  |  {mod_label} ({self.unidade_padrao_projeto})")
        self._reconstruir_planilha()

    def _janela_config_variaveis(self):
        janela = tk.Toplevel(self.root)
        janela.title("Gerenciador de Variáveis e Perfis")
        janela.geometry("780x540")
        janela.transient(self.root)
        janela.grab_set()

        frame_perfis = ttk.LabelFrame(janela, text=" 1. Perfis Rápidos ", padding=6)
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
                ttk.Checkbutton(frame_lista, text=f"Var {i+1}", variable=var_ativa).grid(row=i+1, column=0, padx=2, pady=2)

                e_nome = ttk.Entry(frame_lista, width=22)
                e_nome.insert(0, cfg.get("nome", ""))
                e_nome.grid(row=i+1, column=1, padx=3, pady=2)

                cb_tipo = ttk.Combobox(frame_lista, values=["codigo", "numero", "texto"], width=9, state="readonly")
                cb_tipo.set(cfg.get("tipo", "texto"))
                cb_tipo.grid(row=i+1, column=2, padx=3, pady=2)

                opcoes_visiveis = [o for o in cfg.get("opcoes", []) if o != ""]
                e_opcoes = ttk.Entry(frame_lista, width=44)
                e_opcoes.insert(0, ", ".join(opcoes_visiveis))
                e_opcoes.grid(row=i+1, column=3, padx=3, pady=2)

                entradas_vars.append((var_ativa, e_nome, cb_tipo, e_opcoes))

        ttk.Button(frame_perfis, text="🏙 Urbano", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_PADRAO))).pack(side="left", padx=4)
        ttk.Button(frame_perfis, text="🌾 Rural (Agronômico)", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_RURAL))).pack(side="left", padx=4)
        ttk.Button(frame_perfis, text="💧 COPASA", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_COPASA))).pack(side="left", padx=4)

        frame_cat = ttk.Frame(janela, padding=(10, 2))
        frame_cat.pack(fill="x")
        ttk.Label(frame_cat, text="Catálogo usual:").pack(side="left")
        cb_cat = ttk.Combobox(frame_cat, values=[v["nome"] for v in CATALOGO_VARIAVEIS_USUAIS], width=28, state="readonly")
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

        ttk.Button(frame_cat, text="➕ Incluir", command=adicionar_do_catalogo).pack(side="left")

        frame_lista = ttk.LabelFrame(janela, text=" 2. Variáveis Ativas ", padding=8)
        frame_lista.pack(fill="both", expand=True, padx=10, pady=6)
        recarregar_linhas_editor(self.variaveis_config)

        def salvar():
            novas = []
            for var_ativa, e_nome, cb_tipo, e_opcoes in entradas_vars:
                if var_ativa.get() and e_nome.get().strip():
                    opts = [op.strip() for op in e_opcoes.get().split(",") if op.strip()]
                    if cb_tipo.get() == "codigo" and opts:
                        opts = [""] + opts
                    novas.append({"nome": e_nome.get().strip(), "tipo": cb_tipo.get(), "opcoes": opts})
            self.variaveis_config = novas
            self._salvar_config_variaveis()
            self._reconstruir_planilha()
            janela.destroy()

        ttk.Button(janela, text="💾 Salvar e Atualizar Planilha", style="Primary.TButton", command=salvar).pack(pady=8)

    def _mostrar_tabela_nota_agronomica(self):
        janela = tk.Toplevel(self.root)
        janela.title("Tabela 02 – Nota Agronômica (Kozma 1984 x IN RFB 1877/2019)")
        janela.geometry("620x260")
        janela.transient(self.root)

        cols = ("situacao", "peso", "I", "II", "III", "IV", "V", "VI")
        t_ref = ttk.Treeview(janela, columns=cols, show="headings", height=7)
        t_ref.heading("situacao", text="Situação / Acesso")
        t_ref.heading("peso", text="Peso")
        for rom, p_apt in [("I", "1,00"), ("II", "0,80"), ("III", "0,64"), ("IV", "0,51"), ("V", "0,41"), ("VI", "0,20")]:
            t_ref.heading(rom, text=f"{rom} ({p_apt})")
            t_ref.column(rom, width=65, anchor="center")
        t_ref.column("situacao", width=120, anchor="w")
        t_ref.column("peso", width=55, anchor="center")
        t_ref.pack(fill="both", expand=True, padx=10, pady=8)

        for sit, p_s in [("Ótima", "1,00"), ("Muito Boa", "0,95"), ("Boa", "0,90"), ("Desfavorável", "0,80"), ("Má", "0,75"), ("Péssima", "0,70")]:
            t_ref.insert("", "end", values=(
                sit, p_s,
                TABELA_NOTA_AGRONOMICA["I"][sit], TABELA_NOTA_AGRONOMICA["II"][sit],
                TABELA_NOTA_AGRONOMICA["III"][sit], TABELA_NOTA_AGRONOMICA["IV"][sit],
                TABELA_NOTA_AGRONOMICA["V"][sit], TABELA_NOTA_AGRONOMICA["VI"][sit]
            ))

    # --- ABRIR / SALVAR PROJETO (COM PORTABILIDADE DE FOTOS) ---
    def _salvar_projeto(self):
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Projeto de Pesquisa", "*.json")])
        if not caminho:
            return
        self.caminho_projeto_atual = caminho
        pasta_anexos = self._obter_pasta_anexos_projeto()
        dir_proj = os.path.dirname(caminho)

        for d in self.dados_pesquisas:
            for campo_f in ("foto1", "foto2"):
                orig = self._resolver_caminho_foto(d.get(campo_f, ""))
                if orig and os.path.exists(orig) and not orig.startswith("http"):
                    os.makedirs(pasta_anexos, exist_ok=True)
                    dest = os.path.join(pasta_anexos, os.path.basename(orig))
                    if os.path.abspath(orig) != os.path.abspath(dest):
                        try:
                            shutil.copy2(orig, dest)
                        except Exception:
                            pass
                    d[campo_f] = os.path.relpath(dest, dir_proj)

        dados_salvar = {
            "modelo": self.modelo_ativo,
            "unidade_padrao": self.unidade_padrao_projeto,
            "campos_opcionais": self.campos_opcionais,
            "variaveis_config": self.variaveis_config,
            "pesquisas": self.dados_pesquisas
        }
        try:
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(dados_salvar, f, ensure_ascii=False, indent=4)
            self.alteracoes_pendentes = False
            messagebox.showinfo("Sucesso", "Projeto e anexos salvos com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar arquivo: {e}")

    def _aplicar_conteudo_json(self, conteudo):
        self.modelo_ativo = conteudo.get("modelo", "PADRAO")
        self.variaveis_config = conteudo.get("variaveis_config") or copy.deepcopy(DEFAULT_VARS_PADRAO)
        self.campos_opcionais = conteudo.get("campos_opcionais") or copy.deepcopy(DEFAULT_CAMPOS_OPCIONAIS)
        raw_pesquisas = conteudo.get("pesquisas", [])
        self.dados_pesquisas = [normalizar_dado(d, self.variaveis_config) for d in raw_pesquisas]
        if "unidade_padrao" in conteudo:
            self.unidade_padrao_projeto = conteudo["unidade_padrao"]
        else:
            self.unidade_padrao_projeto = "ha" if any(d.get("unidade") == "ha" for d in self.dados_pesquisas) else "m²"
        self.alteracoes_pendentes = False

    def _executar_abertura_json(self):
        caminho = filedialog.askopenfilename(filetypes=[("Projeto de Pesquisa", "*.json")])
        if not caminho:
            return False
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                conteudo = json.load(f)
            self.caminho_projeto_atual = caminho
            self._aplicar_conteudo_json(conteudo)
            if hasattr(self, "sheet"):
                self._reconstruir_planilha()
                messagebox.showinfo("Sucesso", f"{len(self.dados_pesquisas)} pesquisas carregadas!")
            return True
        except Exception as e:
            messagebox.showerror("Erro ao Abrir", f"Arquivo inválido: {e}")
            return False

    def _abrir_projeto_menu(self):
        self._executar_abertura_json()

    # --- IMPORTAÇÃO EXCEL ---
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
            self._processar_dataframe_importado(df)
        except Exception as e:
            messagebox.showerror("Erro na Importação", f"Falha ao ler planilha: {e}")

    def _processar_dataframe_importado(self, df):
        mapeamento = {}
        unidade_detectada = None
        nomes_vars_config = {normalizar_chave_var(v["nome"]) for v in self.variaveis_config}

        for col in df.columns:
            c_clean = normalizar_chave_var(col)
            if c_clean in nomes_vars_config:
                continue
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
            elif "área constr" in c_clean or "area constr" in c_clean:
                mapeamento["area_construida"] = col
            elif "área terreno" in c_clean or "area terreno" in c_clean or c_clean.startswith(("area", "área")):
                mapeamento["area_terreno"] = col
                if "(ha)" in c_clean:
                    unidade_detectada = "ha"
                elif "(m²)" in c_clean or "(m2)" in c_clean:
                    unidade_detectada = "m²"
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

        novos = []
        for _, row in df.iterrows():
            if pd.isna(row.get(mapeamento.get("informante", ""))) and pd.isna(row.get(mapeamento.get("endereco", ""))):
                continue

            if "unidade" in mapeamento and not pd.isna(row.get(mapeamento["unidade"])):
                unidade = "ha" if "ha" in str(row.get(mapeamento["unidade"])).lower() else "m²"
            elif unidade_detectada:
                unidade = unidade_detectada
            else:
                unidade = self.unidade_padrao_projeto

            v_total = converter_para_float(row.get(mapeamento.get("valor_total", ""), 0.0))
            a_terr = converter_para_float(row.get(mapeamento.get("area_terreno", ""), 0.0), permitir_milhar_ponto=(unidade != "ha"))
            a_const = converter_para_float(row.get(mapeamento.get("area_construida", ""), 0.0))

            utilizado_flag = True
            if "utilizado" in mapeamento and not pd.isna(row.get(mapeamento["utilizado"])):
                val_u = str(row.get(mapeamento["utilizado"])).strip().lower()
                utilizado_flag = val_u not in ("não", "nao", "false", "0", "☐", "não utilizado no modelo")

            str_data = sanitizar_texto_celula(row.get(mapeamento["data"])) if "data" in mapeamento else ""
            if not str_data:
                str_data = data_hoje_br()

            reg = {
                "dado_id": len(self.dados_pesquisas) + len(novos) + 1,
                "utilizado": utilizado_flag,
                "informante": sanitizar_texto_celula(row.get(mapeamento.get("informante", ""), "")),
                "telefone": sanitizar_telefone(row.get(mapeamento.get("telefone", ""), "")),
                "endereco": sanitizar_texto_celula(row.get(mapeamento.get("endereco", ""), "")),
                "bairro": sanitizar_texto_celula(row.get(mapeamento.get("bairro", ""), "")),
                "municipio": sanitizar_texto_celula(row.get(mapeamento.get("municipio", ""), "")),
                "valor_total": v_total,
                "area_terreno": a_terr,
                "area_construida": a_const,
                "unitario": v_total / a_terr if a_terr > 0 else 0.0,
                "unidade": unidade,
                "localizacao": "Rural" if unidade == "ha" else "Urbana",
                "zona_utm": sanitizar_texto_celula(row.get(mapeamento.get("zona_utm", ""), "")),
                "coord_e": limpar_sufixo_coord(sanitizar_texto_celula(row.get(mapeamento.get("coord_e", ""), ""))),
                "coord_s": limpar_sufixo_coord(sanitizar_texto_celula(row.get(mapeamento.get("coord_s", ""), ""))),
                "data": str_data,
                "link": sanitizar_texto_celula(row.get(mapeamento.get("link", ""), "")),
                "foto1": "",
                "foto2": "",
                "variaveis_extras": {}
            }

            for v_cfg in self.variaveis_config:
                chave_alvo = normalizar_chave_var(v_cfg["nome"])
                for col_p in df.columns:
                    if normalizar_chave_var(col_p) == chave_alvo:
                        reg["variaveis_extras"][v_cfg["nome"]] = sanitizar_texto_celula(row.get(col_p, ""))
                        break

            aplicar_nota_agronomica_no_dado(reg, self.variaveis_config)
            novos.append(reg)

        if novos:
            self.dados_pesquisas.extend(novos)
            self.alteracoes_pendentes = True
            self._reconstruir_planilha()
            self._autosave()
            messagebox.showinfo("Importação Concluída", f"{len(novos)} pesquisas importadas com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado válido encontrado.")

    # --- EXPORTAÇÃO EXCEL FORMATADA ---
    def _exportar_excel(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportar.")
            return
        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilha Excel", "*.xlsx")])
        if not caminho:
            return

        linhas_export = []
        for d in self.dados_pesquisas:
            un = d.get("unidade", self.unidade_padrao_projeto)
            linha = {
                "D.": d["dado_id"],
                "Utilizado no Modelo": "Sim" if d.get("utilizado", True) else "Não utilizado no modelo",
                "Informante": d.get("informante", ""),
            }
            if self.campos_opcionais.get("telefone", True):
                linha["Telefone"] = d.get("telefone", "")
            linha["Endereço"] = d.get("endereco", "")
            if self.campos_opcionais.get("bairro", True):
                linha["Bairro"] = d.get("bairro", "")
            linha["Município"] = d.get("municipio", "")
            linha["Valor Total (R$)"] = d.get("valor_total", 0.0)
            linha[f"Área Terreno ({un})"] = d.get("area_terreno", 0.0)
            if self.campos_opcionais.get("area_construida", True):
                linha["Área Construída (m²)"] = d.get("area_construida", 0.0)

            extras = d.get("variaveis_extras", {})
            for v_cfg in self.variaveis_config:
                linha[v_cfg["nome"]] = extras.get(v_cfg["nome"], "")

            linha[f"Unitário (R$/{un})"] = d.get("unitario", 0.0)
            linha["Zona UTM"] = d.get("zona_utm", "")
            linha["Coord. E (m)"] = d.get("coord_e", "")
            linha["Coord. S (m)"] = d.get("coord_s", "")
            if self.campos_opcionais.get("localizacao", True):
                linha["Localização"] = d.get("localizacao", "Urbana")
            linha["Data"] = d.get("data", "")
            if self.campos_opcionais.get("link", True):
                linha["Link"] = d.get("link", "")
            linhas_export.append(linha)

        df = pd.DataFrame(linhas_export)
        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Ficha de Pesquisa", index=False)
            ws = writer.sheets["Ficha de Pesquisa"]
            ws.freeze_panes = "A2"
            for col_cells in ws.columns:
                col_letter = col_cells[0].column_letter
                titulo = str(col_cells[0].value or "")
                max_len = max(len(str(c.value or "")) for c in col_cells)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 42)
                if "R$" in titulo:
                    for cell in col_cells[1:]:
                        cell.number_format = '#,##0.00'

        messagebox.showinfo("Sucesso", "Planilha Excel (SISDEA) formatada e exportada com sucesso!")

    # --- EXPORTAÇÃO WORD ---
    def _baixar_imagem(self, caminho_ou_url):
        if not caminho_ou_url:
            return None
        caminho_ou_url = normalizar_url_gdrive(caminho_ou_url)
        try:
            if caminho_ou_url.startswith("http://") or caminho_ou_url.startswith("https://"):
                resp = requests.get(caminho_ou_url, timeout=12)
                if resp.status_code == 200 and resp.headers.get("Content-Type", "").lower().startswith("image/"):
                    return redimensionar_e_corrigir_imagem(io.BytesIO(resp.content))
            else:
                caminho_real = self._resolver_caminho_foto(caminho_ou_url)
                if caminho_real and os.path.exists(caminho_real):
                    return redimensionar_e_corrigir_imagem(caminho_real)
        except Exception as e:
            logging.warning(f"Erro ao carregar imagem '{caminho_ou_url}': {e}")
        return None

    def _iniciar_exportacao_word_thread(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportar.")
            return

        so_ativos = False
        if any(not d.get("utilizado", True) for d in self.dados_pesquisas):
            resp = messagebox.askyesnocancel(
                "Filtro de Amostras",
                "Deseja exportar TODAS as fichas (marcando as desativadas como 'Não utilizado no modelo')?\n\n"
                "• Sim = Exportar todas\n• Não = Exportar somente as utilizadas (☑)\n• Cancelar = Voltar"
            )
            if resp is None:
                return
            so_ativos = (resp is False)

        caminho = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        if hasattr(self, "progress_bar"):
            self.progress_bar.pack(fill="x", padx=8, pady=2)
            self.progress_bar["value"] = 0

        lista_export = [d for d in self.dados_pesquisas if (not so_ativos or d.get("utilizado", True))]
        thread = threading.Thread(
            target=self._processar_geracao_word,
            args=(caminho, copy.deepcopy(lista_export), copy.deepcopy(self.variaveis_config), self.modelo_ativo, copy.deepcopy(self.campos_opcionais)),
            daemon=True
        )
        thread.start()

    def _inserir_imagem_com_limite(self, paragrafo, img_stream, max_w_in, max_h_in):
        try:
            img_stream.seek(0)
            with Image.open(img_stream) as im:
                w_px, h_px = im.size
            ratio = w_px / float(h_px) if h_px else 1.0
            w_in = max_w_in
            h_in = w_in / ratio
            if h_in > max_h_in:
                h_in = max_h_in
                w_in = h_in * ratio
            img_stream.seek(0)
            paragrafo.add_run().add_picture(img_stream, width=Inches(w_in), height=Inches(h_in))
        except Exception:
            paragrafo.add_run("[ Erro ao inserir imagem ]")

    def _processar_geracao_word(self, caminho, dados_lista, variaveis_lista, modelo_selecionado, campos_opc):
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
                if hdr.paragraphs and not hdr.paragraphs[0].text:
                    p_rem = hdr.paragraphs[0]._element
                    p_rem.getparent().remove(p_rem)

                t_logos = hdr.add_table(rows=1, cols=2, width=Inches(5.90))
                t_logos.alignment = WD_TABLE_ALIGNMENT.CENTER
                c_l1, c_l2 = t_logos.cell(0, 0), t_logos.cell(0, 1)
                c_l1.width, c_l2.width = Inches(2.95), Inches(2.95)

                p_l1 = c_l1.paragraphs[0]
                p_l1.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if os.path.exists(caminho_logo_enprol):
                    p_l1.add_run().add_picture(caminho_logo_enprol, height=Inches(0.42))
                else:
                    r = p_l1.add_run("ENPROL")
                    r.bold = True; r.font.name = "Arial"; r.font.size = Pt(13); r.font.color.rgb = RGBColor(28, 93, 153)

                p_l2 = c_l2.paragraphs[0]
                p_l2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if os.path.exists(caminho_logo_copasa):
                    p_l2.add_run().add_picture(caminho_logo_copasa, height=Inches(0.42))
                else:
                    r2 = p_l2.add_run("copasa")
                    r2.bold = True; r2.font.name = "Arial"; r2.font.size = Pt(13); r2.font.color.rgb = RGBColor(50, 110, 180)

                p_sep = hdr.add_paragraph()
                p_sep.paragraph_format.space_before = Pt(0)
                p_sep.paragraph_format.space_after = Pt(2)
                p_sep.paragraph_format.line_spacing = Pt(2)

                t_faixa = hdr.add_table(rows=1, cols=3, width=Inches(5.90))
                t_faixa.alignment = WD_TABLE_ALIGNMENT.CENTER
                t_faixa.autofit = False
                inserir_ordenado(
                    t_faixa._tbl.tblPr,
                    parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'),
                    SUC_CMAR
                )
                c_f0, c_f1, c_f2 = t_faixa.cell(0, 0), t_faixa.cell(0, 1), t_faixa.cell(0, 2)
                c_f0.width, c_f1.width, c_f2.width = Inches(1.77), Inches(2.16), Inches(1.97)
                for c_f in (c_f0, c_f1, c_f2):
                    inserir_ordenado(c_f._tc.get_or_add_tcPr(), parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="4472C4"/>'), SUC_TCPR[1:])

                for pf, txt, al in [(c_f0.paragraphs[0], "CONTRATO: COPASA I", WD_ALIGN_PARAGRAPH.LEFT), (c_f1.paragraphs[0], "", WD_ALIGN_PARAGRAPH.CENTER), (c_f2.paragraphs[0], "PESQUISA DE MERCADO", WD_ALIGN_PARAGRAPH.RIGHT)]:
                    pf.alignment = al
                    pf.paragraph_format.space_before = Pt(0)
                    pf.paragraph_format.space_after = Pt(0)
                    pf.paragraph_format.line_spacing = 1.0
                    if txt:
                        rf = pf.add_run(txt)
                        rf.bold = True; rf.font.name = "Arial"; rf.font.size = Pt(10); rf.font.color.rgb = RGBColor(255, 255, 255)

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
                        r_tit.bold = True; r_tit.font.name = "Arial"; r_tit.font.size = Pt(10)

                        p_bar = doc.add_paragraph()
                        p_bar.paragraph_format.space_before = Pt(0)
                        p_bar.paragraph_format.space_after = Pt(4)
                        inserir_ordenado(p_bar._p.get_or_add_pPr(), parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="245D8C"/></w:pBdr>'), SUC_PPR)
                    else:
                        p_sp = doc.add_paragraph()
                        p_sp.paragraph_format.space_before = Pt(0)
                        p_sp.paragraph_format.space_after = Pt(4)

                tabela = doc.add_table(rows=0, cols=2)
                tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
                tabela.autofit = False
                if modelo_selecionado == "PADRAO":
                    inserir_ordenado(
                        tabela._tbl.tblPr,
                        parse_xml(f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:insideV w:val="none"/></w:tblBorders>'),
                        SUC_TBLPR
                    )
                else:
                    tabela.style = 'Table Grid'

                for dado in dados_lista[i:i+2]:
                    d_id = dado["dado_id"]
                    sufixo = "" if dado.get("utilizado", True) else " (Não utilizado no modelo)"
                    un = dado.get("unidade", "m²")

                    r_dados = tabela.add_row()
                    c_dados, c_foto = r_dados.cells[0], r_dados.cells[1]
                    c_dados.width = Inches(2.95) if modelo_selecionado == "COPASA" else Inches(3.70)
                    c_foto.width = Inches(3.42) if modelo_selecionado == "COPASA" else Inches(3.70)
                    c_dados.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    c_foto.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    p_dados = c_dados.paragraphs[0]
                    p_dados.paragraph_format.line_spacing = 1.08
                    p_dados.paragraph_format.space_before = Pt(2)
                    p_dados.paragraph_format.space_after = Pt(0)

                    def add_line(p, label, val):
                        r1 = p.add_run(label); r1.bold = True; r1.font.name = "Arial"; r1.font.size = Pt(10)
                        r2 = p.add_run(f" {val}\n"); r2.font.name = "Arial"; r2.font.size = Pt(10)

                    zona_str = f"{dado.get('zona_utm', '').strip()} " if dado.get("zona_utm") else ""
                    ce, cs = dado.get("coord_e", ""), dado.get("coord_s", "")
                    coord_txt = f"{zona_str}{ce} m E / {cs} m S" if (ce or cs) else ""

                    if modelo_selecionado == "PADRAO":
                        p_dados.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_top = p_dados.add_run(f"Pesquisa – {d_id:02d}{sufixo}\n\n")
                        r_top.bold = True; r_top.font.name = "Arial"; r_top.font.size = Pt(10)

                        pc = c_dados.add_paragraph()
                        pc.paragraph_format.line_spacing = 1.08
                        pc.paragraph_format.space_before = Pt(0)
                        pc.paragraph_format.space_after = Pt(0)

                        add_line(pc, "Logradouro:", dado.get("endereco", ""))
                        if campos_opc.get("bairro", True) and dado.get("bairro"):
                            add_line(pc, "Bairro:", dado.get("bairro", ""))
                        add_line(pc, "Município:", dado.get("municipio", ""))
                        if campos_opc.get("telefone", True):
                            add_line(pc, "Contato:", f"{dado.get('telefone', '')} - {dado.get('informante', '')}")
                        else:
                            add_line(pc, "Informante:", dado.get("informante", ""))
                        if campos_opc.get("link", True) and dado.get("link"):
                            add_line(pc, "Link:", dado.get("link", ""))
                        pc.add_run("\n")

                        casas = 4 if un == "ha" else 2
                        add_line(pc, "Área Terreno:", f"{formatar_numero_br(dado.get('area_terreno', 0.0), casas)} {un}")
                        if campos_opc.get("area_construida", True):
                            add_line(pc, "Área Construída:", f"{formatar_numero_br(dado.get('area_construida', 0.0), 2)} m²")
                        add_line(pc, "Valor da Oferta:", f"R$ {formatar_moeda_br(dado.get('valor_total', 0.0))}")
                        add_line(pc, f"Valor Unitário/{un}:", f"R$ {formatar_moeda_br(dado.get('unitario', 0.0))}/{un}")
                        pc.add_run("\n")

                        extras = dado.get("variaveis_extras", {})
                        for v_cfg in variaveis_lista:
                            add_line(pc, f"{v_cfg['nome']}:", extras.get(v_cfg["nome"], ""))

                        add_line(pc, "Coordenadas UTM:", coord_txt)
                        if campos_opc.get("localizacao", True):
                            add_line(pc, "Localização:", dado.get("localizacao", "Urbana"))
                        pc.add_run("\n")
                        add_line(pc, "Data:", dado.get("data", data_hoje_br()))

                        inserir_ordenado(c_dados._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:right w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        inserir_ordenado(c_foto._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                    else:
                        add_line(p_dados, "Logradouro:", dado.get("endereco", ""))
                        if campos_opc.get("bairro", True):
                            add_line(p_dados, "Bairro:", dado.get("bairro", ""))
                        add_line(p_dados, "Município:", dado.get("municipio", ""))
                        p_dados.add_run("\n")
                        extras = dado.get("variaveis_extras", {})
                        add_line(p_dados, "Informações:", extras.get("Informações", ""))
                        p_dados.add_run("\n")
                        add_line(p_dados, "Área Terreno:", f"{formatar_numero_br(dado.get('area_terreno', 0.0), 2)} m²")
                        if campos_opc.get("area_construida", True):
                            add_line(p_dados, "Área Construída:", f"{formatar_numero_br(dado.get('area_construida', 0.0), 2)} m²")
                        add_line(p_dados, "Frente:", extras.get("Frente", "Não informado"))
                        p_dados.add_run("\n")
                        add_line(p_dados, "Via de acesso:", extras.get("Via de acesso", "pavimentada"))
                        add_line(p_dados, "Coordenadas UTM:", coord_txt)
                        p_dados.add_run("\n")
                        add_line(p_dados, "Valor Unitário:", f"R$/m² {formatar_numero_br(dado.get('unitario', 0.0), 2)}")
                        add_line(p_dados, "Valor Total:", f"R$ {formatar_moeda_br(dado.get('valor_total', 0.0))}")
                        p_dados.add_run("\n")
                        add_line(p_dados, "Data:", dado.get("data", data_hoje_br()))

                    img1 = self._baixar_imagem(dado.get("foto1"))
                    img2 = self._baixar_imagem(dado.get("foto2"))
                    p_foto = c_foto.paragraphs[0]
                    p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    max_w = 3.25 if modelo_selecionado == "COPASA" else 2.75
                    if img1 and img2:
                        self._inserir_imagem_com_limite(p_foto, img1, max_w, 1.95)
                        p_foto2 = c_foto.add_paragraph()
                        p_foto2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foto2.paragraph_format.space_before = Pt(2)
                        self._inserir_imagem_com_limite(p_foto2, img2, max_w, 1.95)
                    elif img1 or img2:
                        self._inserir_imagem_com_limite(p_foto, img1 or img2, max_w, 3.60)
                    else:
                        rv = p_foto.add_run("[ Sem fotos anexadas ]")
                        rv.font.name = "Arial"; rv.font.size = Pt(10)

                    if modelo_selecionado == "COPASA":
                        inserir_ordenado(c_dados._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        inserir_ordenado(c_foto._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="nil"/><w:bottom w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        r_lbl = tabela.add_row()
                        c_m = r_lbl.cells[0].merge(r_lbl.cells[1])
                        c_m.width = Inches(6.37)
                        inserir_ordenado(c_m._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        p_lbl = c_m.paragraphs[0]
                        p_lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_lbl.paragraph_format.space_before = Pt(4)
                        p_lbl.paragraph_format.space_after = Pt(4)
                        r_num = p_lbl.add_run(f"Pesquisa {d_id:02d}{sufixo}")
                        r_num.font.name = "Arial"; r_num.font.size = Pt(10)

                    self.fila_eventos_ui.put(("progresso", int(((i + 1) / max(1, total_dados)) * 100)))

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
