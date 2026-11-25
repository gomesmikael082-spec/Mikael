import os
import io
import re
import json
import math
import queue
import shutil
import logging
import threading
import copy
import statistics
from datetime import datetime
from urllib.parse import urlparse
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
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
CATALOGO_FILE = os.path.join(CONFIG_DIR, "catalogo_variaveis.json")
PERFIS_FILE = os.path.join(CONFIG_DIR, "perfis_usuario.json")
AUTOSAVE_FILE = os.path.join(CONFIG_DIR, "autosave_projeto.json")
ANEXOS_TEMP_DIR = os.path.join(CONFIG_DIR, "anexos_temp")
os.makedirs(ANEXOS_TEMP_DIR, exist_ok=True)

MAX_DOWNLOAD_IMAGEM_BYTES = 10 * 1024 * 1024  # 10 MB
LANCZOS_FILTER = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.BICUBIC)

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

OPCOES_SITUACAO = ["", "Ótima", "Muito Boa", "Boa", "Desfavorável", "Má", "Péssima"]

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

CATALOGO_VARIAVEIS_PADRAO = [
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

DEFAULT_CONFIG_WORD = {
    "titulo_padrao": "PESQUISA DE MERCADO - MEMÓRIA DE CÁLCULO",
    "contrato_copasa": "CONTRATO: COPASA I",
    "rodape": ""
}


def data_hoje_br():
    return datetime.now().strftime("%d/%m/%Y")


def normalizar_chave_var(nome):
    return re.sub(r"\s+", " ", str(nome).strip().lower())


def normalizar_unidade(un_str, padrao="m²"):
    s = sanitizar_texto_celula(un_str).lower()
    if "ha" in s or "hectare" in s:
        return "ha"
    if "m" in s:
        return "m²"
    return padrao


def converter_area(valor, unidade_origem, unidade_destino):
    """Converte área numericamente entre m² e ha (1 ha = 10.000 m²) preservando precisão."""
    try:
        val = float(valor)
    except (TypeError, ValueError):
        return 0.0
    if val <= 0:
        return 0.0
    orig = normalizar_unidade(unidade_origem)
    dest = normalizar_unidade(unidade_destino)
    if orig == dest:
        return val
    if orig == "m²" and dest == "ha":
        return round(val / 10000.0, 6)
    if orig == "ha" and dest == "m²":
        return round(val * 10000.0, 4)
    return val


def salvar_json_atomico(caminho, dados, criar_backup=False):
    """Grava JSON em arquivo temporário, sincroniza no disco, cria backup opcional e substitui atomicamente."""
    dir_dest = os.path.dirname(os.path.abspath(caminho))
    os.makedirs(dir_dest, exist_ok=True)
    tmp_path = f"{caminho}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)
        f.flush()
        os.fsync(f.fileno())
    if criar_backup and os.path.exists(caminho):
        bak_path = f"{caminho}.bak"
        try:
            shutil.copy2(caminho, bak_path)
        except Exception as e:
            logging.warning(f"Não foi possível criar backup '{bak_path}': {e}")
    os.replace(tmp_path, caminho)


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
    return s[:-2] if re.match(r"^\d+\.0$", s) else s


def calcular_nota_agronomica_flexivel(variaveis_extras):
    apt, sit = "", ""
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
    mapa_sit = TABELA_NOTA_AGRONOMICA.get(m.group(1).upper(), {})
    sit_clean = sit.strip().lower()
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
    chave_nota = next((v["nome"] for v in variaveis_config if normalizar_chave_var(v["nome"]).startswith("nota agro")), None)
    if not chave_nota:
        chave_nota = next((k for k in extras if normalizar_chave_var(k).startswith("nota agro")), None)
    if chave_nota:
        extras[chave_nota] = nota_auto


def extrair_dados_anuncio(texto):
    res = {"valor_total": None, "area_terreno": None, "unidade": None, "telefone": None, "link": None}
    if not texto:
        return res
    txt = str(texto)

    m_link = re.search(r"https?://[^\s<>\"']+", txt)
    if m_link:
        res["link"] = m_link.group(0).rstrip(".,;)")

    m_tel = re.search(r"(?:\+?55\s*)?(?:\(?\d{2}\)?\s*)?(?:9\d{4}|\d{4})[-\s]?\d{4}", txt)
    if m_tel:
        res["telefone"] = m_tel.group(0).strip()

    m_val = re.search(r"R\$\s*([\d.,]+)\s*(milh(?:ão|ões|ao|oes)|mil)?", txt, re.IGNORECASE)
    if m_val:
        val_num = converter_para_float(m_val.group(1))
        mult = (m_val.group(2) or "").lower()
        if "milh" in mult:
            val_num *= 1_000_000
        elif "mil" in mult:
            val_num *= 1_000
        if val_num > 0:
            res["valor_total"] = val_num

    m_ha = re.search(r"([\d.,]+)\s*(?:ha\b|hectares?\b)", txt, re.IGNORECASE)
    m_m2 = re.search(r"([\d.,]+)\s*(?:m²|m2|metros?\s*quadrados?)", txt, re.IGNORECASE)
    if m_ha:
        area_num = converter_para_float(m_ha.group(1), permitir_milhar_ponto=False)
        if area_num > 0:
            res["area_terreno"] = area_num
            res["unidade"] = "ha"
    elif m_m2:
        area_num = converter_para_float(m_m2.group(1), permitir_milhar_ponto=True)
        if area_num > 0:
            res["area_terreno"] = area_num
            res["unidade"] = "m²"
    return res


def converter_para_float_validado(texto, permitir_milhar_ponto=True):
    """Retorna (valor_float, valido_bool), diferenciando entrada vazia/zero de texto inválido."""
    if texto is None:
        return 0.0, True
    if isinstance(texto, (int, float)):
        val_f = float(texto)
        if math.isnan(val_f) or math.isinf(val_f):
            return 0.0, False
        return val_f, True
    s = str(texto).replace("R$", "").replace("m²", "").replace("ha", "").strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return 0.0, True
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
        return float(s), True
    except ValueError:
        return 0.0, False


def converter_para_float(texto, permitir_milhar_ponto=True):
    val, _ = converter_para_float_validado(texto, permitir_milhar_ponto=permitir_milhar_ponto)
    return val


def formatar_moeda_br(valor):
    try:
        return f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def formatar_numero_br(valor, casas=2):
    try:
        return f"{{:,.{casas}f}}".format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"


def limpar_sufixo_coord(coord_str):
    s = str(coord_str).strip()
    s = re.sub(r'\s*m\s*[ESes]\s*$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s*[ESes]\s*$', '', s, flags=re.IGNORECASE)
    return re.sub(r'\s*m\s*$', '', s, flags=re.IGNORECASE).strip()


def calcular_percentil(valores_ordenados, p):
    """Interpolação linear padrão para quartis (0 <= p <= 100)."""
    n = len(valores_ordenados)
    if n == 0:
        return 0.0
    if n == 1:
        return valores_ordenados[0]
    k = (n - 1) * (p / 100.0)
    f = math.floor(k)
    c = math.ceil(k)
    if f == c:
        return valores_ordenados[int(k)]
    return valores_ordenados[f] + (k - f) * (valores_ordenados[c] - valores_ordenados[f])


def calcular_estatisticas_completas(valores):
    if not valores:
        return None
    ord_v = sorted(valores)
    n = len(ord_v)
    media = statistics.mean(ord_v)
    mediana = statistics.median(ord_v)
    dp = statistics.stdev(ord_v) if n > 1 else 0.0
    cv = (dp / media * 100.0) if media > 0 else 0.0
    q1 = calcular_percentil(ord_v, 25)
    q2 = calcular_percentil(ord_v, 50)
    q3 = calcular_percentil(ord_v, 75)
    iqr = q3 - q1
    return {
        "n": n,
        "minimo": ord_v[0],
        "maximo": ord_v[-1],
        "media": media,
        "mediana": mediana,
        "dp": dp,
        "cv": cv,
        "q1": q1,
        "q2": q2,
        "q3": q3,
        "iqr": iqr,
        "lim_inf_iqr": q1 - 1.5 * iqr,
        "lim_sup_iqr": q3 + 1.5 * iqr,
        "lim_inf_30": media * 0.70,
        "lim_sup_30": media * 1.30,
    }


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


def validar_url_http(url):
    try:
        parsed = urlparse(url)
        return parsed.scheme in ("http", "https") and bool(parsed.netloc)
    except Exception:
        return False


def normalizar_url_gdrive(url):
    if not url:
        return ""
    url = url.strip()
    match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url) or re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    return f"https://drive.google.com/uc?export=download&id={match.group(1)}" if match else url


def normalizar_dado(dado, variaveis_config=None, unidade_padrao="m²"):
    d_id = dado.get("dado_id") if dado.get("dado_id") is not None else dado.get("D.", 1)
    utilizado = dado.get("utilizado") if "utilizado" in dado else dado.get("Utilizado", True)
    if isinstance(utilizado, str):
        utilizado = utilizado.strip().lower() not in ("false", "0", "não", "nao", "☐", "não utilizado no modelo")

    unidade = normalizar_unidade(dado.get("unidade") or dado.get("Unidade") or unidade_padrao, padrao=unidade_padrao)
    v_tot = dado.get("valor_total")
    if v_tot is None:
        v_tot = dado.get("Valor Total (R$)") or dado.get("Valor do Aluguel (R$)") or 0.0
    v_tot = converter_para_float(v_tot)

    a_ter = dado.get("area_terreno")
    if a_ter is None:
        a_ter = dado.get(f"Área Terreno ({unidade})") or dado.get("Área Terreno (ha)") or dado.get("Área Terreno (m²)") or 0.0
    a_ter = converter_para_float(a_ter, permitir_milhar_ponto=(unidade != "ha"))

    a_const = converter_para_float(dado.get("area_construida") if dado.get("area_construida") is not None else dado.get("Área Construída (m²)", 0.0))
    extras_raw = dado.get("variaveis_extras") or dado.get("VariaveisExtras") or {}

    reg = {
        "dado_id": int(d_id),
        "utilizado": bool(utilizado),
        "informante": sanitizar_texto_celula(dado.get("informante") or dado.get("Informante") or ""),
        "telefone": sanitizar_telefone(dado.get("telefone") or dado.get("Telefone") or ""),
        "endereco": sanitizar_texto_celula(dado.get("endereco") or dado.get("Endereço") or ""),
        "bairro": sanitizar_texto_celula(dado.get("bairro") or dado.get("Bairro") or ""),
        "municipio": sanitizar_texto_celula(dado.get("municipio") or dado.get("Município") or ""),
        "valor_total": float(v_tot),
        "area_terreno": float(a_ter),
        "area_construida": float(a_const),
        "unitario": float(v_tot / a_ter) if a_ter > 0 else 0.0,
        "unidade": unidade,
        "localizacao": sanitizar_texto_celula(dado.get("localizacao") or dado.get("Localização") or ("Rural" if unidade == "ha" else "Urbana")),
        "zona_utm": sanitizar_texto_celula(dado.get("zona_utm") or dado.get("Zona UTM") or ""),
        "coord_e": limpar_sufixo_coord(sanitizar_texto_celula(dado.get("coord_e") or dado.get("Coord. E (m)") or "")),
        "coord_s": limpar_sufixo_coord(sanitizar_texto_celula(dado.get("coord_s") or dado.get("Coord. S (m)") or "")),
        "data": sanitizar_texto_celula(dado.get("data") or dado.get("Data") or data_hoje_br()) or data_hoje_br(),
        "link": sanitizar_texto_celula(dado.get("link") or dado.get("Link") or ""),
        "foto1": sanitizar_texto_celula(dado.get("foto1") or dado.get("Foto1") or ""),
        "foto2": sanitizar_texto_celula(dado.get("foto2") or dado.get("Foto2") or ""),
        "variaveis_extras": {k: sanitizar_texto_celula(v) for k, v in extras_raw.items()}
    }
    if variaveis_config:
        aplicar_nota_agronomica_no_dado(reg, variaveis_config)
    return reg


class AppPesquisaMercado:
    def __init__(self, root):
        self.root = root
        self.root.title("ENPROL - Sistema de Pesquisa de Mercado v6.7 (Planilha SISDEA)")
        self.root.geometry("1340x850")
        self.root.minsize(1080, 700)
        self._configurar_estilos()

        self.modelo_ativo = "PADRAO"
        self.finalidade_ativa = "VENDA"
        self.unidade_padrao_projeto = "m²"
        self.destacar_outliers = True
        self.criterio_outlier = "30%"  # "30%" ou "IQR"
        self.dados_pesquisas = []
        self.catalogo_variaveis = self._carregar_catalogo_variaveis()
        self.variaveis_config = self._carregar_config_variaveis()
        self.campos_opcionais = copy.deepcopy(DEFAULT_CAMPOS_OPCIONAIS)
        self.config_word = copy.deepcopy(DEFAULT_CONFIG_WORD)
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
                "Instale a biblioteca no Prompt de Comando (CMD):\npip install tksheet"
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

    # --- PERSISTÊNCIA ---
    def _carregar_catalogo_variaveis(self):
        base = copy.deepcopy(CATALOGO_VARIAVEIS_PADRAO)
        if os.path.exists(CATALOGO_FILE):
            try:
                with open(CATALOGO_FILE, "r", encoding="utf-8") as f:
                    salvos = json.load(f)
                chaves = {normalizar_chave_var(v["nome"]): i for i, v in enumerate(base)}
                for item in salvos:
                    ch = normalizar_chave_var(item.get("nome", ""))
                    if not ch:
                        continue
                    if ch in chaves:
                        base[chaves[ch]] = item
                    else:
                        chaves[ch] = len(base)
                        base.append(item)
            except Exception as e:
                logging.warning(f"Erro ao carregar catálogo: {e}")
        return base

    def _salvar_catalogo_variaveis(self):
        try:
            salvar_json_atomico(CATALOGO_FILE, self.catalogo_variaveis)
        except Exception as e:
            logging.warning(f"Erro ao salvar catálogo: {e}")

    def _registrar_no_catalogo_global(self, lista_vars):
        mapa = {normalizar_chave_var(v["nome"]): i for i, v in enumerate(self.catalogo_variaveis)}
        alterou = False
        for v in lista_vars:
            ch = normalizar_chave_var(v.get("nome", ""))
            if not ch:
                continue
            item = {"nome": v["nome"].strip(), "tipo": v.get("tipo", "texto"), "opcoes": list(v.get("opcoes", []))}
            if ch in mapa:
                if self.catalogo_variaveis[mapa[ch]] != item:
                    self.catalogo_variaveis[mapa[ch]] = item
                    alterou = True
            else:
                mapa[ch] = len(self.catalogo_variaveis)
                self.catalogo_variaveis.append(item)
                alterou = True
        if alterou:
            self._salvar_catalogo_variaveis()

    def _carregar_config_variaveis(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.warning(f"Erro ao carregar config de variáveis: {e}")
        return copy.deepcopy(DEFAULT_VARS_PADRAO)

    def _salvar_config_variaveis(self):
        try:
            salvar_json_atomico(CONFIG_FILE, self.variaveis_config)
        except Exception as e:
            messagebox.showerror("Erro ao salvar variáveis", str(e))

    def _carregar_perfis_usuario(self):
        if os.path.exists(PERFIS_FILE):
            try:
                with open(PERFIS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.warning(f"Erro ao carregar perfis: {e}")
        return {}

    def _salvar_perfis_usuario(self):
        try:
            salvar_json_atomico(PERFIS_FILE, self.perfis_usuario)
        except Exception as e:
            messagebox.showerror("Erro ao salvar perfis", str(e))

    def _autosave(self):
        """Gravação atômica do autosave com limpeza automática quando não há dados."""
        if not self.dados_pesquisas:
            if os.path.exists(AUTOSAVE_FILE):
                try:
                    os.remove(AUTOSAVE_FILE)
                    logging.info("Autosave limpo (projeto sem registros).")
                except Exception as e:
                    logging.error(f"Falha ao limpar autosave antigo: {e}")
            return
        try:
            payload = {
                "modelo": self.modelo_ativo,
                "finalidade": self.finalidade_ativa,
                "unidade_padrao": self.unidade_padrao_projeto,
                "campos_opcionais": self.campos_opcionais,
                "config_word": self.config_word,
                "variaveis_config": self.variaveis_config,
                "pesquisas": self.dados_pesquisas
            }
            salvar_json_atomico(AUTOSAVE_FILE, payload, criar_backup=False)
        except Exception as e:
            logging.error(f"Erro ao gravar autosave: {e}")

    def _ao_fechar_janela(self):
        self._commit_editor_ativo()
        if self.dados_pesquisas and self.alteracoes_pendentes:
            self._autosave()
            resp = messagebox.askyesnocancel("Sair do Sistema", "Existem alterações não salvas. Deseja salvar antes de sair?")
            if resp is None:
                return
            if resp is True:
                self._salvar_projeto()
        self.root.destroy()

    def _rotulo_valor_total(self):
        return "Valor do Aluguel (R$)" if self.finalidade_ativa == "ALUGUEL" else "Valor Total (R$)"

    def _atualizar_rotulo_topo(self):
        if not hasattr(self, "lbl_info_mod"):
            return
        mod = "Modelo 1: Padrão" if self.modelo_ativo == "PADRAO" else "Modelo 2: COPASA"
        fin = "Aluguel" if self.finalidade_ativa == "ALUGUEL" else "Venda"
        pend = " • *" if self.alteracoes_pendentes else ""
        self.lbl_info_mod.config(text=f"  |  {mod} • {fin} ({self.unidade_padrao_projeto}){pend}")

    # --- HUB INICIAL ---
    def _exibir_tela_hub(self):
        self._limpar_container()
        self.root.config(menu="")
        frame_hub = ttk.Frame(self.container_principal, padding=40)
        frame_hub.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(frame_hub, text="SISTEMA DE PESQUISA DE MERCADO v6.7", font=("Segoe UI", 18, "bold"), foreground="#1f3c5b").pack(pady=(0, 5))
        ttk.Label(frame_hub, text="Grade Planilha SISDEA • Conversão Real m²/ha • Estatística & Boxplot", font=("Segoe UI", 10), foreground="#555555").pack(pady=(0, 30))

        ttk.Button(frame_hub, text="📄  Criar Novo Arquivo / Projeto", style="Hub.TButton", width=44, command=self._popup_escolha_modelo).pack(pady=7)
        ttk.Button(frame_hub, text="📂  Carregar Projeto Existente (.json)", style="Hub.TButton", width=44, command=self._abrir_projeto_hub).pack(pady=7)
        ttk.Button(frame_hub, text="📥  Importar Dados de Planilha (.xlsx)", style="Hub.TButton", width=44, command=self._importar_planilha_hub).pack(pady=7)
        if os.path.exists(AUTOSAVE_FILE):
            ttk.Button(frame_hub, text="🔄  Recuperar Sessão Anterior (Autosave)", style="Accent.TButton", width=44, command=self._carregar_autosave).pack(pady=7)
        ttk.Button(frame_hub, text="🚪  Sair do Sistema", width=44, command=self._ao_fechar_janela).pack(pady=(20, 0))

    def _carregar_autosave(self):
        try:
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f:
                self._aplicar_conteudo_json(json.load(f))
            self._exibir_tela_trabalho()
            messagebox.showinfo("Recuperado", f"{len(self.dados_pesquisas)} pesquisas recuperadas do autosave!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao recuperar autosave: {e}")

    def _popup_escolha_modelo(self, callback_pos_confirmacao=None):
        if self.dados_pesquisas and self.alteracoes_pendentes:
            if not messagebox.askyesno("Novo Projeto", "Criar um novo projeto descartará os dados atuais da tela. Deseja continuar?"):
                return

        janela = tk.Toplevel(self.root)
        janela.title("Selecionar Modelo e Perfil Inicial")
        janela.geometry("510x410")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="1. Selecione o modelo de ficha e perfil de variáveis:", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(14, 6))
        var_mod = tk.StringVar(value="PADRAO_URBANO")
        ttk.Radiobutton(janela, text="Modelo 1 - Padrão Urbano (Via, Uso, Testada, PGV, PIB)", value="PADRAO_URBANO", variable=var_mod).pack(anchor="w", padx=35, pady=3)
        ttk.Radiobutton(janela, text="Modelo 1 - Padrão Rural (Aptidão, Situação, Nota Agronômica, Benfeitoria)", value="PADRAO_RURAL", variable=var_mod).pack(anchor="w", padx=35, pady=3)
        ttk.Radiobutton(janela, text="Modelo 2 - Contrato COPASA (Ficha Técnica Oficial)", value="COPASA", variable=var_mod).pack(anchor="w", padx=35, pady=3)

        cb_perfil_custom = None
        if self.perfis_usuario:
            ttk.Radiobutton(janela, text="Usar Perfil Personalizado Salvo:", value="CUSTOM", variable=var_mod).pack(anchor="w", padx=35, pady=(6, 2))
            cb_perfil_custom = ttk.Combobox(janela, values=list(self.perfis_usuario.keys()), width=38, state="readonly")
            cb_perfil_custom.current(0)
            cb_perfil_custom.pack(anchor="w", padx=55, pady=2)

        ttk.Separator(janela, orient="horizontal").pack(fill="x", padx=25, pady=10)
        ttk.Label(janela, text="2. Finalidade da Pesquisa:", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=25, pady=(0, 4))
        var_fin = tk.StringVar(value="VENDA")
        f_fin = ttk.Frame(janela)
        f_fin.pack(anchor="w", padx=35, pady=2)
        ttk.Radiobutton(f_fin, text="Venda (Valor Total / Oferta)", value="VENDA", variable=var_fin).pack(side="left", padx=(0, 18))
        ttk.Radiobutton(f_fin, text="Locação / Aluguel", value="ALUGUEL", variable=var_fin).pack(side="left")

        def confirmar():
            escolha = var_mod.get()
            self.finalidade_ativa = var_fin.get()
            self.campos_opcionais = copy.deepcopy(DEFAULT_CAMPOS_OPCIONAIS)
            self.config_word = copy.deepcopy(DEFAULT_CONFIG_WORD)
            if escolha == "COPASA":
                self.modelo_ativo, self.variaveis_config, self.unidade_padrao_projeto = "COPASA", copy.deepcopy(DEFAULT_VARS_COPASA), "m²"
            elif escolha == "PADRAO_RURAL":
                self.modelo_ativo, self.variaveis_config, self.unidade_padrao_projeto = "PADRAO", copy.deepcopy(DEFAULT_VARS_RURAL), "ha"
            elif escolha == "CUSTOM" and cb_perfil_custom:
                p = self.perfis_usuario[cb_perfil_custom.get()]
                self.modelo_ativo = p.get("modelo", "PADRAO")
                self.finalidade_ativa = p.get("finalidade", self.finalidade_ativa)
                self.variaveis_config = copy.deepcopy(p.get("variaveis_config", DEFAULT_VARS_PADRAO))
                self.unidade_padrao_projeto = normalizar_unidade(p.get("unidade_padrao", "m²"))
                self.campos_opcionais = copy.deepcopy(p.get("campos_opcionais", DEFAULT_CAMPOS_OPCIONAIS))
            else:
                self.modelo_ativo, self.variaveis_config, self.unidade_padrao_projeto = "PADRAO", copy.deepcopy(DEFAULT_VARS_PADRAO), "m²"

            self._salvar_config_variaveis()
            self.caminho_projeto_atual = None
            self.alteracoes_pendentes = False
            self.dados_pesquisas = []
            self._autosave()
            janela.destroy()
            self._exibir_tela_trabalho()
            if callback_pos_confirmacao:
                callback_pos_confirmacao()

        ttk.Button(janela, text="Avançar", style="Primary.TButton", command=confirmar).pack(pady=(16, 0))

    def _abrir_projeto_hub(self):
        if self._executar_abertura_json():
            self._exibir_tela_trabalho()

    def _importar_planilha_hub(self):
        self._popup_escolha_modelo(callback_pos_confirmacao=self._importar_planilha_excel)

    # --- TELA DE TRABALHO ---
    def _exibir_tela_trabalho(self):
        self._limpar_container()
        self._criar_menu_superior()
        self._criar_layout_planilha()
        self._reconstruir_planilha()

    def _criar_menu_superior(self):
        menubar = tk.Menu(self.root)
        m_arq = tk.Menu(menubar, tearoff=0)
        m_arq.add_command(label="🏠 Voltar ao Hub Inicial", command=self._exibir_tela_hub)
        m_arq.add_separator()
        m_arq.add_command(label="Novo Projeto...", command=self._popup_escolha_modelo)
        m_arq.add_command(label="Abrir Projeto (.json)...", command=self._abrir_projeto_menu)
        m_arq.add_command(label="Salvar Projeto (.json)", command=self._salvar_projeto)
        m_arq.add_separator()
        m_arq.add_command(label="📥 Importar Planilha Excel...", command=self._importar_planilha_excel)
        m_arq.add_separator()
        m_arq.add_command(label="Sair", command=self._ao_fechar_janela)
        menubar.add_cascade(label="Arquivo", menu=m_arq)

        m_cfg = tk.Menu(menubar, tearoff=0)
        m_cfg.add_command(label="⚙ Gerenciar Variáveis e Catálogo...", command=self._janela_config_variaveis)
        m_cfg.add_command(label="👁 Campos, Modelo, Word e Venda/Aluguel...", command=self._janela_campos_opcionais)
        m_cfg.add_command(label="💾 Salvar Configuração Atual como Novo Perfil...", command=self._salvar_como_novo_perfil)
        m_cfg.add_separator()
        m_cfg.add_command(label="Aplicar Perfil Urbano Padrão", command=lambda: self._aplicar_perfil_rapido("URBANO"))
        m_cfg.add_command(label="Aplicar Perfil Rural (Nota Agronômica)", command=lambda: self._aplicar_perfil_rapido("RURAL"))
        m_cfg.add_command(label="Aplicar Perfil COPASA", command=lambda: self._aplicar_perfil_rapido("COPASA"))
        m_cfg.add_separator()
        m_cfg.add_command(label="📊 Análise Estatística & Boxplot...", command=self._janela_analise_estatistica)
        m_cfg.add_command(label="📋 Ver Tabela 02 - Nota Agronômica", command=self._mostrar_tabela_nota_agronomica)
        menubar.add_cascade(label="Variáveis e Perfis", menu=m_cfg)
        self.root.config(menu=menubar)

    def _criar_layout_planilha(self):
        toolbar = ttk.Frame(self.container_principal, padding=(8, 4))
        toolbar.pack(fill="x")

        self.btn_add_dados = ttk.Button(toolbar, text="➕ Adicionar Dados ▾", style="Primary.TButton", command=self._menu_adicionar_dados)
        self.btn_add_dados.pack(side="left", padx=2)
        ttk.Button(toolbar, text="📥 Importar Excel", command=self._importar_planilha_excel).pack(side="left", padx=2)
        ttk.Button(toolbar, text="💾 Salvar", command=self._salvar_projeto).pack(side="left", padx=2)
        ttk.Button(toolbar, text="⚙ Variáveis", command=self._janela_config_variaveis).pack(side="left", padx=2)
        ttk.Button(toolbar, text="👁 Campos / Word", command=self._janela_campos_opcionais).pack(side="left", padx=2)
        ttk.Button(toolbar, text="🔄 Padronizar Unid.", command=self._padronizar_unidades_clique).pack(side="left", padx=2)

        ttk.Button(toolbar, text="📄 Gerar Word", style="Primary.TButton", command=self._iniciar_exportacao_word_thread).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📊 Exportar Excel", style="Accent.TButton", command=self._exportar_excel).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📷 Mídia / Fotos", command=self._alternar_painel_lateral).pack(side="right", padx=6)

        self.lbl_info_mod = ttk.Label(toolbar, text="", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b")
        self.lbl_info_mod.pack(side="left", padx=6)
        self._atualizar_rotulo_topo()

        barra_fx = ttk.Frame(self.container_principal, padding=(8, 2))
        barra_fx.pack(fill="x")
        self.lbl_fx = tk.Label(barra_fx, text="fx  [ Célula ]", width=24, anchor="w", bg="#e1e7ef", fg="#1f3c5b", font=("Segoe UI", 9, "bold"), relief="groove", padx=6)
        self.lbl_fx.pack(side="left")
        self.txt_fx = ttk.Entry(barra_fx, font=("Segoe UI", 9))
        self.txt_fx.pack(side="left", fill="x", expand=True, padx=4)
        self.txt_fx.bind("<Return>", lambda e: self._aplicar_barra_fx())

        self.pane_principal = tk.PanedWindow(self.container_principal, orient="horizontal", sashrelief="raised", bg="#d0d7de")
        self.pane_principal.pack(fill="both", expand=True, padx=8, pady=2)

        self.frame_sheet = ttk.Frame(self.pane_principal)
        self.pane_principal.add(self.frame_sheet, stretch="always")

        self.sheet = Sheet(
            self.frame_sheet, data=[], headers=[], font=("Segoe UI", 9, "normal"),
            header_font=("Segoe UI", 9, "bold"), show_x_scrollbar=True, show_y_scrollbar=True,
            empty_horizontal=0, empty_vertical=0
        )
        self.sheet.enable_bindings(
            "single_select", "row_select", "column_select", "drag_select",
            "column_width_resize", "double_click_column_resize", "row_height_resize",
            "arrowkeys", "right_click_popup_menu", "rc_select",
            "copy", "cut", "paste", "delete", "undo", "edit_cell"
        )
        self.sheet.pack(fill="both", expand=True)

        self.sheet.popup_menu_add_command("➕ Adicionar 1 Dado", lambda: self._adicionar_dados_qtd(1))
        self.sheet.popup_menu_add_command("➕ Adicionar Vários Dados...", self._perguntar_qtd_adicionar_dados)
        self.sheet.popup_menu_add_command("📋 Extrair de Texto de Anúncio...", self._janela_extrator_anuncio)
        self.sheet.popup_menu_add_command("📋 Duplicar Linha Selecionada", self._duplicar_linha_selecionada)
        self.sheet.popup_menu_add_command("☑ Alternar Utilizado / Não Utilizado", self._alternar_utilizado_selecionado)
        self.sheet.popup_menu_add_command("⬇ Preencher Para Baixo (Ctrl+D)", self._preencher_para_baixo)
        self.sheet.popup_menu_add_command("🗑 Excluir Linha(s) Selecionada(s)", self._excluir_linhas_selecionadas)

        self.sheet.extra_bindings("cell_select", self._ao_selecionar_celula)
        self.sheet.extra_bindings("end_edit_cell", self._ao_terminar_edicao_celula)
        self.sheet.extra_bindings("end_paste", self._ao_colar_planilha)
        self.sheet.extra_bindings("end_delete", self._ao_deletar_conteudo)
        self.root.bind("<Control-d>", lambda e: self._preencher_para_baixo())

        self.frame_painel_lateral = ttk.LabelFrame(self.pane_principal, text=" Mídia, Fotos & Extrator ", padding=8)
        self._construir_painel_lateral()

        self.progress_bar = ttk.Progressbar(self.container_principal, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=8, pady=1)
        self.progress_bar.pack_forget()

        frame_rodape = tk.Frame(self.container_principal, bg="#e1e7ef", bd=1, relief="solid")
        frame_rodape.pack(fill="x", padx=8, pady=(2, 6))
        self.lbl_estatisticas = tk.Label(frame_rodape, text="📊 0 amostras ativas", bg="#e1e7ef", fg="#1f3c5b", font=("Segoe UI", 9, "bold"), pady=4)
        self.lbl_estatisticas.pack(side="left", padx=10)

        frame_rodape_btns = tk.Frame(frame_rodape, bg="#e1e7ef")
        frame_rodape_btns.pack(side="right", padx=6, pady=2)

        self.var_destaque_out = tk.BooleanVar(value=self.destacar_outliers)
        tk.Checkbutton(
            frame_rodape_btns, text="Destacar", variable=self.var_destaque_out,
            bg="#e1e7ef", activebackground="#e1e7ef", font=("Segoe UI", 8), command=self._alternar_destaque_outliers
        ).pack(side="left", padx=(0, 2))

        self.var_criterio_out = tk.StringVar(value=self.criterio_outlier)
        cb_crit = ttk.Combobox(frame_rodape_btns, textvariable=self.var_criterio_out, values=["30%", "IQR"], width=5, state="readonly")
        cb_crit.pack(side="left", padx=(0, 6))
        cb_crit.bind("<<ComboboxSelected>>", lambda e: self._alternar_criterio_outlier())

        ttk.Button(frame_rodape_btns, text="📈 Boxplot / Estat.", command=self._janela_analise_estatistica).pack(side="left", padx=2)
        ttk.Button(frame_rodape_btns, text="🧹 Saneamento", command=self._desativar_outliers_criterio).pack(side="left", padx=2)
        ttk.Button(frame_rodape_btns, text="📋 Copiar p/ SISDEA", command=self._copiar_tabela_sisdea).pack(side="left", padx=2)

    def _menu_adicionar_dados(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="➕ Adicionar 1 Dado", command=lambda: self._adicionar_dados_qtd(1))
        menu.add_command(label="➕ Adicionar Vários Dados...", command=self._perguntar_qtd_adicionar_dados)
        menu.tk_popup(self.btn_add_dados.winfo_rootx(), self.btn_add_dados.winfo_rooty() + self.btn_add_dados.winfo_height())

    def _perguntar_qtd_adicionar_dados(self):
        qtd = simpledialog.askinteger("Adicionar Vários Dados", "Quantos dados (linhas) deseja adicionar?", minvalue=1, maxvalue=500, initialvalue=10, parent=self.root)
        if qtd:
            self._adicionar_dados_qtd(qtd)

    def _adicionar_dados_qtd(self, qtd):
        self._commit_editor_ativo()
        for _ in range(qtd):
            self._criar_registro_vazio(herdar_anterior=True)
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._reconstruir_planilha()
        self._autosave()

    def _padronizar_unidades_clique(self):
        """Padroniza a unidade de todas as linhas com conversão real de área (1 ha = 10.000 m²)."""
        self._commit_editor_ativo()
        alvo = "ha" if self.unidade_padrao_projeto == "m²" else "m²"
        if not self.dados_pesquisas:
            self.unidade_padrao_projeto = alvo
            self.alteracoes_pendentes = True
            self._atualizar_rotulo_topo()
            messagebox.showinfo("Unidade Padrão", f"Unidade padrão alterada para: {alvo}")
            return

        janela = tk.Toplevel(self.root)
        janela.title("Padronizar e Converter Unidades")
        janela.geometry("440x230")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text=f"Padronizar todas as {len(self.dados_pesquisas)} linhas para:", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=20, pady=(14, 8))
        var_un = tk.StringVar(value=alvo)
        ttk.Radiobutton(janela, text="Hectares (ha) — 1 ha = 10.000 m²", value="ha", variable=var_un).pack(anchor="w", padx=30, pady=3)
        ttk.Radiobutton(janela, text="Metros Quadrados (m²) — 1 m² = 0,0001 ha", value="m²", variable=var_un).pack(anchor="w", padx=30, pady=3)

        var_conv = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            janela,
            text="Converter automaticamente os valores numéricos de área\n(Ex: 50.000 m² → 5,0000 ha)",
            variable=var_conv
        ).pack(anchor="w", padx=25, pady=(12, 10))

        def aplicar():
            un_esc = var_un.get()
            fazer_conv = var_conv.get()
            self.unidade_padrao_projeto = un_esc
            convertidos = 0
            for d in self.dados_pesquisas:
                un_antiga = normalizar_unidade(d.get("unidade", self.unidade_padrao_projeto))
                a_ter = converter_para_float(d.get("area_terreno", 0.0), permitir_milhar_ponto=(un_antiga != "ha"))
                if fazer_conv and un_antiga != un_esc and a_ter > 0:
                    a_ter = converter_area(a_ter, un_antiga, un_esc)
                    convertidos += 1
                d["unidade"] = un_esc
                d["area_terreno"] = a_ter
                d["localizacao"] = "Rural" if un_esc == "ha" else "Urbana"
                v_tot = converter_para_float(d.get("valor_total", 0.0))
                d["unitario"] = v_tot / a_ter if a_ter > 0 else 0.0

            self.alteracoes_pendentes = True
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            janela.destroy()
            msg = f"Todas as linhas foram padronizadas para '{un_esc}'."
            if fazer_conv:
                msg += f"\n{convertidos} área(s) foram convertidas numericamente."
            messagebox.showinfo("Padronização Concluída", msg)

        ttk.Button(janela, text="✔ Aplicar Padronização", style="Primary.TButton", command=aplicar).pack(pady=4)

    def _commit_editor_ativo(self):
        if not hasattr(self, "sheet"):
            return
        try:
            if hasattr(self.sheet, "MT") and getattr(self.sheet.MT, "text_editor", None) is not None:
                self.sheet.MT.close_text_editor(editor_info=None)
        except Exception:
            pass
        self._sincronizar_matriz_para_dados(reconstruir=False)

    def _sincronizar_matriz_para_dados(self, reconstruir=False):
        if not hasattr(self, "sheet") or not hasattr(self, "esquema_colunas"):
            return
        matriz = self.sheet.get_sheet_data()
        while len(self.dados_pesquisas) < len(matriz):
            self._criar_registro_vazio(herdar_anterior=True)

        for r_idx, linha in enumerate(matriz):
            dado = self.dados_pesquisas[r_idx]
            dado["dado_id"] = r_idx + 1
            un_anterior = normalizar_unidade(dado.get("unidade", self.unidade_padrao_projeto))
            un_nova = un_anterior

            for c_idx, (chave, _, _) in enumerate(self.esquema_colunas):
                if c_idx >= len(linha) or chave in ("dado_id", "unitario", "status"):
                    continue
                val = linha[c_idx]
                if chave == "utilizado":
                    dado["utilizado"] = bool(val) if isinstance(val, bool) else str(val).strip().lower() not in ("false", "0", "não", "nao", "")
                elif chave == "unidade":
                    un_nova = normalizar_unidade(val, padrao=un_anterior)
                elif chave == "valor_total":
                    v_num, ok = converter_para_float_validado(val)
                    if ok:
                        dado["valor_total"] = v_num
                elif chave == "area_terreno":
                    a_num, ok = converter_para_float_validado(val, permitir_milhar_ponto=(un_anterior != "ha"))
                    if ok:
                        dado["area_terreno"] = a_num
                elif chave == "area_construida":
                    ac_num, ok = converter_para_float_validado(val)
                    if ok:
                        dado["area_construida"] = ac_num
                elif chave == "telefone":
                    dado["telefone"] = sanitizar_telefone(val)
                elif chave in ("coord_e", "coord_s"):
                    dado[chave] = limpar_sufixo_coord(sanitizar_texto_celula(val))
                elif chave.startswith("var_"):
                    dado.setdefault("variaveis_extras", {})[chave[4:]] = sanitizar_texto_celula(val)
                else:
                    dado[chave] = sanitizar_texto_celula(val)

            if un_nova != un_anterior:
                if dado.get("area_terreno", 0.0) > 0:
                    dado["area_terreno"] = converter_area(dado["area_terreno"], un_anterior, un_nova)
                dado["unidade"] = un_nova
                dado["localizacao"] = "Rural" if un_nova == "ha" else "Urbana"
            else:
                dado["unidade"] = un_anterior

            a_ter = dado.get("area_terreno", 0.0)
            dado["unitario"] = dado.get("valor_total", 0.0) / a_ter if a_ter > 0 else 0.0
            aplicar_nota_agronomica_no_dado(dado, self.variaveis_config)

        if self.dados_pesquisas:
            uns = {d.get("unidade", self.unidade_padrao_projeto) for d in self.dados_pesquisas}
            if len(uns) == 1:
                self.unidade_padrao_projeto = next(iter(uns))
                self._atualizar_rotulo_topo()

        if reconstruir:
            self._reconstruir_planilha()

    def _obter_esquema_colunas(self):
        cols = [("dado_id", "D.", 42), ("utilizado", "☑", 42), ("informante", "Informante", 140)]
        if self.campos_opcionais.get("telefone", True):
            cols.append(("telefone", "Telefone", 110))
        cols.append(("endereco", "Endereço / Logradouro", 180))
        if self.campos_opcionais.get("bairro", True):
            cols.append(("bairro", "Bairro", 115))
        cols.extend([
            ("municipio", "Município", 115),
            ("valor_total", self._rotulo_valor_total(), 130),
            ("area_terreno", "Área Terreno", 95),
            ("unidade", "Unid.", 55),
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
        un = normalizar_unidade(dado.get("unidade") or self.unidade_padrao_projeto, padrao=self.unidade_padrao_projeto)
        dado["unidade"] = un
        casas = 4 if un == "ha" else 2
        v_tot = converter_para_float(dado.get("valor_total", 0.0))
        a_ter = converter_para_float(dado.get("area_terreno", 0.0), permitir_milhar_ponto=(un != "ha"))
        u_val = v_tot / a_ter if a_ter > 0 else 0.0
        dado["unitario"] = u_val
        status_txt = f"{'📷' if (dado.get('foto1') or dado.get('foto2')) else '—'} {'📍' if (dado.get('coord_e') and dado.get('coord_s')) else '—'}"
        extras = dado.get("variaveis_extras", {})

        linha = []
        for chave, _, _ in self.esquema_colunas:
            if chave == "dado_id": linha.append(dado["dado_id"])
            elif chave == "utilizado": linha.append(bool(dado.get("utilizado", True)))
            elif chave == "valor_total": linha.append(formatar_moeda_br(v_tot) if v_tot else "")
            elif chave == "area_terreno": linha.append(formatar_numero_br(a_ter, casas) if a_ter else "")
            elif chave == "area_construida":
                ac = converter_para_float(dado.get("area_construida", 0.0))
                linha.append(formatar_numero_br(ac, 2) if ac else "")
            elif chave == "unitario": linha.append(formatar_moeda_br(u_val) if u_val else "")
            elif chave == "status": linha.append(status_txt)
            elif chave.startswith("var_"): linha.append(extras.get(chave[4:], ""))
            else: linha.append(dado.get(chave, ""))
        return linha

    def _aplicar_dropdown_aberto(self, span_obj, valores):
        try:
            self.sheet.dropdown(span_obj, values=valores, set_value=None, state="normal", validate_input=False)
        except TypeError:
            try:
                self.sheet.dropdown(span_obj, values=valores, set_value=None, state="normal")
            except TypeError:
                self.sheet.dropdown(span_obj, values=valores, set_value=None)

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
                self._aplicar_dropdown_aberto(self.sheet.span(0, c_idx, len(matriz), c_idx + 1), ["m²", "ha"])
                # Garante que o tksheet preserve a unidade real (ex: 'ha') de cada linha após criar o dropdown
                for r_idx, d in enumerate(self.dados_pesquisas):
                    self.sheet.set_cell_data(r_idx, c_idx, d.get("unidade", self.unidade_padrao_projeto), redraw=False)
            elif chave in ("dado_id", "unitario", "status"):
                self.sheet.readonly_columns(columns=[c_idx], readonly=True)
                self.sheet.highlight_columns(columns=[c_idx], bg="#eef2f6", fg="#334e68")
            elif chave.startswith("var_"):
                v_cfg = next((v for v in self.variaveis_config if v["nome"] == chave[4:]), None)
                if v_cfg and v_cfg.get("tipo") == "codigo" and v_cfg.get("opcoes") and matriz:
                    self._aplicar_dropdown_aberto(self.sheet.span(0, c_idx, len(matriz), c_idx + 1), [""] + [o for o in v_cfg["opcoes"] if o != ""])
                    for r_idx, d in enumerate(self.dados_pesquisas):
                        self.sheet.set_cell_data(r_idx, c_idx, d.get("variaveis_extras", {}).get(chave[4:], ""), redraw=False)

        self._atualizar_estatisticas_e_destaques()
        self.sheet.redraw()

    def _alternar_destaque_outliers(self):
        self.destacar_outliers = bool(self.var_destaque_out.get())
        self._atualizar_estatisticas_e_destaques()
        self.sheet.redraw()

    def _alternar_criterio_outlier(self):
        self.criterio_outlier = self.var_criterio_out.get()
        self._atualizar_estatisticas_e_destaques()
        self.sheet.redraw()

    def _eh_outlier(self, valor_unit, est, criterio=None):
        crit = criterio or self.criterio_outlier
        if crit == "IQR":
            return valor_unit < est["lim_inf_iqr"] or valor_unit > est["lim_sup_iqr"]
        return valor_unit < est["lim_inf_30"] or valor_unit > est["lim_sup_30"]

    def _atualizar_estatisticas_e_destaques(self):
        if not hasattr(self, "lbl_estatisticas"):
            return
        unitarios, unidades = [], set()
        for d in self.dados_pesquisas:
            if d.get("utilizado", True) and d.get("unitario", 0.0) > 0:
                unitarios.append(d["unitario"])
                unidades.add(d.get("unidade", self.unidade_padrao_projeto))

        total = len(self.dados_pesquisas)
        col_u = next((i for i, c in enumerate(self.esquema_colunas) if c[0] == "unitario"), None)
        if total > 0:
            self.sheet.dehighlight_rows(rows=list(range(total)), redraw=False)
            if col_u is not None:
                self.sheet.highlight_columns(columns=[col_u], bg="#eef2f6", fg="#334e68", redraw=False)

        if not unitarios:
            self.lbl_estatisticas.config(text=f"📊 Ativos: 0/{total} | Média: R$ 0,00 | Mediana: R$ 0,00 | CV: 0,00%", fg="#1f3c5b")
            return
        if len(unidades) > 1:
            self.lbl_estatisticas.config(text=f"⚠ Mistura de unidades ({', '.join(unidades)})! Use 'Padronizar Unid.' para converter e unificar.", fg="#b22222")
            return

        un_str = next(iter(unidades))
        est = calcular_estatisticas_completas(unitarios)
        outliers = 0
        for r_idx, d in enumerate(self.dados_pesquisas):
            if not d.get("utilizado", True):
                self.sheet.highlight_rows(rows=[r_idx], bg="#eeeeee", fg="#888888", redraw=False)
            elif d.get("unitario", 0.0) > 0 and self._eh_outlier(d["unitario"], est):
                outliers += 1
                if self.destacar_outliers and col_u is not None:
                    self.sheet.highlight_cells(row=r_idx, column=col_u, bg="#fff3cd", fg="#856404", redraw=False)

        lbl_crit = "IQR" if self.criterio_outlier == "IQR" else "±30%"
        alerta = f"  •  ⚠ {outliers} outlier(s) [{lbl_crit}]" if outliers else ""
        self.lbl_estatisticas.config(
            text=(
                f"📊 Ativos: {est['n']}/{total}  •  Média: R$ {formatar_moeda_br(est['media'])}/{un_str}  •  "
                f"Mediana: R$ {formatar_moeda_br(est['mediana'])}  •  DP: R$ {formatar_moeda_br(est['dp'])}  •  "
                f"CV: {formatar_numero_br(est['cv'], 2)}%  •  IQR: R$ {formatar_moeda_br(est['iqr'])}{alerta}"
            ),
            fg="#1f3c5b"
        )

    def _janela_analise_estatistica(self):
        self._commit_editor_ativo()
        ativos = [d for d in self.dados_pesquisas if d.get("utilizado", True) and d.get("unitario", 0.0) > 0]
        if len(ativos) < 2:
            messagebox.showinfo("Análise Estatística", "É necessário pelo menos 2 amostras ativas com unitário válido.")
            return
        unidades = {d.get("unidade", self.unidade_padrao_projeto) for d in ativos}
        if len(unidades) > 1:
            messagebox.showwarning("Unidades Mistas", "Padronize as unidades antes de gerar a análise estatística.")
            return

        un_str = next(iter(unidades))
        unitarios = [d["unitario"] for d in ativos]
        est = calcular_estatisticas_completas(unitarios)

        out_30 = [f"D.{d['dado_id']} (R$ {formatar_moeda_br(d['unitario'])})" for d in ativos if self._eh_outlier(d["unitario"], est, "30%")]
        out_iqr = [f"D.{d['dado_id']} (R$ {formatar_moeda_br(d['unitario'])})" for d in ativos if self._eh_outlier(d["unitario"], est, "IQR")]

        janela = tk.Toplevel(self.root)
        janela.title(f"Análise Estatística & Boxplot ({est['n']} amostras ativas)")
        janela.geometry("760x540")
        janela.transient(self.root)

        f_top = ttk.LabelFrame(janela, text=f" Resumo Estatístico (R$/{un_str}) ", padding=10)
        f_top.pack(fill="x", padx=12, pady=8)

        metricas = [
            ("Amostras Ativas:", f"{est['n']}", "Mínimo:", f"R$ {formatar_moeda_br(est['minimo'])}"),
            ("Média:", f"R$ {formatar_moeda_br(est['media'])}", "Máximo:", f"R$ {formatar_moeda_br(est['maximo'])}"),
            ("Mediana (Q2):", f"R$ {formatar_moeda_br(est['q2'])}", "Desvio Padrão:", f"R$ {formatar_moeda_br(est['dp'])}"),
            ("1º Quartil (Q1):", f"R$ {formatar_moeda_br(est['q1'])}", "Coef. Variação (CV):", f"{formatar_numero_br(est['cv'], 2)}%"),
            ("3º Quartil (Q3):", f"R$ {formatar_moeda_br(est['q3'])}", "Intervalo Interquartil (IQR):", f"R$ {formatar_moeda_br(est['iqr'])}"),
        ]
        for idx, (k1, v1, k2, v2) in enumerate(metricas):
            ttk.Label(f_top, text=k1, font=("Segoe UI", 9, "bold")).grid(row=idx, column=0, sticky="w", padx=6, pady=2)
            ttk.Label(f_top, text=v1).grid(row=idx, column=1, sticky="w", padx=(2, 24), pady=2)
            ttk.Label(f_top, text=k2, font=("Segoe UI", 9, "bold")).grid(row=idx, column=2, sticky="w", padx=6, pady=2)
            ttk.Label(f_top, text=v2).grid(row=idx, column=3, sticky="w", padx=2, pady=2)

        f_box = ttk.LabelFrame(janela, text=" Boxplot dos Valores Unitários ", padding=8)
        f_box.pack(fill="x", padx=12, pady=6)
        canvas = tk.Canvas(f_box, height=130, bg="white", highlightthickness=1, highlightbackground="#cbd5e1")
        canvas.pack(fill="x", expand=True)

        w_c = 700
        pad_x = 45
        v_min_plot = min(est["minimo"], est["lim_inf_iqr"])
        v_max_plot = max(est["maximo"], est["lim_sup_iqr"])
        span = (v_max_plot - v_min_plot) or 1.0

        def map_x(v):
            return pad_x + ((v - v_min_plot) / span) * (w_c - 2 * pad_x)

        y_mid = 62
        x_q1, x_q2, x_q3 = map_x(est["q1"]), map_x(est["q2"]), map_x(est["q3"])
        inliers = [v for v in unitarios if est["lim_inf_iqr"] <= v <= est["lim_sup_iqr"]]
        w_low = min(inliers) if inliers else est["minimo"]
        w_high = max(inliers) if inliers else est["maximo"]
        x_wl, x_wh = map_x(w_low), map_x(w_high)

        # Bigodes
        canvas.create_line(x_wl, y_mid, x_q1, y_mid, fill="#1f3c5b", width=2)
        canvas.create_line(x_q3, y_mid, x_wh, y_mid, fill="#1f3c5b", width=2)
        canvas.create_line(x_wl, y_mid - 14, x_wl, y_mid + 14, fill="#1f3c5b", width=2)
        canvas.create_line(x_wh, y_mid - 14, x_wh, y_mid + 14, fill="#1f3c5b", width=2)
        # Caixa Q1-Q3
        canvas.create_rectangle(x_q1, y_mid - 22, x_q3, y_mid + 22, fill="#dbeafe", outline="#1f3c5b", width=2)
        # Mediana
        canvas.create_line(x_q2, y_mid - 22, x_q2, y_mid + 22, fill="#b91c1c", width=3)
        # Média (losango verde)
        x_med = map_x(est["media"])
        canvas.create_polygon(x_med, y_mid - 7, x_med + 7, y_mid, x_med, y_mid + 7, x_med - 7, y_mid, fill="#15803d", outline="#14532d")

        # Pontos de outliers (IQR)
        for d in ativos:
            v = d["unitario"]
            if v < est["lim_inf_iqr"] or v > est["lim_sup_iqr"]:
                xp = map_x(v)
                canvas.create_oval(xp - 4, y_mid - 4, xp + 4, y_mid + 4, fill="#f59e0b", outline="#92400e")
                canvas.create_text(xp, y_mid - 14, text=f"D.{d['dado_id']}", font=("Segoe UI", 7, "bold"), fill="#92400e")

        canvas.create_text(x_wl, 108, text=f"Min: {formatar_moeda_br(w_low)}", font=("Segoe UI", 8), fill="#334155")
        canvas.create_text(x_q2, 22, text=f"Q2: {formatar_moeda_br(est['q2'])}", font=("Segoe UI", 8, "bold"), fill="#b91c1c")
        canvas.create_text(x_wh, 108, text=f"Max: {formatar_moeda_br(w_high)}", font=("Segoe UI", 8), fill="#334155")

        f_comp = ttk.LabelFrame(janela, text=" Comparação de Critérios de Outliers (Atenção: não exclui automaticamente) ", padding=10)
        f_comp.pack(fill="both", expand=True, padx=12, pady=6)

        ttk.Label(
            f_comp,
            text=f"• Critério ±30% da Média [R$ {formatar_moeda_br(est['lim_inf_30'])} a R$ {formatar_moeda_br(est['lim_sup_30'])}]:",
            font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        ttk.Label(f_comp, text="   " + (", ".join(out_30) if out_30 else "Nenhuma amostra fora da faixa."), wraplength=700).pack(anchor="w", pady=(0, 8))

        ttk.Label(
            f_comp,
            text=f"• Critério IQR / Tukey (1,5 × IQR) [R$ {formatar_moeda_br(est['lim_inf_iqr'])} a R$ {formatar_moeda_br(est['lim_sup_iqr'])}]:",
            font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")
        ttk.Label(f_comp, text="   " + (", ".join(out_iqr) if out_iqr else "Nenhuma amostra fora da faixa."), wraplength=700).pack(anchor="w")

    def _desativar_outliers_criterio(self):
        self._commit_editor_ativo()
        ativos = [d["unitario"] for d in self.dados_pesquisas if d.get("utilizado", True) and d.get("unitario", 0.0) > 0]
        if len(ativos) < 2:
            messagebox.showinfo("Saneamento", "É necessário pelo menos 2 amostras ativas com unitário válido.")
            return
        est = calcular_estatisticas_completas(ativos)
        lbl = "IQR (Tukey)" if self.criterio_outlier == "IQR" else "±30% da média"
        candidatos = [d for d in self.dados_pesquisas if d.get("utilizado", True) and d.get("unitario", 0.0) > 0 and self._eh_outlier(d["unitario"], est)]
        if not candidatos:
            messagebox.showinfo("Saneamento", f"Nenhuma amostra ativa está fora do critério {lbl}.")
            return
        ids_txt = ", ".join(f"D.{d['dado_id']}" for d in candidatos)
        if not messagebox.askyesno("Confirmar Saneamento", f"Critério ativo: {lbl}\n\nDeseja desativar {len(candidatos)} amostra(s) ({ids_txt})?"):
            return
        for d in candidatos:
            d["utilizado"] = False
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._reconstruir_planilha()
        self._autosave()

    def _copiar_tabela_sisdea(self):
        self._commit_editor_ativo()
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado para copiar.")
            return
        cab = ["D.", "Utilizado", "Informante", "Endereço", "Município", self._rotulo_valor_total(), "Área", "Unid."] + [v["nome"] for v in self.variaveis_config] + ["Unitário"]
        linhas = ["\t".join(cab)]
        for d in self.dados_pesquisas:
            un = d.get("unidade", self.unidade_padrao_projeto)
            cols = [
                str(d["dado_id"]), "Sim" if d.get("utilizado", True) else "Não",
                d.get("informante", ""), d.get("endereco", ""), d.get("municipio", ""),
                formatar_moeda_br(d.get("valor_total", 0.0)),
                formatar_numero_br(d.get("area_terreno", 0.0), 4 if un == "ha" else 2), un
            ] + [str(d.get("variaveis_extras", {}).get(v["nome"], "")) for v in self.variaveis_config] + [formatar_moeda_br(d.get("unitario", 0.0))]
            linhas.append("\t".join(cols))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(linhas))
        messagebox.showinfo("Copiado", f"{len(self.dados_pesquisas)} linhas copiadas! Basta colar (Ctrl+V) no Excel ou SISDEA.")

    # --- EVENTOS DA PLANILHA ---
    def _ao_selecionar_celula(self, event=None):
        sel = self.sheet.get_currently_selected()
        if not sel or sel.row is None or sel.column is None:
            return
        r, c = sel.row, sel.column
        if r >= len(self.dados_pesquisas) or c >= len(self.esquema_colunas):
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
        r, c = self.celula_ativa.get("row"), self.celula_ativa.get("col")
        if r is not None and c is not None:
            self._gravar_alteracao_celula(r, c, self.txt_fx.get())
            self.sheet.redraw()

    def _ao_terminar_edicao_celula(self, event=None):
        if event and getattr(event, "row", None) is not None and getattr(event, "column", None) is not None:
            self._gravar_alteracao_celula(event.row, event.column, getattr(event, "value", ""))
            self.sheet.redraw()

    def _gravar_alteracao_celula(self, r, c, novo_val):
        if r >= len(self.dados_pesquisas) or c >= len(self.esquema_colunas):
            return
        chave = self.esquema_colunas[c][0]
        if chave in ("dado_id", "unitario", "status"):
            return
        dado = self.dados_pesquisas[r]

        if chave == "utilizado":
            dado["utilizado"] = bool(novo_val) if isinstance(novo_val, bool) else str(novo_val).strip().lower() not in ("false", "0", "não", "nao", "")
        elif chave == "valor_total":
            val_f, ok = converter_para_float_validado(novo_val)
            if not ok:
                messagebox.showwarning("Valor Inválido", f"'{novo_val}' não é um valor numérico válido.")
            else:
                dado["valor_total"] = val_f
        elif chave == "area_terreno":
            un_atual = normalizar_unidade(dado.get("unidade", self.unidade_padrao_projeto))
            val_f, ok = converter_para_float_validado(novo_val, permitir_milhar_ponto=(un_atual != "ha"))
            if not ok:
                messagebox.showwarning("Área Inválida", f"'{novo_val}' não é uma área numérica válida.")
            else:
                dado["area_terreno"] = val_f
        elif chave == "area_construida":
            val_f, ok = converter_para_float_validado(novo_val)
            if not ok:
                messagebox.showwarning("Área Inválida", f"'{novo_val}' não é uma área numérica válida.")
            else:
                dado["area_construida"] = val_f
        elif chave == "unidade":
            un_antiga = normalizar_unidade(dado.get("unidade", self.unidade_padrao_projeto))
            un_nova = normalizar_unidade(novo_val, padrao=un_antiga)
            if un_nova != un_antiga:
                if dado.get("area_terreno", 0.0) > 0:
                    dado["area_terreno"] = converter_area(dado["area_terreno"], un_antiga, un_nova)
                dado["unidade"] = un_nova
                dado["localizacao"] = "Rural" if un_nova == "ha" else "Urbana"
        elif chave == "telefone":
            dado["telefone"] = sanitizar_telefone(novo_val)
        elif chave in ("coord_e", "coord_s"):
            dado[chave] = limpar_sufixo_coord(sanitizar_texto_celula(novo_val))
        elif chave.startswith("var_"):
            dado.setdefault("variaveis_extras", {})[chave[4:]] = sanitizar_texto_celula(novo_val)
            aplicar_nota_agronomica_no_dado(dado, self.variaveis_config)
        else:
            dado[chave] = sanitizar_texto_celula(novo_val)

        self.alteracoes_pendentes = True
        if self.dados_pesquisas:
            uns = {d.get("unidade", self.unidade_padrao_projeto) for d in self.dados_pesquisas}
            if len(uns) == 1:
                self.unidade_padrao_projeto = next(iter(uns))
        self._atualizar_rotulo_topo()

        for col_i, val_i in enumerate(self._montar_linha_sheet(dado)):
            self.sheet.set_cell_data(r, col_i, val_i, redraw=False)
        self._atualizar_estatisticas_e_destaques()
        self._autosave()

    def _ao_colar_planilha(self, event=None):
        self._sincronizar_matriz_para_dados(reconstruir=True)
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._autosave()

    def _ao_deletar_conteudo(self, event=None):
        self._sincronizar_matriz_para_dados(reconstruir=True)
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._autosave()

    def _criar_registro_vazio(self, herdar_anterior=True):
        ant = self.dados_pesquisas[-1] if (herdar_anterior and self.dados_pesquisas) else {}
        un = normalizar_unidade(ant.get("unidade") or self.unidade_padrao_projeto, padrao=self.unidade_padrao_projeto)
        reg = {
            "dado_id": len(self.dados_pesquisas) + 1, "utilizado": True,
            "informante": "", "telefone": "", "endereco": "", "bairro": "",
            "municipio": ant.get("municipio", ""), "valor_total": 0.0,
            "area_terreno": 0.0, "area_construida": 0.0, "unitario": 0.0,
            "unidade": un, "localizacao": "Rural" if un == "ha" else "Urbana",
            "zona_utm": ant.get("zona_utm", ""), "coord_e": "", "coord_s": "",
            "data": ant.get("data", data_hoje_br()), "link": "",
            "foto1": "", "foto2": "", "variaveis_extras": {}
        }
        self.dados_pesquisas.append(reg)
        return reg

    def _duplicar_linha_selecionada(self):
        sel = self.sheet.get_currently_selected()
        if not sel or sel.row is None or sel.row >= len(self.dados_pesquisas):
            return
        copia = copy.deepcopy(self.dados_pesquisas[sel.row])
        copia["dado_id"] = len(self.dados_pesquisas) + 1
        self.dados_pesquisas.append(copia)
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
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
        self._atualizar_rotulo_topo()
        self._reconstruir_planilha()
        self._autosave()

    def _preencher_para_baixo(self):
        for box in (self.sheet.get_all_selection_boxes() or []):
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
        if not linhas or not messagebox.askyesno("Excluir", f"Deseja excluir {len(linhas)} linha(s)?"):
            return
        for r in linhas:
            if r < len(self.dados_pesquisas):
                self.dados_pesquisas.pop(r)
        for i, d in enumerate(self.dados_pesquisas):
            d["dado_id"] = i + 1
        self.celula_ativa = {"row": None, "col": None, "chave": None}
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._reconstruir_planilha()
        self._autosave()

    # --- EXTRATOR DE TEXTO DE ANÚNCIO ---
    def _janela_extrator_anuncio(self):
        r = self.celula_ativa.get("row")
        if r is None or r >= len(self.dados_pesquisas):
            self._adicionar_dados_qtd(1)
            r = len(self.dados_pesquisas) - 1
            self.celula_ativa = {"row": r, "col": 2, "chave": "informante"}

        janela = tk.Toplevel(self.root)
        janela.title(f"Extrair Dados de Anúncio (Pesquisa D.{self.dados_pesquisas[r]['dado_id']})")
        janela.geometry("500x320")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Cole abaixo a descrição do anúncio (OLX, ZAP, VivaReal, WhatsApp):", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(10, 4))
        txt_box = tk.Text(janela, height=11, font=("Segoe UI", 9))
        txt_box.pack(fill="both", expand=True, padx=12, pady=4)
        txt_box.focus_set()

        def extrair_e_aplicar():
            res = extrair_dados_anuncio(txt_box.get("1.0", tk.END))
            dado = self.dados_pesquisas[r]
            achados = []
            if res["valor_total"]:
                dado["valor_total"] = res["valor_total"]; achados.append(f"Valor: R$ {formatar_moeda_br(res['valor_total'])}")
            if res["area_terreno"]:
                dado["area_terreno"] = res["area_terreno"]; achados.append(f"Área: {res['area_terreno']} {res['unidade']}")
            if res["unidade"]:
                dado["unidade"] = res["unidade"]; dado["localizacao"] = "Rural" if res["unidade"] == "ha" else "Urbana"
            if res["telefone"]:
                dado["telefone"] = res["telefone"]; achados.append(f"Telefone: {res['telefone']}")
            if res["link"]:
                dado["link"] = res["link"]; achados.append("Link encontrado")
            if not achados:
                messagebox.showwarning("Nada encontrado", "Não foi possível detectar valor, área, telefone ou link no texto.")
                return
            self.alteracoes_pendentes = True
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            janela.destroy()
            messagebox.showinfo("Extração Concluída", "Campos preenchidos:\n• " + "\n• ".join(achados))

        ttk.Button(janela, text="✨ Extrair e Preencher na Linha", style="Primary.TButton", command=extrair_e_aplicar).pack(pady=10)

    # --- PAINEL LATERAL DE MÍDIA ---
    def _construir_painel_lateral(self):
        for w in self.frame_painel_lateral.winfo_children():
            w.destroy()

        self.lbl_lat_titulo = ttk.Label(self.frame_painel_lateral, text="Selecione uma linha na grade", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b")
        self.lbl_lat_titulo.pack(anchor="w", pady=(0, 6))

        ttk.Button(self.frame_painel_lateral, text="📋 Colar Texto do Anúncio (Extrair)", style="Accent.TButton", command=self._janela_extrator_anuncio).pack(fill="x", pady=(0, 8))

        ttk.Label(self.frame_painel_lateral, text="Foto 1 (Imóvel / Drone):", font=("Segoe UI", 8, "bold")).pack(anchor="w")
        f1 = ttk.Frame(self.frame_painel_lateral)
        f1.pack(fill="x", pady=2)
        self.lat_foto1 = ttk.Entry(f1, width=20)
        self.lat_foto1.pack(side="left", fill="x", expand=True)
        ttk.Button(f1, text="...", width=3, command=lambda: self._escolher_foto_lateral("foto1")).pack(side="left", padx=2)
        ttk.Button(f1, text="📋 Print", width=7, command=lambda: self._colar_print_clipboard("foto1")).pack(side="left")
        self.lbl_thumb1 = tk.Label(self.frame_painel_lateral, text="[ Sem Foto 1 ]", bg="#e9ecef", width=32, height=6, relief="groove")
        self.lbl_thumb1.pack(fill="x", pady=(2, 8))

        ttk.Label(self.frame_painel_lateral, text="Foto 2 (Print Anúncio):", font=("Segoe UI", 8, "bold")).pack(anchor="w")
        f2 = ttk.Frame(self.frame_painel_lateral)
        f2.pack(fill="x", pady=2)
        self.lat_foto2 = ttk.Entry(f2, width=20)
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
        self.lat_foto1.delete(0, tk.END); self.lat_foto1.insert(0, dado.get("foto1", ""))
        self.lat_foto2.delete(0, tk.END); self.lat_foto2.insert(0, dado.get("foto2", ""))
        self._renderizar_miniatura(dado.get("foto1", ""), self.lbl_thumb1, 1)
        self._renderizar_miniatura(dado.get("foto2", ""), self.lbl_thumb2, 2)

    def _renderizar_miniatura(self, caminho, label_widget, slot):
        caminho_real = self._resolver_caminho_foto(caminho)
        if caminho_real and os.path.exists(caminho_real):
            try:
                img = ImageOps.exif_transpose(Image.open(caminho_real))
                img.thumbnail((240, 110), LANCZOS_FILTER)
                tk_img = ImageTk.PhotoImage(img)
                if slot == 1: self._img_thumb1 = tk_img
                else: self._img_thumb2 = tk_img
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
        entry.delete(0, tk.END); entry.insert(0, caminho)
        self._salvar_fotos_lateral()

    def _colar_print_clipboard(self, campo):
        r = self.celula_ativa.get("row")
        if r is None or r >= len(self.dados_pesquisas):
            messagebox.showwarning("Aviso", "Clique em uma linha da grade antes de colar o print.")
            return
        try:
            img = ImageGrab.grabclipboard()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler área de transferência: {e}")
            return
        if not isinstance(img, Image.Image):
            messagebox.showwarning("Sem Imagem", "Nenhuma imagem na área de transferência. Use Win+Shift+S primeiro.")
            return
        pasta = self._obter_pasta_anexos_projeto()
        os.makedirs(pasta, exist_ok=True)
        arq = os.path.join(pasta, f"D{self.dados_pesquisas[r]['dado_id']:02d}_{campo}_{datetime.now().strftime('%H%M%S')}.png")
        img.save(arq, "PNG")
        entry = self.lat_foto1 if campo == "foto1" else self.lat_foto2
        entry.delete(0, tk.END); entry.insert(0, arq)
        self._salvar_fotos_lateral()

    def _salvar_fotos_lateral(self):
        r = self.celula_ativa.get("row")
        if r is None or r >= len(self.dados_pesquisas):
            return
        dado = self.dados_pesquisas[r]
        dado["foto1"] = self.lat_foto1.get().strip()
        dado["foto2"] = self.lat_foto2.get().strip()
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._atualizar_painel_lateral(r)
        for col_i, val_i in enumerate(self._montar_linha_sheet(dado)):
            self.sheet.set_cell_data(r, col_i, val_i, redraw=False)
        self.sheet.redraw()
        self._autosave()

    def _obter_pasta_anexos_projeto(self):
        return f"{os.path.splitext(self.caminho_projeto_atual)[0]}_anexos" if self.caminho_projeto_atual else ANEXOS_TEMP_DIR

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

    # --- JANELA CAMPOS, MODELO E PERSONALIZAÇÃO DO WORD ---
    def _janela_campos_opcionais(self):
        self._commit_editor_ativo()
        janela = tk.Toplevel(self.root)
        janela.title("Campos, Modelo, Finalidade e Word")
        janela.geometry("480x555")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        f_fin = ttk.LabelFrame(janela, text=" 1. Finalidade da Pesquisa ", padding=8)
        f_fin.pack(fill="x", padx=12, pady=(10, 5))
        var_fin = tk.StringVar(value=self.finalidade_ativa)
        ttk.Radiobutton(f_fin, text="Venda (Valor Total / Oferta)", value="VENDA", variable=var_fin).pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(f_fin, text="Locação / Aluguel (Valor do Aluguel)", value="ALUGUEL", variable=var_fin).pack(anchor="w", padx=10, pady=2)

        f_mod = ttk.LabelFrame(janela, text=" 2. Modelo de Ficha Word ", padding=8)
        f_mod.pack(fill="x", padx=12, pady=5)
        var_mod = tk.StringVar(value=self.modelo_ativo)
        ttk.Radiobutton(f_mod, text="Modelo 1 - Padrão / Memória de Cálculo", value="PADRAO", variable=var_mod).pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(f_mod, text="Modelo 2 - Contrato COPASA", value="COPASA", variable=var_mod).pack(anchor="w", padx=10, pady=2)

        f_opc = ttk.LabelFrame(janela, text=" 3. Campos Opcionais na Grade, Excel e Word ", padding=8)
        f_opc.pack(fill="x", padx=12, pady=5)
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
            ttk.Checkbutton(f_opc, text=rot, variable=v).pack(anchor="w", padx=12, pady=1)
            vars_tk[chave] = v

        f_word = ttk.LabelFrame(janela, text=" 4. Personalização do Relatório Word ", padding=8)
        f_word.pack(fill="x", padx=12, pady=5)
        ttk.Label(f_word, text="Título (Modelo Padrão):").grid(row=0, column=0, sticky="w", pady=2)
        e_tit = ttk.Entry(f_word, width=36)
        e_tit.insert(0, self.config_word.get("titulo_padrao", DEFAULT_CONFIG_WORD["titulo_padrao"]))
        e_tit.grid(row=0, column=1, padx=4, pady=2)

        ttk.Label(f_word, text="Identificação COPASA:").grid(row=1, column=0, sticky="w", pady=2)
        e_cop = ttk.Entry(f_word, width=36)
        e_cop.insert(0, self.config_word.get("contrato_copasa", DEFAULT_CONFIG_WORD["contrato_copasa"]))
        e_cop.grid(row=1, column=1, padx=4, pady=2)

        ttk.Label(f_word, text="Rodapé (opcional):").grid(row=2, column=0, sticky="w", pady=2)
        e_rod = ttk.Entry(f_word, width=36)
        e_rod.insert(0, self.config_word.get("rodape", ""))
        e_rod.grid(row=2, column=1, padx=4, pady=2)

        def aplicar():
            self.finalidade_ativa = var_fin.get()
            self.modelo_ativo = var_mod.get()
            for chave, v in vars_tk.items():
                self.campos_opcionais[chave] = v.get()
            self.config_word["titulo_padrao"] = e_tit.get().strip() or DEFAULT_CONFIG_WORD["titulo_padrao"]
            self.config_word["contrato_copasa"] = e_cop.get().strip() or DEFAULT_CONFIG_WORD["contrato_copasa"]
            self.config_word["rodape"] = e_rod.get().strip()
            self.alteracoes_pendentes = True
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            janela.destroy()

        ttk.Button(janela, text="✔ Aplicar Configurações", style="Primary.TButton", command=aplicar).pack(pady=10)

    def _salvar_como_novo_perfil(self):
        nome = simpledialog.askstring("Salvar Perfil", "Nome para este perfil (ex: Rural Goianésia):", parent=self.root)
        if not nome or not nome.strip():
            return
        self.perfis_usuario[nome.strip()] = {
            "modelo": self.modelo_ativo,
            "finalidade": self.finalidade_ativa,
            "unidade_padrao": self.unidade_padrao_projeto,
            "campos_opcionais": copy.deepcopy(self.campos_opcionais),
            "variaveis_config": copy.deepcopy(self.variaveis_config)
        }
        self._salvar_perfis_usuario()
        messagebox.showinfo("Perfil Salvo", f"Perfil '{nome.strip()}' salvo permanentemente!")

    def _aplicar_perfil_rapido(self, tipo):
        self._commit_editor_ativo()
        if self.dados_pesquisas:
            resp = messagebox.askyesno(
                "Confirmar Troca de Perfil",
                f"Já existem {len(self.dados_pesquisas)} amostra(s) cadastrada(s).\n\n"
                f"Aplicar o perfil '{tipo}' alterará o modelo, as variáveis ativas e a unidade padrão do projeto.\n"
                "Os valores já digitados serão mantidos, mas a estrutura da planilha mudará.\n\n"
                "Deseja continuar?"
            )
            if not resp:
                return

        if tipo == "URBANO":
            self.modelo_ativo, self.variaveis_config, self.unidade_padrao_projeto = "PADRAO", copy.deepcopy(DEFAULT_VARS_PADRAO), "m²"
        elif tipo == "RURAL":
            self.modelo_ativo, self.variaveis_config, self.unidade_padrao_projeto = "PADRAO", copy.deepcopy(DEFAULT_VARS_RURAL), "ha"
        elif tipo == "COPASA":
            self.modelo_ativo, self.variaveis_config, self.unidade_padrao_projeto = "COPASA", copy.deepcopy(DEFAULT_VARS_COPASA), "m²"

        self.alteracoes_pendentes = True
        self._salvar_config_variaveis()
        self._atualizar_rotulo_topo()
        self._reconstruir_planilha()
        self._autosave()

    def _janela_config_variaveis(self):
        self._commit_editor_ativo()
        janela = tk.Toplevel(self.root)
        janela.title("Gerenciador de Variáveis e Catálogo Persistente")
        janela.geometry("780x550")
        janela.transient(self.root)
        janela.grab_set()

        f_perfis = ttk.LabelFrame(janela, text=" 1. Perfis Rápidos ", padding=6)
        f_perfis.pack(fill="x", padx=10, pady=6)
        entradas_vars = []

        def recarregar_linhas_editor(lista_cfg):
            for w in frame_lista.winfo_children():
                w.destroy()
            entradas_vars.clear()
            total_slots = max(10, len(lista_cfg) + 3)
            ttk.Label(frame_lista, text="Ativar", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, padx=2)
            ttk.Label(frame_lista, text="Nome da Variável", font=("Segoe UI", 8, "bold")).grid(row=0, column=1, padx=3, sticky="w")
            ttk.Label(frame_lista, text="Tipo", font=("Segoe UI", 8, "bold")).grid(row=0, column=2, padx=3)
            ttk.Label(frame_lista, text="Opções (separadas por vírgula)", font=("Segoe UI", 8, "bold")).grid(row=0, column=3, padx=3, sticky="w")
            for i in range(total_slots):
                ativa = i < len(lista_cfg)
                cfg = lista_cfg[i] if ativa else {"nome": "", "tipo": "texto", "opcoes": []}
                va = tk.BooleanVar(value=ativa)
                ttk.Checkbutton(frame_lista, text=f"Var {i+1}", variable=va).grid(row=i+1, column=0, padx=2, pady=2)
                en = ttk.Entry(frame_lista, width=22); en.insert(0, cfg.get("nome", "")); en.grid(row=i+1, column=1, padx=3, pady=2)
                ct = ttk.Combobox(frame_lista, values=["codigo", "numero", "texto"], width=9, state="readonly"); ct.set(cfg.get("tipo", "texto")); ct.grid(row=i+1, column=2, padx=3, pady=2)
                eo = ttk.Entry(frame_lista, width=44); eo.insert(0, ", ".join([o for o in cfg.get("opcoes", []) if o != ""])); eo.grid(row=i+1, column=3, padx=3, pady=2)
                entradas_vars.append((va, en, ct, eo))

        ttk.Button(f_perfis, text="🏙 Urbano", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_PADRAO))).pack(side="left", padx=4)
        ttk.Button(f_perfis, text="🌾 Rural (Agronômico)", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_RURAL))).pack(side="left", padx=4)
        ttk.Button(f_perfis, text="💧 COPASA", command=lambda: recarregar_linhas_editor(copy.deepcopy(DEFAULT_VARS_COPASA))).pack(side="left", padx=4)

        f_cat = ttk.Frame(janela, padding=(10, 2))
        f_cat.pack(fill="x")
        ttk.Label(f_cat, text="Catálogo Global (inclui suas variáveis criadas):").pack(side="left")
        cb_cat = ttk.Combobox(f_cat, values=[v["nome"] for v in self.catalogo_variaveis], width=28, state="readonly")
        cb_cat.pack(side="left", padx=6)

        def adicionar_do_catalogo():
            escolha = cb_cat.get()
            item_cat = next((v for v in self.catalogo_variaveis if v["nome"] == escolha), None)
            if not item_cat:
                return
            atuais = []
            for va, en, ct, eo in entradas_vars:
                if va.get() and en.get().strip():
                    atuais.append({"nome": en.get().strip(), "tipo": ct.get(), "opcoes": [o.strip() for o in eo.get().split(",") if o.strip()]})
            if not any(normalizar_chave_var(x["nome"]) == normalizar_chave_var(item_cat["nome"]) for x in atuais):
                atuais.append(copy.deepcopy(item_cat))
            recarregar_linhas_editor(atuais)

        ttk.Button(f_cat, text="➕ Incluir", command=adicionar_do_catalogo).pack(side="left")

        frame_lista = ttk.LabelFrame(janela, text=" 2. Variáveis Ativas (Novas variáveis são salvas automaticamente no Catálogo Global) ", padding=8)
        frame_lista.pack(fill="both", expand=True, padx=10, pady=6)
        recarregar_linhas_editor(self.variaveis_config)

        def salvar():
            novas = []
            for va, en, ct, eo in entradas_vars:
                if va.get() and en.get().strip():
                    opts = [o.strip() for o in eo.get().split(",") if o.strip()]
                    if ct.get() == "codigo" and opts:
                        opts = [""] + opts
                    novas.append({"nome": en.get().strip(), "tipo": ct.get(), "opcoes": opts})
            self.variaveis_config = novas
            self.alteracoes_pendentes = True
            self._registrar_no_catalogo_global(novas)
            self._salvar_config_variaveis()
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            janela.destroy()

        ttk.Button(janela, text="💾 Salvar e Atualizar Planilha", style="Primary.TButton", command=salvar).pack(pady=8)

    def _mostrar_tabela_nota_agronomica(self):
        janela = tk.Toplevel(self.root)
        janela.title("Tabela 02 – Nota Agronômica (Kozma 1984 x IN RFB 1877/2019)")
        janela.geometry("620x260")
        janela.transient(self.root)
        cols = ("situacao", "peso", "I", "II", "III", "IV", "V", "VI")
        t_ref = ttk.Treeview(janela, columns=cols, show="headings", height=7)
        t_ref.heading("situacao", text="Situação / Acesso"); t_ref.heading("peso", text="Peso")
        for rom, p_apt in [("I", "1,00"), ("II", "0,80"), ("III", "0,64"), ("IV", "0,51"), ("V", "0,41"), ("VI", "0,20")]:
            t_ref.heading(rom, text=f"{rom} ({p_apt})"); t_ref.column(rom, width=65, anchor="center")
        t_ref.column("situacao", width=120, anchor="w"); t_ref.column("peso", width=55, anchor="center")
        t_ref.pack(fill="both", expand=True, padx=10, pady=8)
        for sit, p_s in [("Ótima", "1,00"), ("Muito Boa", "0,95"), ("Boa", "0,90"), ("Desfavorável", "0,80"), ("Má", "0,75"), ("Péssima", "0,70")]:
            t_ref.insert("", "end", values=(sit, p_s) + tuple(TABELA_NOTA_AGRONOMICA[r][sit] for r in ("I", "II", "III", "IV", "V", "VI")))

    # --- ABRIR / SALVAR PROJETO (SEGURO E ATÔMICO) ---
    def _salvar_projeto(self):
        self._commit_editor_ativo()
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Projeto de Pesquisa", "*.json")])
        if not caminho:
            return
        self.caminho_projeto_atual = caminho
        pasta_anexos = self._obter_pasta_anexos_projeto()
        dir_proj = os.path.dirname(caminho)

        falhas_fotos = []
        for d in self.dados_pesquisas:
            for cf in ("foto1", "foto2"):
                orig = self._resolver_caminho_foto(d.get(cf, ""))
                if orig and os.path.exists(orig) and not orig.startswith(("http://", "https://")):
                    try:
                        os.makedirs(pasta_anexos, exist_ok=True)
                        dest = os.path.join(pasta_anexos, os.path.basename(orig))
                        if os.path.abspath(orig) != os.path.abspath(dest):
                            shutil.copy2(orig, dest)
                        # Só altera o caminho se a cópia estiver realmente confirmada no destino
                        if os.path.exists(dest) and os.path.getsize(dest) > 0:
                            d[cf] = os.path.relpath(dest, dir_proj)
                        else:
                            falhas_fotos.append(f"D.{d['dado_id']} ({cf}): cópia não confirmada")
                    except Exception as e:
                        logging.error(f"Falha ao copiar foto '{orig}': {e}")
                        falhas_fotos.append(f"D.{d['dado_id']} ({cf}): {e}")

        payload = {
            "modelo": self.modelo_ativo,
            "finalidade": self.finalidade_ativa,
            "unidade_padrao": self.unidade_padrao_projeto,
            "campos_opcionais": self.campos_opcionais,
            "config_word": self.config_word,
            "variaveis_config": self.variaveis_config,
            "pesquisas": self.dados_pesquisas
        }

        try:
            salvar_json_atomico(caminho, payload, criar_backup=True)
            self.alteracoes_pendentes = False
            self._atualizar_rotulo_topo()
            self._autosave()
            if falhas_fotos:
                messagebox.showwarning(
                    "Projeto Salvo com Ressalvas",
                    "O arquivo JSON foi salvo com segurança (e backup criado), mas algumas fotos não puderam ser copiadas "
                    "(seus caminhos originais foram mantidos):\n\n• " + "\n• ".join(falhas_fotos[:8])
                )
            else:
                messagebox.showinfo("Sucesso", "Projeto e anexos salvos com sucesso!")
        except Exception as e:
            logging.error(f"Erro ao salvar projeto '{caminho}': {e}")
            messagebox.showerror("Erro", f"Erro ao salvar arquivo: {e}")

    def _aplicar_conteudo_json(self, conteudo):
        self.modelo_ativo = conteudo.get("modelo", "PADRAO")
        self.finalidade_ativa = conteudo.get("finalidade", "VENDA")
        self.variaveis_config = conteudo.get("variaveis_config") or copy.deepcopy(DEFAULT_VARS_PADRAO)
        self._registrar_no_catalogo_global(self.variaveis_config)
        self.campos_opcionais = conteudo.get("campos_opcionais") or copy.deepcopy(DEFAULT_CAMPOS_OPCIONAIS)
        self.config_word = copy.deepcopy(DEFAULT_CONFIG_WORD)
        self.config_word.update(conteudo.get("config_word") or {})
        raw = conteudo.get("pesquisas", [])
        un_pad = normalizar_unidade(
            conteudo.get("unidade_padrao") or ("ha" if any(normalizar_unidade(d.get("unidade", "")) == "ha" for d in raw) else "m²")
        )
        self.unidade_padrao_projeto = un_pad
        self.dados_pesquisas = [normalizar_dado(d, self.variaveis_config, unidade_padrao=un_pad) for d in raw]
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
                self._atualizar_rotulo_topo()
                self._reconstruir_planilha()
                messagebox.showinfo("Sucesso", f"{len(self.dados_pesquisas)} pesquisas carregadas!")
            return True
        except Exception as e:
            messagebox.showerror("Erro ao Abrir", f"Arquivo inválido: {e}")
            return False

    def _abrir_projeto_menu(self):
        self._executar_abertura_json()

    # --- IMPORTAÇÃO EXCEL COM PRÉVIA ---
    def _importar_planilha_excel(self):
        self._commit_editor_ativo()
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
                if any(k in vals for k in ["informante", "endereço", "endereco", "valor total (r$)", "valor do aluguel (r$)", "área terreno (m²)", "área terreno (ha)"]):
                    header_row = idx
                    break
            df = pd.read_excel(caminho, sheet_name=sheet_name, header=header_row)
            self._mostrar_previa_importacao(df, os.path.basename(caminho))
        except Exception as e:
            messagebox.showerror("Erro na Importação", f"Falha ao ler planilha: {e}")

    def _mostrar_previa_importacao(self, df, nome_arquivo):
        janela = tk.Toplevel(self.root)
        janela.title(f"Prévia da Importação — {nome_arquivo}")
        janela.geometry("820x420")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(
            janela,
            text=f"Foram encontradas {len(df)} linha(s) e {len(df.columns)} coluna(s). Confira as primeiras linhas antes de importar:",
            font=("Segoe UI", 9, "bold")
        ).pack(anchor="w", padx=12, pady=(10, 6))

        cols_prev = list(df.columns)[:10]
        tree = ttk.Treeview(janela, columns=cols_prev, show="headings", height=12)
        for col in cols_prev:
            tree.heading(col, text=str(col))
            tree.column(col, width=110, anchor="w")
        for _, row in df.head(15).iterrows():
            tree.insert("", "end", values=[sanitizar_texto_celula(row.get(c, "")) for c in cols_prev])
        tree.pack(fill="both", expand=True, padx=12, pady=4)

        f_btn = ttk.Frame(janela, padding=8)
        f_btn.pack(fill="x")

        def confirmar():
            janela.destroy()
            self._processar_dataframe_importado(df)

        ttk.Button(f_btn, text="✖ Cancelar", command=janela.destroy).pack(side="right", padx=4)
        ttk.Button(f_btn, text="✔ Confirmar Importação", style="Primary.TButton", command=confirmar).pack(side="right", padx=4)

    def _processar_dataframe_importado(self, df):
        mapeamento, unidade_det = {}, None
        nomes_vars = {normalizar_chave_var(v["nome"]) for v in self.variaveis_config}

        for col in df.columns:
            c = normalizar_chave_var(col)
            if c in nomes_vars:
                continue
            if "informante" in c: mapeamento["informante"] = col
            elif "telefone" in c or "contato" in c: mapeamento["telefone"] = col
            elif "endereço" in c or "endereco" in c or "logradouro" in c: mapeamento["endereco"] = col
            elif "bairro" in c: mapeamento["bairro"] = col
            elif "município" in c or "municipio" in c: mapeamento["municipio"] = col
            elif "valor total" in c or "aluguel" in c or "oferta" in c: mapeamento["valor_total"] = col
            elif "área constr" in c or "area constr" in c: mapeamento["area_construida"] = col
            elif "área terreno" in c or "area terreno" in c or c.startswith(("area", "área")):
                mapeamento["area_terreno"] = col
                if "(ha)" in c: unidade_det = "ha"
                elif "(m²)" in c or "(m2)" in c: unidade_det = "m²"
            elif "coord. e" in c or "coord e" in c or "e (m)" in c: mapeamento["coord_e"] = col
            elif "coord. s" in c or "coord s" in c or "s (m)" in c: mapeamento["coord_s"] = col
            elif "zona" in c: mapeamento["zona_utm"] = col
            elif "data" in c: mapeamento["data"] = col
            elif "link" in c: mapeamento["link"] = col
            elif "unidade" in c or "unid" in c: mapeamento["unidade"] = col
            elif "utilizado" in c or "ativo" in c: mapeamento["utilizado"] = col

        novos = []
        for _, row in df.iterrows():
            if pd.isna(row.get(mapeamento.get("informante", ""))) and pd.isna(row.get(mapeamento.get("endereco", ""))):
                continue
            if "unidade" in mapeamento and not pd.isna(row.get(mapeamento["unidade"])):
                un = normalizar_unidade(row.get(mapeamento["unidade"]), padrao=self.unidade_padrao_projeto)
            else:
                un = unidade_det or self.unidade_padrao_projeto

            v_tot = converter_para_float(row.get(mapeamento.get("valor_total", ""), 0.0))
            a_ter = converter_para_float(row.get(mapeamento.get("area_terreno", ""), 0.0), permitir_milhar_ponto=(un != "ha"))
            a_con = converter_para_float(row.get(mapeamento.get("area_construida", ""), 0.0))
            util = True
            if "utilizado" in mapeamento and not pd.isna(row.get(mapeamento["utilizado"])):
                util = str(row.get(mapeamento["utilizado"])).strip().lower() not in ("não", "nao", "false", "0", "☐", "não utilizado no modelo")
            s_data = sanitizar_texto_celula(row.get(mapeamento["data"])) if "data" in mapeamento else ""

            reg = {
                "dado_id": len(self.dados_pesquisas) + len(novos) + 1, "utilizado": util,
                "informante": sanitizar_texto_celula(row.get(mapeamento.get("informante", ""), "")),
                "telefone": sanitizar_telefone(row.get(mapeamento.get("telefone", ""), "")),
                "endereco": sanitizar_texto_celula(row.get(mapeamento.get("endereco", ""), "")),
                "bairro": sanitizar_texto_celula(row.get(mapeamento.get("bairro", ""), "")),
                "municipio": sanitizar_texto_celula(row.get(mapeamento.get("municipio", ""), "")),
                "valor_total": v_tot, "area_terreno": a_ter, "area_construida": a_con,
                "unitario": v_tot / a_ter if a_ter > 0 else 0.0, "unidade": un,
                "localizacao": "Rural" if un == "ha" else "Urbana",
                "zona_utm": sanitizar_texto_celula(row.get(mapeamento.get("zona_utm", ""), "")),
                "coord_e": limpar_sufixo_coord(sanitizar_texto_celula(row.get(mapeamento.get("coord_e", ""), ""))),
                "coord_s": limpar_sufixo_coord(sanitizar_texto_celula(row.get(mapeamento.get("coord_s", ""), ""))),
                "data": s_data or data_hoje_br(),
                "link": sanitizar_texto_celula(row.get(mapeamento.get("link", ""), "")),
                "foto1": "", "foto2": "", "variaveis_extras": {}
            }
            for v_cfg in self.variaveis_config:
                alvo = normalizar_chave_var(v_cfg["nome"])
                for col_p in df.columns:
                    if normalizar_chave_var(col_p) == alvo:
                        reg["variaveis_extras"][v_cfg["nome"]] = sanitizar_texto_celula(row.get(col_p, ""))
                        break
            aplicar_nota_agronomica_no_dado(reg, self.variaveis_config)
            novos.append(reg)

        if novos:
            self.dados_pesquisas.extend(novos)
            self.alteracoes_pendentes = True
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            messagebox.showinfo("Importação Concluída", f"{len(novos)} pesquisas importadas com sucesso!")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado válido encontrado.")

    # --- EXPORTAÇÃO EXCEL FORMATADA ---
    def _exportar_excel(self):
        self._commit_editor_ativo()
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportar.")
            return
        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilha Excel", "*.xlsx")])
        if not caminho:
            return

        rot_val = self._rotulo_valor_total()
        linhas_export = []
        for d in self.dados_pesquisas:
            un = d.get("unidade", self.unidade_padrao_projeto)
            linha = {"D.": d["dado_id"], "Utilizado no Modelo": "Sim" if d.get("utilizado", True) else "Não utilizado no modelo", "Informante": d.get("informante", "")}
            if self.campos_opcionais.get("telefone", True): linha["Telefone"] = d.get("telefone", "")
            linha["Endereço"] = d.get("endereco", "")
            if self.campos_opcionais.get("bairro", True): linha["Bairro"] = d.get("bairro", "")
            linha["Município"] = d.get("municipio", "")
            linha[rot_val] = d.get("valor_total", 0.0)
            linha[f"Área Terreno ({un})"] = d.get("area_terreno", 0.0)
            if self.campos_opcionais.get("area_construida", True): linha["Área Construída (m²)"] = d.get("area_construida", 0.0)
            for v_cfg in self.variaveis_config:
                linha[v_cfg["nome"]] = d.get("variaveis_extras", {}).get(v_cfg["nome"], "")
            linha[f"Unitário (R$/{un})"] = d.get("unitario", 0.0)
            linha["Zona UTM"] = d.get("zona_utm", "")
            linha["Coord. E (m)"] = d.get("coord_e", "")
            linha["Coord. S (m)"] = d.get("coord_s", "")
            if self.campos_opcionais.get("localizacao", True): linha["Localização"] = d.get("localizacao", "Urbana")
            linha["Data"] = d.get("data", "")
            if self.campos_opcionais.get("link", True): linha["Link"] = d.get("link", "")
            linhas_export.append(linha)

        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            pd.DataFrame(linhas_export).to_excel(writer, sheet_name="Ficha de Pesquisa", index=False)
            ws = writer.sheets["Ficha de Pesquisa"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            header_fill = PatternFill(start_color="1F3C5B", end_color="1F3C5B", fill_type="solid")
            header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for col_cells in ws.columns:
                cab = str(col_cells[0].value or "")
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(max(len(str(c.value or "")) for c in col_cells) + 3, 12), 42)
                if "R$" in cab:
                    for cell in col_cells[1:]:
                        cell.number_format = '#,##0.00'
                elif "(ha)" in cab:
                    for cell in col_cells[1:]:
                        cell.number_format = '#,##0.0000'
                elif "(m²)" in cab:
                    for cell in col_cells[1:]:
                        cell.number_format = '#,##0.00'

        messagebox.showinfo("Sucesso", "Planilha Excel (SISDEA) formatada e exportada com sucesso!")

    # --- EXPORTAÇÃO WORD ---
    def _baixar_imagem(self, caminho_ou_url):
        if not caminho_ou_url:
            return None
        url = normalizar_url_gdrive(caminho_ou_url)
        try:
            if url.startswith(("http://", "https://")):
                if not validar_url_http(url):
                    logging.warning(f"URL inválida ignorada: '{url}'")
                    return None
                with requests.get(url, timeout=12, stream=True) as resp:
                    if resp.status_code != 200:
                        return None
                    ctype = resp.headers.get("Content-Type", "").lower()
                    if not ctype.startswith("image/"):
                        logging.warning(f"Conteúdo não é imagem ({ctype}): '{url}'")
                        return None
                    buf = io.BytesIO()
                    total = 0
                    for chunk in resp.iter_content(chunk_size=65536):
                        if not chunk:
                            continue
                        total += len(chunk)
                        if total > MAX_DOWNLOAD_IMAGEM_BYTES:
                            logging.warning(f"Imagem excedeu o limite de 10 MB: '{url}'")
                            return None
                        buf.write(chunk)
                    buf.seek(0)
                    return redimensionar_e_corrigir_imagem(buf)
            else:
                real = self._resolver_caminho_foto(url)
                if real and os.path.exists(real):
                    return redimensionar_e_corrigir_imagem(real)
        except Exception as e:
            logging.warning(f"Erro ao carregar imagem '{url}': {e}")
        return None

    def _iniciar_exportacao_word_thread(self):
        self._commit_editor_ativo()
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportar.")
            return
        so_ativos = False
        if any(not d.get("utilizado", True) for d in self.dados_pesquisas):
            resp = messagebox.askyesnocancel(
                "Filtro de Amostras",
                "Deseja exportar TODAS as fichas (marcando as desativadas como 'Não utilizado no modelo')?\n\n"
                "• Sim = Todas\n• Não = Somente utilizadas (☑)\n• Cancelar = Voltar"
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

        lista = [d for d in self.dados_pesquisas if (not so_ativos or d.get("utilizado", True))]
        threading.Thread(
            target=self._processar_geracao_word,
            args=(
                caminho, copy.deepcopy(lista), copy.deepcopy(self.variaveis_config),
                self.modelo_ativo, self.finalidade_ativa, copy.deepcopy(self.campos_opcionais),
                copy.deepcopy(self.config_word)
            ),
            daemon=True
        ).start()

    def _inserir_imagem_com_limite(self, paragrafo, img_stream, max_w_in, max_h_in):
        try:
            img_stream.seek(0)
            with Image.open(img_stream) as im:
                w_px, h_px = im.size
            ratio = w_px / float(h_px) if h_px else 1.0
            w_in, h_in = max_w_in, max_w_in / ratio
            if h_in > max_h_in:
                h_in, w_in = max_h_in, max_h_in * ratio
            img_stream.seek(0)
            paragrafo.add_run().add_picture(img_stream, width=Inches(w_in), height=Inches(h_in))
        except Exception:
            paragrafo.add_run("[ Erro ao inserir imagem ]")

    def _processar_geracao_word(self, caminho, dados_lista, variaveis_lista, modelo_sel, finalidade_sel, campos_opc, cfg_word):
        try:
            doc = Document()
            total = len(dados_lista)
            logo_enp = os.path.join(BASE_DIR, "logo_enprol.png")
            logo_cop = os.path.join(BASE_DIR, "logo_copasa.png")
            is_aluguel = (finalidade_sel == "ALUGUEL")

            section = doc.sections[0]
            if cfg_word.get("rodape"):
                p_rod = section.footer.paragraphs[0]
                p_rod.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_rod = p_rod.add_run(cfg_word["rodape"])
                r_rod.font.name = "Arial"
                r_rod.font.size = Pt(8)
                r_rod.font.color.rgb = RGBColor(100, 100, 100)

            if modelo_sel == "COPASA":
                section.top_margin = section.bottom_margin = Inches(0.39375)
                section.left_margin = section.right_margin = Inches(1.18125)
                sectPr = section._sectPr
                for ch in list(sectPr):
                    if ch.tag.endswith("pgBorders"):
                        sectPr.remove(ch)
                pgBrd = parse_xml(f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page"><w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/></w:pgBorders>')
                pgMar = sectPr.find(qn("w:pgMar"))
                pgMar.addnext(pgBrd) if pgMar is not None else sectPr.append(pgBrd)

                hdr = section.header
                if hdr.paragraphs and not hdr.paragraphs[0].text:
                    hdr.paragraphs[0]._element.getparent().remove(hdr.paragraphs[0]._element)

                t_logos = hdr.add_table(rows=1, cols=2, width=Inches(5.90))
                t_logos.alignment = WD_TABLE_ALIGNMENT.CENTER
                c1, c2 = t_logos.cell(0, 0), t_logos.cell(0, 1)
                c1.width = c2.width = Inches(2.95)
                p1 = c1.paragraphs[0]; p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if os.path.exists(logo_enp): p1.add_run().add_picture(logo_enp, height=Inches(0.42))
                else:
                    r = p1.add_run("ENPROL"); r.bold = True; r.font.name = "Arial"; r.font.size = Pt(13); r.font.color.rgb = RGBColor(28, 93, 153)
                p2 = c2.paragraphs[0]; p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                if os.path.exists(logo_cop): p2.add_run().add_picture(logo_cop, height=Inches(0.42))
                else:
                    r2 = p2.add_run("copasa"); r2.bold = True; r2.font.name = "Arial"; r2.font.size = Pt(13); r2.font.color.rgb = RGBColor(50, 110, 180)

                p_sep = hdr.add_paragraph()
                p_sep.paragraph_format.space_before = Pt(0); p_sep.paragraph_format.space_after = Pt(2); p_sep.paragraph_format.line_spacing = Pt(2)

                t_faixa = hdr.add_table(rows=1, cols=3, width=Inches(5.90))
                t_faixa.alignment = WD_TABLE_ALIGNMENT.CENTER; t_faixa.autofit = False
                inserir_ordenado(t_faixa._tbl.tblPr, parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'), SUC_CMAR)
                f0, f1, f2 = t_faixa.cell(0, 0), t_faixa.cell(0, 1), t_faixa.cell(0, 2)
                f0.width, f1.width, f2.width = Inches(1.77), Inches(2.16), Inches(1.97)
                for cf in (f0, f1, f2):
                    inserir_ordenado(cf._tc.get_or_add_tcPr(), parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="4472C4"/>'), SUC_TCPR[1:])
                txt_cop = cfg_word.get("contrato_copasa", DEFAULT_CONFIG_WORD["contrato_copasa"])
                for pf, txt, al in [(f0.paragraphs[0], txt_cop, WD_ALIGN_PARAGRAPH.LEFT), (f1.paragraphs[0], "", WD_ALIGN_PARAGRAPH.CENTER), (f2.paragraphs[0], "PESQUISA DE MERCADO", WD_ALIGN_PARAGRAPH.RIGHT)]:
                    pf.alignment = al; pf.paragraph_format.space_before = Pt(0); pf.paragraph_format.space_after = Pt(0); pf.paragraph_format.line_spacing = 1.0
                    if txt:
                        rf = pf.add_run(txt); rf.bold = True; rf.font.name = "Arial"; rf.font.size = Pt(10); rf.font.color.rgb = RGBColor(255, 255, 255)
                p_sp = hdr.add_paragraph(); p_sp.paragraph_format.space_before = Pt(0); p_sp.paragraph_format.space_after = Pt(4)
            else:
                section.top_margin = section.bottom_margin = Inches(0.35)
                section.left_margin = section.right_margin = Inches(0.55)

            for i in range(0, total, 2):
                if i > 0:
                    doc.add_page_break()
                if modelo_sel == "PADRAO":
                    if i == 0:
                        p_tit = doc.add_paragraph(); p_tit.paragraph_format.space_before = Pt(0); p_tit.paragraph_format.space_after = Pt(2)
                        r_t = p_tit.add_run(cfg_word.get("titulo_padrao", DEFAULT_CONFIG_WORD["titulo_padrao"])); r_t.bold = True; r_t.font.name = "Arial"; r_t.font.size = Pt(10)
                        p_bar = doc.add_paragraph(); p_bar.paragraph_format.space_before = Pt(0); p_bar.paragraph_format.space_after = Pt(4)
                        inserir_ordenado(p_bar._p.get_or_add_pPr(), parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="245D8C"/></w:pBdr>'), SUC_PPR)
                    else:
                        p_sp = doc.add_paragraph(); p_sp.paragraph_format.space_before = Pt(0); p_sp.paragraph_format.space_after = Pt(4)

                tabela = doc.add_table(rows=0, cols=2)
                tabela.alignment = WD_TABLE_ALIGNMENT.CENTER; tabela.autofit = False
                if modelo_sel == "PADRAO":
                    inserir_ordenado(tabela._tbl.tblPr, parse_xml(f'<w:tblBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:insideV w:val="none"/></w:tblBorders>'), SUC_TBLPR)
                else:
                    tabela.style = 'Table Grid'

                for dado in dados_lista[i:i+2]:
                    d_id = dado["dado_id"]
                    sufixo = "" if dado.get("utilizado", True) else " (Não utilizado no modelo)"
                    un = dado.get("unidade", "m²")
                    r_dados = tabela.add_row()
                    c_dados, c_foto = r_dados.cells[0], r_dados.cells[1]
                    c_dados.width = Inches(2.95) if modelo_sel == "COPASA" else Inches(3.70)
                    c_foto.width = Inches(3.42) if modelo_sel == "COPASA" else Inches(3.70)
                    c_dados.vertical_alignment = c_foto.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    p_dados = c_dados.paragraphs[0]
                    p_dados.paragraph_format.line_spacing = 1.08; p_dados.paragraph_format.space_before = Pt(2); p_dados.paragraph_format.space_after = Pt(0)

                    def add_line(p, label, val):
                        r1 = p.add_run(label); r1.bold = True; r1.font.name = "Arial"; r1.font.size = Pt(10)
                        r2 = p.add_run(f" {val}\n"); r2.font.name = "Arial"; r2.font.size = Pt(10)

                    zona_s = f"{dado.get('zona_utm', '').strip()} " if dado.get("zona_utm") else ""
                    ce, cs = dado.get("coord_e", ""), dado.get("coord_s", "")
                    coord_txt = f"{zona_s}{ce} m E / {cs} m S" if (ce or cs) else ""
                    lbl_val = "Valor do Aluguel:" if is_aluguel else "Valor da Oferta:"
                    lbl_unit = f"Valor Unitário Locação/{un}:" if is_aluguel else f"Valor Unitário/{un}:"

                    if modelo_sel == "PADRAO":
                        p_dados.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_top = p_dados.add_run(f"Pesquisa – {d_id:02d}{sufixo}\n\n"); r_top.bold = True; r_top.font.name = "Arial"; r_top.font.size = Pt(10)
                        pc = c_dados.add_paragraph(); pc.paragraph_format.line_spacing = 1.08; pc.paragraph_format.space_before = Pt(0); pc.paragraph_format.space_after = Pt(0)

                        add_line(pc, "Logradouro:", dado.get("endereco", ""))
                        if campos_opc.get("bairro", True) and dado.get("bairro"): add_line(pc, "Bairro:", dado.get("bairro", ""))
                        add_line(pc, "Município:", dado.get("municipio", ""))
                        if campos_opc.get("telefone", True): add_line(pc, "Contato:", f"{dado.get('telefone', '')} - {dado.get('informante', '')}")
                        else: add_line(pc, "Informante:", dado.get("informante", ""))
                        if campos_opc.get("link", True) and dado.get("link"): add_line(pc, "Link:", dado.get("link", ""))
                        pc.add_run("\n")

                        add_line(pc, "Área Terreno:", f"{formatar_numero_br(dado.get('area_terreno', 0.0), 4 if un == 'ha' else 2)} {un}")
                        if campos_opc.get("area_construida", True): add_line(pc, "Área Construída:", f"{formatar_numero_br(dado.get('area_construida', 0.0), 2)} m²")
                        add_line(pc, lbl_val, f"R$ {formatar_moeda_br(dado.get('valor_total', 0.0))}")
                        add_line(pc, lbl_unit, f"R$ {formatar_moeda_br(dado.get('unitario', 0.0))}/{un}")
                        pc.add_run("\n")

                        for v_cfg in variaveis_lista:
                            add_line(pc, f"{v_cfg['nome']}:", dado.get("variaveis_extras", {}).get(v_cfg["nome"], ""))
                        add_line(pc, "Coordenadas UTM:", coord_txt)
                        if campos_opc.get("localizacao", True): add_line(pc, "Localização:", dado.get("localizacao", "Urbana"))
                        pc.add_run("\n")
                        add_line(pc, "Data:", dado.get("data", data_hoje_br()))

                        inserir_ordenado(c_dados._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:right w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        inserir_ordenado(c_foto._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                    else:
                        add_line(p_dados, "Logradouro:", dado.get("endereco", ""))
                        if campos_opc.get("bairro", True): add_line(p_dados, "Bairro:", dado.get("bairro", ""))
                        add_line(p_dados, "Município:", dado.get("municipio", ""))
                        p_dados.add_run("\n")
                        extras = dado.get("variaveis_extras", {})
                        add_line(p_dados, "Informações:", extras.get("Informações", ""))
                        p_dados.add_run("\n")
                        add_line(p_dados, "Área Terreno:", f"{formatar_numero_br(dado.get('area_terreno', 0.0), 4 if un == 'ha' else 2)} {un}")
                        if campos_opc.get("area_construida", True): add_line(p_dados, "Área Construída:", f"{formatar_numero_br(dado.get('area_construida', 0.0), 2)} m²")
                        add_line(p_dados, "Frente:", extras.get("Frente", "Não informado"))
                        p_dados.add_run("\n")
                        add_line(p_dados, "Via de acesso:", extras.get("Via de acesso", "pavimentada"))
                        add_line(p_dados, "Coordenadas UTM:", coord_txt)
                        p_dados.add_run("\n")
                        add_line(p_dados, "Valor Unitário Locação:" if is_aluguel else "Valor Unitário:", f"R$/{un} {formatar_numero_br(dado.get('unitario', 0.0), 2)}")
                        add_line(p_dados, "Valor do Aluguel:" if is_aluguel else "Valor Total:", f"R$ {formatar_moeda_br(dado.get('valor_total', 0.0))}")
                        p_dados.add_run("\n")
                        add_line(p_dados, "Data:", dado.get("data", data_hoje_br()))

                    img1, img2 = self._baixar_imagem(dado.get("foto1")), self._baixar_imagem(dado.get("foto2"))
                    p_foto = c_foto.paragraphs[0]; p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    max_w = 3.25 if modelo_sel == "COPASA" else 2.75
                    if img1 and img2:
                        self._inserir_imagem_com_limite(p_foto, img1, max_w, 1.95)
                        p_f2 = c_foto.add_paragraph(); p_f2.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_f2.paragraph_format.space_before = Pt(2)
                        self._inserir_imagem_com_limite(p_f2, img2, max_w, 1.95)
                    elif img1 or img2:
                        self._inserir_imagem_com_limite(p_foto, img1 or img2, max_w, 3.60)
                    else:
                        rv = p_foto.add_run("[ Sem fotos anexadas ]"); rv.font.name = "Arial"; rv.font.size = Pt(10)

                    if modelo_sel == "COPASA":
                        inserir_ordenado(c_dados._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        inserir_ordenado(c_foto._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="nil"/><w:bottom w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        c_m = tabela.add_row().cells[0].merge(tabela.rows[-1].cells[1]); c_m.width = Inches(6.37)
                        inserir_ordenado(c_m._tc.get_or_add_tcPr(), parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="nil"/></w:tcBorders>'), SUC_TCPR)
                        p_lbl = c_m.paragraphs[0]; p_lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_lbl.paragraph_format.space_before = Pt(4); p_lbl.paragraph_format.space_after = Pt(4)
                        rn = p_lbl.add_run(f"Pesquisa {d_id:02d}{sufixo}"); rn.font.name = "Arial"; rn.font.size = Pt(10)

                    self.fila_eventos_ui.put(("progresso", int(((i + 1) / max(1, total)) * 100)))

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
