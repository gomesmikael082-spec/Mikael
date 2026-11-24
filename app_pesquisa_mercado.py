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
CATALOGO_FILE = os.path.join(CONFIG_DIR, "catalogo_variaveis.json")
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


def extrair_dados_anuncio(texto):
    """Extrai valor, área, unidade, telefone e link de um texto bruto de anúncio."""
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


def normalizar_dado(dado, variaveis_config=None, unidade_padrao="m²"):
    d_id = dado.get("dado_id") if dado.get("dado_id") is not None else dado.get("D.", 1)
    utilizado = dado.get("utilizado") if "utilizado" in dado else dado.get("Utilizado", True)
    if isinstance(utilizado, str):
        utilizado = utilizado.strip().lower() not in ("false", "0", "não", "nao", "☐", "não utilizado no modelo")

    informante = sanitizar_texto_celula(dado.get("informante") or dado.get("Informante") or "")
    telefone = sanitizar_telefone(dado.get("telefone") or dado.get("Telefone") or "")
    endereco = sanitizar_texto_celula(dado.get("endereco") or dado.get("Endereço") or "")
    bairro = sanitizar_texto_celula(dado.get("bairro") or dado.get("Bairro") or "")
    municipio = sanitizar_texto_celula(dado.get("municipio") or dado.get("Município") or "")
    unidade = sanitizar_texto_celula(dado.get("unidade") or dado.get("Unidade") or unidade_padrao) or unidade_padrao

    valor_total = dado.get("valor_total")
    if valor_total is None:
        valor_total = dado.get("Valor Total (R$)") or dado.get("Valor do Aluguel (R$)") or 0.0
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
        self.root.title("ENPROL - Sistema de Pesquisa de Mercado v6.6 (Planilha SISDEA)")
        self.root.geometry("1320x840")
        self.root.minsize(1060, 700)

        self._configurar_estilos()

        self.modelo_ativo = "PADRAO"
        self.finalidade_ativa = "VENDA"  # "VENDA" ou "ALUGUEL"
        self.unidade_padrao_projeto = "m²"
        self.destacar_outliers = True
        self.dados_pesquisas = []
        self.catalogo_variaveis = self._carregar_catalogo_variaveis()
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
                "A Versão 6.6 utiliza a biblioteca 'tksheet' para a grade estilo planilha SISDEA.\n\n"
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

    # --- PERSISTÊNCIA DE CATÁLOGO GLOBAL, VARIÁVEIS E PERFIS ---
    def _carregar_catalogo_variaveis(self):
        base = copy.deepcopy(CATALOGO_VARIAVEIS_PADRAO)
        if os.path.exists(CATALOGO_FILE):
            try:
                with open(CATALOGO_FILE, "r", encoding="utf-8") as f:
                    salvos = json.load(f)
                chaves_base = {normalizar_chave_var(v["nome"]): i for i, v in enumerate(base)}
                for item in salvos:
                    ch = normalizar_chave_var(item.get("nome", ""))
                    if not ch:
                        continue
                    if ch in chaves_base:
                        base[chaves_base[ch]] = item
                    else:
                        chaves_base[ch] = len(base)
                        base.append(item)
            except Exception as e:
                logging.warning(f"Erro ao carregar catálogo de variáveis: {e}")
        return base

    def _salvar_catalogo_variaveis(self):
        try:
            with open(CATALOGO_FILE, "w", encoding="utf-8") as f:
                json.dump(self.catalogo_variaveis, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.warning(f"Erro ao salvar catálogo de variáveis: {e}")

    def _registrar_no_catalogo_global(self, lista_vars):
        """Persiste automaticamente no catálogo global toda variável criada ou editada pelo usuário."""
        mapa = {normalizar_chave_var(v["nome"]): i for i, v in enumerate(self.catalogo_variaveis)}
        alterou = False
        for v in lista_vars:
            ch = normalizar_chave_var(v.get("nome", ""))
            if not ch:
                continue
            item_limpo = {
                "nome": v["nome"].strip(),
                "tipo": v.get("tipo", "texto"),
                "opcoes": list(v.get("opcoes", []))
            }
            if ch in mapa:
                if self.catalogo_variaveis[mapa[ch]] != item_limpo:
                    self.catalogo_variaveis[mapa[ch]] = item_limpo
                    alterou = True
            else:
                mapa[ch] = len(self.catalogo_variaveis)
                self.catalogo_variaveis.append(item_limpo)
                alterou = True
        if alterou:
            self._salvar_catalogo_variaveis()

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
                "finalidade": self.finalidade_ativa,
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
        self._commit_editor_ativo()
        if self.dados_pesquisas and self.alteracoes_pendentes:
            self._autosave()
            resp = messagebox.askyesnocancel("Sair do Sistema", "Existem alterações não salvas. Deseja salvar o projeto antes de sair?")
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
        mod_label = "Modelo 1: Padrão" if self.modelo_ativo == "PADRAO" else "Modelo 2: COPASA"
        fin_label = "Aluguel" if self.finalidade_ativa == "ALUGUEL" else "Venda"
        self.lbl_info_mod.config(text=f"  |  {mod_label} • {fin_label} ({self.unidade_padrao_projeto})")

    # --- HUB INICIAL ---
    def _exibir_tela_hub(self):
        self._limpar_container()
        self.root.config(menu="")

        frame_hub = ttk.Frame(self.container_principal, padding=40)
        frame_hub.place(relx=0.5, rely=0.5, anchor="center")

        ttk.Label(frame_hub, text="SISTEMA DE PESQUISA DE MERCADO v6.6", font=("Segoe UI", 18, "bold"), foreground="#1f3c5b").pack(pady=(0, 5))
        ttk.Label(frame_hub, text="Grade Planilha SISDEA • Catálogo Persistente • Venda & Locação", font=("Segoe UI", 10), foreground="#555555").pack(pady=(0, 30))

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
        ttk.Radiobutton(f_fin, text="Locação / Aluguel (Valor do Aluguel)", value="ALUGUEL", variable=var_fin).pack(side="left")

        def confirmar():
            escolha = var_mod.get()
            self.finalidade_ativa = var_fin.get()
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
                self.finalidade_ativa = p.get("finalidade", self.finalidade_ativa)
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

        ttk.Button(janela, text="Avançar", style="Primary.TButton", command=confirmar).pack(pady=(16, 0))

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
        menu_config.add_command(label="⚙ Gerenciar Variáveis e Catálogo...", command=self._janela_config_variaveis)
        menu_config.add_command(label="👁 Campos, Modelo e Venda/Aluguel...", command=self._janela_campos_opcionais)
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

        self.btn_add_dados = ttk.Button(toolbar, text="➕ Adicionar Dados ▾", style="Primary.TButton", command=self._menu_adicionar_dados)
        self.btn_add_dados.pack(side="left", padx=2)
        ttk.Button(toolbar, text="📥 Importar Excel", command=self._importar_planilha_excel).pack(side="left", padx=2)
        ttk.Button(toolbar, text="💾 Salvar", command=self._salvar_projeto).pack(side="left", padx=2)
        ttk.Button(toolbar, text="⚙ Variáveis", command=self._janela_config_variaveis).pack(side="left", padx=2)
        ttk.Button(toolbar, text="👁 Campos", command=self._janela_campos_opcionais).pack(side="left", padx=2)
        ttk.Button(toolbar, text="🔄 Padronizar Unid.", command=self._padronizar_unidades_clique).pack(side="left", padx=2)

        ttk.Button(toolbar, text="📄 Gerar Word", style="Primary.TButton", command=self._iniciar_exportacao_word_thread).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📊 Exportar Excel", style="Accent.TButton", command=self._exportar_excel).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📷 Mídia / Fotos", command=self._alternar_painel_lateral).pack(side="right", padx=6)

        self.lbl_info_mod = ttk.Label(toolbar, text="", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b")
        self.lbl_info_mod.pack(side="left", padx=6)
        self._atualizar_rotulo_topo()

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

        # 4. Rodapé Estatístico Único + Ações Rápidas Pré-SISDEA
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
        self.lbl_estatisticas.pack(side="left", padx=10)

        frame_rodape_btns = tk.Frame(frame_rodape, bg="#e1e7ef")
        frame_rodape_btns.pack(side="right", padx=6, pady=2)

        self.var_destaque_out = tk.BooleanVar(value=self.destacar_outliers)
        tk.Checkbutton(
            frame_rodape_btns,
            text="Destacar ±30%",
            variable=self.var_destaque_out,
            bg="#e1e7ef",
            activebackground="#e1e7ef",
            font=("Segoe UI", 8),
            command=self._alternar_destaque_outliers
        ).pack(side="left", padx=(0, 6))

        ttk.Button(frame_rodape_btns, text="🧹 Desativar Outliers (>±30%)", command=self._desativar_outliers_30).pack(side="left", padx=2)
        ttk.Button(frame_rodape_btns, text="📋 Copiar p/ SISDEA", command=self._copiar_tabela_sisdea).pack(side="left", padx=2)

    def _menu_adicionar_dados(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="➕ Adicionar 1 Dado", command=lambda: self._adicionar_dados_qtd(1))
        menu.add_command(label="➕ Adicionar Vários Dados...", command=self._perguntar_qtd_adicionar_dados)
        x = self.btn_add_dados.winfo_rootx()
        y = self.btn_add_dados.winfo_rooty() + self.btn_add_dados.winfo_height()
        menu.tk_popup(x, y)

    def _perguntar_qtd_adicionar_dados(self):
        qtd = simpledialog.askinteger(
            "Adicionar Vários Dados",
            "Quantos dados (linhas) deseja adicionar na grade?",
            minvalue=1,
            maxvalue=500,
            initialvalue=10,
            parent=self.root
        )
        if qtd:
            self._adicionar_dados_qtd(qtd)

    def _adicionar_dados_qtd(self, qtd):
        self._commit_editor_ativo()
        for _ in range(qtd):
            self._criar_registro_vazio(herdar_anterior=True)
        self.alteracoes_pendentes = True
        self._reconstruir_planilha()
        self._autosave()

    def _padronizar_unidades_clique(self):
        self._commit_editor_ativo()
        if not self.dados_pesquisas:
            nova_un = "ha" if self.unidade_padrao_projeto == "m²" else "m²"
            self.unidade_padrao_projeto = nova_un
            self._atualizar_rotulo_topo()
            messagebox.showinfo("Unidade Padrão", f"Unidade padrão do projeto alterada para: {nova_un}")
            return

        alvo = "ha" if self.unidade_padrao_projeto == "m²" else "m²"
        resp = messagebox.askyesnocancel(
            "Padronizar Unidade de Todas as Linhas",
            f"Qual unidade deseja aplicar a TODAS as {len(self.dados_pesquisas)} linhas?\n\n"
            f"• Sim = Aplicar '{alvo}' em todas as linhas\n"
            f"• Não = Aplicar '{self.unidade_padrao_projeto}' em todas as linhas\n"
            f"• Cancelar = Voltar"
        )
        if resp is None:
            return
        un_escolhida = alvo if resp is True else self.unidade_padrao_projeto
        self.unidade_padrao_projeto = un_escolhida
        for d in self.dados_pesquisas:
            d["unidade"] = un_escolhida
            d["localizacao"] = "Rural" if un_escolhida == "ha" else "Urbana"
            a_ter = converter_para_float(d.get("area_terreno", 0.0), permitir_milhar_ponto=(un_escolhida != "ha"))
            v_tot = converter_para_float(d.get("valor_total", 0.0))
            d["unitario"] = v_tot / a_ter if a_ter > 0 else 0.0
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._reconstruir_planilha()
        self._autosave()

    def _commit_editor_ativo(self):
        """Força o commit de qualquer célula que esteja aberta em edição no tksheet antes de salvar/exportar."""
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
            for c_idx, (chave, _, _) in enumerate(self.esquema_colunas):
                if c_idx >= len(linha) or chave in ("dado_id", "unitario", "status"):
                    continue
                val = linha[c_idx]
                if chave == "utilizado":
                    dado["utilizado"] = bool(val) if isinstance(val, bool) else str(val).strip().lower() not in ("false", "0", "não", "nao", "")
                elif chave == "unidade":
                    s_un = sanitizar_texto_celula(val).lower()
                    if "ha" in s_un:
                        dado["unidade"] = "ha"
                        dado["localizacao"] = "Rural"
                    elif "m" in s_un:
                        dado["unidade"] = "m²"
                        dado["localizacao"] = "Urbana"
                    else:
                        dado["unidade"] = dado.get("unidade") or self.unidade_padrao_projeto
                elif chave == "valor_total":
                    dado["valor_total"] = converter_para_float(val)
                elif chave == "area_terreno":
                    un = dado.get("unidade", self.unidade_padrao_projeto)
                    dado["area_terreno"] = converter_para_float(val, permitir_milhar_ponto=(un != "ha"))
                elif chave == "area_construida":
                    dado["area_construida"] = converter_para_float(val)
                elif chave == "telefone":
                    dado["telefone"] = sanitizar_telefone(val)
                elif chave in ("coord_e", "coord_s"):
                    dado[chave] = limpar_sufixo_coord(sanitizar_texto_celula(val))
                elif chave.startswith("var_"):
                    dado.setdefault("variaveis_extras", {})[chave[4:]] = sanitizar_texto_celula(val)
                else:
                    dado[chave] = sanitizar_texto_celula(val)
            a_ter = dado.get("area_terreno", 0.0)
            dado["unitario"] = dado.get("valor_total", 0.0) / a_ter if a_ter > 0 else 0.0
            aplicar_nota_agronomica_no_dado(dado, self.variaveis_config)

        if reconstruir:
            self._reconstruir_planilha()

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
            ("valor_total", self._rotulo_valor_total(), 130),
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
        un = dado.get("unidade") or self.unidade_padrao_projeto
        dado["unidade"] = un
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

    def _aplicar_dropdown_aberto(self, span_obj, valores):
        """Aplica dropdown aceitando digitação/colagem livre para evitar bloqueios ao colar do Excel."""
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
            elif chave in ("dado_id", "unitario", "status"):
                self.sheet.readonly_columns(columns=[c_idx], readonly=True)
                self.sheet.highlight_columns(columns=[c_idx], bg="#eef2f6", fg="#334e68")
            elif chave.startswith("var_"):
                nome_v = chave[4:]
                v_cfg = next((v for v in self.variaveis_config if v["nome"] == nome_v), None)
                if v_cfg and v_cfg.get("tipo") == "codigo" and v_cfg.get("opcoes") and matriz:
                    opcoes_limpas = [""] + [o for o in v_cfg["opcoes"] if o != ""]
                    self._aplicar_dropdown_aberto(self.sheet.span(0, c_idx, len(matriz), c_idx + 1), opcoes_limpas)

        self._atualizar_estatisticas_e_destaques()
        self.sheet.redraw()

    def _alternar_destaque_outliers(self):
        self.destacar_outliers = bool(self.var_destaque_out.get())
        self._atualizar_estatisticas_e_destaques()
        self.sheet.redraw()

    def _atualizar_estatisticas_e_destaques(self):
        if not hasattr(self, "lbl_estatisticas"):
            return
        unitarios_ativos = []
        unidades_vistas = set()
        for d in self.dados_pesquisas:
            ativo = bool(d.get("utilizado", True))
            u_val = d.get("unitario", 0.0)
            if ativo and u_val > 0:
                unitarios_ativos.append(u_val)
                unidades_vistas.add(d.get("unidade", self.unidade_padrao_projeto))

        total_geral = len(self.dados_pesquisas)
        n_ativos = len(unitarios_ativos)

        col_unit_idx = next((i for i, c in enumerate(self.esquema_colunas) if c[0] == "unitario"), None)

        if total_geral > 0:
            self.sheet.dehighlight_rows(rows=list(range(total_geral)), redraw=False)
            if col_unit_idx is not None:
                self.sheet.highlight_columns(columns=[col_unit_idx], bg="#eef2f6", fg="#334e68", redraw=False)

        if n_ativos == 0:
            self.lbl_estatisticas.config(text=f"📊 Ativos: 0/{total_geral} | Média: R$ 0,00 | Mediana: R$ 0,00 | CV: 0,00%", fg="#1f3c5b")
            return

        if len(unidades_vistas) > 1:
            self.lbl_estatisticas.config(
                text=f"⚠ Mistura de unidades ({', '.join(unidades_vistas)})! Use 'Padronizar Unid.' para unificar.",
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
                    if self.destacar_outliers and col_unit_idx is not None:
                        self.sheet.highlight_cells(row=r_idx, column=col_unit_idx, bg="#fff3cd", fg="#856404", redraw=False)

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

    def _desativar_outliers_30(self):
        self._commit_editor_ativo()
        ativos = [d["unitario"] for d in self.dados_pesquisas if d.get("utilizado", True) and d.get("unitario", 0.0) > 0]
        if len(ativos) < 2:
            messagebox.showinfo("Outliers", "É necessário pelo menos 2 amostras ativas com unitário válido.")
            return
        media = statistics.mean(ativos)
        desativados = 0
        for d in self.dados_pesquisas:
            if d.get("utilizado", True) and d.get("unitario", 0.0) > 0:
                if abs(d["unitario"] - media) > 0.30 * media:
                    d["utilizado"] = False
                    desativados += 1
        if desativados > 0:
            self.alteracoes_pendentes = True
            self._reconstruir_planilha()
            self._autosave()
            messagebox.showinfo("Saneamento Concluído", f"{desativados} amostra(s) fora de ±30% da média foram marcadas como 'Não utilizado no modelo'.")
        else:
            messagebox.showinfo("Saneamento", "Nenhuma amostra ativa está fora de ±30% da média.")

    def _copiar_tabela_sisdea(self):
        self._commit_editor_ativo()
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado para copiar.")
            return
        linhas_tsv = []
        cab = ["D.", "Utilizado", "Informante", "Endereço", "Município", self._rotulo_valor_total(), "Área", "Unid."]
        for v_cfg in self.variaveis_config:
            cab.append(v_cfg["nome"])
        cab.append("Unitário")
        linhas_tsv.append("\t".join(cab))

        for d in self.dados_pesquisas:
            un = d.get("unidade", self.unidade_padrao_projeto)
            casas = 4 if un == "ha" else 2
            cols = [
                str(d["dado_id"]),
                "Sim" if d.get("utilizado", True) else "Não",
                d.get("informante", ""),
                d.get("endereco", ""),
                d.get("municipio", ""),
                formatar_moeda_br(d.get("valor_total", 0.0)),
                formatar_numero_br(d.get("area_terreno", 0.0), casas),
                un
            ]
            extras = d.get("variaveis_extras", {})
            for v_cfg in self.variaveis_config:
                cols.append(str(extras.get(v_cfg["nome"], "")))
            cols.append(formatar_moeda_br(d.get("unitario", 0.0)))
            linhas_tsv.append("\t".join(cols))

        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(linhas_tsv))
        messagebox.showinfo("Copiado", f"{len(self.dados_pesquisas)} linhas copiadas para a área de transferência! Basta colar (Ctrl+V) no Excel ou SISDEA.")

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
        self._sincronizar_matriz_para_dados(reconstruir=True)
        self.alteracoes_pendentes = True
        self._autosave()

    def _ao_deletar_conteudo(self, event=None):
        self._sincronizar_matriz_para_dados(reconstruir=True)
        self.alteracoes_pendentes = True
        self._autosave()

    def _criar_registro_vazio(self, herdar_anterior=True):
        item_id = len(self.dados_pesquisas) + 1
        ant = self.dados_pesquisas[-1] if (herdar_anterior and self.dados_pesquisas) else {}
        un = ant.get("unidade") or self.unidade_padrao_projeto
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
            "
