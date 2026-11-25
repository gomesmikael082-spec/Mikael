import os
import io
import re
import json
import math
import base64
import queue
import socket
import shutil
import zipfile
import hashlib
import logging
import ipaddress
import threading
import webbrowser
import copy
import statistics
from concurrent.futures import ThreadPoolExecutor, as_completed
from xml.sax.saxutils import escape as xml_escape
from datetime import datetime
from urllib.parse import urlparse
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from PIL import Image, ImageOps, ImageTk, ImageGrab, ImageDraw, ImageFont
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
TILES_CACHE_DIR = os.path.join(CONFIG_DIR, "tiles_cache")
os.makedirs(ANEXOS_TEMP_DIR, exist_ok=True)
os.makedirs(TILES_CACHE_DIR, exist_ok=True)

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
    {"nome": "Frente", "tipo": "texto", "opcoes": []},
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
    {"nome": "Frente", "tipo": "texto", "opcoes": []},
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
    "rodape": "",
    "fator_alqueire_ha": 4.84
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


def normalizar_nota_str(val):
    """Converte entrada de Nota Agronômica (ex: '0.9', 0.8, '0,76') para formato padrão '0,90'."""
    if val is None:
        return ""
    s = sanitizar_texto_celula(val)
    if not s:
        return ""
    num, ok = converter_para_float_validado(s, permitir_milhar_ponto=False)
    if ok and num > 0:
        return formatar_numero_br(num, 2)
    return s


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


def buscar_combinacoes_por_nota_agronomica(nota_str):
    """Retorna lista de (classe_romana, situacao) da Tabela 02 que correspondem à Nota Agronômica informada."""
    num, ok = converter_para_float_validado(nota_str, permitir_milhar_ponto=False)
    if not ok or num <= 0:
        return []
    nota_fmt = formatar_numero_br(num, 2)
    exatos, proximos = [], []
    for classe in ("I", "II", "III", "IV", "V", "VI"):
        for sit, val_tab in TABELA_NOTA_AGRONOMICA[classe].items():
            if val_tab == nota_fmt:
                exatos.append((classe, sit))
            else:
                v_num, _ = converter_para_float_validado(val_tab, permitir_milhar_ponto=False)
                if abs(v_num - num) <= 0.0051:
                    proximos.append((classe, sit))
    return exatos if exatos else proximos


def obter_label_aptidao_configurada(classe_romana, variaveis_config):
    """Retorna o rótulo completo da classe romana conforme opções configuradas na variável Aptidão."""
    for v in variaveis_config:
        if normalizar_chave_var(v["nome"]).startswith("aptid"):
            for opt in v.get("opcoes", []):
                m = re.match(r"^\s*(VI|IV|V|III|II|I)\b", opt.strip(), re.IGNORECASE)
                if m and m.group(1).upper() == classe_romana.upper():
                    return opt
    for opt in OPCOES_APTIDAO:
        m = re.match(r"^\s*(VI|IV|V|III|II|I)\b", opt.strip(), re.IGNORECASE)
        if m and m.group(1).upper() == classe_romana.upper():
            return opt
    return classe_romana


def aplicar_nota_agronomica_no_dado(dado, variaveis_config, origem_edicao="forward", escolhedor_ambiguo=None):
    """
    Sincroniza bidirecionalmente Aptidão Agrícola + Situação <-> Nota Agronômica.
    - origem_edicao == 'forward': Aptidão/Situação editada -> atualiza Nota (ou limpa se alguma ficou vazia).
    - origem_edicao == 'reverse': Nota Agronômica editada -> preenche Aptidão + Situação correspondentes.
    - origem_edicao == 'auto': Importação/normalização -> preenche o lado faltante sem apagar nota manual.
    """
    extras = dado.setdefault("variaveis_extras", {})
    chave_nota = next((v["nome"] for v in variaveis_config if normalizar_chave_var(v["nome"]).startswith("nota agro")), None)
    if not chave_nota:
        chave_nota = next((k for k in extras if normalizar_chave_var(k).startswith("nota agro")), None)
    if not chave_nota:
        return

    chave_apt = next((v["nome"] for v in variaveis_config if normalizar_chave_var(v["nome"]).startswith("aptid")), None)
    if not chave_apt:
        chave_apt = next((k for k in extras if normalizar_chave_var(k).startswith("aptid")), None)

    chave_sit = next((v["nome"] for v in variaveis_config if normalizar_chave_var(v["nome"]).startswith(("situa", "acess"))), None)
    if not chave_sit:
        chave_sit = next((k for k in extras if normalizar_chave_var(k).startswith(("situa", "acess"))), None)

    apt_atual = sanitizar_texto_celula(extras.get(chave_apt, "")) if chave_apt else ""
    sit_atual = sanitizar_texto_celula(extras.get(chave_sit, "")) if chave_sit else ""
    nota_atual = sanitizar_texto_celula(extras.get(chave_nota, ""))

    if origem_edicao == "forward":
        if not apt_atual or not sit_atual:
            extras[chave_nota] = ""
        else:
            nota_auto = calcular_nota_agronomica_flexivel(extras)
            if nota_auto:
                extras[chave_nota] = nota_auto
        return

    if origem_edicao == "reverse":
        if not nota_atual:
            return
        nota_fmt = normalizar_nota_str(nota_atual)
        extras[chave_nota] = nota_fmt
        candidatos = buscar_combinacoes_por_nota_agronomica(nota_fmt)
        if not candidatos:
            return

        m_apt = re.match(r"^\s*(VI|IV|V|III|II|I)\b", apt_atual, re.IGNORECASE)
        classe_existente = m_apt.group(1).upper() if m_apt else ""
        escolhido = None

        if len(candidatos) == 1:
            escolhido = candidatos[0]
        else:
            for cl, st in candidatos:
                if classe_existente and cl == classe_existente:
                    escolhido = (cl, st)
                    break
                if sit_atual and st.lower() in sit_atual.lower():
                    escolhido = (cl, st)
                    break
            if escolhido is None and escolhedor_ambiguo is not None:
                escolhido = escolhedor_ambiguo(nota_fmt, candidatos)
            if escolhido is None:
                escolhido = candidatos[0]

        if escolhido:
            cl_esc, st_esc = escolhido
            if chave_apt:
                extras[chave_apt] = obter_label_aptidao_configurada(cl_esc, variaveis_config)
            if chave_sit:
                extras[chave_sit] = st_esc
        return

    # origem_edicao == "auto"
    if nota_atual:
        extras[chave_nota] = normalizar_nota_str(nota_atual)

    if apt_atual and sit_atual:
        nota_auto = calcular_nota_agronomica_flexivel(extras)
        if nota_auto and not extras.get(chave_nota):
            extras[chave_nota] = nota_auto
    elif extras.get(chave_nota) and (not apt_atual or not sit_atual):
        candidatos = buscar_combinacoes_por_nota_agronomica(extras[chave_nota])
        if candidatos:
            m_apt = re.match(r"^\s*(VI|IV|V|III|II|I)\b", apt_atual, re.IGNORECASE)
            classe_existente = m_apt.group(1).upper() if m_apt else ""
            escolhido = candidatos[0]
            for cl, st in candidatos:
                if (classe_existente and cl == classe_existente) or (sit_atual and st.lower() in sit_atual.lower()):
                    escolhido = (cl, st)
                    break
            if chave_apt and not apt_atual:
                extras[chave_apt] = obter_label_aptidao_configurada(escolhido[0], variaveis_config)
            if chave_sit and not sit_atual:
                extras[chave_sit] = escolhido[1]


def extrair_dados_anuncio(texto, fator_alqueire_ha=4.84):
    res = {"valor_total": None, "area_terreno": None, "unidade": None, "telefone": None, "link": None, "detalhe_area": ""}
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

    m_alq = re.search(r"([\d.,]+)\s*(?:alqueires?\b|alq\.?\b)", txt, re.IGNORECASE)
    m_ha = re.search(r"([\d.,]+)\s*(?:ha\b|hectares?\b)", txt, re.IGNORECASE)
    m_m2 = re.search(r"([\d.,]+)\s*(?:m²|m2|metros?\s*quadrados?)", txt, re.IGNORECASE)

    if m_alq:
        alq_num = converter_para_float(m_alq.group(1), permitir_milhar_ponto=True)
        if alq_num > 0:
            res["area_terreno"] = round(alq_num * float(fator_alqueire_ha), 4)
            res["unidade"] = "ha"
            res["detalhe_area"] = f" ({formatar_numero_br(alq_num, 2)} alq × {formatar_numero_br(fator_alqueire_ha, 2)} ha)"
    elif m_ha:
        area_num = converter_para_float(m_ha.group(1), permitir_milhar_ponto=True)
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
        elif permitir_milhar_ponto and len(partes) == 2 and len(partes[1]) == 3 and len(partes[0]) <= 3:
            try:
                if int(partes[0] or 0) > 0 and partes[1].isdigit():
                    s = s.replace(".", "")
            except ValueError:
                return 0.0, False
    try:
        val_f = float(s)
        if math.isnan(val_f) or math.isinf(val_f):
            return 0.0, False
        return val_f, True
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


# --- SEGURANÇA DE REDE (ANTI-SSRF) ---
def url_eh_segura_publica(url):
    """Verifica se a URL é HTTP(S) e se resolve exclusivamente para IPs públicos (bloqueia localhost/privados)."""
    try:
        parsed = urlparse(url)
        if parsed.scheme not in ("http", "https") or not parsed.hostname:
            return False
        host = parsed.hostname.strip().lower()
        if host in ("localhost", "0.0.0.0") or host.endswith(".local"):
            return False
        infos = socket.getaddrinfo(host, None)
        if not infos:
            return False
        for info in infos:
            ip_str = info[4][0]
            ip_obj = ipaddress.ip_address(ip_str)
            if (
                ip_obj.is_private
                or ip_obj.is_loopback
                or ip_obj.is_link_local
                or ip_obj.is_multicast
                or ip_obj.is_reserved
                or ip_obj.is_unspecified
            ):
                return False
        return True
    except Exception:
        return False


# --- GEOPROCESSAMENTO UTM <-> LAT/LON (SIRGAS 2000 / WGS84) ---
def parse_zona_utm(zona_str, zona_padrao="22K"):
    s = sanitizar_texto_celula(zona_str).upper().replace(" ", "")
    if not s:
        s = zona_padrao
    m = re.match(r"^(\d{1,2})([C-X]?)$", s)
    if not m:
        return None, None
    zona_num = int(m.group(1))
    if not (1 <= zona_num <= 60):
        return None, None
    banda = m.group(2)
    hemisferio_sul = True
    if banda:
        hemisferio_sul = banda < "N"
    return zona_num, hemisferio_sul


def validar_coordenadas_utm(zona_str, coord_e_str, coord_s_str):
    """Valida Zona UTM e coordenadas E/S (ou N). Retorna (valido, mensagem_erro)."""
    ce_txt = limpar_sufixo_coord(coord_e_str)
    cs_txt = limpar_sufixo_coord(coord_s_str)
    if not ce_txt and not cs_txt:
        return True, ""
    if not ce_txt or not cs_txt:
        return False, "Informe tanto Coord. E quanto Coord. S."
    ce, ok_e = converter_para_float_validado(ce_txt, permitir_milhar_ponto=False)
    cs, ok_s = converter_para_float_validado(cs_txt, permitir_milhar_ponto=False)
    if not ok_e or not ok_s:
        return False, "Coordenadas E/S devem ser numéricas."
    if not (100000.0 <= ce <= 900000.0):
        return False, f"Coord. E ({ce}) fora da faixa UTM válida (100.000 a 900.000 m)."
    if not (0.0 <= cs <= 10000000.0):
        return False, f"Coord. S ({cs}) fora da faixa UTM válida (0 a 10.000.000 m)."
    z_num, _ = parse_zona_utm(zona_str or "22K")
    if z_num is None:
        return False, f"Zona UTM '{zona_str}' inválida (use ex: 22K, 23K, 21S)."
    return True, ""


def utm_para_latlon(zona_str, coord_e, coord_s, zona_padrao="22K"):
    """Converte UTM (SIRGAS 2000 / WGS84) para (latitude, longitude) em graus decimais."""
    zona_num, hemisferio_sul = parse_zona_utm(zona_str, zona_padrao=zona_padrao)
    if zona_num is None:
        return None, None
    easting, ok_e = converter_para_float_validado(limpar_sufixo_coord(coord_e), permitir_milhar_ponto=False)
    northing, ok_s = converter_para_float_validado(limpar_sufixo_coord(coord_s), permitir_milhar_ponto=False)
    if not ok_e or not ok_s or not (100000.0 <= easting <= 900000.0) or not (0.0 <= northing <= 10000000.0):
        return None, None

    a = 6378137.0
    f = 1.0 / 298.257223563
    e2 = 2 * f - f * f
    e_prime2 = e2 / (1.0 - e2)
    k0 = 0.9996

    x = easting - 500000.0
    y = northing - 10000000.0 if hemisferio_sul else northing

    M = y / k0
    mu = M / (a * (1.0 - e2 / 4.0 - 3.0 * e2**2 / 64.0 - 5.0 * e2**3 / 256.0))
    e1 = (1.0 - math.sqrt(1.0 - e2)) / (1.0 + math.sqrt(1.0 - e2))

    J1 = 3.0 * e1 / 2.0 - 27.0 * e1**3 / 32.0
    J2 = 21.0 * e1**2 / 16.0 - 55.0 * e1**4 / 32.0
    J3 = 151.0 * e1**3 / 96.0
    J4 = 1097.0 * e1**4 / 512.0

    fp = mu + J1 * math.sin(2.0 * mu) + J2 * math.sin(4.0 * mu) + J3 * math.sin(6.0 * mu) + J4 * math.sin(8.0 * mu)

    C1 = e_prime2 * math.cos(fp)**2
    T1 = math.tan(fp)**2
    R1 = a * (1.0 - e2) / ((1.0 - e2 * math.sin(fp)**2) ** 1.5)
    N1 = a / math.sqrt(1.0 - e2 * math.sin(fp)**2)
    D = x / (N1 * k0)

    Q1 = N1 * math.tan(fp) / R1
    Q2 = (D**2) / 2.0
    Q3 = (5.0 + 3.0 * T1 + 10.0 * C1 - 4.0 * C1**2 - 9.0 * e_prime2) * (D**4) / 24.0
    Q4 = (61.0 + 90.0 * T1 + 298.0 * C1 + 45.0 * T1**2 - 252.0 * e_prime2 - 3.0 * C1**2) * (D**6) / 720.0
    lat = fp - Q1 * (Q2 - Q3 + Q4)

    Q5 = D
    Q6 = (1.0 + 2.0 * T1 + C1) * (D**3) / 6.0
    Q7 = (5.0 - 2.0 * C1 + 28.0 * T1 - 3.0 * C1**2 + 8.0 * e_prime2 + 24.0 * T1**2) * (D**5) / 120.0
    lon0 = math.radians((zona_num - 1) * 6 - 180 + 3)
    lon = lon0 + (Q5 - Q6 + Q7) / math.cos(fp)

    return round(math.degrees(lat), 7), round(math.degrees(lon), 7)


def latlon_para_utm(lat_deg, lon_deg):
    """Converte (latitude, longitude) em graus decimais para (zona_utm_str, coord_e, coord_s) com validação geográfica."""
    if any(math.isnan(v) or math.isinf(v) for v in (lat_deg, lon_deg)):
        raise ValueError("Latitude e longitude devem ser números finitos.")
    if not (-80.0 <= lat_deg <= 84.0):
        raise ValueError(f"Latitude {lat_deg}° fora dos limites válidos do sistema UTM (-80° a +84°).")
    if not (-180.0 <= lon_deg <= 180.0):
        raise ValueError(f"Longitude {lon_deg}° fora dos limites válidos (-180° a +180°).")

    a = 6378137.0
    f = 1.0 / 298.257223563
    e2 = 2 * f - f * f
    e_prime2 = e2 / (1.0 - e2)
    k0 = 0.9996

    zona_num = min(60, max(1, int((lon_deg + 180.0) / 6.0) + 1))
    letras = "CDEFGHJKLMNPQRSTUVWX"
    idx_banda = int((lat_deg + 80.0) / 8.0)
    idx_banda = max(0, min(len(letras) - 1, idx_banda))
    banda = letras[idx_banda]

    lat_rad = math.radians(lat_deg)
    lon_rad = math.radians(lon_deg)
    lon0 = math.radians((zona_num - 1) * 6 - 180 + 3)

    N = a / math.sqrt(1.0 - e2 * math.sin(lat_rad)**2)
    T = math.tan(lat_rad)**2
    C = e_prime2 * math.cos(lat_rad)**2
    A = math.cos(lat_rad) * (lon_rad - lon0)

    M = a * (
        (1.0 - e2 / 4.0 - 3.0 * e2**2 / 64.0 - 5.0 * e2**3 / 256.0) * lat_rad
        - (3.0 * e2 / 8.0 + 3.0 * e2**2 / 32.0 + 45.0 * e2**3 / 1024.0) * math.sin(2.0 * lat_rad)
        + (15.0 * e2**2 / 256.0 + 45.0 * e2**3 / 1024.0) * math.sin(4.0 * lat_rad)
        - (35.0 * e2**3 / 3072.0) * math.sin(6.0 * lat_rad)
    )

    easting = 500000.0 + k0 * N * (
        A + (1.0 - T + C) * (A**3) / 6.0
        + (5.0 - 18.0 * T + T**2 + 72.0 * C - 58.0 * e_prime2) * (A**5) / 120.0
    )
    northing = k0 * (
        M + N * math.tan(lat_rad) * (
            (A**2) / 2.0 + (5.0 - T + 9.0 * C + 4.0 * C**2) * (A**4) / 24.0
            + (61.0 - 58.0 * T + T**2 + 600.0 * C - 330.0 * e_prime2) * (A**6) / 720.0
        )
    )
    if lat_deg < 0:
        northing += 10000000.0

    return f"{zona_num}{banda}", round(easting, 2), round(northing, 2)


def parse_latlon_texto(texto):
    """Extrai (lat, lon) de texto decimal (-16,6868 -49,2648), link do Google Maps ou GMS."""
    if not texto:
        return None, None
    s = str(texto).strip()

    m_url = re.search(r"[@?](?:q=)?(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)", s)
    if m_url:
        lat, lon = float(m_url.group(1)), float(m_url.group(2))
        if -90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0:
            return lat, lon

    m_gms = re.findall(r"(\d{1,3})[°º]\s*(\d{1,2})['′]\s*([\d.,]+)\"?\s*([NSEWO])", s, re.IGNORECASE)
    if len(m_gms) >= 2:
        vals = []
        for g, m, seg, hem in m_gms[:2]:
            dec = float(g) + float(m) / 60.0 + converter_para_float(seg, False) / 3600.0
            if hem.upper() in ("S", "W", "O"):
                dec = -dec
            vals.append(dec)
        if -90.0 <= vals[0] <= 90.0 and -180.0 <= vals[1] <= 180.0:
            return vals[0], vals[1]

    s_norm = re.sub(r"(\d),(\d)", r"\1.\2", s)
    nums = re.findall(r"-?\d+\.\d+", s_norm)
    if len(nums) >= 2:
        lat, lon = float(nums[0]), float(nums[1])
        if -90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0:
            return lat, lon
    return None, None


# --- ESTATÍSTICA, HOMOGENEIZAÇÃO & REGRESSÃO LINEAR MÚLTIPLA (PRÉ-SISDEA) ---
def calcular_percentil(valores_ordenados, p):
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


def t_critico_80(gl):
    """Valor crítico bilateral t de Student para 80% de confiança (NBR 14653-2 IC 80%)."""
    tabela = {
        1: 3.078, 2: 1.886, 3: 1.638, 4: 1.533, 5: 1.476, 6: 1.440, 7: 1.415,
        8: 1.397, 9: 1.383, 10: 1.372, 11: 1.363, 12: 1.356, 13: 1.350, 14: 1.345,
        15: 1.341, 16: 1.337, 17: 1.333, 18: 1.330, 19: 1.328, 20: 1.325,
        25: 1.316, 30: 1.310, 40: 1.303, 60: 1.296, 120: 1.289
    }
    if gl <= 1:
        return 3.078
    if gl in tabela:
        return tabela[gl]
    chaves = sorted(tabela.keys())
    for i in range(len(chaves) - 1):
        if chaves[i] < gl < chaves[i + 1]:
            return tabela[chaves[i]]
    return 1.282


def p_valor_t_aprox(t_abs, gl):
    """Aproximação numérica precisa do p-valor bilateral da distribuição t-Student."""
    if gl <= 0 or t_abs <= 0:
        return 1.0
    steps = 400
    h = t_abs / steps
    expoente = -(gl + 1.0) / 2.0
    c = math.gamma((gl + 1.0) / 2.0) / (math.sqrt(gl * math.pi) * math.gamma(gl / 2.0))

    def pdf(x):
        return c * ((1.0 + (x * x) / gl) ** expoente)

    soma = pdf(0.0) + pdf(t_abs)
    for i in range(1, steps):
        soma += (4.0 if i % 2 == 1 else 2.0) * pdf(i * h)
    area_unilateral = min(0.5, (h / 3.0) * soma)
    return max(0.0, min(1.0, 1.0 - 2.0 * area_unilateral))


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
    t80 = t_critico_80(n - 1) if n > 1 else 0.0
    semi_ic80 = (t80 * dp / math.sqrt(n)) if n > 1 else 0.0
    ic80_inf = media - semi_ic80
    ic80_sup = media + semi_ic80
    amp_ic80_pct = ((ic80_sup - ic80_inf) / media * 100.0) if media > 0 and n > 1 else 0.0
    return {
        "n": n,
        "minimo": ord_v[0],
        "maximo": ord_v[-1],
        "amplitude": ord_v[-1] - ord_v[0],
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
        "ic80_inf": ic80_inf,
        "ic80_sup": ic80_sup,
        "amp_ic80_pct": amp_ic80_pct
    }


def extrair_valor_numerico_variavel(dado, chave_coluna):
    """Extrai valor numérico de campo fixo ou variável extra (incluindo código líder '1 - Local' -> 1.0)."""
    if chave_coluna == "area_terreno":
        return float(dado.get("area_terreno", 0.0))
    if chave_coluna == "area_construida":
        return float(dado.get("area_construida", 0.0))
    if chave_coluna == "unitario":
        return float(dado.get("unitario", 0.0))
    if chave_coluna.startswith("var_"):
        nome_v = chave_coluna[4:]
        bruto = sanitizar_texto_celula(dado.get("variaveis_extras", {}).get(nome_v, ""))
        if not bruto:
            return None
        num, ok = converter_para_float_validado(bruto, permitir_milhar_ponto=True)
        if ok:
            return num
        m_cod = re.match(r"^\s*(-?\d+(?:[.,]\d+)?)\s*[-–:]", bruto)
        if m_cod:
            c_num, c_ok = converter_para_float_validado(m_cod.group(1), permitir_milhar_ponto=False)
            if c_ok:
                return c_num
    return None


def resolver_sistema_linear(A, b):
    """Resolve A * x = b e retorna (x, matriz_inversa_A) via Gauss-Jordan com pivoteamento parcial."""
    n = len(A)
    aug = [list(A[i]) + [1.0 if i == j else 0.0 for j in range(n)] + [b[i]] for i in range(n)]
    for col in range(n):
        piv = max(range(col, n), key=lambda r: abs(aug[r][col]))
        if abs(aug[piv][col]) < 1e-12:
            return None, None
        aug[col], aug[piv] = aug[piv], aug[col]
        div = aug[col][col]
        for j in range(col, 2 * n + 1):
            aug[col][j] /= div
        for r in range(n):
            if r != col and aug[r][col] != 0.0:
                fator = aug[r][col]
                for j in range(col, 2 * n + 1):
                    aug[r][j] -= fator * aug[col][j]
    sol = [aug[i][2 * n] for i in range(n)]
    inv = [aug[i][n:2 * n] for i in range(n)]
    return sol, inv


def calcular_vifs_modelo(linhas_x):
    """Calcula o Variance Inflation Factor (VIF) para cada regressor j = 1..k (multicolinearidade)."""
    n = len(linhas_x)
    p = len(linhas_x[0])
    k = p - 1
    if k <= 1:
        return [None] + [1.0] * k
    vifs = [None]
    for j in range(1, p):
        y_j = [linhas_x[i][j] for i in range(n)]
        y_bar = statistics.mean(y_j)
        sqt_j = sum((y_j[i] - y_bar) ** 2 for i in range(n))
        if sqt_j <= 1e-15:
            vifs.append(float("inf"))
            continue
        cols_outros = [0] + [c for c in range(1, p) if c != j]
        m_sub = len(cols_outros)
        XtX = [[sum(linhas_x[i][r] * linhas_x[i][c] for i in range(n)) for c in cols_outros] for r in cols_outros]
        XtY = [sum(linhas_x[i][r] * y_j[i] for i in range(n)) for r in cols_outros]
        b_sub, _ = resolver_sistema_linear(XtX, XtY)
        if b_sub is None:
            vifs.append(float("inf"))
            continue
        y_hat = [sum(b_sub[idx_c] * linhas_x[i][cols_outros[idx_c]] for idx_c in range(m_sub)) for i in range(n)]
        sqr_j = sum((y_j[i] - y_hat[i]) ** 2 for i in range(n))
        r2_j = max(0.0, min(1.0, 1.0 - (sqr_j / sqt_j)))
        if r2_j >= 0.999999:
            vifs.append(float("inf"))
        else:
            vifs.append(1.0 / (1.0 - r2_j))
    return vifs


def executar_regressao_linear_multipla(amostras, nomes_x, chaves_x, transf_y="LIN", transf_x_log=None):
    """
    Executa Regressão Linear Múltipla (MQO / OLS) com diagnóstico de VIF, correção de viés log-normal
    e proteção de domínio para transformação inversa.
    """
    transf_x_log = transf_x_log or set()
    linhas_x = []
    valores_y_transf = []
    valores_y_orig = []
    amostras_validas = []

    for d in amostras:
        q = float(d.get("unitario", 0.0))
        if q <= 0:
            continue
        if transf_y == "LN":
            y_val = math.log(q)
        elif transf_y == "INV":
            y_val = 1.0 / q
        else:
            y_val = q

        xs = []
        valido = True
        for ch in chaves_x:
            xv = extrair_valor_numerico_variavel(d, ch)
            if xv is None:
                valido = False
                break
            if ch in transf_x_log:
                if xv <= 0:
                    valido = False
                    break
                xv = math.log(xv)
            xs.append(xv)

        if valido:
            linhas_x.append([1.0] + xs)
            valores_y_transf.append(y_val)
            valores_y_orig.append(q)
            amostras_validas.append(d)

    n = len(linhas_x)
    k = len(chaves_x)
    p = k + 1
    if n <= p:
        return None, f"Amostras completas insuficientes ({n}) para {k} variável(is) explicativa(s). São necessárias pelo menos {p + 1} amostras válidas."

    XtX = [[sum(linhas_x[i][r] * linhas_x[i][c] for i in range(n)) for c in range(p)] for r in range(p)]
    XtY = [sum(linhas_x[i][r] * valores_y_transf[i] for i in range(n)) for r in range(p)]

    betas, inv_XtX = resolver_sistema_linear(XtX, XtY)
    if betas is None:
        return None, "Matriz singular (colinearidade perfeita ou variável constante). Remova variáveis redundantes."

    y_bar = statistics.mean(valores_y_transf)
    y_hat_transf = [sum(betas[j] * linhas_x[i][j] for j in range(p)) for i in range(n)]
    residuos = [valores_y_transf[i] - y_hat_transf[i] for i in range(n)]

    sqt = sum((valores_y_transf[i] - y_bar) ** 2 for i in range(n))
    sqr = sum(residuos[i] ** 2 for i in range(n))
    sqe = sqt - sqr

    gl_reg = k
    gl_res = n - p
    r2 = (sqe / sqt) if sqt > 0 else 0.0
    r2_adj = 1.0 - ((1.0 - r2) * (n - 1) / gl_res) if gl_res > 0 else 0.0
    r_corr = math.sqrt(max(0.0, r2))
    s2 = sqr / gl_res if gl_res > 0 else 0.0
    s_erro = math.sqrt(max(0.0, s2))
    f_stat = ((sqe / gl_reg) / s2) if (gl_reg > 0 and s2 > 0) else 0.0

    # Fator de correção de viés de retransformação log-normal: exp(s^2 / 2)
    fator_vies_ln = math.exp(s2 / 2.0) if (transf_y == "LN" and s2 < 20.0) else 1.0
    vifs = calcular_vifs_modelo(linhas_x)

    regressores = []
    labels_modelo = ["Intercepto (β0)"] + [
        f"ln({nomes_x[idx]})" if chaves_x[idx] in transf_x_log else nomes_x[idx]
        for idx in range(k)
    ]
    for j in range(p):
        var_bj = s2 * inv_XtX[j][j]
        se_j = math.sqrt(max(0.0, var_bj))
        t_j = (betas[j] / se_j) if se_j > 0 else 0.0
        pval_j = p_valor_t_aprox(abs(t_j), gl_res)
        if j == 0:
            r_simples = 0.0
        else:
            col_x = [linhas_x[i][j] for i in range(n)]
            x_bar = statistics.mean(col_x)
            num_r = sum((col_x[i] - x_bar) * (valores_y_transf[i] - y_bar) for i in range(n))
            den_r = math.sqrt(sum((col_x[i] - x_bar) ** 2 for i in range(n)) * sqt)
            r_simples = (num_r / den_r) if den_r > 0 else 0.0

        regressores.append({
            "nome": labels_modelo[j],
            "beta": betas[j],
            "se": se_j,
            "t": t_j,
            "p_valor": pval_j,
            "r_simples": r_simples,
            "vif": vifs[j]
        })

    detalhes_amostras = []
    for i, d in enumerate(amostras_validas):
        yh = y_hat_transf[i]
        fora_dominio = False
        if transf_y == "LN":
            q_est = math.exp(yh) * fator_vies_ln
        elif transf_y == "INV":
            if yh <= 1e-12:
                q_est = None
                fora_dominio = True
            else:
                q_est = 1.0 / yh
        else:
            q_est = yh

        q_obs = valores_y_orig[i]
        erro_pct = ((q_obs - q_est) / q_est * 100.0) if (q_est is not None and q_est > 0) else None
        z_res = (residuos[i] / s_erro) if s_erro > 0 else 0.0
        detalhes_amostras.append({
            "dado_id": d["dado_id"],
            "municipio": d.get("municipio", ""),
            "q_obs": q_obs,
            "q_est": q_est,
            "fora_dominio": fora_dominio,
            "erro_pct": erro_pct,
            "z_res": z_res,
            "outlier": abs(z_res) > 2.0
        })

    max_vif = max([v for v in vifs[1:] if v is not None] or [1.0])
    return {
        "n": n,
        "k": k,
        "gl_res": gl_res,
        "transf_y": transf_y,
        "betas": betas,
        "r": r_corr,
        "r2": r2,
        "r2_adj": r2_adj,
        "s_erro": s_erro,
        "f_stat": f_stat,
        "fator_vies_ln": fator_vies_ln,
        "max_vif": max_vif,
        "regressores": regressores,
        "detalhes_amostras": detalhes_amostras
    }, ""


def redimensionar_e_corrigir_imagem(caminho_ou_stream, max_largura=1200, max_altura=850):
    try:
        if isinstance(caminho_ou_stream, io.BytesIO):
            caminho_ou_stream.seek(0)
            with Image.open(caminho_ou_stream) as im_check:
                im_check.verify()
            caminho_ou_stream.seek(0)
            img = Image.open(caminho_ou_stream)
        else:
            if not os.path.exists(caminho_ou_stream):
                return None
            with Image.open(caminho_ou_stream) as im_check:
                im_check.verify()
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
        aplicar_nota_agronomica_no_dado(reg, variaveis_config, origem_edicao="auto")
    return reg


class AppPesquisaMercado:
    def __init__(self, root):
        self.root = root
        self.root.title("ENPROL - Sistema de Pesquisa de Mercado v7.1 (Planilha SISDEA & Geoprocessamento)")
        self.root.geometry("1380x860")
        self.root.minsize(1080, 700)
        self._configurar_estilos()

        self.modelo_ativo = "PADRAO"
        self.finalidade_ativa = "VENDA"
        self.unidade_padrao_projeto = "m²"
        self.destacar_outliers = True
        self.criterio_outlier = "30%"
        self.dados_pesquisas = []
        self.catalogo_variaveis = self._carregar_catalogo_variaveis()
        self.variaveis_config = self._carregar_config_variaveis()
        self.campos_opcionais = copy.deepcopy(DEFAULT_CAMPOS_OPCIONAIS)
        self.config_word = copy.deepcopy(DEFAULT_CONFIG_WORD)
        self.perfis_usuario = self._carregar_perfis_usuario()
        self.caminho_projeto_atual = None
        self.alteracoes_pendentes = False
        self.exportando_word = False

        self.celula_ativa = {"row": None, "col": None, "chave": None}
        self.painel_lateral_visivel = False
        self._img_thumb1 = None
        self._img_thumb2 = None
        self._img_mapa_preview = None
        self.fila_eventos_ui = queue.Queue()

        self.container_principal = ttk.Frame(self.root)
        self.container_principal.pack(fill="both", expand=True)

        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar_janela)
        self.root.bind("<Control-s>", lambda e: self._salvar_projeto())
        self.root.bind("<Control-S>", lambda e: self._salvar_projeto())
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
                    self._concluir_exportacao(payload)
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

    def _autosave(self, forcar_limpeza=False):
        if not self.dados_pesquisas:
            if forcar_limpeza and os.path.exists(AUTOSAVE_FILE):
                try:
                    os.remove(AUTOSAVE_FILE)
                except Exception as e:
                    logging.error(f"Falha ao limpar autosave: {e}")
            return
        try:
            payload = {
                "caminho_projeto_origem": self.caminho_projeto_atual,
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

    def _verificar_descarte_seguro(self, acao_nome="continuar"):
        """Retorna True se puder prosseguir; pergunta Salvar/Descartar/Cancelar se houver alterações pendentes."""
        self._commit_editor_ativo()
        if self.dados_pesquisas and self.alteracoes_pendentes:
            resp = messagebox.askyesnocancel(
                "Alterações Não Salvas",
                f"Existem alterações não salvas no projeto atual. Deseja salvar antes de {acao_nome}?"
            )
            if resp is None:
                return False
            if resp is True:
                return self._salvar_projeto()
        return True

    def _ao_fechar_janela(self):
        if not self._verificar_descarte_seguro("sair do sistema"):
            return
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

        ttk.Label(frame_hub, text="SISTEMA DE PESQUISA DE MERCADO v7.1", font=("Segoe UI", 18, "bold"), foreground="#1f3c5b").pack(pady=(0, 5))
        ttk.Label(frame_hub, text="Planilha SISDEA • Mapa Interativo • Mapa JPG Multi-Camadas • KML/KMZ • Regressão", font=("Segoe UI", 10), foreground="#555555").pack(pady=(0, 30))

        ttk.Button(frame_hub, text="📄  Criar Novo Arquivo / Projeto", style="Hub.TButton", width=48, command=self._popup_escolha_modelo).pack(pady=7)
        ttk.Button(frame_hub, text="📂  Carregar Projeto Existente (.json)", style="Hub.TButton", width=48, command=self._abrir_projeto_hub).pack(pady=7)
        ttk.Button(frame_hub, text="📥  Importar Dados de Planilha (.xlsx)", style="Hub.TButton", width=48, command=self._importar_planilha_hub).pack(pady=7)
        if os.path.exists(AUTOSAVE_FILE):
            ttk.Button(frame_hub, text="🔄  Recuperar Sessão Anterior (Autosave)", style="Accent.TButton", width=48, command=self._carregar_autosave).pack(pady=7)
        ttk.Button(frame_hub, text="🚪  Sair do Sistema", width=48, command=self._ao_fechar_janela).pack(pady=(20, 0))

    def _carregar_autosave(self):
        try:
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f:
                conteudo = json.load(f)
            self.caminho_projeto_atual = conteudo.get("caminho_projeto_origem")
            self._aplicar_conteudo_json(conteudo)
            self._exibir_tela_trabalho()
            messagebox.showinfo("Recuperado", f"{len(self.dados_pesquisas)} pesquisas recuperadas do autosave!")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao recuperar autosave: {e}")

    def _popup_escolha_modelo(self, callback_pos_confirmacao=None):
        if not self._verificar_descarte_seguro("criar um novo projeto"):
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
            self._autosave(forcar_limpeza=True)
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
        m_arq.add_command(label="Salvar Projeto (.json)   [Ctrl+S]", command=self._salvar_projeto)
        m_arq.add_separator()
        m_arq.add_command(label="📥 Importar Planilha Excel...", command=self._importar_planilha_excel)
        m_arq.add_separator()
        m_arq.add_command(label="Sair", command=self._ao_fechar_janela)
        menubar.add_cascade(label="Arquivo", menu=m_arq)

        m_geo = tk.Menu(menubar, tearoff=0)
        m_geo.add_command(label="🌐 Abrir Mapa Interativo (Navegador + Fotos)...", command=self._abrir_mapa_interativo_html)
        m_geo.add_command(label="🌍 Gerar KML / KMZ (Google Earth com Fotos)...", command=self._exportar_kmz_kml)
        m_geo.add_command(label="🗺 Gerar Mapa JPG / PNG das Amostras...", command=self._janela_gerar_mapa_imagem)
        m_geo.add_separator()
        m_geo.add_command(label="🧭 Conversor Lat/Long ↔ UTM...", command=self._janela_conversor_utm)
        menubar.add_cascade(label="Geoprocessamento & Mapas", menu=m_geo)

        m_cfg = tk.Menu(menubar, tearoff=0)
        m_cfg.add_command(label="⚙ Gerenciar Variáveis e Catálogo...", command=self._janela_config_variaveis)
        m_cfg.add_command(label="👁 Campos, Modelo, Word e Alqueire...", command=self._janela_campos_opcionais)
        m_cfg.add_command(label="💾 Salvar Configuração Atual como Novo Perfil...", command=self._salvar_como_novo_perfil)
        m_cfg.add_separator()
        m_cfg.add_command(label="Aplicar Perfil Urbano Padrão", command=lambda: self._aplicar_perfil_rapido("URBANO"))
        m_cfg.add_command(label="Aplicar Perfil Rural (Nota Agronômica)", command=lambda: self._aplicar_perfil_rapido("RURAL"))
        m_cfg.add_command(label="Aplicar Perfil COPASA", command=lambda: self._aplicar_perfil_rapido("COPASA"))
        m_cfg.add_separator()
        m_cfg.add_command(label="📈 Pré-Análise SISDEA (Estatística, Homogeneização & Regressão)...", command=self._janela_analise_estatistica)
        m_cfg.add_command(label="📋 Ver Tabela 02 - Nota Agronômica", command=self._mostrar_tabela_nota_agronomica)
        menubar.add_cascade(label="Variáveis, Perfis & SISDEA", menu=m_cfg)
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
        ttk.Button(toolbar, text="🔄 Unid.", command=self._padronizar_unidades_clique).pack(side="left", padx=2)

        ttk.Button(toolbar, text="📄 Word", style="Primary.TButton", command=self._iniciar_exportacao_word_thread).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📊 Excel", style="Accent.TButton", command=self._exportar_excel).pack(side="right", padx=2)
        ttk.Button(toolbar, text="🗺 Mapa JPG", command=self._janela_gerar_mapa_imagem).pack(side="right", padx=2)
        ttk.Button(toolbar, text="🌐 Mapa Interativo", command=self._abrir_mapa_interativo_html).pack(side="right", padx=2)
        ttk.Button(toolbar, text="🌍 KML/KMZ", command=self._exportar_kmz_kml).pack(side="right", padx=2)
        ttk.Button(toolbar, text="📷 Mídia", command=self._alternar_painel_lateral).pack(side="right", padx=3)

        self.lbl_info_mod = ttk.Label(toolbar, text="", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b")
        self.lbl_info_mod.pack(side="left", padx=4)
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
        try:
            self.sheet.set_options(expand_sheet_if_paste_too_big=True)
        except Exception:
            pass
        self.sheet.enable_bindings(
            "single_select", "row_select", "column_select", "drag_select",
            "column_width_resize", "double_click_column_resize", "row_height_resize",
            "arrowkeys", "right_click_popup_menu", "rc_select",
            "copy", "cut", "paste", "delete", "undo", "redo", "edit_cell"
        )
        self.sheet.pack(fill="both", expand=True)

        self.sheet.popup_menu_add_command("➕ Adicionar 1 Dado", lambda: self._adicionar_dados_qtd(1))
        self.sheet.popup_menu_add_command("➕ Adicionar Vários Dados...", self._perguntar_qtd_adicionar_dados)
        self.sheet.popup_menu_add_command("📋 Extrair de Texto de Anúncio...", self._janela_extrator_anuncio)
        self.sheet.popup_menu_add_command("🧭 Converter Lat/Long ➔ UTM...", self._janela_conversor_utm)
        self.sheet.popup_menu_add_command("📋 Duplicar Linha Selecionada", self._duplicar_linha_selecionada)
        self.sheet.popup_menu_add_command("☑ Alternar Utilizado / Não Utilizado", self._alternar_utilizado_selecionado)
        self.sheet.popup_menu_add_command("⬇ Preencher Para Baixo (Ctrl+D)", self._preencher_para_baixo)
        self.sheet.popup_menu_add_command("🗑 Excluir Linha(s) Selecionada(s)", self._excluir_linhas_selecionadas)

        self.sheet.extra_bindings("cell_select", self._ao_selecionar_celula)
        self.sheet.extra_bindings("end_edit_cell", self._ao_terminar_edicao_celula)
        self.sheet.extra_bindings("end_paste", self._ao_colar_planilha)
        self.sheet.extra_bindings("end_delete", self._ao_deletar_conteudo)
        for ev_extra in ("end_undo", "end_redo", "end_cut"):
            try:
                self.sheet.extra_bindings(ev_extra, self._ao_modificar_grade_em_lote)
            except Exception:
                pass
        self.sheet.bind("<Return>", self._ao_pressionar_enter_grade)
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

        ttk.Button(frame_rodape_btns, text="📈 Pré-Análise SISDEA", style="Primary.TButton", command=self._janela_analise_estatistica).pack(side="left", padx=2)
        ttk.Button(frame_rodape_btns, text="🧹 Saneamento", command=self._desativar_outliers_criterio).pack(side="left", padx=2)
        ttk.Button(frame_rodape_btns, text="📋 Copiar p/ SISDEA", command=self._copiar_tabela_sisdea).pack(side="left", padx=2)

    def _ao_pressionar_enter_grade(self, event=None):
        if not hasattr(self, "sheet") or not self.dados_pesquisas:
            return
        sel = self.sheet.get_currently_selected()
        if sel and sel.row is not None and sel.row == len(self.dados_pesquisas) - 1:
            self._adicionar_dados_qtd(1)
            self.sheet.select_cell(len(self.dados_pesquisas) - 1, sel.column or 2)

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
        self._sincronizar_matriz_para_dados(reconstruir=False, avisar_invalidos=False)

    def _sincronizar_matriz_para_dados(self, reconstruir=False, avisar_invalidos=True):
        if not hasattr(self, "sheet") or not hasattr(self, "esquema_colunas"):
            return
        matriz = self.sheet.get_sheet_data()
        while len(self.dados_pesquisas) < len(matriz):
            self._criar_registro_vazio(herdar_anterior=True)

        celulas_invalidas = []

        for r_idx, linha in enumerate(matriz):
            dado = self.dados_pesquisas[r_idx]
            dado["dado_id"] = r_idx + 1
            un_anterior = normalizar_unidade(dado.get("unidade", self.unidade_padrao_projeto))
            un_nova = un_anterior

            col_un_idx = next((i for i, c in enumerate(self.esquema_colunas) if c[0] == "unidade"), None)
            unidade_foi_alterada = False
            if col_un_idx is not None and col_un_idx < len(linha):
                un_cel = normalizar_unidade(linha[col_un_idx], padrao=un_anterior)
                if un_cel != un_anterior:
                    un_nova = un_cel
                    unidade_foi_alterada = True

            area_foi_alterada_na_matriz = False
            for c_idx, (chave, titulo_col, _) in enumerate(self.esquema_colunas):
                if c_idx >= len(linha) or chave in ("dado_id", "unitario", "status", "unidade"):
                    continue
                val = linha[c_idx]
                if chave == "utilizado":
                    dado["utilizado"] = bool(val) if isinstance(val, bool) else str(val).strip().lower() not in ("false", "0", "não", "nao", "")
                elif chave == "valor_total":
                    v_num, ok = converter_para_float_validado(val)
                    if ok:
                        dado["valor_total"] = v_num
                    else:
                        celulas_invalidas.append(f"Linha {r_idx + 1} ({titulo_col}): '{val}'")
                elif chave == "area_terreno":
                    a_num, ok = converter_para_float_validado(val, permitir_milhar_ponto=(un_nova != "ha"))
                    if ok:
                        if abs(a_num - dado.get("area_terreno", 0.0)) > 1e-9:
                            area_foi_alterada_na_matriz = True
                        dado["area_terreno"] = a_num
                    else:
                        celulas_invalidas.append(f"Linha {r_idx + 1} ({titulo_col}): '{val}'")
                elif chave == "area_construida":
                    ac_num, ok = converter_para_float_validado(val)
                    if ok:
                        dado["area_construida"] = ac_num
                    else:
                        celulas_invalidas.append(f"Linha {r_idx + 1} ({titulo_col}): '{val}'")
                elif chave == "telefone":
                    dado["telefone"] = sanitizar_telefone(val)
                elif chave in ("coord_e", "coord_s"):
                    dado[chave] = limpar_sufixo_coord(sanitizar_texto_celula(val))
                elif chave.startswith("var_"):
                    dado.setdefault("variaveis_extras", {})[chave[4:]] = sanitizar_texto_celula(val)
                else:
                    dado[chave] = sanitizar_texto_celula(val)

            if unidade_foi_alterada:
                if not area_foi_alterada_na_matriz and dado.get("area_terreno", 0.0) > 0:
                    dado["area_terreno"] = converter_area(dado["area_terreno"], un_anterior, un_nova)
                dado["unidade"] = un_nova
                dado["localizacao"] = "Rural" if un_nova == "ha" else "Urbana"
            else:
                dado["unidade"] = un_anterior

            a_ter = dado.get("area_terreno", 0.0)
            dado["unitario"] = dado.get("valor_total", 0.0) / a_ter if a_ter > 0 else 0.0
            aplicar_nota_agronomica_no_dado(dado, self.variaveis_config, origem_edicao="auto")

        if self.dados_pesquisas:
            uns = {d.get("unidade", self.unidade_padrao_projeto) for d in self.dados_pesquisas}
            if len(uns) == 1:
                self.unidade_padrao_projeto = next(iter(uns))
                self._atualizar_rotulo_topo()

        if reconstruir:
            self._reconstruir_planilha()

        if avisar_invalidos and celulas_invalidas:
            messagebox.showwarning(
                "Valores Inválidos Ignorados",
                "As seguintes células continham texto não numérico e mantiveram o valor anterior:\n\n• "
                + "\n• ".join(celulas_invalidas[:10])
            )

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
            cols.append((f"var_{v_cfg['nome']}", v_cfg["nome"], 120))
        cols.append(("unitario", "Unitário (R$/un)", 115))
        if self.campos_opcionais.get("localizacao", True):
            cols.append(("localizacao", "Localização", 90))
        cols.extend([
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
                linha.append(formatar_numero_br(ac, 2))
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
                for r_idx, d in enumerate(self.dados_pesquisas):
                    self.sheet.set_cell_data(r_idx, c_idx, d.get("unidade", self.unidade_padrao_projeto), redraw=False)
            elif chave == "localizacao" and matriz:
                self._aplicar_dropdown_aberto(self.sheet.span(0, c_idx, len(matriz), c_idx + 1), ["", "Urbana", "Rural"])
                for r_idx, d in enumerate(self.dados_pesquisas):
                    self.sheet.set_cell_data(r_idx, c_idx, d.get("localizacao", "Urbana"), redraw=False)
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
            self.lbl_estatisticas.config(text=f"⚠ Mistura de unidades ({', '.join(unidades)})! Use 'Unid.' para converter e unificar.", fg="#b22222")
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
# --- PRÉ-ANÁLISE SISDEA (ESTATÍSTICA + HOMOGENEIZAÇÃO + REGRESSÃO LINEAR MÚLTIPLA) ---
    def _janela_analise_estatistica(self):
        self._commit_editor_ativo()
        ativos = [d for d in self.dados_pesquisas if d.get("utilizado", True) and d.get("unitario", 0.0) > 0]
        if len(ativos) < 2:
            messagebox.showinfo("Pré-Análise SISDEA", "É necessário pelo menos 2 amostras ativas com unitário válido.")
            return
        unidades = {d.get("unidade", self.unidade_padrao_projeto) for d in ativos}
        if len(unidades) > 1:
            messagebox.showwarning("Unidades Mistas", "Padronize as unidades antes de abrir a Pré-Análise SISDEA.")
            return

        un_str = next(iter(unidades))
        unitarios = [d["unitario"] for d in ativos]
        est = calcular_estatisticas_completas(unitarios)

        out_30 = [f"D.{d['dado_id']} (R$ {formatar_moeda_br(d['unitario'])})" for d in ativos if self._eh_outlier(d["unitario"], est, "30%")]
        out_iqr = [f"D.{d['dado_id']} (R$ {formatar_moeda_br(d['unitario'])})" for d in ativos if self._eh_outlier(d["unitario"], est, "IQR")]

        janela = tk.Toplevel(self.root)
        janela.title(f"Pré-Análise SISDEA — NBR 14653 ({est['n']} amostras ativas • R$/{un_str})")
        janela.geometry("1020x690")
        janela.transient(self.root)

        nb = ttk.Notebook(janela)
        nb.pack(fill="both", expand=True, padx=10, pady=8)

        # === ABA 1: ESTATÍSTICA DESCRITIVA & BOXPLOT ===
        tab_est = ttk.Frame(nb, padding=8)
        nb.add(tab_est, text=" 1. Estatística & Boxplot ")

        f_top = ttk.LabelFrame(tab_est, text=f" Resumo Estatístico da Amostra Bruta (R$/{un_str}) ", padding=10)
        f_top.pack(fill="x", pady=4)

        diag_cv = (
            "Baixa dispersão (≤ 15%)" if est["cv"] <= 15
            else ("Dispersão moderada (15%–30%)" if est["cv"] <= 30
                  else "Alta dispersão (> 30% — recomenda-se homogeneizar ou regredir)")
        )
        metricas = [
            ("Amostras Ativas:", f"{est['n']}", "Mínimo / Máximo:", f"R$ {formatar_moeda_br(est['minimo'])}  /  R$ {formatar_moeda_br(est['maximo'])}"),
            ("Média:", f"R$ {formatar_moeda_br(est['media'])}", "Desvio Padrão:", f"R$ {formatar_moeda_br(est['dp'])}"),
            ("Mediana (Q2):", f"R$ {formatar_moeda_br(est['q2'])}", "Coef. Variação (CV):", f"{formatar_numero_br(est['cv'], 2)}% ({diag_cv})"),
            ("Quartis (Q1 / Q3):", f"R$ {formatar_moeda_br(est['q1'])}  /  R$ {formatar_moeda_br(est['q3'])}", "IQR (Q3 - Q1):", f"R$ {formatar_moeda_br(est['iqr'])}"),
            ("IC 80% Média (t-Student):", f"R$ {formatar_moeda_br(est['ic80_inf'])} a R$ {formatar_moeda_br(est['ic80_sup'])}", "Semi-amplitude IC 80%:", f"± {formatar_numero_br(est['amp_ic80_pct'] / 2.0, 2)}% (Total: {formatar_numero_br(est['amp_ic80_pct'], 2)}%)"),
        ]
        for idx, (k1, v1, k2, v2) in enumerate(metricas):
            ttk.Label(f_top, text=k1, font=("Segoe UI", 9, "bold")).grid(row=idx, column=0, sticky="w", padx=6, pady=2)
            ttk.Label(f_top, text=v1).grid(row=idx, column=1, sticky="w", padx=(2, 24), pady=2)
            ttk.Label(f_top, text=k2, font=("Segoe UI", 9, "bold")).grid(row=idx, column=2, sticky="w", padx=6, pady=2)
            ttk.Label(f_top, text=v2).grid(row=idx, column=3, sticky="w", padx=2, pady=2)

        f_box = ttk.LabelFrame(tab_est, text=" Boxplot dos Valores Unitários ", padding=8)
        f_box.pack(fill="x", pady=6)
        canvas = tk.Canvas(f_box, height=135, bg="white", highlightthickness=1, highlightbackground="#cbd5e1")
        canvas.pack(fill="x", expand=True)

        w_c, pad_x, y_mid = 940, 55, 65
        v_min_plot = min(est["minimo"], est["lim_inf_iqr"])
        v_max_plot = max(est["maximo"], est["lim_sup_iqr"])
        span = (v_max_plot - v_min_plot) or 1.0

        def map_x(v):
            return pad_x + ((v - v_min_plot) / span) * (w_c - 2 * pad_x)

        x_q1, x_q2, x_q3 = map_x(est["q1"]), map_x(est["q2"]), map_x(est["q3"])
        inliers = [v for v in unitarios if est["lim_inf_iqr"] <= v <= est["lim_sup_iqr"]]
        w_low = min(inliers) if inliers else est["minimo"]
        w_high = max(inliers) if inliers else est["maximo"]
        x_wl, x_wh = map_x(w_low), map_x(w_high)

        canvas.create_line(x_wl, y_mid, x_q1, y_mid, fill="#1f3c5b", width=2)
        canvas.create_line(x_q3, y_mid, x_wh, y_mid, fill="#1f3c5b", width=2)
        canvas.create_line(x_wl, y_mid - 14, x_wl, y_mid + 14, fill="#1f3c5b", width=2)
        canvas.create_line(x_wh, y_mid - 14, x_wh, y_mid + 14, fill="#1f3c5b", width=2)
        canvas.create_rectangle(x_q1, y_mid - 22, x_q3, y_mid + 22, fill="#dbeafe", outline="#1f3c5b", width=2)
        canvas.create_line(x_q2, y_mid - 22, x_q2, y_mid + 22, fill="#b91c1c", width=3)
        x_med = map_x(est["media"])
        canvas.create_polygon(x_med, y_mid - 7, x_med + 7, y_mid, x_med, y_mid + 7, x_med - 7, y_mid, fill="#15803d", outline="#14532d")

        for d in ativos:
            v = d["unitario"]
            if v < est["lim_inf_iqr"] or v > est["lim_sup_iqr"]:
                xp = map_x(v)
                canvas.create_oval(xp - 4, y_mid - 4, xp + 4, y_mid + 4, fill="#f59e0b", outline="#92400e")
                canvas.create_text(xp, y_mid - 14, text=f"D.{d['dado_id']}", font=("Segoe UI", 7, "bold"), fill="#92400e")

        canvas.create_text(x_wl, 112, text=f"Min: {formatar_moeda_br(w_low)}", font=("Segoe UI", 8), fill="#334155")
        canvas.create_text(x_q2, 22, text=f"Mediana: {formatar_moeda_br(est['q2'])}", font=("Segoe UI", 8, "bold"), fill="#b91c1c")
        canvas.create_text(x_wh, 112, text=f"Max: {formatar_moeda_br(w_high)}", font=("Segoe UI", 8), fill="#334155")

        f_comp = ttk.LabelFrame(tab_est, text=" Diagnóstico de Pontos Atípicos (Outliers) ", padding=10)
        f_comp.pack(fill="both", expand=True, pady=4)
        ttk.Label(f_comp, text=f"• Critério ±30% da Média [R$ {formatar_moeda_br(est['lim_inf_30'])} a R$ {formatar_moeda_br(est['lim_sup_30'])}]:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Label(f_comp, text="   " + (", ".join(out_30) if out_30 else "Nenhuma amostra fora da faixa."), wraplength=940).pack(anchor="w", pady=(0, 8))
        ttk.Label(f_comp, text=f"• Critério IQR / Tukey (1,5 × IQR) [R$ {formatar_moeda_br(est['lim_inf_iqr'])} a R$ {formatar_moeda_br(est['lim_sup_iqr'])}]:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        ttk.Label(f_comp, text="   " + (", ".join(out_iqr) if out_iqr else "Nenhuma amostra fora da faixa."), wraplength=940).pack(anchor="w")

        # === ABA 2: HOMOGENEIZAÇÃO POR FATORES (NBR 14653) ===
        tab_hom = ttk.Frame(nb, padding=8)
        nb.add(tab_hom, text=" 2. Homogeneização (NBR 14653) ")

        f_par_hom = ttk.LabelFrame(tab_hom, text=" Parâmetros dos Fatores de Homogeneização (Validação Estrita) ", padding=8)
        f_par_hom.pack(fill="x", pady=(0, 6))

        areas_ativas = [d["area_terreno"] for d in ativos if d["area_terreno"] > 0]
        area_med_padrao = statistics.median(areas_ativas) if areas_ativas else 1.0

        ttk.Label(f_par_hom, text="Fator Oferta (p/ Condição = Oferta):").grid(row=0, column=0, sticky="w", padx=4, pady=2)
        e_f_oferta = ttk.Entry(f_par_hom, width=8)
        e_f_oferta.insert(0, "0,90")
        e_f_oferta.grid(row=0, column=1, padx=(0, 14), pady=2)

        ttk.Label(f_par_hom, text=f"Área Paradigma ({un_str}):").grid(row=0, column=2, sticky="w", padx=4, pady=2)
        e_area_par = ttk.Entry(f_par_hom, width=12)
        e_area_par.insert(0, formatar_numero_br(area_med_padrao, 4 if un_str == "ha" else 2))
        e_area_par.grid(row=0, column=3, padx=(0, 14), pady=2)

        ttk.Label(f_par_hom, text="Expoente Área (p):").grid(row=0, column=4, sticky="w", padx=4, pady=2)
        e_exp_area = ttk.Entry(f_par_hom, width=8)
        e_exp_area.insert(0, "0,00")
        e_exp_area.grid(row=0, column=5, padx=(0, 14), pady=2)

        var_usar_nota_hom = tk.BooleanVar(value=True)
        ttk.Label(f_par_hom, text="Nota Agronômica Paradigma:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
        e_nota_par = ttk.Entry(f_par_hom, width=8)
        e_nota_par.insert(0, "1,00")
        e_nota_par.grid(row=1, column=1, padx=(0, 14), pady=2)
        ttk.Checkbutton(f_par_hom, text="Aplicar Fator Nota Agronômica (Nota_par / Nota_amostra)", variable=var_usar_nota_hom).grid(row=1, column=2, columnspan=4, sticky="w", padx=4)

        cols_hom = ("id", "mun", "area", "q_orig", "f_of", "f_ar", "f_nt", "f_tot", "q_hom")
        tree_hom = ttk.Treeview(tab_hom, columns=cols_hom, show="headings", height=11)
        headers_hom = [
            ("id", "D.", 42), ("mun", "Município", 130), ("area", f"Área ({un_str})", 95),
            ("q_orig", "Unit. Original", 115), ("f_of", "F. Oferta", 75), ("f_ar", "F. Área", 75),
            ("f_nt", "F. Nota", 75), ("f_tot", "F. Total", 80), ("q_hom", "Unit. Homogeneizado", 135)
        ]
        for cid, ctxt, cw in headers_hom:
            tree_hom.heading(cid, text=ctxt)
            tree_hom.column(cid, width=cw, anchor="center" if cid != "mun" else "w")
        tree_hom.pack(fill="both", expand=True, pady=4)

        lbl_res_hom = tk.Label(tab_hom, text="", bg="#e1e7ef", fg="#1f3c5b", font=("Segoe UI", 9, "bold"), justify="left", anchor="w", padx=10, pady=6, relief="solid", bd=1)
        lbl_res_hom.pack(fill="x", pady=4)

        resultados_hom_cache = []

        def recalcular_homogeneizacao(mostrar_alerta=True):
            f_of_cfg, ok_of = converter_para_float_validado(e_f_oferta.get(), False)
            if not ok_of or not (0.10 <= f_of_cfg <= 2.00):
                if mostrar_alerta:
                    messagebox.showerror("Parâmetro Inválido", "O 'Fator Oferta' deve ser um número válido entre 0,10 e 2,00.", parent=janela)
                return

            a_par, ok_ap = converter_para_float_validado(e_area_par.get(), permitir_milhar_ponto=(un_str != "ha"))
            if not ok_ap or a_par <= 0:
                if mostrar_alerta:
                    messagebox.showerror("Parâmetro Inválido", "A 'Área Paradigma' deve ser um número maior que zero.", parent=janela)
                return

            exp_p, ok_exp = converter_para_float_validado(e_exp_area.get(), False)
            if not ok_exp or not (-2.0 <= exp_p <= 2.0):
                if mostrar_alerta:
                    messagebox.showerror("Parâmetro Inválido", "O 'Expoente Área (p)' deve ser um número válido (ex: 0,00 ou 0,125).", parent=janela)
                return

            n_par, ok_np = converter_para_float_validado(e_nota_par.get(), False)
            if not ok_np or not (0.01 <= n_par <= 10.0):
                if mostrar_alerta:
                    messagebox.showerror("Parâmetro Inválido", "A 'Nota Agronômica Paradigma' deve ser um número maior que zero (ex: 1,00).", parent=janela)
                return

            for item in tree_hom.get_children():
                tree_hom.delete(item)
            resultados_hom_cache.clear()

            q_homs, fatores_tot = [], []
            for d in ativos:
                extras = d.get("variaveis_extras", {})
                cond_txt = ""
                nota_val = 0.0
                for k, v in extras.items():
                    kl = normalizar_chave_var(k)
                    if "condi" in kl:
                        cond_txt = str(v).lower()
                    elif kl.startswith("nota agro"):
                        nota_val, _ = converter_para_float_validado(v, False)

                f_of = 1.0 if "transa" in cond_txt else f_of_cfg
                a_i = d["area_terreno"]
                f_ar = ((a_par / a_i) ** exp_p) if (a_par > 0 and a_i > 0 and exp_p != 0.0) else 1.0
                f_nt = (n_par / nota_val) if (var_usar_nota_hom.get() and nota_val > 0) else 1.0
                f_tot = f_of * f_ar * f_nt
                q_h = d["unitario"] * f_tot

                q_homs.append(q_h)
                fatores_tot.append(f_tot)
                resultados_hom_cache.append((d["dado_id"], d.get("municipio", ""), a_i, d["unitario"], f_of, f_ar, f_nt, f_tot, q_h))

                tree_hom.insert("", "end", values=(
                    f"D.{d['dado_id']}", d.get("municipio", ""),
                    formatar_numero_br(a_i, 4 if un_str == "ha" else 2),
                    f"R$ {formatar_moeda_br(d['unitario'])}",
                    formatar_numero_br(f_of, 3), formatar_numero_br(f_ar, 3),
                    formatar_numero_br(f_nt, 3), formatar_numero_br(f_tot, 3),
                    f"R$ {formatar_moeda_br(q_h)}"
                ))

            est_h = calcular_estatisticas_completas(q_homs)
            f_min, f_max = min(fatores_tot), max(fatores_tot)
            razao_f = (f_max / f_min) if f_min > 0 else 0.0
            status_nbr = (
                "✔ Dentro do limite NBR 14653 (fatores entre 0,50 e 2,00)"
                if (0.50 <= f_min and f_max <= 2.00 and razao_f <= 2.0)
                else "⚠ Atenção: Amplitude de fatores excede faixa 0,50–2,00 da NBR 14653"
            )
            grau_prec = (
                "Grau III (≤ 30%)" if est_h["amp_ic80_pct"] <= 30
                else ("Grau II (≤ 40%)" if est_h["amp_ic80_pct"] <= 40
                      else ("Grau I (≤ 50%)" if est_h["amp_ic80_pct"] <= 50 else "Fora de Grau (> 50%)"))
            )

            lbl_res_hom.config(
                text=(
                    f"📊 Média Homogeneizada: R$ {formatar_moeda_br(est_h['media'])}/{un_str}  (Original: R$ {formatar_moeda_br(est['media'])})   •   "
                    f"CV% Homogeneizado: {formatar_numero_br(est_h['cv'], 2)}%  (Original: {formatar_numero_br(est['cv'], 2)}%)\n"
                    f"🎯 IC 80%: R$ {formatar_moeda_br(est_h['ic80_inf'])} a R$ {formatar_moeda_br(est_h['ic80_sup'])}  •  "
                    f"Amplitude IC 80%: {formatar_numero_br(est_h['amp_ic80_pct'], 2)}% [{grau_prec}]   |   {status_nbr}"
                )
            )

        def exportar_hom_excel():
            if not resultados_hom_cache:
                return
            arq = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], parent=janela)
            if not arq:
                return
            df_h = pd.DataFrame(resultados_hom_cache, columns=[
                "Dado", "Município", f"Área ({un_str})", f"Unitário Original (R$/{un_str})",
                "Fator Oferta", "Fator Área", "Fator Nota", "Fator Resultante", f"Unitário Homogeneizado (R$/{un_str})"
            ])
            df_h.to_excel(arq, sheet_name="Homogeneização", index=False)
            messagebox.showinfo("Exportado", "Tabela de homogeneização exportada com sucesso!", parent=janela)

        f_btns_h = ttk.Frame(f_par_hom)
        f_btns_h.grid(row=0, column=6, rowspan=2, padx=8)
        ttk.Button(f_btns_h, text="🔄 Recalcular", style="Primary.TButton", command=lambda: recalcular_homogeneizacao(True)).pack(fill="x", pady=2)
        ttk.Button(f_btns_h, text="📊 Exportar Excel", command=exportar_hom_excel).pack(fill="x", pady=2)

        recalcular_homogeneizacao(False)

        # === ABA 3: REGRESSÃO LINEAR MÚLTIPLA (PRÉ-ANÁLISE SISDEA) ===
        tab_reg = ttk.Frame(nb, padding=8)
        nb.add(tab_reg, text=" 3. Regressão Linear Múltipla (SISDEA) ")

        f_cfg_reg = ttk.LabelFrame(tab_reg, text=" Configuração do Modelo Econométrico (MQO / OLS) ", padding=8)
        f_cfg_reg.pack(fill="x", pady=(0, 6))

        ttk.Label(f_cfg_reg, text="Transformação de Y (Unitário):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=4)
        var_transf_y = tk.StringVar(value="LN")
        ttk.Radiobutton(f_cfg_reg, text="Semilog: ln(Unitário) [Padrão SISDEA]", value="LN", variable=var_transf_y).grid(row=0, column=1, padx=6, sticky="w")
        ttk.Radiobutton(f_cfg_reg, text="Linear: Unitário", value="LIN", variable=var_transf_y).grid(row=0, column=2, padx=6, sticky="w")
        ttk.Radiobutton(f_cfg_reg, text="Inverso: 1 / Unitário", value="INV", variable=var_transf_y).grid(row=0, column=3, padx=6, sticky="w")

        candidatas_x = [("Área Terreno", "area_terreno")]
        if self.campos_opcionais.get("area_construida", True):
            candidatas_x.append(("Área Construída", "area_construida"))
        for v_cfg in self.variaveis_config:
            candidatas_x.append((v_cfg["nome"], f"var_{v_cfg['nome']}"))

        f_vars_x = ttk.Frame(f_cfg_reg)
        f_vars_x.grid(row=1, column=0, columnspan=5, sticky="w", pady=(6, 2))
        ttk.Label(f_vars_x, text="Variáveis Explicativas (X):", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(4, 8))

        vars_x_checks = []
        for nome_c, chave_c in candidatas_x:
            qtd_validos = sum(1 for d in ativos if extrair_valor_numerico_variavel(d, chave_c) is not None)
            vals_distintos = {extrair_valor_numerico_variavel(d, chave_c) for d in ativos if extrair_valor_numerico_variavel(d, chave_c) is not None}
            ativo_init = (qtd_validos >= max(3, len(ativos) // 2)) and (len(vals_distintos) > 1)
            v_chk = tk.BooleanVar(value=ativo_init)
            v_ln = tk.BooleanVar(value=(chave_c == "area_terreno"))
            f_item = ttk.Frame(f_vars_x)
            f_item.pack(side="left", padx=5)
            ttk.Checkbutton(f_item, text=nome_c, variable=v_chk).pack(side="left")
            ttk.Checkbutton(f_item, text="ln", variable=v_ln).pack(side="left")
            vars_x_checks.append((nome_c, chave_c, v_chk, v_ln))

        lbl_kpi_reg = tk.Label(tab_reg, text="", bg="#e1e7ef", fg="#1f3c5b", font=("Segoe UI", 9, "bold"), justify="left", anchor="w", padx=10, pady=5, relief="solid", bd=1)
        lbl_kpi_reg.pack(fill="x", pady=3)

        pane_reg = ttk.PanedWindow(tab_reg, orient="vertical")
        pane_reg.pack(fill="both", expand=True, pady=2)

        f_coef = ttk.LabelFrame(pane_reg, text=" Significância dos Regressores (Escala Transformada — Teste t & Multicolinearidade VIF) ", padding=4)
        pane_reg.add(f_coef, weight=1)
        cols_coef = ("var", "beta", "se", "t", "pval", "r", "vif", "diag")
        tree_coef = ttk.Treeview(f_coef, columns=cols_coef, show="headings", height=5)
        for cid, ctxt, cw in [
            ("var", "Variável", 165), ("beta", "Coeficiente (β)", 105), ("se", "Erro Padrão", 95),
            ("t", "Estatística t", 82), ("pval", "p-valor", 75), ("r", "Correl. r", 75),
            ("vif", "VIF", 65), ("diag", "Enquadramento NBR 14653", 220)
        ]:
            tree_coef.heading(cid, text=ctxt)
            tree_coef.column(cid, width=cw, anchor="center" if cid not in ("var", "diag") else "w")
        tree_coef.pack(fill="both", expand=True)

        f_res_sim = ttk.Frame(pane_reg)
        pane_reg.add(f_res_sim, weight=1)

        f_res = ttk.LabelFrame(f_res_sim, text=" Resíduos e Pontos Influenciantes (Escala Original — Outliers |z| > 2,0) ", padding=4)
        f_res.pack(side="left", fill="both", expand=True, padx=(0, 6))
        cols_res = ("id", "obs", "est", "err", "z", "st")
        tree_res = ttk.Treeview(f_res, columns=cols_res, show="headings", height=6)
        for cid, ctxt, cw in [
            ("id", "Dado", 55), ("obs", "Observado", 105), ("est", "Calculado", 105),
            ("err", "Erro %", 75), ("z", "Res. Pad. z", 80), ("st", "Status", 110)
        ]:
            tree_res.heading(cid, text=ctxt)
            tree_res.column(cid, width=cw, anchor="center")
        tree_res.pack(fill="both", expand=True)

        f_sim = ttk.LabelFrame(f_res_sim, text=" Simulador (Imóvel Avaliando) ", padding=8)
        f_sim.pack(side="right", fill="y")
        entradas_sim = {}
        modelo_reg_atual = {"res": None, "chaves": [], "transf_x_log": set()}
        sim_ui = {"lbl": None}

        def calcular_simulacao():
            mod = modelo_reg_atual["res"]
            lbl_out = sim_ui["lbl"]
            if not mod or lbl_out is None:
                return
            x_vec = [1.0]
            for ch in modelo_reg_atual["chaves"]:
                ent = entradas_sim.get(ch)
                val_x, ok_x = converter_para_float_validado(ent.get() if ent else "0", permitir_milhar_ponto=True)
                if not ok_x:
                    messagebox.showwarning("Simulador", "Valor numérico inválido no simulador.", parent=janela)
                    return
                if ch in modelo_reg_atual["transf_x_log"]:
                    if val_x <= 0:
                        messagebox.showwarning("Simulador", "Variáveis em ln(X) exigem valor > 0.", parent=janela)
                        return
                    val_x = math.log(val_x)
                x_vec.append(val_x)
            yh = sum(mod["betas"][j] * x_vec[j] for j in range(len(x_vec)))
            if mod["transf_y"] == "LN":
                q_mediana = math.exp(yh)
                q_av = q_mediana * mod["fator_vies_ln"]
                det_ln = f"\n(Mediana: R$ {formatar_moeda_br(q_mediana)} | Viés: ×{mod['fator_vies_ln']:.4f})"
            elif mod["transf_y"] == "INV":
                if yh <= 1e-12:
                    lbl_out.config(text="Fora do domínio:\nŷ transformado ≤ 0")
                    return
                q_av = 1.0 / yh
                det_ln = ""
            else:
                q_av = yh
                det_ln = ""
            area_av = converter_para_float(entradas_sim["area_terreno"].get(), permitir_milhar_ponto=(un_str != "ha")) if "area_terreno" in entradas_sim else 0.0
            tot_av = q_av * area_av
            txt_tot = f"\nValor Total: R$ {formatar_moeda_br(tot_av)}" if area_av > 0 else ""
            lbl_out.config(text=f"Unitário Estimado:\nR$ {formatar_moeda_br(q_av)}/{un_str}{txt_tot}{det_ln}")

        def rodar_regressao():
            for item in tree_coef.get_children():
                tree_coef.delete(item)
            for item in tree_res.get_children():
                tree_res.delete(item)

            nomes_sel, chaves_sel, logs_sel = [], [], set()
            for nome_c, chave_c, v_chk, v_ln in vars_x_checks:
                if v_chk.get():
                    nomes_sel.append(nome_c)
                    chaves_sel.append(chave_c)
                    if v_ln.get():
                        logs_sel.add(chave_c)

            if not chaves_sel:
                lbl_kpi_reg.config(text="⚠ Selecione ao menos 1 variável explicativa (X) para rodar a regressão.", fg="#b22222")
                return

            res_reg, erro_msg = executar_regressao_linear_multipla(ativos, nomes_sel, chaves_sel, var_transf_y.get(), logs_sel)
            if not res_reg:
                lbl_kpi_reg.config(text=f"⚠ {erro_msg}", fg="#b22222")
                return

            modelo_reg_atual["res"] = res_reg
            modelo_reg_atual["chaves"] = chaves_sel
            modelo_reg_atual["transf_x_log"] = logs_sel

            qtd_out_res = sum(1 for r in res_reg["detalhes_amostras"] if r["outlier"])
            escala_lbl = {"LN": "ln(Y)", "INV": "1/Y", "LIN": "Y (Linear)"}[res_reg["transf_y"]]
            alerta_vif = f"  •  ⚠ Colinearidade Alta (VIF máx: {res_reg['max_vif']:.1f})" if res_reg["max_vif"] > 10.0 else ""
            lbl_kpi_reg.config(
                text=(
                    f"📈 Escala {escala_lbl}: n = {res_reg['n']}  •  R = {formatar_numero_br(res_reg['r'], 4)}  •  "
                    f"R² = {formatar_numero_br(res_reg['r2'] * 100, 2)}%  •  R² Ajustado = {formatar_numero_br(res_reg['r2_adj'] * 100, 2)}%  •  "
                    f"F = {formatar_numero_br(res_reg['f_stat'], 2)}  •  Outliers (|z|>2): {qtd_out_res}{alerta_vif}"
                ),
                fg="#1f3c5b"
            )

            for reg in res_reg["regressores"]:
                pv = reg["p_valor"]
                vif_val = reg["vif"]
                vif_txt = "—" if vif_val is None else ("∞" if math.isinf(vif_val) else f"{vif_val:.2f}")
                if reg["nome"].startswith("Intercepto"):
                    diag = "Termo independente"
                else:
                    if pv <= 0.10:
                        diag = "✔ Grau III (p ≤ 10%)"
                    elif pv <= 0.20:
                        diag = "✔ Grau II (p ≤ 20%)"
                    elif pv <= 0.30:
                        diag = "⚠ Grau I (p ≤ 30%)"
                    else:
                        diag = "✖ Não signif. (p > 30%)"
                    if vif_val is not None and vif_val > 10.0:
                        diag += " [VIF>10]"
                tree_coef.insert("", "end", values=(
                    reg["nome"], f"{reg['beta']:.6f}", f"{reg['se']:.6f}",
                    f"{reg['t']:.3f}", f"{pv * 100:.2f}%", f"{reg['r_simples']:.3f}", vif_txt, diag
                ))

            for det in res_reg["detalhes_amostras"]:
                if det["fora_dominio"]:
                    est_str = "Fora domínio (ŷ≤0)"
                    err_str = "—"
                    st_r = "⚠ Inválido (1/ŷ)"
                else:
                    est_str = f"R$ {formatar_moeda_br(det['q_est'])}"
                    err_str = f"{det['erro_pct']:+.1f}%"
                    st_r = "⚠ Outlier (|z|>2)" if det["outlier"] else "✔ Normal"
                tree_res.insert("", "end", values=(
                    f"D.{det['dado_id']}", f"R$ {formatar_moeda_br(det['q_obs'])}",
                    est_str, err_str, f"{det['z_res']:+.2f}", st_r
                ))

            for w in f_sim.winfo_children():
                w.destroy()
            entradas_sim.clear()
            for r_i, (nm_s, ch_s) in enumerate(zip(nomes_sel, chaves_sel)):
                ttk.Label(f_sim, text=f"{nm_s}:", font=("Segoe UI", 8)).grid(row=r_i, column=0, sticky="w", pady=2)
                e_s = ttk.Entry(f_sim, width=12)
                vals_med = [extrair_valor_numerico_variavel(d, ch_s) for d in ativos if extrair_valor_numerico_variavel(d, ch_s) is not None]
                if vals_med:
                    e_s.insert(0, formatar_numero_br(statistics.median(vals_med), 2))
                e_s.grid(row=r_i, column=1, padx=4, pady=2)
                entradas_sim[ch_s] = e_s
            ttk.Button(f_sim, text="Calcular Avaliação", style="Accent.TButton", command=calcular_simulacao).grid(row=len(chaves_sel), column=0, columnspan=2, pady=6)
            lbl_sim_out = tk.Label(f_sim, text="Unitário Estimado:\nR$ 0,00", bg="#eef2f6", fg="#1f3c5b", font=("Segoe UI", 8, "bold"), width=32, pady=8, relief="groove")
            lbl_sim_out.grid(row=len(chaves_sel) + 1, column=0, columnspan=2, pady=2)
            sim_ui["lbl"] = lbl_sim_out
            calcular_simulacao()

        ttk.Button(f_cfg_reg, text="▶ Calcular Regressão", style="Primary.TButton", command=rodar_regressao).grid(row=0, column=4, padx=10)
        rodar_regressao()

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

    # --- MAPA INTERATIVO HTML (LEAFLET NO NAVEGADOR COM FOTOS, CAMADAS E FILTROS) ---
    def _abrir_mapa_interativo_html(self):
        self._commit_editor_ativo()
        pontos = []
        for d in self.dados_pesquisas:
            lat, lon = utm_para_latlon(d.get("zona_utm", "22K"), d.get("coord_e", ""), d.get("coord_s", ""))
            if lat is not None and lon is not None:
                pontos.append((d, lat, lon))

        if not pontos:
            messagebox.showwarning("Mapa Interativo", "Nenhuma amostra possui coordenadas UTM válidas (Coord. E e Coord. S).")
            return

        marcadores_js = []
        for d, lat, lon in pontos:
            d_id = d["dado_id"]
            un = d.get("unidade", self.unidade_padrao_projeto)
            casas = 4 if un == "ha" else 2
            ativo = bool(d.get("utilizado", True))

            foto_b64_html = ""
            for cf in ("foto1", "foto2"):
                src = d.get(cf, "")
                if not src:
                    continue
                stream = self._baixar_imagem(src)
                if stream:
                    b64 = base64.b64encode(stream.getvalue()).decode("ascii")
                    foto_b64_html += f'<img src="data:image/jpeg;base64,{b64}" style="width:100%;max-height:160px;object-fit:cover;border-radius:6px;margin:6px 0;border:1px solid #cbd5e1;"/>'

            link_html = ""
            lk = d.get("link", "").strip()
            if lk:
                href = lk if lk.startswith(("http://", "https://")) else f"https://{lk}"
                link_html = f'<div style="margin-top:6px;"><a href="{xml_escape(href)}" target="_blank" style="color:#1d4ed8;font-weight:bold;text-decoration:none;">🔗 Abrir Anúncio Original</a></div>'

            extras_html = ""
            for v_cfg in self.variaveis_config:
                vv = d.get("variaveis_extras", {}).get(v_cfg["nome"], "")
                if vv:
                    extras_html += f"<div><b>{xml_escape(v_cfg['nome'])}:</b> {xml_escape(vv)}</div>"

            popup_html = (
                f'<div style="font-family:Segoe UI,Arial,sans-serif;font-size:12px;width:290px;color:#1e293b;">'
                f'<div style="background:{"#1f3c5b" if ativo else "#64748b"};color:#fff;padding:6px 10px;border-radius:6px 6px 0 0;font-weight:bold;">'
                f'Pesquisa {d_id:02d}{"" if ativo else " (Descartada)"}</div>'
                f'<div style="padding:8px;border:1px solid #cbd5e1;border-top:none;border-radius:0 0 6px 6px;">'
                f'{foto_b64_html}'
                f'<div><b>Logradouro:</b> {xml_escape(d.get("endereco", ""))}</div>'
                f'<div><b>Município:</b> {xml_escape(d.get("municipio", ""))}</div>'
                f'<div><b>Área:</b> {formatar_numero_br(d.get("area_terreno", 0.0), casas)} {xml_escape(un)}</div>'
                f'<div><b>Valor:</b> R$ {formatar_moeda_br(d.get("valor_total", 0.0))}</div>'
                f'<div><b>Unitário:</b> <span style="color:#0f766e;font-weight:bold;">R$ {formatar_moeda_br(d.get("unitario", 0.0))}/{xml_escape(un)}</span></div>'
                f'{extras_html}'
                f'<div style="color:#64748b;font-size:11px;margin-top:4px;">UTM: {xml_escape(d.get("zona_utm", ""))} {xml_escape(d.get("coord_e", ""))} E / {xml_escape(d.get("coord_s", ""))} S</div>'
                f'{link_html}</div></div>'
            )

            marcadores_js.append({
                "id": d_id,
                "label": f"D.{d_id:02d}",
                "lat": lat,
                "lon": lon,
                "ativo": ativo,
                "municipio": d.get("municipio", ""),
                "unitario": f"R$ {formatar_moeda_br(d.get('unitario', 0.0))}/{un}",
                "popup": popup_html
            })

        html_path = os.path.join(CONFIG_DIR, "mapa_interativo_enprol.html")
        dados_json_js = json.dumps(marcadores_js, ensure_ascii=False)

        html_doc = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<title>ENPROL - Mapa Interativo de Pesquisa de Mercado</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
  body {{ margin:0; padding:0; font-family:'Segoe UI',Arial,sans-serif; display:flex; height:100vh; }}
  #sidebar {{ width:320px; background:#f8fafc; border-right:1px solid #cbd5e1; display:flex; flex-direction:column; }}
  #header {{ background:#1f3c5b; color:#fff; padding:14px; }}
  #header h2 {{ margin:0 0 4px 0; font-size:15px; }}
  #header p {{ margin:0; font-size:12px; color:#cbd5e1; }}
  #controls {{ padding:10px; border-bottom:1px solid #e2e8f0; background:#eef2f6; font-size:12px; }}
  #lista {{ flex:1; overflow-y:auto; padding:8px; }}
  .card {{ background:#fff; border:1px solid #cbd5e1; border-radius:6px; padding:8px 10px; margin-bottom:6px; cursor:pointer; transition:0.15s; }}
  .card:hover {{ border-color:#1f3c5b; background:#f1f5f9; }}
  .card-title {{ font-weight:bold; color:#1f3c5b; font-size:13px; }}
  .card-sub {{ font-size:12px; color:#475569; margin-top:2px; }}
  #map {{ flex:1; }}
  .pin-label {{ background:#facc15; border:2px solid #78350f; border-radius:50%; color:#1e293b; font-weight:bold; font-size:11px; text-align:center; line-height:22px; box-shadow:0 2px 5px rgba(0,0,0,0.35); }}
  .pin-off {{ background:#cbd5e1; border:2px solid #475569; border-radius:50%; color:#334155; font-weight:bold; font-size:11px; text-align:center; line-height:22px; }}
</style>
</head>
<body>
<div id="sidebar">
  <div id="header">
    <h2>ENPROL • Mapa de Amostras</h2>
    <p>{len(pontos)} amostra(s) georreferenciada(s) • {data_hoje_br()}</p>
  </div>
  <div id="controls">
    <label><input type="checkbox" id="chkAtivos" checked onchange="filtrar()"> Mostrar Utilizadas (☑)</label><br/>
    <label><input type="checkbox" id="chkInativos" checked onchange="filtrar()"> Mostrar Descartadas (☐)</label>
  </div>
  <div id="lista"></div>
</div>
<div id="map"></div>
<script>
const pontos = {dados_json_js};
const osm = L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{ maxZoom: 19, attribution: '© OpenStreetMap' }});
const sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}', {{ maxZoom: 19, attribution: '© Esri Satellite' }});
const rotulos = L.tileLayer('https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{{z}}/{{y}}/{{x}}', {{ maxZoom: 19 }});
const hibrido = L.layerGroup([sat, rotulos]);
const topo = L.tileLayer('https://{{s}}.tile.opentopomap.org/{{z}}/{{x}}/{{y}}.png', {{ maxZoom: 17, attribution: '© OpenTopoMap' }});
const limpo = L.tileLayer('https://{{s}}.basemaps.cartocdn.com/light_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{ maxZoom: 19, attribution: '© CartoDB' }});

const map = L.map('map', {{ layers: [osm] }});
L.control.layers({{
  "Mapa de Ruas (OSM)": osm,
  "Satélite (Esri)": sat,
  "Híbrido (Satélite + Ruas)": hibrido,
  "Topográfico / Relevo": topo,
  "Limpo para Laudo (Carto)": limpo
}}).addTo(map);
L.control.scale({{ imperial: false }}).addTo(map);

const markers = [];
const bounds = [];
const listaDiv = document.getElementById('lista');

pontos.forEach(p => {{
  const icon = L.divIcon({{
    className: p.ativo ? 'pin-label' : 'pin-off',
    html: p.id,
    iconSize: [26, 26],
    iconAnchor: [13, 13]
  }});
  const m = L.marker([p.lat, p.lon], {{ icon: icon }}).bindPopup(p.popup, {{ maxWidth: 320 }});
  m.addTo(map);
  markers.push({{ data: p, marker: m }});
  bounds.push([p.lat, p.lon]);

  const card = document.createElement('div');
  card.className = 'card';
  card.id = 'card_' + p.id;
  card.innerHTML = `<div class="card-title">${{p.label}} • ${{p.municipio || 'Sem município'}}</div><div class="card-sub">${{p.unitario}} ${{p.ativo ? '' : '(Descartada)'}}</div>`;
  card.onclick = () => {{ map.setView([p.lat, p.lon], 16); m.openPopup(); }};
  listaDiv.appendChild(card);
}});

if (bounds.length > 0) map.fitBounds(bounds, {{ padding: [45, 45] }});

function filtrar() {{
  const showA = document.getElementById('chkAtivos').checked;
  const showI = document.getElementById('chkInativos').checked;
  markers.forEach(obj => {{
    const vis = obj.data.ativo ? showA : showI;
    const card = document.getElementById('card_' + obj.data.id);
    if (vis) {{ obj.marker.addTo(map); card.style.display = 'block'; }}
    else {{ map.removeLayer(obj.marker); card.style.display = 'none'; }}
  }});
}}
</script>
</body>
</html>"""

        try:
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(html_doc)
            webbrowser.open(f"file:///{os.path.abspath(html_path).replace(os.sep, '/')}")
        except Exception as e:
            messagebox.showerror("Erro ao Abrir Mapa Interativo", str(e))

    # --- EXPORTAÇÃO KML / KMZ (GOOGLE EARTH COM ALFINETE AMARELO, LINK E FOTO) ---
    def _exportar_kmz_kml(self):
        self._commit_editor_ativo()
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para gerar o arquivo KML/KMZ.")
            return

        pontos_validos = []
        ignorados = []
        for d in self.dados_pesquisas:
            lat, lon = utm_para_latlon(d.get("zona_utm", "22K"), d.get("coord_e", ""), d.get("coord_s", ""))
            if lat is None or lon is None:
                ignorados.append(f"D.{d['dado_id']}")
            else:
                pontos_validos.append((d, lat, lon))

        if not pontos_validos:
            messagebox.showwarning(
                "Coordenadas Ausentes",
                "Nenhuma amostra possui coordenadas UTM válidas (Coord. E e Coord. S) para plotar no Google Earth."
            )
            return

        caminho = filedialog.asksaveasfilename(
            defaultextension=".kmz",
            initialfile=f"Pesquisa_Mercado_{datetime.now().strftime('%Y-%m-%d')}.kmz",
            filetypes=[
                ("Google Earth KMZ (Recomendado - com fotos embutidas)", "*.kmz"),
                ("Google Earth KML (Texto puro)", "*.kml")
            ]
        )
        if not caminho:
            return

        is_kmz = caminho.lower().endswith(".kmz")
        fotos_zip = {}
        placemarks_xml = []

        for d, lat, lon in pontos_validos:
            d_id = d["dado_id"]
            un = d.get("unidade", self.unidade_padrao_projeto)
            casas = 4 if un == "ha" else 2
            status_mod = "" if d.get("utilizado", True) else " (Não utilizado no modelo)"
            nome_ponto = f"Pesquisa {d_id:02d}{status_mod}"

            imgs_html = []
            for idx_f, campo_f in enumerate(("foto1", "foto2"), start=1):
                src_f = d.get(campo_f, "")
                if not src_f:
                    continue
                if is_kmz:
                    stream = self._baixar_imagem(src_f)
                    if stream:
                        nome_int = f"fotos/P{d_id:02d}{'' if idx_f == 1 else f'_{idx_f}'}.jpg"
                        fotos_zip[nome_int] = stream.getvalue()
                        imgs_html.append(f'<div style="margin:6px 0;"><img src="{nome_int}" width="330" style="border-radius:4px; border:1px solid #ccc;"/></div>')
                elif src_f.startswith(("http://", "https://")):
                    imgs_html.append(f'<div style="margin:6px 0;"><img src="{xml_escape(src_f)}" width="330" style="border-radius:4px;"/></div>')

            linhas_info = []
            if d.get("endereco"):
                linhas_info.append(f"<b>Logradouro:</b> {xml_escape(d['endereco'])}")
            if d.get("bairro"):
                linhas_info.append(f"<b>Bairro:</b> {xml_escape(d['bairro'])}")
            if d.get("municipio"):
                linhas_info.append(f"<b>Município:</b> {xml_escape(d['municipio'])}")
            if d.get("informante") or d.get("telefone"):
                contato = " - ".join(x for x in (d.get("telefone", ""), d.get("informante", "")) if x)
                linhas_info.append(f"<b>Contato:</b> {xml_escape(contato)}")

            linhas_info.append(f"<b>Área Terreno:</b> {formatar_numero_br(d.get('area_terreno', 0.0), casas)} {xml_escape(un)}")
            ac_val = converter_para_float(d.get("area_construida", 0.0))
            if ac_val > 0:
                linhas_info.append(f"<b>Área Construída:</b> {formatar_numero_br(ac_val, 2)} m²")

            lbl_val = "Valor do Aluguel" if self.finalidade_ativa == "ALUGUEL" else "Valor da Oferta"
            linhas_info.append(f"<b>{lbl_val}:</b> R$ {formatar_moeda_br(d.get('valor_total', 0.0))}")
            linhas_info.append(f"<b>Valor Unitário:</b> R$ {formatar_moeda_br(d.get('unitario', 0.0))}/{xml_escape(un)}")

            extras = d.get("variaveis_extras", {})
            for v_cfg in self.variaveis_config:
                v_val = extras.get(v_cfg["nome"], "")
                if v_val:
                    linhas_info.append(f"<b>{xml_escape(v_cfg['nome'])}:</b> {xml_escape(v_val)}")

            zona_s = f"{d.get('zona_utm', '').strip()} " if d.get("zona_utm") else ""
            linhas_info.append(f"<b>Coordenadas UTM:</b> {xml_escape(zona_s)}{xml_escape(d.get('coord_e', ''))} m E / {xml_escape(d.get('coord_s', ''))} m S")
            linhas_info.append(f"<b>Data:</b> {xml_escape(d.get('data', ''))}")

            link_url = d.get("link", "").strip()
            if link_url:
                href_esc = xml_escape(link_url if link_url.startswith(("http://", "https://")) else f"https://{link_url}")
                linhas_info.append(f'<b>Link da Pesquisa:</b> <a href="{href_esc}" target="_blank">Abrir Anúncio</a>')

            html_balao = (
                '<div style="font-family:Arial,sans-serif; font-size:12px; width:350px; color:#1f2937;">'
                f'<div style="background:#1f3c5b; color:#fff; padding:6px 10px; font-weight:bold; font-size:13px; border-radius:4px 4px 0 0;">{xml_escape(nome_ponto)}</div>'
                f'<div style="padding:8px; border:1px solid #cbd5e1; border-top:none;">'
                f'{"".join(imgs_html)}'
                f'<div style="line-height:1.45;">{"<br/>".join(linhas_info)}</div>'
                '</div></div>'
            )

            placemarks_xml.append(
                "    <Placemark>\n"
                f"      <name>{xml_escape(nome_ponto)}</name>\n"
                "      <styleUrl>#pinoAmarelo</styleUrl>\n"
                f"      <description><![CDATA[{html_balao}]]></description>\n"
                "      <Point>\n"
                f"        <coordinates>{lon:.7f},{lat:.7f},0</coordinates>\n"
                "      </Point>\n"
                "    </Placemark>"
            )

        kml_conteudo = (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<kml xmlns="http://www.opengis.net/kml/2.2">\n'
            '  <Document>\n'
            f'    <name>{xml_escape(os.path.splitext(os.path.basename(caminho))[0])}</name>\n'
            '    <Style id="pinoAmarelo">\n'
            '      <IconStyle>\n'
            '        <scale>1.1</scale>\n'
            '        <Icon>\n'
            '          <href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>\n'
            '        </Icon>\n'
            '        <hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>\n'
            '      </IconStyle>\n'
            '    </Style>\n'
            + "\n".join(placemarks_xml) + "\n"
            '  </Document>\n'
            '</kml>\n'
        )

        try:
            if is_kmz:
                with zipfile.ZipFile(caminho, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr("doc.kml", kml_conteudo.encode("utf-8"))
                    for nome_int, dados_img in fotos_zip.items():
                        zf.writestr(nome_int, dados_img)
            else:
                with open(caminho, "w", encoding="utf-8") as f:
                    f.write(kml_conteudo)

            aviso_ign = f"\n\nAmostras sem coordenadas (ignoradas): {', '.join(ignorados)}" if ignorados else ""
            messagebox.showinfo(
                "KML / KMZ Gerado",
                f"Arquivo gerado com sucesso com {len(pontos_validos)} alfinete(s) amarelo(s)!"
                + (f"\n{len(fotos_zip)} foto(s) embutida(s) no pacote KMZ." if is_kmz else "")
                + aviso_ign
            )
        except Exception as e:
            messagebox.showerror("Erro ao Gerar KML/KMZ", str(e))

    # --- GERADOR DE MAPA JPG / PNG TURBINADO (CACHE + DOWNLOAD PARALELO + 6 CAMADAS + LEGENDA COMPLETA) ---
    def _baixar_tile_com_cache(self, camada_id, url_t):
        chave_hash = hashlib.md5(f"{camada_id}_{url_t}".encode("utf-8")).hexdigest()
        arq_cache = os.path.join(TILES_CACHE_DIR, f"{chave_hash}.png")
        if os.path.exists(arq_cache):
            try:
                return Image.open(arq_cache).convert("RGBA")
            except Exception:
                pass
        try:
            r_t = requests.get(url_t, headers={"User-Agent": "ENPROL-PesquisaMercado/7.1"}, timeout=4)
            if r_t.status_code == 200:
                im = Image.open(io.BytesIO(r_t.content)).convert("RGBA")
                try:
                    im.save(arq_cache, "PNG")
                except Exception:
                    pass
                return im
        except Exception:
            pass
        return None

    def _janela_gerar_mapa_imagem(self):
        self._commit_editor_ativo()
        pontos = []
        for d in self.dados_pesquisas:
            lat, lon = utm_para_latlon(d.get("zona_utm", "22K"), d.get("coord_e", ""), d.get("coord_s", ""))
            if lat is not None and lon is not None:
                pontos.append((d, lat, lon))

        if not pontos:
            messagebox.showwarning("Mapa das Amostras", "Nenhuma amostra possui coordenadas UTM válidas para gerar o mapa.")
            return

        janela = tk.Toplevel(self.root)
        janela.title(f"Gerar Mapa de Localização das Amostras ({len(pontos)} pontos)")
        janela.geometry("980x690")
        janela.transient(self.root)

        f_ctrl = ttk.Frame(janela, padding=8)
        f_ctrl.pack(fill="x")

        ttk.Label(f_ctrl, text="Título do Mapa:", font=("Segoe UI", 9, "bold")).pack(side="left")
        e_tit_mapa = ttk.Entry(f_ctrl, width=30)
        e_tit_mapa.insert(0, "MAPA DE LOCALIZAÇÃO DAS AMOSTRAS")
        e_tit_mapa.pack(side="left", padx=(4, 10))

        camadas_disp = [
            "Mapa de Ruas (OSM)",
            "Satélite (Esri)",
            "Híbrido (Satélite + Ruas)",
            "Topográfico / Relevo",
            "Limpo para Laudo (Carto)",
            "Esquemático Offline"
        ]
        ttk.Label(f_ctrl, text="Camada:").pack(side="left")
        cb_estilo = ttk.Combobox(f_ctrl, values=camadas_disp, width=23, state="readonly")
        cb_estilo.current(0)
        cb_estilo.pack(side="left", padx=(4, 10))
        cb_estilo.bind("<<ComboboxSelected>>", lambda e: disparar_render_async())

        lbl_mapa = tk.Label(janela, text="Gerando prévia do mapa...", bg="#e2e8f0")
        lbl_mapa.pack(fill="both", expand=True, padx=10, pady=4)

        imagem_final_holder = {"img": None}

        def renderizar_mapa_pil(estilo_sel, titulo_txt):
            W, H = 1360, 860
            duas_colunas_leg = len(pontos) > 25
            leg_w = 460 if duas_colunas_leg else 340
            map_w = W - leg_w - 40
            map_h = H - 110
            map_x0, map_y0 = 20, 65

            lats = [p[1] for p in pontos]
            lons = [p[2] for p in pontos]
            lat_min, lat_max = min(lats), max(lats)
            lon_min, lon_max = min(lons), max(lons)

            d_lat = max((lat_max - lat_min) * 0.28, 0.008)
            d_lon = max((lon_max - lon_min) * 0.28, 0.008)
            lat_min -= d_lat
            lat_max += d_lat
            lon_min -= d_lon
            lon_max += d_lon

            def latlon_to_global_px(lat, lon, z):
                n_tiles = 2.0 ** z
                x_px = (lon + 180.0) / 360.0 * n_tiles * 256.0
                lat_rad = math.radians(max(-85.0511, min(85.0511, lat)))
                y_px = (1.0 - math.log(math.tan(lat_rad) + 1.0 / math.cos(lat_rad)) / math.pi) / 2.0 * n_tiles * 256.0
                return x_px, y_px

            zoom = 14
            for z_try in range(17, 4, -1):
                x1, y1 = latlon_to_global_px(lat_max, lon_min, z_try)
                x2, y2 = latlon_to_global_px(lat_min, lon_max, z_try)
                if (x2 - x1) <= map_w * 0.84 and (y2 - y1) <= map_h * 0.84:
                    zoom = z_try
                    break

            lat_c = (lat_min + lat_max) / 2.0
            lon_c = (lon_min + lon_max) / 2.0
            cx_g, cy_g = latlon_to_global_px(lat_c, lon_c, zoom)
            gx0 = cx_g - map_w / 2.0
            gy0 = cy_g - map_h / 2.0

            canvas_img = Image.new("RGB", (W, H), (244, 246, 249))
            draw = ImageDraw.Draw(canvas_img)

            mapa_base = Image.new("RGBA", (map_w, map_h), (235, 240, 245, 255))
            tiles_ok = 0

            if "Offline" not in estilo_sel:
                tx_min = int(math.floor(gx0 / 256.0))
                tx_max = int(math.floor((gx0 + map_w) / 256.0))
                ty_min = int(math.floor(gy0 / 256.0))
                ty_max = int(math.floor((gy0 + map_h) / 256.0))
                max_t = (2 ** zoom) - 1

                tarefas = []
                for tx in range(tx_min, tx_max + 1):
                    for ty in range(ty_min, ty_max + 1):
                        if not (0 <= tx <= max_t and 0 <= ty <= max_t):
                            continue
                        if "Híbrido" in estilo_sel:
                            tarefas.append((tx, ty, 0, "sat", f"https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{zoom}/{ty}/{tx}"))
                            tarefas.append((tx, ty, 1, "ref", f"https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{zoom}/{ty}/{tx}"))
                        elif "Satélite" in estilo_sel:
                            tarefas.append((tx, ty, 0, "sat", f"https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{zoom}/{ty}/{tx}"))
                        elif "Topográfico" in estilo_sel:
                            tarefas.append((tx, ty, 0, "topo", f"https://tile.opentopomap.org/{zoom}/{tx}/{ty}.png"))
                        elif "Limpo" in estilo_sel:
                            tarefas.append((tx, ty, 0, "carto", f"https://basemaps.cartocdn.com/light_all/{zoom}/{tx}/{ty}.png"))
                        else:
                            tarefas.append((tx, ty, 0, "osm", f"https://tile.openstreetmap.org/{zoom}/{tx}/{ty}.png"))

                if len(tarefas) <= 96:
                    resultados_tiles = []
                    with ThreadPoolExecutor(max_workers=10) as pool:
                        futuros = {
                            pool.submit(self._baixar_tile_com_cache, cid, u): (tx, ty, ordem)
                            for (tx, ty, ordem, cid, u) in tarefas
                        }
                        for fut in as_completed(futuros):
                            tx, ty, ordem = futuros[fut]
                            im_t = fut.result()
                            if im_t is not None:
                                resultados_tiles.append((ordem, tx, ty, im_t))

                    resultados_tiles.sort(key=lambda item: item[0])
                    for ordem, tx, ty, im_t in resultados_tiles:
                        px_paste = int(tx * 256 - gx0)
                        py_paste = int(ty * 256 - gy0)
                        if ordem == 0:
                            mapa_base.paste(im_t, (px_paste, py_paste))
                            tiles_ok += 1
                        else:
                            mapa_base.alpha_composite(im_t, (px_paste, py_paste))

            mapa_rgb = mapa_base.convert("RGB")
            draw_m = ImageDraw.Draw(mapa_rgb)
            if tiles_ok == 0:
                for gx in range(0, map_w, 80):
                    draw_m.line([(gx, 0), (gx, map_h)], fill=(214, 222, 232), width=1)
                for gy in range(0, map_h, 80):
                    draw_m.line([(0, gy), (map_w, gy)], fill=(214, 222, 232), width=1)

            try:
                f_tit = ImageFont.truetype("arialbd.ttf", 20)
                f_sub = ImageFont.truetype("arial.ttf", 12)
                f_pin = ImageFont.truetype("arialbd.ttf", 11)
                f_leg = ImageFont.truetype("arial.ttf", 10 if duas_colunas_leg else 11)
                f_leg_b = ImageFont.truetype("arialbd.ttf", 12)
            except Exception:
                f_tit = f_sub = f_pin = f_leg = f_leg_b = ImageFont.load_default()

            for d, lat, lon in pontos:
                px_g, py_g = latlon_to_global_px(lat, lon, zoom)
                px = int(px_g - gx0)
                py = int(py_g - gy0)
                ativo = d.get("utilizado", True)
                cor_pino = (250, 204, 21) if ativo else (160, 160, 160)
                cor_borda = (120, 53, 15) if ativo else (80, 80, 80)

                draw_m.polygon([(px, py), (px - 7, py - 16), (px + 7, py - 16)], fill=cor_borda)
                draw_m.ellipse([px - 10, py - 27, px + 10, py - 7], fill=cor_pino, outline=cor_borda, width=2)

                etiqueta = f"D.{d['dado_id']:02d}"
                draw_m.rectangle([px + 12, py - 24, px + 48, py - 9], fill=(255, 255, 255), outline=(31, 60, 91), width=1)
                draw_m.text((px + 15, py - 22), etiqueta, fill=(31, 60, 91), font=f_pin)

            nx, ny = map_w - 38, 42
            draw_m.ellipse([nx - 22, ny - 22, nx + 22, ny + 22], fill=(255, 255, 255), outline=(31, 60, 91), width=2)
            draw_m.polygon([(nx, ny - 15), (nx - 7, ny + 8), (nx, ny + 3), (nx + 7, ny + 8)], fill=(185, 28, 28))
            draw_m.text((nx - 4, ny - 29), "N", fill=(31, 60, 91), font=f_pin)

            metros_por_px = 156543.03392 * math.cos(math.radians(lat_c)) / (2 ** zoom)
            alvo_m = 100 * metros_por_px
            escala_m = 100
            for cand in (50, 100, 250, 500, 1000, 2000, 5000, 10000, 25000, 50000):
                if cand >= alvo_m * 0.7:
                    escala_m = cand
                    break
            barra_px = max(30, int(escala_m / metros_por_px))
            sx, sy = 20, map_h - 28
            draw_m.rectangle([sx - 6, sy - 18, sx + barra_px + 55, sy + 12], fill=(255, 255, 255), outline=(31, 60, 91))
            draw_m.rectangle([sx, sy, sx + barra_px, sy + 5], fill=(31, 60, 91))
            lbl_esc = f"{escala_m / 1000:.0f} km" if escala_m >= 1000 else f"{escala_m} m"
            draw_m.text((sx + barra_px + 6, sy - 5), lbl_esc, fill=(31, 60, 91), font=f_leg)

            canvas_img.paste(mapa_rgb, (map_x0, map_y0))
            draw.rectangle([map_x0, map_y0, map_x0 + map_w, map_y0 + map_h], outline=(31, 60, 91), width=2)

            draw.rectangle([20, 14, W - 20, 54], fill=(31, 60, 91))
            draw.text((34, 23), titulo_txt.upper() or "MAPA DE LOCALIZAÇÃO DAS AMOSTRAS", fill=(255, 255, 255), font=f_tit)
            draw.text((W - 265, 28), f"SIRGAS 2000 / UTM  •  {data_hoje_br()}", fill=(219, 234, 254), font=f_sub)

            lx0 = map_x0 + map_w + 14
            draw.rectangle([lx0, map_y0, W - 20, map_y0 + map_h], fill=(255, 255, 255), outline=(31, 60, 91), width=2)
            draw.rectangle([lx0, map_y0, W - 20, map_y0 + 32], fill=(43, 122, 120))
            draw.text((lx0 + 10, map_y0 + 8), f"LEGENDA ({len(pontos)} AMOSTRAS)", fill=(255, 255, 255), font=f_leg_b)

            # Legenda dinâmica em 1 ou 2 colunas exibindo TODAS as amostras
            area_leg_h = map_h - 44
            num_cols = 2 if duas_colunas_leg else 1
            itens_por_col = math.ceil(len(pontos) / num_cols)
            passo_y = max(14, min(26, area_leg_h // max(1, itens_por_col)))
            col_w = (leg_w - 16) // num_cols

            for idx_p, (d, _, _) in enumerate(pontos):
                col_i = idx_p // itens_por_col
                row_i = idx_p % itens_por_col
                cx = lx0 + 8 + col_i * col_w
                cy = map_y0 + 40 + row_i * passo_y
                un = d.get("unidade", self.unidade_padrao_projeto)
                mun_curto = d.get("municipio", "")[:10 if duas_colunas_leg else 15]
                txt_item = f"D.{d['dado_id']:02d} • {mun_curto} • R$ {formatar_moeda_br(d.get('unitario', 0.0))}/{un}"
                cor_bol = (250, 204, 21) if d.get("utilizado", True) else (180, 180, 180)
                draw.ellipse([cx, cy + 2, cx + 9, cy + 11], fill=cor_bol, outline=(120, 53, 15))
                draw.text((cx + 13, cy), txt_item, fill=(30, 41, 59), font=f_leg)

            draw.text((20, H - 32), f"ENPROL Engenharia  •  Total de amostras plotadas: {len(pontos)}  •  Camada: {estilo_sel}", fill=(71, 85, 105), font=f_sub)
            return canvas_img

        def disparar_render_async():
            estilo_sel = cb_estilo.get()
            titulo_txt = e_tit_mapa.get()
            lbl_mapa.config(image="", text="Carregando camada cartográfica (com cache rápido)...")

            def worker():
                img_gerada = renderizar_mapa_pil(estilo_sel, titulo_txt)
                if not janela.winfo_exists():
                    return
                imagem_final_holder["img"] = img_gerada
                prev = img_gerada.copy()
                prev.thumbnail((940, 570), LANCZOS_FILTER)

                def aplicar_ui():
                    if janela.winfo_exists():
                        tk_prev = ImageTk.PhotoImage(prev)
                        self._img_mapa_preview = tk_prev
                        lbl_mapa.config(image=tk_prev, text="")

                janela.after(0, aplicar_ui)

            threading.Thread(target=worker, daemon=True).start()

        def salvar_mapa_jpg():
            if imagem_final_holder["img"] is None:
                messagebox.showinfo("Aguarde", "O mapa ainda está sendo gerado.", parent=janela)
                return
            arq = filedialog.asksaveasfilename(
                defaultextension=".jpg",
                initialfile=f"Mapa_Amostras_{datetime.now().strftime('%Y-%m-%d')}.jpg",
                filetypes=[("Imagem JPEG", "*.jpg"), ("Imagem PNG", "*.png")],
                parent=janela
            )
            if not arq:
                return
            if arq.lower().endswith(".png"):
                imagem_final_holder["img"].save(arq, "PNG")
            else:
                imagem_final_holder["img"].save(arq, "JPEG", quality=92)
            messagebox.showinfo("Mapa Salvo", f"Mapa salvo com sucesso em:\n{arq}", parent=janela)

        ttk.Button(f_ctrl, text="🔄 Atualizar", command=disparar_render_async).pack(side="left", padx=4)
        ttk.Button(f_ctrl, text="💾 Salvar Mapa JPG", style="Primary.TButton", command=salvar_mapa_jpg).pack(side="right", padx=4)
        disparar_render_async()

    # --- CONVERSOR LAT/LONG <-> UTM ---
    def _janela_conversor_utm(self):
        janela = tk.Toplevel(self.root)
        janela.title("Conversor Geográfico (Lat/Long ↔ UTM SIRGAS 2000)")
        janela.geometry("490x280")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Cole Latitude/Longitude (graus decimais, GMS ou link do Google Maps):", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=15, pady=(12, 4))
        e_latlon = ttk.Entry(janela, width=56)
        e_latlon.pack(padx=15, pady=4)
        e_latlon.focus_set()

        lbl_prev_utm = ttk.Label(janela, text="Resultado UTM: —", font=("Segoe UI", 9, "bold"), foreground="#1f3c5b")
        lbl_prev_utm.pack(anchor="w", padx=15, pady=8)

        res_utm = {"zona": "", "e": "", "s": ""}

        def converter_agora():
            lat, lon = parse_latlon_texto(e_latlon.get())
            if lat is None or lon is None:
                messagebox.showwarning("Formato não reconhecido", "Informe coordenadas válidas (ex: -16.6868, -49.2648 ou link do Google Maps).", parent=janela)
                return
            try:
                z, ce, cs = latlon_para_utm(lat, lon)
            except ValueError as err_geo:
                messagebox.showerror("Coordenada Fora dos Limites", str(err_geo), parent=janela)
                return
            res_utm["zona"], res_utm["e"], res_utm["s"] = z, f"{ce:.2f}", f"{cs:.2f}"
            z_num = int(re.match(r"^\d+", z).group(0))
            aviso_br = "" if (18 <= z_num <= 25) else "  ⚠ (Fora das zonas do Brasil 18–25)"
            lbl_prev_utm.config(text=f"Resultado UTM: Zona {z}  |  E: {ce:.2f} m  |  S: {cs:.2f} m{aviso_br}")

        def aplicar_na_linha():
            if not res_utm["zona"]:
                converter_agora()
            if not res_utm["zona"]:
                return
            r = self.celula_ativa.get("row")
            if r is None or r >= len(self.dados_pesquisas):
                messagebox.showwarning("Selecione uma Linha", "Clique em uma linha da planilha antes de aplicar.", parent=janela)
                return
            d = self.dados_pesquisas[r]
            d["zona_utm"] = res_utm["zona"]
            d["coord_e"] = res_utm["e"]
            d["coord_s"] = res_utm["s"]
            self.alteracoes_pendentes = True
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            janela.destroy()

        f_b = ttk.Frame(janela, padding=10)
        f_b.pack(fill="x")
        ttk.Button(f_b, text="Converter", command=converter_agora).pack(side="left", padx=15)
        ttk.Button(f_b, text="✔ Aplicar na Linha Selecionada", style="Primary.TButton", command=aplicar_na_linha).pack(side="right", padx=15)

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

    def _perguntar_desambiguacao_nota(self, nota_fmt, candidatos):
        janela = tk.Toplevel(self.root)
        janela.title(f"Nota Agronômica {nota_fmt} — Selecionar Classe")
        janela.geometry("440x175")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(
            janela,
            text=f"A Nota Agronômica {nota_fmt} aparece em duas combinações na Tabela 02.\nSelecione qual Aptidão e Situação deseja preencher:",
            font=("Segoe UI", 9, "bold")
        ).pack(anchor="w", padx=16, pady=(12, 10))

        escolha_holder = {"res": candidatos[0]}
        for cl, st in candidatos:
            lbl_apt = obter_label_aptidao_configurada(cl, self.variaveis_config)

            def selecionar(c_val=cl, s_val=st):
                escolha_holder["res"] = (c_val, s_val)
                janela.destroy()

            ttk.Button(janela, text=f"{lbl_apt}  +  Situação: {st}", command=selecionar).pack(fill="x", padx=20, pady=3)

        self.root.wait_window(janela)
        return escolha_holder["res"]

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
        elif chave == "localizacao":
            dado["localizacao"] = sanitizar_texto_celula(novo_val)
        elif chave == "telefone":
            dado["telefone"] = sanitizar_telefone(novo_val)
        elif chave in ("coord_e", "coord_s"):
            dado[chave] = limpar_sufixo_coord(sanitizar_texto_celula(novo_val))
            ok_u, msg_u = validar_coordenadas_utm(dado.get("zona_utm", "22K"), dado.get("coord_e", ""), dado.get("coord_s", ""))
            if not ok_u and dado.get("coord_e") and dado.get("coord_s"):
                messagebox.showwarning("Coordenada UTM", msg_u)
        elif chave == "zona_utm":
            dado["zona_utm"] = sanitizar_texto_celula(novo_val).upper()
            ok_u, msg_u = validar_coordenadas_utm(dado.get("zona_utm", "22K"), dado.get("coord_e", ""), dado.get("coord_s", ""))
            if not ok_u and dado.get("coord_e") and dado.get("coord_s"):
                messagebox.showwarning("Zona UTM", msg_u)
        elif chave.startswith("var_"):
            nome_v = chave[4:]
            dado.setdefault("variaveis_extras", {})[nome_v] = sanitizar_texto_celula(novo_val)
            k_low = normalizar_chave_var(nome_v)
            if k_low.startswith("nota agro"):
                aplicar_nota_agronomica_no_dado(
                    dado,
                    self.variaveis_config,
                    origem_edicao="reverse",
                    escolhedor_ambiguo=self._perguntar_desambiguacao_nota
                )
            elif k_low.startswith(("aptid", "situa", "acess")):
                aplicar_nota_agronomica_no_dado(
                    dado,
                    self.variaveis_config,
                    origem_edicao="forward"
                )
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
        self._sincronizar_matriz_para_dados(reconstruir=True, avisar_invalidos=True)
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._autosave()

    def _ao_deletar_conteudo(self, event=None):
        self._sincronizar_matriz_para_dados(reconstruir=True, avisar_invalidos=False)
        self.alteracoes_pendentes = True
        self._atualizar_rotulo_topo()
        self._autosave()

    def _ao_modificar_grade_em_lote(self, event=None):
        self._sincronizar_matriz_para_dados(reconstruir=True, avisar_invalidos=False)
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

        fator_alq = float(self.config_word.get("fator_alqueire_ha", 4.84))
        janela = tk.Toplevel(self.root)
        janela.title(f"Extrair Dados de Anúncio (Pesquisa D.{self.dados_pesquisas[r]['dado_id']})")
        janela.geometry("510x340")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(
            janela,
            text=f"Cole abaixo a descrição do anúncio (1 alqueire = {formatar_numero_br(fator_alq, 2)} ha):",
            font=("Segoe UI", 9, "bold")
        ).pack(anchor="w", padx=12, pady=(10, 4))
        txt_box = tk.Text(janela, height=11, font=("Segoe UI", 9))
        txt_box.pack(fill="both", expand=True, padx=12, pady=4)
        txt_box.focus_set()

        def extrair_e_aplicar():
            res = extrair_dados_anuncio(txt_box.get("1.0", tk.END), fator_alqueire_ha=fator_alq)
            dado = self.dados_pesquisas[r]
            achados = []
            if res["valor_total"]:
                dado["valor_total"] = res["valor_total"]
                achados.append(f"Valor: R$ {formatar_moeda_br(res['valor_total'])}")
            if res["area_terreno"]:
                dado["area_terreno"] = res["area_terreno"]
                achados.append(f"Área: {res['area_terreno']} {res['unidade']}{res['detalhe_area']}")
            if res["unidade"]:
                dado["unidade"] = res["unidade"]
                dado["localizacao"] = "Rural" if res["unidade"] == "ha" else "Urbana"
            if res["telefone"]:
                dado["telefone"] = res["telefone"]
                achados.append(f"Telefone: {res['telefone']}")
            if res["link"]:
                dado["link"] = res["link"]
                achados.append("Link encontrado")
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
            self.pane_principal.sash_place(0, 960, 0)
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
                img = ImageOps.exif_transpose(Image.open(caminho_real))
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
        entry.delete(0, tk.END)
        entry.insert(0, arq)
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

    def _resolver_caminho_foto(self, caminho, dir_base_override=None):
        if not caminho:
            return ""
        if os.path.exists(caminho):
            return caminho
        base_dir = dir_base_override or (os.path.dirname(self.caminho_projeto_atual) if self.caminho_projeto_atual else None)
        if base_dir:
            rel = os.path.join(base_dir, caminho)
            if os.path.exists(rel):
                return rel
        return caminho

    # --- JANELA CAMPOS, MODELO E PERSONALIZAÇÃO DO WORD ---
    def _janela_campos_opcionais(self):
        self._commit_editor_ativo()
        janela = tk.Toplevel(self.root)
        janela.title("Campos, Modelo, Finalidade e Word")
        janela.geometry("490x590")
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

        f_word = ttk.LabelFrame(janela, text=" 4. Personalização do Relatório Word & Alqueire ", padding=8)
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

        ttk.Label(f_word, text="1 Alqueire em ha (Extrator):").grid(row=3, column=0, sticky="w", pady=2)
        e_alq = ttk.Entry(f_word, width=14)
        e_alq.insert(0, formatar_numero_br(self.config_word.get("fator_alqueire_ha", 4.84), 2))
        e_alq.grid(row=3, column=1, sticky="w", padx=4, pady=2)

        def aplicar():
            self.finalidade_ativa = var_fin.get()
            self.modelo_ativo = var_mod.get()
            for chave, v in vars_tk.items():
                self.campos_opcionais[chave] = v.get()
            self.config_word["titulo_padrao"] = e_tit.get().strip() or DEFAULT_CONFIG_WORD["titulo_padrao"]
            self.config_word["contrato_copasa"] = e_cop.get().strip() or DEFAULT_CONFIG_WORD["contrato_copasa"]
            self.config_word["rodape"] = e_rod.get().strip()
            f_alq_val = converter_para_float(e_alq.get(), False)
            if f_alq_val > 0:
                self.config_word["fator_alqueire_ha"] = f_alq_val
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

    # --- NOVO GERENCIADOR DE VARIÁVEIS (MESTRE-DETALHE MODERNO E INTUITIVO) ---
    def _janela_config_variaveis(self):
        self._commit_editor_ativo()
        janela = tk.Toplevel(self.root)
        janela.title("Gerenciador de Variáveis & Catálogo Global")
        janela.geometry("820x540")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        lista_trabalho = copy.deepcopy(self.variaveis_config)
        # Guarda nome original no projeto para migrar valores caso o usuário renomeie
        for item in lista_trabalho:
            item["_nome_orig"] = item["nome"]

        estado = {"idx": 0 if lista_trabalho else None, "bloqueado": False}

        # Barra superior: Perfis Rápidos + Inclusão do Catálogo
        f_top = ttk.LabelFrame(janela, text=" 1. Carregar Perfil Base ou Incluir do Catálogo Global ", padding=8)
        f_top.pack(fill="x", padx=12, pady=(10, 6))

        def aplicar_preset(preset_lista):
            lista_trabalho.clear()
            for it in copy.deepcopy(preset_lista):
                it["_nome_orig"] = it["nome"]
                lista_trabalho.append(it)
            estado["idx"] = 0 if lista_trabalho else None
            atualizar_listbox()

        ttk.Button(f_top, text="🏙 Urbano", command=lambda: aplicar_preset(DEFAULT_VARS_PADRAO)).pack(side="left", padx=3)
        ttk.Button(f_top, text="🌾 Rural (Agronômico)", command=lambda: aplicar_preset(DEFAULT_VARS_RURAL)).pack(side="left", padx=3)
        ttk.Button(f_top, text="💧 COPASA", command=lambda: aplicar_preset(DEFAULT_VARS_COPASA)).pack(side="left", padx=3)

        ttk.Separator(f_top, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Label(f_top, text="Catálogo:").pack(side="left")
        cb_cat = ttk.Combobox(f_top, values=[v["nome"] for v in self.catalogo_variaveis], width=24, state="readonly")
        if self.catalogo_variaveis:
            cb_cat.current(0)
        cb_cat.pack(side="left", padx=4)

        def incluir_do_catalogo():
            sincronizar_editor_para_item()
            escolha = cb_cat.get()
            item_cat = next((v for v in self.catalogo_variaveis if v["nome"] == escolha), None)
            if not item_cat:
                return
            if any(normalizar_chave_var(x["nome"]) == normalizar_chave_var(item_cat["nome"]) for x in lista_trabalho):
                messagebox.showinfo("Já existe", f"A variável '{item_cat['nome']}' já está na lista ativa.", parent=janela)
                return
            novo = copy.deepcopy(item_cat)
            novo["_nome_orig"] = novo["nome"]
            lista_trabalho.append(novo)
            estado["idx"] = len(lista_trabalho) - 1
            atualizar_listbox()

        ttk.Button(f_top, text="➕ Incluir", style="Accent.TButton", command=incluir_do_catalogo).pack(side="left", padx=2)

        # Corpo Mestre-Detalhe
        f_corpo = ttk.Frame(janela, padding=(12, 4))
        f_corpo.pack(fill="both", expand=True)

        # Esquerda: Lista de Variáveis Ativas
        f_esq = ttk.LabelFrame(f_corpo, text=" 2. Colunas / Variáveis Ativas ", padding=8)
        f_esq.pack(side="left", fill="both", expand=True, padx=(0, 8))

        lb_vars = tk.Listbox(f_esq, font=("Segoe UI", 10), selectmode="browse", activestyle="none", relief="solid", bd=1)
        lb_vars.pack(fill="both", expand=True, pady=(0, 6))

        f_btns_esq = ttk.Frame(f_esq)
        f_btns_esq.pack(fill="x")

        # Direita: Painel de Edição da Variável Selecionada
        f_dir = ttk.LabelFrame(f_corpo, text=" 3. Propriedades da Variável Selecionada ", padding=12)
        f_dir.pack(side="right", fill="both", expand=True)

        ttk.Label(f_dir, text="Nome da Variável (Cabeçalho da Coluna):", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        e_nome = ttk.Entry(f_dir, width=38, font=("Segoe UI", 10))
        e_nome.pack(fill="x", pady=(2, 10))

        ttk.Label(f_dir, text="Tipo de Dado:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        var_tipo = tk.StringVar(value="texto")
        f_tipos = ttk.Frame(f_dir)
        f_tipos.pack(anchor="w", pady=(2, 10))
        ttk.Radiobutton(f_tipos, text="Lista Suspensa (Código)", value="codigo", variable=var_tipo, command=lambda: on_editar_campo()).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(f_tipos, text="Número", value="numero", variable=var_tipo, command=lambda: on_editar_campo()).pack(side="left", padx=(0, 10))
        ttk.Radiobutton(f_tipos, text="Texto Livre", value="texto", variable=var_tipo, command=lambda: on_editar_campo()).pack(side="left")

        ttk.Label(f_dir, text="Opções da Lista Suspensa (1 por linha ou separadas por vírgula):", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        txt_opcoes = tk.Text(f_dir, height=9, font=("Segoe UI", 9), relief="solid", bd=1)
        txt_opcoes.pack(fill="both", expand=True, pady=(4, 6))

        lbl_dica = ttk.Label(
            f_dir,
            text="Dica: Para variáveis numéricas na Regressão SISDEA, use prefixo numérico\n(ex.: '1 - Baixo', '2 - Normal', '3 - Alto').",
            font=("Segoe UI", 8),
            foreground="#475569"
        )
        lbl_dica.pack(anchor="w")

        def parse_opcoes_widget():
            bruto = txt_opcoes.get("1.0", tk.END)
            partes = re.split(r"[\n,;]+", bruto)
            return [p.strip() for p in partes if p.strip()]

        def sincronizar_editor_para_item():
            idx = estado["idx"]
            if estado["bloqueado"] or idx is None or not (0 <= idx < len(lista_trabalho)):
                return
            item = lista_trabalho[idx]
            nm = e_nome.get().strip()
            if nm:
                item["nome"] = nm
            tp = var_tipo.get()
            item["tipo"] = tp
            opts = parse_opcoes_widget()
            item["opcoes"] = ([""] + opts) if (tp == "codigo" and opts) else opts

        def carregar_item_no_editor(idx):
            estado["bloqueado"] = True
            if idx is None or not (0 <= idx < len(lista_trabalho)):
                e_nome.delete(0, tk.END)
                txt_opcoes.delete("1.0", tk.END)
                estado["bloqueado"] = False
                return
            item = lista_trabalho[idx]
            e_nome.delete(0, tk.END)
            e_nome.insert(0, item.get("nome", ""))
            var_tipo.set(item.get("tipo", "texto"))
            txt_opcoes.config(state="normal")
            txt_opcoes.delete("1.0", tk.END)
            opts_limpas = [o for o in item.get("opcoes", []) if o != ""]
            txt_opcoes.insert("1.0", "\n".join(opts_limpas))
            if item.get("tipo") == "numero":
                txt_opcoes.config(state="disabled")
            estado["bloqueado"] = False

        def atualizar_listbox():
            lb_vars.delete(0, tk.END)
            tipo_lbl = {"codigo": "Lista", "numero": "Número", "texto": "Texto"}
            for i, it in enumerate(lista_trabalho):
                lb_vars.insert(tk.END, f" {i + 1:02d}. {it['nome']}   [{tipo_lbl.get(it.get('tipo', 'texto'), 'Texto')}]")
            idx = estado["idx"]
            if idx is not None and 0 <= idx < len(lista_trabalho):
                lb_vars.selection_clear(0, tk.END)
                lb_vars.selection_set(idx)
                lb_vars.activate(idx)
                carregar_item_no_editor(idx)
            else:
                carregar_item_no_editor(None)

        def on_select_listbox(event=None):
            sel = lb_vars.curselection()
            if not sel:
                return
            sincronizar_editor_para_item()
            estado["idx"] = sel[0]
            carregar_item_no_editor(sel[0])

        def on_editar_campo(event=None):
            if estado["bloqueado"]:
                return
            txt_opcoes.config(state="disabled" if var_tipo.get() == "numero" else "normal")
            sincronizar_editor_para_item()
            idx = estado["idx"]
            if idx is not None and 0 <= idx < len(lista_trabalho):
                tipo_lbl = {"codigo": "Lista", "numero": "Número", "texto": "Texto"}
                it = lista_trabalho[idx]
                lb_vars.delete(idx)
                lb_vars.insert(idx, f" {idx + 1:02d}. {it['nome']}   [{tipo_lbl.get(it.get('tipo', 'texto'), 'Texto')}]")
                lb_vars.selection_set(idx)

        def nova_variavel():
            sincronizar_editor_para_item()
            n_num = len(lista_trabalho) + 1
            lista_trabalho.append({"nome": f"Variável {n_num}", "tipo": "codigo", "opcoes": ["", "1 - Opção A", "2 - Opção B"], "_nome_orig": None})
            estado["idx"] = len(lista_trabalho) - 1
            atualizar_listbox()
            e_nome.focus_set()
            e_nome.selection_range(0, tk.END)

        def mover_variavel(delta):
            sincronizar_editor_para_item()
            idx = estado["idx"]
            if idx is None:
                return
            novo_idx = idx + delta
            if 0 <= novo_idx < len(lista_trabalho):
                lista_trabalho[idx], lista_trabalho[novo_idx] = lista_trabalho[novo_idx], lista_trabalho[idx]
                estado["idx"] = novo_idx
                atualizar_listbox()

        def remover_variavel():
            idx = estado["idx"]
            if idx is None or not (0 <= idx < len(lista_trabalho)):
                return
            lista_trabalho.pop(idx)
            estado["idx"] = min(idx, len(lista_trabalho) - 1) if lista_trabalho else None
            atualizar_listbox()

        lb_vars.bind("<<ListboxSelect>>", on_select_listbox)
        e_nome.bind("<KeyRelease>", on_editar_campo)
        txt_opcoes.bind("<KeyRelease>", on_editar_campo)

        ttk.Button(f_btns_esq, text="➕ Nova Variável", style="Primary.TButton", command=nova_variavel).pack(side="left", padx=2)
        ttk.Button(f_btns_esq, text="▲ Subir", width=8, command=lambda: mover_variavel(-1)).pack(side="left", padx=2)
        ttk.Button(f_btns_esq, text="▼ Descer", width=8, command=lambda: mover_variavel(1)).pack(side="left", padx=2)
        ttk.Button(f_btns_esq, text="🗑 Remover", width=9, command=remover_variavel).pack(side="right", padx=2)

        atualizar_listbox()

        def salvar_e_fechar():
            sincronizar_editor_para_item()
            novas = []
            chaves_vistas = set()
            mapa_renomeacao = {}

            for it in lista_trabalho:
                nm = it.get("nome", "").strip()
                if not nm:
                    continue
                ch = normalizar_chave_var(nm)
                if ch in chaves_vistas:
                    messagebox.showerror("Nome Duplicado", f"Existem duas variáveis com o nome '{nm}'. Renomeie uma delas.", parent=janela)
                    return
                chaves_vistas.add(ch)
                orig = it.get("_nome_orig")
                if orig and orig != nm:
                    mapa_renomeacao[orig] = nm
                novas.append({
                    "nome": nm,
                    "tipo": it.get("tipo", "texto"),
                    "opcoes": list(it.get("opcoes", []))
                })

            if mapa_renomeacao:
                for d in self.dados_pesquisas:
                    ex = d.setdefault("variaveis_extras", {})
                    for antigo, novo in mapa_renomeacao.items():
                        if antigo in ex and novo not in ex:
                            ex[novo] = ex.pop(antigo)

            self.variaveis_config = novas
            self.alteracoes_pendentes = True
            self._registrar_no_catalogo_global(novas)
            self._salvar_config_variaveis()
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            janela.destroy()

        f_rodape_v = ttk.Frame(janela, padding=8)
        f_rodape_v.pack(fill="x")
        ttk.Button(f_rodape_v, text="✖ Cancelar", command=janela.destroy).pack(side="right", padx=12)
        ttk.Button(f_rodape_v, text="💾 Salvar e Atualizar Planilha", style="Primary.TButton", command=salvar_e_fechar).pack(side="right", padx=4)

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
            t_ref.insert("", "end", values=(sit, p_s) + tuple(TABELA_NOTA_AGRONOMICA[r][sit] for r in ("I", "II", "III", "IV", "V", "VI")))

    # --- ABRIR / SALVAR PROJETO (SEGURO, PROTEGIDO E ATÔMICO) ---
    def _salvar_projeto(self):
        self._commit_editor_ativo()
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Projeto de Pesquisa", "*.json")])
        if not caminho:
            return False
        dir_origem_anterior = os.path.dirname(self.caminho_projeto_atual) if self.caminho_projeto_atual else None
        self.caminho_projeto_atual = caminho
        pasta_anexos = self._obter_pasta_anexos_projeto()
        dir_proj = os.path.dirname(caminho)

        falhas_fotos = []
        for d in self.dados_pesquisas:
            for cf in ("foto1", "foto2"):
                orig = self._resolver_caminho_foto(d.get(cf, ""), dir_base_override=dir_origem_anterior)
                if orig and os.path.exists(orig) and not orig.startswith(("http://", "https://")):
                    try:
                        os.makedirs(pasta_anexos, exist_ok=True)
                        dest = os.path.join(pasta_anexos, os.path.basename(orig))
                        if os.path.abspath(orig) != os.path.abspath(dest):
                            shutil.copy2(orig, dest)
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
            return True
        except Exception as e:
            logging.error(f"Erro ao salvar projeto '{caminho}': {e}")
            messagebox.showerror("Erro", f"Erro ao salvar arquivo: {e}")
            return False

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
        if not self._verificar_descarte_seguro("abrir outro projeto"):
            return False
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

    # --- IMPORTAÇÃO EXCEL COM PRÉVIA E VALIDAÇÃO ---
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
            elif "localiza" in c: mapeamento["localizacao"] = col
            elif "data" in c: mapeamento["data"] = col
            elif "link" in c: mapeamento["link"] = col
            elif "unidade" in c or "unid" in c: mapeamento["unidade"] = col
            elif "utilizado" in c or "ativo" in c: mapeamento["utilizado"] = col

        novos, avisos_num = [], []
        for idx_df, row in df.iterrows():
            if pd.isna(row.get(mapeamento.get("informante", ""))) and pd.isna(row.get(mapeamento.get("endereco", ""))):
                continue
            if "unidade" in mapeamento and not pd.isna(row.get(mapeamento["unidade"])):
                un = normalizar_unidade(row.get(mapeamento["unidade"]), padrao=self.unidade_padrao_projeto)
            else:
                un = unidade_det or self.unidade_padrao_projeto

            v_tot, ok_v = converter_para_float_validado(row.get(mapeamento.get("valor_total", ""), 0.0))
            a_ter, ok_a = converter_para_float_validado(row.get(mapeamento.get("area_terreno", ""), 0.0), permitir_milhar_ponto=(un != "ha"))
            a_con, ok_ac = converter_para_float_validado(row.get(mapeamento.get("area_construida", ""), 0.0))
            if not (ok_v and ok_a and ok_ac):
                avisos_num.append(f"Linha {idx_df + 2}: valor/área inválido convertido como 0")

            util = True
            if "utilizado" in mapeamento and not pd.isna(row.get(mapeamento["utilizado"])):
                util = str(row.get(mapeamento["utilizado"])).strip().lower() not in ("não", "nao", "false", "0", "☐", "não utilizado no modelo")
            s_data = sanitizar_texto_celula(row.get(mapeamento["data"])) if "data" in mapeamento else ""
            loc_imp = sanitizar_texto_celula(row.get(mapeamento["localizacao"])) if "localizacao" in mapeamento else ""

            reg = {
                "dado_id": len(self.dados_pesquisas) + len(novos) + 1, "utilizado": util,
                "informante": sanitizar_texto_celula(row.get(mapeamento.get("informante", ""), "")),
                "telefone": sanitizar_telefone(row.get(mapeamento.get("telefone", ""), "")),
                "endereco": sanitizar_texto_celula(row.get(mapeamento.get("endereco", ""), "")),
                "bairro": sanitizar_texto_celula(row.get(mapeamento.get("bairro", ""), "")),
                "municipio": sanitizar_texto_celula(row.get(mapeamento.get("municipio", ""), "")),
                "valor_total": v_tot, "area_terreno": a_ter, "area_construida": a_con,
                "unitario": v_tot / a_ter if a_ter > 0 else 0.0, "unidade": un,
                "localizacao": loc_imp or ("Rural" if un == "ha" else "Urbana"),
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
                        val_bruto = row.get(col_p, "")
                        if alvo.startswith("nota agro"):
                            reg["variaveis_extras"][v_cfg["nome"]] = normalizar_nota_str(val_bruto)
                        else:
                            reg["variaveis_extras"][v_cfg["nome"]] = sanitizar_texto_celula(val_bruto)
                        break
            aplicar_nota_agronomica_no_dado(reg, self.variaveis_config, origem_edicao="auto")
            novos.append(reg)

        if novos:
            self.dados_pesquisas.extend(novos)
            self.alteracoes_pendentes = True
            self._atualizar_rotulo_topo()
            self._reconstruir_planilha()
            self._autosave()
            msg_ok = f"{len(novos)} pesquisas importadas com sucesso!"
            if avisos_num:
                msg_ok += "\n\n⚠ Atenção em células não numéricas:\n• " + "\n• ".join(avisos_num[:8])
            messagebox.showinfo("Importação Concluída", msg_ok)
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
                val_v = d.get("variaveis_extras", {}).get(v_cfg["nome"], "")
                if normalizar_chave_var(v_cfg["nome"]).startswith("nota agro") and val_v:
                    num_n, ok_n = converter_para_float_validado(val_v, permitir_milhar_ponto=False)
                    linha[v_cfg["nome"]] = num_n if ok_n else val_v
                else:
                    linha[v_cfg["nome"]] = val_v
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
                if "R$" in cab or "(m²)" in cab or "nota agro" in cab.lower():
                    for cell in col_cells[1:]:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0.00'
                elif "(ha)" in cab:
                    for cell in col_cells[1:]:
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0.0000'

        messagebox.showinfo("Sucesso", "Planilha Excel (SISDEA) formatada e exportada com sucesso!")

    # --- EXPORTAÇÃO WORD (SEGURA CONTRA SSRF E CONCORRÊNCIA) ---
    def _baixar_imagem(self, caminho_ou_url):
        if not caminho_ou_url:
            return None
        url = normalizar_url_gdrive(caminho_ou_url)
        try:
            if url.startswith(("http://", "https://")):
                if not url_eh_segura_publica(url):
                    logging.warning(f"URL bloqueada por política de segurança (host inválido/privado): '{url}'")
                    return None
                with requests.get(url, timeout=12, stream=True, allow_redirects=True) as resp:
                    if not url_eh_segura_publica(resp.url):
                        logging.warning(f"Redirecionamento bloqueado por política de segurança: '{resp.url}'")
                        return None
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
        if self.exportando_word:
            messagebox.showinfo("Exportação em Andamento", "Aguarde a conclusão da geração do documento Word atual.")
            return
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
        self.exportando_word = True
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
                    hdr.paragraphs[0]._element.get_parent().remove(hdr.paragraphs[0]._element)

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

                lote = dados_lista[i:i+2]
                for idx_lote, dado in enumerate(lote):
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
                        if campos_opc.get("localizacao", True): add_line(pc, "Localização:", dado.get("localizacao", ""))
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
                        add_line(p_dados, "Frente:", extras.get("Frente", ""))
                        p_dados.add_run("\n")
                        add_line(p_dados, "Via de acesso:", extras.get("Via de acesso", ""))
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

                    processadas = min(i + idx_lote + 1, total)
                    self.fila_eventos_ui.put(("progresso", int((processadas / max(1, total)) * 100)))

            doc.save(caminho)
            self.fila_eventos_ui.put(("concluido", None))
        except Exception as e:
            self.fila_eventos_ui.put(("erro", e))

    def _atualizar_progresso(self, valor):
        if hasattr(self, "progress_bar"):
            self.progress_bar["value"] = valor

    def _concluir_exportacao(self, payload=None):
        self.exportando_word = False
        if hasattr(self, "progress_bar"):
            self.progress_bar.pack_forget()
        messagebox.showinfo("Sucesso", "Fichas de Pesquisa no Word geradas com sucesso!")

    def _falha_exportacao(self, err):
        self.exportando_word = False
        if hasattr(self, "progress_bar"):
            self.progress_bar.pack_forget()
        messagebox.showerror("Erro na Exportação", f"Ocorreu um erro ao gerar o documento Word: {err}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppPesquisaMercado(root)
    root.mainloop()
