import os
import io
import re
import json
import logging
import threading
import copy
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pesquisa_mercado")
os.makedirs(CONFIG_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(CONFIG_DIR, "config_variaveis.json")

DEFAULT_VARS_PADRAO = [
    {"nome": "Via", "tipo": "codigo", "opcoes": ["1 - Local", "2 - Coletora", "3 - Arterial"]},
    {"nome": "Uso", "tipo": "codigo", "opcoes": ["1 - Residencial", "2 - Comercial", "3 - Misto"]},
    {"nome": "Testada (m)", "tipo": "numero", "opcoes": []},
    {"nome": "PGV (R$)", "tipo": "numero", "opcoes": []}
]

DEFAULT_VARS_COPASA = [
    {"nome": "Frente", "tipo": "texto", "opcoes": ["Não informado"]},
    {"nome": "Via de acesso", "tipo": "texto", "opcoes": []},
    {"nome": "Informações", "tipo": "texto", "opcoes": []}
]

def converter_para_float(texto):
    if texto is None:
        return 0.0
    s = str(texto).replace("R$", "").replace("m²", "").replace("ha", "").strip()
    if not s:
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
    """Garante compatibilidade total entre arquivos JSON antigos e novos."""
    d_id = dado.get("dado_id") if dado.get("dado_id") is not None else dado.get("D.", 1)
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
    data = dado.get("data") or dado.get("Data") or "18/09/2026"
    link = dado.get("link") or dado.get("Link") or ""
    foto1 = dado.get("foto1") or dado.get("Foto1") or ""
    foto2 = dado.get("foto2") or dado.get("Foto2") or ""
    localizacao = dado.get("localizacao") or dado.get("Localização") or ("Rural" if unidade == "ha" else "Urbana")

    vars_extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras") or {}

    return {
        "dado_id": int(d_id),
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
        self.root.title("ENPROL - Sistema de Pesquisa de Mercado")
        self.root.geometry("1100x820")
        self.root.minsize(980, 700)

        self._configurar_estilos()

        self.modelo_ativo = "PADRAO"
        self.dados_pesquisas = []
        self.item_em_edicao = None
        self.variaveis_config = self._carregar_config_variaveis()

        self.container_principal = ttk.Frame(self.root)
        self.container_principal.pack(fill="both", expand=True)

        self._exibir_tela_hub()

    def _configurar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        self.root.configure(bg="#f4f6f9")
        style.configure("TFrame", background="#f4f6f9")
        style.configure("TLabelframe", background="#f4f6f9", font=("Segoe UI", 9, "bold"))
        style.configure("TLabelframe.Label", background="#f4f6f9", foreground="#1f3c5b", font=("Segoe UI", 10, "bold"))
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

    def _carregar_config_variaveis(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logging.warning(f"Erro ao carregar configurações: {e}")
        return list(DEFAULT_VARS_PADRAO)

    def _salvar_config_variaveis(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.variaveis_config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("Erro ao salvar variáveis", str(e))

    # --- TELA 1: HUB INICIAL ---
    def _exibir_tela_hub(self):
        self._limpar_container()
        self.root.config(menu="")

        frame_hub = ttk.Frame(self.container_principal, padding=40)
        frame_hub.place(relx=0.5, rely=0.5, anchor="center")

        lbl_logo = ttk.Label(frame_hub, text="SISTEMA DE PESQUISA DE MERCADO", font=("Segoe UI", 18, "bold"), foreground="#1f3c5b")
        lbl_logo.pack(pady=(0, 5))

        lbl_sub = ttk.Label(frame_hub, text="Selecione uma opção para iniciar os trabalhos", font=("Segoe UI", 10), foreground="#666666")
        lbl_sub.pack(pady=(0, 30))

        btn_novo = ttk.Button(frame_hub, text="📄  Criar Novo Arquivo / Projeto", style="Hub.TButton", width=38, command=self._popup_escolha_modelo)
        btn_novo.pack(pady=8)

        btn_abrir = ttk.Button(frame_hub, text="📂  Carregar Projeto Existente (.json)", style="Hub.TButton", width=38, command=self._abrir_projeto_hub)
        btn_abrir.pack(pady=8)

        btn_importar = ttk.Button(frame_hub, text="📥  Importar Dados de Planilha (.xlsx)", style="Hub.TButton", width=38, command=self._importar_planilha_hub)
        btn_importar.pack(pady=8)

        btn_sair = ttk.Button(frame_hub, text="🚪  Sair do Sistema", width=38, command=self.root.quit)
        btn_sair.pack(pady=(20, 0))

    def _popup_escolha_modelo(self, callback_pos_confirmacao=None):
        janela = tk.Toplevel(self.root)
        janela.title("Selecionar Modelo de Ficha")
        janela.geometry("450x260")
        janela.resizable(False, False)
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Selecione o modelo do contrato:", font=("Segoe UI", 11, "bold")).pack(pady=(20, 15))

        var_mod = tk.StringVar(value="PADRAO")
        r1 = ttk.Radiobutton(janela, text="Modelo 1 - Padrão / Memória de Cálculo (Geral)", value="PADRAO", variable=var_mod)
        r1.pack(anchor="w", padx=40, pady=5)

        r2 = ttk.Radiobutton(janela, text="Modelo 2 - Contrato COPASA (Ficha Técnica)", value="COPASA", variable=var_mod)
        r2.pack(anchor="w", padx=40, pady=5)

        def confirmar():
            self.modelo_ativo = var_mod.get()
            if self.modelo_ativo == "COPASA":
                self.variaveis_config = list(DEFAULT_VARS_COPASA)
            else:
                self.variaveis_config = list(DEFAULT_VARS_PADRAO)
            janela.destroy()
            self.dados_pesquisas = []
            self._exibir_tela_trabalho()
            if callback_pos_confirmacao:
                callback_pos_confirmacao()

        ttk.Button(janela, text="Avançar", style="Primary.TButton", command=confirmar).pack(pady=(25, 0))

    def _abrir_projeto_hub(self):
        if self._executar_abertura_json():
            self._exibir_tela_trabalho()

    def _importar_planilha_hub(self):
        self._popup_escolha_modelo(callback_pos_confirmacao=self._importar_planilha_excel)

    # --- TELA 2: ÁREA DE TRABALHO ---
    def _exibir_tela_trabalho(self):
        self._limpar_container()
        self._criar_menu_superior()
        self._criar_layout_operacional()
        self._atualizar_interface_variaveis()
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
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Sair", command=self.root.quit)
        menubar.add_cascade(label="Arquivo", menu=menu_arquivo)

        menu_config = tk.Menu(menubar, tearoff=0)
        menu_config.add_command(label="Gerenciar Variáveis da Avaliação...", command=self._janela_config_variaveis)
        menubar.add_cascade(label="Configurações", menu=menu_config)

        self.root.config(menu=menubar)

    def _criar_layout_operacional(self):
        mod_label = "Modelo 1: Padrão / Memória de Cálculo" if self.modelo_ativo == "PADRAO" else "Modelo 2: Contrato COPASA"
        lbl_info_mod = ttk.Label(self.container_principal, text=f"Modo Ativo: {mod_label}", font=("Segoe UI", 9, "italic"), foreground="#1f3c5b")
        lbl_info_mod.pack(anchor="w", padx=12, pady=(4, 0))

        frame_form = ttk.LabelFrame(self.container_principal, text=" Cadastro e Edição do Dado de Mercado ", padding=8)
        frame_form.pack(fill="x", padx=10, pady=4)

        ttk.Label(frame_form, text="Informante:").grid(row=0, column=0, sticky="w")
        self.txt_informante = ttk.Entry(frame_form, width=22)
        self.txt_informante.grid(row=0, column=1, padx=4, pady=2)

        ttk.Label(frame_form, text="Telefone:").grid(row=0, column=2, sticky="w")
        self.txt_telefone = ttk.Entry(frame_form, width=22)
        self.txt_telefone.grid(row=0, column=3, padx=4, pady=2)

        ttk.Label(frame_form, text="Logradouro/Endereço:").grid(row=1, column=0, sticky="w")
        self.txt_endereco = ttk.Entry(frame_form, width=22)
        self.txt_endereco.grid(row=1, column=1, padx=4, pady=2)

        ttk.Label(frame_form, text="Bairro:").grid(row=1, column=2, sticky="w")
        self.txt_bairro = ttk.Entry(frame_form, width=22)
        self.txt_bairro.grid(row=1, column=3, padx=4, pady=2)

        ttk.Label(frame_form, text="Município:").grid(row=2, column=0, sticky="w")
        self.txt_municipio = ttk.Entry(frame_form, width=22)
        self.txt_municipio.grid(row=2, column=1, padx=4, pady=2)

        ttk.Label(frame_form, text="Valor Total (R$):").grid(row=2, column=2, sticky="w")
        self.txt_valor = ttk.Entry(frame_form, width=22)
        self.txt_valor.grid(row=2, column=3, padx=4, pady=2)

        ttk.Label(frame_form, text="Área Terreno:").grid(row=3, column=0, sticky="w")
        frame_area = ttk.Frame(frame_form)
        frame_area.grid(row=3, column=1, sticky="w", padx=4, pady=2)
        self.txt_area = ttk.Entry(frame_area, width=13)
        self.txt_area.pack(side="left")
        self.var_unidade = tk.StringVar(value="ha" if self.modelo_ativo == "PADRAO" else "m²")
        self.cb_unidade = ttk.Combobox(frame_area, textvariable=self.var_unidade, values=["m²", "ha"], width=5, state="readonly")
        self.cb_unidade.pack(side="left", padx=2)

        ttk.Label(frame_form, text="Área Const. (m²):").grid(row=3, column=2, sticky="w")
        self.txt_area_const = ttk.Entry(frame_form, width=22)
        self.txt_area_const.grid(row=3, column=3, padx=4, pady=2)

        frame_coords = ttk.Frame(frame_form)
        frame_coords.grid(row=4, column=0, columnspan=4, sticky="w", pady=3)

        ttk.Label(frame_coords, text="Zona UTM:", foreground="#b22222", font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 2))
        self.txt_zona = ttk.Entry(frame_coords, width=7)
        self.txt_zona.pack(side="left", padx=(0, 10))

        ttk.Label(frame_coords, text="Coord. E (m):").pack(side="left", padx=(0, 2))
        self.txt_coord_e = ttk.Entry(frame_coords, width=16)
        self.txt_coord_e.pack(side="left", padx=(0, 10))

        ttk.Label(frame_coords, text="Coord. S (m):").pack(side="left", padx=(0, 2))
        self.txt_coord_s = ttk.Entry(frame_coords, width=16)
        self.txt_coord_s.pack(side="left", padx=(0, 10))

        ttk.Label(frame_coords, text="Data:").pack(side="left", padx=(0, 2))
        self.txt_data = ttk.Entry(frame_coords, width=12)
        self.txt_data.insert(0, "18/09/2026")
        self.txt_data.pack(side="left")

        ttk.Label(frame_form, text="Link do Anúncio:").grid(row=5, column=0, sticky="w")
        self.txt_link = ttk.Entry(frame_form, width=65)
        self.txt_link.grid(row=5, column=1, columnspan=3, sticky="w", padx=4, pady=2)

        ttk.Label(frame_form, text="Foto 1 (Imóvel/Drone):").grid(row=6, column=0, sticky="w")
        self.txt_foto1 = ttk.Entry(frame_form, width=52)
        self.txt_foto1.grid(row=6, column=1, columnspan=2, sticky="w", padx=4, pady=2)
        ttk.Button(frame_form, text="Buscar", command=lambda: self._buscar_arquivo_foto(self.txt_foto1)).grid(row=6, column=3, sticky="w")

        ttk.Label(frame_form, text="Foto 2 (Print Anúncio):").grid(row=7, column=0, sticky="w")
        self.txt_foto2 = ttk.Entry(frame_form, width=52)
        self.txt_foto2.grid(row=7, column=1, columnspan=2, sticky="w", padx=4, pady=2)
        ttk.Button(frame_form, text="Buscar", command=lambda: self._buscar_arquivo_foto(self.txt_foto2)).grid(row=7, column=3, sticky="w")

        # Variáveis Dinâmicas
        self.frame_vars = ttk.LabelFrame(self.container_principal, text=" Variáveis da Avaliação ", padding=8)
        self.frame_vars.pack(fill="x", padx=10, pady=3)
        self.widgets_dinamicos = {}

        frame_btn_cad = ttk.Frame(self.container_principal, padding=4)
        frame_btn_cad.pack(fill="x", padx=10)

        self.btn_salvar_dado = ttk.Button(frame_btn_cad, text="➕ Adicionar à Lista", style="Primary.TButton", command=self._adicionar_ou_salvar_dado)
        self.btn_salvar_dado.pack(side="left", padx=4)

        ttk.Button(frame_btn_cad, text="🧹 Novo / Limpar Campos", command=self._limpar_formulario).pack(side="left", padx=4)

        self.btn_cancelar_edicao = ttk.Button(frame_btn_cad, text="✖ Cancelar Edição", command=self._limpar_formulario, state="disabled")
        self.btn_cancelar_edicao.pack(side="left", padx=4)

        ttk.Button(frame_btn_cad, text="⚙ Configurar Variáveis", command=self._janela_config_variaveis).pack(side="right", padx=4)

        # Tabela
        frame_tabela = ttk.LabelFrame(self.container_principal, text=" Dados Cadastrados (Clique duplo para editar) ", padding=8)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=3)

        colunas = ("dado", "informante", "endereco", "municipio", "valor", "area", "unidade", "unitario")
        self.tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=6)
        self.tree.heading("dado", text="D.")
        self.tree.heading("informante", text="Informante")
        self.tree.heading("endereco", text="Endereço")
        self.tree.heading("municipio", text="Município")
        self.tree.heading("valor", text="Valor Total (R$)")
        self.tree.heading("area", text="Área")
        self.tree.heading("unidade", text="Unid.")
        self.tree.heading("unitario", text="Unitário (R$/un)")

        self.tree.column("dado", width=35, anchor="center")
        self.tree.column("unidade", width=50, anchor="center")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>", lambda event: self._carregar_para_edicao())

        frame_botoes_grid = ttk.Frame(frame_tabela)
        frame_botoes_grid.pack(fill="x", pady=3)
        ttk.Button(frame_botoes_grid, text="✏ Editar Selecionado", command=self._carregar_para_edicao).pack(side="left", padx=4)
        ttk.Button(frame_botoes_grid, text="🗑 Excluir Selecionado", command=self._excluir_dado).pack(side="left", padx=4)
        ttk.Button(frame_botoes_grid, text="📥 Importar Planilha (.xlsx)", command=self._importar_planilha_excel).pack(side="right", padx=4)

        self.progress_bar = ttk.Progressbar(self.container_principal, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=12, pady=2)
        self.progress_bar.pack_forget()

        frame_acoes = ttk.Frame(self.container_principal, padding=6)
        frame_acoes.pack(fill="x", padx=10, pady=4)

        ttk.Button(frame_acoes, text="📊 Exportar Planilha Excel (SISDEA)", style="Accent.TButton", command=self._exportar_excel).pack(side="left", padx=6, expand=True, fill="x")
        ttk.Button(frame_acoes, text="📄 Gerar Relatório Word (Fichas Prontas)", style="Primary.TButton", command=self._iniciar_exportacao_word_thread).pack(side="right", padx=6, expand=True, fill="x")

    def _atualizar_interface_variaveis(self):
        for w in self.frame_vars.winfo_children():
            w.destroy()
        self.widgets_dinamicos.clear()

        for idx, var in enumerate(self.variaveis_config):
            col = (idx % 2) * 2
            row = idx // 2

            lbl = ttk.Label(self.frame_vars, text=f"{var['nome']}:")
            lbl.grid(row=row, column=col, sticky="w", padx=4, pady=2)

            if var["tipo"] == "codigo" and var.get("opcoes"):
                cb = ttk.Combobox(self.frame_vars, values=var["opcoes"], width=24)
                if var["opcoes"]:
                    cb.set(var["opcoes"][0])
                cb.grid(row=row, column=col + 1, sticky="w", padx=4, pady=2)
                self.widgets_dinamicos[var["nome"]] = cb
            else:
                ent = ttk.Entry(self.frame_vars, width=26)
                ent.grid(row=row, column=col + 1, sticky="w", padx=4, pady=2)
                self.widgets_dinamicos[var["nome"]] = ent

    def _janela_config_variaveis(self):
        janela = tk.Toplevel(self.root)
        janela.title("Gerenciador de Variáveis")
        janela.geometry("640x460")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Configure as variáveis do modelo ativo:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=6)

        frame_lista = ttk.Frame(janela, padding=6)
        frame_lista.pack(fill="both", expand=True)

        entradas_vars = []
        total_slots = max(6, len(self.variaveis_config) + 2)

        for i in range(total_slots):
            f_linha = ttk.Frame(frame_lista)
            f_linha.pack(fill="x", pady=2)

            ativa = i < len(self.variaveis_config)
            cfg = self.variaveis_config[i] if ativa else {"nome": "", "tipo": "texto", "opcoes": []}

            var_ativa = tk.BooleanVar(value=ativa)
            chk = ttk.Checkbutton(f_linha, text=f"Var {i+1}", variable=var_ativa)
            chk.grid(row=0, column=0, padx=2)

            e_nome = ttk.Entry(f_linha, width=16)
            e_nome.insert(0, cfg.get("nome", ""))
            e_nome.grid(row=0, column=1, padx=3)

            cb_tipo = ttk.Combobox(f_linha, values=["codigo", "numero", "texto"], width=8, state="readonly")
            cb_tipo.set(cfg.get("tipo", "texto"))
            cb_tipo.grid(row=0, column=2, padx=3)

            e_opcoes = ttk.Entry(f_linha, width=28)
            e_opcoes.insert(0, ", ".join(cfg.get("opcoes", [])))
            e_opcoes.grid(row=0, column=3, padx=3)

            entradas_vars.append((var_ativa, e_nome, cb_tipo, e_opcoes))

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
            janela.destroy()
            messagebox.showinfo("Sucesso", "Variáveis atualizadas com sucesso!")

        ttk.Button(janela, text="💾 Salvar Configurações", command=salvar).pack(pady=8)

    def _buscar_arquivo_foto(self, entry_widget):
        caminho = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg;*.webp")])
        if caminho:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, caminho)

    def _limpar_formulario(self):
        self.item_em_edicao = None
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
        self.txt_data.insert(0, "18/09/2026")
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

        self.btn_salvar_dado.config(text="➕ Adicionar à Lista")
        self.btn_cancelar_edicao.config(state="disabled")

    def _adicionar_ou_salvar_dado(self):
        try:
            valor_total = converter_para_float(self.txt_valor.get())
            area_num = converter_para_float(self.txt_area.get())
            area_const = converter_para_float(self.txt_area_const.get())
            unidade = self.var_unidade.get()
            unitario = valor_total / area_num if area_num > 0 else 0.0

            dado_num = self.item_em_edicao.get("dado_id", self.item_em_edicao.get("D.", 1)) if self.item_em_edicao else (len(self.dados_pesquisas) + 1)

            coord_e = limpar_sufixo_coord(self.txt_coord_e.get())
            coord_s = limpar_sufixo_coord(self.txt_coord_s.get())

            registro = {
                "dado_id": dado_num,
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
                    messagebox.showinfo("Atualizado", f"Pesquisa {dado_num:02d} atualizada com sucesso!")
            else:
                self.dados_pesquisas.append(registro)

            self._recarregar_grid()
            self._limpar_formulario()
        except Exception as e:
            messagebox.showerror("Erro de Preenchimento", f"Verifique os dados numéricos: {e}")

    def _recarregar_grid(self):
        if not hasattr(self, 'tree'):
            return
        for item in self.tree.get_children():
            self.tree.delete(item)

        for dado in self.dados_pesquisas:
            d_id = dado.get("dado_id") or dado.get("D.", 1)
            un = dado.get("unidade") or dado.get("Unidade", "m²")
            casas = 4 if un == "ha" else 2
            v_total = dado.get("valor_total") if dado.get("valor_total") is not None else dado.get("Valor Total (R$)", 0.0)
            a_total = dado.get("area_terreno") if dado.get("area_terreno") is not None else (dado.get(f"Área Terreno ({un})") or dado.get("Área Terreno (ha)") or dado.get("Área Terreno (m²)", 0.0))
            u_unit = dado.get("unitario") if dado.get("unitario") is not None else (dado.get(f"Unitário (R$/{un})") or 0.0)

            self.tree.insert("", "end", values=(
                d_id,
                dado.get("informante") or dado.get("Informante", ""),
                dado.get("endereco") or dado.get("Endereço", ""),
                dado.get("municipio") or dado.get("Município", ""),
                f"R$ {formatar_moeda_br(v_total)}",
                f"{formatar_numero_br(a_total, casas)}",
                un,
                f"R$ {formatar_moeda_br(u_unit)}"
            ))

    def _carregar_para_edicao(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma pesquisa na tabela para editar.")
            return

        item_val = self.tree.item(sel[0])["values"]
        dado_id = item_val[0]
        dado = next((d for d in self.dados_pesquisas if (d.get("dado_id") or d.get("D.")) == dado_id), None)
        if not dado:
            return

        self.item_em_edicao = dado
        self.btn_salvar_dado.config(text=f"💾 Atualizar Dado {dado_id}")
        self.btn_cancelar_edicao.config(state="normal")

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

        a_total = dado.get("area_terreno") if dado.get("area_terreno") is not None else (dado.get(f"Área Terreno ({un})") or dado.get("Área Terreno (ha)") or dado.get("Área Terreno (m²)", 0.0))
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
            val = extras.get(nome, dado.get(nome, ""))
            if isinstance(widget, ttk.Combobox):
                widget.set(val)
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(val))

    def _excluir_dado(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma pesquisa para excluir.")
            return

        item_val = self.tree.item(sel[0])["values"]
        dado_id = item_val[0]
        if messagebox.askyesno("Confirmar Exclusão", f"Deseja excluir a Pesquisa {dado_id}?"):
            self.dados_pesquisas = [d for d in self.dados_pesquisas if (d.get("dado_id") or d.get("D.")) != dado_id]
            for i, d in enumerate(self.dados_pesquisas):
                d["dado_id"] = i + 1
            self._recarregar_grid()
            self._limpar_formulario()

    def _salvar_projeto(self):
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Projeto de Pesquisa (*.json)", "*.json")])
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
        caminho = filedialog.askopenfilename(filetypes=[("Projeto de Pesquisa (*.json)", "*.json")])
        if not caminho:
            return False
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                conteudo = json.load(f)
            self.modelo_ativo = conteudo.get("modelo", "PADRAO")
            self.variaveis_config = conteudo.get("variaveis_config") or conteudo.get("config_variaveis") or list(DEFAULT_VARS_PADRAO)
            
            raw_pesquisas = conteudo.get("pesquisas", [])
            self.dados_pesquisas = [normalizar_dado(d) for d in raw_pesquisas]

            if hasattr(self, 'tree'):
                self._atualizar_interface_variaveis()
                self._recarregar_grid()
                self._limpar_formulario()
                messagebox.showinfo("Sucesso", f"{len(self.dados_pesquisas)} pesquisas carregadas com sucesso!")

            return True
        except Exception as e:
            messagebox.showerror("Erro ao Abrir", f"Arquivo corrompido ou formato inválido: {e}")
            return False

    def _abrir_projeto_menu(self):
        self._executar_abertura_json()

    # --- IMPORTADOR INTELIGENTE DE PLANILHAS EXCEL ---
    def _importar_planilha_excel(self):
        caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx;*.xls")])
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

            mapeamento = {}
            unidade_detectada_coluna = None

            for col in df.columns:
                c_clean = str(col).strip().lower()
                if "informante" in c_clean: mapeamento["informante"] = col
                elif "telefone" in c_clean or "contato" in c_clean: mapeamento["telefone"] = col
                elif "endereço" in c_clean or "endereco" in c_clean or "logradouro" in c_clean: mapeamento["endereco"] = col
                elif "bairro" in c_clean: mapeamento["bairro"] = col
                elif "município" in c_clean or "municipio" in c_clean: mapeamento["municipio"] = col
                elif "valor total" in c_clean or "oferta" in c_clean: mapeamento["valor_total"] = col
                elif "área terreno" in c_clean or "area terreno" in c_clean:
                    mapeamento["area_terreno"] = col
                    if "(ha)" in c_clean:
                        unidade_detectada_coluna = "ha"
                    elif "(m²)" in c_clean or "(m2)" in c_clean:
                        unidade_detectada_coluna = "m²"
                elif "área construída" in c_clean or "area construida" in c_clean: mapeamento["area_construida"] = col
                elif "coord. e" in c_clean or "coord e" in c_clean or "e (m)" in c_clean: mapeamento["coord_e"] = col
                elif "coord. s" in c_clean or "coord s" in c_clean or "s (m)" in c_clean: mapeamento["coord_s"] = col
                elif "zona" in c_clean: mapeamento["zona_utm"] = col
                elif "data" in c_clean: mapeamento["data"] = col
                elif "link" in c_clean: mapeamento["link"] = col
                elif "unidade" in c_clean or "unid" in c_clean: mapeamento["unidade"] = col

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
                    unidade = self.var_unidade.get()

                u_calc = v_total / a_terr if a_terr > 0 else 0.0
                item_id = len(self.dados_pesquisas) + len(novos_dados) + 1

                registro = {
                    "dado_id": item_id,
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
                    "data": str(row.get(mapeamento.get("data", "18/09/2026"))).replace("nan", "").strip(),
                    "link": str(row.get(mapeamento.get("link", ""), "")).replace("nan", "").strip(),
                    "foto1": "",
                    "foto2": "",
                    "variaveis_extras": {}
                }

                for v_cfg in self.variaveis_config:
                    for col_planilha in df.columns:
                        if v_cfg["nome"].lower() in str(col_planilha).lower():
                            registro["variaveis_extras"][v_cfg["nome"]] = str(row.get(col_planilha, "")).replace("nan", "").strip()

                novos_dados.append(registro)

            if novos_dados:
                self.dados_pesquisas.extend(novos_dados)
                self._recarregar_grid()
                messagebox.showinfo("Importação Concluída", f"{len(novos_dados)} pesquisas importadas com sucesso!\nConfira as unidades e anexe as fotos correspondentes.")
            else:
                messagebox.showwarning("Aviso", "Nenhum dado válido encontrado para importação.")

        except Exception as e:
            messagebox.showerror("Erro na Importação", f"Falha ao ler a planilha: {e}")

    # --- DOWNLOADS E EXPORTAÇÃO ---
    def _baixar_imagem(self, caminho_ou_url):
        if not caminho_ou_url:
            return None
        caminho_ou_url = normalizar_url_gdrive(caminho_ou_url)
        try:
            if caminho_ou_url.startswith("http://") or caminho_ou_url.startswith("https://"):
                resp = requests.get(caminho_ou_url, timeout=12)
                if resp.status_code == 200:
                    return io.BytesIO(resp.content)
            elif os.path.exists(caminho_ou_url):
                return caminho_ou_url
        except Exception as e:
            logging.warning(f"Erro ao carregar imagem '{caminho_ou_url}': {e}")
            return None
        return None

    def _exportar_excel(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportar.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilha Excel (*.xlsx)", "*.xlsx")])
        if not caminho:
            return

        linhas_export = []
        for d in self.dados_pesquisas:
            d_id = d.get("dado_id") or d.get("D.", 1)
            un = d.get("unidade") or d.get("Unidade", "m²")
            v_total = d.get("valor_total") if d.get("valor_total") is not None else d.get("Valor Total (R$)", 0.0)
            a_total = d.get("area_terreno") if d.get("area_terreno") is not None else (d.get(f"Área Terreno ({un})") or 0.0)
            a_const = d.get("area_construida") if d.get("area_construida") is not None else d.get("Área Construída (m²)", 0.0)
            u_unit = d.get("unitario") if d.get("unitario") is not None else (d.get(f"Unitário (R$/{un})") or 0.0)

            linha = {
                "D.": d_id,
                "Informante": d.get("informante") or d.get("Informante", ""),
                "Telefone": d.get("telefone") or d.get("Telefone", ""),
                "Endereço": d.get("endereco") or d.get("Endereço", ""),
                "Bairro": d.get("bairro") or d.get("Bairro", ""),
                "Município": d.get("municipio") or d.get("Município", ""),
                "Valor Total (R$)": v_total,
                f"Área Terreno ({un})": a_total,
                "Área Construída (m²)": a_const,
                f"Unitário (R$/{un})": u_unit,
                "Zona UTM": d.get("zona_utm") or d.get("Zona UTM", ""),
                "Coord. E (m)": d.get("coord_e") or d.get("Coord. E (m)", ""),
                "Coord. S (m)": d.get("coord_s") or d.get("Coord. S (m)", ""),
                "Localização": d.get("localizacao") or d.get("Localização", "Urbana"),
                "Data": d.get("data") or d.get("Data", ""),
                "Link": d.get("link") or d.get("Link", "")
            }
            extras = d.get("variaveis_extras") or d.get("VariaveisExtras", {})
            for k, v in extras.items():
                linha[k] = v
            linhas_export.append(linha)

        df = pd.DataFrame(linhas_export)
        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Ficha de Pesquisa", index=False)

        messagebox.showinfo("Sucesso", "Planilha exportada com sucesso!")

    def _definir_bordas_tabela(self, table):
        """Aplica moldura externa retangular e separador horizontal, SEM a listra vertical do meio."""
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            f'<w:insideV w:val="none"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def _iniciar_exportacao_word_thread(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado para exportação.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Documento Word (*.docx)", "*.docx")])
        if not caminho:
            return

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
            for section in doc.sections:
                section.top_margin = Inches(0.35)
                section.bottom_margin = Inches(0.35)
                section.left_margin = Inches(0.45)
                section.right_margin = Inches(0.45)

            total_dados = len(dados_lista)

            for i in range(0, total_dados, 2):
                if i > 0:
                    doc.add_page_break()

                # CABEÇALHOS DE PÁGINA
                if modelo_selecionado == "PADRAO":
                    # Modelo 1: Título e barra azul apenas na Página 1
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
                            f'<w:pBrd {nsdecls("w")}>'
                            f'<w:bottom w:val="single" w:sz="12" w:space="1" w:color="245D8C"/>'
                            f'</w:pBrd>'
                        )
                        p_bar._p.get_or_add_pPr().append(pBrd)
                    else:
                        p_sp = doc.add_paragraph()
                        p_sp.paragraph_format.space_before = Pt(0)
                        p_sp.paragraph_format.space_after = Pt(4)

                elif modelo_selecionado == "COPASA":
                    # Modelo 2: Logos Enprol e Copasa alinhados
                    t_cab = doc.add_table(rows=1, cols=2)
                    t_cab.alignment = WD_TABLE_ALIGNMENT.CENTER
                    t_cab.autofit = False
                    c_logo1, c_logo2 = t_cab.cell(0, 0), t_cab.cell(0, 1)
                    c_logo1.width = Inches(3.7)
                    c_logo2.width = Inches(3.7)

                    p_enp = c_logo1.paragraphs[0]
                    p_enp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    if os.path.exists("logo_enprol.png"):
                        try:
                            p_enp.add_run().add_picture("logo_enprol.png", height=Inches(0.40))
                        except Exception:
                            r = p_enp.add_run("ENPROL")
                            r.bold = True
                            r.font.name = "Arial"
                            r.font.size = Pt(13)
                            r.font.color.rgb = RGBColor(28, 93, 153)
                    else:
                        r = p_enp.add_run("ENPROL")
                        r.bold = True
                        r.font.name = "Arial"
                        r.font.size = Pt(13)
                        r.font.color.rgb = RGBColor(28, 93, 153)

                    p_cop = c_logo2.paragraphs[0]
                    p_cop.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    if os.path.exists("logo_copasa.png"):
                        try:
                            p_cop.add_run().add_picture("logo_copasa.png", height=Inches(0.40))
                        except Exception:
                            r2 = p_cop.add_run("copasa")
                            r2.bold = True
                            r2.font.name = "Arial"
                            r2.font.size = Pt(13)
                            r2.font.color.rgb = RGBColor(50, 110, 180)
                    else:
                        r2 = p_cop.add_run("copasa")
                        r2.bold = True
                        r2.font.name = "Arial"
                        r2.font.size = Pt(13)
                        r2.font.color.rgb = RGBColor(50, 110, 180)

                    # Faixa roxa/azul: Tabela de 1 linha e 2 colunas com alinhamento perfeito
                    t_faixa = doc.add_table(rows=1, cols=2)
                    t_faixa.alignment = WD_TABLE_ALIGNMENT.CENTER
                    t_faixa.autofit = False
                    cf_esq, cf_dir = t_faixa.cell(0, 0), t_faixa.cell(0, 1)
                    cf_esq.width = Inches(3.7)
                    cf_dir.width = Inches(3.7)

                    shd_esq = parse_xml(f'<w:shd {nsdecls("w")} w:fill="8FA8D6"/>')
                    shd_dir = parse_xml(f'<w:shd {nsdecls("w")} w:fill="8FA8D6"/>')
                    cf_esq._tc.get_or_add_tcPr().append(shd_esq)
                    cf_dir._tc.get_or_add_tcPr().append(shd_dir)

                    p_fe = cf_esq.paragraphs[0]
                    p_fe.paragraph_format.space_before = Pt(2)
                    p_fe.paragraph_format.space_after = Pt(2)
                    r_fe = p_fe.add_run("  CONTRATO: COPASA |")
                    r_fe.bold = True
                    r_fe.font.name = "Arial"
                    r_fe.font.size = Pt(9)
                    r_fe.font.color.rgb = RGBColor(255, 255, 255)

                    p_fd = cf_dir.paragraphs[0]
                    p_fd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    p_fd.paragraph_format.space_before = Pt(2)
                    p_fd.paragraph_format.space_after = Pt(2)
                    r_fd = p_fd.add_run("PESQUISA DE MERCADO  ")
                    r_fd.bold = True
                    r_fd.font.name = "Arial"
                    r_fd.font.size = Pt(9)
                    r_fd.font.color.rgb = RGBColor(255, 255, 255)

                    p_espaco = doc.add_paragraph()
                    p_espaco.paragraph_format.space_before = Pt(0)
                    p_espaco.paragraph_format.space_after = Pt(2)

                # TABELA DE PESQUISAS (Moldura retangular sem divisória vertical)
                tabela = doc.add_table(rows=0, cols=2)
                tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
                tabela.autofit = False
                self._definir_bordas_tabela(tabela)

                lote = dados_lista[i:i+2]
                for dado in lote:
                    d_id = dado.get("dado_id") or dado.get("D.", 1)
                    un = dado.get("unidade") or dado.get("Unidade", "m²")
                    row = tabela.add_row()
                    celula_dados, celula_fotos = row.cells[0], row.cells[1]
                    celula_dados.width = Inches(3.7)
                    celula_fotos.width = Inches(3.7)
                    celula_dados.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    celula_fotos.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    p_dados = celula_dados.paragraphs[0]
                    p_dados.paragraph_format.line_spacing = 1.10
                    p_dados.paragraph_format.space_before = Pt(0)
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
                        r_top = p_dados.add_run(f"Pesquisa – {d_id:02d}\n\n")
                        r_top.bold = True
                        r_top.font.name = "Arial"
                        r_top.font.size = Pt(10)

                        p_corpo = celula_dados.add_paragraph()
                        p_corpo.paragraph_format.line_spacing = 1.10
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

                        add_f_line(p_corpo, "Coordenadas Geográfica:", coord_texto)
                        loc_val = dado.get("localizacao") or dado.get("Localização", "Urbana")
                        add_f_line(p_corpo, "Localização:", loc_val)
                        p_corpo.add_run("\n")
                        add_f_line(p_corpo, "Data:", dado.get("data") or dado.get("Data", ""))

                    else:
                        # MODELO 2: COPASA
                        add_f_line(p_dados, "Logradouro:", dado.get("endereco") or dado.get("Endereço", ""))
                        add_f_line(p_dados, "Bairro:", dado.get("bairro") or dado.get("Bairro", ""))
                        add_f_line(p_dados, "Município:", dado.get("municipio") or dado.get("Município", ""))
                        p_dados.add_run("\n")

                        extras = dado.get("variaveis_extras") or dado.get("VariaveisExtras", {})
                        add_f_line(p_dados, "Informações:", extras.get("Informações", dado.get("Informações", "")))
                        p_dados.add_run("\n")

                        a_terr = dado.get("area_terreno") if dado.get("area_terreno") is not None else (dado.get("Área Terreno (m²)") or 0.0)
                        add_f_line(p_dados, "Área Terreno:", f"{formatar_numero_br(a_terr, 2)} m²")

                        a_const = dado.get("area_construida") if dado.get("area_construida") is not None else dado.get("Área Construída (m²)", 0.0)
                        add_f_line(p_dados, "Área Construída:", f"{formatar_numero_br(a_const, 2)} m²")

                        add_f_line(p_dados, "Frente:", extras.get("Frente", dado.get("Frente", "Não informado")))
                        p_dados.add_run("\n")
                        add_f_line(p_dados, "Via de acesso:", extras.get("Via de acesso", dado.get("Via de acesso", "")))

                        zona_val = dado.get("zona_utm") or dado.get("Zona UTM", "")
                        zona_str = f"{str(zona_val).strip()} " if zona_val else ""
                        coord_e = limpar_sufixo_coord(dado.get("coord_e") or dado.get("Coord. E (m)", ""))
                        coord_s = limpar_sufixo_coord(dado.get("coord_s") or dado.get("Coord. S (m)", ""))
                        coord_texto = f"{zona_str}{coord_e} m E / {coord_s} m S" if coord_e or coord_s else ""
                        add_f_line(p_dados, "Coordenadas UTM:", coord_texto)
                        p_dados.add_run("\n")

                        u_val = dado.get("unitario") if dado.get("unitario") is not None else (dado.get("Unitário (R$/m²)") or 0.0)
                        add_f_line(p_dados, "Valor Unitário:", f"R$ {formatar_moeda_br(u_val)}/m²")

                        v_tot = dado.get("valor_total") if dado.get("valor_total") is not None else dado.get("Valor Total (R$)", 0.0)
                        add_f_line(p_dados, "Valor Total:", f"R$ {formatar_moeda_br(v_tot)}")
                        p_dados.add_run("\n")
                        add_f_line(p_dados, "Data:", dado.get("data") or dado.get("Data", ""))

                        # Pesquisa centralizada na parte inferior do card
                        p_cop_num = celula_dados.add_paragraph()
                        p_cop_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_cop_num.paragraph_format.space_before = Pt(6)
                        p_cop_num.paragraph_format.space_after = Pt(0)
                        r_c_num = p_cop_num.add_run(f"Pesquisa {d_id:02d}")
                        r_c_num.bold = True
                        r_c_num.font.name = "Arial"
                        r_c_num.font.size = Pt(10)

                    # Inserção das Fotos
                    img1 = self._baixar_imagem(dado.get("foto1") or dado.get("Foto1"))
                    img2 = self._baixar_imagem(dado.get("foto2") or dado.get("Foto2"))

                    p_foto = celula_fotos.paragraphs[0]
                    p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    if img1 and img2:
                        try:
                            p_foto.add_run().add_picture(img1, width=Inches(2.75))
                        except Exception:
                            p_foto.add_run("[ Erro na Foto 1 ]\n")

                        p_foto2 = celula_fotos.add_paragraph()
                        p_foto2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_foto2.paragraph_format.space_before = Pt(2)
                        try:
                            p_foto2.add_run().add_picture(img2, width=Inches(2.75))
                        except Exception:
                            p_foto2.add_run("[ Erro na Foto 2 ]")

                    elif img1 or img2:
                        img_unica = img1 if img1 else img2
                        try:
                            p_foto.add_run().add_picture(img_unica, width=Inches(2.75))
                        except Exception:
                            p_foto.add_run("[ Erro ao carregar imagem ]")
                    else:
                        r_vazio = p_foto.add_run("[ Sem fotos anexadas ]")
                        r_vazio.font.name = "Arial"
                        r_vazio.font.size = Pt(10)

                    progresso = int(((i + 1) / total_dados) * 100)
                    self.root.after(0, lambda p=progresso: self._atualizar_progresso(p))

            doc.save(caminho)
            self.root.after(0, self._concluir_exportacao)
        except Exception as e:
            self.root.after(0, lambda err=e: self._falha_exportacao(err))

    def _atualizar_progresso(self, valor):
        self.progress_bar["value"] = valor

    def _concluir_exportacao(self):
        self.progress_bar.pack_forget()
        messagebox.showinfo("Sucesso", "Fichas de Pesquisa no Word geradas com sucesso!")

    def _falha_exportacao(self, err):
        self.progress_bar.pack_forget()
        messagebox.showerror("Erro na Exportação", f"Ocorreu um erro ao gerar o documento Word: {err}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppPesquisaMercado(root)
    root.mainloop()
