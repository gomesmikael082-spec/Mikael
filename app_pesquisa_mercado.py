import os
import io
import re
import json
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

CONFIG_FILE = "config_variaveis.json"

DEFAULT_VARS = [
    {"nome": "Via", "tipo": "codigo", "opcoes": ["1 - Local", "2 - Coletora", "3 - Arterial"]},
    {"nome": "Uso", "tipo": "codigo", "opcoes": ["1 - Residencial", "2 - Comercial", "3 - Misto"]},
    {"nome": "Testada (m)", "tipo": "numero", "opcoes": []},
    {"nome": "PGV (R$)", "tipo": "numero", "opcoes": []}
]

class AppPesquisaMercado:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Pesquisa de Mercado - SISDEA / Word")
        self.root.geometry("1020x780")

        self.dados_pesquisas = []
        self.item_em_edicao = None
        self.variaveis_config = self._carregar_config_variaveis()

        self._criar_menu()
        self._criar_layout()
        self._atualizar_interface_variaveis()

    def _carregar_config_variaveis(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return list(DEFAULT_VARS)

    def _salvar_config_variaveis(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.variaveis_config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("Erro ao salvar configurações", str(e))

    def _criar_menu(self):
        menubar = tk.Menu(self.root)
        menu_arquivo = tk.Menu(menubar, tearoff=0)
        menu_arquivo.add_command(label="Novo Projeto", command=self._novo_projeto)
        menu_arquivo.add_command(label="Abrir Projeto...", command=self._abrir_projeto)
        menu_arquivo.add_command(label="Salvar Projeto", command=self._salvar_projeto)
        menu_arquivo.add_separator()
        menu_arquivo.add_command(label="Sair", command=self.root.quit)
        menubar.add_cascade(label="Arquivo", menu=menu_arquivo)

        menu_config = tk.Menu(menubar, tearoff=0)
        menu_config.add_command(label="Configurar Variáveis...", command=self._janela_config_variaveis)
        menubar.add_cascade(label="Configurações", menu=menu_config)

        self.root.config(menu=menubar)

    def _criar_layout(self):
        # Frame de Entrada de Dados Principais
        frame_form = ttk.LabelFrame(self.root, text=" Cadastro do Dado de Mercado ", padding=10)
        frame_form.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_form, text="Informante:").grid(row=0, column=0, sticky="w")
        self.txt_informante = ttk.Entry(frame_form, width=24)
        self.txt_informante.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Telefone:").grid(row=0, column=2, sticky="w")
        self.txt_telefone = ttk.Entry(frame_form, width=24)
        self.txt_telefone.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Endereço:").grid(row=1, column=0, sticky="w")
        self.txt_endereco = ttk.Entry(frame_form, width=24)
        self.txt_endereco.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Bairro:").grid(row=1, column=2, sticky="w")
        self.txt_bairro = ttk.Entry(frame_form, width=24)
        self.txt_bairro.grid(row=1, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Município:").grid(row=2, column=0, sticky="w")
        self.txt_municipio = ttk.Entry(frame_form, width=24)
        self.txt_municipio.grid(row=2, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Valor Total (R$):").grid(row=2, column=2, sticky="w")
        self.txt_valor = ttk.Entry(frame_form, width=24)
        self.txt_valor.grid(row=2, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Área Terreno:").grid(row=3, column=0, sticky="w")
        frame_area = ttk.Frame(frame_form)
        frame_area.grid(row=3, column=1, sticky="w", padx=5, pady=2)
        self.txt_area = ttk.Entry(frame_area, width=14)
        self.txt_area.pack(side="left")
        self.var_unidade = tk.StringVar(value="m²")
        self.cb_unidade = ttk.Combobox(frame_area, textvariable=self.var_unidade, values=["m²", "ha"], width=5, state="readonly")
        self.cb_unidade.pack(side="left", padx=2)

        ttk.Label(frame_form, text="Área Const. (m²):").grid(row=3, column=2, sticky="w")
        self.txt_area_const = ttk.Entry(frame_form, width=24)
        self.txt_area_const.insert(0, "0")
        self.txt_area_const.grid(row=3, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Coord. E (m):").grid(row=4, column=0, sticky="w")
        self.txt_coord_e = ttk.Entry(frame_form, width=24)
        self.txt_coord_e.grid(row=4, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Coord. S (m):").grid(row=4, column=2, sticky="w")
        self.txt_coord_s = ttk.Entry(frame_form, width=24)
        self.txt_coord_s.grid(row=4, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Data:").grid(row=5, column=0, sticky="w")
        self.txt_data = ttk.Entry(frame_form, width=24)
        self.txt_data.insert(0, "31/08/2026")
        self.txt_data.grid(row=5, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Link do Anúncio:").grid(row=5, column=2, sticky="w")
        self.txt_link = ttk.Entry(frame_form, width=24)
        self.txt_link.grid(row=5, column=3, padx=5, pady=2)

        # Campos de 2 Fotos
        ttk.Label(frame_form, text="Foto 1 (Imóvel/Satélite):").grid(row=6, column=0, sticky="w")
        self.txt_foto1 = ttk.Entry(frame_form, width=40)
        self.txt_foto1.grid(row=6, column=1, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Button(frame_form, text="Buscar", command=lambda: self._buscar_arquivo_foto(self.txt_foto1)).grid(row=6, column=3, sticky="w")

        ttk.Label(frame_form, text="Foto 2 (Print Anúncio):").grid(row=7, column=0, sticky="w")
        self.txt_foto2 = ttk.Entry(frame_form, width=40)
        self.txt_foto2.grid(row=7, column=1, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Button(frame_form, text="Buscar", command=lambda: self._buscar_arquivo_foto(self.txt_foto2)).grid(row=7, column=3, sticky="w")

        # Frame de Variáveis Dinâmicas
        self.frame_vars = ttk.LabelFrame(self.root, text=" Variáveis da Avaliação (Até 4) ", padding=10)
        self.frame_vars.pack(fill="x", padx=10, pady=5)
        self.widgets_dinamicos = {}

        # Botões de Ação do Cadastro
        frame_btn_cad = ttk.Frame(self.root, padding=5)
        frame_btn_cad.pack(fill="x", padx=10)

        self.btn_salvar_dado = ttk.Button(frame_btn_cad, text="➕ Adicionar Dado à Lista", command=self._adicionar_ou_salvar_dado)
        self.btn_salvar_dado.pack(side="left", padx=5)

        self.btn_cancelar_edicao = ttk.Button(frame_btn_cad, text="✖ Cancelar Edição", command=self._limpar_formulario, state="disabled")
        self.btn_cancelar_edicao.pack(side="left", padx=5)

        ttk.Button(frame_btn_cad, text="⚙ Gerenciar Variáveis", command=self._janela_config_variaveis).pack(side="right", padx=5)

        # Tabela Visualização
        frame_tabela = ttk.LabelFrame(self.root, text=" Dados Cadastrados (Clique duplo para editar) ", padding=10)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=5)

        colunas = ("dado", "informante", "endereco", "bairro", "valor", "area", "unidade", "unitario")
        self.tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=7)
        self.tree.heading("dado", text="D.")
        self.tree.heading("informante", text="Informante")
        self.tree.heading("endereco", text="Endereço")
        self.tree.heading("bairro", text="Bairro")
        self.tree.heading("valor", text="Valor Total (R$)")
        self.tree.heading("area", text="Área")
        self.tree.heading("unidade", text="Unid.")
        self.tree.heading("unitario", text="Unitário (R$/un)")

        self.tree.column("dado", width=40, anchor="center")
        self.tree.column("unidade", width=60, anchor="center")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>", lambda event: self._carregar_para_edicao())

        frame_botoes_grid = ttk.Frame(frame_tabela)
        frame_botoes_grid.pack(fill="x", pady=5)
        ttk.Button(frame_botoes_grid, text="✏ Editar Selecionado", command=self._carregar_para_edicao).pack(side="left", padx=5)
        ttk.Button(frame_botoes_grid, text="🗑 Excluir Selecionado", command=self._excluir_dado).pack(side="left", padx=5)

        # Frame Exportações
        frame_acoes = ttk.Frame(self.root, padding=10)
        frame_acoes.pack(fill="x", padx=10, pady=5)

        ttk.Button(frame_acoes, text="📊 Exportar para Excel (SISDEA)", command=self._exportar_excel).pack(side="left", padx=10, expand=True, fill="x")
        ttk.Button(frame_acoes, text="📄 Exportar para Word (Fichas Técnicas)", command=self._exportar_word).pack(side="right", padx=10, expand=True, fill="x")

    def _atualizar_interface_variaveis(self):
        for w in self.frame_vars.winfo_children():
            w.destroy()
        self.widgets_dinamicos.clear()

        for idx, var in enumerate(self.variaveis_config[:4]):
            col = (idx % 2) * 2
            row = idx // 2

            lbl = ttk.Label(self.frame_vars, text=f"{var['nome']}:")
            lbl.grid(row=row, column=col, sticky="w", padx=5, pady=3)

            if var["tipo"] == "codigo" and var.get("opcoes"):
                cb = ttk.Combobox(self.frame_vars, values=var["opcoes"], width=26)
                if var["opcoes"]:
                    cb.set(var["opcoes"][0])
                cb.grid(row=row, column=col + 1, sticky="w", padx=5, pady=3)
                self.widgets_dinamicos[var["nome"]] = cb
            else:
                ent = ttk.Entry(self.frame_vars, width=28)
                ent.grid(row=row, column=col + 1, sticky="w", padx=5, pady=3)
                self.widgets_dinamicos[var["nome"]] = ent

    def _janela_config_variaveis(self):
        janela = tk.Toplevel(self.root)
        janela.title("Gerenciador de Variáveis da Pesquisa")
        janela.geometry("620x420")
        janela.transient(self.root)
        janela.grab_set()

        ttk.Label(janela, text="Defina até 4 variáveis adicionais para o projeto e ficha técnica:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=8)

        frame_lista = ttk.Frame(janela, padding=10)
        frame_lista.pack(fill="both", expand=True)

        entradas_vars = []

        for i in range(4):
            f_linha = ttk.LabelFrame(frame_lista, text=f" Variável {i+1} ", padding=5)
            f_linha.pack(fill="x", pady=4)

            ativa = i < len(self.variaveis_config)
            cfg = self.variaveis_config[i] if ativa else {"nome": "", "tipo": "texto", "opcoes": []}

            var_ativa = tk.BooleanVar(value=ativa)
            chk = ttk.Checkbutton(f_linha, text="Ativar", variable=var_ativa)
            chk.grid(row=0, column=0, padx=4)

            ttk.Label(f_linha, text="Nome:").grid(row=0, column=1)
            e_nome = ttk.Entry(f_linha, width=18)
            e_nome.insert(0, cfg.get("nome", ""))
            e_nome.grid(row=0, column=2, padx=4)

            ttk.Label(f_linha, text="Tipo:").grid(row=0, column=3)
            cb_tipo = ttk.Combobox(f_linha, values=["codigo", "numero", "texto"], width=8, state="readonly")
            cb_tipo.set(cfg.get("tipo", "texto"))
            cb_tipo.grid(row=0, column=4, padx=4)

            ttk.Label(f_linha, text="Opções (sep. vírgula):").grid(row=0, column=5)
            e_opcoes = ttk.Entry(f_linha, width=24)
            e_opcoes.insert(0, ", ".join(cfg.get("opcoes", [])))
            e_opcoes.grid(row=0, column=6, padx=4)

            entradas_vars.append((var_ativa, e_nome, cb_tipo, e_opcoes))

        def salvar_configs():
            novas_configs = []
            for var_ativa, e_nome, cb_tipo, e_opcoes in entradas_vars:
                if var_ativa.get() and e_nome.get().strip():
                    opts = [op.strip() for op in e_opcoes.get().split(",") if op.strip()]
                    novas_configs.append({
                        "nome": e_nome.get().strip(),
                        "tipo": cb_tipo.get(),
                        "opcoes": opts
                    })
            self.variaveis_config = novas_configs
            self._salvar_config_variaveis()
            self._atualizar_interface_variaveis()
            janela.destroy()
            messagebox.showinfo("Configurações", "Variáveis atualizadas com sucesso!")

        ttk.Button(janela, text="💾 Salvar Configurações", command=salvar_configs).pack(pady=10)

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
        self.txt_area_const.insert(0, "0")
        self.txt_coord_e.delete(0, tk.END)
        self.txt_coord_s.delete(0, tk.END)
        self.txt_link.delete(0, tk.END)
        self.txt_foto1.delete(0, tk.END)
        self.txt_foto2.delete(0, tk.END)

        for nome, widget in self.widgets_dinamicos.items():
            if isinstance(widget, ttk.Combobox):
                vals = widget.cget("values")
                if vals:
                    widget.set(vals[0])
            else:
                widget.delete(0, tk.END)

        self.btn_salvar_dado.config(text="➕ Adicionar Dado à Lista")
        self.btn_cancelar_edicao.config(state="disabled")

    def _adicionar_ou_salvar_dado(self):
        try:
            valor_raw = self.txt_valor.get().replace("R$", "").replace(".", "").replace(",", ".").strip()
            area_raw = self.txt_area.get().replace(".", "").replace(",", ".").strip()
            const_raw = self.txt_area_const.get().replace(".", "").replace(",", ".").strip()

            valor_total = float(valor_raw)
            area_num = float(area_raw)
            area_const = float(const_raw) if const_raw else 0.0
            unidade = self.var_unidade.get()
            unitario = valor_total / area_num if area_num > 0 else 0

            dado_num = self.item_em_edicao["D."] if self.item_em_edicao else (len(self.dados_pesquisas) + 1)

            registro = {
                "D.": dado_num,
                "Informante": self.txt_informante.get(),
                "Telefone": self.txt_telefone.get(),
                "Endereço": self.txt_endereco.get(),
                "Bairro": self.txt_bairro.get(),
                "Município": self.txt_municipio.get(),
                "Valor Total (R$)": valor_total,
                f"Área Terreno ({unidade})": area_num,
                "Área Construída (m²)": area_const,
                f"Unitário (R$/{unidade})": unitario,
                "Localização": "Urbana",
                "Coord. E (m)": self.txt_coord_e.get(),
                "Coord. S (m)": self.txt_coord_s.get(),
                "Data": self.txt_data.get(),
                "Link": self.txt_link.get(),
                "Foto1": self.txt_foto1.get(),
                "Foto2": self.txt_foto2.get(),
                "Unidade": unidade,
                "VariaveisExtras": {}
            }

            for nome, widget in self.widgets_dinamicos.items():
                val = widget.get().strip()
                registro[nome] = val
                registro["VariaveisExtras"][nome] = val

            if self.item_em_edicao:
                idx = next(i for i, d in enumerate(self.dados_pesquisas) if d["D."] == dado_num)
                self.dados_pesquisas[idx] = registro
                messagebox.showinfo("Atualização", f"Dado {dado_num} atualizado com sucesso!")
            else:
                self.dados_pesquisas.append(registro)

            self._recarregar_grid()
            self._limpar_formulario()
        except Exception as e:
            messagebox.showerror("Erro de Preenchimento", f"Verifique os campos: {e}")

    def _recarregar_grid(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for dado in self.dados_pesquisas:
            un = dado.get("Unidade", "m²")
            self.tree.insert("", "end", values=(
                dado["D."], dado["Informante"], dado["Endereço"],
                dado["Bairro"], f"R$ {dado['Valor Total (R$)']:,.2f}",
                f"{dado.get(f'Área Terreno ({un})', 0):,.2f}",
                un, f"R$ {dado.get(f'Unitário (R$/{un})', 0):,.2f}"
            ))

    def _carregar_para_edicao(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um dado na tabela para editar.")
            return

        item_val = self.tree.item(sel[0])["values"]
        dado_id = item_val[0]
        dado = next((d for d in self.dados_pesquisas if d["D."] == dado_id), None)
        if not dado:
            return

        self.item_em_edicao = dado
        self.btn_salvar_dado.config(text=f"💾 Atualizar Dado {dado_id}")
        self.btn_cancelar_edicao.config(state="normal")

        self.txt_informante.delete(0, tk.END)
        self.txt_informante.insert(0, dado.get("Informante", ""))

        self.txt_telefone.delete(0, tk.END)
        self.txt_telefone.insert(0, dado.get("Telefone", ""))

        self.txt_endereco.delete(0, tk.END)
        self.txt_endereco.insert(0, dado.get("Endereço", ""))

        self.txt_bairro.delete(0, tk.END)
        self.txt_bairro.insert(0, dado.get("Bairro", ""))

        self.txt_municipio.delete(0, tk.END)
        self.txt_municipio.insert(0, dado.get("Município", ""))

        self.txt_valor.delete(0, tk.END)
        self.txt_valor.insert(0, str(dado.get("Valor Total (R$)", "")))

        un = dado.get("Unidade", "m²")
        self.var_unidade.set(un)
        self.txt_area.delete(0, tk.END)
        self.txt_area.insert(0, str(dado.get(f"Área Terreno ({un})", "")))

        self.txt_area_const.delete(0, tk.END)
        self.txt_area_const.insert(0, str(dado.get("Área Construída (m²)", 0)))

        self.txt_coord_e.delete(0, tk.END)
        self.txt_coord_e.insert(0, dado.get("Coord. E (m)", ""))

        self.txt_coord_s.delete(0, tk.END)
        self.txt_coord_s.insert(0, dado.get("Coord. S (m)", ""))

        self.txt_data.delete(0, tk.END)
        self.txt_data.insert(0, dado.get("Data", ""))

        self.txt_link.delete(0, tk.END)
        self.txt_link.insert(0, dado.get("Link", ""))

        self.txt_foto1.delete(0, tk.END)
        self.txt_foto1.insert(0, dado.get("Foto1", ""))

        self.txt_foto2.delete(0, tk.END)
        self.txt_foto2.insert(0, dado.get("Foto2", ""))

        extras = dado.get("VariaveisExtras", {})
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
            messagebox.showwarning("Aviso", "Selecione um dado na tabela para excluir.")
            return

        item_val = self.tree.item(sel[0])["values"]
        dado_id = item_val[0]
        if messagebox.askyesno("Confirmar Exclusão", f"Deseja realmente excluir o Dado {dado_id}?"):
            self.dados_pesquisas = [d for d in self.dados_pesquisas if d["D."] != dado_id]
            # Renumerar
            for i, d in enumerate(self.dados_pesquisas):
                d["D."] = i + 1
            self._recarregar_grid()
            self._limpar_formulario()

    def _novo_projeto(self):
        if messagebox.askyesno("Novo Projeto", "Deseja criar um novo projeto? Dados não salvos serão limpos."):
            self.dados_pesquisas = []
            self._recarregar_grid()
            self._limpar_formulario()

    def _salvar_projeto(self):
        caminho = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Projeto Pesquisa (*.json)", "*.json")])
        if not caminho:
            return
        dados_salvar = {
            "config_variaveis": self.variaveis_config,
            "pesquisas": self.dados_pesquisas
        }
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(dados_salvar, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("Sucesso", "Projeto salvo com sucesso!")

    def _abrir_projeto(self):
        caminho = filedialog.askopenfilename(filetypes=[("Projeto Pesquisa (*.json)", "*.json")])
        if not caminho:
            return
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                conteudo = json.load(f)
            self.variaveis_config = conteudo.get("config_variaveis", self.variaveis_config)
            self.dados_pesquisas = conteudo.get("pesquisas", [])
            self._atualizar_interface_variaveis()
            self._recarregar_grid()
            self._limpar_formulario()
            messagebox.showinfo("Sucesso", "Projeto carregado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao abrir", f"Não foi possível abrir o arquivo: {e}")

    def _baixar_imagem(self, caminho_ou_url):
        if not caminho_ou_url:
            return None
        try:
            if caminho_ou_url.startswith("http://") or caminho_ou_url.startswith("https://"):
                url = caminho_ou_url
                if "drive.google.com" in url and "id=" in url:
                    file_id = re.search(r"id=([a-zA-Z0-9_-]+)", url).group(1)
                    url = f"https://drive.google.com/uc?export=download&id={file_id}"
                resp = requests.get(url, timeout=10)
                if resp.status_code == 200:
                    return io.BytesIO(resp.content)
            elif os.path.exists(caminho_ou_url):
                return caminho_ou_url
        except Exception:
            return None
        return None

    def _exportar_excel(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not caminho:
            return

        df = pd.DataFrame(self.dados_pesquisas)
        df_sisdea = df.drop(columns=["Foto1", "Foto2", "Unidade", "VariaveisExtras"], errors="ignore")

        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df_sisdea.to_excel(writer, sheet_name="Ficha de Pesquisa", index=False)

        messagebox.showinfo("Concluído", "Planilha Excel para SISDEA exportada com sucesso!")

    def _definir_bordas_tabela(self, table):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def _exportar_word(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if not caminho:
            return

        doc = Document()
        for section in doc.sections:
            section.top_margin = Inches(0.4)
            section.bottom_margin = Inches(0.4)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        # Cabeçalho Superior Técnico
        p_topo = doc.add_paragraph()
        run_topo = p_topo.add_run("ENPROL")
        run_topo.bold = True
        run_topo.font.name = "Arial"
        run_topo.font.size = Pt(13)
        run_topo.font.color.rgb = RGBColor(90, 125, 154)

        p_linha = doc.add_paragraph()
        p_linha.paragraph_format.space_after = Pt(4)
        run_linha = p_linha.add_run("―" * 70)
        run_linha.font.color.rgb = RGBColor(90, 125, 154)

        tabela = doc.add_table(rows=0, cols=2)
        tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabela.autofit = False
        self._definir_bordas_tabela(tabela)

        for dado in self.dados_pesquisas:
            un = dado.get("Unidade", "m²")
            row = tabela.add_row()
            celula_dados, celula_fotos = row.cells[0], row.cells[1]
            celula_dados.width = Inches(3.7)
            celula_fotos.width = Inches(3.7)
            celula_dados.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            celula_fotos.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            # Montagem das Variáveis Extras da Pesquisa
            bloco_vars = ""
            for v_cfg in self.variaveis_config[:4]:
                v_nome = v_cfg["nome"]
                v_val = dado.get("VariaveisExtras", {}).get(v_nome, dado.get(v_nome, ""))
                bloco_vars += f"<b>{v_nome}:</b> {v_val}<br>"

            p_dados = celula_dados.paragraphs[0]
            p_dados.paragraph_format.line_spacing = 1.15
            p_dados.paragraph_format.space_after = Pt(0)

            def add_f_line(p, label, val):
                r1 = p.add_run(label)
                r1.bold = True
                r1.font.name = "Arial"
                r1.font.size = Pt(8.5)
                r2 = p.add_run(f" {val}\n")
                r2.font.name = "Arial"
                r2.font.size = Pt(8.5)

            add_f_line(p_dados, "Logradouro:", dado.get("Endereço", ""))
            add_f_line(p_dados, "Bairro:", dado.get("Bairro", ""))
            add_f_line(p_dados, "Município:", dado.get("Município", ""))
            add_f_line(p_dados, "Contato:", f"{dado.get('Telefone', '')} - {dado.get('Informante', '')}")
            add_f_line(p_dados, "Link:", dado.get("Link", ""))
            p_dados.add_run("\n")

            add_f_line(p_dados, "Área Terreno:", f"{dado.get(f'Área Terreno ({un})', 0):,.2f} {un}")
            add_f_line(p_dados, "Área Construída:", f"{dado.get('Área Construída (m²)', 0):,.2f} m²")
            add_f_line(p_dados, "Valor da Oferta:", f"R$ {dado.get('Valor Total (R$)', 0):,.2f}")
            add_f_line(p_dados, f"Valor Unitário/{un}:", f"R$ {dado.get(f'Unitário (R$/{un})', 0):,.2f}/{un}")
            p_dados.add_run("\n")

            for v_cfg in self.variaveis_config[:4]:
                v_nome = v_cfg["nome"]
                v_val = dado.get("VariaveisExtras", {}).get(v_nome, dado.get(v_nome, ""))
                add_f_line(p_dados, f"{v_nome}:", v_val)

            add_f_line(p_dados, "Coordenadas Geográfica:", f"{dado.get('Coord. E (m)', '')} / {dado.get('Coord. S (m)', '')}")
            add_f_line(p_dados, "Localização:", dado.get("Localização", "Urbana"))
            p_dados.add_run("\n")
            add_f_line(p_dados, "Data:", dado.get("Data", ""))

            p_num = celula_dados.add_paragraph()
            p_num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r_num = p_num.add_run(f"Pesquisa – {dado['D.']:02d}")
            r_num.bold = True
            r_num.font.name = "Arial"
            r_num.font.size = Pt(9)

            # Inserção das Fotos
            img1 = self._baixar_imagem(dado.get("Foto1"))
            img2 = self._baixar_imagem(dado.get("Foto2"))

            p_foto = celula_fotos.paragraphs[0]
            p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER

            if img1 and img2:
                try:
                    p_foto.add_run().add_picture(img1, width=Inches(3.3))
                except Exception:
                    p_foto.add_run("[ Erro na Foto 1 ]\n")
                
                p_foto2 = celula_fotos.add_paragraph()
                p_foto2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    p_foto2.add_run().add_picture(img2, width=Inches(3.3))
                except Exception:
                    p_foto2.add_run("[ Erro na Foto 2 ]")
            elif img1:
                try:
                    p_foto.add_run().add_picture(img1, width=Inches(3.4), height=Inches(3.2))
                except Exception:
                    p_foto.add_run("[ Erro na Foto 1 ]")
            elif img2:
                try:
                    p_foto.add_run().add_picture(img2, width=Inches(3.4), height=Inches(3.2))
                except Exception:
                    p_foto.add_run("[ Erro na Foto 2 ]")
            else:
                p_foto.add_run("[ Sem fotos anexadas ]")

        doc.save(caminho)
        messagebox.showinfo("Exportação Concluída", "Fichas de Pesquisa no Word geradas com sucesso!")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppPesquisaMercado(root)
    root.mainloop()
