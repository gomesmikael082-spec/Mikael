import os
import io
import re
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

class AppPesquisaMercado:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Pesquisa de Mercado - SISDEA / Word")
        self.root.geometry("900x700")

        self.dados_pesquisas = []
        self.variaveis_extras = ["Via", "Uso", "PGV (R$)", "Testada (m)"]

        self._criar_layout()

    def _criar_layout(self):
        # Frame de Entrada de Dados
        frame_form = ttk.LabelFrame(self.root, text=" Cadastro do Dado de Mercado ", padding=10)
        frame_form.pack(fill="x", padx=10, pady=5)

        # Campos Padrão
        ttk.Label(frame_form, text="Informante:").grid(row=0, column=0, sticky="w")
        self.txt_informante = ttk.Entry(frame_form, width=22)
        self.txt_informante.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Telefone:").grid(row=0, column=2, sticky="w")
        self.txt_telefone = ttk.Entry(frame_form, width=22)
        self.txt_telefone.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Endereço:").grid(row=1, column=0, sticky="w")
        self.txt_endereco = ttk.Entry(frame_form, width=22)
        self.txt_endereco.grid(row=1, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Bairro:").grid(row=1, column=2, sticky="w")
        self.txt_bairro = ttk.Entry(frame_form, width=22)
        self.txt_bairro.grid(row=1, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Município:").grid(row=2, column=0, sticky="w")
        self.txt_municipio = ttk.Entry(frame_form, width=22)
        self.txt_municipio.grid(row=2, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Valor Total (R$):").grid(row=2, column=2, sticky="w")
        self.txt_valor = ttk.Entry(frame_form, width=22)
        self.txt_valor.grid(row=2, column=3, padx=5, pady=2)

        # Unidade de Área e Valores
        ttk.Label(frame_form, text="Área Terreno:").grid(row=3, column=0, sticky="w")
        self.txt_area = ttk.Entry(frame_form, width=15)
        self.txt_area.grid(row=3, column=1, sticky="w", padx=5, pady=2)

        self.var_unidade = tk.StringVar(value="m²")
        cb_unidade = ttk.Combobox(frame_form, textvariable=self.var_unidade, values=["m²", "ha"], width=5, state="readonly")
        cb_unidade.grid(row=3, column=1, sticky="e", padx=5, pady=2)

        ttk.Label(frame_form, text="Área Const. (m²):").grid(row=3, column=2, sticky="w")
        self.txt_area_const = ttk.Entry(frame_form, width=22)
        self.txt_area_const.insert(0, "0")
        self.txt_area_const.grid(row=3, column=3, padx=5, pady=2)

        # Coordenadas e Link
        ttk.Label(frame_form, text="Coord. E (m):").grid(row=4, column=0, sticky="w")
        self.txt_coord_e = ttk.Entry(frame_form, width=22)
        self.txt_coord_e.grid(row=4, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Coord. S (m):").grid(row=4, column=2, sticky="w")
        self.txt_coord_s = ttk.Entry(frame_form, width=22)
        self.txt_coord_s.grid(row=4, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Data:").grid(row=5, column=0, sticky="w")
        self.txt_data = ttk.Entry(frame_form, width=22)
        self.txt_data.insert(0, "01/09/2026")
        self.txt_data.grid(row=5, column=1, padx=5, pady=2)

        ttk.Label(frame_form, text="Link do Anúncio:").grid(row=5, column=2, sticky="w")
        self.txt_link = ttk.Entry(frame_form, width=22)
        self.txt_link.grid(row=5, column=3, padx=5, pady=2)

        ttk.Label(frame_form, text="Link/Caminho da Imagem:").grid(row=6, column=0, sticky="w")
        self.txt_imagem = ttk.Entry(frame_form, width=45)
        self.txt_imagem.grid(row=6, column=1, columnspan=2, sticky="w", padx=5, pady=2)

        btn_arquivo_foto = ttk.Button(frame_form, text="Buscar Foto", command=self._selecionar_foto_local)
        btn_arquivo_foto.grid(row=6, column=3, sticky="w", padx=5, pady=2)

        # Frame de Variáveis Dinâmicas
        frame_vars = ttk.LabelFrame(self.root, text=" Variáveis da Avaliação (Editáveis) ", padding=10)
        frame_vars.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_vars, text="Via:").grid(row=0, column=0, sticky="w")
        self.txt_via = ttk.Entry(frame_vars, width=15)
        self.txt_via.insert(0, "1 - Local")
        self.txt_via.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(frame_vars, text="Uso:").grid(row=0, column=2, sticky="w")
        self.txt_uso = ttk.Entry(frame_vars, width=15)
        self.txt_uso.insert(0, "1 - Residencial")
        self.txt_uso.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(frame_vars, text="PGV (R$):").grid(row=0, column=4, sticky="w")
        self.txt_pgv = ttk.Entry(frame_vars, width=15)
        self.txt_pgv.insert(0, "400")
        self.txt_pgv.grid(row=0, column=5, padx=5, pady=2)

        ttk.Label(frame_vars, text="Testada (m):").grid(row=1, column=0, sticky="w")
        self.txt_testada = ttk.Entry(frame_vars, width=15)
        self.txt_testada.grid(row=1, column=1, padx=5, pady=2)

        # Botão Adicionar
        btn_add = ttk.Button(frame_vars, text="➕ Adicionar Dado à Lista", command=self._adicionar_dado)
        btn_add.grid(row=1, column=4, columnspan=2, sticky="ew", padx=5, pady=5)

        # Tabela Visualização
        frame_tabela = ttk.LabelFrame(self.root, text=" Dados Cadastrados ", padding=10)
        frame_tabela.pack(fill="both", expand=True, padx=10, pady=5)

        colunas = ("dado", "informante", "endereco", "bairro", "valor", "area", "unidade", "unitario")
        self.tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=8)
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

        # Frame Exportações
        frame_acoes = ttk.Frame(self.root, padding=10)
        frame_acoes.pack(fill="x", padx=10, pady=5)

        btn_exp_excel = ttk.Button(frame_acoes, text="📊 Exportar para Excel (SISDEA)", command=self._exportar_excel)
        btn_exp_excel.pack(side="left", padx=10, expand=True, fill="x")

        btn_exp_word = ttk.Button(frame_acoes, text="📄 Exportar para Word (Fichas com Foto)", command=self._exportar_word)
        btn_exp_word.pack(side="right", padx=10, expand=True, fill="x")

    def _selecionar_foto_local(self):
        caminho = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg;*.webp")])
        if caminho:
            self.txt_imagem.delete(0, tk.END)
            self.txt_imagem.insert(0, caminho)

    def _adicionar_dado(self):
        try:
            valor_total = float(self.txt_valor.get().replace(".", "").replace(",", "."))
            area_num = float(self.txt_area.get().replace(".", "").replace(",", "."))
            unidade = self.var_unidade.get()

            # Cálculo Unitário considerando m² ou Hectares
            unitario = valor_total / area_num if area_num > 0 else 0

            novo_dado = {
                "D.": len(self.dados_pesquisas) + 1,
                "Informante": self.txt_informante.get(),
                "Telefone": self.txt_telefone.get(),
                "Endereço": self.txt_endereco.get(),
                "Bairro": self.txt_bairro.get(),
                "Município": self.txt_municipio.get(),
                "Valor Total (R$)": valor_total,
                f"Área Terreno ({unidade})": area_num,
                "Área Construída (m²)": float(self.txt_area_const.get() or 0),
                "Via": self.txt_via.get(),
                "Uso": self.txt_uso.get(),
                "PGV (R$)": self.txt_pgv.get(),
                "Testada (m)": self.txt_testada.get(),
                f"Unitário (R$/{unidade})": unitario,
                "Localização": "Urbana",
                "Coord. E (m)": self.txt_coord_e.get(),
                "Coord. S (m)": self.txt_coord_s.get(),
                "Data": self.txt_data.get(),
                "Link": self.txt_link.get(),
                "Imagem": self.txt_imagem.get(),
                "Unidade": unidade
            }

            self.dados_pesquisas.append(novo_dado)
            self.tree.insert("", "end", values=(
                novo_dado["D."], novo_dado["Informante"], novo_dado["Endereço"],
                novo_dado["Bairro"], f"R$ {valor_total:,.2f}", f"{area_num:,.2f}",
                unidade, f"R$ {unitario:,.2f}"
            ))

            messagebox.showinfo("Sucesso", f"Pesquisa – {novo_dado['D.']:02d} adicionada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro de Preenchimento", f"Verifique os campos numéricos: {e}")

    def _exportar_excel(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not caminho:
            return

        df = pd.DataFrame(self.dados_pesquisas)
        # Remove a coluna de caminho de imagem na tabela pura do SISDEA
        df_sisdea = df.drop(columns=["Imagem", "Unidade"], errors="ignore")
        
        with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
            df_sisdea.to_excel(writer, sheet_name="Ficha de Pesquisa", index=False)

        messagebox.showinfo("Exportação Concluída", "Arquivo Excel gerado com sucesso!")

    def _baixar_imagem(self, caminho_ou_url):
        """Baixa a imagem da web/nuvem ou carrega arquivo local."""
        if not caminho_ou_url:
            return None
        try:
            if caminho_ou_url.startswith("http://") or caminho_ou_url.startswith("https://"):
                # Suporte a links diretos e conversão básica para Google Drive se aplicável
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

    def _exportar_word(self):
        if not self.dados_pesquisas:
            messagebox.showwarning("Aviso", "Nenhum dado cadastrado.")
            return

        caminho = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if not caminho:
            return

        doc = Document()
        
        # Cabeçalho Principal
        titulo = doc.add_paragraph()
        run_titulo = titulo.add_run("PESQUISA DE MERCADO - MEMÓRIA DE CÁLCULO")
        run_titulo.bold = True
        run_titulo.font.size = Pt(12)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for dado in self.dados_pesquisas:
            un = dado.get("Unidade", "m²")
            tabela = doc.add_table(rows=1, cols=2)
            tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
            tabela.autofit = False

            celula_dados = tabela.cell(0, 0)
            celula_foto = tabela.cell(0, 1)

            # Larguras aproximadas
            celula_dados.width = Inches(3.8)
            celula_foto.width = Inches(3.2)

            # Preenchimento das informações
            p = celula_dados.paragraphs[0]
            texto_info = (
                f"Logradouro: {dado.get('Endereço', '')}\n"
                f"Bairro: {dado.get('Bairro', '')}\n"
                f"Município: {dado.get('Município', '')}\n"
                f"Contato: {dado.get('Informante', '')} - {dado.get('Telefone', '')}\n"
                f"Link: {dado.get('Link', '')}\n\n"
                f"Área Terreno: {dado.get(f'Área Terreno ({un})', '')} {un}\n"
                f"Área Construída: {dado.get('Área Construída (m²)', '0')} m²\n"
                f"Valor da Oferta: R$ {dado.get('Valor Total (R$)', 0):,.2f}\n"
                f"Valor Unitário/{un}: R$ {dado.get(f'Unitário (R$/{un})', 0):,.2f}/{un}\n\n"
                f"Via: {dado.get('Via', '')}\n"
                f"Uso: {dado.get('Uso', '')}\n"
                f"Testada: {dado.get('Testada (m)', '')} m\n"
                f"PGV: R$ {dado.get('PGV (R$)', '')}\n"
                f"Coordenadas Geográfica: {dado.get('Coord. E (m)', '')} / {dado.get('Coord. S (m)', '')}\n"
                f"Localização: {dado.get('Localização', 'Urbana')}\n"
                f"Data: {dado.get('Data', '')}\n"
                f"Pesquisa – {dado['D.']:02d}"
            )
            p.text = texto_info

            # Inserção da Imagem do Anúncio
            img_source = self._baixar_imagem(dado.get("Imagem"))
            p_foto = celula_foto.paragraphs[0]
            p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if img_source:
                try:
                    p_foto.add_run().add_picture(img_source, width=Inches(3.0))
                except Exception:
                    p_foto.text = "[ Erro ao carregar imagem ]"
            else:
                p_foto.text = "[ inserir foto / print do anúncio ]"

            # Espaço entre pesquisas
            doc.add_paragraph()

        doc.save(caminho)
        messagebox.showinfo("Exportação Concluída", "Arquivo Word gerado com sucesso!")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppPesquisaMercado(root)
    root.mainloop()
