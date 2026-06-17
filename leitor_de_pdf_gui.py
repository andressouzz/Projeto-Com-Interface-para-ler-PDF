#!/usr/bin/env python3
"""
Leitor de PDF - Interface Gráfica
Leitor de Notas Fiscais de Hardware, Software e Serviço
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# Adiciona o diretório ao path para importar o módulo principal
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
from leitor_de_pdf import LeitordePDF, CruzamentoDados


class LeitorDePDFGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Leitor de Notas Fiscais de Hardware, Software e Serviço")
        self.window.geometry("850x600")
        self.window.configure(bg="#f0f0f0")

        self.dir_path = tk.StringVar(value=str(Path.cwd()))
        from datetime import datetime
        nome_padrao = f"Leitura das Notas Fiscais - {datetime.now():%Y-%m-%d %Hh%M}.xlsx"
        self.excel_path = tk.StringVar(
            value=str(Path.cwd() / nome_padrao)
        )
        self.cruzamento_path = tk.StringVar()
        self._nf_counter = 0
        self._total_pdfs = 0
        self._total_nfs = 0
        self._leitor = None

        self._build_ui()

    def _build_ui(self):
        # Título
        title_frame = tk.Frame(self.window, bg="#f0f0f0", pady=15)
        title_frame.pack(fill="x")

        title = tk.Label(
            title_frame,
            text="Leitor de Notas Fiscais de Hardware, Software e Serviço",
            font=("Segoe UI", 16, "bold"),
            fg="#000080",
            bg="#f0f0f0",
        )
        title.pack()

        # Notebook com abas
        notebook = ttk.Notebook(self.window)
        notebook.pack(fill="both", expand=True, padx=10, pady=5)

        # Estilo: aba selecionada em laranja
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except:
            pass
        style.configure("TNotebook.Tab", background="#F0F0F0")
        style.map("TNotebook.Tab", background=[("selected", "#ED7D31")])

        # --- Aba 1: Leitor ---
        tab_leitor = tk.Frame(notebook, bg="#f0f0f0", padx=20, pady=10)
        notebook.add(tab_leitor, text="Leitor")

        # Local das Notas Fiscais
        self._criar_linha(
            tab_leitor,
            0,
            "Local Onde as Notas Fiscais Estão Salvas",
            self.dir_path,
            "Selecionar Pasta",
            self._selecionar_diretorio,
        )

        # Destino do Excel
        self._criar_linha(
            tab_leitor,
            1,
            "Salvar Planilha como",
            self.excel_path,
            "Selecionar Arquivo",
            self._selecionar_destino_excel,
        )

        # Planilha de Cruzamento
        self._criar_linha(
            tab_leitor,
            2,
            "Planilha para Cruzamento",
            self.cruzamento_path,
            "Selecionar Planilha",
            self._selecionar_cruzamento,
        )

        # Separador
        ttk.Separator(tab_leitor, orient="horizontal").grid(
            row=3, column=0, columnspan=3, sticky="ew", pady=15
        )

        # Botões Executar e Abrir Planilha
        btn_frame = tk.Frame(tab_leitor, bg="#f0f0f0")
        btn_frame.grid(row=4, column=0, columnspan=3, pady=5, sticky="ew")
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(2, weight=1)

        self.btn_executar = tk.Button(
            btn_frame,
            text="Executar",
            font=("Segoe UI", 12, "bold"),
            bg="#000080",
            fg="white",
            padx=30,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self._executar,
        )
        self.btn_executar.grid(row=0, column=1)

        self.btn_abrir = tk.Button(
            btn_frame,
            text="Abrir Planilha",
            font=("Segoe UI", 10, "bold"),
            bg="#28a745",
            fg="white",
            padx=20,
            pady=5,
            relief="flat",
            cursor="hand2",
            state="disabled",
            command=self._abrir_planilha,
        )
        self.btn_abrir.grid(row=0, column=2, sticky="e", padx=(0, 5))

        # Barra de progresso
        self.progress = ttk.Progressbar(
            tab_leitor, mode="determinate", length=600
        )
        self.progress.grid(row=5, column=0, columnspan=3, pady=(10, 5), sticky="ew")

        self.lbl_status = tk.Label(
            tab_leitor,
            text="Pronto para executar",
            font=("Segoe UI", 9),
            fg="#555555",
            bg="#f0f0f0",
        )
        self.lbl_status.grid(row=6, column=0, columnspan=3, pady=(0, 2))

        self.lbl_contagem = tk.Label(
            tab_leitor,
            text="",
            font=("Segoe UI", 10, "bold"),
            fg="#000080",
            bg="#f0f0f0",
        )
        self.lbl_contagem.grid(row=7, column=0, columnspan=3, pady=(0, 5))

        # Console output
        console_frame = tk.Frame(tab_leitor, bg="#f0f0f0")
        console_frame.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(5, 0))
        tab_leitor.grid_rowconfigure(8, weight=1)
        tab_leitor.grid_columnconfigure(0, weight=1)

        self.txt_console = tk.Text(
            console_frame,
            height=8,
            bg="#1e1e1e",
            fg="#d4d4d4",
            font=("Consolas", 9),
            relief="flat",
            padx=5,
            pady=5,
        )
        scrollbar = tk.Scrollbar(console_frame, command=self.txt_console.yview)
        self.txt_console.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.txt_console.pack(fill="both", expand=True)

        # --- Aba 2: Notas Duplicadas ---
        tab_duplicatas = tk.Frame(notebook, bg="#f0f0f0", padx=20, pady=10)
        notebook.add(tab_duplicatas, text="Notas Duplicadas")

        # Treeview para listar duplicatas
        tree_frame = tk.Frame(tab_duplicatas, bg="#f0f0f0")
        tree_frame.pack(fill="both", expand=True)

        self.tree_duplicatas = ttk.Treeview(
            tree_frame,
            columns=("arquivo", "nf", "tipo"),
            show="headings",
            selectmode="extended",
        )
        self.tree_duplicatas.heading("arquivo", text="Arquivo")
        self.tree_duplicatas.heading("nf", text="NF")
        self.tree_duplicatas.heading("tipo", text="Tipo")
        self.tree_duplicatas.column("arquivo", width=500, minwidth=200)
        self.tree_duplicatas.column("nf", width=80, anchor="center")
        self.tree_duplicatas.column("tipo", width=100, anchor="center")

        tree_scroll = tk.Scrollbar(tree_frame, orient="vertical", command=self.tree_duplicatas.yview)
        self.tree_duplicatas.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y")
        self.tree_duplicatas.pack(fill="both", expand=True)

        # Botões
        btn_dup_frame = tk.Frame(tab_duplicatas, bg="#f0f0f0", pady=10)
        btn_dup_frame.pack(fill="x")

        btns = [
            ("Excluir",      "#d9534f", self._excluir_duplicatas),
            ("Mover Para",   "#f0ad4e", self._mover_duplicatas),
            ("Copiar Para",  "#5bc0de", self._copiar_duplicatas),
        ]
        for texto, cor, comando in btns:
            btn = tk.Button(
                btn_dup_frame,
                text=texto,
                font=("Segoe UI", 10, "bold"),
                bg=cor,
                fg="white",
                padx=15,
                pady=5,
                relief="flat",
                cursor="hand2",
                command=comando,
            )
            btn.pack(side="left", padx=5)

    def _criar_linha(self, parent, row, label_text, var, btn_text, btn_cmd):
        tk.Label(
            parent,
            text=label_text,
            font=("Segoe UI", 10, "bold"),
            bg="#f0f0f0",
            anchor="w",
        ).grid(row=row, column=0, sticky="w", pady=(8, 2))

        entry_frame = tk.Frame(parent, bg="#f0f0f0")
        entry_frame.grid(row=row, column=1, sticky="ew", padx=(0, 10), pady=(8, 2))
        parent.grid_columnconfigure(1, weight=1)

        entry = tk.Entry(
            entry_frame,
            textvariable=var,
            font=("Segoe UI", 9),
            bg="white",
            relief="solid",
            bd=1,
        )
        entry.pack(fill="x", expand=True)

        btn = tk.Button(
            parent,
            text=btn_text,
            font=("Segoe UI", 9),
            bg="#e0e0e0",
            padx=10,
            width=18,
            cursor="hand2",
            command=btn_cmd,
        )
        btn.grid(row=row, column=2, sticky="w", pady=(8, 2))

    def _selecionar_diretorio(self):
        path = filedialog.askdirectory(initialdir=self.dir_path.get())
        if path:
            self.dir_path.set(path)

    def _selecionar_destino_excel(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=os.path.basename(self.excel_path.get()),
        )
        if path:
            self.excel_path.set(path)

    def _selecionar_cruzamento(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls")],
            title="Selecionar planilha para cruzamento de dados",
        )
        if path:
            self.cruzamento_path.set(path)

    def _log(self, msg):
        self.txt_console.insert("end", msg + "\n")
        self.txt_console.see("end")
        self.window.update_idletasks()

    def _executar(self):
        if not self.dir_path.get():
            messagebox.showwarning("Aviso", "Selecione o diretório onde as notas fiscais estão salvas.")
            return

        self.btn_executar.configure(state="disabled", text="Processando...")
        self.progress["value"] = 0
        self.txt_console.delete("1.0", "end")
        self.lbl_status.configure(text="Processando...")

        # Limpa treeview de duplicatas
        for item in self.tree_duplicatas.get_children():
            self.tree_duplicatas.delete(item)

        thread = threading.Thread(target=self._processar, daemon=True)
        thread.start()

    def _processar(self):
        try:
            base_dir = self.dir_path.get()
            dirs = [base_dir]
            base = Path(base_dir)
            # Adiciona subdiretórios separadamente (como faz o CLI) para detectar duplicatas
            for subdir in sorted(base.iterdir()):
                if subdir.is_dir() and not subdir.name.startswith('.'):
                    if list(subdir.rglob("*.pdf")):
                        dirs.append(str(subdir))
            pdfs = sorted(base.rglob("*.pdf"))
            self._total_pdfs = len(pdfs)
            self._nf_counter = 0
            self._total_nfs = 0
            self.window.after(0, self._atualizar_contagem)

            # --- Etapa 1: Processar PDFs ---
            self._log("=" * 50)
            self._log("LEITOR DE PDF - Notas Fiscais")
            self._log("=" * 50)

            leitor = LeitordePDF()
            leitor.excel_path = self.excel_path.get()

            class ConsoleRedirect:
                def __init__(self, gui):
                    self.gui = gui

                def write(self, text):
                    if text.strip():
                        stripped = text.rstrip()
                        self.gui._log(stripped)
                        if "NF encontrada" in stripped:
                            self.gui._nf_counter += 1
                            self.gui.window.after(0, self.gui._atualizar_contagem)

                def flush(self):
                    pass

            sys.stdout = ConsoleRedirect(self)

            leitor.processar_pdfs(dirs)

            self._total_nfs = len(leitor.dados_nf)
            self._leitor = leitor

            if not leitor.dados_nf:
                self._log("\nNenhuma nota fiscal encontrada.")
                sys.stdout = sys.__stdout__
                self._finalizar()
                return

            self._log(f"\nTotal de NFs extraidas: {self._total_nfs}")
            self.window.after(0, self._atualizar_contagem)
            leitor.criar_planilha_excel()
            self._log(f"\nPlanilha salva em: {self.excel_path.get()}")

            # --- Etapa 2: Cruzamento de dados ---
            if self.cruzamento_path.get():
                self._log("\n" + "=" * 50)
                self._log("Iniciando cruzamento de dados...")
                self._log("=" * 50)

                self.window.after(0, self._atualizar_progresso, 0, 1, "Cruzando dados...")

                try:
                    CruzamentoDados.executar(
                        self.excel_path.get(),
                        self.cruzamento_path.get(),
                        progress_callback=self._progress_callback,
                    )
                    self._log("\nCruzamento de dados concluido!")
                    self._log("   Colunas adicionadas: CGC, CIAUS, Field Responsavel")
                    leitor._recarregar_dados_cruzamento()
                    self._log(f"\nTotal de NFs apos cruzamento: {len(leitor.dados_nf)}")
                except Exception as e:
                    self._log(f"\nErro no cruzamento: {e}")

            sys.stdout = sys.__stdout__

            self._exibir_preview_console(leitor)
            self._carregar_duplicatas(leitor)

            self._finalizar()

        except Exception as e:
            sys.stdout = sys.__stdout__
            self._log(f"\nErro: {e}")
            self._finalizar(erro=True)

    def _carregar_duplicatas(self, leitor):
        duplicatas = {
            nf: paths for nf, paths in leitor._nf_para_arquivos.items()
            if len(paths) > 1
        }
        self.window.after(0, self._popular_treeview_duplicatas, duplicatas)

    def _popular_treeview_duplicatas(self, duplicatas):
        self.tree_duplicatas.delete(*self.tree_duplicatas.get_children())
        if not duplicatas:
            self._log("Nenhuma duplicata encontrada.")
            return
        for nf in sorted(duplicatas.keys()):
            paths = duplicatas[nf]
            for i, path in enumerate(paths):
                tipo = "Principal" if i == 0 else "Duplicata"
                self.tree_duplicatas.insert("", "end", values=(path, nf, tipo))
        self._log(f"{len(duplicatas)} NF(s) com duplicatas carregadas na aba.")

    def _obter_selecionados(self):
        selecao = []
        for item_id in self.tree_duplicatas.selection():
            valores = self.tree_duplicatas.item(item_id, "values")
            if len(valores) >= 1:
                selecao.append(valores[0])
        return selecao

    def _excluir_duplicatas(self):
        selecionados = self._obter_selecionados()
        if not selecionados:
            messagebox.showinfo("Aviso", "Selecione ao menos um arquivo na lista.")
            return
        qtd = len(selecionados)
        if not messagebox.askyesno("Confirmar", f"Excluir permanentemente {qtd} arquivo(s)?"):
            return
        excluidos = 0
        for path in selecionados:
            try:
                os.remove(path)
                excluidos += 1
            except Exception as e:
                self._log(f"Erro ao excluir {path}: {e}")
        for item_id in self.tree_duplicatas.selection():
            self.tree_duplicatas.delete(item_id)
        self._log(f"{excluidos} arquivo(s) excluido(s).")
        messagebox.showinfo("Concluido", f"{excluidos} arquivo(s) excluido(s).")

    def _mover_duplicatas(self):
        selecionados = self._obter_selecionados()
        if not selecionados:
            messagebox.showinfo("Aviso", "Selecione ao menos um arquivo na lista.")
            return
        destino = filedialog.askdirectory(title="Selecionar pasta de destino")
        if not destino:
            return
        movidos = 0
        for path in selecionados:
            try:
                nome = os.path.basename(path)
                dest_path = os.path.join(destino, nome)
                os.rename(path, dest_path)
                movidos += 1
            except Exception as e:
                self._log(f"Erro ao mover {path}: {e}")
        for item_id in self.tree_duplicatas.selection():
            self.tree_duplicatas.delete(item_id)
        self._log(f"{movidos} arquivo(s) movido(s) para {destino}.")
        messagebox.showinfo("Concluido", f"{movidos} arquivo(s) movido(s).")

    def _copiar_duplicatas(self):
        selecionados = self._obter_selecionados()
        if not selecionados:
            messagebox.showinfo("Aviso", "Selecione ao menos um arquivo na lista.")
            return
        destino = filedialog.askdirectory(title="Selecionar pasta de destino")
        if not destino:
            return
        import shutil
        copiados = 0
        for path in selecionados:
            try:
                nome = os.path.basename(path)
                dest_path = os.path.join(destino, nome)
                shutil.copy2(path, dest_path)
                copiados += 1
            except Exception as e:
                self._log(f"Erro ao copiar {path}: {e}")
        self._log(f"{copiados} arquivo(s) copiado(s) para {destino}.")
        messagebox.showinfo("Concluido", f"{copiados} arquivo(s) copiado(s).")

    def _abrir_planilha(self):
        path = self.excel_path.get()
        if not path or not os.path.exists(path):
            messagebox.showinfo("Aviso", "Nenhuma planilha encontrada. Execute o processamento primeiro.")
            return
        try:
            os.startfile(path)
        except Exception as e:
            self._log(f"Erro ao abrir planilha: {e}")

    def _progress_callback(self, atual, total):
        if total > 0:
            pct = int((atual / total) * 100)
            self.window.after(0, self._atualizar_progresso, pct, total,
                              f"Cruzando... {atual + 1}/{total}")

    def _atualizar_contagem(self):
        if self._total_nfs > 0:
            texto = f"Notas lidas: {self._total_nfs} de {self._total_pdfs} PDFs processados"
        elif self._nf_counter > 0:
            texto = f"Lendo nota {self._nf_counter} de {self._total_pdfs} PDFs..."
        else:
            texto = f"Encontrados {self._total_pdfs} arquivos PDF"
        self.lbl_contagem.configure(text=texto)
        self.window.update_idletasks()

    def _atualizar_progresso(self, valor, total, texto):
        self.progress["value"] = valor
        self.lbl_status.configure(text=texto)
        self.window.update_idletasks()

    def _exibir_preview_console(self, leitor):
        if not leitor.dados_nf:
            return
        self._log("\n" + "=" * 80)
        self._log("PREVIEW FINAL DA PLANILHA")
        self._log("=" * 80)
        header = (
            f"{'Arquivo':20} | {'NF':>8} | {'Tipo':10} | "
            f"{'Cliente':20} | {'Data':10} | {'UF':3} | "
            f"{'Valor':>10} | {'SAP':12} | {'CGC':>8} | {'Field':10}"
        )
        self._log(header)
        self._log("-" * 80)
        for dado in leitor.dados_nf:
            cgc = dado.get('cgc', 'N/A')
            field = dado.get('field_responsavel', 'N/A')
            self._log(
                f"{dado.get('arquivo','')[:18]:20} | "
                f"{dado['numero_nf']:>8} | "
                f"{dado.get('tipo_nota_fiscal','')[:8]:10} | "
                f"{dado['cliente'][:18]:20} | "
                f"{dado['data_emissao']:10} | "
                f"{dado['uf_destino']:3} | "
                f"R$ {dado['valor_total']:>7} | "
                f"{dado['codigo_sap']:12} | "
                f"{str(cgc)[:8]:>8} | "
                f"{str(field)[:8]:10}"
            )
        self._log("=" * 80)
        self._log(f"Total: {len(leitor.dados_nf)} linha(s)")

    def _finalizar(self, erro=False):
        self.window.after(0, lambda: self._ui_finalizar(erro))

    def _ui_finalizar(self, erro=False):
        self.progress["value"] = 100 if not erro else 0
        self.btn_executar.configure(state="normal", text="Executar")
        if erro:
            self.lbl_status.configure(text="Erro na execucao")
        else:
            self.btn_abrir.configure(state="normal")
            self.lbl_status.configure(text="Concluido!")
            self.lbl_contagem.configure(
                text=f"Leitura concluida - {self._total_nfs} notas fiscais lidas"
            )
            messagebox.showinfo(
                "Concluido",
                "Leitura das notas foi feita com sucesso e o arquivo excel foi criado no destino especificado."
            )
        self.window.update_idletasks()

    def iniciar(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = LeitorDePDFGUI()
    app.iniciar()
