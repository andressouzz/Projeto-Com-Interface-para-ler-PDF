#!/usr/bin/env python3
"""
Leitor de PDF - Interface Gráfica
Leitor de Notas Ficais de Hardware
"""

import os
import sys
import difflib
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Adiciona o diretório ao path para importar o módulo principal
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
from leitor_de_pdf import LeitordePDF


class CruzamentoDados:
    """Cruzamento de dados aproximado (fuzzy matching) entre planilhas."""

    COLUNAS_ALVO = ['CGC', 'CIAUS', 'Field Responsável']

    # Padrões para identificar colunas de endereço na planilha de cruzamento
    PADROES_ENDERECO = {
        'endereco': ['endereço', 'endereco', 'logradouro', 'end'],
        'bairro': ['bairro', 'bairro/distrito'],
        'cidade': ['cidade', 'município', 'municipio'],
        'cep': ['cep', 'postal', 'zip'],
    }

    @staticmethod
    def _normalizar(texto):
        if not texto:
            return ""
        return re.sub(r'\s+', ' ', str(texto).lower().strip())

    @staticmethod
    def _encontrar_coluna(headers, padroes):
        headers_lower = [str(h).lower().strip() if h else "" for h in headers]
        for i, h in enumerate(headers_lower):
            for padrao in padroes:
                if padrao in h:
                    return i
        return None

    @classmethod
    def _mapear_colunas_cruzamento(cls, headers):
        mapeamento = {}
        for campo, padroes in cls.PADROES_ENDERECO.items():
            idx = cls._encontrar_coluna(headers, padroes)
            if idx is not None:
                mapeamento[campo] = idx
        for alvo in cls.COLUNAS_ALVO:
            idx = cls._encontrar_coluna(headers, [alvo.lower()])
            if idx is not None:
                mapeamento[alvo] = idx
        return mapeamento

    @staticmethod
    def _criar_chave(endereco, bairro, cidade, cep):
        partes = [
            str(endereco or ""),
            str(bairro or ""),
            str(cidade or ""),
            str(cep or ""),
        ]
        return re.sub(r'\s+', ' ', ' '.join(partes).lower().strip())

    @classmethod
    def executar(cls, caminho_excel_gerado, caminho_cruzamento, progress_callback=None):
        wb = load_workbook(caminho_excel_gerado)
        ws = wb.active

        wb_cruz = load_workbook(caminho_cruzamento, data_only=True)
        ws_cruz = wb_cruz.active

        headers_cruz = [cell.value for cell in ws_cruz[1]]
        mapa = cls._mapear_colunas_cruzamento(headers_cruz)

        # Verifica se encontrou as colunas necessárias no cruzamento
        cols_necessarias = ['endereco', 'cidade']
        for col in cols_necessarias:
            if col not in mapa:
                raise ValueError(
                    f"Coluna '{col}' não encontrada na planilha de cruzamento. "
                    f"Cabeçalhos encontrados: {headers_cruz}"
                )

        # Constrói lista de referência da planilha de cruzamento
        ref_data = []
        for row in ws_cruz.iter_rows(min_row=2, values_only=True):
            end = row[mapa['endereco']] if mapa['endereco'] < len(row) else ""
            bai = row[mapa['bairro']] if mapa.get('bairro', 0) < len(row) else ""
            cid = row[mapa['cidade']] if mapa['cidade'] < len(row) else ""
            cep = row[mapa['cep']] if mapa.get('cep', 0) < len(row) else ""

            ref_data.append({
                'chave': cls._criar_chave(end, bai, cid, cep),
                'cgc': row[mapa['CGC']] if mapa.get('CGC', 0) < len(row) else "",
                'ciaus': row[mapa['CIAUS']] if mapa.get('CIAUS', 0) < len(row) else "",
                'field': row[mapa['Field Responsável']] if mapa.get('Field Responsável', 0) < len(row) else "",
            })

        if not ref_data:
            raise ValueError("Planilha de cruzamento está vazia (sem linhas de dados).")

        headers_gerado = [cell.value for cell in ws[1]]

        # Encontra índices das colunas de endereço no Excel gerado
        def idx_col(nome):
            for i, h in enumerate(headers_gerado):
                if h and nome.lower() in str(h).lower():
                    return i
            return None

        idx_end = idx_col("Endereço")
        idx_bairro = idx_col("Bairro")
        idx_cidade = idx_col("Cidade")
        idx_cep = idx_col("CEP")

        if idx_end is None or idx_cidade is None:
            raise ValueError(
                "Colunas de endereço não encontradas na planilha gerada."
            )

        # Adiciona cabeçalhos das novas colunas (preservando formatação)
        ultima_col = ws.max_column
        estilo_cabecalho = None
        cell_ref = ws.cell(1, 1)
        if cell_ref.font:
            estilo_cabecalho = {
                'fill': cell_ref.fill,
                'font': cell_ref.font,
                'alignment': cell_ref.alignment,
            }

        for i, nome_col in enumerate(['CGC (Cruzamento)', 'CIAUS', 'Field Responsável']):
            col = ultima_col + 1 + i
            cell = ws.cell(1, col)
            cell.value = nome_col
            if estilo_cabecalho:
                cell.fill = estilo_cabecalho['fill']
                cell.font = estilo_cabecalho['font']
                cell.alignment = estilo_cabecalho['alignment']

        # Define largura das novas colunas
        col_letter = lambda n: chr(64 + n) if n <= 26 else None
        widths = [18, 15, 20]
        for i, w in enumerate(widths):
            col = ultima_col + 1 + i
            letra = col_letter(col)
            if letra:
                ws.column_dimensions[letra].width = w

        total_linhas = ws.max_row - 1
        for row_idx in range(2, ws.max_row + 1):
            if progress_callback:
                progress_callback(row_idx - 2, total_linhas)

            # Constrói chave da linha atual
            end = cls._normalizar(ws.cell(row_idx, idx_end + 1).value)
            bai = cls._normalizar(ws.cell(row_idx, idx_bairro + 1).value) if idx_bairro else ""
            cid = cls._normalizar(ws.cell(row_idx, idx_cidade + 1).value)
            cep = cls._normalizar(ws.cell(row_idx, idx_cep + 1).value) if idx_cep else ""

            chave_origem = f"{end} {bai} {cid} {cep}"

            # Encontra melhor correspondência fuzzy
            melhor_score = 0
            melhor_ref = None
            for ref in ref_data:
                score = difflib.SequenceMatcher(None, chave_origem, ref['chave']).ratio()
                if score > melhor_score:
                    melhor_score = score
                    melhor_ref = ref

            # Escreve dados se o score for aceitável
            threshold = 0.4
            if melhor_ref and melhor_score >= threshold:
                ws.cell(row_idx, ultima_col + 1).value = melhor_ref['cgc']
                ws.cell(row_idx, ultima_col + 2).value = melhor_ref['ciaus']
                ws.cell(row_idx, ultima_col + 3).value = melhor_ref['field']

        # Aplica bordas nas novas colunas
        borda = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )
        for row_idx in range(1, ws.max_row + 1):
            for col in range(ultima_col + 1, ultima_col + 4):
                ws.cell(row_idx, col).border = borda
                if row_idx > 1:
                    ws.cell(row_idx, col).font = Font(color="000080", size=10)
                    ws.cell(row_idx, col).alignment = Alignment(horizontal="center", vertical="center")

        wb.save(caminho_excel_gerado)
        return True


class LeitorDePDFGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Leitor de Notas Ficais de Hardware")
        self.window.geometry("750x500")
        self.window.configure(bg="#f0f0f0")

        self.dir_path = tk.StringVar(value=str(Path.cwd()))
        self.excel_path = tk.StringVar(
            value=str(Path.cwd() / "Leitura das Notas Fiscais.xlsx")
        )
        self.cruzamento_path = tk.StringVar()
        self._nf_counter = 0
        self._total_pdfs = 0
        self._total_nfs = 0

        self._build_ui()

    def _build_ui(self):
        # Título
        title_frame = tk.Frame(self.window, bg="#f0f0f0", pady=20)
        title_frame.pack(fill="x")

        title = tk.Label(
            title_frame,
            text="Leitor de Notas Ficais de Hardware",
            font=("Segoe UI", 18, "bold"),
            fg="#000080",
            bg="#f0f0f0",
        )
        title.pack()

        # Frame principal
        main_frame = tk.Frame(self.window, bg="#f0f0f0", padx=30, pady=10)
        main_frame.pack(fill="both", expand=True)

        # --- Botão 1: Diretório dos PDFs ---
        self._criar_linha(
            main_frame,
            0,
            "📂 Diretório dos PDFs",
            self.dir_path,
            "Selecionar Pasta",
            self._selecionar_diretorio,
        )

        # --- Botão 2: Destino do Excel ---
        self._criar_linha(
            main_frame,
            1,
            "📊 Salvar Planilha como",
            self.excel_path,
            "Selecionar Arquivo",
            self._selecionar_destino_excel,
        )

        # --- Botão 3: Planilha de Cruzamento ---
        self._criar_linha(
            main_frame,
            2,
            "🔗 Planilha para Cruzamento",
            self.cruzamento_path,
            "Selecionar Planilha",
            self._selecionar_cruzamento,
        )

        # Separador
        ttk.Separator(main_frame, orient="horizontal").grid(
            row=3, column=0, columnspan=3, sticky="ew", pady=20
        )

        # Botão Executar
        btn_frame = tk.Frame(main_frame, bg="#f0f0f0")
        btn_frame.grid(row=4, column=0, columnspan=3, pady=10)

        self.btn_executar = tk.Button(
            btn_frame,
            text="▶ Executar",
            font=("Segoe UI", 12, "bold"),
            bg="#000080",
            fg="white",
            padx=30,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self._executar,
        )
        self.btn_executar.pack()

        # Barra de progresso
        self.progress = ttk.Progressbar(
            main_frame, mode="determinate", length=600
        )
        self.progress.grid(row=5, column=0, columnspan=3, pady=(10, 5), sticky="ew")

        self.lbl_status = tk.Label(
            main_frame,
            text="Pronto para executar",
            font=("Segoe UI", 9),
            fg="#555555",
            bg="#f0f0f0",
        )
        self.lbl_status.grid(row=6, column=0, columnspan=3, pady=(0, 2))

        self.lbl_contagem = tk.Label(
            main_frame,
            text="",
            font=("Segoe UI", 10, "bold"),
            fg="#000080",
            bg="#f0f0f0",
        )
        self.lbl_contagem.grid(row=7, column=0, columnspan=3, pady=(0, 5))

        # Console output
        console_frame = tk.Frame(main_frame, bg="#f0f0f0")
        console_frame.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(5, 0))
        main_frame.grid_rowconfigure(8, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

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

    def _criar_linha(self, parent, row, label_text, var, btn_text, btn_cmd):
        tk.Label(
            parent,
            text=label_text,
            font=("Segoe UI", 10, "bold"),
            bg="#f0f0f0",
            anchor="w",
        ).grid(row=row, column=0, sticky="w", pady=(10, 2))

        entry_frame = tk.Frame(parent, bg="#f0f0f0")
        entry_frame.grid(row=row, column=1, sticky="ew", padx=(0, 10), pady=(10, 2))
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
            cursor="hand2",
            command=btn_cmd,
        )
        btn.grid(row=row, column=2, sticky="w", pady=(10, 2))

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
            messagebox.showwarning("Aviso", "Selecione o diretório dos PDFs.")
            return

        self.btn_executar.configure(state="disabled", text="⏳ Processando...")
        self.progress["value"] = 0
        self.txt_console.delete("1.0", "end")
        self.lbl_status.configure(text="Processando...")

        thread = threading.Thread(target=self._processar, daemon=True)
        thread.start()

    def _processar(self):
        import re  # needed by CruzamentoDados._normalizar

        try:
            # Conta PDFs antes de processar
            pdfs = sorted(Path(self.dir_path.get()).rglob("*.pdf"))
            self._total_pdfs = len(pdfs)
            self._nf_counter = 0
            self._total_nfs = 0
            self.window.after(0, self._atualizar_contagem)

            # --- Etapa 1: Processar PDFs ---
            self._log("=" * 50)
            self._log("🔍 LEITOR DE PDF - Notas Fiscais")
            self._log("=" * 50)

            leitor = LeitordePDF()
            leitor.excel_path = self.excel_path.get()

            # Redireciona prints para o console e conta NFs
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

            leitor.processar_pdfs(self.dir_path.get())

            self._total_nfs = len(leitor.dados_nf)

            if not leitor.dados_nf:
                self._log("\n⚠️  Nenhuma nota fiscal encontrada.")
                sys.stdout = sys.__stdout__
                self._finalizar()
                return

            self._log(f"\n📊 Total de NFs extraídas: {self._total_nfs}")
            self.window.after(0, self._atualizar_contagem)
            leitor.criar_planilha_excel()
            self._log(f"\n✅ Planilha salva em: {self.excel_path.get()}")

            # --- Etapa 2: Cruzamento de dados ---
            if self.cruzamento_path.get():
                self._log("\n" + "=" * 50)
                self._log("🔗 Iniciando cruzamento de dados...")
                self._log("=" * 50)

                self.window.after(0, self._atualizar_progresso, 0, 1, "Cruzando dados...")

                try:
                    CruzamentoDados.executar(
                        self.excel_path.get(),
                        self.cruzamento_path.get(),
                        progress_callback=self._progress_callback,
                    )
                    self._log("\n✅ Cruzamento de dados concluído!")
                    self._log(f"   Colunas adicionadas: CGC, CIAUS, Field Responsável")
                except Exception as e:
                    self._log(f"\n⚠️  Erro no cruzamento: {e}")

            sys.stdout = sys.__stdout__
            self._finalizar()

        except Exception as e:
            sys.stdout = sys.__stdout__
            self._log(f"\n❌ Erro: {e}")
            self._finalizar(erro=True)

    def _progress_callback(self, atual, total):
        if total > 0:
            pct = int((atual / total) * 100)
            self.window.after(0, self._atualizar_progresso, pct, total,
                              f"Cruzando... {atual + 1}/{total}")

    def _atualizar_contagem(self):
        if self._total_nfs > 0:
            texto = f"✅ Notas lidas: {self._total_nfs} de {self._total_pdfs} PDFs processados"
        elif self._nf_counter > 0:
            texto = f"📖 Lendo nota {self._nf_counter} de {self._total_pdfs} PDFs..."
        else:
            texto = f"📄 Encontrados {self._total_pdfs} arquivos PDF"
        self.lbl_contagem.configure(text=texto)
        self.window.update_idletasks()

    def _atualizar_progresso(self, valor, total, texto):
        self.progress["value"] = valor
        self.lbl_status.configure(text=texto)
        self.window.update_idletasks()

    def _finalizar(self, erro=False):
        self.window.after(0, lambda: self._ui_finalizar(erro))

    def _ui_finalizar(self, erro=False):
        self.progress["value"] = 100 if not erro else 0
        self.btn_executar.configure(state="normal", text="▶ Executar")
        if erro:
            self.lbl_status.configure(text="Erro na execução")
        else:
            self.lbl_status.configure(text="Concluído!")
            self.lbl_contagem.configure(
                text=f"✅ Leitura concluída - {self._total_nfs} notas fiscais lidas"
            )
            messagebox.showinfo(
                "Concluído",
                "Leitura das notas foi feita com sucesso e o arquivo excel foi criado no destino especificado."
            )
        self.window.update_idletasks()

    def iniciar(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = LeitorDePDFGUI()
    app.iniciar()
