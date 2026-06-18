#Código contruido por André Souza 100% no Ipad

"""
Leitor de PDF - Extrai dados de notas fiscais em PDFs
e gera uma planilha Excel formatada.
"""

import re
from pathlib import Path
from PyPDF2 import PdfReader
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from rich.console import Console
from rich.table import Table


class LeitordePDF:
    """Classe para ler PDFs e extrair dados de notas fiscais."""

    UFS_BRASIL = frozenset({
        'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS',
        'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC',
        'SP', 'SE', 'TO',
    })
    FORMATO_CONTABIL_BR = '_-R$ * #,##0.00_-;R$ * (#,##0.00)_-;_-R$ * "-"??_-;_-@_-'
    
    def __init__(self):
        self.dados_nf = []
        self._nf_para_arquivos = {}
        from datetime import datetime
        self.excel_path = f"Leitura das Notas Fiscais - {datetime.now():%Y-%m-%d %Hh%M}.xlsx"
        self.console = Console()
    
    def _identificar_tipo_nf(self, numero_nf):
        if not numero_nf:
            return None
        if (len(numero_nf) == 5 and numero_nf[0] in ('7', '8')) or \
           (len(numero_nf) == 8 and numero_nf[0] == '2'):
            return "Hardware"
        if (len(numero_nf) == 6 and numero_nf[0] in ('6', '7', '8', '9')) or \
           (len(numero_nf) == 7 and numero_nf[0] == '1'):
            return "Software"
        return None

    def extrair_numero_nf(self, texto):
        """
        Extrai o número da nota fiscal do texto.
        Suporta hardware (Nº 000074734) e software (NFS-e, Número:).
        """
        candidatos = []

        # Padrões comuns: Nº, NF-e, Nota Fiscal etc.
        padoes = [
            r'Nº\s+(\d+)',
            r'N[oº]\s+(\d+)',
            r'NF-?[e]?\s+(\d+)',
            r'Nota\s+Fiscal\s+(\d+)',
            r'NF\s*[#:-]?\s*(\d+)',
        ]
        for padrao in padoes:
            for m in re.finditer(padrao, texto, re.IGNORECASE):
                candidatos.append(m.group(1))

        # Software: "Nmero da NFS-e\n1021539" - específico, prioridade máxima
        match = re.search(r'N.mero\s+da\s+NFS[-\s]e\s*(\d+)', texto, re.DOTALL)
        if match:
            return str(int(match.group(1)))

        # Software: "Nmero: 601146" ou "NMERO DA NOTA\n750645"
        for m in re.finditer(r'N.mero.{0,40}?(\d{5,})', texto, re.DOTALL | re.IGNORECASE):
            candidatos.append(m.group(1))

        # Filtra: limpa zeros à esquerda, descarta números curtos
        limpos = []
        for num in candidatos:
            try:
                limpo = str(int(num))
                if len(limpo) >= 5:
                    limpos.append(limpo)
            except (ValueError, TypeError):
                continue

        if not limpos:
            return None

        # Retorna o mais longo (maior chance de ser o NF real)
        return max(limpos, key=len)
    
    def extrair_cliente(self, texto):
        # Hardware: NOME/RAZÃO SOCIAL
        m = re.search(r'NOME/RAZ[AÃ]O\s+SOCIAL\s*(.+?)\s*C\.\s*N\.\s*P\.\s*J\.', texto, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        # Software: busca "CAIXA ECONOMICA FEDERAL" (presente em todos os formatos)
        m = re.search(r'CAIXA\s+ECON[OÔ]MICA\s+FEDERAL', texto, re.IGNORECASE)
        if m:
            return m.group(0).strip()
        return None

    def extrair_data_emissao(self, texto):
        # Hardware: DATA DA EMISSÃO DD.MM.YYYY
        m = re.search(r'DATA\s*DA\s*EMISS[AÃ]O\s*(\d{2})\.(\d{2})\.(\d{4})', texto, re.IGNORECASE)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        # Software Format A: Data e Hora da emissão da NFS-e\nDD/MM/YYYY
        m = re.search(r'[Dd]ata\s+e\s+[Hh]ora\s+da\s+emiss[ãa]o\s+da\s+NFS[-\s]e\s*(\d{2})/(\d{2})/(\d{4})', texto, re.DOTALL)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        # Software Format B: Data Emissão:\nDD/MM/YYYY
        m = re.search(r'Data\s+Emiss[ãa]o:\s*(\d{2})/(\d{2})/(\d{4})', texto, re.IGNORECASE)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        # Software Format C: DATA DE EMISSÃO\nDD/MM/YYYY
        m = re.search(r'DATA\s+DE\s+EMISS[ÃA]O\s*(\d{2})/(\d{2})/(\d{4})', texto)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        # Serviço: Data de Emissão: DD/MM/YYYY
        m = re.search(r'Data\s+de\s+Emiss[ãa]o:\s*(\d{2})/(\d{2})/(\d{4})', texto, re.IGNORECASE)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        # Generic: DD/MM/YYYY near "emiss" keyword
        m = re.search(r'[Ee]miss[ãa]o[^d]*?(\d{2})/(\d{2})/(\d{4})', texto)
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        return None
    
    def extrair_uf_destino(self, texto):
        # Hardware: FONE/FAX UF
        m = re.search(r'FONE/FAX\s+UF\s*([A-Z]{2})', texto, re.IGNORECASE)
        if m:
            return m.group(1)
        # Software Format B/D: UF:\nXX (DADOS DO TOMADOR section)
        m = re.search(r'UF[:\s]*\n\s*([A-Z]{2})\b', texto)
        if m and m.group(1) in self.UFS_BRASIL:
            return m.group(1)
        # Software: "CIDADE/UF" pattern (e.g., JARAGUA DO SUL/SC)
        m = re.search(r'([A-Za-zÀ-ÿ\s]+)/(AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO)\b', texto)
        if m:
            return m.group(2)
        # Software: UF: XX (inline)
        m = re.search(r'\bUF:[ \t]*([A-Z]{2})\b', texto)
        if m and m.group(1) in self.UFS_BRASIL:
            return m.group(1)
        return None
    
    def extrair_valor_total(self, texto):
        # Hardware: VALOR TOTAL DA NOTA
        m = re.search(r'VALOR\s+TOTAL\s+DA\s*NOTA\s*([\d.]+,\d+)', texto, re.IGNORECASE)
        if m:
            return m.group(1)
        # Software Format B: VALOR LÍQUIDO DA NOTA
        m = re.search(r'VALOR\s+L[IÍ]QUIDO\s+DA\s+NOTA\s*R?\$?\s*([\d.]+,\d+)', texto, re.IGNORECASE)
        if m:
            return m.group(1)
        # Software Format C: VALOR LÍQUIDO: R$ ...
        m = re.search(r'VALOR\s+L[IÍ]QUIDO[:\s]+\s*R?\$?\s*([\d.]+,\d+)', texto, re.IGNORECASE)
        if m:
            return m.group(1)
        # Software Format C: VALOR DOS SERVIÇOS: R$ ...
        m = re.search(r'VALOR\s+DOS\s+SERVI[ÇC]OS[:\s]+\s*R?\$?\s*([\d.]+,\d+)', texto, re.IGNORECASE)
        if m:
            return m.group(1)
        # Software Format A: Valor do Serviço (first occurrence)
        m = re.search(r'Valor\s+do\s+SERVI[ÇC]O\s*R?\$?\s*([\d.]+,\d+)', texto, re.IGNORECASE)
        if m:
            return m.group(1)
        # Serviço: VALOR TOTAL DA NOTA (com quebra)
        m = re.search(r'VALOR\s+TOTAL\s+DA\s+NOTA\s*R?\$?\s*([\d.]+,\d+)', texto, re.IGNORECASE)
        if m:
            return m.group(1)
        return None
    
    def extrair_cgc(self, texto):
        """
        Extrai o CGC (CNPJ) da seção de informações complementares.
        Procura por "CGC: X" na seção de informações complementares.
        """
        padrao = r'CGC[:\s]+(\d+)'
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            return match.group(1)
        return None
    
    def _texto_endereco_entrega(self, texto):
        """Extrai o texto do campo Endereço Entrega (informações complementares)."""
        matches = list(re.finditer(
            r'endere[cç]o\s+(?:de\s+)?entrega\s*:\s*(.+?)(?=\s*-?\s*Contrato\s|\s*-?\s*Conta\s+banc[aá]ria|$)',
            texto,
            re.IGNORECASE | re.DOTALL,
        ))
        if not matches:
            return None
        melhor = max(matches, key=lambda m: len(m.group(1)))
        return melhor.group(1).replace('\n', ' ')

    def _normalizar_texto_endereco(self, raw):
        """Padroniza separadores colados pelo extrator do PDF (ex: -Campus, GO -)."""
        raw = re.sub(r'\s+', ' ', raw).strip()
        raw = re.sub(r'-\s*([A-Za-zÀ-ÿ])', r'- \1', raw)
        raw = re.sub(r'\s*-\s*', ' - ', raw)
        return raw

    def _eh_parte_logradouro(self, parte, continua_logradouro=False):
        """Indica se o trecho pertence ao logradouro (rua/número), não ao bairro."""
        if re.search(r'\b(KM|NR|LOTE|num\.?|n[ºo]\.?)\b', parte, re.IGNORECASE):
            return True
        if re.match(
            r'^(R\.?|Av\.?|Est\.?|Rod\.?|Rua|Avenida|Travessa|Praça|Alameda)\b',
            parte,
            re.IGNORECASE,
        ):
            return True
        if continua_logradouro and re.search(r'\d', parte):
            return True
        if continua_logradouro and re.match(r'^Campus\s', parte, re.IGNORECASE):
            return True
        return False

    def _indice_fim_logradouro(self, partes, uf_idx):
        """Último índice das partes que compõem o logradouro, antes do bairro."""
        if uf_idx < 2:
            return 0
        fim = 0
        for i in range(1, uf_idx - 1):
            if self._eh_parte_logradouro(partes[i], continua_logradouro=True):
                fim = i
            else:
                break
        return fim

    @staticmethod
    def _formatar_cep(texto):
        if not texto or str(texto).strip() in ('', 'N/A', '#N/A'):
            return "N/A"
        digits = re.sub(r'\D', '', str(texto))
        if len(digits) == 8:
            return f"{digits[:5]}-{digits[5:]}"
        return str(texto)

    def _parsear_endereco_entrega(self, texto):
        """
        Interpreta o campo Endereço Entrega: logradouro - bairro - cidade - UF - CEP.
        Ex: Est do Campus KM 8 - Campus II - Samambaia UFG - Campus Universitário - GOIÂNIA - GO - 74690-900
        """
        resultado = {
            'logradouro': None,
            'bairro': 'N/A',
            'cidade': 'N/A',
            'cep': 'N/A',
        }
        raw = self._texto_endereco_entrega(texto)
        if not raw:
            return resultado

        cep_match = re.search(r'(\d{2})[.\s]?(\d{3})[.\s-](\d{3})', raw)
        if cep_match:
            resultado['cep'] = f"{cep_match.group(1)}{cep_match.group(2)}-{cep_match.group(3)}"
            raw = raw[:cep_match.start()].strip()

        partes = [
            p.strip().strip('-').strip()
            for p in self._normalizar_texto_endereco(raw).split(' - ')
            if p.strip()
        ]
        if not partes:
            return resultado

        uf_idx = None
        for i in range(len(partes) - 1, -1, -1):
            uf = partes[i].strip().upper()
            if len(uf) == 2 and uf in self.UFS_BRASIL:
                uf_idx = i
                break

        if uf_idx is None or uf_idx < 1:
            resultado['logradouro'] = partes[0][:50]
            return resultado

        resultado['cidade'] = partes[uf_idx - 1]
        fim_logradouro = self._indice_fim_logradouro(partes, uf_idx)
        resultado['logradouro'] = ' - '.join(partes[: fim_logradouro + 1])[:50]

        if fim_logradouro < uf_idx - 2:
            resultado['bairro'] = ' - '.join(partes[fim_logradouro + 1 : uf_idx - 1])

        return resultado

    def extrair_endereco_instalacao(self, texto):
        """Extrai o logradouro do campo Endereço Entrega (informações complementares)."""
        dados = self._parsear_endereco_entrega(texto)
        return dados['logradouro']

    def extrair_bairro_cidade_cep(self, texto):
        """Extrai bairro, cidade e CEP do campo Endereço Entrega."""
        dados = self._parsear_endereco_entrega(texto)
        return {
            'bairro': dados['bairro'],
            'cidade': dados['cidade'],
            'cep': dados['cep'],
        }
    
    def _bloco_produtos(self, texto):
        """Recorta a seção de produtos/serviços do texto da NF."""
        inicio = re.search(r'C[ÓO]D\.?\s*PROD', texto, re.IGNORECASE)
        if not inicio:
            return texto
        resto = texto[inicio.start():]
        matches = list(re.finditer(
            r'RECEBEMOS\s+DE|DADOS\s+ADICIONAIS|INFORMAÇÕES\s+COMPLEMENTARES',
            resto,
            re.IGNORECASE,
        ))
        if matches:
            fim = matches[-1]
            return resto[: fim.start()]
        return resto

    def _limpar_descricao_produto(self, trecho):
        """Remove NCM, série e dados fiscais; mantém a descrição do material."""
        trecho = re.sub(r'^\s*\d{8}\s*-\s*', '', trecho)
        trecho = re.split(
            r'N[rº]?\s*(?:de\s*)?S[ée]rie|NrodeS[ée]rie|N[ºo]?\s*Pedido\s*de\s*Compra',
            trecho,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0]
        trecho = re.split(r'\s+(?:UNI|UN)\s+\d+', trecho, maxsplit=1)[0]
        trecho = re.sub(r'\s+', ' ', trecho).strip()
        return trecho[:80] if trecho else "N/A"

    def extrair_produtos(self, texto):
        """
        Extrai os produtos (código SAP, descrição e quantidade) da nota fiscal.
        Funciona com produtos em linhas separadas ou colados na mesma linha do PDF.
        """
        bloco = self._bloco_produtos(texto)
        codigos = list(re.finditer(r'(\d{4}-\d{4}-\d+)', bloco))
        if not codigos:
            return []

        produtos = []
        for i, match in enumerate(codigos):
            fim = codigos[i + 1].start() if i + 1 < len(codigos) else len(bloco)
            trecho = bloco[match.end() : fim]

            qtd_match = re.search(r'(?:UNI|UN)\s+(\d+)\s+[\d.,]+', trecho)
            quantidade = qtd_match.group(1) if qtd_match else "0"

            produtos.append({
                'codigo': match.group(1),
                'descricao': self._limpar_descricao_produto(trecho),
                'quantidade': quantidade,
            })

        return produtos

    def extrair_produtos_software(self, texto):
        """Extrai produtos de PDFs de software (licenças/serviços com descrição e quantidade)."""
        produtos = []

        # Pattern: DESCRICAO DA LICENCA/SERVICO: <desc> - QUANTIDADE: <N>
        padrao = r'(?:DESCRI[CÇ][AÃ]O\s+DA\s+LICEN[CÇ]A/SERVIC[OA]|Descri[cç][aã]o\s+da\s+licen[cç]a/servi[cç]o)\s*:\s*(.+?)\s*-\s*(?:QUANTIDADE|Quantidade)[:\s]*(\d+)'
        for m in re.finditer(padrao, texto, re.IGNORECASE | re.DOTALL):
            desc = re.sub(r'\s+', ' ', m.group(1)).strip()[:80]
            qtd = m.group(2)
            if desc:
                produtos.append({
                    'codigo': 'N/A',
                    'descricao': desc,
                    'quantidade': qtd,
                })

        # Pattern: ITEM N - QUANTIDADE (N) DESCRIÇÃO (Format B)
        if not produtos:
            padrao_item = r'ITEM\s+\d+\s*[-:]?\s*QUANTIDADE\s*\((\d+)\)\s*(.+?)(?=\s*(?:\n|DOC\s+FATURAMENTO|ENQUADRAMENTO|CGC|$))'
            for m in re.finditer(padrao_item, texto, re.IGNORECASE | re.DOTALL):
                desc = re.sub(r'\s+', ' ', m.group(2)).strip()[:80]
                qtd = m.group(1)
                if desc:
                    produtos.append({
                        'codigo': 'N/A',
                        'descricao': desc,
                        'quantidade': qtd,
                    })

        return produtos

    def extrair_ciaus(self, texto):
        """Extrai CIAUS do texto da nota fiscal (presente em software)."""
        m = re.search(r'CIAUS\s+([A-Z]{2,})', texto)
        if m:
            return m.group(1).strip()
        return None

    def extrair_produtos_consolidados(self, texto):
        """
        Extrai produtos do texto completo e consolida por código SAP.
        Se o mesmo código SAP aparece mais de uma vez (ex: em páginas
        diferentes de um mesmo PDF), soma as quantidades em uma única linha.
        """
        produtos = self.extrair_produtos(texto)
        consolidados = {}
        for prod in produtos:
            codigo = prod['codigo']
            if codigo in consolidados:
                qtd_existente = int(consolidados[codigo]['quantidade'])
                qtd_nova = int(prod['quantidade'])
                consolidados[codigo]['quantidade'] = str(qtd_existente + qtd_nova)
            else:
                consolidados[codigo] = prod
        return list(consolidados.values())

    def ler_pdf(self, caminho_pdf):
        """Lê um arquivo PDF e extrai todo o texto."""
        try:
            with open(caminho_pdf, 'rb') as arquivo:
                leitor = PdfReader(arquivo)
                texto = ""
                for pagina in leitor.pages:
                    texto += pagina.extract_text()
                return texto
        except Exception as e:
            print(f" Erro ao ler {caminho_pdf}: {e}")
            return ""

    def contar_paginas(self, caminho_pdf):
        """Retorna o número de páginas de um arquivo PDF."""
        try:
            with open(caminho_pdf, 'rb') as arquivo:
                leitor = PdfReader(arquivo)
                return len(leitor.pages)
        except Exception as e:
            print(f" Erro ao contar páginas de {caminho_pdf}: {e}")
            return 0
    
    def exibir_preview_planilha(self):
        """Exibe um preview da planilha no terminal usando tabela formatada."""
        table = Table(title=" PREVIEW DA PLANILHA", style="cyan")
        
        table.add_column("Arquivo", justify="left", style="blue", width=22)
        table.add_column("NF", justify="center", style="blue", width=12)
        table.add_column("Tipo", justify="center", style="blue", width=10)
        table.add_column("Págs", justify="center", style="blue", width=5)
        table.add_column("Cliente", justify="left", style="blue", width=22)
        table.add_column("Data", justify="center", style="blue", width=12)
        table.add_column("UF", justify="center", style="blue", width=5)
        table.add_column("Valor", justify="right", style="blue", width=12)
        table.add_column("SAP", justify="center", style="blue", width=12)
        table.add_column("Descrição", justify="left", style="blue", width=20)
        table.add_column("Qtde", justify="center", style="blue", width=6)
        table.add_column("CGC", justify="center", style="blue", width=10)
        table.add_column("CIAUS", justify="center", style="blue", width=10)
        table.add_column("Field", justify="center", style="blue", width=10)
        table.add_column("Endereço", justify="left", style="blue", width=20)
        table.add_column("Bairro", justify="center", style="blue", width=12)
        table.add_column("Cidade", justify="center", style="blue", width=14)
        table.add_column("CEP", justify="center", style="blue", width=10)
        
        for dado in self.dados_nf:
            arquivo_curto = dado['arquivo'][:20] + "..." if len(dado['arquivo']) > 20 else dado['arquivo']
            descricao_curta = dado['descricao_material'][:18] + "..." if len(dado['descricao_material']) > 18 else dado['descricao_material']
            endereco_curto = dado['endereco_instalacao'][:18] + "..." if len(dado['endereco_instalacao']) > 18 else dado['endereco_instalacao']
            
            table.add_row(
                arquivo_curto,
                str(dado['numero_nf']),
                str(dado['tipo_nota_fiscal']),
                str(dado['qtde_paginas']),
                str(dado['cliente']),
                str(dado['data_emissao']),
                str(dado['uf_destino']),
                f"R$ {dado['valor_total']}",
                str(dado['codigo_sap']),
                descricao_curta,
                str(dado['quantidade']),
                str(dado['cgc']),
                str(dado.get('ciaus', 'N/A')),
                str(dado.get('field_responsavel', 'N/A')),
                endereco_curto,
                str(dado['bairro']),
                str(dado['cidade']),
                str(dado['cep'])
            )
        
        self.console.print(table)
    
    def processar_pdfs(self, diretorios=None):
        """Processa todos os PDFs em um ou mais diretórios (e subdiretórios)."""
        if diretorios is None:
            diretorios = ["."]
        if isinstance(diretorios, str):
            diretorios = [diretorios]

        arquivos_pdf = []
        for d in diretorios:
            path = Path(d)
            if not path.is_dir():
                print(f"  Diretório não encontrado: {d}")
                continue
            arquivos_pdf.extend(sorted(path.rglob("*.pdf")))

        if not arquivos_pdf:
            print("  Nenhum PDF encontrado nos diretórios especificados")
            return

        nfs_processadas = set()

        print(f" Encontrados {len(arquivos_pdf)} arquivos PDF")

        for caminho_pdf in arquivos_pdf:
            nome_arquivo = str(caminho_pdf)
            print(f"\n Processando: {nome_arquivo}")

            texto = self.ler_pdf(caminho_pdf)
            qtde_paginas = self.contar_paginas(caminho_pdf)
            if not texto:
                continue

            numero_nf = self.extrair_numero_nf(texto)

            tipo_nf = self._identificar_tipo_nf(numero_nf) if numero_nf else None

            # Pula NFs que não são hardware nem software
            if numero_nf and not tipo_nf:
                print(f"   NF {numero_nf} não reconhecida (pulando)")
                continue

            # Rastreia caminhos para relatório de duplicatas
            if numero_nf:
                if numero_nf not in self._nf_para_arquivos:
                    self._nf_para_arquivos[numero_nf] = []
                self._nf_para_arquivos[numero_nf].append(nome_arquivo)

            # Pula duplicatas da mesma NF (ex: backup em subpastas)
            if numero_nf and numero_nf in nfs_processadas:
                print(f"   NF {numero_nf} já processada (pulando duplicata)")
                continue
            if numero_nf:
                nfs_processadas.add(numero_nf)

            cliente = self.extrair_cliente(texto)
            data_emissao = self.extrair_data_emissao(texto)
            uf_destino = self.extrair_uf_destino(texto)
            valor_total = self.extrair_valor_total(texto)

            # Extrai produtos conforme o tipo
            if tipo_nf == "Software":
                produtos = self.extrair_produtos_software(texto)
            else:
                produtos = self.extrair_produtos_consolidados(texto)

            # Extrai informações complementares
            cgc = self.extrair_cgc(texto)
            ciaus = self.extrair_ciaus(texto) if tipo_nf == "Software" else None
            endereco = self.extrair_endereco_instalacao(texto)
            endereco_dados = self.extrair_bairro_cidade_cep(texto)

            if numero_nf:
                print(f"    NF encontrada: {numero_nf}")
                if data_emissao:
                    print(f"    Data de emissão: {data_emissao}")
                else:
                    print(f"     Data de emissão não encontrada")
                if uf_destino:
                    print(f"    UF de destino: {uf_destino}")
                else:
                    print(f"     UF de destino não encontrada")
                if valor_total:
                    print(f"    Valor total: R$ {valor_total}")
                else:
                    print(f"     Valor total não encontrado")
                
                print(f"    Produtos encontrados: {len(produtos)}")
                
                # Extrai informações complementares
                if cgc:
                    print(f"    CGC: {cgc}")
                if endereco:
                    print(f"    Endereço: {endereco}")
                
                # Se houver produtos, cria uma linha para cada produto
                if produtos:
                    for produto in produtos:
                        self.dados_nf.append({
                            'arquivo': nome_arquivo,
                            'numero_nf': numero_nf,
                            'tipo_nota_fiscal': tipo_nf,
                            'qtde_paginas': qtde_paginas,
                            'cliente': cliente or "N/A",
                            'data_emissao': data_emissao or "N/A",
                            'uf_destino': uf_destino or "N/A",
                            'valor_total': valor_total or "N/A",
                            'codigo_sap': produto['codigo'],
                            'descricao_material': produto['descricao'],
                            'quantidade': produto['quantidade'],
                            'cgc': cgc or "N/A",
                            'ciaus': ciaus or "N/A",
                            'endereco_instalacao': endereco or "N/A",
                            'bairro': endereco_dados['bairro'],
                            'cidade': endereco_dados['cidade'],
                            'cep': endereco_dados['cep']
                        })
                else:
                    # Se não houver produtos, cria apenas com os dados da NF
                        self.dados_nf.append({
                            'arquivo': nome_arquivo,
                            'numero_nf': numero_nf,
                            'tipo_nota_fiscal': tipo_nf,
                            'qtde_paginas': qtde_paginas,
                            'cliente': cliente or "N/A",
                            'data_emissao': data_emissao or "N/A",
                            'uf_destino': uf_destino or "N/A",
                            'valor_total': valor_total or "N/A",
                            'codigo_sap': "N/A",
                            'descricao_material': "N/A",
                            'quantidade': "0",
                            'cgc': cgc or "N/A",
                            'ciaus': ciaus or "N/A",
                            'endereco_instalacao': endereco or "N/A",
                            'bairro': endereco_dados['bairro'],
                            'cidade': endereco_dados['cidade'],
                            'cep': endereco_dados['cep']
                        })
            else:
                print(f"     NF não encontrada neste documento")
        
        # Relatório de duplicatas
        duplicatas = {nf: paths for nf, paths in self._nf_para_arquivos.items() if len(paths) > 1}
        if duplicatas:
            print(f"\n Arquivos duplicados encontrados: {sum(len(v) for v in duplicatas.values())} ocorrências de {len(duplicatas)} NF(s)")
            print(f"   Detalhes na aba 'Duplicatas' da planilha")
    
    def criar_aba_duplicatas(self, wb):
        """Adiciona uma aba 'Duplicatas' com arquivos repetidos (mesma NF em múltiplos caminhos)."""
        duplicatas = {nf: paths for nf, paths in self._nf_para_arquivos.items() if len(paths) > 1}
        if not duplicatas:
            return

        ws = wb.create_sheet("Duplicatas")

        fill_cinza = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        fonte_azul = Font(color="000080", bold=True, size=11)
        fonte_dados = Font(color="000080", size=10)
        align_center = Alignment(horizontal="center", vertical="center")

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 50

        headers = ["Número da NF", "Caminhos dos Arquivos"]
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.value = h
            cell.fill = fill_cinza
            cell.font = fonte_azul
            cell.alignment = align_center

        row = 2
        for nf in sorted(duplicatas.keys()):
            for path in duplicatas[nf]:
                ws.cell(row=row, column=1, value=int(nf)).font = fonte_dados
                ws.cell(row=row, column=1).alignment = align_center
                ws.cell(row=row, column=2, value=path).font = fonte_dados
                row += 1

        ws.sheet_view.showGridLines = False
        borda_celula = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )
        for linha in range(1, row):
            for coluna in range(1, len(headers) + 1):
                ws.cell(row=linha, column=coluna).border = borda_celula

        ws.auto_filter.ref = f"A1:B{row - 1}"
        ws.freeze_panes = 'A2'
        ws.sheet_view.tabColor = "ED7D31"

    def _valor_para_float(self, valor):
        """Converte valor monetário brasileiro (ex: 6.784,16) para float."""
        if not valor or valor == "N/A":
            return None
        try:
            valor_limpo = str(valor).replace('.', '').replace(',', '.')
            return round(float(valor_limpo), 2)
        except ValueError:
            return None
    
    def criar_planilha_excel(self):
        """Cria uma planilha Excel com os dados extraídos."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Notas Fiscais"
        
        # Define largura das colunas
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 15
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 22
        ws.column_dimensions['I'].width = 15
        ws.column_dimensions['J'].width = 35
        ws.column_dimensions['K'].width = 12
        ws.column_dimensions['L'].width = 15
        ws.column_dimensions['M'].width = 15
        ws.column_dimensions['N'].width = 20
        ws.column_dimensions['O'].width = 30
        ws.column_dimensions['P'].width = 15
        ws.column_dimensions['Q'].width = 15
        ws.column_dimensions['R'].width = 12
        
        # Estilo do cabeçalho: fundo cinza e texto azul marinho
        fill_cinza = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        fonte_azul = Font(color="000080", bold=True, size=11)
        alignment_centralizado = Alignment(horizontal="center", vertical="center")
        alignment_esquerda = Alignment(horizontal="left", vertical="center")
        
        # Adiciona cabeçalhos
        headers = ["Nome Original do Arquivo", "Número da NF", "Tipo de Nota Fiscal", "Qtde Páginas PDF", "Cliente", "Data de Emissão", "UF de Destino", 
                   "Valor Total da NF", "Código SAP", "Descrição do Material", "Quantidade", 
                   "CGC", "CIAUS", "Field Responsável", "Endereço de Instalação", "Bairro", "Cidade", "CEP"]
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = fill_cinza
            cell.font = fonte_azul
            cell.alignment = alignment_centralizado
        
        # Adiciona dados
        fonte_dados = Font(color="000080", size=10)
        alignment_direita = Alignment(horizontal="right", vertical="center")
        
        for idx, dado in enumerate(self.dados_nf, start=2):
            # Nome Original do Arquivo
            ws[f'A{idx}'] = dado['arquivo']
            ws[f'A{idx}'].font = fonte_dados
            ws[f'A{idx}'].alignment = alignment_centralizado
            
            # Número da NF
            try:
                ws[f'B{idx}'] = int(dado['numero_nf'])
            except (ValueError, TypeError):
                ws[f'B{idx}'] = dado['numero_nf']
            ws[f'B{idx}'].font = fonte_dados
            ws[f'B{idx}'].alignment = alignment_centralizado
            
            # Tipo de Nota Fiscal
            ws[f'C{idx}'] = dado['tipo_nota_fiscal']
            ws[f'C{idx}'].font = fonte_dados
            ws[f'C{idx}'].alignment = alignment_centralizado
            
            # Qtde Páginas PDF
            ws[f'D{idx}'] = dado['qtde_paginas']
            ws[f'D{idx}'].font = fonte_dados
            ws[f'D{idx}'].alignment = alignment_centralizado
            
            # Cliente
            ws[f'E{idx}'] = dado['cliente']
            ws[f'E{idx}'].font = fonte_dados
            ws[f'E{idx}'].alignment = alignment_centralizado
            
            # Data de Emissão
            ws[f'F{idx}'] = dado['data_emissao']
            ws[f'F{idx}'].font = fonte_dados
            ws[f'F{idx}'].alignment = alignment_centralizado
            
            # UF de Destino
            ws[f'G{idx}'] = dado['uf_destino']
            ws[f'G{idx}'].font = fonte_dados
            ws[f'G{idx}'].alignment = alignment_centralizado
            
            # Valor Total (formatação Contábil do Excel)
            celula_valor = ws[f'H{idx}']
            valor_num = self._valor_para_float(dado['valor_total'])
            if valor_num is not None:
                celula_valor.value = round(valor_num, 2)
                celula_valor.number_format = self.FORMATO_CONTABIL_BR
            else:
                celula_valor.value = "N/A"
            celula_valor.font = fonte_dados
            celula_valor.alignment = alignment_direita
            
            # Código SAP
            ws[f'I{idx}'] = dado['codigo_sap']
            ws[f'I{idx}'].font = fonte_dados
            ws[f'I{idx}'].alignment = alignment_centralizado
            
            # Descrição do Material
            ws[f'J{idx}'] = dado['descricao_material']
            ws[f'J{idx}'].font = fonte_dados
            ws[f'J{idx}'].alignment = alignment_esquerda
            
            # Quantidade
            try:
                ws[f'K{idx}'] = int(dado['quantidade'])
            except (ValueError, TypeError):
                ws[f'K{idx}'] = dado['quantidade']
            ws[f'K{idx}'].font = fonte_dados
            ws[f'K{idx}'].alignment = alignment_centralizado
            
            # CGC
            if dado['cgc'] != "N/A":
                try:
                    ws[f'L{idx}'] = int(dado['cgc'])
                except (ValueError, TypeError):
                    ws[f'L{idx}'] = dado['cgc']
            else:
                ws[f'L{idx}'] = "N/A"
            ws[f'L{idx}'].font = fonte_dados
            ws[f'L{idx}'].alignment = alignment_centralizado
            
            # CIAUS
            ws[f'M{idx}'] = dado.get('ciaus', "N/A")
            ws[f'M{idx}'].font = fonte_dados
            ws[f'M{idx}'].alignment = alignment_centralizado
            
            # Field Responsável
            ws[f'N{idx}'] = dado.get('field_responsavel', "N/A")
            ws[f'N{idx}'].font = fonte_dados
            ws[f'N{idx}'].alignment = alignment_centralizado
            
            # Endereço de Instalação
            ws[f'O{idx}'] = dado['endereco_instalacao']
            ws[f'O{idx}'].font = fonte_dados
            ws[f'O{idx}'].alignment = alignment_centralizado
            
            # Bairro
            ws[f'P{idx}'] = dado['bairro']
            ws[f'P{idx}'].font = fonte_dados
            ws[f'P{idx}'].alignment = alignment_centralizado
            
            # Cidade
            ws[f'Q{idx}'] = dado['cidade']
            ws[f'Q{idx}'].font = fonte_dados
            ws[f'Q{idx}'].alignment = alignment_centralizado
            
            # CEP (formatado como 00000-000)
            ws[f'R{idx}'] = self._formatar_cep(dado['cep'])
            ws[f'R{idx}'].font = fonte_dados
            ws[f'R{idx}'].alignment = alignment_centralizado
        
        # Grade só na área com conteúdo; restante da planilha sem linhas
        ws.sheet_view.showGridLines = False
        borda_celula = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )
        ultima_linha = 1 + len(self.dados_nf)
        ultima_coluna = len(headers)
        for linha in range(1, ultima_linha + 1):
            for coluna in range(1, ultima_coluna + 1):
                ws.cell(row=linha, column=coluna).border = borda_celula

        ws.auto_filter.ref = f"A1:R{ultima_linha}"
        ws.freeze_panes = 'A2'
        ws.sheet_view.tabColor = "4472C4"

        # Aba de duplicatas
        self.criar_aba_duplicatas(wb)

        # Salva a planilha
        try:
            wb.save(self.excel_path)
            print(f"\n Planilha criada com sucesso: {self.excel_path}")
        except Exception as e:
            print(f" Erro ao salvar planilha: {e}")
    
    def _recarregar_dados_cruzamento(self):
        """Recarrega CGC, CIAUS e Field do Excel após o cruzamento para atualizar self.dados_nf."""
        wb = load_workbook(self.excel_path)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        col_map = {}
        for i, h in enumerate(headers):
            col_map[h] = i
        for idx, dado in enumerate(self.dados_nf, start=2):
            row = ws[idx]
            if 'CGC' in col_map:
                val = row[col_map['CGC']].value
                if val is not None and str(val).strip() not in ('', 'N/A', '0'):
                    dado['cgc'] = str(val)
            if 'CIAUS' in col_map:
                val = row[col_map['CIAUS']].value
                if val is not None and str(val).strip():
                    dado['ciaus'] = str(val)
            if 'Field Responsável' in col_map:
                val = row[col_map['Field Responsável']].value
                if val is not None and str(val).strip():
                    dado['field_responsavel'] = str(val)
        wb.close()

    def executar(self, diretorios=None, caminho_cruzamento=None):
        """Executa o fluxo completo de leitura, geração da planilha e cruzamento."""
        print("=" * 50)
        print(" LEITOR DE PDF - Notas Fiscais")
        print("=" * 50)
        
        self.processar_pdfs(diretorios)
        
        if self.dados_nf:
            print(f"\n Total de NFs extraídas: {len(self.dados_nf)}\n")
            self.criar_planilha_excel()
            
            # Etapa de cruzamento
            if caminho_cruzamento:
                print("\n" + "=" * 50)
                print(" Iniciando cruzamento de dados...")
                print("=" * 50)
                try:
                    CruzamentoDados.executar(
                        self.excel_path,
                        caminho_cruzamento,
                    )
                    print("\n Cruzamento de dados concluído!")
                    print("   Colunas adicionadas: CGC, CIAUS, Field Responsável")
                except Exception as e:
                    print(f"\n  Erro no cruzamento: {e}")
                finally:
                    self._recarregar_dados_cruzamento()
            
            self.exibir_preview_planilha()
        else:
            print("\n  Nenhuma nota fiscal foi encontrada nos PDFs.")


class CruzamentoDados:
    """Cruzamento entre planilha gerada e aba Base do SDLAN CEF."""

    PESOS = {'endereco': 25, 'bairro': 20, 'cidade': 25, 'cep': 30}

    COLUNAS_BASE = {
        'endereco': ['endereço', 'endereco', 'logradouro', 'end'],
        'bairro': ['bairro'],
        'cidade': ['cidade', 'município', 'municipio'],
        'cep': ['cep', 'postal', 'zip'],
        'cgc_unidade': ['cgc unidade', 'cgc da unidade', 'cgc unid'],
        'ciaus': ['ciaus'],
        'field': ['field responsável', 'field'],
    }

    COLUNAS_GERADO = {
        'endereco': 'endereço de instalação',
        'bairro': 'bairro',
        'cidade': 'cidade',
        'cep': 'cep',
        'cgc': 'cgc',
        'ciaus': 'ciaus',
        'field': 'field responsável',
    }

    _RE_ACENTOS = re.compile(r'[àáâãäå]', re.I)
    _RE_ACENTOS_E = re.compile(r'[èéêë]', re.I)
    _RE_ACENTOS_I = re.compile(r'[ìíîï]', re.I)
    _RE_ACENTOS_O = re.compile(r'[òóôõö]', re.I)
    _RE_ACENTOS_U = re.compile(r'[ùúûü]', re.I)
    _RE_ACENTOS_C = re.compile(r'[ç]', re.I)

    _ABREVIACOES = {
        'r': 'rua', 'av': 'avenida', 'est': 'estrada', 'rod': 'rodovia',
        'pca': 'praca', 'trav': 'travessa', 'al': 'alameda',
    }

    @staticmethod
    def _normalizar(texto):
        if not texto:
            return ""
        s = str(texto)
        # Remove pontuação (vírgula, ponto, hífen, barra, parênteses etc)
        s = re.sub(r'[,./\-\\()\[\]{}:;!?@#&*+=_~<>"\'°]', ' ', s)
        s = CruzamentoDados._RE_ACENTOS.sub('a', s)
        s = CruzamentoDados._RE_ACENTOS_E.sub('e', s)
        s = CruzamentoDados._RE_ACENTOS_I.sub('i', s)
        s = CruzamentoDados._RE_ACENTOS_O.sub('o', s)
        s = CruzamentoDados._RE_ACENTOS_U.sub('u', s)
        s = CruzamentoDados._RE_ACENTOS_C.sub('c', s)
        s = re.sub(r'\s+', ' ', s.lower().strip())
        # Expande abreviações comuns de endereço
        tokens = s.split()
        tokens = [CruzamentoDados._ABREVIACOES.get(t, t) for t in tokens]
        return ' '.join(tokens)

    @staticmethod
    def _normalizar_cep(texto):
        if not texto:
            return ""
        return re.sub(r'\D', '', str(texto))

    @staticmethod
    def _formatar_cep(texto):
        if not texto or str(texto).strip() in ('', 'N/A', '#N/A'):
            return "N/A"
        digits = re.sub(r'\D', '', str(texto))
        if len(digits) == 8:
            return f"{digits[:5]}-{digits[5:]}"
        return str(texto)

    @staticmethod
    def _encontrar_coluna(headers, padroes):
        headers_lower = [str(h).lower().strip() if h else "" for h in headers]
        for i, h in enumerate(headers_lower):
            for padrao in padroes:
                if padrao in h:
                    return i
        return None

    @classmethod
    def _mapear_colunas_base(cls, headers):
        mapeamento = {}
        for campo, padroes in cls.COLUNAS_BASE.items():
            idx = cls._encontrar_coluna(headers, padroes)
            if idx is not None:
                mapeamento[campo] = idx
        return mapeamento

    @classmethod
    def _mapear_colunas_gerado(cls, headers):
        mapeamento = {}
        for campo, nome_busca in cls.COLUNAS_GERADO.items():
            for i, h in enumerate(headers):
                if h and nome_busca in str(h).lower().strip():
                    mapeamento[campo] = i
                    break
        return mapeamento

    @staticmethod
    def _ratio_tokens(a, b):
        tokens_a = set(a.split())
        tokens_b = set(b.split())
        if not tokens_a or not tokens_b:
            return 0.0
        inter = tokens_a & tokens_b
        return len(inter) / max(len(tokens_a), len(tokens_b))

    @staticmethod
    def _score_match(linha, ref, pesos):
        score = 0.0
        total_peso = 0.0
        for campo, peso in pesos.items():
            val_linha = linha.get(campo, "")
            val_ref = ref.get(campo, "")
            if not val_linha or not val_ref:
                continue
            if campo == 'cep':
                if val_linha == val_ref:
                    score += peso
                total_peso += peso
            else:
                score += CruzamentoDados._ratio_tokens(val_linha, val_ref) * peso
                total_peso += peso
        return score / total_peso if total_peso > 0 else 0.0

    @classmethod
    def executar(cls, caminho_excel_gerado, caminho_cruzamento, progress_callback=None):
        wb = load_workbook(caminho_excel_gerado)
        ws = wb.active
        headers_gerado = [cell.value for cell in ws[1]]
        cols_gerado = cls._mapear_colunas_gerado(headers_gerado)

        wb_base = load_workbook(caminho_cruzamento, data_only=True, read_only=True)
        if 'Base' not in wb_base.sheetnames:
            raise ValueError(f"Aba 'Base' nao encontrada. Abas: {wb_base.sheetnames}")
        ws_base = wb_base['Base']
        headers_base = [cell.value for cell in ws_base[1]]
        cols_base = cls._mapear_colunas_base(headers_base)

        obrigatorias_base = ['endereco', 'bairro', 'cidade', 'cep', 'cgc_unidade']
        for col in obrigatorias_base:
            if col not in cols_base:
                raise ValueError(f"Coluna '{col}' nao encontrada na aba Base. Cabecalhos: {headers_base}")

        obrigatorias_gerado = ['endereco', 'bairro', 'cidade', 'cep', 'cgc']
        for col in obrigatorias_gerado:
            if col not in cols_gerado:
                raise ValueError(f"Coluna '{col}' nao encontrada na planilha gerada. Cabecalhos: {headers_gerado}")

        ref_por_cidade = {}
        ref_por_cgc = {}
        ref_por_cep = {}
        idx_cgc_unid = cols_base['cgc_unidade']
        idx_ciaus = cols_base.get('ciaus')
        idx_field = cols_base.get('field')

        for row in ws_base.iter_rows(min_row=2, values_only=True):
            cgc_unidade = row[idx_cgc_unid] if idx_cgc_unid < len(row) else None
            if cgc_unidade is None or str(cgc_unidade).strip() in ('', 'N/A', '#N/A'):
                continue
            try:
                cgc_int = int(float(str(cgc_unidade).replace(',', '.')))
            except (ValueError, TypeError, OverflowError):
                continue

            end = cls._normalizar(row[cols_base['endereco']]) if cols_base['endereco'] < len(row) else ""
            bai = cls._normalizar(row[cols_base['bairro']]) if cols_base['bairro'] < len(row) else ""
            cid = cls._normalizar(row[cols_base['cidade']]) if cols_base['cidade'] < len(row) else ""
            cep_raw = row[cols_base['cep']] if cols_base.get('cep', 0) < len(row) else ""
            cep = cls._normalizar_cep(cep_raw)

            if not end and not bai and not cid and not cep:
                continue

            ciaus_val = str(row[idx_ciaus]).strip() if idx_ciaus is not None and idx_ciaus < len(row) and row[idx_ciaus] is not None else ""
            field_val = str(row[idx_field]).strip() if idx_field is not None and idx_field < len(row) and row[idx_field] is not None else ""

            registro = {
                'endereco': end, 'bairro': bai, 'cidade': cid, 'cep': cep,
                'cgc': cgc_int, 'ciaus': ciaus_val, 'field': field_val,
            }

            chave_cidade = cid if cid else "_sem_cidade"
            if chave_cidade not in ref_por_cidade:
                ref_por_cidade[chave_cidade] = []
            ref_por_cidade[chave_cidade].append(registro)

            # Também indexa por CGC para lookup direto
            if cgc_int not in ref_por_cgc:
                ref_por_cgc[cgc_int] = registro

            # Indexa por prefixo CEP (5, 4 e 3 dígitos) para fallback progressivo
            for cep_len in (5, 4, 3):
                if len(cep) >= cep_len:
                    cep_prefix = cep[:cep_len]
                    if cep_prefix not in ref_por_cep:
                        ref_por_cep[cep_prefix] = []
                    ref_por_cep[cep_prefix].append(registro)

        if not ref_por_cidade:
            raise ValueError("Nenhuma linha com CGC Unidade valido na aba Base.")

        idx_end = cols_gerado['endereco']
        idx_bai = cols_gerado['bairro']
        idx_cid = cols_gerado['cidade']
        idx_cep = cols_gerado['cep']
        idx_cgc = cols_gerado['cgc']
        idx_ciaus = cols_gerado.get('ciaus')
        idx_field = cols_gerado.get('field')

        fonte_dados = Font(color="000080", size=10)
        alignment_centralizado = Alignment(horizontal="center", vertical="center")

        total_linhas = ws.max_row - 1
        substituidos_cgc = 0
        preenchidos_ciaus_field = 0
        preenchidos_endereco_cgc = 0

        cache_grupo = {}

        for row_idx in range(2, ws.max_row + 1):
            if progress_callback:
                progress_callback(row_idx - 2, total_linhas)

            cgc_atual = ws.cell(row_idx, idx_cgc + 1).value

            end_linha = cls._normalizar(ws.cell(row_idx, idx_end + 1).value)
            bai_linha = cls._normalizar(ws.cell(row_idx, idx_bai + 1).value)
            cid_linha = cls._normalizar(ws.cell(row_idx, idx_cid + 1).value)
            cep_linha = cls._normalizar_cep(ws.cell(row_idx, idx_cep + 1).value)

            if not end_linha and not bai_linha and not cid_linha and not cep_linha:
                continue

            linha_atual = {'endereco': end_linha, 'bairro': bai_linha, 'cidade': cid_linha, 'cep': cep_linha}

            candidatos = ref_por_cidade.get(cid_linha, [])
            if not candidatos:
                candidatos = [r for grupo in ref_por_cidade.values() for r in grupo]

            melhor_score = 0.0
            melhor_dados = None
            for ref in candidatos:
                score = cls._score_match(linha_atual, ref, cls.PESOS)
                if score > melhor_score:
                    melhor_score = score
                    melhor_dados = ref

            if melhor_dados and melhor_score >= 0.15:
                cgc_precisa_substituir = (cgc_atual is None or str(cgc_atual).strip() in ('N/A', '', '0'))
                if not cgc_precisa_substituir:
                    try:
                        cgc_int = int(float(str(cgc_atual).replace(',', '.')))
                        if cgc_int not in ref_por_cgc:
                            cgc_precisa_substituir = True
                    except (ValueError, TypeError, OverflowError):
                        cgc_precisa_substituir = True
                if cgc_precisa_substituir:
                    ws.cell(row_idx, idx_cgc + 1).value = melhor_dados['cgc']
                    substituidos_cgc += 1

                if idx_ciaus is not None and melhor_dados['ciaus']:
                    ws.cell(row_idx, idx_ciaus + 1).value = melhor_dados['ciaus']
                    ws.cell(row_idx, idx_ciaus + 1).font = fonte_dados
                    ws.cell(row_idx, idx_ciaus + 1).alignment = alignment_centralizado
                    preenchidos_ciaus_field += 1
                if idx_field is not None and melhor_dados['field']:
                    ws.cell(row_idx, idx_field + 1).value = melhor_dados['field']
                    ws.cell(row_idx, idx_field + 1).font = fonte_dados
                    ws.cell(row_idx, idx_field + 1).alignment = alignment_centralizado

            # Fallback por CEP: se não achou por endereço, tenta por prefixo progressivo do CEP
            if not (melhor_dados and melhor_score >= 0.15) and cep_linha and len(cep_linha) >= 5:
                melhor_cep_dados = None
                melhor_cep_score = 0.0
                for cep_len in (5, 4, 3):
                    cep_prefix = cep_linha[:cep_len]
                    candidatos_cep = ref_por_cep.get(cep_prefix, [])
                    if not candidatos_cep:
                        continue
                    for cand in candidatos_cep:
                        s = cls._score_match(linha_atual, cand, cls.PESOS)
                        if s > melhor_cep_score:
                            melhor_cep_score = s
                            melhor_cep_dados = cand
                    if melhor_cep_dados and melhor_cep_score >= 0.15:
                        break
                if melhor_cep_dados:
                    cgc_precisa_substituir = (cgc_atual is None or str(cgc_atual).strip() in ('N/A', '', '0'))
                    if not cgc_precisa_substituir:
                        try:
                            cgc_int = int(float(str(cgc_atual).replace(',', '.')))
                            if cgc_int not in ref_por_cgc:
                                cgc_precisa_substituir = True
                        except (ValueError, TypeError, OverflowError):
                            cgc_precisa_substituir = True
                    if cgc_precisa_substituir:
                        ws.cell(row_idx, idx_cgc + 1).value = melhor_cep_dados['cgc']
                        substituidos_cgc += 1
                    if idx_ciaus is not None and melhor_cep_dados['ciaus']:
                        ws.cell(row_idx, idx_ciaus + 1).value = melhor_cep_dados['ciaus']
                        ws.cell(row_idx, idx_ciaus + 1).font = fonte_dados
                        ws.cell(row_idx, idx_ciaus + 1).alignment = alignment_centralizado
                        preenchidos_ciaus_field += 1
                    if idx_field is not None and melhor_cep_dados['field']:
                        ws.cell(row_idx, idx_field + 1).value = melhor_cep_dados['field']
                        ws.cell(row_idx, idx_field + 1).font = fonte_dados
                        ws.cell(row_idx, idx_field + 1).alignment = alignment_centralizado

        # Segundo passe: busca por CGC para preencher endereço/bairro/cidade/CEP faltantes
        for row_idx in range(2, ws.max_row + 1):
            cgc_val = ws.cell(row_idx, idx_cgc + 1).value
            if cgc_val is None or str(cgc_val).strip() in ('', 'N/A', '0'):
                continue
            try:
                cgc_int = int(float(str(cgc_val).replace(',', '.')))
            except (ValueError, TypeError, OverflowError):
                continue

            base_row = ref_por_cgc.get(cgc_int)
            if base_row is None:
                continue

            cell_end = ws.cell(row_idx, idx_end + 1)
            if cell_end.value is None or str(cell_end.value).strip() in ('', 'N/A'):
                cell_end.value = base_row['endereco'].upper()
                cell_end.font = fonte_dados
                preenchidos_endereco_cgc += 1

            cell_bai = ws.cell(row_idx, idx_bai + 1)
            if cell_bai.value is None or str(cell_bai.value).strip() in ('', 'N/A'):
                cell_bai.value = base_row['bairro'].upper()
                cell_bai.font = fonte_dados
                preenchidos_endereco_cgc += 1

            cell_cid = ws.cell(row_idx, idx_cid + 1)
            if cell_cid.value is None or str(cell_cid.value).strip() in ('', 'N/A'):
                cell_cid.value = base_row['cidade'].upper()
                cell_cid.font = fonte_dados
                preenchidos_endereco_cgc += 1

            cell_cep = ws.cell(row_idx, idx_cep + 1)
            if cell_cep.value is None or str(cell_cep.value).strip() in ('', 'N/A'):
                cell_cep.value = cls._formatar_cep(base_row['cep'])
                cell_cep.font = fonte_dados
                cell_cep.alignment = alignment_centralizado
                preenchidos_endereco_cgc += 1

        wb.save(caminho_excel_gerado)
        print(f"\n   Cruzamento concluido: {substituidos_cgc} CGC(s) preenchido(s), "
              f"{preenchidos_ciaus_field} linha(s) com CIAUS/Field preenchidos, "
              f"{preenchidos_endereco_cgc} campo(s) de endereço preenchidos por CGC")
        return True


if __name__ == "__main__":
    # Cria instância e executa
    leitor = LeitordePDF()
    
    # Procura automaticamente pela planilha Base para cruzamento
    base_files = sorted(Path(".").glob("Base do Projeto SDLAN CEF*.xlsx"))
    caminho_cruzamento = str(base_files[0]) if base_files else None
    if caminho_cruzamento:
        print(f" Planilha Base encontrada: {base_files[0].name}")
    
    # Processa PDFs e faz cruzamento se houver Base
    leitor.executar(diretorios=["."], caminho_cruzamento=caminho_cruzamento)
