#Código contruido por André Souza 100% no Ipad

"""
Leitor de PDF - Extrai dados de notas fiscais em PDFs
e gera uma planilha Excel formatada.
"""

import re
from pathlib import Path
from PyPDF2 import PdfReader
from openpyxl import Workbook
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
        self.excel_path = "Leitura das Notas Fiscais.xlsx"
        self.console = Console()
    
    def extrair_numero_nf(self, texto):
        """
        Extrai o número da nota fiscal do texto.
        Procura por padrões como "Nº", "NF-e", "Nota Fiscal" seguido de números.
        """
        # Padrões comuns para nota fiscal
        padoes = [
            r'Nº\s+(\d+)',  # Nº 000074734
            r'N[oº]\s+(\d+)',  # No ou Nº seguido de número
            r'NF-?[e]?\s+(\d+)',  # NF-e ou NFe
            r'Nota\s+Fiscal\s+(\d+)',  # Nota Fiscal
            r'NF\s*[#:-]?\s*(\d+)',  # NF com separadores
        ]
        
        for padrao in padoes:
            match = re.search(padrao, texto, re.IGNORECASE)
            if match:
                numero = match.group(1)
                # Remove zeros a esquerda
                numero_limpo = str(int(numero))
                return numero_limpo
        
        return None
    
    def extrair_cliente(self, texto):
        """
        Extrai o nome do cliente/razão social da nota fiscal.
        Procura por "NOME/RAZÃO SOCIAL" seguido do nome do cliente.
        """
        padrao = r'NOME/RAZ[AÃ]O\s+SOCIAL\s*(.+?)\s*C\.\s*N\.\s*P\.\s*J\.'
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return None

    def extrair_data_emissao(self, texto):
        """
        Extrai a data de emissão do bloco "DATA DA EMISSÃO" (formato DD.MM.YYYY).
        O PDF pode vir com ou sem espaços entre o rótulo e a data (ex: DATADAEMISSÃO20.08.2025).
        Retorna no formato DD/MM/YYYY.
        """
        padrao = r'DATA\s*DA\s*EMISS[AÃ]O\s*(\d{2})\.(\d{2})\.(\d{4})'
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            dia, mes, ano = match.groups()
            return f"{dia}/{mes}/{ano}"
        
        return None
    
    def extrair_uf_destino(self, texto):
        """
        Extrai a UF de destino da nota fiscal.
        Procura por "FONE/FAX UF" seguido da sigla do estado.
        """
        padrao = r'FONE/FAX\s+UF\s*([A-Z]{2})'
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            return match.group(1)
        
        return None
    
    def extrair_valor_total(self, texto):
        """
        Extrai o valor total da nota fiscal.
        Procura por "VALOR TOTAL DA NOTA" seguido do valor em formato brasileiro.
        """
        padrao = r'VALOR\s+TOTAL\s+DA\s*NOTA\s*([\d.]+,\d+)'
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            return match.group(1)
        
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
            r'endereço\s+(?:de\s+)?entrega\s*:\s*(.+?)(?=\s*-?\s*Contrato\s|\s*-?\s*Conta\s+banc[aá]ria|$)',
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
        fim = re.search(
            r'RECEBEMOS\s+DE|DADOS\s+ADICIONAIS|INFORMAÇÕES\s+COMPLEMENTARES',
            resto,
            re.IGNORECASE,
        )
        return resto[: fim.start()] if fim else resto

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
            print(f"❌ Erro ao ler {caminho_pdf}: {e}")
            return ""

    def contar_paginas(self, caminho_pdf):
        """Retorna o número de páginas de um arquivo PDF."""
        try:
            with open(caminho_pdf, 'rb') as arquivo:
                leitor = PdfReader(arquivo)
                return len(leitor.pages)
        except Exception as e:
            print(f"❌ Erro ao contar páginas de {caminho_pdf}: {e}")
            return 0
    
    def exibir_preview_planilha(self):
        """Exibe um preview da planilha no terminal usando tabela formatada."""
        table = Table(title="📊 PREVIEW DA PLANILHA", style="cyan")
        
        table.add_column("Número da NF", justify="center", style="blue", width=12)
        table.add_column("Qtde Páginas", justify="center", style="blue", width=12)
        table.add_column("Cliente", justify="left", style="blue", width=22)
        table.add_column("Data", justify="center", style="blue", width=12)
        table.add_column("UF", justify="center", style="blue", width=5)
        table.add_column("Valor", justify="right", style="blue", width=12)
        table.add_column("Código SAP", justify="center", style="blue", width=12)
        table.add_column("Descrição", justify="left", style="blue", width=20)
        table.add_column("Qtde.", justify="center", style="blue", width=6)
        table.add_column("CGC", justify="center", style="blue", width=8)
        table.add_column("Endereço", justify="left", style="blue", width=20)
        table.add_column("Bairro", justify="center", style="blue", width=12)
        table.add_column("Cidade", justify="center", style="blue", width=12)
        table.add_column("CEP", justify="center", style="blue", width=10)
        
        for dado in self.dados_nf:
            valor_formatado = f"R$ {dado['valor_total']}"
            descricao_curta = dado['descricao_material'][:20] + "..." if len(dado['descricao_material']) > 20 else dado['descricao_material']
            endereco_curto = dado['endereco_instalacao'][:20] + "..." if len(dado['endereco_instalacao']) > 20 else dado['endereco_instalacao']
            
            table.add_row(
                str(dado['numero_nf']),
                str(dado['qtde_paginas']),
                str(dado['cliente']),
                str(dado['data_emissao']),
                str(dado['uf_destino']),
                valor_formatado,
                str(dado['codigo_sap']),
                descricao_curta,
                str(dado['quantidade']),
                str(dado['cgc']),
                endereco_curto,
                str(dado['bairro']),
                str(dado['cidade']),
                str(dado['cep'])
            )
        
        self.console.print(table)
    
    def processar_pdfs(self, diretorio="."):
        """Processa todos os PDFs em um diretório e subdiretórios."""
        path = Path(diretorio)
        arquivos_pdf = sorted(path.rglob("*.pdf"))
        
        if not arquivos_pdf:
            print(f"⚠️  Nenhum PDF encontrado em {diretorio}")
            return
        
        nfs_processadas = set()
        
        print(f"📄 Encontrados {len(arquivos_pdf)} arquivos PDF")
        
        for caminho_pdf in arquivos_pdf:
            nome_arquivo = str(caminho_pdf.relative_to(path))
            print(f"\n📖 Processando: {nome_arquivo}")
            
            texto = self.ler_pdf(caminho_pdf)
            qtde_paginas = self.contar_paginas(caminho_pdf)
            if not texto:
                continue
            
            numero_nf = self.extrair_numero_nf(texto)
            
            # Pula duplicatas da mesma NF (ex: backup em subpastas)
            if numero_nf and numero_nf in nfs_processadas:
                print(f"   ⏭️  NF {numero_nf} já processada (pulando duplicata)")
                continue
            if numero_nf:
                nfs_processadas.add(numero_nf)
            
            cliente = self.extrair_cliente(texto)
            data_emissao = self.extrair_data_emissao(texto)
            uf_destino = self.extrair_uf_destino(texto)
            valor_total = self.extrair_valor_total(texto)

            if qtde_paginas > 1:
                produtos = self.extrair_produtos_consolidados(texto)
            else:
                produtos = self.extrair_produtos(texto)
            
            # Extrai informações complementares
            cgc = self.extrair_cgc(texto)
            endereco = self.extrair_endereco_instalacao(texto)
            endereco_dados = self.extrair_bairro_cidade_cep(texto)
            
            if numero_nf:
                print(f"   ✓ NF encontrada: {numero_nf}")
                if data_emissao:
                    print(f"   ✓ Data de emissão: {data_emissao}")
                else:
                    print(f"   ⚠️  Data de emissão não encontrada")
                if uf_destino:
                    print(f"   ✓ UF de destino: {uf_destino}")
                else:
                    print(f"   ⚠️  UF de destino não encontrada")
                if valor_total:
                    print(f"   ✓ Valor total: R$ {valor_total}")
                else:
                    print(f"   ⚠️  Valor total não encontrado")
                
                print(f"   ✓ Produtos encontrados: {len(produtos)}")
                
                # Extrai informações complementares
                if cgc:
                    print(f"   ✓ CGC: {cgc}")
                if endereco:
                    print(f"   ✓ Endereço: {endereco}")
                
                # Se houver produtos, cria uma linha para cada produto
                if produtos:
                    for produto in produtos:
                        self.dados_nf.append({
                            'arquivo': nome_arquivo,
                            'numero_nf': numero_nf,
                            'qtde_paginas': qtde_paginas,
                            'cliente': cliente or "N/A",
                            'data_emissao': data_emissao or "N/A",
                            'uf_destino': uf_destino or "N/A",
                            'valor_total': valor_total or "N/A",
                            'codigo_sap': produto['codigo'],
                            'descricao_material': produto['descricao'],
                            'quantidade': produto['quantidade'],
                            'cgc': cgc or "N/A",
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
                        'qtde_paginas': qtde_paginas,
                        'cliente': cliente or "N/A",
                        'data_emissao': data_emissao or "N/A",
                        'uf_destino': uf_destino or "N/A",
                        'valor_total': valor_total or "N/A",
                        'codigo_sap': "N/A",
                        'descricao_material': "N/A",
                        'quantidade': "0",
                        'cgc': cgc or "N/A",
                        'endereco_instalacao': endereco or "N/A",
                        'bairro': endereco_dados['bairro'],
                        'cidade': endereco_dados['cidade'],
                        'cep': endereco_dados['cep']
                    })
            else:
                print(f"   ⚠️  NF não encontrada neste documento")
    
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
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 22
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 35
        ws.column_dimensions['I'].width = 12
        ws.column_dimensions['J'].width = 15
        ws.column_dimensions['K'].width = 15
        ws.column_dimensions['L'].width = 20
        ws.column_dimensions['M'].width = 30
        ws.column_dimensions['N'].width = 15
        ws.column_dimensions['O'].width = 15
        ws.column_dimensions['P'].width = 12
        
        # Estilo do cabeçalho: fundo cinza e texto azul marinho
        fill_cinza = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        fonte_azul = Font(color="000080", bold=True, size=11)
        alignment_centralizado = Alignment(horizontal="center", vertical="center")
        alignment_esquerda = Alignment(horizontal="left", vertical="center")
        
        # Adiciona cabeçalhos
        headers = ["Número da NF", "Qtde Páginas PDF", "Cliente", "Data de Emissão", "UF de Destino", 
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
            # Número da NF
            try:
                ws[f'A{idx}'] = int(dado['numero_nf'])
            except (ValueError, TypeError):
                ws[f'A{idx}'] = dado['numero_nf']
            ws[f'A{idx}'].font = fonte_dados
            ws[f'A{idx}'].alignment = alignment_centralizado
            
            # Qtde Páginas PDF
            ws[f'B{idx}'] = dado['qtde_paginas']
            ws[f'B{idx}'].font = fonte_dados
            ws[f'B{idx}'].alignment = alignment_centralizado
            
            # Cliente
            ws[f'C{idx}'] = dado['cliente']
            ws[f'C{idx}'].font = fonte_dados
            ws[f'C{idx}'].alignment = alignment_centralizado
            
            # Data de Emissão
            ws[f'D{idx}'] = dado['data_emissao']
            ws[f'D{idx}'].font = fonte_dados
            ws[f'D{idx}'].alignment = alignment_centralizado
            
            # UF de Destino
            ws[f'E{idx}'] = dado['uf_destino']
            ws[f'E{idx}'].font = fonte_dados
            ws[f'E{idx}'].alignment = alignment_centralizado
            
            # Valor Total (formatação Contábil do Excel)
            celula_valor = ws[f'F{idx}']
            valor_num = self._valor_para_float(dado['valor_total'])
            if valor_num is not None:
                celula_valor.value = round(valor_num, 2)
                celula_valor.number_format = self.FORMATO_CONTABIL_BR
            else:
                celula_valor.value = "N/A"
            celula_valor.font = fonte_dados
            celula_valor.alignment = alignment_direita
            
            # Código SAP
            ws[f'G{idx}'] = dado['codigo_sap']
            ws[f'G{idx}'].font = fonte_dados
            ws[f'G{idx}'].alignment = alignment_centralizado
            
            # Descrição do Material
            ws[f'H{idx}'] = dado['descricao_material']
            ws[f'H{idx}'].font = fonte_dados
            ws[f'H{idx}'].alignment = alignment_esquerda
            
            # Quantidade
            try:
                ws[f'I{idx}'] = int(dado['quantidade'])
            except (ValueError, TypeError):
                ws[f'I{idx}'] = dado['quantidade']
            ws[f'I{idx}'].font = fonte_dados
            ws[f'I{idx}'].alignment = alignment_centralizado
            
            # CGC
            if dado['cgc'] != "N/A":
                try:
                    ws[f'J{idx}'] = int(dado['cgc'])
                except (ValueError, TypeError):
                    ws[f'J{idx}'] = dado['cgc']
            else:
                ws[f'J{idx}'] = "N/A"
            ws[f'J{idx}'].font = fonte_dados
            ws[f'J{idx}'].alignment = alignment_centralizado
            
            # CIAUS
            ws[f'K{idx}'] = dado.get('ciaus', "N/A")
            ws[f'K{idx}'].font = fonte_dados
            ws[f'K{idx}'].alignment = alignment_centralizado
            
            # Field Responsável
            ws[f'L{idx}'] = dado.get('field_responsavel', "N/A")
            ws[f'L{idx}'].font = fonte_dados
            ws[f'L{idx}'].alignment = alignment_centralizado
            
            # Endereço de Instalação
            ws[f'M{idx}'] = dado['endereco_instalacao']
            ws[f'K{idx}'].font = fonte_dados
            ws[f'K{idx}'].alignment = alignment_esquerda
            
            # Bairro
            ws[f'N{idx}'] = dado['bairro']
            ws[f'N{idx}'].font = fonte_dados
            ws[f'N{idx}'].alignment = alignment_centralizado
            
            # Cidade
            ws[f'O{idx}'] = dado['cidade']
            ws[f'O{idx}'].font = fonte_dados
            ws[f'O{idx}'].alignment = alignment_centralizado
            
            # CEP
            ws[f'P{idx}'] = dado['cep']
            ws[f'P{idx}'].font = fonte_dados
            ws[f'P{idx}'].alignment = alignment_centralizado
        
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
        
        # Salva a planilha
        try:
            wb.save(self.excel_path)
            print(f"\n✅ Planilha criada com sucesso: {self.excel_path}")
        except Exception as e:
            print(f"❌ Erro ao salvar planilha: {e}")
    
    def executar(self, diretorio="."):
        """Executa o fluxo completo de leitura e geração da planilha."""
        print("=" * 50)
        print("🔍 LEITOR DE PDF - Notas Fiscais")
        print("=" * 50)
        
        self.processar_pdfs(diretorio)
        
        if self.dados_nf:
            print(f"\n📊 Total de NFs extraídas: {len(self.dados_nf)}\n")
            self.exibir_preview_planilha()
            self.criar_planilha_excel()
        else:
            print("\n⚠️  Nenhuma nota fiscal foi encontrada nos PDFs.")


if __name__ == "__main__":
    # Cria instância e executa
    leitor = LeitordePDF()
    
    # Processa PDFs no diretório atual
    leitor.executar()
