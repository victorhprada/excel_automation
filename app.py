"""
Ferramenta de Validação de Faturamento Excel
Aplicação Streamlit para upload e processamento de arquivos Excel
"""

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import get_column_letter
from copy import copy

# ========================================
# Funções Auxiliares
# ========================================

def copiar_estilo(celula_origem, celula_destino):
    """
    Copia atributos de formatação de uma célula para outra.
    
    Atributos copiados: font, border, fill, number_format, alignment
    
    Args:
        celula_origem: Célula de onde copiar o estilo
        celula_destino: Célula para onde copiar o estilo
    """
    if celula_origem.has_style:
        celula_destino.font = copy(celula_origem.font)
        celula_destino.border = copy(celula_origem.border)
        celula_destino.fill = copy(celula_origem.fill)
        celula_destino.number_format = celula_origem.number_format
        celula_destino.alignment = copy(celula_origem.alignment)


def validar_abas_necessarias(parceiro_wb, base_wb):
    """
    Valida se todas as abas necessárias existem nos workbooks.
    Retorna (sucesso: bool, mensagem: str)
    """
    abas_parceiro_necessarias = ['Parcelas Pagas', 'Produção']
    abas_base_necessarias = ['BASE', 'INADIMPLENTES', 'JAN.26']  # Incluir JAN.26 como template
    
    # Verificar PARCEIRO
    for aba in abas_parceiro_necessarias:
        if aba not in parceiro_wb.sheetnames:
            return False, f"Aba '{aba}' não encontrada no arquivo PARCEIRO"
    
    # Verificar BASE
    for aba in abas_base_necessarias:
        if aba not in base_wb.sheetnames:
            return False, f"Aba '{aba}' não encontrada no arquivo BASE"
    
    return True, "Todas as abas necessárias estão presentes (incluindo template JAN.26)"


def encontrar_ultima_linha(ws):
    """
    Encontra a última linha preenchida em uma worksheet.
    Retorna o número da linha.
    """
    for row in range(ws.max_row, 0, -1):
        # Verificar se há algum valor não-nulo na linha
        if any(ws.cell(row=row, column=col).value is not None 
               for col in range(1, ws.max_column + 1)):
            return row
    return 0  # Se worksheet vazia, retornar 0


def copiar_dados_aba(ws_origem, ws_destino, incluir_header=False):
    """
    Copia todos os dados de uma worksheet origem para destino.
    Usa encontrar_ultima_linha() para escrever na posição correta,
    evitando problemas com formatação em células vazias.
    """
    linhas_copiadas = 0
    start_row_origem = 1 if incluir_header else 2  # Pular header se incluir_header=False
    
    # Encontrar a próxima linha vazia no destino
    ultima_linha_destino = encontrar_ultima_linha(ws_destino)
    proxima_linha_destino = ultima_linha_destino + 1
    
    # Se destino estiver completamente vazio, começar da linha 1
    if ultima_linha_destino == 0:
        proxima_linha_destino = 1
    
    # Iterar sobre linhas da origem
    for row in ws_origem.iter_rows(min_row=start_row_origem, values_only=True):
        # Pular linhas completamente vazias
        if all(cell is None for cell in row):
            continue
        
        # Escrever célula por célula na linha de destino
        for col_idx, valor in enumerate(row, start=1):
            ws_destino.cell(row=proxima_linha_destino, column=col_idx, value=valor)
        
        proxima_linha_destino += 1
        linhas_copiadas += 1
    
    return linhas_copiadas


def copiar_producao_para_base(ws_origem, ws_destino):
    """
    Copia dados da aba 'Produção' para 'BASE' de forma explícita e controlada.
    
    CRÍTICO: Usa mapeamento segmentado de colunas:
    - A-G (1-7): Cópia direta origem -> destino
    - H (8): Fórmula injetada =F{row} (não vem da origem)
    - H-J origem (8-10) -> I-K destino (9-11): Deslocamento +1
    
    Não copia formatação ou fórmulas da origem (exceto fórmula injetada em H).
    Copia a formatação da última linha existente na BASE para manter consistência visual.
    
    Args:
        ws_origem: Worksheet de origem (Produção)
        ws_destino: Worksheet de destino (BASE)
    
    Returns:
        int: Número de linhas copiadas
    """
    # 1. Encontrar última linha real em BASE (onde coluna A tem valor)
    last_row_base = 0
    for row in range(1, ws_destino.max_row + 1):
        if ws_destino.cell(row=row, column=1).value is not None:
            last_row_base = row
    
    # Se BASE está vazia, começar da linha 2 (linha 1 é header)
    if last_row_base == 0:
        last_row_base = 1
    
    new_row = last_row_base + 1
    linhas_copiadas = 0
    
    # 2. Iterar sobre linhas da aba 'Produção' (começando da linha 2)
    for source_row in range(2, ws_origem.max_row + 1):
        # Verificar se linha tem dados na coluna A (se não, parar)
        if ws_origem.cell(row=source_row, column=1).value is None:
            break
        
        # 3. Copiar colunas com mapeamento segmentado
        # Etapa 3.1: Colunas A-G (1-7) - Cópia direta
        for col in range(1, 8):  # 1 a 7 (A até G)
            valor = ws_origem.cell(row=source_row, column=col).value
            cell_nova = ws_destino.cell(row=new_row, column=col, value=valor)
            
            # Copiar formatação da linha molde
            if last_row_base > 1:
                cell_molde = ws_destino.cell(row=last_row_base, column=col)
                copiar_estilo(cell_molde, cell_nova)
        
        # Etapa 3.2: Coluna H (8) - Injetar fórmula =F{row}
        cell_nova = ws_destino.cell(row=new_row, column=8, value=f"=F{new_row}")
        
        # Copiar formatação da linha molde
        if last_row_base > 1:
            cell_molde = ws_destino.cell(row=last_row_base, column=8)
            copiar_estilo(cell_molde, cell_nova)
        
        # Etapa 3.3: Colunas H-J da origem (8-10) -> I-K do destino (9-11)
        # Deslocamento: origem_col + 1 = destino_col
        for origem_col in range(8, 11):  # 8, 9, 10 (H, I, J da origem)
            destino_col = origem_col + 1  # 9, 10, 11 (I, J, K do destino)
            valor = ws_origem.cell(row=source_row, column=origem_col).value
            cell_nova = ws_destino.cell(row=new_row, column=destino_col, value=valor)
            
            # Copiar formatação da linha molde
            if last_row_base > 1:
                cell_molde = ws_destino.cell(row=last_row_base, column=destino_col)
                copiar_estilo(cell_molde, cell_nova)
        
        new_row += 1
        linhas_copiadas += 1
    
    return linhas_copiadas


def filtrar_inadimplentes(ws_origem, coluna_validacao='VALIDAÇÃO'):
    """
    Filtra linhas onde a coluna VALIDAÇÃO é igual a 'Não'.
    Retorna lista de tuplas com os dados das linhas filtradas.
    """
    inadimplentes = []
    
    # Encontrar índice da coluna VALIDAÇÃO no header (linha 1)
    header_row = list(ws_origem.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    
    try:
        col_idx = header_row.index(coluna_validacao)
    except ValueError:
        raise ValueError(f"Coluna '{coluna_validacao}' não encontrada na aba")
    
    # Filtrar linhas onde VALIDAÇÃO = 'Não'
    for row in ws_origem.iter_rows(min_row=2, values_only=True):  # Pular header
        if row[col_idx] == 'Não':
            inadimplentes.append(row)
    
    return inadimplentes


def validar_template_jan26(workbook):
    """
    Valida se a aba 'JAN.26' (template padrão) existe no workbook BASE.
    Esta aba é usada como matriz para criar todas as novas abas de mês.
    
    Args:
        workbook: Workbook do openpyxl
    
    Returns:
        tuple: (existe: bool, mensagem: str)
    """
    template_nome = 'JAN.26'
    
    if template_nome in workbook.sheetnames:
        return True, f"Template '{template_nome}' encontrado"
    else:
        return False, f"ERRO CRÍTICO: Aba '{template_nome}' não encontrada. Esta aba é necessária como template padrão."


def capturar_formulas_colunas(ws, linha=2, col_inicio=17, col_fim=24):
    """
    Captura fórmulas de colunas específicas de uma linha.
    Retorna dicionário {coluna_idx: formula_string}
    
    Args:
        ws: Worksheet do openpyxl
        linha: Linha de onde extrair fórmulas (default: 2)
        col_inicio: Primeira coluna (default: 17 = Q)
        col_fim: Última coluna (default: 24 = X)
    
    Returns:
        dict: {col_idx: formula} apenas para colunas que têm fórmulas
    """
    formulas = {}
    
    for col_idx in range(col_inicio, col_fim + 1):
        cell = ws.cell(row=linha, column=col_idx)
        
        # Verificar se a célula tem fórmula
        if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
            formulas[col_idx] = cell.value
    
    return formulas


def atualizar_formula_linha(formula, linha_nova):
    """
    Atualiza referências de linha em uma fórmula Excel.
    
    Args:
        formula: String da fórmula (ex: '=VLOOKUP(@Q:Q;BASE!A:K;11;0)')
        linha_nova: Número da nova linha
    
    Returns:
        str: Fórmula com referências de linha atualizadas
    
    Exemplos:
        atualizar_formula_linha('=IF(ISNUMBER(MATCH(V2;Q:Q;0));"Sim";"Não")', 5)
        -> '=IF(ISNUMBER(MATCH(V5;Q:Q;0));"Sim";"Não")'
    """
    import re
    
    # Padrão para referências de célula com linha específica (ex: A2, V2, Q2)
    # Captura letra(s) seguida(s) de número
    padrao = r'([A-Z]+)(\d+)'
    
    def substituir_linha(match):
        coluna = match.group(1)
        # Substituir qualquer número de linha pelo novo
        return f"{coluna}{linha_nova}"
    
    # Substituir todas as referências de linha na fórmula
    formula_atualizada = re.sub(padrao, substituir_linha, formula)
    
    return formula_atualizada


def limpar_dados_worksheet(ws, manter_linha_1=True):
    """
    Limpa todos os dados de uma worksheet, mantendo a linha 1 (header).
    
    Args:
        ws: Worksheet do openpyxl
        manter_linha_1: Se True, mantém linha 1 intacta
    """
    linha_inicial = 2 if manter_linha_1 else 1
    
    # Iterar de trás para frente para evitar problemas com índices
    for row_idx in range(ws.max_row, linha_inicial - 1, -1):
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_idx, column=col_idx).value = None


def aplicar_regras_colunas_n_x(ws, target_month, linha_inicio=2):
    """
    Aplica regras de negócio explícitas para as colunas N até X.
    
    Args:
        ws: Worksheet do openpyxl onde aplicar as regras
        target_month: String do mês alvo (ex: 'JAN.26')
        linha_inicio: Linha inicial (default: 2, primeira linha de dados)
    
    Returns:
        dict: {'linhas_n_o': int, 'linhas_q_w': int, 'ccbs_unicos': int}
    
    Regras:
        MOMENTO A (Colunas N-O para todas as linhas):
            Col N (14) - Mês Faturado: target_month formatado (minúsculo, hífen)
            Col O (15) - Data Desembolso: =VLOOKUP(A{row},'BASE'!A:H,8,0)
            Col P (16) - Separador: None (vazio)
        
        MOMENTO B (Colunas Q-W apenas para CCBs únicos):
            Col Q (17) - CCB: Valor único da coluna A
            Col R (18) - Mês Originação: =VLOOKUP(Q{row},'BASE'!A:K,11,0)
            Col S (19) - Repasse: =SUMIF(A:A,Q{row},L:L)
            Col T (20) - Data Desemb 1: =VLOOKUP(Q{row},'BASE'!A:H,8,0)
            Col U (21) - Separador: None (vazio)
            Col V (22) e W (23): None (vazio)
            Col X (24) - Vazio (removida fórmula)
    """
    # ========================================
    # PREPARAÇÃO: Formatar target_month
    # ========================================
    # Converter 'JAN.26' -> 'jan-26' (minúsculo com hífen)
    mes_faturado = target_month.replace('.', '-').lower()
    
    # ========================================
    # ENCONTRAR ÚLTIMA LINHA COM DADOS
    # ========================================
    ultima_linha = linha_inicio - 1
    for row in range(linha_inicio, ws.max_row + 1):
        if ws.cell(row=row, column=1).value is not None:
            ultima_linha = row
        else:
            break
    
    if ultima_linha < linha_inicio:
        return {'linhas_n_o': 0, 'linhas_q_w': 0, 'ccbs_unicos': 0}
    
    # ========================================
    # MOMENTO A: Preencher Colunas N-O (todas as linhas)
    # ========================================
    linhas_n_o = 0
    
    for row in range(linha_inicio, ultima_linha + 1):
        # Col N (14) - Mês Faturado: String formatada
        ws.cell(row=row, column=14, value=mes_faturado)
        
        # Col O (15) - Data Desembolso: Fórmula VLOOKUP
        ws.cell(row=row, column=15, value=f"=VLOOKUP(A{row},'BASE'!A:H,8,0)")
        
        # Col P (16) - Separador: Vazio
        ws.cell(row=row, column=16, value=None)
        
        linhas_n_o += 1
    
    # ========================================
    # MOMENTO B: Preencher Colunas Q-W (apenas CCBs únicos)
    # ========================================
    
    # Extrair todos os valores da coluna A (CCBs)
    ccbs_todos = []
    for row in range(linha_inicio, ultima_linha + 1):
        valor_a = ws.cell(row=row, column=1).value
        if valor_a is not None:
            ccbs_todos.append(valor_a)
    
    # Gerar lista de CCBs únicos (preservando ordem de primeira aparição)
    ccbs_unicos = []
    vistos = set()
    for ccb in ccbs_todos:
        if ccb not in vistos:
            ccbs_unicos.append(ccb)
            vistos.add(ccb)
    
    # Preencher colunas Q-W para cada CCB único
    linhas_q_w = 0
    row_destino = linha_inicio
    
    for ccb_unico in ccbs_unicos:
        # Col Q (17) - CCB: Valor único
        ws.cell(row=row_destino, column=17, value=ccb_unico)
        
        # Col R (18) - Mês Originação: Fórmula VLOOKUP
        ws.cell(row=row_destino, column=18, value=f"=VLOOKUP(Q{row_destino},'BASE'!A:K,11,0)")
        
        # Col S (19) - Repasse: Fórmula SUMIF
        ws.cell(row=row_destino, column=19, value=f"=SUMIF(A:A,Q{row_destino},L:L)")
        
        # Col T (20) - Data Desemb 1: Fórmula VLOOKUP
        ws.cell(row=row_destino, column=20, value=f"=VLOOKUP(Q{row_destino},'BASE'!A:H,8,0)")
        
        # Col U (21) - Separador: Vazio
        ws.cell(row=row_destino, column=21, value=None)
        
        # Col V (22) e W (23): Vazios
        ws.cell(row=row_destino, column=22, value=None)
        ws.cell(row=row_destino, column=23, value=None)
        
        # Col X (24) - Vazio (sem fórmula)
        ws.cell(row=row_destino, column=24, value=None)
        
        row_destino += 1
        linhas_q_w += 1
    
    return {
        'linhas_n_o': linhas_n_o,
        'linhas_q_w': linhas_q_w,
        'ccbs_unicos': len(ccbs_unicos)
    }


def encontrar_colunas_meses(ws_base):
    """
    Identifica colunas de meses na aba BASE.
    
    Returns:
        list: [
            {'nome': 'Setembro', 'indice': 17, 'letra': 'Q'},
            {'nome': 'Outubro', 'indice': 18, 'letra': 'R'},
            ...
        ]
    """
    colunas_meses = []
    
    # Encontrar índice da coluna P (última coluna antes dos meses)
    col_p_index = 16  # P = 16
    
    # Encontrar índice da coluna V (DATA) - buscar pelo header
    col_v_index = None
    for col in range(1, ws_base.max_column + 1):
        header = ws_base.cell(row=1, column=col).value
        if header == 'DATA':
            col_v_index = col
            break
    
    if not col_v_index:
        # Se não encontrar V, usar max_column
        col_v_index = ws_base.max_column + 1
    
    # Iterar entre P+1 e V-1
    for col_idx in range(col_p_index + 1, col_v_index):
        header = ws_base.cell(row=1, column=col_idx).value
        if header:  # Se tem cabeçalho, é coluna de mês
            colunas_meses.append({
                'nome': header,
                'indice': col_idx,
                'letra': get_column_letter(col_idx)
            })
    
    return colunas_meses


def inserir_coluna_mes(ws_base, target_month, colunas_meses):
    """
    Insere nova coluna de mês na aba BASE.
    
    Args:
        ws_base: Worksheet da BASE
        target_month: String do mês (ex: 'JAN.26')
        colunas_meses: Lista de colunas de meses existentes
    
    Returns:
        dict: {'nome': 'JAN.26', 'indice': 22, 'letra': 'V'}
    """
    # Determinar posição de inserção
    if colunas_meses:
        # Inserir após a última coluna de mês
        ultimo_mes_idx = colunas_meses[-1]['indice']
        pos_insercao = ultimo_mes_idx + 1
    else:
        # Se não há colunas de meses, inserir após P
        pos_insercao = 17  # Q
    
    # Inserir coluna
    ws_base.insert_cols(pos_insercao)
    
    # Definir cabeçalho
    ws_base.cell(row=1, column=pos_insercao, value=target_month)
    
    # Aplicar fórmula COUNTIF em todas as linhas (da linha 2 até última)
    ultima_linha = encontrar_ultima_linha(ws_base)
    
    for row in range(2, ultima_linha + 1):
        # Fórmula: =COUNTIF('JAN.26'!A:A, BASE!A2)
        formula = f"=COUNTIF('{target_month}'!A:A,BASE!A{row})"
        ws_base.cell(row=row, column=pos_insercao, value=formula)
    
    return {
        'nome': target_month,
        'indice': pos_insercao,
        'letra': get_column_letter(pos_insercao)
    }


def aplicar_formulas_dinamicas(ws_base, colunas_meses, base_wb):
    """
    Aplica fórmulas dinâmicas L, M, N em TODAS as linhas da BASE.
    
    CRÍTICO: Deve processar TODAS as linhas (da 2 até última), não apenas novas,
    pois registros antigos podem ter pago no novo mês e precisam ser atualizados.
    
    CRÍTICO: Usa APENAS abas locais do workbook, sem referências externas.
    
    Args:
        ws_base: Worksheet da BASE
        colunas_meses: Lista de colunas de meses (incluindo a nova)
        base_wb: Workbook da BASE (para validar sheetnames)
    
    Returns:
        int: Número de linhas processadas
    """
    ultima_linha = encontrar_ultima_linha(ws_base)
    linhas_processadas = 0
    
    # CORREÇÃO: Construir lista de abas validando contra workbook.sheetnames
    # Isso garante que APENAS abas locais sejam usadas nas fórmulas
    abas_meses_validas = []
    sheetnames_disponiveis = base_wb.sheetnames
    
    for col_mes in colunas_meses:
        nome_header = col_mes['nome']
        
        # Validar: essa aba existe localmente no workbook?
        if nome_header in sheetnames_disponiveis:
            abas_meses_validas.append(nome_header)
        # Se não existe, ignorar (não adicionar warning para não poluir UI)
    
    # Se não houver abas válidas, retornar erro
    if not abas_meses_validas:
        raise ValueError("Nenhuma aba de mês válida encontrada no workbook BASE")
    
    # IMPORTANTE: Processar TODAS as linhas (2 até última), não apenas novas
    # Usar abas_meses_validas (sem referências externas)
    for row in range(2, ultima_linha + 1):
        # ===== COLUNA L (12) - Parcela Paga? =====
        # =IF(OR(NOT(ISERROR(VLOOKUP(A2,'Setembro'!A:A,1,0))), ...), "Sim", "Não")
        vlookup_parts = []
        for aba in abas_meses_validas:
            # Garantir que a referência seja APENAS 'NomeAba'!A:A
            vlookup_parts.append(f"NOT(ISERROR(VLOOKUP(A{row},'{aba}'!A:A,1,0)))")
        
        formula_l = f'=IF(OR({",".join(vlookup_parts)}),"Sim","Não")'
        cell_l = ws_base.cell(row=row, column=12, value=formula_l)
        
        # Copiar formatação da linha anterior
        if row > 2:
            linha_molde = row - 1
            copiar_estilo(ws_base.cell(row=linha_molde, column=12), cell_l)
        
        # ===== COLUNA M (13) - Data Pagamento =====
        # =IFERROR(VLOOKUP(...,'Setembro'!A:N,14,0), IFERROR(..., "Pendente"))
        formula_m = ""
        for aba in abas_meses_validas:
            if formula_m == "":
                formula_m = f"IFERROR(VLOOKUP(A{row},'{aba}'!A:N,14,0)"
            else:
                formula_m += f",IFERROR(VLOOKUP(A{row},'{aba}'!A:N,14,0)"
        
        # Fechar todos os IFERRORs e adicionar fallback
        formula_m += "," + '"Pendente de pagamento"' + ")" * len(abas_meses_validas)
        formula_m = "=" + formula_m
        
        cell_m = ws_base.cell(row=row, column=13, value=formula_m)
        
        # Copiar formatação da linha anterior
        if row > 2:
            linha_molde = row - 1
            copiar_estilo(ws_base.cell(row=linha_molde, column=13), cell_m)
        
        # ===== COLUNA N (14) - Parcelas Recebidas =====
        # =COUNTIF('Setembro'!A:A,BASE!A2) + COUNTIF('Outubro'!A:A,BASE!A2) + ...
        countif_parts = []
        for aba in abas_meses_validas:
            countif_parts.append(f"COUNTIF('{aba}'!A:A,BASE!A{row})")
        
        formula_n = f'={"+".join(countif_parts)}'
        cell_n = ws_base.cell(row=row, column=14, value=formula_n)
        
        # Copiar formatação da linha anterior
        if row > 2:
            linha_molde = row - 1
            copiar_estilo(ws_base.cell(row=linha_molde, column=14), cell_n)
        
        linhas_processadas += 1
    
    return linhas_processadas


def aplicar_formulas_estaticas(ws_base, linha_inicio):
    """
    Aplica fórmulas estáticas O, P, V nas novas linhas.
    
    Args:
        ws_base: Worksheet da BASE
        linha_inicio: Primeira linha onde começaram os novos dados
    
    Returns:
        int: Número de linhas processadas
    """
    ultima_linha = encontrar_ultima_linha(ws_base)
    
    # Encontrar índice da coluna V (DATA) dinamicamente
    col_v_index = None
    for col in range(1, ws_base.max_column + 1):
        header = ws_base.cell(row=1, column=col).value
        if header == 'DATA':
            col_v_index = col
            break
    
    if not col_v_index:
        # Se não encontrar, assumir que está após as colunas de meses
        # (pode precisar ajustar dependendo da inserção)
        col_v_index = ws_base.max_column
    
    linhas_processadas = 0
    
    for row in range(linha_inicio, ultima_linha + 1):
        # Linha molde: linha anterior (row - 1)
        linha_molde = row - 1
        
        # Col O (15) - % Recebimento: =N2/E2
        cell_o = ws_base.cell(row=row, column=15, value=f"=N{row}/E{row}")
        if linha_molde >= 2:
            copiar_estilo(ws_base.cell(row=linha_molde, column=15), cell_o)
        
        # Col P (16) - Pendentes: =E2-N2
        cell_p = ws_base.cell(row=row, column=16, value=f"=E{row}-N{row}")
        if linha_molde >= 2:
            copiar_estilo(ws_base.cell(row=linha_molde, column=16), cell_p)
        
        # Col V (índice dinâmico) - Helper: =LEFT(F2,10)
        cell_v = ws_base.cell(row=row, column=col_v_index, value=f"=LEFT(F{row},10)")
        if linha_molde >= 2:
            copiar_estilo(ws_base.cell(row=linha_molde, column=col_v_index), cell_v)
        
        linhas_processadas += 1
    
    return linhas_processadas


def atualizar_aba_base(base_wb, parceiro_wb, target_month, linha_inicio_append):
    """
    Atualiza a aba BASE com novos dados e fórmulas dinâmicas.
    
    IMPORTANTE: As fórmulas dinâmicas (L, M, N) são aplicadas em TODAS as linhas,
    não apenas nas novas, pois registros antigos podem ter pago no novo mês.
    
    Args:
        base_wb: Workbook do arquivo BASE
        parceiro_wb: Workbook do arquivo PARCEIRO
        target_month: String do mês (ex: 'JAN.26')
        linha_inicio_append: Primeira linha onde foram adicionados dados de Produção
                           (usado apenas para fórmulas estáticas O, P, V)
    
    Returns:
        dict: {
            'linhas_producao': int,
            'coluna_mes_inserida': str,
            'abas_meses_encontradas': list,
            'linhas_formulas_aplicadas': int,    # Total de linhas (L, M, N)
            'linhas_novas_estaticas': int        # Apenas novas (O, P, V)
        }
    """
    # 1. Obter referências
    ws_base = base_wb['BASE']
    ws_producao = parceiro_wb['Produção']
    
    # 2. Identificar colunas de meses existentes (entre P e V)
    colunas_meses = encontrar_colunas_meses(ws_base)
    
    # 3. Inserir nova coluna de mês
    col_inserida = inserir_coluna_mes(ws_base, target_month, colunas_meses)
    
    # 4. Atualizar colunas_meses com a nova coluna
    colunas_meses.append(col_inserida)
    
    # 5. Aplicar fórmulas dinâmicas (L, M, N) em TODAS as linhas
    # CRÍTICO: Atualiza todas as linhas, não apenas novas, pois registros
    # antigos podem ter pago no novo mês e precisam ser atualizados
    # CORREÇÃO: Passar base_wb para validação de abas locais
    linhas_processadas = aplicar_formulas_dinamicas(
        ws_base, 
        colunas_meses,
        base_wb  # NOVO: passar workbook para validação
    )
    
    # 6. Aplicar fórmulas estáticas (O, P, V) nas novas linhas
    linhas_novas = aplicar_formulas_estaticas(ws_base, linha_inicio_append)
    
    # 7. Retornar métricas
    return {
        'coluna_mes_inserida': target_month,
        'abas_meses_encontradas': [col['nome'] for col in colunas_meses],
        'linhas_formulas_aplicadas': linhas_processadas,  # L, M, N (todas)
        'linhas_novas_estaticas': linhas_novas           # O, P, V (apenas novas)
    }


def inserir_dados_colunas_especificas(ws_origem, ws_destino, col_inicio=1, col_fim=13, linha_destino_inicio=2):
    """
    Copia dados de worksheet origem para destino, mas apenas em colunas específicas.
    
    Args:
        ws_origem: Worksheet de origem
        ws_destino: Worksheet de destino
        col_inicio: Primeira coluna a copiar (default: 1 = A)
        col_fim: Última coluna a copiar (default: 13 = M)
        linha_destino_inicio: Linha inicial no destino (default: 2)
    
    Returns:
        int: Número de linhas copiadas
    
    Nota:
        Colunas N-X são preenchidas pela função aplicar_regras_colunas_n_x()
    """
    linhas_copiadas = 0
    linha_destino = linha_destino_inicio
    
    # Iterar sobre linhas da origem (pulando header - linha 1)
    for row in ws_origem.iter_rows(min_row=2, values_only=True):
        # Pular linhas vazias
        if all(cell is None for cell in row):
            continue
        
        # Copiar apenas colunas especificadas
        for col_idx in range(col_inicio, min(col_fim + 1, len(row) + 1)):
            valor = row[col_idx - 1] if col_idx <= len(row) else None
            ws_destino.cell(row=linha_destino, column=col_idx, value=valor)
        
        linha_destino += 1
        linhas_copiadas += 1
    
    return linhas_copiadas


def reaplicar_formulas(ws, formulas_dict, linha_inicio=2, linha_fim=None):
    """
    Aplica fórmulas capturadas em um range de linhas, atualizando referências.
    
    Args:
        ws: Worksheet do openpyxl
        formulas_dict: Dict {col_idx: formula_template}
        linha_inicio: Primeira linha onde aplicar (default: 2)
        linha_fim: Última linha (default: None = até última linha com dados)
    
    Returns:
        int: Número de fórmulas aplicadas
    """
    if linha_fim is None:
        linha_fim = encontrar_ultima_linha(ws)
    
    formulas_aplicadas = 0
    
    for linha in range(linha_inicio, linha_fim + 1):
        for col_idx, formula_template in formulas_dict.items():
            # Atualizar referências de linha na fórmula
            formula_atualizada = atualizar_formula_linha(formula_template, linha)
            
            # Aplicar fórmula na célula
            ws.cell(row=linha, column=col_idx, value=formula_atualizada)
            formulas_aplicadas += 1
    
    return formulas_aplicadas


# ========================================
# Configuração da Página
# ========================================

st.set_page_config(
    page_title="Validação de Faturamento Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========================================
# Título Principal
# ========================================

st.title("📊 Ferramenta de Validação de Faturamento")
st.markdown("---")

# ========================================
# Sidebar - Configurações
# ========================================

st.sidebar.header("⚙️ Configurações")
st.sidebar.markdown("### 📅 Período de Análise")

# Selectbox para Mês
meses = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 
         'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
mes_selecionado = st.sidebar.selectbox(
    "Mês",
    options=meses,
    index=0
)

# Selectbox para Ano
anos = ['24', '25', '26']
ano_selecionado = st.sidebar.selectbox(
    "Ano",
    options=anos,
    index=1  # Default para '25'
)

# Concatenar para formar target_month no formato MMM.YY
target_month = f"{mes_selecionado}.{ano_selecionado}"

# Exibir o período selecionado
st.sidebar.success(f"**Período Selecionado:** {target_month}")
st.sidebar.markdown("---")

# ========================================
# Área Principal - Upload de Arquivos
# ========================================

st.header("📁 Upload de Arquivos")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Arquivo PARCEIRO")
    arquivo_parceiro = st.file_uploader(
        "Selecione o arquivo PARCEIRO (.xlsx)",
        type=['xlsx'],
        key='parceiro',
        help="Arquivo Excel com dados do parceiro"
    )
    
    if arquivo_parceiro:
        st.success(f"✅ {arquivo_parceiro.name}")
        st.info(f"Tamanho: {arquivo_parceiro.size / 1024:.2f} KB")

with col2:
    st.subheader("Arquivo BASE")
    arquivo_base = st.file_uploader(
        "Selecione o arquivo BASE (.xlsx ou .xlsm)",
        type=['xlsx', 'xlsm'],
        key='base',
        help="Arquivo Excel base (fórmulas serão preservadas)"
    )
    
    if arquivo_base:
        st.success(f"✅ {arquivo_base.name}")
        st.info(f"Tamanho: {arquivo_base.size / 1024:.2f} KB")

st.markdown("---")

# ========================================
# Botão de Processamento
# ========================================

st.header("🚀 Processamento")

# Verificar se ambos os arquivos foram carregados
arquivos_prontos = arquivo_parceiro is not None and arquivo_base is not None

if not arquivos_prontos:
    st.warning("⚠️ Por favor, faça upload dos dois arquivos para continuar.")

# Botão de processamento
processar = st.button(
    "🔄 Iniciar Processamento",
    type="primary",
    disabled=not arquivos_prontos,
    use_container_width=True
)

# ========================================
# Lógica de Processamento e Session State
# ========================================

if processar and arquivos_prontos:
    try:
        with st.spinner("Processando arquivos..."):
            
            # Armazenar target_month
            st.session_state['target_month'] = target_month
            
            # ==================================================
            # ETAPA 1: Carregar Arquivos com Openpyxl
            # ==================================================
            st.info("📄 Carregando arquivos...")
            
            # Carregar PARCEIRO
            arquivo_parceiro.seek(0)
            parceiro_wb = openpyxl.load_workbook(
                BytesIO(arquivo_parceiro.read()),
                data_only=True
            )
            
            # Carregar BASE
            arquivo_base.seek(0)
            base_wb = openpyxl.load_workbook(
                BytesIO(arquivo_base.read()),
                data_only=False  # Preservar fórmulas
            )
            
            # ==================================================
            # ETAPA 2: Validar Abas Necessárias
            # ==================================================
            st.info("🔍 Validando estrutura dos arquivos...")
            
            valido, mensagem = validar_abas_necessarias(parceiro_wb, base_wb)
            if not valido:
                st.error(f"❌ {mensagem}")
                st.stop()
            
            st.success(f"✅ {mensagem}")
            
            # ==================================================
            # ETAPA 3: Clonar Template 'JAN.26' para target_month
            # ==================================================
            st.info(f"📝 Preparando aba '{target_month}' a partir do template 'JAN.26'...")
            
            # Validar que template JAN.26 existe
            template_existe, mensagem_template = validar_template_jan26(base_wb)
            
            if not template_existe:
                st.error(f"❌ {mensagem_template}")
                st.error("A aba 'JAN.26' deve existir no arquivo BASE como template padrão.")
                st.stop()
            
            st.success(f"✅ {mensagem_template}")
            
            # Remover aba target_month se já existir
            if target_month in base_wb.sheetnames:
                st.warning(f"⚠️ Aba '{target_month}' já existe. Será substituída.")
                del base_wb[target_month]
            
            # Clonar aba JAN.26 para criar nova aba
            st.info("📋 Clonando estrutura de 'JAN.26'...")
            ws_template = base_wb['JAN.26']
            ws_mes = base_wb.copy_worksheet(ws_template)
            ws_mes.title = target_month
            
            st.success(f"✅ Aba '{target_month}' criada com estrutura idêntica a 'JAN.26'")
            st.info("ℹ️ Estrutura clonada: Headers, larguras de coluna, formatação")
            
            # ==================================================
            # ETAPA 4: Limpar, Inserir Dados e Aplicar Regras
            # ==================================================
            st.info("📋 Processando dados na nova aba...")
            
            # Sub-etapa 4.1: Limpar dados antigos (manter header)
            st.info("🧹 Limpando dados da linha 2 para baixo...")
            limpar_dados_worksheet(ws_mes, manter_linha_1=True)
            st.success("✅ Dados antigos removidos (Linha 1 - Headers preservados)")
            
            # Sub-etapa 4.2: Inserir dados do parceiro nas colunas A-M
            st.info("📥 Inserindo dados de 'Parcelas Pagas' (colunas A-M)...")
            ws_parcela_paga = parceiro_wb['Parcelas Pagas']
            
            linhas_copiadas = inserir_dados_colunas_especificas(
                ws_parcela_paga,
                ws_mes,
                col_inicio=1,   # Coluna A
                col_fim=13,     # Coluna M
                linha_destino_inicio=2
            )
            
            st.success(f"✅ {linhas_copiadas} linhas inseridas nas colunas A-M")
            
            # Sub-etapa 4.3: Aplicar regras de negócio nas colunas N-X
            st.info("🔧 Aplicando regras de negócio nas colunas N-X...")
            
            resultado = aplicar_regras_colunas_n_x(
                ws_mes,
                target_month,
                linha_inicio=2
            )
            
            st.success(f"✅ Regras aplicadas com sucesso!")
            
            # Mostrar métricas
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Linhas N-O", resultado['linhas_n_o'])
            with col2:
                st.metric("CCBs Únicos", resultado['ccbs_unicos'])
            with col3:
                st.metric("Linhas Q-W", resultado['linhas_q_w'])
            
            # Detalhar o que foi aplicado
            with st.expander("📋 Detalhes das Regras Aplicadas"):
                st.write("**MOMENTO A - Colunas N-O (todas as linhas):**")
                st.write(f"- Col N: Mês Faturado formatado ('{target_month.replace('.', '-').lower()}')")
                st.write("- Col O: Data Desembolso (VLOOKUP)")
                st.write("- Col P: Separador (vazio)")
                st.write("")
                st.write("**MOMENTO B - Colunas Q-W (apenas CCBs únicos):**")
                st.write("- Col Q: CCB único (deduplicated)")
                st.write("- Col R: Mês Originação (VLOOKUP)")
                st.write("- Col S: Repasse (SUMIF)")
                st.write("- Col T: Data Desemb 1 (VLOOKUP)")
                st.write("- Col U: Separador (vazio)")
                st.write("- Col V, W: Vazios")
                st.write("- Col X: Vazio (sem fórmula)")
                st.write("")
                st.info(f"ℹ️ Tabela esquerda (A-P): {resultado['linhas_n_o']} linhas")
                st.info(f"ℹ️ Tabela direita (Q-W): {resultado['linhas_q_w']} linhas (apenas CCBs únicos)")
            
            st.success(f"✅ Aba '{target_month}' configurada com sucesso!")
            st.write(f"📊 Estrutura: A-M (dados), N-O (todas linhas), Q-W (CCBs únicos)")
            
            # ==================================================
            # ETAPA 5: Atualizar Aba BASE
            # ==================================================
            st.info("📊 Atualizando aba BASE (Produção + Fórmulas)...")
            
            # Sub-etapa 5.1: Identificar linha inicial para append
            ultima_linha_base_antes = encontrar_ultima_linha(base_wb['BASE'])
            linha_inicio_append = ultima_linha_base_antes + 1
            
            st.write(f"Última linha em BASE antes do append: {ultima_linha_base_antes}")
            
            # Sub-etapa 5.2: Append dados de Produção (colunas A-J APENAS)
            ws_producao = parceiro_wb['Produção']
            ws_base = base_wb['BASE']
            
            # CORREÇÃO: Usar nova função que copia explicitamente apenas A-J
            linhas_append = copiar_producao_para_base(
                ws_producao,
                ws_base
            )
            
            st.success(f"✅ {linhas_append} linhas de Produção adicionadas (colunas A-J)")
            st.info("ℹ️ Copiados apenas valores das colunas A-J, sem formatação")
            
            # Sub-etapa 5.3: Atualizar BASE completa
            st.info("🔧 Atualizando colunas dinâmicas e fórmulas...")
            st.warning("⚠️ Atualizando fórmulas em TODAS as linhas (registros antigos + novos)")
            
            resultado_base = atualizar_aba_base(
                base_wb,
                parceiro_wb,
                target_month,
                linha_inicio_append
            )
            
            st.success(f"✅ Aba BASE atualizada com sucesso!")
            
            # Métricas
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Coluna Inserida", resultado_base['coluna_mes_inserida'])
            with col2:
                st.metric("Abas de Meses", len(resultado_base['abas_meses_encontradas']))
            with col3:
                st.metric("Fórmulas L-M-N", resultado_base['linhas_formulas_aplicadas'])
            with col4:
                st.metric("Fórmulas O-P-V", resultado_base['linhas_novas_estaticas'])
            
            # Detalhes
            with st.expander("📋 Detalhes da Atualização"):
                st.write(f"**Abas de meses referenciadas:** {', '.join(resultado_base['abas_meses_encontradas'])}")
                st.write(f"**Fórmulas dinâmicas (L, M, N):** Atualizadas em TODAS as {resultado_base['linhas_formulas_aplicadas']} linhas")
                st.write(f"**Fórmulas estáticas (O, P, V):** Aplicadas nas {resultado_base['linhas_novas_estaticas']} novas linhas")
                st.write(f"**Nova coluna '{target_month}' inserida com fórmula:** =COUNTIF('{target_month}'!A:A,BASE!A#)")
                st.info("ℹ️ Registros antigos que pagaram no novo mês agora mostram 'Sim' em 'Parcela Paga?'")
            
            # ==================================================
            # ETAPA 6: Filtrar Inadimplentes
            # ==================================================
            st.info("🔍 Filtrando inadimplentes (VALIDAÇÃO = 'Não')...")
            
            try:
                inadimplentes = filtrar_inadimplentes(ws_mes)
                
                if inadimplentes:
                    ws_inadimplentes = base_wb['INADIMPLENTES']
                    
                    # Encontrar próxima linha vazia em INADIMPLENTES
                    ultima_linha_inad = encontrar_ultima_linha(ws_inadimplentes)
                    proxima_linha_inad = ultima_linha_inad + 1
                    
                    # Adicionar inadimplentes célula por célula (não usar .append())
                    for row_data in inadimplentes:
                        for col_idx, valor in enumerate(row_data, start=1):
                            ws_inadimplentes.cell(row=proxima_linha_inad, column=col_idx, value=valor)
                        proxima_linha_inad += 1
                    
                    st.success(f"✅ {len(inadimplentes)} inadimplentes adicionados")
                else:
                    st.info("ℹ️ Nenhum inadimplente encontrado")
                    
            except ValueError as e:
                st.warning(f"⚠️ {str(e)}")
            
            # ==================================================
            # ETAPA 7: Armazenar em Session State
            # ==================================================
            st.session_state['base_workbook_modificado'] = base_wb
            st.session_state['base_filename'] = arquivo_base.name
            st.session_state['processado'] = True
            
        st.success("✅ Processamento concluído com sucesso!")
        st.balloons()
        
    except Exception as e:
        st.error(f"❌ Erro ao processar arquivos: {str(e)}")
        st.exception(e)

# ========================================
# Resumo das Operações (se processado)
# ========================================

if st.session_state.get('processado', False):
    st.markdown("---")
    st.header("📊 Resumo das Operações")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Aba Criada", st.session_state['target_month'])
    
    with col2:
        # Contar linhas na aba do mês
        ws_mes = st.session_state['base_workbook_modificado'][st.session_state['target_month']]
        st.metric("Linhas em " + st.session_state['target_month'], ws_mes.max_row)
    
    with col3:
        ws_base = st.session_state['base_workbook_modificado']['BASE']
        st.metric("Total em BASE", ws_base.max_row)
    
    with col4:
        ws_inad = st.session_state['base_workbook_modificado']['INADIMPLENTES']
        st.metric("Total Inadimplentes", ws_inad.max_row)

# ========================================
# Botão de Download do Arquivo BASE Modificado
# ========================================

if st.session_state.get('processado', False):
    st.markdown("---")
    st.header("💾 Download do Arquivo Processado")
    
    # Preparar arquivo para download
    base_wb_modificado = st.session_state.get('base_workbook_modificado')
    
    if base_wb_modificado:
        # Salvar workbook em BytesIO
        output = BytesIO()
        base_wb_modificado.save(output)
        output.seek(0)
        
        # Nome do arquivo de saída
        nome_original = st.session_state.get('base_filename', 'BASE.xlsx')
        nome_saida = nome_original.replace('.xlsx', f'_{target_month}_processado.xlsx')
        nome_saida = nome_saida.replace('.xlsm', f'_{target_month}_processado.xlsx')
        
        # Botão de download
        st.download_button(
            label="⬇️ Download Arquivo BASE Processado",
            data=output.getvalue(),
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        st.success(f"✅ Arquivo pronto: {nome_saida}")

# ========================================
# Footer
# ========================================

st.sidebar.markdown("---")
st.sidebar.markdown("### 📌 Instruções")
st.sidebar.markdown("""
1. Selecione o **mês** e **ano**
2. Faça upload do arquivo **PARCEIRO**
3. Faça upload do arquivo **BASE**
4. Clique em **Iniciar Processamento**
5. Aguarde o processamento das operações:
   - Criação da aba do mês
   - Cópia de dados 'Parcelas Pagas'
   - Append de dados 'Produção'
   - Filtro de inadimplentes
6. Faça o **download** do arquivo processado
""")

st.sidebar.markdown("---")
st.sidebar.caption("Ferramenta de Validação de Faturamento v2.0")
