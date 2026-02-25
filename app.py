import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border
from openpyxl.cell.cell import MergedCell
from copy import copy
from datetime import date
from dateutil.relativedelta import relativedelta
import re
from datetime import datetime
from openpyxl.utils import column_index_from_string
import gc

# =========================================================
# 1. FUNÇÕES AUXILIARES E REGRAS DE NEGÓCIO (Intactas)
# =========================================================

def copiar_estilo(celula_origem, celula_destino):
    if celula_origem.has_style:
        celula_destino.font = copy(celula_origem.font)
        b_origem = celula_origem.border
        if b_origem:
            celula_destino.border = Border(
                left=copy(b_origem.left), right=copy(b_origem.right),
                top=copy(b_origem.top), bottom=copy(b_origem.bottom),
                diagonal=copy(b_origem.diagonal), diagonal_direction=b_origem.diagonal_direction,
                outline=b_origem.outline, vertical=b_origem.vertical, horizontal=b_origem.horizontal
            )
        celula_destino.fill = copy(celula_origem.fill)
        celula_destino.number_format = celula_origem.number_format
        celula_destino.alignment = copy(celula_origem.alignment)

def encontrar_coluna_por_header(ws, nome_header):
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == nome_header:
            return col
    return None

def validar_abas_necessarias(parceiro_wb, base_wb):
    abas_parceiro_necessarias = ['Parcelas Pagas', 'Produção']
    abas_base_necessarias = ['BASE', 'INADIMPLENTES', 'JAN.26'] 
    for aba in abas_parceiro_necessarias:
        if aba not in parceiro_wb.sheetnames: return False, f"Aba '{aba}' não encontrada no arquivo PARCEIRO"
    for aba in abas_base_necessarias:
        if aba not in base_wb.sheetnames: return False, f"Aba '{aba}' não encontrada no arquivo BASE"
    return True, "Todas as abas necessárias estão presentes"

def encontrar_ultima_linha(ws):
    for row in range(ws.max_row, 0, -1):
        if any(ws.cell(row=row, column=col).value is not None for col in range(1, ws.max_column + 1)):
            return row
    return 0

def calcular_mes_anterior(mes_str):
    meses_eng = {'JAN': 1, 'FEV': 2, 'MAR': 3, 'ABR': 4, 'MAI': 5, 'JUN': 6, 'JUL': 7, 'AGO': 8, 'SET': 9, 'OUT': 10, 'NOV': 11, 'DEZ': 12}
    meses_pt = {1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun', 7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'}
    partes = mes_str.upper().strip().split('.')
    mes_num = meses_eng[partes[0]]
    ano = 2000 + int(partes[1])
    d = date(ano, mes_num, 1)
    d_ant = date(ano - 1, 12, 1) if mes_num == 1 else date(ano, mes_num - 1, 1)
    return f"{meses_pt[d_ant.month]}/{str(d_ant.year)[-2:]}"

def encontrar_ultima_coluna_resumo(ws):
    ultima = 1
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=2, column=col).value is not None: ultima = col
    return ultima

def atualizar_resumo_mes_faturamento(base_wb, target_month):
    ws_resumo = base_wb['RESUMO']
    ultima_col = encontrar_ultima_coluna_resumo(ws_resumo)
    nova_coluna = ultima_col + 1
    
    linha2_vazia = ws_resumo.cell(row=2, column=nova_coluna).value is None
    linha9_vazia = ws_resumo.cell(row=9, column=nova_coluna).value is None
    linha9_valor = ws_resumo.cell(row=9, column=nova_coluna).value
    eh_header_regras = linha9_valor and 'REGRA' in str(linha9_valor).upper()
    
    if not (linha2_vazia and linha9_vazia and not eh_header_regras): ws_resumo.insert_cols(nova_coluna)
    
    letra = get_column_letter(nova_coluna)
    mes_faturado = target_month.replace('.', '/').lower()
    mes_ref = calcular_mes_anterior(target_month)

    ws_resumo.cell(row=2, column=nova_coluna, value=mes_faturado)
    ws_resumo.cell(row=3, column=nova_coluna, value=mes_ref)
    ws_resumo.cell(row=4, column=nova_coluna, value=f"=SUMIF(BASE!$K:$K,RESUMO!{letra}3,BASE!$D:$D)")
    ws_resumo.cell(row=5, column=nova_coluna, value=f"=COUNTIF(BASE!$K:$K,RESUMO!{letra}3)")
    ws_resumo.cell(row=6, column=nova_coluna, value=f"={letra}4*3%")

    col_molde = nova_coluna - 1
    while col_molde >= 1:
        if ws_resumo.cell(row=4, column=col_molde).value is not None: break
        col_molde -= 1
    
    if col_molde >= 1:
        for r in range(2, 7): copiar_estilo(ws_resumo.cell(row=r, column=col_molde), ws_resumo.cell(row=r, column=nova_coluna))
    return nova_coluna

def atualizar_resumo_ciclo_pmt(base_wb, target_month):
    ws_resumo = base_wb['RESUMO']
    mes_faturado = target_month.replace('.', '/').lower()
    
    col_idx = None
    for col in range(1, ws_resumo.max_column + 1):
        val = ws_resumo.cell(row=2, column=col).value
        if val and str(val).strip().lower() == mes_faturado:
            col_idx = col
            break
            
    if not col_idx: raise ValueError(f"Coluna com '{mes_faturado}' não encontrada na aba RESUMO")
    
    meses_eng = {'JAN': 1, 'FEV': 2, 'MAR': 3, 'ABR': 4, 'MAI': 5, 'JUN': 6, 'JUL': 7, 'AGO': 8, 'SET': 9, 'OUT': 10, 'NOV': 11, 'DEZ': 12}
    partes = target_month.upper().strip().split('.')
    data_ref = date(2000 + int(partes[1]), meses_eng[partes[0]], 1)
    data_ref_menos_4 = data_ref - relativedelta(months=4)
    
    data_ini = date(data_ref_menos_4.year, data_ref_menos_4.month, 23)
    data_fim_mes = data_ref_menos_4 + relativedelta(months=1)
    data_fim = date(data_fim_mes.year, data_fim_mes.month, 20)
    
    data_ini_str = data_ini.strftime("%d/%m/%Y")
    data_fim_str = data_fim.strftime("%d/%m/%Y")
    header_str = f"{data_ini.strftime('%d/%m')} a {data_fim.strftime('%d/%m')} - {data_ini.year}"
    letra = get_column_letter(col_idx)
    
    ws_resumo.cell(row=9, column=col_idx, value=header_str)
    ws_resumo.cell(row=10, column=col_idx, value=f'=COUNTIFS(BASE!$H:$H,">={data_ini_str}",BASE!$H:$H,"<={data_fim_str}")')
    ws_resumo.cell(row=11, column=col_idx, value=f'=SUMIFS(BASE!$D:$D,BASE!$H:$H,">={data_ini_str}",BASE!$H:$H,"<={data_fim_str}")')
    ws_resumo.cell(row=12, column=col_idx, value=f"=SUM('{target_month}'!L:L)")
    ws_resumo.cell(row=13, column=col_idx, value=f"=COUNTA('{target_month}'!O:O)-1")
    ws_resumo.cell(row=14, column=col_idx, value=f'=COUNTIFS(\'{target_month}\'!R:R,">={data_ini_str}",\'{target_month}\'!R:R,"<={data_fim_str}")')
    ws_resumo.cell(row=15, column=col_idx, value=f"={letra}13-{letra}14")
    ws_resumo.cell(row=17, column=col_idx, value=f"={letra}14-{letra}10")
    
    celula_esq_18 = ws_resumo.cell(row=18, column=col_idx - 1)
    ws_resumo.cell(row=18, column=col_idx, value=celula_esq_18.value if celula_esq_18.value else None)
    
    col_molde = col_idx - 1
    while col_molde >= 1:
        if ws_resumo.cell(row=10, column=col_molde).value is not None: break
        col_molde -= 1
        
    if col_molde >= 1:
        for r in range(9, 19): copiar_estilo(ws_resumo.cell(row=r, column=col_molde), ws_resumo.cell(row=r, column=col_idx))

def verificar_e_corrigir_headers_regras(ws):
    col_regra = None
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(row=9, column=col).value
        if valor and 'REGRA' in str(valor).upper() and 'PARCELAMENTO' in str(valor).upper():
            col_regra = col
            break
            
    if not col_regra: return
    
    headers = ['CICLO PARCELAS', 'Repasse DataPrev p/Paketa', 'Receita Wiipo']
    for i, header in enumerate(headers, start=1):
        col_atual = col_regra + i
        coord = f"{get_column_letter(col_atual)}9"
        
        for merged_range in list(ws.merged_cells.ranges):
            if coord in merged_range:
                ws.unmerge_cells(str(merged_range))
                if (9, col_atual) in ws._cells: del ws._cells[(9, col_atual)]
                break
                
        celula = ws.cell(row=9, column=col_atual)
        celula.value = header
        copiar_estilo(ws.cell(row=9, column=col_regra), celula)

def atualizar_resumo_bloco_final(base_wb, target_month, col_idx):
    ws = base_wb['RESUMO']
    letra = get_column_letter(col_idx)
    valor_linha2 = ws.cell(row=2, column=col_idx).value or target_month.replace('.', '/').lower()
    linhas_alvo = [20, 21, 22, 23]
    
    for linha_num in linhas_alvo:
        coord = f"{letra}{linha_num}"
        for merged_range in list(ws.merged_cells.ranges):
            if coord in merged_range:
                ws.unmerge_cells(str(merged_range))
                if (linha_num, col_idx) in ws._cells: del ws._cells[(linha_num, col_idx)]
                break
                
    ws.cell(row=20, column=col_idx).value = valor_linha2
    ws.cell(row=21, column=col_idx).value = f"={letra}6"
    ws.cell(row=22, column=col_idx).value = f"={letra}12"
    ws.cell(row=23, column=col_idx).value = f"=SUM({letra}21:{letra}22)"
    
    col_anterior = col_idx - 1
    while col_anterior >= 1:
        if ws.cell(row=20, column=col_anterior).value is not None: break
        col_anterior -= 1
        
    if col_anterior >= 1:
        letra_anterior = get_column_letter(col_anterior)
        ws.column_dimensions[letra].width = ws.column_dimensions[letra_anterior].width
        for r in linhas_alvo:
            if ws.cell(row=r, column=col_anterior).has_style:
                try: copiar_estilo(ws.cell(row=r, column=col_anterior), ws.cell(row=r, column=col_idx))
                except: pass

def copiar_producao_para_base(ws_origem, ws_destino):
    last_row_base = 0
    for row in range(1, ws_destino.max_row + 1):
        if ws_destino.cell(row=row, column=1).value is not None: last_row_base = row
    if last_row_base == 0: last_row_base = 1
    
    new_row = last_row_base + 1
    linhas_copiadas = 0
    
    for source_row in range(2, ws_origem.max_row + 1):
        if ws_origem.cell(row=source_row, column=1).value is None: break
        
        for col in range(1, 8):
            cell_nova = ws_destino.cell(row=new_row, column=col, value=ws_origem.cell(row=source_row, column=col).value)
            if last_row_base > 1: copiar_estilo(ws_destino.cell(row=last_row_base, column=col), cell_nova)
                
        cell_nova = ws_destino.cell(row=new_row, column=8, value=f"=F{new_row}")
        if last_row_base > 1: copiar_estilo(ws_destino.cell(row=last_row_base, column=8), cell_nova)
            
        for origem_col in range(8, 11):
            destino_col = origem_col + 1
            cell_nova = ws_destino.cell(row=new_row, column=destino_col, value=ws_origem.cell(row=source_row, column=origem_col).value)
            if last_row_base > 1: copiar_estilo(ws_destino.cell(row=last_row_base, column=destino_col), cell_nova)
                
        new_row += 1
        linhas_copiadas += 1
    return linhas_copiadas

def validar_template_jan26(workbook):
    template_nome = 'JAN.26'
    if template_nome in workbook.sheetnames: return True, f"Template '{template_nome}' encontrado"
    return False, f"ERRO CRÍTICO: Aba '{template_nome}' não encontrada."

def limpar_dados_worksheet(ws, manter_linha_1=True):
    linha_inicial = 2 if manter_linha_1 else 1
    for row_idx in range(ws.max_row, linha_inicial - 1, -1):
        for col_idx in range(1, ws.max_column + 1): ws.cell(row=row_idx, column=col_idx).value = None

def aplicar_regras_colunas_n_x(ws, target_month, linha_inicio=2):
    mes_faturado = target_month.replace('.', '-').lower()
    ultima_linha = linha_inicio - 1
    for row in range(linha_inicio, ws.max_row + 1):
        if ws.cell(row=row, column=1).value is not None: ultima_linha = row
        else: break
            
    if ultima_linha < linha_inicio: return {'linhas_n_o': 0, 'linhas_q_w': 0, 'ccbs_unicos': 0}
    
    for row in range(linha_inicio, ultima_linha + 1):
        ws.cell(row=row, column=14, value=mes_faturado)
        ws.cell(row=row, column=15, value=f"=VLOOKUP(A{row},'BASE'!A:H,8,0)")
        ws.cell(row=row, column=16, value=None)
        
    ccbs_todos = [ws.cell(row=row, column=1).value for row in range(linha_inicio, ultima_linha + 1) if ws.cell(row=row, column=1).value is not None]
    ccbs_unicos = []
    vistos = set()
    for ccb in ccbs_todos:
        if ccb not in vistos:
            ccbs_unicos.append(ccb)
            vistos.add(ccb)
            
    row_destino = linha_inicio
    for ccb_unico in ccbs_unicos:
        ws.cell(row=row_destino, column=17, value=ccb_unico)
        ws.cell(row=row_destino, column=18, value=f"=VLOOKUP(Q{row_destino},'BASE'!A:K,11,0)")
        ws.cell(row=row_destino, column=19, value=f"=SUMIF(A:A,Q{row_destino},L:L)")
        ws.cell(row=row_destino, column=20, value=f"=VLOOKUP(Q{row_destino},'BASE'!A:H,8,0)")
        for col in [21, 22, 23, 24]: ws.cell(row=row_destino, column=col, value=None)
        row_destino += 1
    return {'linhas_n_o': ultima_linha - linha_inicio + 1, 'linhas_q_w': len(ccbs_unicos), 'ccbs_unicos': len(ccbs_unicos)}

def encontrar_colunas_meses(ws_base):
    colunas_meses = []
    col_data_index = encontrar_coluna_por_header(ws_base, 'DATA') or (ws_base.max_column + 1)
    for col_idx in range(17, col_data_index):
        header = ws_base.cell(row=1, column=col_idx).value
        if header: colunas_meses.append({'nome': header, 'indice': col_idx, 'letra': get_column_letter(col_idx)})
    return colunas_meses

def inserir_coluna_mes(ws_base, target_month, colunas_meses):
    pos_insercao = colunas_meses[-1]['indice'] + 1 if colunas_meses else 17
    ws_base.insert_cols(pos_insercao)
    ws_base.cell(row=1, column=pos_insercao, value=target_month)
    ultima_linha = encontrar_ultima_linha(ws_base)
    for row in range(2, ultima_linha + 1): ws_base.cell(row=row, column=pos_insercao, value=f"=COUNTIF('{target_month}'!A:A,BASE!A{row})")
    return {'nome': target_month, 'indice': pos_insercao, 'letra': get_column_letter(pos_insercao)}

def aplicar_formulas_dinamicas(ws_base, colunas_meses, base_wb):
    ultima_linha = ws_base.max_row
    while ultima_linha > 1 and ws_base.cell(row=ultima_linha, column=1).value is None: ultima_linha -= 1
    if ultima_linha < 2 or not colunas_meses: return 0
    target_month_sheet = colunas_meses[-1]['nome']

    formula_l_limpa = str(ws_base.cell(row=2, column=12).value or "").replace(";", ",")
    nova_formula_l = f'=IF(OR(NOT(ISERROR(VLOOKUP(A2,\'{target_month_sheet}\'!A:A,1,0)))),"Sim","Não")' if not formula_l_limpa.startswith("=") else formula_l_limpa.replace('),"Sim"', f",NOT(ISERROR(VLOOKUP(A2,'{target_month_sheet}'!A:A,1,0)))" + '),"Sim"') if target_month_sheet not in formula_l_limpa else formula_l_limpa

    formula_m_limpa = str(ws_base.cell(row=2, column=13).value or "").replace(";", ",")
    nova_formula_m = '="Pendente de pagamento"' if not formula_m_limpa.startswith("=") else formula_m_limpa.replace('"Pendente de pagamento"', f"IFERROR(VLOOKUP(A2,'{target_month_sheet}'!A:N,14,0), " + '"Pendente de pagamento"') + ")" if target_month_sheet not in formula_m_limpa else formula_m_limpa

    formula_n_limpa = str(ws_base.cell(row=2, column=14).value or "").replace(";", ",")
    nova_formula_n = f"=COUNTIF('{target_month_sheet}'!A:A,BASE!A2)" if not formula_n_limpa.startswith("=") else formula_n_limpa + f"+COUNTIF('{target_month_sheet}'!A:A,BASE!A2)" if target_month_sheet not in formula_n_limpa else formula_n_limpa

    linhas_processadas = 0
    for row in range(2, ultima_linha + 1):
        ws_base.cell(row=row, column=12, value=nova_formula_l.replace("A2", f"A{row}"))
        ws_base.cell(row=row, column=13, value=nova_formula_m.replace("A2", f"A{row}"))
        ws_base.cell(row=row, column=14, value=nova_formula_n.replace("A2", f"A{row}"))
        if row > 2:
            try:
                for col in [12, 13, 14]: copiar_estilo(ws_base.cell(row-1, col), ws_base.cell(row, col))
            except: pass
        linhas_processadas += 1
    return linhas_processadas

def processar_inadimplentes(dados_filtrados, ws_destino, base_wb, nome_coluna_id):
    def limpar_id(valor): return "" if pd.isna(valor) else str(valor).strip().replace('.0', '')
    valores_coluna_q = {limpar_id(c.value) for c in ws_destino['Q'] if c.value is not None}
    ids_novos = dados_filtrados[nome_coluna_id].dropna().apply(limpar_id)
    ids_inadimplentes = [id_val for id_val in ids_novos.unique() if id_val and id_val not in valores_coluna_q]
    
    if ids_inadimplentes:
        if 'INADIMPLENTES' not in base_wb.sheetnames: raise ValueError("Aba 'INADIMPLENTES' não encontrada na planilha.")
        ws_inad, ws_base = base_wb['INADIMPLENTES'], base_wb['BASE']
        linha_destino = ws_inad.max_row + 1
        formula_molde_L, formula_molde_M = str(ws_base.cell(row=2, column=12).value or ""), str(ws_base.cell(row=2, column=13).value or "")

        df_temp = dados_filtrados.copy()
        df_temp['ID_LIMPO'] = df_temp[nome_coluna_id].apply(limpar_id)
        df_inadimplentes = df_temp[df_temp['ID_LIMPO'].isin(ids_inadimplentes)].drop_duplicates(subset=['ID_LIMPO'])

        def extrair_numero(val):
            if pd.isna(val) or str(val).strip().startswith('='): return 0.0
            if isinstance(val, (int, float)): return float(val)
            try: return float(str(val).replace('.', '').replace(',', '.'))
            except: return 0.0
            
        for _, row in df_inadimplentes.iterrows():
            valores_linha = row.iloc[0:16].tolist()
            val_valor_emp, val_parcelas, val_fee, val_recebidas = extrair_numero(valores_linha[3]), extrair_numero(valores_linha[4]), extrair_numero(valores_linha[8]), extrair_numero(valores_linha[13])
            
            valores_linha[7] = valores_linha[5]
            valores_linha[9] = val_valor_emp * val_fee
            valores_linha[14] = (val_recebidas / val_parcelas) if val_parcelas > 0 else 0.0
            valores_linha[15] = max(0, val_parcelas - val_recebidas)

            for col_idx, valor in enumerate(valores_linha, start=1):
                texto_valor = str(valor).strip()
                if col_idx == 12 and formula_molde_L.startswith('='): valor_excel = re.sub(r'\$?[Aa]\$?\d+', f'A{linha_destino}', formula_molde_L)
                elif col_idx == 13 and formula_molde_M.startswith('='): valor_excel = re.sub(r'\$?[Aa]\$?\d+', f'A{linha_destino}', formula_molde_M)
                elif pd.isna(valor) or texto_valor.startswith('='): valor_excel = None
                elif col_idx == 8:
                    if isinstance(valor, str) and 'T' in valor:
                        try: valor_excel = pd.to_datetime(valor.split('T')[0]).date()
                        except: valor_excel = valor.split('T')[0]
                    elif isinstance(valor, pd.Timestamp): valor_excel = valor.date()
                    else: valor_excel = valor
                elif isinstance(valor, pd.Timestamp): valor_excel = valor.to_pydatetime()
                else: valor_excel = valor

                celula_nova = ws_inad.cell(row=linha_destino, column=col_idx, value=valor_excel)
                if linha_destino > 2:
                    try: copiar_estilo(ws_inad.cell(row=linha_destino - 1, column=col_idx), celula_nova)
                    except: pass 
                    
                if col_idx == 6: celula_nova.number_format = 'yyyy-mm-ddThh:mm:ss'
                elif col_idx in [8, 13]: celula_nova.number_format = 'dd/mm/yyyy'
                elif col_idx in [4, 10]: celula_nova.number_format = '#,##0.00'
                elif col_idx in [9, 15]: celula_nova.number_format = '0.00%'
            linha_destino += 1
    if 'INADIMPLENTES' in base_wb.sheetnames: base_wb['INADIMPLENTES'].auto_filter.ref = base_wb['INADIMPLENTES'].dimensions
    return base_wb

def processar_ciclo_validacao(base_df, base_wb, target_month_name, data_inicio, data_fim):
    COLUNA_DATA_LETRA = 'F' 
    ws_destino = base_wb[target_month_name]
    idx_data = column_index_from_string(COLUNA_DATA_LETRA) - 1
    nome_coluna_data = base_df.columns[idx_data]
    nome_coluna_id = base_df.columns[0]
    
    coluna_datas_limpas = pd.to_datetime(base_df[nome_coluna_data], errors='coerce', format='mixed').dt.date
    mask = (coluna_datas_limpas >= data_inicio) & (coluna_datas_limpas <= data_fim)
    dados_filtrados = base_df[mask].copy()
    qtd = len(dados_filtrados)
    
    max_row = ws_destino.max_row
    if max_row >= 2:
        for row in ws_destino.iter_rows(min_row=2, max_row=max_row, min_col=22, max_col=24):
            for cell in row: cell.value = None
            
    linha_atual = 2
    for index, row in dados_filtrados.iterrows():
        ws_destino.cell(row=linha_atual, column=22).value = row[nome_coluna_id]
        ws_destino.cell(row=linha_atual, column=23).value = coluna_datas_limpas[index]
        ws_destino.cell(row=linha_atual, column=24).value = f'=IF(ISNUMBER(MATCH(V{linha_atual},Q:Q,0)),"Sim","Não")'
        if linha_atual > 2:
             try:
                 for col in [22, 23, 24]: copiar_estilo(ws_destino.cell(linha_atual-1, col), ws_destino.cell(linha_atual, col))
             except: pass
        linha_atual += 1
    processar_inadimplentes(dados_filtrados, ws_destino, base_wb, nome_coluna_id)
    return qtd

def aplicar_formulas_estaticas(ws_base, linha_inicio):
    ultima_linha = encontrar_ultima_linha(ws_base)
    col_data_index = encontrar_coluna_por_header(ws_base, 'DATA')
    linhas_processadas = 0
    for row in range(linha_inicio, ultima_linha + 1):
        linha_molde = row - 1
        for col, formula in [(15, f"=N{row}/E{row}"), (16, f"=E{row}-N{row}"), (col_data_index, f'=TEXT(F{row},"dd/mm/aaaa")')]:
            cell = ws_base.cell(row=row, column=col, value=formula)
            if linha_molde >= 2: copiar_estilo(ws_base.cell(row=linha_molde, column=col), cell)
        linhas_processadas += 1
    return linhas_processadas

def atualizar_aba_base(base_wb, parceiro_wb, target_month, linha_inicio_append):
    ws_base, ws_producao = base_wb['BASE'], parceiro_wb['Produção']
    colunas_meses = encontrar_colunas_meses(ws_base)
    col_inserida = inserir_coluna_mes(ws_base, target_month, colunas_meses)
    colunas_meses.append(col_inserida)
    aplicar_formulas_dinamicas(ws_base, colunas_meses, base_wb)
    aplicar_formulas_estaticas(ws_base, linha_inicio_append)

def inserir_dados_colunas_especificas(ws_origem, ws_destino, col_inicio=1, col_fim=13, linha_destino_inicio=2):
    linha_destino = linha_destino_inicio
    for row in ws_origem.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row): continue
        for col_idx in range(col_inicio, min(col_fim + 1, len(row) + 1)):
            ws_destino.cell(row=linha_destino, column=col_idx, value=row[col_idx - 1] if col_idx <= len(row) else None)
        linha_destino += 1

# =========================================================
# 2. INTERFACE DO STREAMLIT
# =========================================================

st.set_page_config(page_title="Faturamento Zen", page_icon="📊", layout="centered")

st.title("📊 Validação de Faturamento")
st.markdown("Processador de comissões turbinado com **16GB de RAM**! 🚀")

with st.form("form_processamento"):
    arquivo_parceiro = st.file_uploader("1️⃣ Arquivo PARCEIRO (.xlsx)", type=["xlsx"])
    arquivo_base = st.file_uploader("2️⃣ Arquivo BASE (.xlsx, .xlsm)", type=["xlsx", "xlsm"])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        target_month = st.text_input("Mês Alvo", value="FEV.26")
    with col2:
        dt_inicio = st.date_input("Início do Ciclo")
    with col3:
        dt_fim = st.date_input("Fim do Ciclo")
        
    submit = st.form_submit_button("Iniciar Processamento", type="primary")

if submit:
    if not arquivo_parceiro or not arquivo_base:
        st.error("⚠️ Por favor, envie as duas planilhas antes de processar.")
    else:
        with st.status("🚀 Processando planilhas com força total...", expanded=True) as status:
            try:
                st.write("📥 Lendo planilhas pesadas para a memória (16GB RAM ativado)...")
                # Lê direto dos arquivos em memória
                parceiro_wb = openpyxl.load_workbook(arquivo_parceiro, data_only=True)
                base_wb = openpyxl.load_workbook(arquivo_base, data_only=False)
                
                st.write("⚙️ Executando validações iniciais...")
                valido, mensagem = validar_abas_necessarias(parceiro_wb, base_wb)
                if not valido: raise ValueError(mensagem)
                    
                template_existe, msg_template = validar_template_jan26(base_wb)
                if not template_existe: raise ValueError(msg_template)

                st.write("🔄 Manipulando abas e copiando dados...")
                if target_month in base_wb.sheetnames: del base_wb[target_month]
                ws_mes = base_wb.copy_worksheet(base_wb['JAN.26'])
                ws_mes.title = target_month

                limpar_dados_worksheet(ws_mes, manter_linha_1=True)
                inserir_dados_colunas_especificas(parceiro_wb['Parcelas Pagas'], ws_mes, 1, 13, 2)
                aplicar_regras_colunas_n_x(ws_mes, target_month, 2)

                ultima_linha_base = encontrar_ultima_linha(base_wb['BASE'])
                linha_inicio_append = ultima_linha_base + 1
                
                copiar_producao_para_base(parceiro_wb['Produção'], base_wb['BASE'])
                atualizar_aba_base(base_wb, parceiro_wb, target_month, linha_inicio_append)

                st.write("🧮 Processando Pandas DataFrame e Inadimplentes...")
                ws_base_ativa = base_wb['BASE']
                data = list(ws_base_ativa.values)
                if data:
                    cols = data[0]
                    base_df_atualizado = pd.DataFrame(data[1:], columns=cols)
                else:
                    base_df_atualizado = pd.DataFrame()

                processar_ciclo_validacao(base_df_atualizado, base_wb, target_month, dt_inicio, dt_fim)

                if 'RESUMO' in base_wb.sheetnames:
                    st.write("📝 Atualizando aba de RESUMO...")
                    coluna_alvo = atualizar_resumo_mes_faturamento(base_wb, target_month)
                    atualizar_resumo_ciclo_pmt(base_wb, target_month)
                    verificar_e_corrigir_headers_regras(base_wb['RESUMO'])
                    atualizar_resumo_bloco_final(base_wb, target_month, col_idx=coluna_alvo)

                st.write("💾 Gerando arquivo final Excel...")
                output = BytesIO()
                base_wb.save(output)
                output.seek(0)
                
                # Esvazia a RAM do servidor
                del parceiro_wb
                del base_wb
                del base_df_atualizado
                gc.collect()
                
                status.update(label="✅ Processamento Concluído com Sucesso!", state="complete", expanded=False)
                
                st.download_button(
                    label="📥 Baixar Excel Processado",
                    data=output,
                    file_name=f"Processado_{target_month}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
            except Exception as e:
                status.update(label="❌ Erro no Processamento", state="error")
                st.error(f"Erro detalhado: {str(e)}")