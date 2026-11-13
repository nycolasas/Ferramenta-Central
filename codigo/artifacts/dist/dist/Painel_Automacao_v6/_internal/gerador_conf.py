import os
import sys
import time
import re
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import numpy as np
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from portable_paths import BASE_WORKBOOK, RESULT_WORKBOOK, ensure_portable_data_files

ensure_portable_data_files()

# ==============================================================================
# FUNÇÕES AUXILIARES DE TRATAMENTO DE TEXTO E COLUNAS
# ==============================================================================
def normalizar_texto(texto):
    """Remove acentos, caracteres especiais e converte para maiúsculas."""
    if not isinstance(texto, str):
        texto = str(texto)
    nfkd = unicodedata.normalize('NFKD', texto)
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return sem_acento.strip().upper()

def carregar_aba_flexivel(caminho, nome_aba):
    """Carrega uma aba da planilha tratando diferenças de caixa e ignorando linhas em branco no topo."""
    if not os.path.exists(caminho):
        return pd.DataFrame()
    try:
        xls = pd.ExcelFile(caminho)
        abas_existentes = xls.sheet_names
        
        aba_encontrada = None
        for aba in abas_existentes:
            if normalizar_texto(aba) == normalizar_texto(nome_aba):
                aba_encontrada = aba
                break
                
        if not aba_encontrada:
            print(f"⚠️ Aba '{nome_aba}' não encontrada na planilha. Usando estrutura vazia de fallback.")
            return pd.DataFrame()
            
        df = pd.read_excel(caminho, sheet_name=aba_encontrada)
        df.dropna(how='all', inplace=True)
        return df
    except Exception as e:
        print(f"⚠️ Erro ao carregar aba '{nome_aba}': {e}")
        return pd.DataFrame()

def buscar_coluna_segura(df, termos_busca, coluna_padrao=""):
    """Busca uma coluna no DataFrame. Se não encontrar ou o DF estiver vazio, retorna a coluna padrão ou None."""
    if df.empty or len(df.columns) == 0:
        return None
    
    colunas_norm = {col: normalizar_texto(col) for col in df.columns}
    
    for termo in termos_busca:
        termo_norm = normalizar_texto(termo)
        for col_orig, col_n in colunas_norm.items():
            if termo_norm in col_n:
                return col_orig
                
    # Fallback para primeira coluna disponível se não encontrar por nome
    return df.columns[0] if len(df.columns) > 0 else None

# ==============================================================================
# 1. CARREGAMENTO DAS BASES E LIMPEZA INICIAL
# ==============================================================================
caminho_arquivo = str(BASE_WORKBOOK)
caminho_saida = str(RESULT_WORKBOOK)

print("📂 Carregando os dados...")
df_saldo = carregar_aba_flexivel(caminho_arquivo, 'f_saldo')
df_4260 = carregar_aba_flexivel(caminho_arquivo, 'f_4260')
df_vt = carregar_aba_flexivel(caminho_arquivo, 'f_vt')
df_registro = carregar_aba_flexivel(caminho_arquivo, 'f_registro')
df_desconsiderar = carregar_aba_flexivel(caminho_arquivo, 'desconsiderar')
df_canais = carregar_aba_flexivel(caminho_arquivo, 'canais')
df_atendente = carregar_aba_flexivel(caminho_arquivo, 'atendente')

# Tenta carregar o histórico anterior do resultado_processamento.xlsx
df_hist_saldo = pd.DataFrame()
if os.path.exists(caminho_saida):
    try:
        df_hist_saldo = pd.read_excel(caminho_saida, sheet_name='HISTORICO SALDO')
        print("📜 Histórico anterior carregado com sucesso a partir do resultado_processamento.xlsx!")
    except Exception:
        pass

if df_hist_saldo.empty:
    df_hist_saldo = carregar_aba_flexivel(caminho_arquivo, 'HISTORICO SALDO')
    if not df_hist_saldo.empty:
        print("📜 Histórico carregado a partir da planilha_bruta.xlsx.")
    else:
        print("⚠️ Aba HISTORICO SALDO não encontrada. Um novo histórico será criado.")

# Limpando espaços ocultos nos cabeçalhos
for df in [df_saldo, df_4260, df_vt, df_registro, df_desconsiderar, df_hist_saldo, df_canais, df_atendente]:
    if not df.empty:
        df.columns = df.columns.str.strip()

# EXTRAÇÃO LIMPA DOS DÍGITOS NUMÉRICOS DO PEDIDO
def extrair_pedido_limpo(serie):
    if serie is None or serie.empty:
        return pd.Series(dtype=str)
    s = serie.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    num = s.str.extract(r'(\d+)\s*$', expand=False)
    return num.str.lstrip('0').fillna('')

def achar_coluna_pedido(df):
    if df.empty: return None
    for col in ['Pedido', 'PEDIDO', 'Pedido Venda', 'Nr. Pedido', 'Nº Pedido', 'Num. Pedido', 'Pedido_Venda', 'Local - Sit']:
        if col in df.columns: return col
    cols = [c for c in df.columns if 'PEDIDO' in c.upper()]
    return cols[0] if cols else df.columns[0]

col_ped_4260 = achar_coluna_pedido(df_4260)
col_ped_reg = achar_coluna_pedido(df_registro)
col_ped_des = achar_coluna_pedido(df_desconsiderar)
col_ped_vt = achar_coluna_pedido(df_vt)

# Filtro da Empresa 11 + Chave Numérica Limpa
if not df_saldo.empty and 'Pedido' in df_saldo.columns:
    cond_emp_11 = df_saldo['Pedido'].astype(str).str.contains(r'^\(?11\)?', regex=True, na=False)
    df_saldo = df_saldo[cond_emp_11].copy()
    df_saldo['Pedido_Chave'] = extrair_pedido_limpo(df_saldo['Pedido'])
    df_saldo['Pedido'] = '(11) ' + df_saldo['Pedido_Chave']
else:
    df_saldo['Pedido_Chave'] = ''

df_4260['Pedido_Chave'] = extrair_pedido_limpo(df_4260[col_ped_4260]) if col_ped_4260 else ''
df_registro['Pedido_Chave'] = extrair_pedido_limpo(df_registro[col_ped_reg]) if col_ped_reg else ''
df_desconsiderar['Pedido_Chave'] = extrair_pedido_limpo(df_desconsiderar[col_ped_des]) if col_ped_des else ''
df_vt['Pedido_Chave'] = extrair_pedido_limpo(df_vt[col_ped_vt]) if col_ped_vt else ''

if not df_saldo.empty:
    if 'Situação' in df_saldo.columns:
        df_saldo['Situação'] = df_saldo['Situação'].astype(str).str.strip().str.upper()
    if 'Cod. Item' in df_saldo.columns:
        df_saldo['Cod. Item'] = df_saldo['Cod. Item'].astype(str).str.strip()
    if 'Linha de Produto' in df_saldo.columns:
        df_saldo['Linha de Produto'] = df_saldo['Linha de Produto'].astype(str).str.strip().str.upper()

    # FILTRO 1: REMOVER ITENS QUE COMEÇAM COM "3TAXA"
    cond_nao_taxa = ~df_saldo['Cod. Item'].str.upper().str.startswith('3TAXA', na=False)
    df_saldo = df_saldo[cond_nao_taxa].copy()

col_data_emissao = None
for c in ['Data de Emissão', 'Data Emissão', 'DATA DE EMISSAO', 'DATA EMISSAO', 'Emissão', 'Data']:
    if c in df_4260.columns:
        col_data_emissao = c
        break

if col_data_emissao:
    df_4260['Data de Emissão'] = pd.to_datetime(df_4260[col_data_emissao], dayfirst=True, errors='coerce').dt.normalize()

if not df_saldo.empty:
    df_saldo['Chave_Item'] = df_saldo['Pedido_Chave'] + "_" + df_saldo['Cod. Item']

df_4260_dedup = df_4260[df_4260['Pedido_Chave'] != ''].drop_duplicates(subset=['Pedido_Chave'], keep='first') if not df_4260.empty else pd.DataFrame(columns=['Pedido_Chave', 'Data de Emissão'])

# ==============================================================================
# 1.1 MAPEAMENTO DE ATENDENTE E REGIÃO
# ==============================================================================
print("👤 Mapeando Atendentes e Região...")

dict_rep_para_loja = {}
dict_rep_para_regiao = {}
dict_loja_soltos = {}
dict_loja_planejados = {}

if not df_canais.empty:
    col_rep_canais = buscar_coluna_segura(df_canais, ['REPRESENTANTE', 'REP', 'VENDEDOR'])
    col_loja_canais = buscar_coluna_segura(df_canais, ['LOJA', 'CANAL'])
    col_regiao_canais = buscar_coluna_segura(df_canais, ['REGIAO', 'REGIONAL', 'UF'])

    if col_rep_canais and col_loja_canais and col_regiao_canais:
        df_canais['Rep_Clean'] = pd.to_numeric(df_canais[col_rep_canais], errors='coerce')
        df_canais['Loja_Clean'] = df_canais[col_loja_canais].astype(str).str.strip().str.upper()
        df_canais['Regiao_Clean'] = df_canais[col_regiao_canais].astype(str).str.strip()

        dict_rep_para_loja = dict(zip(df_canais['Rep_Clean'], df_canais['Loja_Clean']))
        dict_rep_para_regiao = dict(zip(df_canais['Rep_Clean'], df_canais['Regiao_Clean']))
else:
    print("⚠️ Aba 'canais' está vazia ou ausente. O mapeamento de Loja/Região utilizará valores padrão.")

if not df_atendente.empty:
    col_loja_atend = buscar_coluna_segura(df_atendente, ['LOJA', 'CANAL'])
    col_soltos_atend = buscar_coluna_segura(df_atendente, ['SOLTO', 'SOLTOS', 'PADRAO'])
    col_plan_atend = buscar_coluna_segura(df_atendente, ['PLANEJADO', 'PLANEJADOS', 'PROJETO'])

    if col_loja_atend and col_soltos_atend and col_plan_atend:
        df_atendente['Loja_Clean'] = df_atendente[col_loja_atend].astype(str).str.strip().str.upper()
        dict_loja_soltos = dict(zip(df_atendente['Loja_Clean'], df_atendente[col_soltos_atend].astype(str).str.strip()))
        dict_loja_planejados = dict(zip(df_atendente['Loja_Clean'], df_atendente[col_plan_atend].astype(str).str.strip()))
else:
    print("⚠️ Aba 'atendente' está vazia ou ausente. O mapeamento de atendentes utilizará valores padrão.")

# Identifica pedidos com item planejado
pedidos_com_planejados = set(
    df_saldo.loc[df_saldo['Linha de Produto'] == '4-PLANEJADOS', 'Pedido_Chave']
) if not df_saldo.empty else set()

if not df_saldo.empty:
    df_saldo['Tipo_Produto'] = np.where(
        df_saldo['Pedido_Chave'].isin(pedidos_com_planejados), 
        'PLANEJADOS', 
        'SOLTOS'
    )

    df_saldo['Rep_Num'] = pd.to_numeric(df_saldo['Representante'], errors='coerce') if 'Representante' in df_saldo.columns else np.nan
    df_saldo['Loja'] = df_saldo['Rep_Num'].map(dict_rep_para_loja).fillna('')
    df_saldo['Região'] = df_saldo['Rep_Num'].map(dict_rep_para_regiao).fillna('')

    def definir_atendente(row):
        loja = row['Loja']
        tipo = row['Tipo_Produto']
        if tipo == 'PLANEJADOS':
            return dict_loja_planejados.get(loja, '')
        else:
            return dict_loja_soltos.get(loja, '')

    df_saldo['Atendente'] = df_saldo.apply(definir_atendente, axis=1)

# ==============================================================================
# 1.2 CORREÇÃO DE ERRO HUMANO: PLANEJADOS EM 'N' COM DATA DE ENTREGA 2040
# ==============================================================================
col_data_entrega = None
if not df_saldo.empty:
    for c in ['Data Entrega', 'Dt. Previsão', 'Data Previsão', 'Data Prev. Entrega', 'Previsão Entrega', 'DT ENTREGA', 'DATA ENTREGA']:
        if c in df_saldo.columns:
            col_data_entrega = c
            break

    if col_data_entrega is None:
        cols_ent = [c for c in df_saldo.columns if 'ENTREGA' in c.upper() or 'PREVIS' in c.upper()]
        if cols_ent:
            col_data_entrega = cols_ent[0]

df_erro_2040 = pd.DataFrame()

if not df_saldo.empty and col_data_entrega:
    df_saldo['Ano_Entrega'] = pd.to_datetime(df_saldo[col_data_entrega], dayfirst=True, errors='coerce').dt.year
    
    cond_erro_2040 = (
        (df_saldo['Tipo_Produto'] == 'PLANEJADOS') & 
        (df_saldo['Situação'] == 'N') & 
        (df_saldo['Ano_Entrega'] == 2040)
    )
    
    df_erro_2040 = df_saldo[cond_erro_2040].copy()
    
    if not df_erro_2040.empty:
        df_erro_2040['Situação Original ERP'] = 'N'
        df_erro_2040['Situação Ajustada Python'] = 'S'
        df_erro_2040['Data Entrega ERP'] = pd.to_datetime(df_erro_2040[col_data_entrega], errors='coerce').dt.strftime('%d/%m/%Y')
        
        cols_erro_export = ['Cliente', 'Pedido', 'Cod. Item', 'Linha de Produto', 'Atendente', 'Situação Original ERP', 'Situação Ajustada Python', 'Data Entrega ERP']
        cols_existentes = [c for c in cols_erro_export if c in df_erro_2040.columns]
        df_erro_2040 = df_erro_2040[cols_existentes].copy()

    df_saldo.loc[cond_erro_2040, 'Situação'] = 'S'
    print(f"🛠️ Correção virtual aplicada: {cond_erro_2040.sum()} itens de planejados em 'N' com data = 2040 convertidos temporariamente para 'S'.")

# ==============================================================================
# 2. IDENTIFICAÇÃO DE ITENS QUE SAÍRAM DE 'S'
# ==============================================================================
chaves_sairam_de_S = set()
dict_status_anterior_s = {}

if not df_hist_saldo.empty and 'Data Saldo' in df_hist_saldo.columns:
    df_hist_saldo['Pedido_Chave'] = extrair_pedido_limpo(df_hist_saldo['Pedido'])
    df_hist_saldo['Cod. Item'] = df_hist_saldo['Cod. Item'].astype(str).str.strip()
    df_hist_saldo['Situação'] = df_hist_saldo['Situação'].astype(str).str.strip().str.upper()
    df_hist_saldo['Chave_Item'] = df_hist_saldo['Pedido_Chave'] + "_" + df_hist_saldo['Cod. Item']
    df_hist_saldo['Data Saldo'] = pd.to_datetime(df_hist_saldo['Data Saldo'], errors='coerce')

    hoje_dt = pd.Timestamp('today').normalize()
    df_hist_anterior = df_hist_saldo[df_hist_saldo['Data Saldo'] < hoje_dt].copy()

    if not df_hist_anterior.empty and not df_saldo.empty:
        df_ultimo_status = (
            df_hist_anterior.sort_values(by=['Chave_Item', 'Data Saldo'], ascending=[True, True])
            .drop_duplicates(subset=['Chave_Item'], keep='last')
        )
        dict_status_anterior_s = dict(zip(df_ultimo_status['Chave_Item'], df_ultimo_status['Situação']))

        for _, row in df_saldo.iterrows():
            chave = row['Chave_Item']
            status_atual = row['Situação']
            status_ant = dict_status_anterior_s.get(chave, None)

            if status_ant == 'S' and status_atual != 'S':
                chaves_sairam_de_S.add(chave)

# ==============================================================================
# 3. REGRA D-2 E MERGE DE DATAS
# ==============================================================================
hoje = pd.Timestamp('today').normalize()

eh_terca = hoje.weekday() == 1

if eh_terca:
    data_inicio_d2 = hoje - pd.Timedelta(days=4)
    data_fim_d2 = hoje - pd.Timedelta(days=2)
    data_corte_atrasados = data_fim_d2
else:
    data_corte = hoje - pd.offsets.CustomBusinessDay(n=2)
    data_inicio_d2 = data_corte
    data_fim_d2 = data_corte
    data_corte_atrasados = data_corte

DATA_LIMITE_ALEM_D2 = pd.Timestamp('2026-07-31')

if not df_saldo.empty:
    df_saldo_completo = pd.merge(
        df_saldo, 
        df_4260_dedup[['Pedido_Chave', 'Data de Emissão']], 
        on='Pedido_Chave', 
        how='left'
    )

    df_atrasados = df_saldo_completo[
        (df_saldo_completo['Data de Emissão'] <= data_corte_atrasados) | (df_saldo_completo['Data de Emissão'].isna())
    ].copy()

    cond_normais = (df_atrasados['Situação'] == 'N') & (~df_atrasados['Pedido_Chave'].isin(pedidos_com_planejados))
    cond_planejados_etapa1 = (df_atrasados['Situação'] == 'S') & (df_atrasados['Pedido_Chave'].isin(pedidos_com_planejados))
    cond_planejados_etapa2 = (df_atrasados['Situação'] == 'N') & (df_atrasados['Pedido_Chave'].isin(pedidos_com_planejados))

    df_saldo_cliente = df_atrasados[cond_normais | cond_planejados_etapa1 | cond_planejados_etapa2].copy()
    df_saldo_cliente['Chave_Etapa'] = df_saldo_cliente['Pedido_Chave'] + "_" + df_saldo_cliente['Situação']
else:
    df_saldo_cliente = pd.DataFrame()

# ==============================================================================
# 4. FILTRO DE EXCLUSÃO DOS ITENS VT
# ==============================================================================
if not df_vt.empty:
    df_vt['Item_Clean'] = df_vt['Item'].astype(str).str.strip() if 'Item' in df_vt.columns else ''
    chaves_vt = set(df_vt['Pedido_Chave'] + "_" + df_vt['Item_Clean'])
else:
    chaves_vt = set()

if not df_saldo_cliente.empty:
    df_saldo_cliente['Item_Clean'] = df_saldo_cliente['Cod. Item'].astype(str).str.strip()
    df_saldo_cliente['Chave_Temp'] = df_saldo_cliente['Pedido_Chave'] + "_" + df_saldo_cliente['Item_Clean']

    df_confirmacao_saldo = df_saldo_cliente[~df_saldo_cliente['Chave_Temp'].isin(chaves_vt)].copy()
    df_confirmacao_saldo.drop(columns=['Item_Clean', 'Chave_Temp'], inplace=True)
else:
    df_confirmacao_saldo = pd.DataFrame()

# ==============================================================================
# 5. MAPEAMENTO DA ABA "DESCONSIDERAR"
# ==============================================================================
if not df_desconsiderar.empty and 'MOTIVO' in df_desconsiderar.columns:
    df_desconsiderar['MOTIVO'] = df_desconsiderar['MOTIVO'].astype(str).str.strip().str.upper()

    motivos_bloqueio = ['ENTREGA AGENDADA', 'JÁ FATURADO', 'JA FATURADO']
    pedidos_desconsiderar = set(
        df_desconsiderar.loc[df_desconsiderar['MOTIVO'].isin(motivos_bloqueio), 'Pedido_Chave']
    )
else:
    pedidos_desconsiderar = set()

# ==============================================================================
# 6. REGRA DE VALIDAÇÃO (F_REGISTRO)
# ==============================================================================
print("🔍 Validando confirmações na aba f_registro...")
dict_ultimo_motivo = {}
pedidos_confirmados_registro = set()
chaves_etapa_confirmadas_registro = set()

if not df_registro.empty:
    col_motivo = None
    for col in ['Motivo', 'MOTIVO', 'motivo', 'Registro', 'Observação', 'OBSERVAÇÃO']:
        if col in df_registro.columns:
            col_motivo = col
            break
    if col_motivo is None:
        col_motivo = df_registro.columns[9] if len(df_registro.columns) > 9 else df_registro.columns[-1]

    col_tipo_info = None
    for col in ['Tipo informação', 'Tipo Informação', 'TIPO INFORMAÇÃO', 'Tipo Informacao', 'TIPO INFORMACAO']:
        if col in df_registro.columns:
            col_tipo_info = col
            break

    df_registro['Motivo_Clean'] = df_registro[col_motivo].astype(str).str.strip().str.upper()
    df_registro['Tipo_Info_Clean'] = df_registro[col_tipo_info].astype(str).str.strip().str.upper() if col_tipo_info else ''
    df_registro['Situação'] = df_registro['Situação'].astype(str).str.strip().str.upper() if 'Situação' in df_registro.columns else 'N'
    df_registro['Chave_Etapa'] = df_registro['Pedido_Chave'] + "_" + df_registro['Situação']

    motivos_diretos = ['AGUARDANDO CARGA PARA O INTERIOR', 'AGUARDANDO CARGA', 'AGENDADO', 'ATRASO ENTREGA AVISADO']
    padrao_direto = '|'.join(motivos_diretos)
    cond_excluir_direto = df_registro['Motivo_Clean'].str.contains(padrao_direto, regex=True, na=False)

    cond_motivo_confirmacao = df_registro['Motivo_Clean'].str.contains(r'CONFIRMA[CÇ][AÃ]O DE DATA DE ENTREGA', regex=True, na=False)
    cond_tipo_contato = df_registro['Tipo_Info_Clean'] == 'CONTATO CLIENTE'
    cond_excluir_confirmacao = cond_motivo_confirmacao & cond_tipo_contato

    cond_confirmado = cond_excluir_direto | cond_excluir_confirmacao

    pedidos_confirmados_registro = set(df_registro.loc[cond_confirmado, 'Pedido_Chave'].dropna().unique())
    chaves_etapa_confirmadas_registro = set(df_registro.loc[cond_confirmado, 'Chave_Etapa'].dropna().unique())

    if 'Data Contato' in df_registro.columns:
        df_registro['Data Contato'] = pd.to_datetime(df_registro['Data Contato'], errors='coerce')
        df_reg_ordenado = df_registro.sort_values(by=['Pedido_Chave', 'Data Contato'], ascending=[True, True])
    else:
        df_reg_ordenado = df_registro.copy()

    df_ultimo_motivo_ped = df_reg_ordenado[df_reg_ordenado['Pedido_Chave'] != ''].drop_duplicates(subset=['Pedido_Chave'], keep='last')
    dict_ultimo_motivo = dict(zip(df_ultimo_motivo_ped['Pedido_Chave'], df_ultimo_motivo_ped[col_motivo]))

# ==============================================================================
# 7. GERAR BASE GERAL DE CONFIRMAÇÃO
# ==============================================================================
if not df_confirmacao_saldo.empty:
    nat_oper_clean = df_confirmacao_saldo['Nat. Oper'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True) if 'Nat. Oper' in df_confirmacao_saldo.columns else pd.Series('')
    nat_permitidas = ['7002', '7005', '8025', '7025', '7004']
    cond_natureza = nat_oper_clean.isin(nat_permitidas)

    prefixo_linha = df_confirmacao_saldo['Linha de Produto'].astype(str).str.strip().str[:2]
    cond_linha = ~prefixo_linha.isin(['24', '25', '99'])

    rep_numerico = pd.to_numeric(df_confirmacao_saldo['Representante'], errors='coerce') if 'Representante' in df_confirmacao_saldo.columns else pd.Series(np.nan)
    cond_representante = ~((rep_numerico >= 5000) & (rep_numerico <= 5999))

    cond_desconsiderar = ~df_confirmacao_saldo['Pedido_Chave'].isin(pedidos_desconsiderar)

    eh_planejado = df_confirmacao_saldo['Pedido_Chave'].isin(pedidos_com_planejados)

    cond_confirmado_normal = (~eh_planejado) & (df_confirmacao_saldo['Pedido_Chave'].isin(pedidos_confirmados_registro))
    cond_confirmado_planejado = (eh_planejado) & (df_confirmacao_saldo['Chave_Etapa'].isin(chaves_etapa_confirmadas_registro))

    cond_pendente_registro = ~(cond_confirmado_normal | cond_confirmado_planejado)

    df_confirmacao_base = df_confirmacao_saldo[
        cond_natureza & cond_linha & cond_representante & cond_desconsiderar & cond_pendente_registro
    ].copy()

    df_confirmacao_base.drop_duplicates(subset=['Pedido_Chave', 'Situação'], keep='first', inplace=True)

    def definir_rotulo_confirmacao(row):
        if row['Tipo_Produto'] == 'PLANEJADOS':
            if row['Situação'] == 'S':
                return '1ª Confirmação (S)'
            elif row['Situação'] == 'N':
                return '2ª Confirmação (N)'
        return 'Confirmação Única (N)'

    df_confirmacao_base['Etapa_Confirmacao'] = df_confirmacao_base.apply(definir_rotulo_confirmacao, axis=1)
else:
    df_confirmacao_base = pd.DataFrame()

# ==============================================================================
# FUNÇÕES AUXILIARES PARA AS ABAS FINAIS
# ==============================================================================
def desambiguar_nomes_duplicados(df):
    if df.empty or 'Cliente' not in df.columns:
        return df

    contagem = df['Cliente'].value_counts()
    nomes_duplicados = set(contagem[contagem > 1].index)

    if not nomes_duplicados:
        return df

    contadores = {nome: 0 for nome in nomes_duplicados}
    novos_nomes = []

    for nome in df['Cliente']:
        if nome in nomes_duplicados:
            contadores[nome] += 1
            novos_nomes.append(f"{nome} ({contadores[nome]})")
        else:
            novos_nomes.append(nome)

    df['Cliente'] = novos_nomes
    return df

def filtrar_motivos_exclusao(df):
    if df.empty or 'Último Motivo' not in df.columns:
        return df

    tabela_acentos = str.maketrans({
        'À': 'A', 'Á': 'A', 'Â': 'A', 'Ã': 'A', 'Ä': 'A',
        'È': 'E', 'É': 'E', 'Ê': 'E', 'Ë': 'E',
        'Ì': 'I', 'Í': 'I', 'Î': 'I', 'Ï': 'I',
        'Ò': 'O', 'Ó': 'O', 'Ô': 'O', 'Õ': 'O', 'Ö': 'O',
        'Ù': 'U', 'Ú': 'U', 'Û': 'U', 'Ü': 'U',
        'Ç': 'C', 'Ñ': 'N',
        'à': 'A', 'á': 'A', 'â': 'A', 'ã': 'A', 'ä': 'A',
        'è': 'E', 'é': 'E', 'ê': 'E', 'ë': 'E',
        'ì': 'I', 'í': 'I', 'î': 'I', 'ï': 'I',
        'ò': 'O', 'ó': 'O', 'ô': 'O', 'õ': 'O', 'ö': 'O',
        'ù': 'U', 'ú': 'U', 'û': 'U', 'ü': 'U',
        'ç': 'C', 'ñ': 'N',
        'º': ' ', 'ª': ' ',
        '-': ' ', '/': ' ', '.': ' ', ',': ' ', '_': ' ', '%': ' ',
        ';': ' ', ':': ' ', '(': ' ', ')': ' ', '[': ' ', ']': ' ',
        '\t': ' ', '\n': ' ', '\r': ' '
    })

    def normalizar_motivo(value):
        if pd.isna(value):
            return ''
        texto = str(value).translate(tabela_acentos).upper()
        texto = re.sub(r'[^A-Z0-9]+', ' ', texto)
        texto = re.sub(r'\s+', ' ', texto).strip()
        return texto

    motivos_proibidos = {
        normalizar_motivo('CLIENTE NÃO QUER RECEBER - OBRAS'),
        normalizar_motivo('CLIENTE NÃO QUER RECEBER/OBRAS'),
        normalizar_motivo('AGUARDANDO CARGA'),
        normalizar_motivo('CONFIRMAÇÃO DE DATA DE ENTREGA'),
        normalizar_motivo('CLIENTE NÃO QUER RECEBER - MOTIVO P'),
        normalizar_motivo('BOLETO 50%')
    }

    motivos_normalizados = df['Último Motivo'].map(normalizar_motivo)
    cond_excluir = motivos_normalizados.isin(motivos_proibidos)
    return df[~cond_excluir].copy()

# ==============================================================================
# 8. SEPARAÇÃO DAS ABAS D-2 E ALÉM DO D-2
# ==============================================================================
if not df_confirmacao_base.empty and 'Data de Emissão' in df_confirmacao_base.columns:
    datas_emissao_norm = pd.to_datetime(df_confirmacao_base['Data de Emissão'], errors='coerce').dt.normalize()

    cond_alem_d2 = datas_emissao_norm <= DATA_LIMITE_ALEM_D2
    df_confirmacao_alem_d2 = df_confirmacao_base[cond_alem_d2].copy()

    cond_eh_d2 = datas_emissao_norm > DATA_LIMITE_ALEM_D2
    df_confirmacao_d2 = df_confirmacao_base[cond_eh_d2].copy()
else:
    df_confirmacao_alem_d2 = pd.DataFrame()
    df_confirmacao_d2 = pd.DataFrame()

df_relatorio_alem_d2 = pd.DataFrame()
if not df_confirmacao_alem_d2.empty:
    df_relatorio_alem_d2['Cliente'] = df_confirmacao_alem_d2['Cliente']
    df_relatorio_alem_d2['Pedido'] = df_confirmacao_alem_d2['Pedido']
    df_relatorio_alem_d2['Item'] = df_confirmacao_alem_d2['Cod. Item']
    df_relatorio_alem_d2['Tipo'] = df_confirmacao_alem_d2['Tipo_Produto']
    df_relatorio_alem_d2['Etapa / Confirmação'] = df_confirmacao_alem_d2['Etapa_Confirmacao']
    df_relatorio_alem_d2['Atendente'] = df_confirmacao_alem_d2['Atendente']
    df_relatorio_alem_d2['Data Emissão'] = pd.to_datetime(df_confirmacao_alem_d2['Data de Emissão']).dt.strftime('%d/%m/%Y')
    df_relatorio_alem_d2['Último Motivo'] = df_confirmacao_alem_d2['Pedido_Chave'].map(dict_ultimo_motivo).fillna('')
    df_relatorio_alem_d2['Data Envio'] = hoje.strftime('%d/%m/%Y')
    df_relatorio_alem_d2['Região'] = df_confirmacao_alem_d2['Região']

    df_relatorio_alem_d2 = filtrar_motivos_exclusao(df_relatorio_alem_d2)
    df_relatorio_alem_d2 = desambiguar_nomes_duplicados(df_relatorio_alem_d2)

df_relatorio_d2 = pd.DataFrame()
if not df_confirmacao_d2.empty:
    df_relatorio_d2['Cliente'] = df_confirmacao_d2['Cliente']
    df_relatorio_d2['Pedido'] = df_confirmacao_d2['Pedido']
    df_relatorio_d2['Item'] = df_confirmacao_d2['Cod. Item']
    df_relatorio_d2['Tipo'] = df_confirmacao_d2['Tipo_Produto']
    df_relatorio_d2['Etapa / Confirmação'] = df_confirmacao_d2['Etapa_Confirmacao']
    df_relatorio_d2['Atendente'] = df_confirmacao_d2['Atendente']
    df_relatorio_d2['Data Emissão'] = pd.to_datetime(df_confirmacao_d2['Data de Emissão']).dt.strftime('%d/%m/%Y')
    df_relatorio_d2['Último Motivo'] = df_confirmacao_d2['Pedido_Chave'].map(dict_ultimo_motivo).fillna('')
    df_relatorio_d2['Data Envio'] = hoje.strftime('%d/%m/%Y')
    df_relatorio_d2['Região'] = df_confirmacao_d2['Região']

    df_relatorio_d2 = filtrar_motivos_exclusao(df_relatorio_d2)
    df_relatorio_d2 = desambiguar_nomes_duplicados(df_relatorio_d2)

df_sairam_s = df_saldo[df_saldo['Chave_Item'].isin(chaves_sairam_de_S)].copy() if not df_saldo.empty else pd.DataFrame()

df_relatorio_sairam_s = pd.DataFrame()
if not df_sairam_s.empty:
    df_relatorio_sairam_s['Cliente'] = df_sairam_s['Cliente']
    df_relatorio_sairam_s['Pedido'] = df_sairam_s['Pedido']
    df_relatorio_sairam_s['Item'] = df_sairam_s['Cod. Item']
    df_relatorio_sairam_s['Tipo'] = df_sairam_s['Tipo_Produto']
    df_relatorio_sairam_s['Atendente'] = df_sairam_s['Atendente']
    df_relatorio_sairam_s['Status Anterior'] = 'S'
    df_relatorio_sairam_s['Status Atual'] = df_sairam_s['Situação']
    df_relatorio_sairam_s['Data Alteração'] = hoje.strftime('%d/%m/%Y')

print(f"\n📊 RESULTADO FINAL DAS ABAS APÓS NOVA REGRA DE CORTE DE DATAS E MOTIVOS:")
print(f"📌 Regra aplicada para 'Além do D-2': Pedidos com data de emissão até {DATA_LIMITE_ALEM_D2.strftime('%d/%m/%Y')}")
print(f"✅ Pedidos na aba CONFIRMACAO_D2 (Agosto em diante): {len(df_relatorio_d2)}")
print(f"✅ Pedidos na aba CONFIRMACAO_ALEM_D2 (Até 31/07/2026): {len(df_relatorio_alem_d2)}")
print(f"✅ Pedidos na aba SAIRAM_DE_S: {len(df_relatorio_sairam_s)}")
print(f"⚠️ Pedidos com ERRO DE SITUAÇÃO (2040): {len(df_erro_2040)}")

for df_temp in [df_confirmacao_base, df_confirmacao_saldo, df_saldo_cliente]:
    if 'Pedido_Chave' in df_temp.columns:
        df_temp.drop(columns=['Pedido_Chave'], inplace=True)

# ==============================================================================
# 9. ATUALIZAR E CONSOLIDAR ABA "HISTORICO SALDO"
# ==============================================================================
if not df_saldo.empty:
    df_snapshot_hoje = df_saldo[['Pedido', 'Cod. Item', 'Situação', 'Linha de Produto']].copy()
    df_snapshot_hoje['Data Saldo'] = hoje.strftime('%Y-%m-%d')
else:
    df_snapshot_hoje = pd.DataFrame(columns=['Pedido', 'Cod. Item', 'Situação', 'Linha de Produto', 'Data Saldo'])

if not df_hist_saldo.empty and 'Data Saldo' in df_hist_saldo.columns:
    df_hist_saldo['Data Saldo'] = pd.to_datetime(df_hist_saldo['Data Saldo'], errors='coerce').dt.strftime('%Y-%m-%d')

df_hist_saldo_atualizado = pd.concat([df_hist_saldo, df_snapshot_hoje], ignore_index=True)
df_hist_saldo_atualizado.drop_duplicates(subset=['Pedido', 'Cod. Item', 'Data Saldo'], keep='last', inplace=True)

for df_temp in [df_saldo_cliente, df_confirmacao_saldo, df_confirmacao_base]:
    for col in df_temp.select_dtypes(include=['datetime64[ns]', 'datetime64']).columns:
        df_temp[col] = df_temp[col].dt.strftime('%d/%m/%Y')

# ==============================================================================
# 10. EXPORTAR RESULTADOS
# ==============================================================================
print(f"\n💾 Salvando resultados em: {caminho_saida}")

if os.path.exists(caminho_saida):
    try:
        os.remove(caminho_saida)
    except PermissionError:
        print("❌ ERRO: O arquivo 'resultado_processamento.xlsx' está aberto no Excel! Feche-o e tente novamente.")
        sys.exit(1)

with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
    df_relatorio_d2.to_excel(writer, sheet_name='CONFIRMACAO_D2', index=False)
    df_relatorio_alem_d2.to_excel(writer, sheet_name='CONFIRMACAO_ALEM_D2', index=False)
    df_relatorio_sairam_s.to_excel(writer, sheet_name='SAIRAM_DE_S', index=False)
    df_erro_2040.to_excel(writer, sheet_name='ERRO_CADASTRO_2040', index=False)
    df_confirmacao_base.to_excel(writer, sheet_name='CONFIRMAÇÃO FINAL', index=False)
    df_confirmacao_saldo.to_excel(writer, sheet_name='CONFIRMAÇÃO SALDO', index=False)
    df_saldo_cliente.to_excel(writer, sheet_name='SALDO CLIENTE', index=False)
    df_hist_saldo_atualizado.to_excel(writer, sheet_name='HISTORICO SALDO', index=False)

print("🎉 Processamento concluído com sucesso!")