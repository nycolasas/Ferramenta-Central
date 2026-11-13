import os
import re
import sys
import unicodedata
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

from src.paths import BASE_WORKBOOK, FERIADOS_WORKBOOK, ensure_portable_data_files

# Permite os logs com emojis em terminais Windows configurados com cp1252.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ensure_portable_data_files()

# ==============================================================================
# FUNÇÕES AUXILIARES DE TRATAMENTO DE TEXTO E BUSCA FLEXÍVEL
# ==============================================================================
def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    txt = str(texto).strip()
    txt = unicodedata.normalize('NFKD', txt).encode('ASCII', 'ignore').decode('utf-8')
    txt = txt.upper()
    txt = re.sub(r'[^A-Z0-9]', '', txt)
    return txt

def encontrar_coluna_flexivel(df, palavras_chave, posicao_padrao=0):
    for col in df.columns:
        col_str = str(col).strip().upper()
        if any(p.upper() in col_str for p in palavras_chave):
            return col
    if len(df.columns) > posicao_padrao:
        return df.columns[posicao_padrao]
    return None

def subtrair_dias_uteis(data_inicio, dias_uteis, feriados_set):
    cur = data_inicio
    count = 0
    while count < dias_uteis:
        cur -= pd.Timedelta(days=1)
        if cur.weekday() < 5 and cur not in feriados_set:
            count += 1
    return cur

# ==============================================================================
# FERRAMENTA DE AUDITORIA E DEPURAÇÃO DETALHADA NO TERMINAL
# ==============================================================================
def auditar_pedido(termo_busca, caminho_arquivo_consolidado):
    """
    Busca parcialmente pelo número do pedido, exibe o representante, região,
    passa por todas as etapas de validação e informa o status de atraso detalhado.
    """
    print(f"\n🔍 [AUDITORIA DE PEDIDO] Buscando ocorrências para: '{termo_busca}'...\n" + "="*70)
    
    try:
        df_saldo = pd.read_excel(caminho_arquivo_consolidado, sheet_name='f_saldo')
    except Exception as e:
        print(f"❌ Erro ao carregar a aba 'f_saldo': {e}")
        return

    try:
        df_hist = pd.read_excel(caminho_arquivo_consolidado, sheet_name='f_registro')
    except Exception:
        try:
            df_hist = pd.read_excel(caminho_arquivo_consolidado, sheet_name='HISTORICO')
        except Exception:
            df_hist = pd.DataFrame()

    try:
        df_canais = pd.read_excel(caminho_arquivo_consolidado, sheet_name='canais')
        df_canais.columns = df_canais.columns.astype(str).str.strip()
    except Exception:
        df_canais = pd.DataFrame()

    try:
        df_feriados = pd.read_excel(str(FERIADOS_WORKBOOK), sheet_name='Feriados')
        col_feriado = [c for c in df_feriados.columns if 'FERIADO' in str(c).upper()][0]
        lista_feriados = pd.to_datetime(df_feriados[col_feriado], errors='coerce').dt.normalize().dropna().tolist()
        set_feriados = set(lista_feriados)
    except Exception:
        set_feriados = set()

    df_saldo.columns = df_saldo.columns.astype(str).str.strip()
    df_hist.columns = df_hist.columns.astype(str).str.strip()

    # Mapeamento de canais para região
    dict_rep_para_regiao = {}
    if not df_canais.empty:
        col_rep_c = encontrar_coluna_flexivel(df_canais, ['REPRES', 'REP', 'VENDEDOR', 'CANAL', 'COD'], 0)
        col_regiao_c = encontrar_coluna_flexivel(df_canais, ['REGI', 'UF', 'ESTADO'], 1)
        for _, row in df_canais.iterrows():
            rep_norm = normalizar_texto(row[col_rep_c])
            reg_val = str(row[col_regiao_c]).strip() if col_regiao_c and pd.notna(row[col_regiao_c]) else "N/A"
            if rep_norm:
                dict_rep_para_regiao[rep_norm] = reg_val

    def buscar_dados_rep(representante):
        if pd.isna(representante):
            return "N/A"
        chave_norm = normalizar_texto(representante)
        if chave_norm in dict_rep_para_regiao:
            return dict_rep_para_regiao[chave_norm]
        for k in dict_rep_para_regiao:
            if k in chave_norm or chave_norm in k:
                return dict_rep_para_regiao[k]
        return "N/A"

    def obter_cliente_final(val_cliente, val_destino):
        c_orig = str(val_cliente).strip() if pd.notna(val_cliente) else ""
        c_dest = str(val_destino).strip() if pd.notna(val_destino) else ""
        if c_orig.upper() != c_dest.upper():
            if c_orig.upper().startswith("LIDER") and c_dest != "":
                return c_dest
            elif c_orig.upper().startswith("LIDER") and c_dest == "":
                return c_orig
        return c_orig

    # Mapeamento de colunas em f_saldo
    col_cliente = encontrar_coluna_flexivel(df_saldo, ['CLIENTE'], 0)
    col_cliente_destino = encontrar_coluna_flexivel(df_saldo, ['CLIENTE DESTINO', 'DESTINO'], 2)
    col_pedido = encontrar_coluna_flexivel(df_saldo, ['PEDIDO', 'NUMERO'], 1)
    col_situacao = encontrar_coluna_flexivel(df_saldo, ['SITUACAO', 'SITUAÇÃO', 'ST', 'SIT'], 3)
    col_dt_entrega = encontrar_coluna_flexivel(df_saldo, ['PRAZO ENTREGA', 'DATA ENTREGA', 'DT ENTREGA', 'PREVISAO'], 4)
    col_cod_item = encontrar_coluna_flexivel(df_saldo, ['ITEM', 'COD', 'PRODUTO'], 5)
    col_natureza = encontrar_coluna_flexivel(df_saldo, ['NATUREZA', 'OPERA', 'NAT', 'LINHA'], 8)
    col_qtd = encontrar_coluna_flexivel(df_saldo, ['QUANTIDADE', 'QTD', 'SALDO'], 9)
    col_est = encontrar_coluna_flexivel(df_saldo, ['ESTOQUE', 'EST'], 10)
    col_rep_saldo = encontrar_coluna_flexivel(df_saldo, ['REPRESENTANTE', 'CANAL', 'VENDEDOR', 'REP'], 18)
    col_lider = encontrar_coluna_flexivel(df_saldo, ['LIDER', 'PERFIL', 'GRUPO'], 21)

    # Mapeamento de histórico de avisos
    dict_avisos = {}
    if not df_hist.empty:
        col_hist_ped = encontrar_coluna_flexivel(df_hist, ['PEDIDO', 'NUMERO'], 1)
        col_hist_contato = encontrar_coluna_flexivel(df_hist, ['CONTATO', 'DATA'], 4)
        col_hist_prev = encontrar_coluna_flexivel(df_hist, ['PREVISAO ENTREGA', 'PREVIS', 'NOVA DATA'], 5)
        col_hist_motivo = encontrar_coluna_flexivel(df_hist, ['MOTIVO', 'ACAO'], 8)
        col_hist_tipo = encontrar_coluna_flexivel(df_hist, ['TIPO', 'FORMA'], 9)

        if col_hist_motivo and col_hist_tipo:
            cond_motivo = df_hist[col_hist_motivo].astype(str).str.strip().str.upper() == "ATRASO ENTREGA AVISADO"
            cond_tipo = df_hist[col_hist_tipo].astype(str).str.strip().str.upper() == "CONTATO CLIENTE"
            df_hist_filtrado = df_hist[cond_motivo & cond_tipo].copy()
            df_hist_filtrado['Dt_Previsao_Hist'] = pd.to_datetime(df_hist_filtrado[col_hist_prev], errors='coerce')
            for _, h_row in df_hist_filtrado.iterrows():
                p_id = str(h_row[col_hist_ped]).strip()
                dt_h = h_row['Dt_Previsao_Hist']
                if pd.notna(dt_h) and dt_h.year != 2040:
                    dict_avisos[p_id] = dt_h.normalize()

    # Filtro flexível (contém o termo digitado, ignorando maiúsculas/minúsculas)
    mask_pedido = df_saldo[col_pedido].astype(str).str.contains(str(termo_busca), case=False, na=False)
    df_pedido = df_saldo[mask_pedido]

    if df_pedido.empty:
        print(f"❌ MOTIVO DO DESCARTE: Nenhum pedido contendo '{termo_busca}' foi encontrado na aba 'f_saldo'.")
        return

    pedidos_encontrados = df_pedido[col_pedido].astype(str).unique()
    print(f"📦 Pedido(s) encontrado(s) na base: {list(pedidos_encontrados)}")
    print(f"📊 Total de linhas/itens correspondentes: {len(df_pedido)}\n" + "-"*70)

    hoje = pd.Timestamp('today').normalize()

    # Analisar linha por linha dos pedidos encontrados
    for idx, row in df_pedido.iterrows():
        ped_raw = str(row[col_pedido]).strip()
        rep_val = row[col_rep_saldo] if col_rep_saldo and pd.notna(row[col_rep_saldo]) else "Não informado"
        regiao_val = buscar_dados_rep(row[col_rep_saldo])
        regiao_norm = normalizar_texto(regiao_val)

        print(f"📌 [ITEM / LINHA {idx}] - Pedido: {ped_raw}")
        print(f"  👤 REPRESENTANTE: {rep_val}")
        print(f"  🗺️  REGIÃO: {regiao_val}")
        print("  |-- 📋 ETAPA 1: Verificação da Empresa (Prefixo 11)")

        # 1. Filtro Empresa 11
        if not (ped_raw.startswith("(11)-") or ped_raw.startswith("11-")):
            print(f"  |   ❌ REJEITADO: O pedido não pertence à Empresa 11 (Valor: '{ped_raw[:10]}').\n" + "-"*50)
            continue
        print("  |   ✅ APROVADO: Pertence à Empresa 11.")

        # 2. Filtro Perfil Lider
        print("  |-- 📋 ETAPA 2: Verificação de Perfil/Liderança e Cliente Final")
        val_lider_raw = str(row[col_lider]).strip().upper() if col_lider and pd.notna(row[col_lider]) else ""
        if val_lider_raw == "LIDER":
            print(f"  |   ❌ REJEITADO: O campo de perfil está marcado como 'LIDER'.\n" + "-"*50)
            continue
            
        cliente_final = obter_cliente_final(row[col_cliente], row.get(col_cliente_destino, ""))
        if cliente_final.strip().upper().startswith("LIDER"):
            print(f"  |   ❌ REJEITADO: O cliente final ('{cliente_final}') começa com 'LIDER'.\n" + "-"*50)
            continue
        print(f"  |   ✅ APROVADO: Perfil e Cliente Final válidos (Cliente: {cliente_final}).")

        # 3. Filtros de Natureza e Itens Especiais
        print("  |-- 📋 ETAPA 3: Análise do Código do Item e Natureza da Operação")
        cod_item = str(row[col_cod_item]).replace('\xa0', '').strip().upper() if col_cod_item and pd.notna(row[col_cod_item]) else ""
        natureza = str(row[col_natureza]).replace('\xa0', '').strip().upper() if col_natureza and pd.notna(row[col_natureza]) else ""
        
        if cod_item.startswith("3TAXA"):
            print(f"  |   ⏩ IGNORADO: O item é uma taxa ('{cod_item}').\n" + "-"*50)
            continue
        if natureza in ["24-SERVIÇOS E INSUMOS", "25-PROTECAO DE TECIDOS", "99-GENERICO"]:
            print(f"  |   ⏩ IGNORADO: Natureza de serviço/genérica ('{natureza}').\n" + "-"*50)
            continue
        print(f"  |   ✅ APROVADO: Item produtivo válido (Código: {cod_item} | Natureza: {natureza}).")

        # 4. Filtro Planejado MG
        print("  |-- 📋 ETAPA 4: Regra de Pedidos Planejados em Minas Gerais (MG)")
        eh_mg = "MG" in regiao_norm or "MINAS" in regiao_norm or "MINASGERAIS" in regiao_norm
        eh_planejado_item = "PLAN" in cod_item or "PLAN" in natureza or natureza.startswith("4-")
        
        if eh_mg and eh_planejado_item:
            print(f"  |   ❌ REJEITADO: Item planejado para o estado de Minas Gerais (Região: {regiao_val}).\n" + "-"*50)
            continue
        print(f"  |   ✅ APROVADO: Região e Planejamento regulares.")

        # 5. Análise de Estoque vs Quantidade Pendente
        print("  |-- 📋 ETAPA 5: Comparação de Estoque vs Quantidade Necessária")
        try:
            val_qtd_str = str(row[col_qtd]).replace('.', '').replace(',', '.')
            qtd_val = float(re.sub(r'[^0-9.]', '', val_qtd_str))
        except Exception:
            qtd_val = 0.0

        try:
            val_est_str = str(row[col_est]).replace('.', '').replace(',', '.')
            est_val = float(re.sub(r'[^0-9.]', '', val_est_str))
        except Exception:
            est_val = 0.0

        print(f"  |   ℹ️ Quantidade Pendente: {qtd_val} | Estoque Disponível: {est_val}")
        if est_val >= qtd_val:
            print(f"  |   ❌ REJEITADO: Estoque atende ou supera a quantidade necessária. Sem pendência crítica.\n" + "-"*50)
            continue
        print(f"  |   ✅ APROVADO: Estoque insuficiente (há saldo pendente crítico).")

        # 6. Validação da Data de Entrega
        print("  |-- 📋 ETAPA 6: Validação do Prazo/Data de Entrega e Feriados")
        dt_entrega_raw = row[col_dt_entrega]
        dt_entrega = pd.to_datetime(dt_entrega_raw, dayfirst=True, errors='coerce')
        
        if pd.isna(dt_entrega) or dt_entrega.year == 2040:
            print(f"  |   ❌ REJEITADO: Data de entrega inválida ou igual a 2040 ({dt_entrega_raw}).\n" + "-"*50)
            continue
        dt_entrega = dt_entrega.normalize()
        print(f"  |   ✅ APROVADO: Data de entrega válida: {dt_entrega.strftime('%d/%m/%Y')}")

        # Status de Suspensão
        situacao_val = str(row[col_situacao]).strip().upper() if col_situacao and pd.notna(row[col_situacao]) else ""
        eh_suspenso = (situacao_val == "S")
        if eh_suspenso:
            print(f"  |   ⚠️ ALERTA: Este item específico está marcado como **SUSPENSO** (Situação = 'S').")

        # Prazos por Região da Matriz Oficial
        if eh_mg:
            prazo_dias_base = 3
        elif "BAHIA" in regiao_norm or "BA" in regiao_norm:
            prazo_dias_base = 9
        else:
            prazo_dias_base = 7

        # 7. AVALIAÇÃO OFICIAL DE ATRASO / MATRIZ DE REAVISO
        print("  |-- 🎯 ETAPA 7: Aplicação da Matriz de Atraso e Limite de Dias Úteis")
        print(f"  |   Data de Referência (Hoje):   {hoje.strftime('%d/%m/%Y')}")
        print(f"  |   Prazo base da Região:        {prazo_dias_base} dia(s) útil(eis)")

        status_acao = None
        data_limite = subtrair_dias_uteis(dt_entrega, prazo_dias_base, set_feriados)
        print(f"  |   Data Limite de Corte (Útil): {data_limite.strftime('%d/%m/%Y')}")

        if ped_raw not in dict_avisos:
            if (hoje > dt_entrega) or (hoje > data_limite):
                status_acao = "AVISAR"
            else:
                motivo_dia = f"Ainda está dentro do prazo limite (limite de corte útil era {data_limite.strftime('%d/%m/%Y')})."
        else:
            dt_previsao_hist = dict_avisos[ped_raw]
            print(f"  |   Histórico de Aviso Anterior: {dt_previsao_hist.strftime('%d/%m/%Y')}")
            if dt_entrega > dt_previsao_hist:
                nova_data_limite = subtrair_dias_uteis(dt_entrega, prazo_dias_base, set_feriados)
                if (hoje > dt_entrega) or (hoje > nova_data_limite):
                    status_acao = "REAVISAR ATRASO"
                else:
                    motivo_dia = "Data foi estendida no saldo e ainda está dentro do novo prazo estendido."
            else:
                data_limite_reavisar = subtrair_dias_uteis(dt_previsao_hist, prazo_dias_base, set_feriados)
                if (hoje > dt_previsao_hist) or (hoje > data_limite_reavisar):
                    status_acao = "REAVISAR ATRASO"
                else:
                    motivo_dia = "Já possui aviso anterior e está dentro do prazo de reaviso baseado na previsão histórica."

        if status_acao:
            print(f"  |   🔴 **STATUS FINAL: {status_acao} (ATRASADO / CRÍTICO)**")
            print(f"  |   📝 **Critério Atendido:** O pedido atingiu os critérios de atraso da matriz oficial e precisa ser acionado/reavisado.\n" + "="*70)
        else:
            print(f"  |   🟢 **STATUS FINAL: EM DIA (NÃO ATRASADO)**")
            print(f"  |   📝 **Motivo do porquê não está atrasado:** {motivo_dia}\n" + "="*70)

if __name__ == "__main__":
    ARQUIVO_CONSOLIDADO = str(BASE_WORKBOOK)
    
    print("="*70)
    print("🛠️  AUDITOR E DEPURADOR DE PEDIDOS - LIDER INTERIORS")
    print("="*70)
    termo_alvo = input("Digite o número ou parte do pedido (ex: 201631): ").strip()
    
    if termo_alvo:
        auditar_pedido(termo_alvo, ARQUIVO_CONSOLIDADO)
    else:
        print("⚠️ Nenhum termo foi inserido.")
