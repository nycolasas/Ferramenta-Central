import os
import re
import sys
import time
import unicodedata
import pandas as pd
import numpy as np

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.remote.client_config import ClientConfig
from selenium.webdriver.support.ui import WebDriverWait

from src.paths import ATRASADOS_WORKBOOK, BASE_WORKBOOK, FERIADOS_WORKBOOK, ensure_portable_data_files

# Permite os logs com emojis em terminais Windows configurados com cp1252.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ensure_portable_data_files()

# ==============================================================================
# CONFIGURAÇÃO E INICIALIZAÇÃO DO SELENIUM
# ==============================================================================
def iniciar_driver():
    options = Options()
    options.page_load_strategy = 'eager'
    
    client_config = ClientConfig(timeout=300)
    driver = webdriver.Chrome(options=options, client_config=client_config)
    driver.maximize_window()
    return driver

# ==============================================================================
# 1. AUTOMAÇÃO E EXTRAÇÃO WEB (SELENIUM)
# ==============================================================================
def extrair_dados_web():
    print("🌐 Iniciando navegação Web via Selenium...")
    driver = iniciar_driver()
    
    try:
        wait = WebDriverWait(driver, 30)
        print("⏳ Aguardando conclusão da consulta no servidor...")
        time.sleep(5)
        
    except Exception as e:
        print(f"❌ Erro na automação Web: {e}")
        raise e
    finally:
        driver.quit()

# ==============================================================================
# 2. FUNÇÕES AUXILIARES DE TRATAMENTO DE TEXTO E BUSCA FLEXÍVEL
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

def subtrair_dias_uteis(data_inicio, dias_uteis, feriados_set):
    cur = data_inicio
    count = 0
    while count < dias_uteis:
        cur -= pd.Timedelta(days=1)
        if cur.weekday() < 5 and cur not in feriados_set:
            count += 1
    return cur

# ==============================================================================
# 3. LÓGICA PRINCIPAL DE PROCESSAMENTO DOS DADOS (PANDAS)
# ==============================================================================
def executar_saldo(caminho_arquivo_consolidado):
    
    caminho_bruta = caminho_arquivo_consolidado
    caminho_feriados = str(FERIADOS_WORKBOOK)
    caminho_saida = str(ATRASADOS_WORKBOOK)

    print("📂 Carregando bases de dados...")

    try:
        df_saldo = pd.read_excel(caminho_bruta, sheet_name='f_saldo')
    except Exception as e:
        raise FileNotFoundError(f"❌ Erro ao carregar a aba 'f_saldo': {e}")

    try:
        df_hist = pd.read_excel(caminho_bruta, sheet_name='f_registro')
    except Exception:
        try:
            df_hist = pd.read_excel(caminho_bruta, sheet_name='HISTORICO')
        except Exception as e:
            raise FileNotFoundError(f"❌ Erro ao carregar a aba de histórico ('f_registro' ou 'HISTORICO'): {e}")

    try:
        df_canais = pd.read_excel(caminho_bruta, sheet_name='canais')
        df_canais.columns = df_canais.columns.astype(str).str.strip()
    except Exception:
        df_canais = pd.DataFrame()
        print("⚠️ Aba 'canais' não encontrada.")

    try:
        df_atendente = pd.read_excel(caminho_bruta, sheet_name='atendente')
        df_atendente.columns = df_atendente.columns.astype(str).str.strip()
    except Exception:
        df_atendente = pd.DataFrame()
        print("⚠️ Aba 'atendente' não encontrada.")

    try:
        df_feriados = pd.read_excel(caminho_feriados, sheet_name='Feriados')
        col_feriado = [c for c in df_feriados.columns if 'FERIADO' in str(c).upper()][0]
        lista_feriados = pd.to_datetime(df_feriados[col_feriado], errors='coerce').dt.normalize().dropna().tolist()
        set_feriados = set(lista_feriados)
        print(f"🗓️ Carregados {len(set_feriados)} feriados registrados.")
    except Exception as e:
        print(f"⚠️ Aviso ao carregar feriados: {e}. Prosseguindo sem cálculo de feriados específicos.")
        set_feriados = set()

    df_saldo.columns = df_saldo.columns.astype(str).str.strip()
    df_hist.columns = df_hist.columns.astype(str).str.strip()

    dict_rep_para_regiao = {}
    dict_rep_para_loja = {}

    dict_chave_solto = {}
    dict_chave_planejado = {}

    if not df_canais.empty:
        col_rep_c    = encontrar_coluna_flexivel(df_canais, ['REPRES', 'REP', 'VENDEDOR', 'CANAL', 'COD'], 0)
        col_regiao_c = encontrar_coluna_flexivel(df_canais, ['REGI', 'UF', 'ESTADO'], 1)
        col_loja_c   = encontrar_coluna_flexivel(df_canais, ['LOJA', 'CIDADE', 'UNIDADE'], 2)
        
        print(f"ℹ️ Aba 'canais' ligada: Rep='{col_rep_c}' ➔ Região='{col_regiao_c}', Loja='{col_loja_c}'")

        for _, row in df_canais.iterrows():
            rep_norm = normalizar_texto(row[col_rep_c])
            reg_val  = str(row[col_regiao_c]).strip() if col_regiao_c and pd.notna(row[col_regiao_c]) else "N/A"
            loja_val = str(row[col_loja_c]).strip() if col_loja_c and pd.notna(row[col_loja_c]) else "N/A"
            
            if rep_norm:
                dict_rep_para_regiao[rep_norm] = reg_val
                dict_rep_para_loja[rep_norm]   = loja_val

    if not df_atendente.empty:
        col_chave_a  = encontrar_coluna_flexivel(df_atendente, ['LOJA', 'REGI', 'ESTADO', 'UF', 'CIDADE'], 0)
        col_solto_a  = encontrar_coluna_flexivel(df_atendente, ['SOLTO', 'ESTOFADO', 'SOLT'], 1)
        col_plan_a   = encontrar_coluna_flexivel(df_atendente, ['PLAN'], 2)

        print(f"ℹ️ Aba 'atendente' ligada: Chave='{col_chave_a}', Solto='{col_solto_a}', Plan='{col_plan_a}'")

        for _, row in df_atendente.iterrows():
            if col_chave_a and pd.notna(row[col_chave_a]):
                chave_norm = normalizar_texto(row[col_chave_a])
                if chave_norm:
                    dict_chave_solto[chave_norm]     = str(row[col_solto_a]).strip() if col_solto_a and pd.notna(row[col_solto_a]) else "N/A"
                    dict_chave_planejado[chave_norm] = str(row[col_plan_a]).strip() if col_plan_a and pd.notna(row[col_plan_a]) else "N/A"

    def buscar_dados_rep(representante):
        if pd.isna(representante):
            return "N/A", "N/A"
        chave_norm = normalizar_texto(representante)
        
        if chave_norm in dict_rep_para_regiao:
            return dict_rep_para_regiao[chave_norm], dict_rep_para_loja.get(chave_norm, "N/A")
        
        for k in dict_rep_para_regiao:
            if k in chave_norm or chave_norm in k:
                return dict_rep_para_regiao[k], dict_rep_para_loja.get(k, "N/A")
                
        return "N/A", "N/A"

    def buscar_responsavel(regiao, loja, eh_planejado=False):
        chaves_para_testar = [normalizar_texto(regiao), normalizar_texto(loja)]
        
        for chave in chaves_para_testar:
            if not chave or chave == "NA":
                continue
                
            if eh_planejado:
                res = dict_chave_planejado.get(chave, "N/A")
                if res == "N/A":
                    res = dict_chave_solto.get(chave, "N/A")
                if res != "N/A":
                    return res
            else:
                res = dict_chave_solto.get(chave, "N/A")
                if res != "N/A":
                    return res
                    
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

    print("🔍 Mapeando histórico de avisos e datas previstas...")

    col_hist_ped     = encontrar_coluna_flexivel(df_hist, ['PEDIDO', 'NUMERO'], 1)
    col_hist_contato = encontrar_coluna_flexivel(df_hist, ['CONTATO', 'DATA'], 4)
    col_hist_prev    = encontrar_coluna_flexivel(df_hist, ['PREVISAO ENTREGA', 'PREVIS', 'NOVA DATA'], 5)
    col_hist_motivo  = encontrar_coluna_flexivel(df_hist, ['MOTIVO', 'ACAO'], 8)
    col_hist_tipo    = encontrar_coluna_flexivel(df_hist, ['TIPO', 'FORMA'], 9)

    cond_motivo = df_hist[col_hist_motivo].astype(str).str.strip().str.upper() == "ATRASO ENTREGA AVISADO"
    cond_tipo   = df_hist[col_hist_tipo].astype(str).str.strip().str.upper() == "CONTATO CLIENTE"

    df_hist_filtrado = df_hist[cond_motivo & cond_tipo].copy()
    df_hist_filtrado['Dt_Contato'] = pd.to_datetime(df_hist_filtrado[col_hist_contato], errors='coerce')
    df_hist_filtrado['Dt_Previsao_Hist'] = pd.to_datetime(df_hist_filtrado[col_hist_prev], errors='coerce')

    df_hist_filtrado = df_hist_filtrado.sort_values(by=['Dt_Contato'], ascending=True)

    dict_avisos = {}
    for _, row in df_hist_filtrado.iterrows():
        ped = str(row[col_hist_ped]).strip()
        dt_hist = row['Dt_Previsao_Hist']
        # Desconsidera datas do ano 2040 gravadas no histórico
        if pd.notna(dt_hist) and dt_hist.year != 2040:
            dict_avisos[ped] = dt_hist.normalize()

    print("⚡ Mapeando colunas e analisando composição dos pedidos...")

    col_cliente         = encontrar_coluna_flexivel(df_saldo, ['CLIENTE'], 0)
    col_cliente_destino = encontrar_coluna_flexivel(df_saldo, ['CLIENTE DESTINO', 'DESTINO'], 2)
    col_pedido          = encontrar_coluna_flexivel(df_saldo, ['PEDIDO', 'NUMERO'], 1)
    col_situacao        = encontrar_coluna_flexivel(df_saldo, ['SITUACAO', 'SITUAÇÃO', 'ST', 'SIT'], 3)
    col_dt_entrega      = encontrar_coluna_flexivel(df_saldo, ['PRAZO ENTREGA', 'DATA ENTREGA', 'DT ENTREGA', 'PREVISAO'], 4)
    col_cod_item        = encontrar_coluna_flexivel(df_saldo, ['ITEM', 'COD', 'PRODUTO'], 5)
    col_natureza        = encontrar_coluna_flexivel(df_saldo, ['NATUREZA', 'OPERA', 'NAT', 'LINHA'], 8)
    col_qtd             = encontrar_coluna_flexivel(df_saldo, ['QUANTIDADE', 'QTD', 'SALDO'], 9)
    col_est             = encontrar_coluna_flexivel(df_saldo, ['ESTOQUE', 'EST'], 10)
    col_rep_saldo       = encontrar_coluna_flexivel(df_saldo, ['REPRESENTANTE', 'CANAL', 'VENDEDOR', 'REP'], 18)
    col_lider           = encontrar_coluna_flexivel(df_saldo, ['LIDER', 'PERFIL', 'GRUPO'], 21)

    # Identificação de móvel planejado no pedido
    dict_pedido_contem_planejado = {}
    for ped_id, group in df_saldo.groupby(col_pedido):
        ped_str = str(ped_id).strip()
        tem_planejado = False
        
        for _, row_item in group.iterrows():
            cod_i = str(row_item[col_cod_item]).replace('\xa0', '').strip().upper()
            nat_i = str(row_item[col_natureza]).replace('\xa0', '').strip().upper()
            
            if cod_i.startswith("3TAXA") or nat_i in ["24-SERVIÇOS E INSUMOS", "25-PROTECAO DE TECIDOS", "99-GENERICO"]:
                continue
                
            if "PLAN" in cod_i or "PLAN" in nat_i or nat_i.startswith("4-"):
                tem_planejado = True
                break
                
        dict_pedido_contem_planejado[ped_str] = tem_planejado

    hoje = pd.Timestamp('today').normalize()
    data_hora_agora = pd.Timestamp.now().strftime('%d/%m/%Y %H:%M')

    resultados = []
    suspensos = []
    pedidos_processados = set()
    pedidos_planejados_mg = set()

    for idx, row in df_saldo.iterrows():
        ped_raw = str(row[col_pedido]).strip() if pd.notna(row[col_pedido]) else ""
        if not ped_raw or ped_raw in pedidos_processados:
            continue
        
        # Filtro Empresa 11
        if not (ped_raw.startswith("(11)-") or ped_raw.startswith("11-")):
            continue
            
        # Filtro Perfis "LIDER"
        if str(row[col_lider]).strip().upper() == "LIDER":
            continue

        cliente_final = obter_cliente_final(row[col_cliente], row.get(col_cliente_destino, ""))

        if cliente_final.strip().upper().startswith("LIDER"):
            continue
            
        rep_val = row[col_rep_saldo]
        regiao_val, loja_val = buscar_dados_rep(rep_val)
        regiao_norm = normalizar_texto(regiao_val)
        
        eh_mg = "MG" in regiao_norm or "MINAS" in regiao_norm or "MINASGERAIS" in regiao_norm

        eh_planejado_pedido = dict_pedido_contem_planejado.get(ped_raw, False)
        if eh_mg and eh_planejado_pedido:
            pedidos_planejados_mg.add(ped_raw)
            pedidos_processados.add(ped_raw)
            continue

        cod_item = str(row[col_cod_item]).replace('\xa0', '').strip().upper()
        if cod_item.startswith("3TAXA"):
            continue
            
        natureza = str(row[col_natureza]).replace('\xa0', '').strip().upper()
        if natureza in ["24-SERVIÇOS E INSUMOS", "25-PROTECAO DE TECIDOS", "99-GENERICO"]:
            continue

        eh_planejado_item = "PLAN" in cod_item or "PLAN" in natureza or natureza.startswith("4-")
        responsavel = buscar_responsavel(regiao_val, loja_val, eh_planejado=eh_planejado_item)

        situacao_val = str(row[col_situacao]).strip().upper() if pd.notna(row[col_situacao]) else ""
        eh_suspenso = (situacao_val == "S")

        if eh_suspenso:
            suspensos.append({
                "Pedido": ped_raw,
                "Cliente": cliente_final,
                "Responsável": responsavel,
                "Região": regiao_val,
                "Loja": loja_val,
                "Situação": "SUSPENSO",
                "Data/Hora Identificação": data_hora_agora
            })

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
            
        if est_val >= qtd_val:
            continue
            
        dt_entrega = pd.to_datetime(row[col_dt_entrega], dayfirst=True, errors='coerce')
        
        # Ignora pedidos sem data válida ou que possuem data controle 2040 no Saldo
        if pd.isna(dt_entrega) or dt_entrega.year == 2040:
            continue
        dt_entrega = dt_entrega.normalize()

        # Prazos por Região
        if eh_mg:
            prazo_dias_base = 3
        elif "BAHIA" in regiao_norm or "BA" in regiao_norm:
            prazo_dias_base = 9
        else:
            prazo_dias_base = 7
            
        # ----------------------------------------------------------------------
        # VALIDAÇÃO ROBUSTA DE ATRASO / REAVISO
        # ----------------------------------------------------------------------
        status_acao = None
        dt_exibicao = dt_entrega

        if ped_raw not in dict_avisos:
            # CASO 1: Nunca foi avisado
            data_limite = subtrair_dias_uteis(dt_entrega, prazo_dias_base, set_feriados)
            if (hoje > dt_entrega) or (hoje > data_limite):
                status_acao = "AVISAR"
                dt_exibicao = dt_entrega

        else:
            # CASO 2: Já foi avisado anteriormente com data real (< 2040)
            dt_previsao_hist = dict_avisos[ped_raw]

            # Se a data do saldo foi estendida (maior que a do histórico)
            if dt_entrega > dt_previsao_hist:
                nova_data_limite = subtrair_dias_uteis(dt_entrega, prazo_dias_base, set_feriados)
                
                # Só gera reaviso se a nova data também já estiver em atraso
                if (hoje > dt_entrega) or (hoje > nova_data_limite):
                    status_acao = "REAVISAR ATRASO"
                    dt_exibicao = dt_entrega
                else:
                    status_acao = None # Inelegível por estar dentro do novo prazo estendido

            else:
                # Mantém validação com base na data do aviso gravado
                data_limite_reavisar = subtrair_dias_uteis(dt_previsao_hist, prazo_dias_base, set_feriados)
                if (hoje > dt_previsao_hist) or (hoje > data_limite_reavisar):
                    status_acao = "REAVISAR ATRASO"
                    dt_exibicao = dt_previsao_hist

        if status_acao:
            resultados.append({
                "Pedido": ped_raw,
                "Cliente": cliente_final,
                "Responsável": responsavel,
                "Região": regiao_val,
                "Loja": loja_val,
                "Suspenso": "SIM" if eh_suspenso else "NÃO",
                "Data Entrega": dt_exibicao.strftime('%d/%m/%Y'),
                "Status": status_acao
            })
            pedidos_processados.add(ped_raw)

    # GERAÇÃO DAS ABAS DA PLANILHA DE SAÍDA
    df_analise = pd.DataFrame(resultados)
    if not df_analise.empty and pedidos_planejados_mg:
        df_analise = df_analise[
            ~df_analise["Pedido"].astype(str).str.strip().isin(pedidos_planejados_mg)
        ].copy()
    df_analise = desambiguar_nomes_duplicados(df_analise)
    df_suspensos_hoje = pd.DataFrame(suspensos)

    df_historico_hoje = df_analise.copy()
    if not df_historico_hoje.empty:
        df_historico_hoje['Data Processamento'] = data_hora_agora

    if os.path.exists(caminho_saida):
        try:
            df_hist_ant = pd.read_excel(caminho_saida, sheet_name='HISTORICO_ATRASOS')
            df_historico_acumulado = pd.concat([df_hist_ant, df_historico_hoje], ignore_index=True)
        except Exception:
            df_historico_acumulado = df_historico_hoje

        try:
            df_susp_ant = pd.read_excel(caminho_saida, sheet_name='PEDIDO_SUSPENSO')
            df_suspensos_acumulado = pd.concat([df_susp_ant, df_suspensos_hoje], ignore_index=True).drop_duplicates(subset=['Pedido'], keep='last')
        except Exception:
            df_suspensos_acumulado = df_suspensos_hoje
    else:
        df_historico_acumulado = df_historico_hoje
        df_suspensos_acumulado = df_suspensos_hoje

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_analise.to_excel(writer, sheet_name='ANALISE', index=False)
        df_historico_acumulado.to_excel(writer, sheet_name='HISTORICO_ATRASOS', index=False)
        df_suspensos_acumulado.to_excel(writer, sheet_name='PEDIDO_SUSPENSO', index=False)

    print("\n" + "="*60)
    print("✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
    print(f"📊 Pedidos a AVISAR/REAVISAR nesta rodada: {len(df_analise)}")
    print(f"🛑 Pedidos SUSPENSOS identificados: {len(df_suspensos_hoje)}")
    print(f"📁 Arquivo gerado em: {caminho_saida}")
    print("="*60)

if __name__ == "__main__":
    ARQUIVO_CONSOLIDADO = str(BASE_WORKBOOK)
    executar_saldo(ARQUIVO_CONSOLIDADO)
