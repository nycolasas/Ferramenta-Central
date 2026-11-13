import os
import openpyxl
import pandas as pd

from portable_paths import BASE_WORKBOOK, SOURCE_WORKBOOK, ensure_portable_data_files

ensure_portable_data_files()

# ==============================================================================
# 1. CAMINHOS DOS ARQUIVOS
# ==============================================================================
caminho_origem = str(SOURCE_WORKBOOK)
caminho_destino = str(BASE_WORKBOOK)

# Mapeamento: 'Aba Origem': 'Aba Destino'
mapeamento_abas = {
    'SALDO': 'f_saldo',
    'AVISO_ENTREGA': 'f_registro',
    'PEDIDOS': 'f_4260',
    'ESTOQUE_LOTE': 'f_vt'
}

# (Opcional) Se quiser mapear colunas de data específicas para cada aba, adicione abaixo:
# Exemplo: 'SALDO': ['dt_emissao', 'dt_vencimento']
colunas_datas_por_aba = {
    'SALDO': [],
    'AVISO_ENTREGA': [],
    'PEDIDOS': [],
    'ESTOQUE_LOTE': []
}

# ==============================================================================
# 2. VERIFICAÇÃO DE EXISTÊNCIA DOS ARQUIVOS
# ==============================================================================
if not os.path.exists(caminho_origem):
    raise FileNotFoundError(f"❌ Arquivo de origem não encontrado em: {caminho_origem}")

if not os.path.exists(caminho_destino):
    raise FileNotFoundError(f"❌ Arquivo de destino não encontrado em: {caminho_destino}")

print("📂 Lendo os arquivos de Excel...")

# ==============================================================================
# 3. LEITURA E TRANSFERÊNCIA DE DADOS
# ==============================================================================
try:
    # Abre a planilha de destino usando openpyxl para manipular as abas
    wb_destino = openpyxl.load_workbook(caminho_destino)
    
    # Lê todas as abas da planilha de origem via Pandas
    dict_origem = pd.read_excel(caminho_origem, sheet_name=None)
    
    for aba_origem, aba_destino in mapeamento_abas.items():
        print(f"\n🔄 Processando: [{aba_origem}] ➡️ [{aba_destino}]")
        
        # Verifica se a aba existe na origem
        if aba_origem not in dict_origem:
            print(f"⚠️ Aba de origem '{aba_origem}' não foi encontrada em relatorios_consolidados.xlsx. Pulando...")
            continue
            
        # Verifica se a aba existe no destino
        if aba_destino not in wb_destino.sheetnames:
            print(f"⚠️ Aba de destino '{aba_destino}' não foi encontrada em planilha_bruta.xlsx. Criando aba...")
            ws_dest = wb_destino.create_sheet(title=aba_destino)
        else:
            ws_dest = wb_destino[aba_destino]
            
        # 1. Limpa todo o conteúdo antigo da aba de destino
        ws_dest.delete_rows(1, ws_dest.max_row + 1)
        
        # 2. Pega os novos dados da origem
        df_dados = dict_origem[aba_origem]
        
        if df_dados.empty:
            print(f"ℹ️ A aba '{aba_origem}' está vazia. Conteúdo antigo de '{aba_destino}' foi limpo.")
            continue

        # ----------------------------------------------------------------------
        # 🛠️ NOVO BLOCO: TRATAMENTO E CONVERSÃO DE COLUNAS DE DATA
        # ----------------------------------------------------------------------
        # Identifica colunas mapeadas manualmente ou tenta detectar via padrão de texto no nome da coluna
        cols_explicitadas = colunas_datas_por_aba.get(aba_origem, [])
        cols_detectadas = [
            col for col in df_dados.columns 
            if any(termo in str(col).lower() for termo in ['data', 'dt', 'date'])
        ]
        
        # Une as colunas encontradas sem duplicatas
        cols_datas = list(set(cols_explicitadas + cols_detectadas))
        
        if cols_datas:
            print(f"📅 Colunas de data identificadas para conversão: {cols_datas}")
            for col in cols_datas:
                if col in df_dados.columns:
                    # Converte para datetime tratando erros e garantindo o formato da data (AAAA-MM-DD)
                    df_dados[col] = pd.to_datetime(
                        df_dados[col], 
                        dayfirst=True,   # Ajuste para True se datas numéricas forem dia/mês/ano
                        errors='coerce'  # Registra NaT se a conversão falhar para algum item
                    ).dt.date            # Mantém no formato de data sem o horário (00:00:00)
        # ----------------------------------------------------------------------
            
        # 3. Escreve o cabeçalho
        cabecalho = list(df_dados.columns)
        ws_dest.append(cabecalho)
        
        # 4. Escreve as linhas de dados
        # Trata valores NaN/NaT/None para não quebrar no Excel
        df_dados_limpos = df_dados.fillna('')
        for linha in df_dados_limpos.values.tolist():
            ws_dest.append(linha)
            
        print(f"✅ {len(df_dados)} linhas copiadas com sucesso para '{aba_destino}'.")

    # ==============================================================================
    # 4. SALVAMENTO
    # ==============================================================================
    print(f"\n💾 Salvando alterações em {caminho_destino}...")
    wb_destino.save(caminho_destino)
    wb_destino.close()
    print("🎉 Atualização da planilha_bruta.xlsx concluída com sucesso!")

except PermissionError:
    print("\n❌ ERRO DE PERMISSÃO: O arquivo 'planilha_bruta.xlsx' está aberto no Excel! Feche o arquivo e tente novamente.")
except Exception as e:
    print(f"\n❌ Ocorreu um erro durante o processamento: {e}")