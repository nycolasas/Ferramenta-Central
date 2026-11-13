import os
import sys
import json
import pandas as pd
import requests

from src.paths import ATRASADOS_WORKBOOK, ensure_portable_data_files

ensure_portable_data_files()

# ==============================================================================
# CONFIGURAÇÃO
# ==============================================================================
# Insira a URL do seu Web App do Apps Script aqui:
URL_WEBHOOK_GOOGLE = "https://script.google.com/macros/s/AKfycbydJ--n1koGPtmUDCcyNTlGRKUFWHo_SqRve5gSXrt7lnzj7YMyIgy6SCQadt3QDr4S/exec"
caminho_resultado = str(ATRASADOS_WORKBOOK)

# ==============================================================================
# PROCESSAMENTO E ENVIO
# ==============================================================================
def executar():
    print("📂 Lendo dados da aba ANALISE...")
    
    if not os.path.exists(caminho_resultado):
        print(f"❌ Arquivo não encontrado: {caminho_resultado}")
        sys.exit(1)

    try:
        # Lê a aba ANALISE do Excel gerado
        df_analise = pd.read_excel(caminho_resultado, sheet_name='ANALISE')
    except Exception as e:
        print(f"❌ Erro ao ler a aba ANALISE: {e}")
        sys.exit(1)

    if not df_analise.empty:
        # Substitui nulos por string vazia
        df_analise = df_analise.fillna('')

        # Função auxiliar para extrair coluna pela letra do Excel (A=0, B=1, etc.)
        def col_excel(df, letra):
            idx = ord(letra.upper()) - ord('A')
            if df.shape[1] > idx:
                return df.iloc[:, idx].astype(str)
            return ''

        # ----------------------------------------------------------------------
        # 🎯 MAPEAMENTO DE COLUNAS (Origem Excel ANALISE -> Destino Google Sheets)
        # ----------------------------------------------------------------------
        df_mapeado = pd.DataFrame()

        df_mapeado['Col_A'] = col_excel(df_analise, 'A')  # Excel A -> Google A (Pedido)
        df_mapeado['Col_B'] = col_excel(df_analise, 'B')  # Excel B -> Google B (Cliente)
        df_mapeado['Col_C'] = ''                          # Vazio
        df_mapeado['Col_D'] = col_excel(df_analise, 'G')  # Excel G -> Google D (Data Entrega)
        df_mapeado['Col_E'] = col_excel(df_analise, 'H')  # Excel H -> Google E (Status)
        df_mapeado['Col_F'] = ''                          # Vazio
        df_mapeado['Col_G'] = ''                          # Vazio
        df_mapeado['Col_H'] = ''                          # Vazio
        df_mapeado['Col_I'] = col_excel(df_analise, 'C')  # Excel C -> Google I (Responsável)

        # Converte em matriz (lista de listas)
        matriz_dados = df_mapeado.values.tolist()

        print("🚀 Enviando dados para salvar histórico e atualizar aba Atraso...")

        try:
            resposta = requests.post(
                URL_WEBHOOK_GOOGLE, 
                data=json.dumps(matriz_dados),
                headers={'Content-Type': 'application/json'},
                timeout=30
            )

            if resposta.status_code == 200 and "html" not in resposta.text.lower():
                print("✅ Sincronização realizada com sucesso!")
                print(f"📄 Resposta do Servidor: {resposta.text}")
            else:
                print(f"⚠️ Resposta com atenção do Google (Status {resposta.status_code}):")
                print(resposta.text[:300])
                sys.exit(1)

        except Exception as e:
            print(f"❌ Erro de conexão com o Google Sheets: {e}")
            sys.exit(1)

    else:
        print("⚠️ A aba ANALISE está vazia. Nenhuma atualização enviada.")

if __name__ == "__main__":
    executar()
