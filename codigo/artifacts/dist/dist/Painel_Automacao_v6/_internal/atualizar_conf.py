import os
import sys
import json
import pandas as pd
import requests

from portable_paths import RESULT_WORKBOOK, ensure_portable_data_files

ensure_portable_data_files()

# ==============================================================================
# CONFIGURAÇÃO
# ==============================================================================
URL_WEBHOOK_GOOGLE = "https://script.google.com/macros/s/AKfycbzA-yGgYxO-tEZDpZJOpaILmbfQnzbQdpphwkpegL0AD95wGLf83SDZjOcLA6C8mP4v/exec"
caminho_resultado = str(RESULT_WORKBOOK)

# ==============================================================================
# PROCESSAMENTO E ENVIO
# ==============================================================================
def executar():
    print("📂 Lendo dados da aba CONFIRMACAO_D2...")
    
    if not os.path.exists(caminho_resultado):
        print(f"❌ Arquivo não encontrado: {caminho_resultado}")
        sys.exit(1)

    try:
        # Lê a aba CONFIRMACAO_D2 do Excel
        df_confirmacao_d2 = pd.read_excel(caminho_resultado, sheet_name='CONFIRMACAO_D2')
    except Exception as e:
        print(f"❌ Erro ao ler a aba CONFIRMACAO_D2: {e}")
        sys.exit(1)

    if not df_confirmacao_d2.empty:
        # Substitui nulos por string vazia
        df_confirmacao_d2 = df_confirmacao_d2.fillna('')

        # Função auxiliar para extrair coluna pela letra do Excel (A=0, B=1, etc.)
        def col_excel(df, letra):
            idx = ord(letra.upper()) - ord('A')
            if df.shape[1] > idx:
                return df.iloc[:, idx].astype(str)
            return ''

        # ----------------------------------------------------------------------
        # 🎯 MAPEAMENTO DE COLUNAS (Origem Excel -> Destino Google Sheets)
        # ----------------------------------------------------------------------
        df_mapeado = pd.DataFrame()

        df_mapeado['Col_A'] = col_excel(df_confirmacao_d2, 'A')  # A -> A
        df_mapeado['Col_B'] = col_excel(df_confirmacao_d2, 'B')  # B -> B
        df_mapeado['Col_C'] = col_excel(df_confirmacao_d2, 'G')  # G -> C
        df_mapeado['Col_D'] = ''                                 # Vazio
        df_mapeado['Col_E'] = ''                                 # Vazio
        df_mapeado['Col_F'] = ''                                 # Vazio
        df_mapeado['Col_G'] = ''                                 # Vazio
        df_mapeado['Col_H'] = ''                                 # Vazio
        df_mapeado['Col_I'] = ''                                 # Vazio
        df_mapeado['Col_J'] = col_excel(df_confirmacao_d2, 'F')  # F -> J
        df_mapeado['Col_K'] = ''                                 # Vazio
        df_mapeado['Col_L'] = col_excel(df_confirmacao_d2, 'J')  # J -> L

        # Converte em matriz (lista de listas)
        matriz_dados = df_mapeado.values.tolist()

        print("🚀 Enviando dados mapeados para o Google Sheets...")

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
        print("⚠️ A aba CONFIRMACAO_D2 está vazia. Nenhuma atualização enviada.")

if __name__ == "__main__":
    executar()