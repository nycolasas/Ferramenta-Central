# Passos 5 e 6: atualizar_conf.py e atualizar_atr.py

Esses dois scripts compartilham a mesma arquitetura. Sua responsabilidade é funcionar como **pontes de integração (ETL Webhooks)** enviando os relatórios processados localmente em Python via API (JSON Payload) para instâncias do Google Sheets/Google Apps Script.

## 🎯 Objetivo
Transmitir os resultados consolidados (`resultado_processamento.xlsx` e `relatorio_atrasados.xlsx`) para a Nuvem de forma estruturada para consumo pelas equipes de linha de frente usando planilhas e formulários do Google Workspace.

## ⚙️ Regras e Procedimentos (atualizar_conf.py)

### Foco: Dados de Confirmações D-2
* **Entrada:** Lê a aba `CONFIRMACAO_D2` de `resultado_processamento.xlsx`.
* **Mapeamento de Colunas:** 
  O script mapeia explicitamente as colunas do Excel em colunas rígidas (A-L) esperadas pelo Google Apps Script. 
  * `Col A (Google)` -> `Col A (Excel)`
  * `Col B (Google)` -> `Col B (Excel)`
  * `Col C (Google)` -> `Col G (Excel)`
  * `Col J (Google)` -> `Col F (Excel)`
  * `Col L (Google)` -> `Col J (Excel)`
  *(As colunas não listadas são preenchidas com strings vazias `''` garantindo integridade estrutural no Sheets)*
* **Ação:** Realiza um `POST Request` no webhook: `https://script.google.com/macros/s/AKfycbzA-yGgYxO-tEZDpZJOpaILmbfQnzbQdpphwkpegL0AD95wGLf83SDZjOcLA6C8mP4v/exec`

## ⚙️ Regras e Procedimentos (atualizar_atr.py)

### Foco: Dados de Análise de Atraso
* **Entrada:** Lê a aba `ANALISE` de `relatorio_atrasados.xlsx`.
* **Mapeamento de Colunas:**
  * `Col A (Google)` -> `Col A (Excel - Pedido)`
  * `Col B (Google)` -> `Col B (Excel - Cliente)`
  * `Col D (Google)` -> `Col G (Excel - Data Entrega)`
  * `Col E (Google)` -> `Col H (Excel - Status Ação)`
  * `Col I (Google)` -> `Col C (Excel - Responsável)`
  *(As demais colunas são transmitidas vazias)*
* **Ação:** Realiza um `POST Request` no webhook: `https://script.google.com/macros/s/AKfycbydJ--n1koGPtmUDCcyNTlGRKUFWHo_SqRve5gSXrt7lnzj7YMyIgy6SCQadt3QDr4S/exec`

### Tratamento de Falhas e Verificações em Ambos os Scripts
* **Formatação Nula:** Ambos trocam dados vazios (NaN/NaT do Pandas) por strings nulas `""` antes da serialização em JSON, prevenindo corrupção no Javascript da Google Cloud.
* **Validação 200 HTTP:** O script verifica se o Status Code da requisição foi 200 e se o Google Apps Script não retornou porcamente uma página HTML ("Service Unavailable" ou erro de permissão no Google Cloud).
* **Bloqueio Fatal:** Se o webhook falhar, o script encerra a execução (`sys.exit(1)`) imprimindo os primeiros 300 caracteres da resposta do Google para depuração.
