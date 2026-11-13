# Passo 3: gerador_conf.py

O script `gerador_conf.py` é o "cérebro" das regras de negócio relacionadas a confirmações de entrega. Ele cruza os relatórios extraídos, mapeia os responsáveis por cada pedido e define quem receberá as confirmações.

## 🎯 Objetivo
Processar a base bruta, limpar dados irrelevantes, aplicar cruzamentos relacionais, corrigir erros de ERP e aplicar cortes de data (D-2) gerando as planilhas finais prontas para despacho na planilha `resultado_processamento.xlsx`.

## ⚙️ Regras e Procedimentos

### 1. Carregamento Flexível
O script implementa funções flexíveis de carregamento e busca de colunas, conseguindo localizar nomes de abas e colunas mesmo com pequenas variações de acentuação, espaços e maiúsculas/minúsculas.
* Carrega: `f_saldo`, `f_4260`, `f_vt`, `f_registro`, `desconsiderar`, `canais`, `atendente`.
* Tenta carregar o Histórico (`HISTORICO SALDO`) do processamento anterior para saber quais itens mudaram de status.

### 2. Limpeza de Pedidos (Chaves)
* Cria chaves únicas de Pedidos extraindo apenas a raiz numérica (ignorando sufixos e zeros à esquerda).
* Garante que os pedidos na aba `f_saldo` pertençam à Empresa 11 (aplicando prefixo `(11)`).
* Remove do processamento itens com código iniciando por **"3TAXA"**.

### 3. Cruzamento Relacional (Atendentes e Regiões)
* Cruza informações da aba `canais` para mapear cada "Representante" para uma "Loja" e "Região".
* O pedido é separado em `PLANEJADOS` (se tiver a Linha de Produto "4-PLANEJADOS") ou `SOLTOS`.
* Com base no tipo do produto e na loja, busca na aba `atendente` quem é a pessoa específica encarregada de entrar em contato com o cliente.

### 4. Correção de Erro Humano (Virtualização de Status)
* **Regra 2040:** Itens do tipo `PLANEJADOS`, que estão com a Situação `N` (Normal), mas cuja Data de Entrega está preenchida como o ano de **2040**, possuem um erro de registro.
* O sistema aplica uma "correção virtual": altera momentaneamente a Situação de `N` para `S` (Suspenso) para essas linhas no DataFrame. Registra esses itens corrigidos numa aba especial de log `ERRO_CADASTRO_2040`.

### 5. Itens que "Saíram do S"
* Compara o snapshot histórico do dia anterior com a extração de hoje.
* Se um item (Chave do Pedido + Código do Item) estava com a Situação `S` (Suspenso) ontem e hoje está diferente de `S`, ele é isolado na aba `SAIRAM_DE_S`.

### 6. Regra do D-2 (Corte de Datas)
* O processo separa itens atrasados/normais utilizando uma janela de retrocesso de 2 Dias Úteis (D-2).
* **Condição de Terça-feira:** Se hoje for terça-feira, o intervalo D-2 considera de -4 dias até -2 dias atrás.
* Separa o escopo "Normal/Planejados": Filtra apenas itens com Situação 'N' ou 'S' dependendo de serem soltos ou planejados, com base na data do D-2.

### 7. Filtros e Exclusões (Exceções de Negócio)
* **Filtro VT:** Cruza e remove itens que constam na aba `f_vt` (Estoque Lote Endereçado VT).
* **Filtro Desconsiderar:** Remove pedidos que estejam cadastrados na aba `desconsiderar` com motivos específicos como 'ENTREGA AGENDADA', 'JÁ FATURADO'.
* **Naturezas Permitidas:** Apenas itens com as Naturezas de Operação: `7002`, `7005`, `8025`, `7025`, `7004`.
* **Linhas de Produto Excluídas:** Prefixos `24`, `25`, `99`.
* **Representantes Excluídos:** Numeração de representantes entre `5000` e `5999`.

### 8. Validação de Confirmações Prévias (f_registro)
* Analisa o histórico de ligações na aba `f_registro`.
* Procura por motivos como `AGUARDANDO CARGA`, `AGENDADO` ou ligações de `CONFIRMAÇÃO DE DATA DE ENTREGA`.
* Se o pedido/etapa já tiver um histórico de sucesso validado, ele é retirado da fila pendente de confirmação.

### 9. Desambiguação e Geração de Bases
* Separa os resultados finais confirmados em duas abas principais baseando-se num **limite fixo de data: 31/07/2026**.
  * Antes desse limite: Aba `CONFIRMACAO_ALEM_D2`.
  * Depois desse limite: Aba `CONFIRMACAO_D2`.
* Gera rótulos descritivos (`1ª Confirmação (S)`, `2ª Confirmação (N)`, `Confirmação Única (N)`).
* Normaliza e filtra motivos indevidos.
* Atualiza a aba do histórico com a foto dos status de hoje.

### 10. Saída
Salva o arquivo `resultado_processamento.xlsx` contendo as múltiplas abas relatoriais para os times operacionais e integrações futuras.
