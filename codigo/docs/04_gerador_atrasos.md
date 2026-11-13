# Passo 4: gerador_atrasos.py

O script `gerador_atrasos.py` faz a detecção analítica de pedidos que ultrapassaram o tempo limite estipulado pela empresa e define se a equipe deve acionar o cliente, baseado num histórico de reavisos e prazos regionais.

## 🎯 Objetivo
Avaliar o saldo pendente, cruzar com o histórico de contato (`f_registro`), descontar feriados, validar volumes de estoque e emitir ações claras: "AVISAR" ou "REAVISAR ATRASO". O resultado é exportado para `relatorio_atrasados.xlsx`.

## ⚙️ Regras e Procedimentos

### 1. Ingestão de Dados
* Lê `planilha_bruta.xlsx` (abas `f_saldo`, `f_registro`/`HISTORICO`, `canais`, `atendente`).
* Lê os feriados parametrizados a partir do arquivo `Feriados.xlsx`.
* O script executa uma instância rápida do Selenium na inicialização ("simulando" uma consulta web passiva/aguardando sincronia do servidor) para garantir latências antes da quebra da base de dados local.

### 2. Filtros de Exclusão Básicos
Ao varrer o saldo, ignora e remove os seguintes pedidos do escopo de atrasados:
* Pedidos que não começam com `11-` (Empresa 11).
* Pedidos em que o campo de Liderança (`LIDER`, `PERFIL`, `GRUPO`) seja igual a "LIDER".
* O Cliente Final (Ou Cliente de Destino) inicie com a palavra "LIDER".
* Itens específicos iniciados com "3TAXA".
* Naturezas operacionais que correspondam a "24-SERVIÇOS E INSUMOS", "25-PROTECAO DE TECIDOS" e "99-GENERICO".

### 3. Exceção Região de Minas Gerais (Planejados)
* Há um filtro específico que impede que itens "Planejados" (`PLAN` na natureza ou que comecem com `4-`) enviados para o estado de **Minas Gerais (MG)** sejam classificados neste fluxo. Eles são retirados da lista geral.

### 4. Validação Físico-Lógica de Estoque
* Antes de considerar que um item está "atrasado para a entrega", o sistema verifica o estoque fisicamente disponível vs a quantidade comprada (`QTD`).
* Se o **Estoque Disponível >= Quantidade Pendente**, o item é descartado, pois há produto em armazém para honrar o saldo e a falha de entrega provavelmente reside em logística em trânsito/faturamento, e não em atraso produtivo.

### 5. Matriz de Prazos (Dias Úteis)
A data limite não é um simples comparativo D-0, o sistema estipula janelas de limite de transporte/fabricação com base na região de destino, subtraindo "Dias Úteis" e "Feriados" a partir da data de entrega prevista:
* **Minas Gerais (MG):** 3 dias úteis.
* **Bahia (BA):** 9 dias úteis.
* **Demais Estados:** 7 dias úteis.

### 6. Motor Lógico de Aviso e Reaviso
A avaliação compara a **Data de Entrega Prometida** com o **Histórico de Avisos** do cliente (quando ele foi avisado na última vez).

**Cenário 1: Cliente Nunca foi Avisado (AVISAR)**
* Verifica a `Data de Entrega`. Subtrai os dias úteis (Ex: 7 dias).
* Se a data de hoje ultrapassar a data limite deduzida ou a data de entrega final, recebe o status **"AVISAR"**.

**Cenário 2: Cliente Já foi Avisado (REAVISAR ATRASO)**
* O sistema localiza o histórico de previsões de entrega passadas no `f_registro`.
* Se a data de entrega atual do sistema tiver **aumentado** (foi estendida em relação ao aviso que demos antes): O sistema calcula a nova janela de prazo baseada na nova data estendida. Ele só aciona o reaviso se a data limite dessa nova data também for estourada.
* Se a data não foi estendida, baseia-se na data do último aviso. Se já passou o prazo a partir dela, gera o status **"REAVISAR ATRASO"**.

### 7. Pedidos Suspensos (Status 'S')
Pedidos cujo ERP aponta situação 'S' não recebem cálculo de dias de atraso comuns. Eles são extraídos, rotulados como `SIM` em um log isolado para análise de itens "Suspensos".

### 8. Saídas (Excel)
Gera o `relatorio_atrasados.xlsx` contendo 3 abas essenciais:
* `ANALISE`: Ações diárias a serem tomadas para itens elegíveis de aviso e reaviso.
* `HISTORICO_ATRASOS`: Banco de dados evolutivo com todas as análises feitas com data/hora de execução anexada.
* `PEDIDO_SUSPENSO`: Acumulado de pedidos flagrados em suspensão.
