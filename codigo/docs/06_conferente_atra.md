# Passo 7: conferente_atra.py

O `conferente_atra.py` é uma **ferramenta de auditoria e depuração local**, projetada para rodar de forma isolada, via terminal interativo, para analisar casos isolados quando há dúvidas sobre a classificação de um determinado pedido no sistema de atrasos.

## 🎯 Objetivo
Fornecer à equipe de monitoramento transparência do algoritmo. Quando um operador questionar "Por que o pedido 201631 não está listado nos avisos de atraso de hoje?", ele deve usar esta ferramenta para imprimir a validação, etapa por etapa.

## ⚙️ Regras e Procedimentos

### 1. Inicialização e Interação
* Exige uma chamada manual interativa via terminal.
* O sistema solicita o número parcial ou completo do pedido via input: `Digite o número ou parte do pedido (ex: 201631):`.

### 2. Busca Híbrida e Extração de Dependências
* Carrega `planilha_bruta.xlsx` contendo saldo, canais e histórico, exatamente como o script produtivo principal faz.
* Realiza um `.str.contains()` ignorando maiúsculas e minúsculas para encontrar todas as linhas no ERP relacionadas ao termo buscado (o pedido pode conter dezenas de itens diferentes na mesma nota).

### 3. Pipeline de Avaliação Logada (7 Etapas)
O script itera sobre as instâncias encontradas do pedido e escreve no console se ela "Aprova" ou "Rejeita" baseado nos rígidos critérios de negócios, um a um:

* **Etapa 1:** Verifica o Prefixo (Empresa `11`).
* **Etapa 2:** Verifica se a natureza do perfil do representante ou cliente aponta como 'LIDER'.
* **Etapa 3:** Valida a Natureza Operacional (Serviços) e o Item (`3TAXA`).
* **Etapa 4:** Verifica o cruzamento híbrido de Região de Minas Gerais (`MG`) sendo do tipo Planejado (`PLAN`).
* **Etapa 5:** Depura o estoque daquele momento específico (Se `Estoque >= Quantidade`, exibe rejeição).
* **Etapa 6:** Validação da Data (descobre se possui data vazia ou data espúria "2040"). Exibe suspensões com alertas amarelos se a Situação for "S".
* **Etapa 7 (Final):** Aplicação de cálculo de corte de dias úteis com matriz de feriados.
  * O script imprime a Data Base do sistema (hoje), os Feriados encontrados e qual foi a "Data de Corte" útil calculada para aquele caso.
  * Exibe se ele entrou na condição `AVISAR`, se acionou histórico prévio, ou, em caso negativo, imprime de forma amigável **o porquê o item não atingiu atraso** (ex: *"Ainda está dentro do prazo limite, o corte seria dia 15/10/2026"*).

### Conclusão
O output é projetado para legibilidade e clareza, imprimindo emojis coloridos no terminal, facilitando o diagnóstico do atendente. Esta ferramenta não salva nem altera dados, é uma simulação de depuração passiva da lógica de produção.
