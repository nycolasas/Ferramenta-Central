# Passo 2: atualizar_bruta.py

O script `atualizar_bruta.py` é responsável por pegar os dados "crus" baixados pelo script anterior e injetá-los na base de trabalho principal do sistema.

## 🎯 Objetivo
Atualizar a planilha base (`planilha_bruta.xlsx`) com os dados mais recentes extraídos do ERP e armazenados em `dados/relatorios_consolidados.xlsx`.

## ⚙️ Regras e Procedimentos

### 1. Mapeamento de Abas
O script lê as abas de origem e as direciona para abas de destino específicas na planilha bruta:
* `SALDO` ➔ `f_saldo`
* `AVISO_ENTREGA` ➔ `f_registro`
* `PEDIDOS` ➔ `f_4260`
* `ESTOQUE_LOTE` ➔ `f_vt`

### 2. Leitura e Limpeza
* Verifica se ambos os arquivos existem (origem e destino). Se não, levanta uma exceção `FileNotFoundError`.
* Utiliza a biblioteca `pandas` para ler os dados da origem.
* Utiliza a biblioteca `openpyxl` para interagir com o arquivo de destino sem corromper as outras abas existentes (como as abas de configurações manuais `canais`, `atendente`, etc).
* Para cada aba mapeada, ele limpa todo o conteúdo antigo da aba de destino (`ws_dest.delete_rows()`) antes de colar os dados novos, garantindo que não fiquem dados residuais (lixo) na base.

### 3. Tratamento e Conversão de Datas
* O sistema conta com uma lógica inteligente de identificação de colunas de data. Ele busca por colunas com nomes contendo `data`, `dt` ou `date`.
* Todas as datas identificadas são convertidas e forçadas para o formato Data sem horário (00:00:00).
* Trata valores nulos/NaNs/NaTs convertendo-os para strings vazias `""` para evitar quebra no momento de escrever no arquivo `.xlsx`.

### 4. Salvamento
* Reescreve os cabeçalhos e as linhas de dados tratados.
* Salva e fecha a `planilha_bruta.xlsx`. O script também intercepta o `PermissionError` alertando o usuário caso ele esteja com a planilha aberta no Excel bloqueando a gravação.
