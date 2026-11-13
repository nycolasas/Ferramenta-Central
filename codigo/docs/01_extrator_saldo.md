# Passo 1: extrator_saldo.py

O `extrator_saldo.py` autentica no ERP interno e extrai quatro relatórios para
abas do arquivo `dados/relatorios_consolidados.xlsx`. As consultas são feitas
pelos formulários web existentes usando `requests.Session`; o HTML retornado é
interpretado com `BeautifulSoup`. O extrator não inicia um navegador.

## Autenticação e conexão

- Pela interface gráfica, o usuário e a senha são recebidos em `GUI_USER` e
  `GUI_PASS`. Ao executar diretamente em um terminal interativo, as credenciais
  são solicitadas; a senha é digitada sem aparecer na tela. Em execução não
  interativa, defina as duas variáveis de ambiente.
- Uma única sessão HTTP autenticada é compartilhada pelas quatro extrações,
  preservando cookies e evitando logins repetidos.
- As respostas passam por validação de status HTTP e de tela de login antes de
  o conteúdo ser interpretado, evitando salvar uma página de autenticação como
  se fosse um relatório vazio.
- As requisições de relatório têm timeout de 300 segundos. Há uma pausa de
  1,5 segundo entre as consultas em lote do relatório SALDO e entre as
  consultas de cada relatório que possui mais de um lote.
- Cada etapa pode ser repetida até três vezes em caso de falha.

## Relatórios extraídos

1. **SALDO** (`/pcp/situacao_pedidos2.php`): empresa 11, período fixo de
   01/01/2001 a 31/12/2050, dividido em intervalos de representantes.
2. **AVISO_ENTREGA** (`/vdp/consulta_aviso_entrega2.php`): datas informadas na
   interface ou no terminal, tipos AE/OE/AD e pedidos em aberto.
3. **PEDIDOS** (`/vdp/consulta_pedidos.php`): datas informadas e intervalos de
   representantes.
4. **ESTOQUE_LOTE** (`/sup/estoque_lote_ender.php`): situação L, endereço VT,
   tipo F e formato de relatório R.

As datas são recebidas no formato `DD/MM/AAAA` pela interface e convertidas
para `YYYY-MM-DD` antes de serem enviadas ao ERP. Se o script for executado
diretamente sem as datas da interface, solicita o período no terminal.

Os resultados substituem as abas correspondentes no arquivo consolidado.
Cabeçalhos são padronizados e textos de data/número são convertidos para
valores apropriados no Excel.

## Dependências e navegador

`requests` e `beautifulsoup4` são dependências do extrator. O Selenium continua
necessário em outras etapas da automação, em particular no gerador de atrasos;
por isso a interface ainda verifica a disponibilidade do Chrome para executar
a pipeline completa.
