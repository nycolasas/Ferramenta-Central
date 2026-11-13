# Automação de Pedidos e Atrasos

Este projeto automatiza a extração, processamento e atualização de dados de pedidos, saldo e atrasos em planilhas Excel e integrações com Google Sheets.

## Visão geral

A pipeline principal segue esta ordem:

1. Extrator de saldo
2. Atualização da planilha bruta
3. Geração de confirmação D-2
4. Geração de relatório de atrasos
5. Atualização do Google Sheets de confirmação
6. Atualização do Google Sheets de atrasos

## Estrutura do projeto

```text
.
├─ data/                         # planilhas de entrada, saída e referência
│  ├─ raw/
│  ├─ processed/
│  └─ reference/
├─ docs/                         # documentação por etapa
├─ src/                          # utilitários compartilhados e constantes do sistema
│  ├─ __init__.py
│  ├─ paths.py                  # resolução de caminhos, arquivos base e inicialização dos dados
│  ├─ etl/                      # scripts da pipeline
│  └─ ui/                       # implementação da interface e controle do pipeline
├─ ui/                           # compatibilidade para imports antigos
│  ├─ __init__.py
│  └─ interface_executor.py
├─ interface_executor.py          # wrapper de compatibilidade com o entrypoint legado
├─ portable_paths.py             # compatibilidade para imports antigos
├─ scripts/
│  └─ build.py                   # empacotamento com PyInstaller
├─ requirements.txt              # dependências do projeto
├─ README.md                     # documentação principal
├─ .gitignore                    # arquivos localmente gerados e artefatos de build
└─ ...
```

A organização foi separada em duas camadas principais:
- `src/`: lógica compartilhada, configurações e arquivos base do app
- `src/ui/`: interface e orquestração visual da pipeline

O arquivo `interface_executor.py` permanece como ponto de entrada compatível para não quebrar scripts ou scripts de execução antigos.

## Como executar

### Ambiente Python

```bash
pip install -r requirements.txt
```

### Fluxo principal

```bash
python extrator_saldo.py
python atualizar_bruta.py
python gerador_conf.py
python gerador_atrasos.py
python atualizar_conf.py
python atualizar_atr.py
```

### Interface gráfica

```bash
python interface_executor.py
```

## Observações importantes

- O extrator consulta os relatórios do ERP por HTTP (`requests` + `BeautifulSoup`), mantendo o login e os filtros dos formulários existentes. O Selenium continua sendo usado na etapa de geração de atrasos.
- Os caminhos de arquivos foram centralizados para funcionar tanto em ambiente de desenvolvimento quanto em execução empacotada.
- A pasta `build/` e `dist/` são artefatos gerados e não fazem parte da lógica de negócio.

## Documentação detalhada

A documentação por etapa está em `docs/` com a visão geral e explicações de cada script.
