from __future__ import annotations

import sys
import shutil
from pathlib import Path

from openpyxl import Workbook


def resolve_app_root() -> Path:
    """Resolve a pasta base do app, tanto no código-fonte quanto no executável empacotado."""
    if getattr(sys, "frozen", False):
        # Use a pasta do executável para manter planilhas e configurações
        # persistentes e permitir executar o aplicativo fora de diretórios de sistema.
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_ROOT = resolve_app_root()
RESOURCE_ROOT = (
    Path(sys._MEIPASS).resolve()
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")
    else APP_ROOT
)
DATA_DIR = APP_ROOT / "dados"
BASE_WORKBOOK = APP_ROOT / "planilha_bruta.xlsx"
RESULT_WORKBOOK = APP_ROOT / "resultado_processamento.xlsx"
ATRASADOS_WORKBOOK = APP_ROOT / "relatorio_atrasados.xlsx"
FERIADOS_WORKBOOK = APP_ROOT / "Feriados.xlsx"
SOURCE_WORKBOOK = DATA_DIR / "relatorios_consolidados.xlsx"


def _preparar_planilha(caminho: Path, abas: list[str], modelo: Path | None = None) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    if caminho.exists():
        return
    if modelo and modelo.exists():
        shutil.copy2(modelo, caminho)
        return

    wb = Workbook()
    # Mantém apenas as abas que a automação precisa; os arquivos iniciais são vazios.
    padrao = wb.active
    padrao.title = abas[0] if abas else "Planilha"

    for nome in abas[1:]:
        wb.create_sheet(title=nome)

    wb.save(caminho)
    wb.close()


def ensure_portable_data_files() -> dict[str, Path]:
    """Cria os arquivos base do projeto quando eles não existem, mantendo a lógica atual do fluxo."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    _preparar_planilha(
        SOURCE_WORKBOOK,
        ["SALDO", "AVISO_ENTREGA", "PEDIDOS", "ESTOQUE_LOTE"],
        RESOURCE_ROOT / "dados" / SOURCE_WORKBOOK.name,
    )
    _preparar_planilha(
        BASE_WORKBOOK,
        [
            "f_saldo",
            "f_4260",
            "f_vt",
            "f_registro",
            "desconsiderar",
            "canais",
            "atendente",
            "HISTORICO SALDO",
        ],
        RESOURCE_ROOT / BASE_WORKBOOK.name,
    )
    _preparar_planilha(RESULT_WORKBOOK, ["CONFIRMACAO_D2", "HISTORICO SALDO"])
    _preparar_planilha(ATRASADOS_WORKBOOK, ["ANALISE"])
    _preparar_planilha(
        FERIADOS_WORKBOOK,
        ["Feriados"],
        RESOURCE_ROOT / FERIADOS_WORKBOOK.name,
    )

    return {
        "origem": SOURCE_WORKBOOK,
        "base": BASE_WORKBOOK,
        "resultado": RESULT_WORKBOOK,
        "atrasados": ATRASADOS_WORKBOOK,
        "feriados": FERIADOS_WORKBOOK,
    }
