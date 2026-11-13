from __future__ import annotations

import shutil
import sys
from pathlib import Path

from openpyxl import Workbook


def resolve_app_root() -> Path:
    """Resolve a pasta base do app, tanto no código-fonte quanto no executável empacotado."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


APP_ROOT = resolve_app_root()
RESOURCE_ROOT = (
    Path(sys._MEIPASS).resolve()
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")
    else APP_ROOT
)
DATA_DIR = APP_ROOT / "data"
RAW_DATA_DIR = DATA_DIR / "raw"
PROCESSED_DATA_DIR = DATA_DIR / "processed"
REFERENCE_DATA_DIR = DATA_DIR / "reference"
BASE_WORKBOOK = PROCESSED_DATA_DIR / "planilha_bruta.xlsx"
RESULT_WORKBOOK = PROCESSED_DATA_DIR / "resultado_processamento.xlsx"
ATRASADOS_WORKBOOK = PROCESSED_DATA_DIR / "relatorio_atrasados.xlsx"
FERIADOS_WORKBOOK = REFERENCE_DATA_DIR / "Feriados.xlsx"
SOURCE_WORKBOOK = RAW_DATA_DIR / "relatorios_consolidados.xlsx"


def _preparar_planilha(caminho: Path, abas: list[str], modelo: Path | None = None) -> None:
    """Cria planilhas base quando o arquivo ainda não existe."""
    caminho.parent.mkdir(parents=True, exist_ok=True)
    if caminho.exists():
        return
    if modelo and modelo.exists():
        shutil.copy2(modelo, caminho)
        return

    wb = Workbook()
    padrao = wb.active
    padrao.title = abas[0] if abas else "Planilha"

    for nome in abas[1:]:
        wb.create_sheet(title=nome)

    wb.save(caminho)
    wb.close()


def ensure_portable_data_files() -> dict[str, Path]:
    """Cria os arquivos base do projeto quando eles não existem, preservando o fluxo atual."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DATA_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSED_DATA_DIR.mkdir(parents=True, exist_ok=True)
    REFERENCE_DATA_DIR.mkdir(parents=True, exist_ok=True)

    _preparar_planilha(
        SOURCE_WORKBOOK,
        ["SALDO", "AVISO_ENTREGA", "PEDIDOS", "ESTOQUE_LOTE"],
        RESOURCE_ROOT / "data" / "raw" / SOURCE_WORKBOOK.name,
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
        RESOURCE_ROOT / "data" / "processed" / BASE_WORKBOOK.name,
    )
    _preparar_planilha(RESULT_WORKBOOK, ["CONFIRMACAO_D2", "HISTORICO SALDO"])
    _preparar_planilha(ATRASADOS_WORKBOOK, ["ANALISE"])
    _preparar_planilha(
        FERIADOS_WORKBOOK,
        ["Feriados"],
        RESOURCE_ROOT / "data" / "reference" / FERIADOS_WORKBOOK.name,
    )

    return {
        "origem": SOURCE_WORKBOOK,
        "base": BASE_WORKBOOK,
        "resultado": RESULT_WORKBOOK,
        "atrasados": ATRASADOS_WORKBOOK,
        "feriados": FERIADOS_WORKBOOK,
    }


__all__ = [
    "APP_ROOT",
    "RESOURCE_ROOT",
    "DATA_DIR",
    "BASE_WORKBOOK",
    "RESULT_WORKBOOK",
    "ATRASADOS_WORKBOOK",
    "FERIADOS_WORKBOOK",
    "SOURCE_WORKBOOK",
    "resolve_app_root",
    "ensure_portable_data_files",
]
