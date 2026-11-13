"""
build.py — Script de empacotamento PyInstaller
================================================
Gera o executável portátil que qualquer usuário pode rodar sem ter Python
instalado. O .exe resultante funciona como GUI e como "interpretador"
de scripts internos (via --run-script).

Uso:
    python build.py
"""

import PyInstaller.__main__
import sys
import shutil
from pathlib import Path

# Força UTF-8 para evitar erros com emojis no terminal Windows (cp1252)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

NOME_APP = "Painel_Automacao_v6"
BASE_DIR = Path(__file__).resolve().parents[1]

# Scripts que devem ser incluídos como dados (executados via --run-script)
SCRIPTS_DADOS = [
    Path("src") / "etl" / "extrator_saldo.py",
    Path("src") / "etl" / "atualizar_bruta.py",
    Path("src") / "etl" / "gerador_conf.py",
    Path("src") / "etl" / "gerador_atrasos.py",
    Path("src") / "etl" / "atualizar_conf.py",
    Path("src") / "etl" / "atualizar_atr.py",
    Path("src") / "etl" / "conferente_atra.py",
]

def limpar_build_anterior():
    """Remove artefatos de build anteriores para evitar conflitos."""
    for pasta in ["build", "dist"]:
        caminho = BASE_DIR / pasta
        if caminho.exists():
            print(f"🗑️  Removendo {pasta}/...")
            shutil.rmtree(caminho, ignore_errors=True)

    spec = BASE_DIR / f"{NOME_APP}.spec"
    if spec.exists():
        spec.unlink()
        print(f"🗑️  Removendo {spec.name}...")


def construir():
    print("=" * 60)
    print(f"🏗️  Empacotando: {NOME_APP}")
    print("=" * 60)

    limpar_build_anterior()

    # Monta os argumentos do PyInstaller
    args = [
        str(BASE_DIR / "interface_executor.py"),
        "--noconfirm",
        "--onedir",
        "--windowed",
        f"--name={NOME_APP}",
        f"--distpath={BASE_DIR / 'dist'}",
        f"--workpath={BASE_DIR / 'build'}",
        f"--specpath={BASE_DIR}",
        # Coleta todas as dependências do projeto
        "--collect-all=selenium",
        "--collect-all=pandas",
        "--collect-all=numpy",
        "--collect-all=openpyxl",
        "--collect-all=requests",
        "--collect-all=bs4",
        "--collect-all=customtkinter",
        # Hidden imports necessários
        "--hidden-import=selenium.webdriver.chrome.service",
        "--hidden-import=selenium.webdriver.chrome.options",
        "--hidden-import=selenium.webdriver.common.by",
        "--hidden-import=selenium.webdriver.support.ui",
        "--hidden-import=selenium.webdriver.support.expected_conditions",
        "--hidden-import=selenium.webdriver.remote.client_config",
        "--hidden-import=openpyxl.styles",
        "--hidden-import=numpy",
        "--hidden-import=pandas",
        "--hidden-import=json",
        "--hidden-import=unicodedata",
        "--hidden-import=requests",
        "--hidden-import=bs4",
    ]

    # Adiciona scripts como dados
    for script in SCRIPTS_DADOS:
        caminho = BASE_DIR / script
        if caminho.exists():
            args.append(f"--add-data={caminho};{script.parent.as_posix()}")
            print(f"  📄 Incluindo script: {script.as_posix()}")
        else:
            print(f"  ⚠️  Script não encontrado: {script.as_posix()}")

    # Os arquivos modelo são lidos de RESOURCE_ROOT/data e escritos em APP_ROOT/data.
    data_dir = BASE_DIR / "data"
    if data_dir.exists():
        args.append(f"--add-data={data_dir};data")
        print("  📁 Incluindo diretório: data/")
    else:
        print(f"  ⚠️  Diretório de dados não encontrado: {data_dir}")

    print(f"\n🔧 Executando PyInstaller...")
    print("-" * 60)

    PyInstaller.__main__.run(args)

    # Verifica resultado
    dist_dir = BASE_DIR / "dist" / NOME_APP
    exe_path = dist_dir / f"{NOME_APP}.exe"

    if exe_path.exists():
        tamanho_mb = exe_path.stat().st_size / (1024 * 1024)
        print("\n" + "=" * 60)
        print(f"✅ Empacotamento concluído com sucesso!")
        print(f"   Executável: {exe_path}")
        print(f"   Tamanho do .exe: {tamanho_mb:.1f} MB")

        # Calcula tamanho total da pasta dist
        total_dist = sum(f.stat().st_size for f in dist_dir.rglob("*") if f.is_file())
        print(f"   Tamanho total (dist): {total_dist / (1024 * 1024):.1f} MB")
        print("=" * 60)
    else:
        print("\n❌ ERRO: Executável não foi gerado!")
        sys.exit(1)


if __name__ == "__main__":
    construir()
