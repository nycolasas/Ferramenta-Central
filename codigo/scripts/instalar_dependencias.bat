@echo off
chcp 65001 >nul 2>&1
title Instalador de Dependencias - Painel de Automacao v5
color 0A

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   INSTALADOR DE DEPENDENCIAS - Painel de Automacao v5      ║
echo ╠══════════════════════════════════════════════════════════════╣
echo ║  Este script instala tudo que voce precisa para compilar    ║
echo ║  ou executar a ferramenta a partir do codigo-fonte.         ║
echo ║                                                             ║
echo ║  Se voce ja possui o .exe empacotado, NAO precisa rodar     ║
echo ║  este instalador (basta executar o .exe diretamente).       ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: ==================================================================
:: 1. Verificar se Python esta instalado
:: ==================================================================
echo [1/5] Verificando Python...
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Python NAO encontrado!
    echo.
    echo    Faca o download em: https://www.python.org/downloads/
    echo    IMPORTANTE: Marque "Add Python to PATH" durante a instalacao.
    echo.
    echo    Apos instalar, feche e reabra este terminal e execute
    echo    este script novamente.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYTHON_VER=%%i
echo    ✅ %PYTHON_VER% encontrado.
echo.

:: ==================================================================
:: 2. Verificar/criar ambiente virtual (opcional mas recomendado)
:: ==================================================================
echo [2/5] Verificando pip...
python -m pip --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo    ⚠️ pip nao encontrado. Instalando...
    python -m ensurepip --upgrade
)
echo    ✅ pip disponivel.
echo.

:: ==================================================================
:: 3. Atualizar pip
:: ==================================================================
echo [3/5] Atualizando pip...
python -m pip install --upgrade pip --quiet
echo    ✅ pip atualizado.
echo.

:: ==================================================================
:: 4. Instalar dependencias do requirements.txt
:: ==================================================================
echo [4/5] Instalando dependencias Python...
echo    Isso pode levar alguns minutos na primeira vez...
echo.

python -m pip install customtkinter --quiet
if %ERRORLEVEL% NEQ 0 (
    echo    ❌ Falha ao instalar customtkinter
    goto :erro
)
echo    ✅ customtkinter

python -m pip install pandas --quiet
if %ERRORLEVEL% NEQ 0 (
    echo    ❌ Falha ao instalar pandas
    goto :erro
)
echo    ✅ pandas

python -m pip install numpy --quiet
if %ERRORLEVEL% NEQ 0 (
    echo    ❌ Falha ao instalar numpy
    goto :erro
)
echo    ✅ numpy

python -m pip install openpyxl --quiet
if %ERRORLEVEL% NEQ 0 (
    echo    ❌ Falha ao instalar openpyxl
    goto :erro
)
echo    ✅ openpyxl

python -m pip install selenium --quiet
if %ERRORLEVEL% NEQ 0 (
    echo    ❌ Falha ao instalar selenium
    goto :erro
)
echo    ✅ selenium

python -m pip install requests --quiet
if %ERRORLEVEL% NEQ 0 (
    echo    ❌ Falha ao instalar requests
    goto :erro
)
echo    ✅ requests

python -m pip install pyinstaller --quiet
if %ERRORLEVEL% NEQ 0 (
    echo    ❌ Falha ao instalar pyinstaller
    goto :erro
)
echo    ✅ pyinstaller

echo.

:: ==================================================================
:: 5. Verificar Google Chrome
:: ==================================================================
echo [5/5] Verificando Google Chrome...

set CHROME_FOUND=0

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    set CHROME_FOUND=1
)
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    set CHROME_FOUND=1
)
if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    set CHROME_FOUND=1
)

if %CHROME_FOUND%==1 (
    echo    ✅ Google Chrome encontrado.
) else (
    echo    ⚠️ Google Chrome NAO encontrado.
    echo    A ferramenta precisa do Chrome para as etapas de extracao.
    echo    Download: https://www.google.com/chrome/
)

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║               INSTALACAO CONCLUIDA!                         ║
echo ╠══════════════════════════════════════════════════════════════╣
echo ║                                                             ║
echo ║  Para EXECUTAR (modo desenvolvimento):                      ║
echo ║    python interface_executor.py                             ║
echo ║                                                             ║
echo ║  Para EMPACOTAR em .exe (portavel):                         ║
echo ║    python build.py                                          ║
echo ║                                                             ║
echo ║  O .exe sera gerado em: dist\Painel_Automacao_v5\           ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
pause
exit /b 0

:erro
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║           ❌ ERRO NA INSTALACAO                             ║
echo ║  Verifique sua conexao com a internet e tente novamente.    ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
pause
exit /b 1
