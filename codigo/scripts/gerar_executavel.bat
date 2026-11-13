@echo off
chcp 65001 >nul 2>&1
title Empacotador - Gerar Executavel Portavel
color 0E

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║          EMPACOTADOR - Gerar Executavel Portavel            ║
echo ╠══════════════════════════════════════════════════════════════╣
echo ║  Este script gera o .exe portavel que pode ser executado    ║
echo ║  em qualquer maquina Windows (sem precisar de Python).      ║
echo ║                                                             ║
echo ║  Pre-requisito: rode instalar_dependencias.bat primeiro.    ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: Verificar Python
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ Python nao encontrado. Execute instalar_dependencias.bat primeiro.
    pause
    exit /b 1
)

:: Verificar PyInstaller
python -c "import PyInstaller" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ PyInstaller nao encontrado. Execute instalar_dependencias.bat primeiro.
    pause
    exit /b 1
)

echo 🏗️  Iniciando empacotamento...
echo    Isso pode levar de 3 a 10 minutos dependendo da maquina.
echo.

cd /d "%~dp0"
python build.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Falha no empacotamento. Verifique os erros acima.
    pause
    exit /b 1
)

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║              ✅ EXECUTAVEL GERADO COM SUCESSO!              ║
echo ╠══════════════════════════════════════════════════════════════╣
echo ║                                                             ║
echo ║  Local: dist\Painel_Automacao_v5\Painel_Automacao_v5.exe    ║
echo ║                                                             ║
echo ║  Para distribuir, copie TODA a pasta                        ║
echo ║  dist\Painel_Automacao_v5\ para a maquina de destino.       ║
echo ║                                                             ║
echo ║  Requisito na maquina destino: Google Chrome instalado.     ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: Perguntar se quer abrir a pasta
set /p ABRIR="Deseja abrir a pasta do executavel? (S/N): "
if /I "%ABRIR%"=="S" (
    explorer "%~dp0dist\Painel_Automacao_v5"
)

pause
exit /b 0
