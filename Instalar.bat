@echo off
setlocal
title Instalando - Emissao de Recibos NetBicas
color 0B
cd /d "%~dp0"

echo ============================================================
echo   Emissao de Recibos - NetBicas FIBRA
echo   Instalacao no computador
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERRO] O Python nao foi encontrado neste computador.
    echo.
    echo Antes de continuar, instale o Python:
    echo   1. Acesse https://www.python.org/downloads/
    echo   2. Baixe e instale a versao mais recente do Python 3.
    echo   3. IMPORTANTE: marque a opcao "Add python.exe to PATH"
    echo      durante a instalacao.
    echo   4. Depois de instalar, execute este arquivo novamente.
    echo.
    pause
    exit /b 1
)

echo [1/4] Python encontrado. Criando ambiente do aplicativo...
python -m venv "%~dp0venv"
if errorlevel 1 (
    echo [ERRO] Nao foi possivel criar o ambiente do aplicativo.
    pause
    exit /b 1
)

echo [2/4] Instalando componentes necessarios (isso pode levar 1-2 minutos)...
"%~dp0venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
"%~dp0venv\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt" --quiet
if errorlevel 1 (
    echo [ERRO] Falha ao instalar os componentes. Verifique sua conexao com a internet.
    pause
    exit /b 1
)

echo [3/4] Criando pasta de dados...
if not exist "%~dp0dados" mkdir "%~dp0dados"

echo [4/4] Criando atalho na Area de Trabalho...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$s = (New-Object -COM WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Recibos NetBicas.lnk');" ^
    "$s.TargetPath = '%~dp0Abrir_Recibos.bat';" ^
    "$s.WorkingDirectory = '%~dp0';" ^
    "$s.IconLocation = '%~dp0assets\icon.ico';" ^
    "$s.Description = 'Emissao de Recibos - NetBicas FIBRA';" ^
    "$s.Save()"

echo.
echo ============================================================
echo   Instalacao concluida!
echo   Um atalho "Recibos NetBicas" foi criado na Area de Trabalho.
echo   Basta clicar duas vezes nele para abrir o aplicativo.
echo ============================================================
echo.
pause
