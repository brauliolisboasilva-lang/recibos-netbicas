@echo off
cd /d "%~dp0"

if not exist "%~dp0venv\Scripts\pythonw.exe" (
    echo O aplicativo ainda nao foi instalado neste computador.
    echo Execute primeiro o arquivo "Instalar.bat".
    pause
    exit /b 1
)

start "" "%~dp0venv\Scripts\pythonw.exe" "%~dp0recibo_app.py"
