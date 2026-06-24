@echo off
echo ========================================
echo    GOFO Payroll System
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado.
    echo Descarga Python en: https://www.python.org/downloads/
    pause
    exit
)

REM Install dependencies if needed
echo Verificando dependencias...
pip install flask pandas openpyxl werkzeug --quiet

echo.
echo Iniciando el sistema...
echo.
echo *** Abre tu navegador en: http://localhost:5050 ***
echo.
echo Para cerrar el programa presiona CTRL+C
echo.

cd /d "%~dp0"
python app.py
pause
