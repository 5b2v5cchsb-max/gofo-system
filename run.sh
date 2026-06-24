#!/bin/bash
echo "========================================"
echo "   GOFO Payroll System"
echo "========================================"
echo ""

# Install dependencies
echo "Verificando dependencias..."
pip3 install flask pandas openpyxl werkzeug --quiet

echo ""
echo "Iniciando el sistema..."
echo ""
echo "*** Abre tu navegador en: http://localhost:5050 ***"
echo ""
echo "Para cerrar el programa presiona CTRL+C"
echo ""

cd "$(dirname "$0")"
python3 app.py
