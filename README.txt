============================================
  GOFO PAYROLL SYSTEM — Instrucciones
============================================

CÓMO USAR CADA SEMANA:
───────────────────────
Windows: Doble clic en run.bat
Mac:     Doble clic en run.sh

Luego abre tu navegador en:
  http://localhost:5050

PASOS:
1. Sube los archivos Excel de GOFO (todos a la vez)
2. Agrega bonos extra, cobros o ajusta tarifas si es necesario
3. Revisa los números en la vista previa
4. Descarga el Excel con el reporte completo


PARA CAMBIAR TARIFAS PERMANENTES:
───────────────────────────────────
Abre el archivo app.py con cualquier editor de texto
Busca "CITY_CONFIG" al inicio del archivo
Modifica los valores de "stop" y "extra" por ciudad y ruta


CIUDADES CONFIGURADAS:
───────────────────────
ORD (Chicago)
  Ruta 032: $1.60/stop · $0.25/extra
  Ruta 007: $1.70/stop · $0.40/extra
  Bono fijo: Susy $800

SDF (Louisville)
  Ruta 014: $1.60/stop · $0.50/extra
  Ruta 027: $1.70/stop · $0.50/extra
  Bono fijo: Yadrian $800

MEM (Memphis)
  Ruta 014: $1.60/stop · $0.40/extra
  Ruta 020: $1.79/stop · $0.40/extra
  Ruta 024: $1.89/stop · $0.40/extra
  Bono fijo: Areli $500

BTR (Baton Rouge)
  Ruta 007: $1.50/stop · $0.25/extra
  Bono fijo: Heather $500

SHV (Shreveport)
  Ruta 002: $1.50/stop · $0.25/extra
  Bono fijo: Germain $500


SOPORTE:
─────────
Si algo no funciona, revisa que Python esté
instalado (python.org) y que los archivos Excel
tengan el formato estándar de GOFO.
============================================
