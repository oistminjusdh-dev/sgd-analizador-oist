@echo off
cd /d "%~dp0"
python -m pip install --upgrade pip >NUL 2>&1
pip install Flask pandas pdfplumber openpyxl pytesseract pillow pdf2image >NUL 2>&1
echo Iniciando servidor Flask...
python app.py
pause
