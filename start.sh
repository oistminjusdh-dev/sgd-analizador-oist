#!/usr/bin/env bash

# 1. Instalar librerías de sistema necesarias (Tesseract y Poppler)
# La imagen base de Render usa Debian/Ubuntu, por eso usamos 'apt-get'.
echo "Instalando dependencias de sistema (tesseract-ocr y poppler-utils)..."
apt-get update
apt-get install -y tesseract-ocr poppler-utils

# 2. Iniciar la aplicación usando Gunicorn
# 'app:app' indica que ejecute el objeto 'app' que está en el archivo 'app.py'
echo "Iniciando servidor Gunicorn..."
gunicorn --workers 4 --bind 0.0.0.0:$PORT app:app