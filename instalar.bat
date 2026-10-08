@echo off
title Instalacao - Gerador de Boletim COGER v2
echo ============================================
echo GERADOR DE BOLETIM COGER V2 - INSTALACAO
echo ============================================
echo.
py -m venv .venv
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
echo.
echo Dependencias Python instaladas.
echo.
echo IMPORTANTE PARA OCR:
echo Instale o Tesseract OCR em:
echo C:\Program Files\Tesseract-OCR\
echo e inclua o idioma Portugues (por).
echo.
pause
