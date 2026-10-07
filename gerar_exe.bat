@echo off
title Gerar EXE - Gerador de Boletim COGER v5
call .venv\Scripts\activate
echo ============================================
echo GERANDO EXECUTAVEL - V5
echo ============================================
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name Gerador_Boletim_COGER ^
  --add-data "assets;assets" ^
  --collect-all customtkinter ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  main.py
echo.
echo Executavel gerado em:
echo dist\Gerador_Boletim_COGER.exe
echo.
echo Microsoft Word e necessario para atualizar o indice e gerar o PDF.
echo Tesseract OCR e necessario apenas para importar prints.
pause
