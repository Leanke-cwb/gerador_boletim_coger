@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Execute instalar.bat primeiro.
  pause
  exit /b 1
)
if not exist "vendor\tesseract\tesseract.exe" (
  echo Execute preparar_tesseract.bat primeiro.
  pause
  exit /b 1
)
if not exist "vendor\tesseract\tessdata\por.traineddata" (
  echo Idioma Portugues ausente no Tesseract embutido.
  pause
  exit /b 1
)
call .venv\Scripts\activate
if exist build rmdir /s /q build
if exist "dist\Gerador_Boletim_COGER" rmdir /s /q "dist\Gerador_Boletim_COGER"
pyinstaller --noconfirm --clean Gerador_Boletim_COGER.spec
if errorlevel 1 (
  echo Falha no build.
  pause
  exit /b 1
)
echo Build concluido em dist\Gerador_Boletim_COGER
pause
