@echo off
setlocal
cd /d "%~dp0"
set "ORIGEM=C:\Program Files\Tesseract-OCR"
set "DESTINO=%CD%\vendor\tesseract"
echo Preparando Tesseract embutido...
if not exist "%ORIGEM%\tesseract.exe" (
  echo ERRO: Tesseract nao encontrado em %ORIGEM%
  pause
  exit /b 1
)
if not exist "%ORIGEM%\tessdata\por.traineddata" (
  echo ERRO: idioma Portugues nao encontrado.
  pause
  exit /b 1
)
if exist "%DESTINO%" rmdir /s /q "%DESTINO%"
mkdir "%DESTINO%"
xcopy "%ORIGEM%\*" "%DESTINO%\" /E /I /H /Y >nul
echo Tesseract copiado para vendor\tesseract
pause
