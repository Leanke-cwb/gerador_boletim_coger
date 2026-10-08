@echo off
setlocal
title Gerar instalador sem administrador - COGER

cd /d "%~dp0"

echo ============================================
echo GERADOR DE INSTALADOR - COGER V6.10
echo ============================================
echo.

if not exist "dist\Gerador_Boletim_COGER\Gerador_Boletim_COGER.exe" (
    echo O aplicativo ainda nao foi compilado.
    echo Execute gerar_exe_portatil.bat primeiro.
    pause
    exit /b 1
)

set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"

if not exist "%ISCC%" (
    echo Inno Setup 6 nao foi encontrado.
    echo Instale o Inno Setup 6 e execute novamente.
    pause
    exit /b 1
)

"%ISCC%" "instalador.iss"

if errorlevel 1 (
    echo.
    echo Falha ao gerar o instalador.
    pause
    exit /b 1
)

echo.
echo ============================================
echo INSTALADOR CRIADO
echo ============================================
echo.
echo Arquivo:
echo instalador\Instalador_Gerador_Boletim_COGER_v6_10.exe
echo.
echo Esta versao instala no perfil do usuario e nao exige administrador.
echo.
pause
