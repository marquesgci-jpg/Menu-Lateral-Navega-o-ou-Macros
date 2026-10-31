@echo off
setlocal
title Instalador do Menu Lateral de Navegacao para Excel

if not exist "%~dp0setup.exe" (
    echo [ERRO] setup.exe nao foi encontrado.
    echo Extraia todos os arquivos do pacote ZIP para a mesma pasta.
    pause
    exit /b 1
)

echo Abrindo o instalador do suplemento...
start /wait "" "%~dp0setup.exe"
set "installResult=%ERRORLEVEL%"

if not "%installResult%"=="0" (
    echo.
    echo [ERRO] O instalador terminou com codigo %installResult%.
    pause
    exit /b %installResult%
)

echo.
echo O instalador foi encerrado. Abra o Excel e confira Suplementos COM.
pause
exit /b 0
