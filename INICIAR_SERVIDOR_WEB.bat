@echo off
chcp 65001 > nul
title Servidor Web Local — Sistema de Pesquisa de Preços (Lei 14.133/2021)
color 0B

echo ===============================================================================
echo     🌐  SISTEMA DE PESQUISA DE PREÇOS — SERVIDOR WEB LOCAL
echo     Conforme Instrução Normativa SEGES e Lei Federal nº 14.133/2021
echo ===============================================================================
echo.

:: 1. Localização automática do Python
set "PYTHON_EXE="
where python >nul 2>nul
if %errorlevel% equ 0 set "PYTHON_EXE=python"
if "%PYTHON_EXE%"=="" (
    where py >nul 2>nul
    if %errorlevel% equ 0 set "PYTHON_EXE=py"
)
if "%PYTHON_EXE%"=="" (
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
        if exist "%%D\python.exe" set "PYTHON_EXE=%%D\python.exe"
    )
)

if "%PYTHON_EXE%"=="" (
    color 0C
    echo [ERRO] Python não foi encontrado no seu computador!
    echo Instale o Python pelo site https://www.python.org/
    pause
    exit /b 1
)

echo [*] Python detectado: %PYTHON_EXE%

:: 2. Verifica se o Flask está instalado
"%PYTHON_EXE%" -c "import flask" >nul 2>nul
if %errorlevel% neq 0 (
    echo [*] Instalando framework web (Flask)...
    "%PYTHON_EXE%" -m pip install flask
)

:: 3. Inicia o servidor web e abre no navegador padrão
echo.
echo ===============================================================================
echo     🚀 INICIANDO SERVIDOR WEB LOCAL EM: http://127.0.0.1:5000/
echo     O navegador abrirá automaticamente em instantes...
echo     Para encerrar o servidor, basta fechar esta janela ou teclar Ctrl + C.
echo ===============================================================================
echo.

"%PYTHON_EXE%" app.py --web

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [AVISO] O servidor encerrou com código %errorlevel%.
    pause
)
