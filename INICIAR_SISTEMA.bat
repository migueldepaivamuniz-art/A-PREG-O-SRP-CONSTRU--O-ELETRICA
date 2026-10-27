@echo off
chcp 65001 > nul
title Sistema de Pesquisa de Preços — Pregão SRP (Lei 14.133/2021)
color 0B

echo ===============================================================================
echo     🏛️  SISTEMA DE PESQUISA DE PREÇOS DE MERCADO — PREGÃO SRP
echo     Conforme Instrução Normativa SEGES e Lei Federal nº 14.133/2021
echo ===============================================================================
echo.

:: 1. Localização automática do interpretador Python
set "PYTHON_EXE="

:: Testa python no PATH
where python >nul 2>nul
if %errorlevel% equ 0 (
    set "PYTHON_EXE=python"
    goto :python_found
)

:: Testa py launcher
where py >nul 2>nul
if %errorlevel% equ 0 (
    set "PYTHON_EXE=py"
    goto :python_found
)

:: Procura na pasta padrão do Windows AppData
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
    if exist "%%D\python.exe" (
        set "PYTHON_EXE=%%D\python.exe"
        goto :python_found
    )
)

:python_found
if "%PYTHON_EXE%"=="" (
    color 0C
    echo [ERRO] Python não foi encontrado no seu computador!
    echo Por favor, instale o Python 3.10 ou superior pelo site https://www.python.org/
    echo Certifique-se de marcar a opção "Add Python to PATH" durante a instalação.
    echo.
    pause
    exit /b 1
)

echo [*] Python detectado: %PYTHON_EXE%

:: 2. Verifica configuração do arquivo .env
if not exist ".env" (
    if exist ".env.example" (
        copy ".env.example" ".env" >nul
        echo [*] Arquivo .env criado a partir de .env.example.
    )
)

:: 3. Verifica se as bibliotecas essenciais estão instaladas
"%PYTHON_EXE%" -c "import PySide6, playwright, pandas, openpyxl, PIL" >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo [*] Instalando dependências necessárias... Por favor, aguarde alguns instantes.
    echo.
    "%PYTHON_EXE%" -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        color 0C
        echo.
        echo [ERRO] Falha ao instalar as dependências. Verifique sua conexão com a internet.
        pause
        exit /b 1
    )
    echo [*] Instalando navegador Chromium para o robô de coleta...
    "%PYTHON_EXE%" -m playwright install chromium
)

:: 4. Executa a aplicação gráfica
echo.
echo [*] Iniciando aplicação desktop...
echo.
"%PYTHON_EXE%" app.py

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [AVISO] A aplicação encerrou com código de saída %errorlevel%.
    echo Se ocorreu algum erro inesperado, revise as mensagens acima.
    echo.
    pause
)
