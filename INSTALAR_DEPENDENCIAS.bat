@echo off
chcp 65001 > nul
title Instalador de Dependências — Sistema de Pesquisa de Preços
color 0A

echo ===============================================================================
echo     📦  INSTALADOR DE DEPENDÊNCIAS DO SISTEMA
echo ===============================================================================
echo.

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
    echo [ERRO] Python não foi localizado!
    pause
    exit /b 1
)

echo [*] Utilizando interpretador: %PYTHON_EXE%
echo.
echo [*] 1/2 Atualizando pip e instalando pacotes de requirements.txt...
"%PYTHON_EXE%" -m pip install --upgrade pip
"%PYTHON_EXE%" -m pip install -r requirements.txt

echo.
echo [*] 2/2 Baixando e configurando o navegador Chromium (Playwright)...
"%PYTHON_EXE%" -m playwright install chromium

echo.
echo ===============================================================================
echo     ✅  TODAS AS DEPENDÊNCIAS FORAM INSTALADAS COM SUCESSO!
echo     Agora você pode abrir o sistema clicando em INICIAR_SISTEMA.bat
echo ===============================================================================
echo.
pause
