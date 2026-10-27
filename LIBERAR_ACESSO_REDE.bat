@echo off
chcp 65001 > nul
title Liberar Acesso na Rede Local (Porta 5000) — Lei 14.133/2021
color 0A

echo ===============================================================================
echo     🛡️  CONFIGURAÇÃO DE FIREWALL DO WINDOWS PARA REDE LOCAL
echo ===============================================================================
echo.
echo Liberando permissão de entrada na porta 5000 (TCP)...
netsh advfirewall firewall delete rule name="Sistema Cotacao Pregao (Porta 5000)" >nul 2>nul
netsh advfirewall firewall add rule name="Sistema Cotacao Pregao (Porta 5000)" dir=in action=allow protocol=TCP localport=5000 profile=any >nul 2>nul

if %errorlevel% equ 0 (
    echo.
    echo ===============================================================================
    echo     ✅  PORTA 5000 LIBERADA COM SUCESSO NO FIREWALL!
    echo.
    echo     Seus colegas agora podem acessar diretamente pelo navegador no PC deles:
    for /f "tokens=4" %%a in ('route print^|findstr 0.0.0.0.*0.0.0.0^|findstr /v "Default"') do (
        echo     👉 http://%%a:5000/
    )
    echo ===============================================================================
) else (
    color 0C
    echo.
    echo [AVISO] Para aplicar a regra do Firewall, clique com o botão direito neste
    echo arquivo e escolha "Executar como Administrador".
)
echo.
pause
