@echo off
chcp 65001 >nul
setlocal EnableExtensions

rem Todo caminho sai de %~dp0: aberto por caminho de rede, o CMD nao
rem suporta UNC e cai em C:\Windows. Aspas ficam no ponto de uso.

set "PS1=%~dp0aplicar-indice-cnpj.ps1"

echo ============================================================
echo  INDICE UNICO DE CNPJ - CRM Comercial Zapromaq (PRODUCAO)
echo ============================================================
echo.
echo  Impede DOIS clientes com o MESMO CNPJ preenchido.
echo  NAO impede duplicata de cadastro SEM CNPJ - o Access aceita
echo  varios nulos num indice unico, e 54%% da base esta sem CNPJ.
echo.
echo  Feche o CRM em TODAS as estacoes antes de continuar.
echo.

if not exist "%PS1%" (
    echo  ERRO: aplicar-indice-cnpj.ps1 nao encontrado em:
    echo    %~dp0
    goto :fim
)

pause
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
set "RC=%errorlevel%"
echo.

rem Nao confiar so no errorlevel: confira o log citado acima.
if not "%RC%"=="0" (
    echo  TERMINOU COM FALHA - codigo %RC%. Leia o log-indice-cnpj-*.txt.
) else (
    echo  CONCLUIDO. Leia o log-indice-cnpj-*.txt para a conferencia.
)

:fim
echo.
pause
endlocal
