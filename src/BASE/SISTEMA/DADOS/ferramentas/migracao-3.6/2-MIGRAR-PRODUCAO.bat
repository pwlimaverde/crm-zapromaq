@echo off
chcp 65001 >nul
setlocal
set "BANCO=%~dp0..\..\..\..\..\01 - CRM\01 - CONTROLE\BASE\crm_zapromaq.accdb"
set "CARGA=%~dp0..\40 - DICIONARIO\carga"
echo ============================================================
echo  MIGRACAO CRM ZAPROMAQ - PRODUCAO
echo ============================================================
echo  Banco: %BANCO%
echo.
pause
echo.
echo [1/4] criando o banco...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0criar-banco.ps1" -Banco "%BANCO%"
if errorlevel 1 goto :erro
echo.
echo [2/4] carregando os dados...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0carregar.ps1" -Banco "%BANCO%" -Dados "%CARGA%"
if errorlevel 1 goto :erro
echo.
echo [3/4] aplicando integridade...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0aplicar-integridade.ps1" -Banco "%BANCO%"
if errorlevel 1 goto :erro
echo.
echo [4/4] conferindo...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0conferir-base.ps1" -Banco "%BANCO%" -Dados "%CARGA%"
if errorlevel 1 goto :erro
echo.
echo ============================================================
echo  CONCLUIDO SEM DIVERGENCIA
echo ============================================================
goto :fim
:erro
echo.
echo ############################################################
echo  INTERROMPIDO - confira o log da etapa acima
echo ############################################################
:fim
echo.
pause
