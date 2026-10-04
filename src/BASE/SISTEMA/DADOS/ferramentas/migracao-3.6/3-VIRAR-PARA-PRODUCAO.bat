@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
set "BANCO=%~dp0..\..\..\..\..\01 - CRM\01 - CONTROLE\BASE\crm_zapromaq.accdb"
set "CARGA=%~dp0..\40 - DICIONARIO\carga"
echo ============================================================
echo  VIRADA PARA PRODUCAO - CRM Comercial Zapromaq
echo ============================================================
echo.
echo  ANTES DE CONTINUAR, confirme:
echo.
echo   [ ] Ninguem com a planilha CRM-Comercial-Zapromaq-3.6 aberta
echo   [ ] Ninguem com o CRM_Zapromaq.xlsm aberto
echo   [ ] A extracao foi refeita depois da ultima alteracao na planilha
echo.
echo  Este processo cria o banco de PRODUCAO e reconfigura o modelo.
echo  Se o banco de producao ja existir, o processo para - e proposital.
echo.
pause
echo.
echo [1/5] criando o banco de producao...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0criar-banco.ps1" -Banco "%BANCO%"
if errorlevel 1 goto :erro
echo.
echo [2/5] carregando os dados...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0carregar.ps1" -Banco "%BANCO%" -Dados "%CARGA%"
if errorlevel 1 goto :erro
echo.
echo [3/5] aplicando integridade...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0aplicar-integridade.ps1" -Banco "%BANCO%"
if errorlevel 1 goto :erro
echo.
echo [4/5] conferindo a carga...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0conferir-base.ps1" -Banco "%BANCO%" -Dados "%CARGA%"
if errorlevel 1 goto :erro
echo.
echo [5/5] regravando o modelo apontando para PRODUCAO...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0montar-frontend.ps1" -Banco "%BANCO%" -Ambiente PRODUCAO
if errorlevel 1 goto :erro
echo.
echo ============================================================
echo  VIRADA CONCLUIDA
echo ============================================================
echo.
echo  Agora:
echo   1. Teste a primeira gravacao em producao pela sua copia local
echo   2. Peca a equipe para abrir o modelo da rede e responder SIM
echo   3. Rode o backup uma vez: schtasks /Run /TN "CRM Zapromaq - backup 0900"
echo   4. Mova a planilha 3.6 para uma pasta de consulta, fora do uso diario
echo.
goto :fim
:erro
echo.
echo ############################################################
echo  INTERROMPIDO - confira o log da etapa acima.
echo  O banco de producao pode ter ficado incompleto: apague-o antes
echo  de rodar de novo.
echo ############################################################
:fim
popd
echo.
pause
