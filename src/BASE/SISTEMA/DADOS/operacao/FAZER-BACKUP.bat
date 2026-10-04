@echo off
setlocal EnableExtensions
rem FAZER-BACKUP.bat - backup manual do banco do CRM.
rem Tenta o backup COMPLETO (exige ninguem conectado); se houver alguem,
rem oferece a copia simples. O resultado e conferido pelo arquivo gerado.
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
set "PASTA=%~dp0..\execucao\backup"
echo ============================================================
echo  BACKUP MANUAL - CRM Zapromaq
echo ============================================================
echo.
pause
call :contar ANTES
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\backup-banco.ps1" -Modo Diario
set "RC=%errorlevel%"
call :contar DEPOIS
if %DEPOIS% GTR %ANTES% (echo. & echo  BACKUP COMPLETO GERADO - serve para restaurar. & goto :fim)
if not "%RC%"=="2" goto :erro
echo.
echo  Ha alguem com o CRM aberto: o backup completo nao foi feito.
echo  A copia simples cobre o caso comum, mas pode sair inconsistente
echo  se for tirada no meio de uma gravacao.
set "R="
set /p "R=Gerar a copia simples? (S/N): "
if /i not "%R%"=="S" goto :fim
call :contar ANTES
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\backup-banco.ps1" -Modo Periodico
set "RC=%errorlevel%"
call :contar DEPOIS
if %DEPOIS% GTR %ANTES% (echo. & echo  COPIA SIMPLES GERADA. & goto :fim)
:erro
echo.
echo  NENHUMA COPIA NOVA FOI GERADA (codigo %RC%). Veja execucao\backup\backup.log
goto :fim
:contar
set "_n=0"
for %%A in ("%PASTA%\*.accdb") do set /a _n+=1
set "%~1=%_n%"
goto :eof
:fim
echo.
pause
