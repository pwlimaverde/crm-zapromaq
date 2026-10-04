@echo off
setlocal
rem APLICAR-MIGRACOES.bat - aplica as migracoes pendentes do banco (faz backup antes).
rem Migracao estrutural exige todos fora do CRM. Para so listar: APLICAR-MIGRACOES.bat -Simular
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0aplicar-migracoes.ps1" %*
echo.
pause
