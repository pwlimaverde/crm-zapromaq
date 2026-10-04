@echo off
setlocal
rem CHECAR-AMBIENTE.bat - confere se esta maquina tem o que o CRM precisa (nao altera nada).
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0checar-ambiente.ps1"
echo.
pause
