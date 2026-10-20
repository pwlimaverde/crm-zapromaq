@echo off
setlocal
rem Publica o conector de dados numa pasta de versao (IA\conector\versoes) e troca o ponteiro (DEV-05).
rem Voltar para uma versao anterior: PUBLICAR-CONECTOR.bat -Voltar 1.6.0 -Motivo "por que"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0mcp\publicar-conector.ps1" %*
pause
