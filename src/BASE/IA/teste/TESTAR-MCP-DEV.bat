@echo off
setlocal
rem Testa o MCP de desenvolvimento numa COPIA de SISTEMA\DADOS (nada do projeto muda).
rem Sem Edge na maquina: TESTAR-MCP-DEV.bat -SemPrevia
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0testar-mcp-dev.ps1" %*
echo.
pause
