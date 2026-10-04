@echo off
setlocal
rem Registra o MCP de DESENVOLVIMENTO do front (altera SISTEMA\DADOS\fonte).
rem Instalar so na maquina de quem mantem o sistema. Remover: INSTALAR-MCP-DEV.bat -Remover
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0mcp\instalar-mcp.ps1" -Dev %*
echo.
pause
