@echo off
setlocal
rem Configura MCP local e skill para Claude Desktop e Codex.
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0mcp\instalar-mcp.ps1" -Aplicativo Ambos %*
if errorlevel 1 (echo FALHA na instalacao.) else (echo Configuracao concluida. Reabra os aplicativos.)
pause
