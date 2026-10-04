@echo off
setlocal
rem Registra o servidor MCP do CRM no Claude Desktop deste usuario.
rem Nao precisa de administrador. Para remover: INSTALAR-MCP.bat -Remover
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0mcp\instalar-mcp.ps1" %*
echo.
pause
