@echo off
setlocal
rem Conversa com o servidor MCP como o Claude Desktop faria e confere as respostas.
rem Grava no banco de teste: rode CRIAR-BANCO-TESTE.bat antes.
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0testar-mcp.ps1" %*
echo.
pause
