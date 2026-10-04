@echo off
setlocal
rem Cria BASE\crm_zapromaq.accdb com dados FICTICIOS para testar o MCP.
rem Exige o motor ACE 64 bits (Office 64 bits ou Access Database Engine 2016 x64).
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0criar-banco-teste.ps1" -Recriar
echo.
pause
