@echo off
setlocal EnableExtensions
rem MEDIR-COMPACTACAO.bat - mede, numa COPIA, quanto compactar o banco ganharia (DEV-25).
rem Nunca toca o banco de origem; o relatorio fica em execucao\medicoes.
rem Uso: MEDIR-COMPACTACAO.bat [-Banco <copia de backup validada>] [-Repeticoes 7]
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\medir-compactacao.ps1" %*
set "RC=%errorlevel%"
pause
exit /b %RC%
