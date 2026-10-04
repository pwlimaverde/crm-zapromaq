@echo off
setlocal
rem AGENDAR-BACKUPS.bat - cria as tarefas de backup (09:00 e 16:40) para ESTE usuario.
rem Nao precisa de administrador. Para remover: AGENDAR-BACKUPS.bat -Remover
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\agendar-backups.ps1" %*
echo.
pause
