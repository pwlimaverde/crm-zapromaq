@echo off
setlocal
rem TESTAR-RESTAURACAO.bat - restaura o backup mais recente numa COPIA e confere.
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\restaurar-banco.ps1" %*
echo.
pause
