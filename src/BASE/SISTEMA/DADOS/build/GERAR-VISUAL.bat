@echo off
setlocal
rem GERAR-VISUAL.bat - gera fundos, previas e modTema a partir de fonte\layout (PowerShell + Edge).
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0gerar-visual.ps1" %*
echo.
pause
