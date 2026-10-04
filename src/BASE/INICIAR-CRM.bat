@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
title CRM Comercial Zapromaq

rem ============================================================
rem  INICIAR-CRM.bat - lancador do CRM Comercial Zapromaq
rem
rem  ARQUIVO UNICO: vive SO na pasta da rede. Na estacao fica um
rem  ATALHO apontando para ele.
rem
rem  Ordem:
rem    1. acha a pasta da rede
rem    2. acha a pasta Documentos deste usuario (pelo registro)
rem    3. cria o atalho, se ainda nao existe
rem    4. le a versao publicada em VERSAO-FRONT.txt (1a linha)
rem       e compara com HKCU\Software\Zapromaq\CRM\VersaoInstaladaCRM
rem    5. se precisar, confere se o CRM local esta aberto, copia
rem       (via arquivo temporario) e SO ENTAO grava o registro
rem    6. abre a copia local
rem
rem  Nao abre Excel para descobrir versao. Sem rede, abre a copia
rem  local com aviso.
rem ============================================================

pushd "%~dp0" 2>nul

set "ARQ=CRM_Zapromaq.xlsm"
set "ICO=crm_zapromaq.ico"
set "ARQVER=VERSAO-FRONT.txt"
set "REGKEY=HKCU\Software\Zapromaq\CRM"
set "REGVAL=VersaoInstaladaCRM"
set "SUB=01 - CRM\01 - CONTROLE\BASE"
set "REDE="

rem --- 1. onde esta a rede -------------------------------------
call :tentar "%~dp0"
call :tentar "%~dp0..\..\.."
call :tentar "G:\1 - COMERCIAL ZAPROMAQ"
call :tentar "\\servidor\COMERCIAL\1 - COMERCIAL ZAPROMAQ"

rem --- 2. pasta Documentos deste usuario ------------------------
set "DOCS="
for /f "tokens=2*" %%A in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" /v Personal 2^>nul') do set "DOCS=%%B"
if not defined DOCS set "DOCS=%USERPROFILE%\Documents"
set "LOCAL=%DOCS%\CRM Zapromaq"
if not exist "%LOCAL%" mkdir "%LOCAL%" 2>nul

echo ============================================================
echo  CRM COMERCIAL ZAPROMAQ
echo ============================================================
echo.
echo  Copia local: %LOCAL%
if defined REDE echo  Rede ......: %REDE%

rem --- 3. atalho -----------------------------------------------
if defined REDE call :atalho

rem --- 4. versoes ----------------------------------------------
if not defined REDE goto :semrede

set "VERSAO_REDE="
if exist "%REDE%\%ARQVER%" (
    set /p "VERSAO_REDE=" < "%REDE%\%ARQVER%"
)
rem limpa espacos e CR eventuais
if defined VERSAO_REDE for /f "tokens=1 delims= " %%V in ("!VERSAO_REDE!") do set "VERSAO_REDE=%%V"

set "VERSAO_LOCAL="
for /f "tokens=2,*" %%A in ('reg query "%REGKEY%" /v %REGVAL% 2^>nul ^| find "%REGVAL%"') do set "VERSAO_LOCAL=%%B"

echo.
echo  Versao publicada: !VERSAO_REDE!
echo  Versao instalada: !VERSAO_LOCAL!

if not defined VERSAO_REDE (
    echo.
    echo  AVISO: nao consegui ler %ARQVER% na rede.
    echo  A copia local NAO sera substituida.
    goto :abrir
)

rem Sem copia local ou sem registro: instala a publicada.
if not exist "%LOCAL%\%ARQ%" goto :atualizar
if not defined VERSAO_LOCAL goto :atualizar

call :compararstrings "!VERSAO_REDE!" "!VERSAO_LOCAL!"
set "VC=!errorlevel!"

if "!VC!"=="0" (
    echo  Sua copia ja esta atualizada.
    goto :abrir
)
if "!VC!"=="2" (
    echo  A copia local tem versao mais nova que a da rede. Nada sera substituido.
    goto :abrir
)
if "!VC!"=="9" (
    echo  AVISO: formato de versao invalido. A copia local NAO sera substituida.
    goto :abrir
)

rem --- 5. atualizar --------------------------------------------
:atualizar
echo.
echo  Instalando a versao !VERSAO_REDE!...

:checatrava
call :emuso "%LOCAL%\%ARQ%"
if errorlevel 1 (
    echo.
    echo  O CRM esta ABERTO nesta maquina e nao pode ser substituido.
    echo  Feche o Excel (salve o que precisar^) e pressione uma tecla.
    echo  Para abrir sem atualizar, feche esta janela.
    pause >nul
    goto :checatrava
)

set "TMPARQ=%LOCAL%\~novo_%ARQ%"
copy /y "%REDE%\%ARQ%" "%TMPARQ%" >nul 2>nul
if errorlevel 1 goto :falhacopia
if not exist "%TMPARQ%" goto :falhacopia
move /y "%TMPARQ%" "%LOCAL%\%ARQ%" >nul 2>nul
if errorlevel 1 goto :falhacopia

reg add "%REGKEY%" /v %REGVAL% /t REG_SZ /d "!VERSAO_REDE!" /f >nul 2>nul
if errorlevel 1 (
    echo  AVISO: arquivo atualizado, mas nao consegui gravar o registro.
) else (
    echo  Versao !VERSAO_REDE! instalada e registrada.
)
goto :abrir

:semrede
echo.
echo  NAO ACHEI A PASTA DA REDE.
echo  Abrindo a copia desta maquina - pode estar desatualizada.
echo.
goto :abrir

:falhacopia
del "%TMPARQ%" >nul 2>nul
echo.
echo  NAO CONSEGUI ATUALIZAR. Confira o acesso a rede e se o
echo  Excel esta fechado, e rode de novo.
echo  Abrindo a copia atual assim mesmo. O registro NAO foi alterado.
echo.
timeout /t 5 >nul

:abrir
if not exist "%LOCAL%\%ARQ%" goto :naotem
start "" "%LOCAL%\%ARQ%"
popd 2>nul
exit /b 0

:naotem
echo.
echo  Nao existe copia do CRM nesta maquina e a rede nao respondeu.
echo  Confira se a pasta da rede esta acessivel e rode de novo.
echo.
pause
popd 2>nul
exit /b 1

rem ------------------------------------------------------------
rem  EM USO: errorlevel 1 se o arquivo existe e esta travado
rem  (aberto no Excel). Abre para acrescimo sem gravar nada.
rem ------------------------------------------------------------
:emuso
if not exist "%~1" exit /b 0
2>nul (>>"%~1" call ) && exit /b 0
exit /b 1

rem ------------------------------------------------------------
rem  ATALHO
rem ------------------------------------------------------------
:atalho
if not exist "%LOCAL%\%ICO%" (
    if exist "%REDE%\%ICO%" (
        copy /y "%REDE%\%ICO%" "%LOCAL%\%ICO%" >nul 2>nul
    ) else if exist "%REDE%\v1.4\%ICO%" (
        copy /y "%REDE%\v1.4\%ICO%" "%LOCAL%\%ICO%" >nul 2>nul
    )
)
if exist "%LOCAL%\CRM Zapromaq.lnk" goto :eof

set "VBS=%TEMP%\crm_atalho_%RANDOM%.vbs"
> "%VBS%" echo Set s = CreateObject("WScript.Shell")
>>"%VBS%" echo Set a = s.CreateShortcut("%LOCAL%\CRM Zapromaq.lnk")
>>"%VBS%" echo a.TargetPath = "%REDE%\INICIAR-CRM.bat"
>>"%VBS%" echo a.WorkingDirectory = "%LOCAL%"
>>"%VBS%" echo a.IconLocation = "%LOCAL%\%ICO%"
>>"%VBS%" echo a.WindowStyle = 7
>>"%VBS%" echo a.Description = "Abre o CRM Comercial Zapromaq, atualizando da rede"
>>"%VBS%" echo a.Save
cscript //nologo "%VBS%" >nul 2>nul
del "%VBS%" >nul 2>nul

if exist "%LOCAL%\CRM Zapromaq.lnk" (
    echo.
    echo  Atalho "CRM Zapromaq" criado em:
    echo    %LOCAL%
) else (
    echo.
    echo  AVISO: nao consegui criar o atalho.
)
goto :eof

rem ------------------------------------------------------------
rem  COMPARAR STRINGS DE VERSAO  (%1 rede, %2 local)
rem  0 iguais, 1 rede mais nova, 2 local mais nova, 9 invalida
rem ------------------------------------------------------------
:compararstrings
set "VA=%~1"
set "VB=%~2"
if not defined VA exit /b 9
if not defined VB exit /b 9
for /f "delims=0123456789." %%Z in ("%VA%") do exit /b 9
for /f "delims=0123456789." %%Z in ("%VB%") do exit /b 9
for /f "tokens=1-4 delims=." %%a in ("%VA%.0.0.0") do (
    set /a A1=1%%a-100, A2=1%%b-100, A3=1%%c-100, A4=1%%d-100
)
for /f "tokens=1-4 delims=." %%a in ("%VB%.0.0.0") do (
    set /a B1=1%%a-100, B2=1%%b-100, B3=1%%c-100, B4=1%%d-100
)
for %%N in (1 2 3 4) do (
    if !A%%N! GTR !B%%N! exit /b 1
    if !A%%N! LSS !B%%N! exit /b 2
)
exit /b 0

rem --- procura a pasta BASE a partir de uma raiz ---------------
:tentar
if defined REDE goto :eof
set "TENTA=%~1\%SUB%"
if exist "%TENTA%\%ARQ%" set "REDE=%TENTA%"
goto :eof
