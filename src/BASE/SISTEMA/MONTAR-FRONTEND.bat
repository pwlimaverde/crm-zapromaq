@echo off
setlocal EnableExtensions
rem ============================================================
rem  MONTAR-FRONTEND.bat - unico comando da publicacao do CRM.
rem
rem  Faz a sequencia inteira, e so publica se tudo der certo:
rem    1. confere o ambiente (Excel, banco, .xlsm livre)
rem    2. BACKUP do banco e do .xlsm atuais
rem    3. MIGRACOES pendentes do banco
rem    4. monta o .xlsm novo, compila e roda o autoteste
rem    5. confere o pacote e roda o teste de uso (grades e Painel)
rem    6. publica na raiz de BASE e sobe a versao (1.4 -> 1.5)
rem
rem  Falhou em qualquer etapa: NADA e publicado; o .xlsm da raiz
rem  continua o de antes e o backup da etapa 2 fica guardado.
rem
rem  Opcoes:  MONTAR-FRONTEND.bat -Teste          so monta e confere
rem           MONTAR-FRONTEND.bat -SemMigracoes   pula as migracoes
rem           MONTAR-FRONTEND.bat -SemTesteDeUso  pula o teste de uso
rem ============================================================
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "PS=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
echo ============================================================
echo  PUBLICACAO DO CRM ZAPROMAQ
echo ============================================================
echo.
echo  Esta maquina precisa ter o Excel 2019 64 bits.
echo.
echo  Antes de continuar:
echo    - feche o CRM_Zapromaq.xlsm da pasta da rede em todas as maquinas
echo      (as copias locais em Documentos podem ficar abertas);
echo    - de preferencia, ninguem gravando no sistema agora.
echo.
echo  O que vai acontecer: backup, migracoes do banco, montagem,
echo  compilacao, autoteste, teste de uso e publicacao da versao nova.
echo.
pause
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0DADOS\build\montar-frontend.ps1" %*
set "RC=%errorlevel%"
echo.
if "%RC%"=="0" (
  echo  CONCLUIDO. A versao nova esta publicada na raiz de BASE.
  echo.
  echo  Proximo passo: copie o CRM_Zapromaq.xlsm para a sua maquina,
  echo  abra e confira. Estando tudo certo, avise a equipe para abrir
  echo  o atalho INICIAR-CRM - ele atualiza sozinho a copia de cada um.
) else (
  echo  FALHOU - veja o log indicado acima. NADA foi publicado:
  echo  o CRM_Zapromaq.xlsm da raiz continua sendo o de antes.
)
echo.
pause
exit /b %RC%
