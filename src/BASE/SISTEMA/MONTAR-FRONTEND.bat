@echo off
setlocal EnableExtensions
rem ============================================================
rem  MONTAR-FRONTEND.bat - unico comando da publicacao do CRM.
rem
rem  Faz a sequencia inteira, na ordem, e so publica se tudo der certo:
rem    1 pacote de municipios   2 ambiente   3 trava (.laccdb)   4 esquema
rem    5 backup (H0)   6 copia de trabalho   7 migracoes NA COPIA
rem    8 conferencias da copia   9 monta o front e roda o autoteste
rem    10 teste de uso   11 troca do banco (so com trava livre e H0 igual)
rem    12 publica na raiz de BASE e sobe a versao
rem
rem  Falhou antes do passo 11: NADA mudou (banco de producao e raiz).
rem  Falhou no 12: o banco volta ao backup.
rem
rem  Opcoes:  MONTAR-FRONTEND.bat -Ensaio         passos 1 a 8 numa copia; nao troca o banco, nao precisa de Excel
rem           MONTAR-FRONTEND.bat -Teste          so sobre banco TESTE; nao publica
rem           MONTAR-FRONTEND.bat -SemTesteDeUso  pula o teste de uso
rem           MONTAR-FRONTEND.bat -Volta -Motivo "por que"   VOLTA: publica o pacote anterior
rem                (extraido por cima da BASE) como versao NOVA, maior; motivo obrigatorio
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
if "%RC%"=="0" if "%~1"=="-Ensaio" (
  echo  ENSAIO APROVADO. Nada foi publicado e o banco de producao nao foi tocado.
  echo  Leia o relatorio indicado acima ^(DADOS\execucao\ensaio^).
  goto :fim
)
if "%RC%"=="0" if not "%~1"=="-Teste" (
  echo  CONCLUIDO. A versao nova esta publicada na raiz de BASE.
  echo.
  echo  Proximo passo: copie o CRM_Zapromaq.xlsm para a sua maquina,
  echo  abra e confira. Estando tudo certo, avise a equipe para abrir
  echo  o atalho INICIAR-CRM - ele atualiza sozinho a copia de cada um.
) else (
if "%RC%"=="0" if "%~1"=="-Teste" (
  echo  TESTE CONCLUIDO. Nada foi publicado; o arquivo de teste esta em DADOS\execucao\teste.
) else if "%RC%"=="0" (
  echo  CONCLUIDO. A versao nova esta publicada na raiz de BASE.
  echo.
  echo  Proximo passo: copie o CRM_Zapromaq.xlsm para a sua maquina,
  echo  abra e confira. Estando tudo certo, avise a equipe para abrir
  echo  o atalho INICIAR-CRM - ele atualiza sozinho a copia de cada um.
) else (
  echo  FALHOU - veja o log indicado acima. NADA foi publicado:
  echo  o CRM_Zapromaq.xlsm da raiz continua sendo o de antes.
)
:fim
echo.
pause
exit /b %RC%
