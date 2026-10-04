# Como o INICIAR-CRM sabe que há versão nova

O `INICIAR-CRM.bat` das estações não faz parte deste projeto, mas depende de uma
informação que o build publica. Existem **duas** formas de ler a versão
publicada, e as duas continuam funcionando:

| Fonte | Onde | Custo para ler | Observação |
|---|---|---|---|
| `VERSAO-FRONT.txt` | raiz de `BASE`, ao lado do `.xlsm` | instantâneo (é texto) | **recomendado** |
| `Início!B7` | dentro do `.xlsm` | abrir o Excel (segundos) ou descompactar o `.xlsm` | continua no mesmo formato e na mesma célula |

O `VERSAO-FRONT.txt` é gravado **depois** de o `.xlsm` já estar publicado na
raiz, e só quando a publicação deu certo. Então a estação nunca vê um número
novo apontando para um arquivo velho. O arquivo é ANSI com CRLF, porque quem lê
é o `cmd.exe`:

```
1.5
Ambiente: PRODUCAO   |   Versao 1.5   |   Publicada em 20/09/2026 10:21
arquivo=CRM_Zapromaq.xlsm
publicado=2026-09-20 10:21:45
por=ESTACAO01\paulo
```

A **primeira linha** é a versão, e é só ela que o `.bat` precisa comparar.

## Trecho para o INICIAR-CRM.bat

Lê a versão publicada, compara com a do registro do usuário (`HKCU`, não precisa
de administrador) e só copia quando mudou:

```bat
set "REDE=\\servidor\...\BASE"
set "LOCAL=%USERPROFILE%\Documents\CRM Zapromaq"
set "CHAVE=HKCU\Software\Zapromaq\CRM"

rem --- versão publicada (1a linha do VERSAO-FRONT.txt)
set "VERPUB="
if exist "%REDE%\VERSAO-FRONT.txt" set /p VERPUB=<"%REDE%\VERSAO-FRONT.txt"
if not defined VERPUB goto :copiar          & rem sem o arquivo: copia por garantia

rem --- versão que esta máquina já tem
set "VERLOC="
for /f "tokens=2,*" %%a in ('reg query "%CHAVE%" /v Versao 2^>nul ^| find "Versao"') do set "VERLOC=%%b"

if /i "%VERPUB%"=="%VERLOC%" goto :abrir    & rem igual: nem copia

:copiar
if not exist "%LOCAL%" mkdir "%LOCAL%"
copy /y "%REDE%\CRM_Zapromaq.xlsm" "%LOCAL%\CRM_Zapromaq.xlsm" >nul
if errorlevel 1 goto :erro
reg add "%CHAVE%" /v Versao /t REG_SZ /d "%VERPUB%" /f >nul

:abrir
start "" "%LOCAL%\CRM_Zapromaq.xlsm"
exit /b 0

:erro
echo Nao foi possivel copiar o CRM da rede. Verifique a conexao e tente de novo.
pause
exit /b 1
```

Detalhes que evitam dor de cabeça:

- **Copiar antes de gravar o registro.** Se a cópia falhar, a versão no registro
  continua a antiga e a próxima tentativa copia de novo.
- **Sem o `VERSAO-FRONT.txt`, copie assim mesmo.** É o que acontece na primeira
  vez e se alguém apagar o arquivo: o pior caso é uma cópia a mais.
- **A cópia local aberta não atrapalha a publicação.** O que não pode estar
  aberto na hora de publicar é o `.xlsm` da rede.
- A versão também fica em `config.versao_front`, no banco. O próprio CRM compara
  na abertura e avisa quem estiver com cópia antiga — é a rede de segurança caso
  o atalho não tenha rodado.
