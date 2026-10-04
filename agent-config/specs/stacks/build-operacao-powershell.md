# Stack: build e operação — Windows PowerShell 5.1 / `.bat`

## Papel

Montar e publicar o `.xlsm` na rede (`MONTAR-FRONTEND.bat`), operar o banco (backup,
restauração, agendamento) e atualizar as estações (`INICIAR-CRM.bat`). Tudo roda numa estação
só com Windows + Office: nada de módulo extra, Python ou Node.

## Versões e ambiente

- **Windows PowerShell 5.1** (`powershell.exe`, chamado pelo caminho completo, 64 bits —
  `Sysnative` quando o `.bat` roda em 32). Nunca `pwsh`.
- Automação COM do Excel; "Confiar no acesso ao modelo de projeto do VBA" ligado só durante a
  montagem (HKCU) e restaurado no fim.
- Agendador: `schtasks.exe /Create /XML` (o `Register-ScheduledTask` dá "Acesso negado" para
  usuário comum na rede da empresa).

## Onde fica

```
src/BASE/INICIAR-CRM.bat                     lançador (fica na rede; estação tem atalho)
src/BASE/SISTEMA/MONTAR-FRONTEND.bat         único ponto de entrada do build
src/BASE/SISTEMA/DADOS/
  VERSAO.txt                                 versão de partida (o build sobe e grava ao publicar)
  build/montar-frontend.ps1, gerar-visual.ps1, lib/{Excel,Pacote,Versao}.ps1
  lib/Comum.ps1                              raiz, caminhos, banco, log
  operacao/{FAZER-BACKUP,AGENDAR-BACKUPS,TESTAR-RESTAURACAO}.bat + scripts/
  testes/                                    testar-build-lib, testar-uso, checar-ambiente...
  execucao/                                  GERADO na rede: logs, backups, histórico, build (fora do git)
```

## O build (`montar-frontend.ps1`)

pré-checagens → **backup** (`execucao\backup\antes-da-publicacao`) → **migrações** pendentes →
visual em dia → versão (parte da **maior** entre `VERSAO.txt` e `VERSAO-FRONT.txt`, +1) → monta em
arquivo temporário → compila + `modAutoteste` com vigia → confere o pacote pelo zip → teste de
uso → publica (anterior para `execucao\historico-front`, grava `VERSAO.txt`, `VERSAO-FRONT.txt`,
`config.versao_front`, `CHANGELOG.md`). Falhou → nada publicado. Sem o banco na raiz, não publica.
`-Teste` monta em `execucao\teste` sem publicar nem subir versão.

## Contratos

- `Início!B7`: `Ambiente: X   |   Versao X.Y   |   Publicada em dd/MM/yyyy HH:mm` — padrão e célula fixos.
- `VERSAO-FRONT.txt` na raiz de `BASE`: 1ª linha = versão; o `INICIAR-CRM.bat` compara com
  `HKCU\Software\Zapromaq\CRM\VersaoInstaladaCRM` e copia o `.xlsm` se a rede for mais nova.
- O `INICIAR-CRM.bat` acha a rede pela própria pasta (`..\..\..` + `01 - CRM\01 - CONTROLE\BASE`),
  depois `G:\1 - COMERCIAL ZAPROMAQ`, depois o UNC.

## Comandos

```powershell
src\BASE\SISTEMA\MONTAR-FRONTEND.bat [-Teste] [-SemMigracoes] [-SemTesteDeUso]
powershell -File src\BASE\SISTEMA\DADOS\testes\testar-build-lib.ps1
src\BASE\SISTEMA\DADOS\testes\CHECAR-AMBIENTE.bat
```

## Estilo

- `.ps1` em **UTF-8 com BOM + CRLF** (sem BOM o 5.1 lê como ANSI). `.bat` em ASCII + CRLF.
- `$ErrorActionPreference = 'Stop'` por padrão; onde se chama executável nativo que escreve no
  stderr (`git`, `schtasks`), `Continue` local e decisão pelo `$LASTEXITCODE`.
- Caminhos com `Join-Path`; `-LiteralPath`; `ProviderPath` em caminho de rede; nada de letra de
  unidade fixa. Arquivo de texto gravado com encoding explícito (`UTF8Encoding($false)`, cp1252 para VBA).
- Log em `execucao\logs` via `New-Log`; mensagens ao usuário em pt-BR sem acento nos `.bat`.

## Testes

`testar-build-lib.ps1` (`Conferir nome, condição`) para funções puras do build (versão,
B7, pacote). O build completo é o teste de integração (`verificar-tudo -Completo`). Mudança no
`INICIAR-CRM.bat` ou nos agendamentos: teste manual 🖥 numa estação, registrado no plano.

## Limites

- Sempre: falhar sem publicar; backup antes de mexer no banco ou no `.xlsm` publicado.
- Perguntar antes: mudar contrato (B7, `VERSAO-FRONT.txt`, nomes fixos de arquivo), o
  `INICIAR-CRM.bat`, os agendamentos de backup.
- Nunca: exigir administrador; instalar algo na estação; caminho absoluto da máquina de
  desenvolvimento; publicar a partir de `-Teste`.

## Revisão

- [ ] BOM/CRLF; sintaxe (o `verificar.py` passa no analisador do 5.1).
- [ ] Nenhum `Write-Host` onde a saída é capturada por outro script; códigos de saída corretos.
- [ ] Caminhos com espaço entre aspas no `.bat`.
- [ ] Falha no meio não deixa a rede pela metade.

## Referências offline

`referencias/` 09-powershell-5.1, 04-excel-modelo-objetos, 03-vbe-modelo-extensibilidade.
