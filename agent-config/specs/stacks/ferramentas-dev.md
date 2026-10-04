# Stack: ferramentas de desenvolvimento

Só na máquina de desenvolvimento — **nada disto pode ser pré-requisito** para montar ou usar
o sistema na rede.

## O que há

| Ferramenta | Para | Onde |
|---|---|---|
| Python 3 (3.14 aqui) | `verificar.py` (estático), `conferir-*.py` | `src/BASE/SISTEMA/DADOS/ferramentas/verificacao/` |
| ANTLR 4.13 (jar + gramática MS-VBAL) | parser formal do VBA | `.../verificacao/antlr-vba/` (gerado commitado; regerar só se mudar a gramática) |
| git + git-flow (GitFlow .NET 2.3.0) | versionamento | finalizar com `agent-config/ferramentas/finalizar-branch.ps1` |
| agent-skills (Addy Osmani, MIT) | processo spec → ship | `agent-config/vendor/agent-skills` (subtree) + `adaptacao/` → `.claude/` |
| Claude Code | desenvolvimento | aberto em `agent-config/` |
| Excel 365 + ACE 16 | build `-Teste`, testes com banco | nesta máquina |

## Scripts de `agent-config/ferramentas`

```powershell
verificar-tudo.ps1 [-Rapido | -Completo]   # a suíte (o /test); prova de "pronto"
instalar-skills.ps1 [-Conferir]            # regera .claude a partir de vendor + adaptacao
atualizar-agent-skills.ps1                 # git subtree pull do original + reinstala
finalizar-branch.ps1 [-Branch x] [-Enviar] # finish do git-flow (--no-ff, tag vX.Y em release/hotfix)
empacotar-base.ps1                         # dist\BASE-vX.Y-....zip para a empresa (do commit)
```

## Estilo

Scripts de desenvolvimento em PowerShell 5.1 (UTF-8 com BOM + CRLF) ou Python; comentário de
cabeçalho com uso; saída em pt-BR; código de saída 0/1. Python em UTF-8, só biblioteca padrão
(+ o runtime do ANTLR que já está no projeto).

## Limites

- Ferramenta nova de desenvolvimento: perguntar antes; nunca entra em `src/BASE` como dependência.
- `vendor/` não se edita; `.claude/skills|commands|agents|references` não se editam (gerados).
- Nada de credencial em script ou config versionada.

## Revisão

- [ ] Funciona a partir de qualquer pasta do repositório (usa `git rev-parse` / `$PSScriptRoot`).
- [ ] Não altera nada fora do repositório sem dizer; falha com mensagem clara.
