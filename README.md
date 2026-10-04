# CRM Zapromaq

CRM comercial da **Zapromaq** (vendas de máquinas industriais): cadastro de clientes,
contatos e oportunidades, painel de metas e acompanhamento de atendimentos.

O sistema roda **100% local, sem internet**: um banco Access (`.accdb`) numa pasta
compartilhada da rede e um front em Excel (`.xlsm` com VBA) que cada estação copia para
a própria máquina. Este repositório é a modernização do sistema legado — fonte VBA
organizada, build automatizado, layout visual desenhado em HTML e executado em MSForms,
migrações de banco versionadas e um servidor MCP para o Claude Desktop consultar e
atualizar os dados.

## Estrutura do repositório

```
crm-zapromaq/
├── agent-config/          configuração dos agentes de IA (não vai para a rede)
│   ├── CLAUDE.md          instruções do projeto; o Claude Code é aberto nesta pasta
│   │                      e trabalha sobre a pasta acima
│   ├── DIRETRIZES-IA.md   protocolo de trabalho: /spec → /plan → /build → /test → /review → /ship
│   ├── specs/             spec do projeto, das 6 stacks e de cada funcionalidade
│   ├── vendor/agent-skills/  addyosmani/agent-skills na íntegra (git subtree, MIT)
│   ├── adaptacao/         o que se soma ao original para este projeto
│   ├── .claude/           skills, comandos e agentes do Claude Code (GERADO)
│   ├── ferramentas/       verificar-tudo, instalar-skills, atualizar-agent-skills,
│   │                      empacotar-base, finalizar-branch
│   └── doc_dev/planejamento/
│       ├── PLANO.md       checklist de fases, decisões, correções e pendências
│       └── referencias/   documentação oficial offline (VBA, MSForms, Excel, ADO, ACE…)
└── src/
    └── BASE/              ENTREGÁVEL — a pasta copiada inteira para a rede
        ├── INICIAR-CRM.bat           lançador das estações (atalho aponta para ele na rede)
        ├── SISTEMA/
        │   ├── MONTAR-FRONTEND.bat   único ponto de entrada do build
        │   ├── LEIA-ME-PRIMEIRO.md   roteiro de publicação
        │   └── DADOS/                fonte VBA, layout, build, banco, operação, testes
        └── IA/                       servidores MCP (dados e desenvolvimento) e skills
```

> **`src/BASE` tem estrutura fixa.** Os scripts usam caminhos relativos a ela e as
> estações dependem do nome e da posição dos arquivos. Não renomeie nem mova nada lá
> dentro. Na rede, o banco `crm_zapromaq.accdb` e o `CRM_Zapromaq.xlsm` montado ficam na
> raiz de `BASE` — ambos fora do git.

## Restrições de implantação

- Sem internet, CDN, API externa, pacote instalado ou OCX/DLL de terceiros.
- Estações: Windows + **Office 2019 64 bits** (nada de recursos do Microsoft 365).
- Build e operação só com o que vem no Windows: **PowerShell 5.1**, automação COM do
  Excel e Edge headless (prévia visual).
- Sem caminho absoluto nem letra de unidade: copiar a pasta `BASE` para a rede basta.
- Python, git e linters são ferramentas **só de desenvolvimento**.

## Stack

| Camada            | Tecnologia                                                        |
|-------------------|-------------------------------------------------------------------|
| Banco             | Access `.accdb` (UNC), `Microsoft.ACE.OLEDB.16.0`                 |
| Front             | Excel `.xlsm`, VBA + ADO (late binding), UserForms MSForms        |
| Layout visual     | JSON → prévia HTML/JS no Edge → fundos BMP + `modTema.bas`        |
| Build / operação  | PowerShell 5.1 disparado por `.bat`                               |
| IA                | Servidor MCP em PowerShell (stdio) para o Claude Desktop          |
| Verificação (dev) | Python: parser VBA (ANTLR MS-VBAL), conferência de layout e DDL   |

## Comandos

Executados a partir de `src/BASE` com `powershell.exe` (5.1).

```powershell
# Verificação estática completa (VBA, layout x código, esquema x DDL, .ps1, .bat, JSON) — dev
python SISTEMA\DADOS\ferramentas\verificacao\verificar.py

# Prévia visual, fundos e módulo de tema
powershell -File SISTEMA\DADOS\build\gerar-visual.ps1 [-Tela frmCRM]

# Build do .xlsm (exige Excel 2019); -Teste monta sem publicar nem subir versão
SISTEMA\MONTAR-FRONTEND.bat [-Teste]

# Testes
powershell -File SISTEMA\DADOS\testes\testar-build-lib.ps1
IA\teste\TESTAR-MCP.bat -SoProtocolo          # MCP sem banco
IA\teste\CRIAR-BANCO-TESTE.bat ; IA\teste\TESTAR-MCP.bat
IA\teste\TESTAR-MCP-DEV.bat

# Banco
SISTEMA\DADOS\banco\APLICAR-MIGRACOES.bat
powershell -File SISTEMA\DADOS\testes\criar-banco-demo.ps1 [-Recriar]   # banco fictício: 20 clientes, contatos e atendimentos
```

Para experimentar sem a rede da empresa: gere o banco de demonstração e monte com
`SISTEMA\MONTAR-FRONTEND.bat -Teste`. O `.xlsm` sai em
`SISTEMA\DADOS\execucao\teste\CRM_Zapromaq.xlsm`, já apontando para o banco local.

A máquina de desenvolvimento não precisa de Excel: o VBA é conferido estaticamente
pelo `verificar.py` e pela prévia; compilação e autoteste acontecem no build, numa
estação com Excel 2019.

## Publicação

Roteiro completo em [`src/BASE/SISTEMA/LEIA-ME-PRIMEIRO.md`](src/BASE/SISTEMA/LEIA-ME-PRIMEIRO.md)
e [`src/BASE/SISTEMA/DADOS/docs/implantacao.md`](src/BASE/SISTEMA/DADOS/docs/implantacao.md).
A BASE vai para a empresa como um pacote:

```powershell
powershell -File agent-config\ferramentas\empacotar-base.ps1   # -> dist\BASE-vX.Y-....zip
```

O pacote sai do commit e nunca leva banco, planilha, logs nem backups. Na empresa:

1. **Extraia o `.zip` por cima da pasta `BASE` da rede**, substituindo os arquivos.
   **Não apague a `BASE` antes**: o banco, a planilha publicada, o ícone e
   `SISTEMA\DADOS\execucao` (backups, histórico) só existem lá.
2. Rode `SISTEMA\MONTAR-FRONTEND.bat`. Ele faz backup, aplica as migrações do banco,
   monta, compila, testa e só então publica a versão seguinte. Sem o banco na raiz,
   não publica; a numeração nunca repete uma versão já publicada.

**A versão da rede tem prioridade.** Ajustes feitos lá pelo Claude Desktop (MCP de
desenvolvimento) voltam para o repositório **antes** de qualquer ciclo novo: a BASE
da rede é trazida inteira e comparada com o último commit. Roteiro em
[`levar-para-a-empresa.md`](src/BASE/SISTEMA/DADOS/docs/levar-para-a-empresa.md) e
[`sincronizar-rede.md`](src/BASE/SISTEMA/DADOS/docs/sincronizar-rede.md).

## Desenvolvimento com IA

Abra o Claude Code dentro de `agent-config/`. O `CLAUDE.md` descreve arquitetura,
restrições e regras aprendidas na produção; o código manipulado fica em `../src/BASE` e
o plano em `agent-config/doc_dev/planejamento/PLANO.md`.

O trabalho segue o protocolo de [`agent-config/DIRETRIZES-IA.md`](agent-config/DIRETRIZES-IA.md),
baseado em [addyosmani/agent-skills](https://github.com/addyosmani/agent-skills):

```
/spec  ->  /plan  ->  /build  ->  /test  ->  /review  ->  /code-simplify  ->  /ship
```

- Especificações em [`agent-config/specs/`](agent-config/specs/README.md): a do projeto
  (objetivo, limites, definição de pronto), uma por stack e uma pasta por funcionalidade.
- O agent-skills fica **na íntegra** em `agent-config/vendor/agent-skills` (git subtree) e é
  adaptado por `agent-config/adaptacao/`; o `.claude/` é gerado dos dois:

```powershell
powershell -File agent-config\ferramentas\instalar-skills.ps1          # regera .claude
powershell -File agent-config\ferramentas\atualizar-agent-skills.ps1   # traz a versão nova do original
powershell -File agent-config\ferramentas\verificar-tudo.ps1 [-Rapido | -Completo]   # a verificação (/test)
```

### Fluxo de branches (git-flow)

| Branch | Uso |
|---|---|
| `main` | versão publicada na rede; cada publicação tem a tag `vX.Y` |
| `develop` | integração do que vai para a próxima publicação |
| `feature/<nome>` | trabalho novo, sai de `develop` e volta para ela |
| `bugfix/<nome>` | correção em `develop` |
| `release/X.Y` | preparação da publicação; ao terminar, merge na `main` + tag `vX.Y` |
| `hotfix/<nome>` | correção urgente do que está publicado, sai da `main` |

Para abrir uma branch use `git flow <tipo> start <nome>`. Para finalizar, use
`powershell -File agent-config\ferramentas\finalizar-branch.ps1 [-Enviar]` em vez de
`git flow <tipo> finish`: o GitFlow .NET 2.3.0 (winget `Kubis1982.GitFlow`) falha em todo
merge `--no-ff`. O script faz o mesmo finish (com a tag `vX.Y` em release e hotfix), e
`-Enviar` faz o push.

Os merges são `--no-ff`. O número da tag é o que o build gravou em `VERSAO.txt` ao
publicar; para montar sem publicar (e sem subir a versão), use `MONTAR-FRONTEND.bat -Teste`.

Documentação, comentários e mensagens em pt-BR. O repositório é público: exemplos e dados
de teste são sempre fictícios.

## Fora do versionamento

Banco (`*.accdb`), planilhas (`*.xlsm`), `VERSAO-FRONT.txt`, o `CHANGELOG.md` gerado
pelo build e tudo em `SISTEMA/DADOS/execucao/` (saídas de build, prévias, logs) são
gerados ou contêm dados reais e ficam fora do git (ver `.gitignore`).
