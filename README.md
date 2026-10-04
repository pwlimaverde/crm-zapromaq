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
│   └── doc_dev/planejamento/
│       ├── PLANO.md       checklist de fases, decisões, correções e pendências
│       └── referencias/   documentação oficial offline (VBA, MSForms, Excel, ADO, ACE…)
└── src/
    └── BASE/              ENTREGÁVEL — a pasta copiada inteira para a rede
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
Em resumo: copiar `src/BASE/SISTEMA` (e `IA`, se usar o Claude Desktop) para a pasta da
rede e executar `MONTAR-FRONTEND.bat`. O build faz backup, aplica migrações, monta,
compila, testa e só então substitui o `.xlsm` publicado e sobe a versão.

Ajustes feitos diretamente na rede pelo MCP de desenvolvimento voltam ao repositório com
`SISTEMA\DADOS\ferramentas\sincronizar\trazer-da-rede.ps1`.

## Desenvolvimento com IA

Abra o Claude Code dentro de `agent-config/`. O `CLAUDE.md` descreve arquitetura,
restrições e regras aprendidas na produção; o código manipulado fica em `../src/BASE` e
o plano em `agent-config/doc_dev/planejamento/PLANO.md`.

Fluxo de branches: cada fase do plano é desenvolvida em `fase-N-...` e integrada em
`main` com merge `--no-ff` + tag `fase-N`. Documentação, comentários e mensagens em pt-BR.

## Fora do versionamento

Banco (`*.accdb`), planilhas (`*.xlsm`), `VERSAO-FRONT.txt`, o `CHANGELOG.md` gerado
pelo build e tudo em `SISTEMA/DADOS/execucao/` (saídas de build, prévias, logs) são
gerados ou contêm dados reais e ficam fora do git (ver `.gitignore`).
