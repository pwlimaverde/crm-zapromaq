# Spec do projeto: CRM Zapromaq

## Objetivo

CRM comercial da Zapromaq (vendas de máquinas industriais): cadastro de clientes e pré-clientes,
contatos, atendimentos/oportunidades com próxima ação, painel de metas e indicadores, e uma IA
(Claude Desktop, via MCP) que consulta e atualiza os dados e mantém o sistema.

**Usuários:** vendedores e gestão comercial (front Excel em cada estação); o mantenedor do
sistema (desenvolvimento aqui com Claude Code; ajustes na empresa com Claude Desktop).

**Situação:** em produção desde 16/09/2026 (migração da planilha 3.6); versão publicada na rede:
1.5. Modernização em fases no `agent-config/doc_dev/planejamento/PLANO.md`.

## Tech stack

| Camada | Tecnologia | Spec |
|---|---|---|
| Front | Excel **2019** 64 bits, VBA 7, MSForms, ADO late binding | [front-vba](stacks/front-vba.md) |
| Banco | Access `.accdb` em pasta de rede (UNC), `Microsoft.ACE.OLEDB.16.0` | [banco-access](stacks/banco-access.md) |
| Build e operação | Windows PowerShell **5.1**, automação COM do Excel, `.bat` | [build-operacao-powershell](stacks/build-operacao-powershell.md) |
| Visual | layout em JSON, prévia HTML/JS no Edge headless, fundos BMP | [visual-layout](stacks/visual-layout.md) |
| IA | servidores MCP em PowerShell (stdio), skills do Claude Desktop/Codex | [ia-mcp](stacks/ia-mcp.md) |
| Desenvolvimento | Python 3, ANTLR, git + git-flow, agent-skills | [ferramentas-dev](stacks/ferramentas-dev.md) |

## Comandos

```powershell
# verificação (o /test): -Rapido | padrão | -Completo
powershell -File agent-config\ferramentas\verificar-tudo.ps1 [-Rapido | -Completo]
# montar sem publicar (Excel): resultado em src\BASE\SISTEMA\DADOS\execucao\teste
src\BASE\SISTEMA\MONTAR-FRONTEND.bat -Teste
# banco fictício para ver o front (20 clientes, contatos e atendimentos)
powershell -File src\BASE\SISTEMA\DADOS\testes\criar-banco-demo.ps1 -Recriar
# pacote para a empresa (sai do commit)
powershell -File agent-config\ferramentas\empacotar-base.ps1
# git-flow
git flow feature start <nome>
powershell -File agent-config\ferramentas\finalizar-branch.ps1 [-Enviar]
```

## Estrutura

```
crm-zapromaq/
  README.md
  agent-config/       Claude Code é aberto aqui: CLAUDE.md, .claude (gerado), adaptacao,
                      vendor/agent-skills, specs, ferramentas, doc_dev (plano e referências)
  src/BASE/           o entregável — estrutura FIXA (a rede depende dela)
    INICIAR-CRM.bat   lançador das estações
    SISTEMA/          MONTAR-FRONTEND.bat + DADOS/ (fonte, build, banco, operacao, testes, docs, lib, ferramentas, execucao)
    IA/               mcp/ (dados), mcp-dev/ (desenvolvimento), instrucoes/ (skills), teste/
  dist/               pacotes gerados (fora do git)
```

## Estilo (geral)

- pt-BR em código, comentários, mensagens e documentos. Nomes descritivos em português.
- Codificação por tipo: VBA **Windows-1252 + CRLF**; `.ps1` **UTF-8 com BOM + CRLF**; `.bat`
  **ASCII + CRLF**; Markdown/JSON UTF-8 sem BOM. O git não converte (`.gitattributes: * -text`).
- Comentário explica o **porquê** (o legado registra o motivo de cada "esquisitice").
- Detalhes por linguagem nas specs de stack.

## Estratégia de teste (TDD por stack)

| Mudança em | Teste que falha primeiro | Nível do `verificar-tudo` |
|---|---|---|
| Lógica VBA pura | `Confere` em `fonte/modulos/modAutoteste.bas` | `-Completo` |
| Regra estática VBA/esquema/encoding | checagem em `ferramentas/verificacao/verificar.py` | `-Rapido` |
| Build, versão, pacote, B7 | `Conferir` em `testes/testar-build-lib.ps1` | `-Rapido` |
| Banco / SQL / MCP de dados | caso em `IA/teste/testar-mcp.ps1` | padrão |
| MCP de desenvolvimento | caso em `IA/teste/testar-mcp-dev.ps1` | padrão |
| Grades e Painel com dados | `testes/testar-uso.ps1` (no build) | `-Completo` |
| Aparência | prévia (`gerar-visual.ps1`) + 🖥 calibração na estação | manual |
| Ferramenta de desenvolvimento (`agent-config/`) | `roteador/testar_roteador.py` ou checagem no `verificar-tudo.ps1` | `-Rapido` |

Dados de teste sempre fictícios. Teste que altera banco restaura o estado.

## Definição de pronto

1. Teste novo cobrindo a mudança (falhou antes, passa agora).
2. `verificar-tudo` **padrão** sem falha; **`-Completo`** se tocou `fonte\`, `build\` ou layout.
   A saída é colada como prova.
3. Spec de stack atualizada se a stack mudou; `PLANO.md` atualizado se mudou fase/decisão/pendência.
4. Commit atômico na branch do git-flow, mensagem em pt-BR com o porquê.
5. Itens 🖥 (Excel 2019 / rede) listados para o usuário conferir — "pronto aqui" não é
   "pronto na empresa".

## Limites

**Sempre**
- 100% local: nada de internet, CDN, API externa, pacote ou OCX/DLL nas estações.
- Excel **2019**: nada de matriz dinâmica, `LET`, `XLOOKUP`, `LAMBDA` etc. (esta máquina tem 365).
- O entregável é a pasta `src/BASE`: caminhos relativos ou UNC; nada da máquina de desenvolvimento.
- SQL parametrizado; data como parâmetro; controle de versão otimista; log na mesma transação.
- A versão da rede tem prioridade: a BASE de lá volta para o repositório antes de qualquer ciclo.

**Perguntar antes**
- Migração estrutural do banco; mudar contrato (B7, `VERSAO-FRONT.txt`, nomes de controle,
  ferramentas MCP, códigos `CT-`/`AT-`); publicar (`MONTAR-FRONTEND.bat` sem `-Teste`); gerar
  pacote para a empresa; push na `main`; ferramenta nova de desenvolvimento.

**Nunca**
- Dado real no repositório (público). Editar o `.xlsm` à mão. `Microsoft.Jet.OLEDB.4.0`.
  Recalcular/reformatar códigos `CT-`/`AT-`. Gravar `ctx_*` pelo formulário. Subclassing no
  UserForm. Apagar a BASE da rede no lugar de extrair o pacote por cima.

## Critérios de sucesso do projeto

Os critérios de aceite do legado continuam valendo (`PLANO.md` §7): nenhuma sobrescrita
silenciosa, nenhum registro duplicado em falha, nenhum código alterado, painel coerente com o
banco, tela que cabe em 1366×768.

## Questões em aberto

As pendências do `PLANO.md` §8 (P01…P09).
