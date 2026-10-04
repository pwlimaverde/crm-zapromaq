## Neste projeto: branch, teste primeiro e paradas obrigatórias

- **Branch:** se estiver em `develop`, abra antes da primeira tarefa
  `git flow feature start <nome>` (ou `bugfix`). Spec e plano vão no primeiro commit dela.
- **RED por stack** (detalhe em `agent-config/specs/stacks/`): lógica VBA pura → novo `Confere`
  em `modAutoteste.bas` (falha no `verificar-tudo -Completo`); banco/SQL/MCP → caso em
  `IA/teste/testar-mcp.ps1` ou teste PowerShell; build/operação → `Conferir` em
  `testar-build-lib.ps1`; regra estática → checagem no `verificar.py`; tela → prévia
  (`gerar-visual.ps1`) + item 🖥 na estação.
- **Commit por tarefa**, só com os arquivos dela, mensagem em pt-BR dizendo o porquê, com as
  linhas de atribuição da sessão.
- **Pare e pergunte também quando:** migração **estrutural** do banco (`EsquemaPara`), qualquer
  coisa que toque o banco de produção, `MONTAR-FRONTEND.bat` **sem** `-Teste` (publica e sobe
  versão), push na `main`, pacote para a empresa (`empacotar-base.ps1`), mudança no contrato
  de `Início!B7`, nos nomes de controle dos formulários ou nos códigos `CT-`/`AT-`.
