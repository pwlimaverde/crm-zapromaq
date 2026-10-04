## Neste projeto: onde fica o plano e como é cada tarefa

- `plan.md` e `todo.md` ficam na pasta da funcionalidade
  (`agent-config/specs/funcionalidades/NNN-nome/`), não em `tasks/`. Use os modelos de `_modelo/`.
- Cada tarefa diz: stack(s) tocada(s), **teste que falha primeiro** (ver a tabela de TDD em
  `agent-config/specs/PROJETO.md`), nível de verificação (`verificar-tudo -Rapido`, padrão ou
  `-Completo`) e arquivos.
- Fatias verticais atravessam as camadas (ex.: campo novo = migração do banco + `modSchema` +
  consulta em `modCRM` + MCP + teste), não "todo o banco, depois todo o VBA".
- Marque com 🖥 a tarefa que só se confere numa estação com **Excel 2019** (esta máquina tem
  Excel 365) ou na rede da empresa; ela vira um checkpoint humano.
