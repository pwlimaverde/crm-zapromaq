---
description: Break work into small verifiable tasks with acceptance criteria and dependency ordering
---

Invoke the planning-and-task-breakdown skill.

Read the existing spec (SPEC.md or equivalent) and the relevant codebase sections. Then:

1. Enter plan mode — read only, no code changes
2. Identify the dependency graph between components
3. Slice work vertically (one complete path per task, not horizontal layers)
4. Write tasks with acceptance criteria and verification steps
5. Add checkpoints between phases
6. Present the plan for human review

Save the plan to tasks/plan.md and task list to tasks/todo.md.

If tasks/plan.md or tasks/todo.md already exists with unchecked tasks for different work, stop and ask before writing — never silently overwrite an incomplete plan.

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

---

## Neste projeto: CRM Zapromaq

- **Idioma:** pt-BR em perguntas, documentos, comentários e mensagens de commit.
- **Contexto antes de agir:** `agent-config/CLAUDE.md` (regras) › `agent-config/specs/PROJETO.md`
  (spec do projeto) › `agent-config/specs/stacks/*.md` das camadas que a tarefa toca ›
  spec da funcionalidade. Em sessão longa, releia a spec em vez de supor.
- **Caminhos:** o Claude Code é aberto em `agent-config/`; a raiz do git é `..`; o entregável
  é `../src/BASE` (estrutura fixa, não reorganizar).
- **Artefatos de uma funcionalidade:** `agent-config/specs/funcionalidades/NNN-nome/`
  com `SPEC.md`, `plan.md` e `todo.md` — este é o lugar de `SPEC.md` e `tasks/` do original.
- **Verificação:** `powershell -File ferramentas\verificar-tudo.ps1` (`-Rapido` | padrão | `-Completo`).
  "Pronto" exige a saída dele (definição em `specs/PROJETO.md`).
- **Git:** git-flow. Nunca commitar em `main`/`develop`; finalizar com
  `ferramentas\finalizar-branch.ps1 [-Enviar]`.
- **Repositório público:** nenhum dado real (cliente, CNPJ, telefone, valor) em código, teste ou doc.
