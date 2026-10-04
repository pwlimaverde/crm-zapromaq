---
description: Simplify code for clarity and maintainability — reduce complexity without changing behavior
---

Invoke the code-simplification skill.

Simplify recently changed code (or the specified scope) while preserving exact behavior:

1. Read CLAUDE.md and study project conventions
2. Identify the target code — recent changes unless a broader scope is specified
3. Understand the code's purpose, callers, edge cases, and test coverage before touching it
4. Scan for simplification opportunities:
   - Deep nesting → guard clauses or extracted helpers
   - Long functions → split by responsibility
   - Nested ternaries → if/else or switch
   - Generic names → descriptive names
   - Duplicated logic → shared functions
   - Dead code → remove after confirming
5. Apply each simplification incrementally — run tests after each change
6. Verify all tests pass, the build succeeds, and the diff is clean

If tests fail after a simplification, revert that change and reconsider. Use `code-review-and-quality` to review the result.

## Neste projeto: o que não se simplifica

Siga o trecho "Neste projeto" da skill `code-simplification`: blocos marcados com
`simplify-ignore` não se tocam, e vários padrões "estranhos" do código são decisões
registradas (`src/BASE/SISTEMA/DADOS/docs/legado/estado-migracao.md`, `PLANO.md` §3–4c).

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
