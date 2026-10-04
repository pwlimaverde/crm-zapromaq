---
description: Conduct a five-axis code review — correctness, readability, architecture, security, performance
---

Invoke the code-review-and-quality skill.

Review the current changes (staged or recent commits) across all five axes:

1. **Correctness** — Does it match the spec? Edge cases handled? Tests adequate?
2. **Readability** — Clear names? Straightforward logic? Well-organized?
3. **Architecture** — Follows existing patterns? Clean boundaries? Right abstraction level?
4. **Security** — Input validated? Secrets safe? Auth checked? (Use security-and-hardening skill)
5. **Performance** — No N+1 queries? No unbounded ops? (Use performance-optimization skill)

Categorize findings as Critical, Important, or Suggestion.
Output a structured review with specific file:line references and fix recommendations.

## Neste projeto: o que mais olhar

Além dos cinco eixos, confira a seção **Revisão** de cada spec de stack tocada
(`agent-config/specs/stacks/`). O mínimo, sempre:

- SQL parametrizado; data só como parâmetro; `UPDATE` com `WHERE id=? AND versao=?` e log na
  mesma transação; nenhuma consulta dentro de laço.
- VBA em Windows-1252 + CRLF; nada do Microsoft 365 (Excel é 2019); nada de OCX, URL ou
  subclassing; nomes de controle e `Início!B7` intactos.
- `.ps1` UTF-8 com BOM + CRLF; nada no stdout do servidor MCP além de JSON-RPC.
- Nenhum dado real; nenhum caminho absoluto da máquina de desenvolvimento nem letra de unidade.

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
