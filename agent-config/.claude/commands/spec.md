---
description: Start spec-driven development — write a structured specification before writing code
---

Invoke the spec-driven-development skill.

Begin by understanding what the user wants to build. Ask clarifying questions about:
1. The objective and target users
2. Core features and acceptance criteria
3. Tech stack preferences and constraints
4. Known boundaries (what to always do, ask first about, and never do)

Then generate a structured spec covering all six core areas: objective, commands, project structure, code style, testing strategy, and boundaries.

If the request bundles several independently testable capabilities, first propose a capability map (module ids, dependency direction, build order) per the skill's Phase 0 and get it approved, then spec each module in dependency order.

Save the spec as SPEC.md in the project root and confirm with the user before proceeding.

## Neste projeto: onde e como fica a spec

- Salve em `agent-config/specs/funcionalidades/NNN-nome-curto/SPEC.md` (NNN = próximo número
  livre), partindo de `agent-config/specs/funcionalidades/_modelo/SPEC.md` — não em `SPEC.md` na raiz.
- Tech Stack, Commands, Project Structure, Code Style e Testing Strategy: **referencie** as specs
  de stack (`agent-config/specs/stacks/`) e escreva só o que for específico da funcionalidade.
- Boundaries: as restrições inegociáveis de `agent-config/specs/PROJETO.md` valem sempre; liste
  só as específicas.
- Entrevista em pt-BR, uma ou duas perguntas por vez. Regra de negócio e decisão de interface
  são do usuário; consulte também `src/BASE/SISTEMA/DADOS/docs/legado/` (estudo da migração
  e diário de decisões) antes de propor algo que o legado já decidiu.
- A spec é commitada junto com o início da branch (`git flow feature start <nome>`).

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
