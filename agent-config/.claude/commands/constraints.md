---
description: Define and enforce this project's quality bar — interview, sane defaults, CONSTRAINTS.md
---

Invoke the constraint-driven-development skill.

$ARGUMENTS

Default behaviour with no arguments: set up constraints for this repository.

1. **Detect first.** Read package.json / pyproject.toml / go.mod, the test runner, existing lint configs, current coverage output, CI workflows, and the agent harness in use. Report what you found in two lines. Never ask for anything you can read.

2. **Interview, at most four questions.** One at a time, each with your best guess and a usable default so "I don't know" still produces a working config:
   - Which dimensions beyond the floor (coverage, security, performance, accessibility, architecture)
   - Block or warn when a check fails mid-task
   - Target numbers, or measure today's values and hold them
   - Slowest check tolerated before handing work back

3. **Write CONSTRAINTS.md** at the repo root with a Floor section, enforced numbers, measured-only metrics with today's values, and an exceptions table with owners and expiry dates. Every number needs a stated reason.

4. **Install what each picked dimension needs.** A dimension with a number and no tool behind it is an aspiration. Use the de facto tool so existing config keeps working: Semgrep for code scanning, gitleaks (always `--redact`) for secrets, osv-scanner for dependencies, axe-core for accessibility, Lighthouse for web vitals, size-limit for bundles, dependency-cruiser for boundaries, Stryker for assertion quality. Record the exact command next to each rule in CONSTRAINTS.md. Accessibility and performance need a running URL; if the project has none, say so and drop the dimension rather than inventing a check. Add the commands to package.json as check:fast / check:task / check:full.

5. **Place each check by cost.** Types, lint and secrets in the edit loop (seconds). Related tests and changed-line coverage at task end (under 90s). Everything else at review or in CI. Scope checks to the diff, not the whole repo.

6. **Point the agent at it.** Add a line to CLAUDE.md telling agents to read CONSTRAINTS.md and never weaken it to make a change pass.

7. **Verify.** Run the constraints against the current branch. If anything fails that the user disagrees with, fix the constraint now rather than leaving a gate people will learn to ignore.

Sub-commands:
- `/constraints check` — run the current constraints against this branch and report
- `/constraints guard` — inspect the diff for a weakened bar: lowered thresholds, skipped or deleted tests, new suppression comments, unfinished stubs, new exceptions
- `/constraints ratchet` — record today's measured values as the floor that must not fall

## Neste projeto

Sem `package.json` nem CI: a barra de qualidade é a do `ferramentas\verificar-tudo.ps1` e das
seções **Revisão** das specs de stack. Se for criar `CONSTRAINTS.md`, ele fica em
`agent-config/specs/CONSTRAINTS.md` e os comandos registrados são os níveis do `verificar-tudo`.
Ferramentas que exigem internet ou Node (Semgrep em nuvem, Lighthouse, axe) não entram no
fluxo do sistema; no desenvolvimento, só com aprovação do usuário.

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
