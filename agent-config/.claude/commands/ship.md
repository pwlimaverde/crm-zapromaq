---
description: Run the pre-launch checklist via parallel fan-out to specialist personas, then synthesize a go/no-go decision
---

Invoke the shipping-and-launch skill.

`/ship` is a **fan-out orchestrator**. It runs three specialist personas in parallel against the current change, then merges their reports into a single go/no-go decision with a rollback plan. The personas operate independently — no shared state, no ordering — which is what makes parallel execution safe and useful here.

## Phase A — Parallel fan-out

Spawn three subagents concurrently using the Agent tool. **Issue all three Agent tool calls in a single assistant turn so they execute in parallel** — sequential calls defeat the purpose of this command.

In Claude Code, each call passes `subagent_type` matching the persona's `name` field:

1. **`code-reviewer`** — Run a five-axis review (correctness, readability, architecture, security, performance) on the staged changes or recent commits. Output the standard review template.
2. **`security-auditor`** — Run a vulnerability and threat-model pass. Check OWASP Top 10, secrets handling, auth/authz, dependency CVEs. Output the standard audit report.
3. **`test-engineer`** — Analyze test coverage for the change. Identify gaps in happy path, edge cases, error paths, and concurrency scenarios. Output the standard coverage analysis.

In other harnesses without an Agent tool, invoke each persona's system prompt sequentially and treat their outputs as if returned in parallel — the merge phase still works.

Constraints (from Claude Code's subagent model):
- Subagents cannot spawn other subagents — do not let one persona delegate to another.
- Each subagent gets its own context window and returns only its report to this main session.
- If you need teammates that talk to each other instead of just reporting back, use Claude Code Agent Teams and reference these personas as teammate types (see `references/orchestration-patterns.md`).

**Persona resolution.** If you've defined your own `code-reviewer`, `security-auditor`, or `test-engineer` in `.claude/agents/` or `~/.claude/agents/`, those take precedence over this plugin's versions — `/ship` picks up your customizations automatically. This is intentional: plugin subagents sit at the bottom of Claude Code's scope priority table, so user-level definitions win by design.

## Phase B — Merge in main context

Once all three reports are back, the main agent (not a sub-persona) synthesizes them:

1. **Code Quality** — Aggregate Critical/Important findings from `code-reviewer` and any failing tests, lint, or build output. Resolve duplicates between reviewers.
2. **Security** — Promote any Critical/High `security-auditor` findings to launch blockers. Cross-reference with `code-reviewer`'s security axis.
3. **Performance** — Pull from `code-reviewer`'s performance axis; cross-check Core Web Vitals if applicable.
4. **Accessibility** — Verify keyboard nav, screen reader support, contrast (not covered by the three personas — handle directly here, or invoke the accessibility checklist).
5. **Infrastructure** — Env vars, migrations, monitoring, feature flags. Verify directly.
6. **Documentation** — README, ADRs, changelog. Verify directly.

## Phase C — Decision and rollback

Produce a single output:

```markdown
## Ship Decision: GO | NO-GO

### Blockers (must fix before ship)
- [Source persona: Critical finding + file:line]

### Recommended fixes (should fix before ship)
- [Source persona: Important finding + file:line]

### Acknowledged risks (shipping anyway)
- [Risk + mitigation]

### Rollback plan
- Trigger conditions: [what signals would prompt rollback]
- Rollback procedure: [exact steps]
- Recovery time objective: [target]

### Specialist reports (full)
- [code-reviewer report]
- [security-auditor report]
- [test-engineer report]
```

## Rules

1. The three Phase A personas run in parallel — never sequentially.
2. Personas do not call each other. The main agent merges in Phase B.
3. The rollback plan is mandatory before any GO decision.
4. If any persona returns a Critical finding, the default verdict is NO-GO unless the user explicitly accepts the risk.
5. **Skip the fan-out only if all of the following are true:** the change touches 2 files or fewer, the diff is under 50 lines, and it does not touch auth, payments, data access, or config/env. Otherwise, default to fan-out. `/ship` is designed for production-bound changes — when the blast radius is non-trivial, run the parallel review even if the diff looks small.

## Neste projeto: o que é "entregar"

Há dois níveis — confirme com o usuário qual é:

1. **Integrar a branch:** `verificar-tudo` (padrão; `-Completo` se tocou `fonte\`/`build\`) →
   revisão → `ferramentas\finalizar-branch.ps1 -Enviar` (merge `--no-ff` em `develop` + push).
2. **Publicar na empresa:** `git flow release start X.Y` (X.Y = a versão que a rede vai publicar:
   a de `VERSAO-FRONT.txt` de lá + 1) → `verificar-tudo -Completo` →
   `ferramentas\empacotar-base.ps1` → o usuário extrai o pacote **por cima** da BASE da rede e
   roda `MONTAR-FRONTEND.bat` → só então `finalizar-branch.ps1 -Enviar` (main + tag `vX.Y`).

Itens extras do go/no-go: pacote conferido pelo script; migrações revisadas (estrutural exige
banco exclusivo); a BASE da rede foi trazida antes (a versão da rede tem prioridade);
plano de volta = `execucao\backup\antes-da-publicacao\<data>` + `execucao\historico-front`.
Acessibilidade web e Core Web Vitals não se aplicam.

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
