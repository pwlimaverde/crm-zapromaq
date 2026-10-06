---
description: "Orquestra o ciclo completo de desenvolvimento de ponta a ponta (diagnóstico, plano, TDD autônomo, revisão modular e entrega)"
argument-hint: "<funcionalidade ou correção a ser feita>"
---

Invoque a skill `ciclo` para conduzir a demanda informada em `$ARGUMENTS`.

Siga estritamente o ciclo de 5 estágios:
1. **Diagnóstico e Mapeamento:** Avalie o impacto com base nas especificações do projeto (`CLAUDE.md`, `specs/PROJETO.md`, specs das stacks) e mapeie as skills e agentes que atuarão em cada fase.
2. **Plano Envelopado:** Crie a pasta da funcionalidade em `agent-config/specs/funcionalidades/NNN-nome/` com `SPEC.md`, `plan.md` e `todo.md`. Apresente o plano ao usuário e **pare no GATE 1 para obter aprovação explícita**.
3. **Execução Passo a Passo Autônoma:** Abra a branch no GitFlow e execute cada fatia com TDD (RED → GREEN → REGRESSÃO → COMMIT). Em caso de falha de teste ou erro, **autocorrija em loop fechado sem interromper o usuário**. Só avance após cada fatia estar 100% validada.
4. **Verificação e Revisão Modular:** Execute `powershell -File agent-config\ferramentas\verificar-tudo.ps1 -Completo` e realize a auditoria com os revisores especializados (`code-reviewer`, `security-auditor`, `test-engineer` e `code-simplification`).
5. **Relatório e Entrega:** Apresente o resumo executivo completo e **pergunte formalmente no GATE 2** se pode finalizar a branch com `powershell -File agent-config\ferramentas\finalizar-branch.ps1 -Enviar`.
