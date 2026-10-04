# GEMINI.md

Este arquivo define as diretrizes de trabalho do **Antigravity** neste projeto, espelhando as configurações e instruções do Claude Code para manter uma **única fonte da verdade**.

---

## 1. Diretriz Central e Fontes Oficiais da Verdade

Toda a engenharia, arquitetura e decisões deste repositório estão documentadas em arquivos canônicos. **Você DEVE consultar e seguir rigorosamente:**

1. **[CLAUDE.md](file:///c:/PROJETOS/VBA/crm-zapromaq/agent-config/CLAUDE.md)**: Arquitetura completa, layout do repositório, restrições inegociáveis de ambiente (Excel 2019, ACE 16 64-bit, sem internet), convenções GitFlow, ferramentas de build e testes.
2. **[DIRETRIZES-IA.md](file:///c:/PROJETOS/VBA/crm-zapromaq/agent-config/DIRETRIZES-IA.md)**: Protocolo de engenharia com IA (*Beyond Vibe Coding*), os 4 pilares (Contexto, Spec, Fatias pequenas, Verificação real) e critérios de parada para validação com o humano.
3. **[specs/PROJETO.md](file:///c:/PROJETOS/VBA/crm-zapromaq/agent-config/specs/PROJETO.md)**: Objetivo, limites, definição de pronto e TDD por stack.
4. **[doc_dev/planejamento/PLANO.md](file:///c:/PROJETOS/VBA/crm-zapromaq/agent-config/doc_dev/planejamento/PLANO.md)**: Checklist de fases, decisões tomadas (D01...), correções (C01...) e pendências (P01...).

---

## 2. Layout de Trabalho

O workspace ativo é `agent-config/` e o trabalho incide sobre a pasta raiz do repositório (`..`).

- Entregável: `../src/BASE/` (estrutura fixa, nunca reorganizar).
- Skills do projeto: localizadas em `.claude/skills/` e registradas para o Antigravity via `.agents/skills.json`.
- Scripts de desenvolvimento: `ferramentas/` (ex.: `instalar-skills.ps1`, `verificar-tudo.ps1`, `finalizar-branch.ps1`).

---

## 3. O Ciclo de Desenvolvimento e Habilidades (Skills)

Todo trabalho segue o ciclo estruturado:
```
DEFINIR      PLANEJAR     CONSTRUIR    VERIFICAR    REVISAR      SIMPLIFICAR      ENTREGAR
/spec   ->   /plan   ->   /build  ->   /test   ->   /review  ->  /code-simplify -> /ship
```

As 23 habilidades especializadas do projeto estão em [agent-config/.claude/skills/](file:///c:/PROJETOS/VBA/crm-zapromaq/agent-config/.claude/skills/) e são carregadas sob demanda:
- `spec-driven-development` (para `/spec`)
- `planning-and-task-breakdown` (para `/plan`)
- `test-driven-development` / `incremental-implementation` (para `/build`)
- `code-review-and-quality` / `security-and-hardening` (para `/review`)
- `code-simplification` (para `/code-simplify`)
- `shipping-and-launch` / `git-workflow-and-versioning` (para `/ship`)

---

## 4. Regras Inegociáveis do Projeto

- **Excel 2019 travado**: Nada de recursos exclusivos do Microsoft 365 (`XLOOKUP`, `LET`, `LAMBDA`, matrizes dinâmicas).
- **Sem internet / 100% local**: Sem APIs externas, CDN ou pacotes externos em tempo de execução.
- **Codificação de arquivos**:
  - VBA: Windows-1252 (cp1252) com CRLF.
  - PowerShell (`.ps1`): UTF-8 com BOM (PowerShell 5.1 lê sem BOM como ANSI).
- **Verificação real**: Nada é considerado pronto sem a prova de execução de `powershell -File ferramentas\verificar-tudo.ps1` sem falhas.
- **GitFlow disciplinado**: Branches `feature/`, `bugfix/`, `release/`, `hotfix/`. Finalização via `ferramentas\finalizar-branch.ps1`.
