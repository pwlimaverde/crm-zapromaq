---
name: development-lifecycle
description: Orquestra o ciclo completo de desenvolvimento de ponta a ponta (diagnóstico prévio de skills/agentes, planejamento envelopado em fatias, execução autônoma passo a passo com TDD e autocorreção, revisão modular por personas e portão formal de finalização/merge/push). Acionável via /ciclo.
---

# Development Lifecycle (Orquestrador de Desenvolvimento)

## Visão Geral

Esta skill orquestra o ciclo completo de desenvolvimento de software no projeto **CRM Zapromaq**, unificando as habilidades de engenharia (*agent-skills*) e as restrições arquiteturais em um fluxo guiado, autônomo e seguro.

Em vez de exigir a chamada manual de cada etapa (`/spec`, `/plan`, `/build`, `/test`, `/review`, `/code-simplify`, `/ship`), o orquestrador conduz o processo através de **fases claras com portões de aprovação (gates) estratégicos**, dando autonomia total ao agente para corrigir problemas durante a construção sem interromper o usuário desnecessariamente.

---

## Os 5 Estágios do Ciclo de Desenvolvimento

```
[Entrada do Usuário]
         │
         ▼
[Fase 1: Diagnóstico e Mapeamento de Agentes/Skills]
         │
         ▼
[Fase 2: Envelopamento do Plano em Fatias (plan.md / todo.md)]
         │
         ▼
 🛑 GATE 1: Aprovação Explícita do Plano pelo Usuário
         │
         ▼
[Fase 3: Execução Passo a Passo Autônoma]
         ├─ GitFlow: git flow feature start <nome> | bugfix/
         ├─ Para cada fatia:
         │   ├─ Teste que falha primeiro (RED)
         │   ├─ Código mínimo (GREEN)
         │   ├─ Autocorreção autônoma em caso de falha (sem interrupção)
         │   └─ Commit atômico (só avança se validada 100%)
         ▼
[Fase 4: Verificação Integral e Revisão Modular por Personas]
         ├─ Prova real de execução: verificar-tudo.ps1
         ├─ Revisor 1: code-reviewer.md
         ├─ Revisor 2: security-auditor.md
         ├─ Revisor 3: test-engineer.md
         └─ Simplificação: code-simplification
         ▼
[Fase 5: Relatório Consolidado de Entrega]
         │
         ▼
 🛑 GATE 2: Questionamento Formal de Finalização
         │  "Posso finalizar a branch, realizar o merge no develop e o push?"
         ▼
[Execução de finalizar-branch.ps1 -Enviar e Fechamento]
```

---

## Detalhamento dos Estágios

### Estágio 1: Diagnóstico e Mapeamento de Agentes/Skills

Ao receber a demanda do usuário:
1. **Consulte a hierarquia canônica de contexto:**
   - `agent-config/CLAUDE.md` e `agent-config/GEMINI.md` (regras centrais e restrições)
   - `agent-config/specs/PROJETO.md` (objetivo, definição de pronto, TDD por stack)
   - `agent-config/specs/stacks/*.md` (especificações das stacks afetadas: `front-vba`, `banco-access`, `build-operacao-powershell`, etc.)
   - `agent-config/doc_dev/planejamento/PLANO.md` (decisões históricas e roadmap)
2. **Avalie o escopo e identifique ambiguidades:** se houver requisitos fundamentais ausentes ou conflitos de regra de negócio, esclareça antes de formalizar o plano (usando `interview-me` ou perguntas diretas).
3. **Mapeie explicitamente os Agentes e Skills que atuarão no ciclo:**
   - **Especificação / Requisitos:** `spec-driven-development`
   - **Planejamento:** `planning-and-task-breakdown`, `constraint-driven-development`
   - **Construção e TDD:** `test-driven-development`, `incremental-implementation` (+ `source-driven-development` para APIs/docs oficiais)
   - **Depuração e Autocorreção:** `debugging-and-error-recovery`
   - **Verificação:** `verificar-tudo.ps1`
   - **Revisão Modular:** personas `code-reviewer`, `security-auditor`, `test-engineer` e skill `code-simplification`
   - **Entrega / Git:** `git-workflow-and-versioning`, `shipping-and-launch`

---

### Estágio 2: Envelopamento e Persistência do Plano

1. **Defina a numeração e diretório da funcionalidade:**
   - Local: `agent-config/specs/funcionalidades/NNN-nome/` (baseado no último número existente).
2. **Gere os três artefatos canônicos:**
   - `SPEC.md`: problema a resolver, critérios de aceite objetivos, limites e restrições inegociáveis.
   - `plan.md`: quebra em fatias verticais finas, dependências, estratégia de teste para cada fatia e marcação de itens que exigem a estação física (`🖥 estação`).
   - `todo.md`: lista de verificação passo a passo com caixas de checagem.
3. **Apresente o plano envelopado no chat:**
   - Exiba o diagnóstico, o mapeamento de agentes/skills convocados e o checklist das fatias.
4. **🛑 GATE 1 — Parada Obrigatória de Aprovação:**
   - **Aguarde a validação explícita do usuário.**
   - Não inicie a codificação nem crie branches antes da confirmação afirmativa do usuário ("aprovado", "pode iniciar", "ok").

---

### Estágio 3: Execução Passo a Passo Autônoma (TDD + Autocorreção)

Uma vez aprovado o plano:
1. **Abertura da branch no GitFlow:**
   - Se nova funcionalidade: `git flow feature start <nome>` (ou crie a branch `feature/<nome>` a partir de `develop`).
   - Se correção: `git flow bugfix start <nome>`.
2. **Execução sequencial fatia por fatia:**
   - Para cada fatia descrita no `plan.md`:
     1. **RED:** Escreva o teste que falha primeiro cobrindo a alteração proposta, conforme a matriz de TDD de `specs/PROJETO.md`:
        - Lógica VBA: `modAutoteste.bas` (`Confere`)
        - Regra estática / encoding: `verificar.py`
        - Build / versão: `testar-build-lib.ps1`
        - Banco / MCP: `testar-mcp.ps1` / `testar-mcp-dev.ps1`
     2. **GREEN:** Implemente o código estritamente necessário para fazer o teste passar.
     3. **REGRESSÃO:** Execute os testes existentes para garantir que nada foi quebrado.
     4. **COMMIT ATÔMICO:** Faça o commit no Git com mensagem descritiva em pt-BR explicando o motivo da mudança.
3. **Loop Fechado de Autocorreção (Sem Interrupções Desnecessárias):**
   - Se um teste quebrar, o build falhar ou ocorrer um erro de sintaxe/linter:
     - **NÃO pare para pedir autorização para corrigir.**
     - Ative a lógica de `debugging-and-error-recovery`: examine as mensagens de erro, localize o ponto da falha, aplique a correção necessária e reexecute o teste.
     - Repita o ciclo até que a fatia esteja 100% verde.
4. **Critério Inegociável de Avanço:**
   - **SÓ avance para a próxima fatia após a fatia anterior estar validada e commitada.**
5. **Critérios de Parada Excepcional na Fase 3:**
   - Apenas interrompa o fluxo autônomo se:
     - Encontrar uma ambiguidade irresolúvel nas regras de negócio;
     - Deparar-se com uma ação destrutiva ou irreversível (ex.: migração estrutural de banco de dados);
     - Atingir um item explicitamente marcado como `🖥 estação` (que requer validação física no Excel 2019 / rede).

---

### Estágio 4: Verificação Integral e Revisão Modular por Personas

Com todas as fatias de código concluídas:
1. **Verificação Real Completa:**
   - Execute a suíte de testes do projeto:
     ```powershell
     powershell -File agent-config\ferramentas\verificar-tudo.ps1 -Completo
     ```
   - A saída limpa (sem erros) é a prova obrigatória de conclusão técnica.
2. **Revisão Modular por Personas (Agentes):**
   - Convoque sequencialmente os perfis especializados definidos em `agent-config/.claude/agents/`:
     - **Persona 1 — `code-reviewer`:** Avalia legibilidade, adesão às convenções (inglês nos identificadores, português nos comentários), respeito ao princípio DRY, modularidade e ausência de código morto.
     - **Persona 2 — `security-auditor`:** Audita sanitização de entradas, queries parametrizadas (zero SQL concatenado), integridade das transações de gravação (`log_alteracoes`), ausência de segredos/hardcoded e conformidade com ambiente 100% offline.
     - **Persona 3 — `test-engineer`:** Audita os testes criados, verificando se cobrem casos de borda, se não contêm falsos positivos e se não burlam as checagens reais.
     - **Passada de Simplificação (`code-simplification`):** Examina se a solução pode ser reduzida e simplificada sem alterar comportamento.
3. **Ajustes pós-revisão:** Aplique correções imediatas de eventuais achados críticos identificados pelos revisores e revalide com `verificar-tudo.ps1`.

---

### Estágio 5: Resumo Executivo e Portão de Finalização (Merge & Push)

1. **Gere o relatório consolidado de entrega contendo:**
   - Resumo das mudanças implementadas;
   - Lista de arquivos criados e modificados;
   - Evidência dos testes (status de `verificar-tudo.ps1`);
   - Parecer consolidado das 3 personas de revisão;
   - Identificação da branch atual e histórico de commits atômicos realizados.
2. **🛑 GATE 2 — Questionamento Formal de Finalização:**
   - Pergunte claramente ao usuário:
     > *"Todas as etapas foram desenvolvidas, testadas e auditadas com sucesso. Posso finalizar a branch, realizar o merge no `develop` e efetuar o push para o repositório remoto?"*
3. **Execução do Fechamento e Envio:**
   - Após a autorização explícita do usuário:
     ```powershell
     powershell -File agent-config\ferramentas\finalizar-branch.ps1 -Enviar
     ```
   - Confirme a conclusão do merge no `develop`, a remoção limpa da branch local de trabalho e o push no repositório.
