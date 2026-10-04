# Especificações

Fonte de verdade do **o quê** e do **com quê** deste projeto, no formato do
`spec-driven-development` (agent-skills). Hierarquia de contexto:

```
agent-config/CLAUDE.md            regras e armadilhas (sempre carregado)
  specs/PROJETO.md                spec do projeto: objetivo, limites, definição de pronto, TDD por stack
    specs/stacks/<stack>.md       uma por camada técnica (versões, comandos, estilo, testes, limites, revisão)
      specs/funcionalidades/NNN-nome/   SPEC.md + plan.md + todo.md de cada trabalho
```

| Stack | Spec |
|---|---|
| Front: Excel 2019 / VBA 7 / MSForms | [stacks/front-vba.md](stacks/front-vba.md) |
| Banco: Access `.accdb` / ACE / ADO / migrações | [stacks/banco-access.md](stacks/banco-access.md) |
| Build e operação: PowerShell 5.1 / `.bat` | [stacks/build-operacao-powershell.md](stacks/build-operacao-powershell.md) |
| Visual: layout JSON / prévia HTML no Edge | [stacks/visual-layout.md](stacks/visual-layout.md) |
| IA: servidores MCP / skills do Claude Desktop | [stacks/ia-mcp.md](stacks/ia-mcp.md) |
| Ferramentas de desenvolvimento | [stacks/ferramentas-dev.md](stacks/ferramentas-dev.md) |

## Ciclo de uma funcionalidade

`/spec` → `funcionalidades/NNN-nome/SPEC.md` (a partir de `_modelo/`) → `/plan` → `plan.md` +
`todo.md` → `/build` (uma tarefa por vez, teste primeiro, commit por tarefa) → `/test`
(`verificar-tudo`) → `/review` → `/code-simplify` → `/ship`. Ao terminar, a SPEC recebe
"Concluída em dd/mm/aaaa — commit/tag" no topo e a pasta fica como registro.

Spec de stack desatualizada é defeito: quem muda a stack atualiza a spec no mesmo commit.
