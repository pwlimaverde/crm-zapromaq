---

## Neste projeto: CRM Zapromaq

- Cada fatia deixa o projeto montável: `verificar-tudo -Rapido` a cada fatia, padrão no fim da
  tarefa, `-Completo` quando tocar `fonte\` ou `build\`.
- O `.xlsm` nunca é editado: toda mudança é em `src/BASE/SISTEMA/DADOS/fonte` (texto) e o build monta.
- Campo novo é uma fatia vertical: migração numerada → `modSchema` → consulta em `modCRM` →
  MCP (se a IA lê/grava) → testes. Ver `specs/stacks/banco-access.md`.
- Um commit por fatia, só com os arquivos dela.
