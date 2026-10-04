---

## Neste projeto: os contratos que existem

Não há API HTTP. Os contratos são: as ferramentas MCP (nome, esquema de entrada, `readOnlyHint`),
o texto de `Início!B7` e `VERSAO-FRONT.txt` (lidos pelo `INICIAR-CRM.bat`), os nomes de controle
dos formulários (`formularios/<tela>.txt` x `layout/formularios/<tela>.json`), as colunas do
banco (DDL x `modSchema`) e os códigos `CT-`/`AT-`. Mudou contrato = mudança das duas pontas no
mesmo commit + teste. Ver `specs/stacks/ia-mcp.md` e `specs/stacks/front-vba.md`.
