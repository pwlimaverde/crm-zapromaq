---

## Neste projeto: "Discover the Stack First" já respondido

Não há Jest/pytest no sistema. Os testes são os do próprio projeto, sem dependência externa:

| Mudança em | Teste que falha primeiro | Roda com |
|---|---|---|
| Lógica VBA pura (cálculo, normalização, validação, montagem de SQL) | `Confere "nome", obtido, esperado` em `fonte/modulos/modAutoteste.bas` | `verificar-tudo -Completo` (compila e roda no Excel) |
| Regra estática do VBA/esquema (encoding, chamada inexistente, campo fora do DDL) | checagem em `ferramentas/verificacao/verificar.py` | `verificar-tudo -Rapido` |
| Build, versão, pacote, contrato B7 | `Conferir` em `testes/testar-build-lib.ps1` | `-Rapido` |
| Banco, SQL, regras do MCP de dados | caso em `IA/teste/testar-mcp.ps1` (banco fictício de `criar-banco-teste.ps1`) | padrão |
| MCP de desenvolvimento | caso em `IA/teste/testar-mcp-dev.ps1` | padrão |
| Grades e Painel com dados | `testes/testar-uso.ps1` (dentro do build) | `-Completo` |
| Aparência das telas | prévia `gerar-visual.ps1` + item 🖥 (calibração na estação) | manual |

- `modAutoteste` não acessa banco nem grava na planilha; acento por `ChrW`.
- Teste que muda o banco devolve o estado no fim (o `testar-mcp.ps1` roda de novo sem recriar).
- Dados de teste são sempre fictícios (repositório público).
