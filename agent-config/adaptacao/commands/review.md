## Neste projeto: o que mais olhar

Além dos cinco eixos, confira a seção **Revisão** de cada spec de stack tocada
(`agent-config/specs/stacks/`). O mínimo, sempre:

- SQL parametrizado; data só como parâmetro; `UPDATE` com `WHERE id=? AND versao=?` e log na
  mesma transação; nenhuma consulta dentro de laço.
- VBA em Windows-1252 + CRLF; nada do Microsoft 365 (Excel é 2019); nada de OCX, URL ou
  subclassing; nomes de controle e `Início!B7` intactos.
- `.ps1` UTF-8 com BOM + CRLF; nada no stdout do servidor MCP além de JSON-RPC.
- Nenhum dado real; nenhum caminho absoluto da máquina de desenvolvimento nem letra de unidade.
